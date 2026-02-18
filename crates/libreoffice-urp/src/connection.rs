//! High-level URP connection management.
//!
//! `UrpConnection` manages the TCP transport, protocol state, and provides
//! methods for calling remote UNO methods and handling protocol negotiation.

use std::sync::atomic::{AtomicU64, Ordering};

use bytes::{BufMut, BytesMut};
use tokio::net::TcpStream;

use crate::error::{Result, UrpError};
use crate::interface::{self, MethodDef};
use crate::marshal;
use crate::protocol::{
    self, ReaderState, UrpMessage, WriterState,
    OID_PROTOCOL_PROPERTIES, TID_PROTOCOL_PROPERTIES,
    FN_REQUEST_CHANGE, FN_COMMIT_CHANGE,
};
use crate::proxy::{self, UnoProxy};
use crate::transport::Transport;
use crate::types::{Type, UnoValue};

/// A URP connection to a LibreOffice instance.
pub struct UrpConnection {
    transport: Transport,
    reader_state: ReaderState,
    writer_state: WriterState,
    tid_counter: AtomicU64,
    /// Whether CurrentContext mode is active (after successful negotiation).
    current_context_mode: bool,
}

impl UrpConnection {
    /// Connect to a LibreOffice instance listening on the given address.
    pub async fn connect(host: &str, port: u16) -> Result<Self> {
        let addr = format!("{host}:{port}");
        let stream = TcpStream::connect(&addr).await.map_err(|e| {
            UrpError::Io(std::io::Error::new(
                e.kind(),
                format!("failed to connect to LibreOffice at {addr}: {e}"),
            ))
        })?;

        tracing::info!("Connected to LibreOffice at {addr}");

        let mut conn = Self {
            transport: Transport::new(stream),
            reader_state: ReaderState::new(),
            writer_state: WriterState::new(),
            tid_counter: AtomicU64::new(1),
            current_context_mode: false,
        };

        // Perform protocol negotiation
        conn.negotiate_protocol().await?;

        Ok(conn)
    }

    /// Generate a unique TID (transaction ID) for a new request.
    fn next_tid(&self) -> Vec<u8> {
        let n = self.tid_counter.fetch_add(1, Ordering::Relaxed);
        format!("tid-{n}").into_bytes()
    }

    // ========================================================================
    // Protocol negotiation
    // ========================================================================

    /// Perform the URP protocol properties negotiation.
    ///
    /// Sequence:
    /// 1. Send requestChange with a random number
    /// 2. Read reply (or incoming requestChange from the other side)
    /// 3. If we win (reply=1), send commitChange with CurrentContext property
    /// 4. If we lose (reply=0), wait for their commitChange
    async fn negotiate_protocol(&mut self) -> Result<()> {
        let random: i32 = rand::random();
        tracing::debug!("Sending requestChange with random={random}");

        // Send requestChange
        let tid = TID_PROTOCOL_PROPERTIES.to_vec();
        let proto_type = Type::interface("com.sun.star.bridge.XProtocolProperties");

        let mut body = BytesMut::new();
        marshal::write_value(&mut body, &UnoValue::Long(random), &Type::long());

        let msg = self.writer_state.encode_request(
            FN_REQUEST_CHANGE,
            &proto_type,
            OID_PROTOCOL_PROPERTIES,
            &tid,
            true,
            &body,
        );

        self.transport.send_message(&msg).await?;

        // Read response — could be a reply to our requestChange, or an incoming
        // requestChange from the other side
        loop {
            let data = self.transport.recv_message().await?;
            let message = self.reader_state.decode_message(data)?;

            match message {
                UrpMessage::Reply(reply) => {
                    if reply.is_exception {
                        // The other side doesn't support protocol properties negotiation.
                        // This is fine — we just skip CurrentContext mode.
                        tracing::debug!("Protocol negotiation: exception in reply, skipping CurrentContext");
                        return Ok(());
                    }

                    // Parse the reply value (should be a long)
                    let mut body = reply.body;
                    let result = if body.len() >= 4 {
                        marshal::read_value(&mut body, &Type::long())?
                    } else {
                        tracing::debug!("Protocol negotiation: reply body too short ({}), assuming 1", body.len());
                        UnoValue::Long(1) // Assume success if empty
                    };

                    match result {
                        UnoValue::Long(1) => {
                            // We won — send commitChange
                            tracing::debug!("Protocol negotiation: we won, sending commitChange");
                            self.send_commit_change(&tid, &proto_type).await?;
                            self.current_context_mode = true;
                            return Ok(());
                        }
                        UnoValue::Long(0) => {
                            // We lost — wait for their commitChange, then reply
                            tracing::debug!("Protocol negotiation: we lost, waiting for their commitChange");
                            self.wait_for_commit_change().await?;
                            self.current_context_mode = true;
                            return Ok(());
                        }
                        UnoValue::Long(-1) => {
                            // Tie — retry with new random number
                            tracing::debug!("Protocol negotiation: tie, retrying");
                            // For simplicity in the prototype, just accept without CurrentContext
                            return Ok(());
                        }
                        other => {
                            tracing::warn!("Protocol negotiation: unexpected reply value: {other:?}");
                            return Ok(());
                        }
                    }
                }
                UrpMessage::Request(req) => {
                    // Incoming requestChange from the other side
                    if req.function_id == FN_REQUEST_CHANGE {
                        let their_random = if req.body.len() >= 4 {
                            let mut body = req.body;
                            marshal::read_value(&mut body, &Type::long())
                                .unwrap_or(UnoValue::Long(0))
                        } else {
                            UnoValue::Long(0)
                        };

                        let their_val = match their_random {
                            UnoValue::Long(n) => n,
                            _ => 0,
                        };

                        // Determine who wins: larger random number wins
                        let reply_val = if random > their_val {
                            0i32 // They lose
                        } else if random < their_val {
                            1i32 // They win
                        } else {
                            -1i32 // Tie
                        };

                        tracing::debug!(
                            "Protocol negotiation collision: our={random} vs their={their_val}, reply_val={reply_val}"
                        );

                        // Send reply
                        let mut reply_body = BytesMut::new();
                        marshal::write_value(&mut reply_body, &UnoValue::Long(reply_val), &Type::long());
                        let reply_msg = self.writer_state.encode_reply(
                            &req.tid,
                            false,
                            &reply_body,
                        );
                        self.transport.send_message(&reply_msg).await?;

                        if reply_val == 1 {
                            // They won — wait for their commitChange
                            tracing::debug!("Protocol negotiation: they won, waiting for commitChange");
                            self.wait_for_commit_change().await?;
                            self.current_context_mode = true;
                            return Ok(());
                        }
                        // Otherwise continue looping for our reply
                    } else if req.function_id == FN_COMMIT_CHANGE {
                        // They're committing changes
                        tracing::debug!("Protocol negotiation: received commitChange");
                        let reply_msg = self.writer_state.encode_reply(&req.tid, false, &[]);
                        self.transport.send_message(&reply_msg).await?;
                        self.current_context_mode = true;
                        return Ok(());
                    }
                }
            }
        }
    }

    async fn send_commit_change(&mut self, tid: &[u8], proto_type: &Type) -> Result<()> {
        // commitChange(Sequence<ProtocolProperty>)
        // We send CurrentContext=true
        let mut body = BytesMut::new();

        // Sequence of ProtocolProperty, length 1
        marshal::write_compressed(&mut body, 1);
        // ProtocolProperty struct: Name (string) + Value (any)
        marshal::write_string(&mut body, "CurrentContext");
        // Value as Any: type=void, value=void (LO sends Void for the value)
        marshal::write_type(&mut body, &Type::void(), 0xFFFF, false);

        let msg = self.writer_state.encode_request(
            FN_COMMIT_CHANGE,
            proto_type,
            OID_PROTOCOL_PROPERTIES,
            tid,
            true,
            &body,
        );
        self.transport.send_message(&msg).await?;

        // Wait for reply
        let data = self.transport.recv_message().await?;
        let message = self.reader_state.decode_message(data)?;
        match message {
            UrpMessage::Reply(reply) => {
                if reply.is_exception {
                    tracing::warn!("commitChange raised exception");
                }
            }
            _ => {
                tracing::warn!("Expected reply to commitChange, got request");
            }
        }

        Ok(())
    }

    async fn wait_for_commit_change(&mut self) -> Result<()> {
        loop {
            let data = self.transport.recv_message().await?;
            let message = self.reader_state.decode_message(data)?;
            match &message {
                UrpMessage::Request(req) => {
                    if req.function_id == FN_COMMIT_CHANGE {
                        // Reply with success
                        let reply_msg = self.writer_state.encode_reply(&req.tid, false, &[]);
                        self.transport.send_message(&reply_msg).await?;
                        tracing::debug!("Protocol negotiation: replied to commitChange");
                        return Ok(());
                    }
                }
                UrpMessage::Reply(_reply) => {
                    // This is the reply to our requestChange — ignore it and keep waiting
                    tracing::trace!("Ignoring reply while waiting for commitChange");
                }
            }
        }
    }

    // ========================================================================
    // Method invocation
    // ========================================================================

    /// Call a method on a remote UNO object.
    pub async fn call(
        &mut self,
        proxy: &UnoProxy,
        method: &MethodDef,
        args: &[UnoValue],
    ) -> Result<UnoValue> {
        let body = proxy::serialize_params_cached(method, args, &mut self.writer_state.oid_cache)?;
        let tid = self.next_tid();

        let mut full_body = BytesMut::new();

        // In CurrentContext mode, prepend a null XCurrentContext interface reference
        if self.current_context_mode && !method.one_way
            && method.name != "requestChange" && method.name != "commitChange"
        {
            // Null interface reference: empty string + 0xFFFF cache index
            marshal::write_string(&mut full_body, "");
            full_body.put_u16(0xFFFF);
        }
        full_body.extend_from_slice(&body);

        let msg = self.writer_state.encode_request(
            method.index,
            &proxy.interface_type,
            &proxy.oid,
            &tid,
            !method.one_way,
            &full_body,
        );

        tracing::trace!(
            "Calling {}() on OID={}, fn_id={}",
            method.name, proxy.oid, method.index
        );

        self.transport.send_message(&msg).await?;

        if method.one_way {
            return Ok(UnoValue::Void);
        }

        // Wait for the reply
        // In a full implementation, we'd match on TID to handle out-of-order replies.
        // For the prototype, we assume replies come in order.
        loop {
            let data = self.transport.recv_message().await?;
            let message = self.reader_state.decode_message(data)?;

            match message {
                UrpMessage::Reply(reply) => {
                    if reply.is_exception {
                        // Parse the exception from the body (it's an Any)
                        let mut body = reply.body;
                        if body.is_empty() {
                            return Err(UrpError::RemoteException("(empty exception)".into()));
                        }
                        let exc_value = marshal::read_value_cached(
                            &mut body,
                            &Type::any(),
                            &mut self.reader_state.oid_cache,
                        )?;
                        let msg = match &exc_value {
                            UnoValue::Any(a) => match &a.value {
                                UnoValue::Exception(e) => e.message.clone(),
                                UnoValue::Struct(members) => {
                                    // Try to extract message from first member
                                    members.first()
                                        .and_then(|v| v.as_string())
                                        .unwrap_or("unknown exception")
                                        .to_string()
                                }
                                other => format!("{other:?}"),
                            },
                            other => format!("{other:?}"),
                        };
                        return Err(UrpError::RemoteException(msg));
                    }
                    let result = proxy::deserialize_return_cached(
                        method,
                        reply.body,
                        &mut self.reader_state.oid_cache,
                    )?;
                    tracing::trace!("{}() returned {:?}", method.name, result);
                    return Ok(result);
                }
                UrpMessage::Request(req) => {
                    // Handle incoming requests while waiting for our reply.
                    // The most common case is `release` (fire-and-forget).
                    if req.function_id == protocol::FN_RELEASE {
                        // Ignore release — we don't track local objects
                        tracing::trace!("Ignoring incoming release for OID={}", req.oid);
                        continue;
                    }

                    // For other incoming requests, send back a void reply
                    // (this is a simplification — a full impl would dispatch to handlers)
                    if req.must_reply {
                        let reply_msg = self.writer_state.encode_reply(&req.tid, false, &[]);
                        self.transport.send_message(&reply_msg).await?;
                    }
                }
            }
        }
    }

    /// Call queryInterface on a remote object to get a proxy for a different interface.
    ///
    /// Always uses XInterface as the request type, since queryInterface is
    /// defined on XInterface (function ID 0).
    pub async fn query_interface(
        &mut self,
        proxy: &UnoProxy,
        target_type: Type,
    ) -> Result<Option<UnoProxy>> {
        // queryInterface must always be called via XInterface type,
        // regardless of the proxy's current interface type.
        let xi_proxy = UnoProxy::new(
            proxy.oid.clone(),
            Type::interface("com.sun.star.uno.XInterface"),
        );
        let method = interface::query_interface();
        let args = [UnoValue::Type(target_type.clone())];
        let result = self.call(&xi_proxy, &method, &args).await?;
        proxy::extract_query_interface_result(result, target_type)
    }

    /// Send a `release` message for a remote object (one-way, no reply).
    pub async fn release(&mut self, proxy: &UnoProxy) -> Result<()> {
        // release is on XInterface, use XInterface type
        let xi_proxy = UnoProxy::new(
            proxy.oid.clone(),
            Type::interface("com.sun.star.uno.XInterface"),
        );
        let method = interface::release();
        self.call(&xi_proxy, &method, &[]).await?;
        Ok(())
    }

    // ========================================================================
    // Bootstrap helpers
    // ========================================================================

    /// Get the initial object from the bridge.
    ///
    /// The initial object is identified by the well-known name provided in the
    /// `--accept` string (e.g., "StarOffice.ComponentContext").
    pub async fn get_initial_object(&mut self, name: &str) -> Result<UnoProxy> {
        // The initial object is obtained by sending queryInterface to the
        // well-known OID with XInterface type.
        let initial_proxy = UnoProxy::new(
            name.to_string(),
            Type::interface("com.sun.star.uno.XInterface"),
        );

        let xi_type = Type::interface("com.sun.star.uno.XInterface");
        let result = self.query_interface(&initial_proxy, xi_type.clone()).await?;

        match result {
            Some(proxy) => Ok(proxy),
            None => {
                // The initial object should always support XInterface.
                // If queryInterface returned void, the OID itself may be valid.
                Ok(initial_proxy)
            }
        }
    }

    /// Bootstrap the full UNO environment: get XComponentContext, then
    /// XMultiComponentFactory (ServiceManager), then create the Desktop.
    ///
    /// Returns (component_context, service_manager, desktop).
    pub async fn bootstrap(&mut self) -> Result<(UnoProxy, UnoProxy, UnoProxy)> {
        use crate::types::type_names;

        tracing::debug!("Bootstrapping UNO environment...");

        // 1. Get the initial XComponentContext
        let ctx_proxy_raw = self.get_initial_object("StarOffice.ComponentContext").await?;
        tracing::debug!("Got initial object OID={}", ctx_proxy_raw.oid);

        // queryInterface for XComponentContext
        let ctx_type = Type::interface(type_names::X_COMPONENT_CONTEXT);
        let ctx_proxy = self.query_interface(&ctx_proxy_raw, ctx_type.clone()).await?
            .unwrap_or_else(|| UnoProxy::new(ctx_proxy_raw.oid.clone(), ctx_type));

        tracing::debug!("Got XComponentContext OID={}", ctx_proxy.oid);

        // 2. Get the ServiceManager from the context
        let get_sm = interface::get_service_manager();
        let sm_result = self.call(&ctx_proxy, &get_sm, &[]).await?;
        let sm_oid = proxy::extract_oid_from_return(&sm_result)
            .ok_or_else(|| UrpError::Protocol("getServiceManager returned null".into()))?;
        let sm_proxy = UnoProxy::new(
            sm_oid,
            Type::interface(type_names::X_MULTI_COMPONENT_FACTORY),
        );
        tracing::info!("Got ServiceManager: {}", sm_proxy.oid);

        // 3. Create the Desktop via the ServiceManager
        let create_inst = interface::create_instance_with_context();
        let desktop_result = self.call(&sm_proxy, &create_inst, &[
            UnoValue::String(type_names::SERVICE_DESKTOP.to_string()),
            UnoValue::Interface(ctx_proxy.oid.clone()),
        ]).await?;

        let desktop_oid = proxy::extract_oid_from_return(&desktop_result)
            .ok_or_else(|| UrpError::Protocol("createInstanceWithContext(Desktop) returned null".into()))?;

        // queryInterface for XComponentLoader on the Desktop
        let loader_type = Type::interface(type_names::X_COMPONENT_LOADER);
        let desktop_raw = UnoProxy::new(desktop_oid, loader_type.clone());
        let desktop_proxy = self.query_interface(&desktop_raw, loader_type.clone()).await?
            .unwrap_or(desktop_raw);

        tracing::info!("Got Desktop/ComponentLoader: {}", desktop_proxy.oid);

        Ok((ctx_proxy, sm_proxy, desktop_proxy))
    }
}
