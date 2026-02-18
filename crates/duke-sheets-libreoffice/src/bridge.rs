//! LibreOffice bridge: manages soffice process and URP connection.

use std::path::PathBuf;
use std::process::Stdio;
use std::time::Duration;

use libreoffice_urp::connection::UrpConnection;
use libreoffice_urp::interface;
use libreoffice_urp::proxy::{self, UnoProxy};
use libreoffice_urp::types::{Type, UnoValue, type_names};
use tokio::process::{Child, Command};
use tokio::time::sleep;

use crate::error::{BridgeError, Result};
use crate::workbook::Workbook;

/// Configuration for the LibreOffice bridge.
pub struct LibreOfficeConfig {
    /// Path to the `soffice` executable. If None, searches PATH.
    pub soffice_path: Option<PathBuf>,
    /// Host to connect to (or listen on). Default: "localhost".
    pub host: String,
    /// Port for URP socket communication. Default: 2002.
    pub port: u16,
    /// Timeout for waiting for LibreOffice to start. Default: 30 seconds.
    pub startup_timeout: Duration,
    /// Extra arguments to pass to soffice.
    pub extra_args: Vec<String>,
}

impl Default for LibreOfficeConfig {
    fn default() -> Self {
        Self {
            soffice_path: None,
            host: "localhost".to_string(),
            port: 2002,
            startup_timeout: Duration::from_secs(30),
            extra_args: Vec::new(),
        }
    }
}

/// The main handle for communicating with LibreOffice via URP.
pub struct LibreOfficeBridge {
    conn: UrpConnection,
    /// The child process, if we spawned it.
    _child: Option<Child>,
    /// Bootstrap objects
    _ctx: UnoProxy,
    _sm: UnoProxy,
    desktop: UnoProxy,
}

impl LibreOfficeBridge {
    /// Connect to an already-running LibreOffice instance.
    pub async fn connect(host: &str, port: u16) -> Result<Self> {
        let mut conn = UrpConnection::connect(host, port).await?;
        let (ctx, sm, desktop) = conn.bootstrap().await?;

        Ok(Self {
            conn,
            _child: None,
            _ctx: ctx,
            _sm: sm,
            desktop,
        })
    }

    /// Start a new LibreOffice instance and connect to it.
    pub async fn start(config: LibreOfficeConfig) -> Result<Self> {
        let soffice = config
            .soffice_path
            .unwrap_or_else(|| PathBuf::from("soffice"));

        let accept_arg = format!(
            "socket,host={},port={};urp;StarOffice.ComponentContext",
            config.host, config.port
        );

        let mut cmd = Command::new(&soffice);
        cmd.arg("--headless")
            .arg("--invisible")
            .arg("--nocrashreport")
            .arg("--nodefault")
            .arg("--nologo")
            .arg("--nofirststartwizard")
            .arg("--norestore")
            .arg(format!("--accept={accept_arg}"));

        for arg in &config.extra_args {
            cmd.arg(arg);
        }

        cmd.stdin(Stdio::null())
            .stdout(Stdio::null())
            .stderr(Stdio::null());

        tracing::info!("Starting LibreOffice: {:?}", cmd);
        let child = cmd.spawn().map_err(|e| {
            if e.kind() == std::io::ErrorKind::NotFound {
                BridgeError::NotFound
            } else {
                BridgeError::SpawnFailed(e)
            }
        })?;

        // Wait for LibreOffice to start listening
        let timeout = config.startup_timeout;
        let start = tokio::time::Instant::now();
        let mut conn = loop {
            if start.elapsed() > timeout {
                return Err(BridgeError::Timeout(timeout.as_secs()));
            }

            match UrpConnection::connect(&config.host, config.port).await {
                Ok(conn) => break conn,
                Err(_) => {
                    sleep(Duration::from_millis(500)).await;
                }
            }
        };

        let (ctx, sm, desktop) = conn.bootstrap().await?;

        Ok(Self {
            conn,
            _child: Some(child),
            _ctx: ctx,
            _sm: sm,
            desktop,
        })
    }

    /// Create a new empty spreadsheet workbook.
    pub async fn create_workbook(&mut self) -> Result<Workbook<'_>> {
        // loadComponentFromURL("private:factory/scalc", "_blank", 0, ())
        let method = interface::load_component_from_url();
        let result = self.conn.call(&self.desktop, &method, &[
            UnoValue::String("private:factory/scalc".to_string()),
            UnoValue::String("_blank".to_string()),
            UnoValue::Long(0),
            UnoValue::Sequence(vec![]), // empty PropertyValue sequence
        ]).await?;

        let doc_oid = proxy::extract_oid_from_return(&result)
            .ok_or_else(|| BridgeError::OperationFailed(
                "loadComponentFromURL returned null".into()
            ))?;

        let doc_proxy = UnoProxy::new(
            doc_oid,
            Type::interface(type_names::X_COMPONENT),
        );

        tracing::info!("Created new spreadsheet: {}", doc_proxy.oid);

        Ok(Workbook::new(&mut self.conn, doc_proxy))
    }

    /// Open an existing workbook from a file path.
    pub async fn open_workbook(&mut self, path: &str) -> Result<Workbook<'_>> {
        // Convert to file:// URL
        let url = if path.starts_with("file://") {
            path.to_string()
        } else {
            let abs = if path.starts_with('/') {
                path.to_string()
            } else {
                std::env::current_dir()
                    .unwrap_or_default()
                    .join(path)
                    .display()
                    .to_string()
            };
            format!("file://{abs}")
        };

        let method = interface::load_component_from_url();
        let result = self.conn.call(&self.desktop, &method, &[
            UnoValue::String(url),
            UnoValue::String("_blank".to_string()),
            UnoValue::Long(0),
            UnoValue::Sequence(vec![]),
        ]).await?;

        let doc_oid = proxy::extract_oid_from_return(&result)
            .ok_or_else(|| BridgeError::OperationFailed(
                "loadComponentFromURL returned null for open".into()
            ))?;

        let doc_proxy = UnoProxy::new(
            doc_oid,
            Type::interface(type_names::X_COMPONENT),
        );

        Ok(Workbook::new(&mut self.conn, doc_proxy))
    }

    /// Get a mutable reference to the URP connection.
    pub fn conn(&mut self) -> &mut UrpConnection {
        &mut self.conn
    }

    /// Get a reference to the Desktop proxy.
    pub fn desktop(&self) -> &UnoProxy {
        &self.desktop
    }

    /// Create a new empty spreadsheet workbook using an externally-owned
    /// connection and desktop proxy. Used by `FixtureBuilder`.
    pub async fn create_workbook_with<'a>(
        conn: &'a mut UrpConnection,
        desktop: &UnoProxy,
    ) -> Result<Workbook<'a>> {
        let method = interface::load_component_from_url();
        let result = conn.call(desktop, &method, &[
            UnoValue::String("private:factory/scalc".to_string()),
            UnoValue::String("_blank".to_string()),
            UnoValue::Long(0),
            UnoValue::Sequence(vec![]),
        ]).await?;

        let doc_oid = proxy::extract_oid_from_return(&result)
            .ok_or_else(|| BridgeError::OperationFailed(
                "loadComponentFromURL returned null".into()
            ))?;

        let doc_proxy = UnoProxy::new(
            doc_oid,
            Type::interface(type_names::X_COMPONENT),
        );

        tracing::info!("Created new spreadsheet: {}", doc_proxy.oid);

        Ok(Workbook::new(conn, doc_proxy))
    }

    /// Shut down the bridge. If we spawned LibreOffice, kills the process.
    pub async fn shutdown(self) -> Result<()> {
        // Try to close gracefully via Desktop.terminate()
        // (Desktop extends XDesktop which has terminate())
        // For the prototype, just drop the connection.

        if let Some(mut child) = self._child {
            let _ = child.kill().await;
        }

        Ok(())
    }
}
