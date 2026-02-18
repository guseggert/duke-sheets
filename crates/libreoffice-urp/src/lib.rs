//! Pure Rust implementation of the UNO Remote Protocol (URP) for communicating
//! with LibreOffice.
//!
//! URP is the binary protocol that LibreOffice uses for interprocess communication.
//! When LibreOffice is started with a socket listener:
//!
//! ```text
//! soffice --headless --accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager"
//! ```
//!
//! This crate can connect and make UNO method calls to control the application:
//! create/open documents, manipulate cells, save files, etc.
//!
//! # Architecture
//!
//! The protocol has three layers:
//! - **Transport** (`transport.rs`): TCP connection with block-layer framing
//! - **Marshal** (`marshal.rs`): Binary serialization of UNO types
//! - **Protocol** (`protocol.rs`): Request/reply messages with caching
//!
//! On top of these, `connection.rs` provides the `UrpConnection` which manages
//! the full lifecycle: connect, negotiate protocol properties, bootstrap the
//! UNO environment, and invoke methods on remote objects.
//!
//! # Example
//!
//! ```rust,no_run
//! use libreoffice_urp::connection::UrpConnection;
//!
//! # async fn example() -> libreoffice_urp::error::Result<()> {
//! let mut conn = UrpConnection::connect("localhost", 2002).await?;
//! let (ctx, sm, desktop) = conn.bootstrap().await?;
//! // Now you can use `desktop` to load/create documents...
//! # Ok(())
//! # }
//! ```

pub mod connection;
pub mod error;
pub mod interface;
pub mod marshal;
pub mod protocol;
pub mod proxy;
pub mod transport;
pub mod types;

// Re-export key types
pub use connection::UrpConnection;
pub use error::{Result, UrpError};
pub use proxy::UnoProxy;
pub use types::{Any, Type, TypeClass, UnoValue};
