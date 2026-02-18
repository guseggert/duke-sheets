//! Block-layer framing for URP over TCP.
//!
//! The URP block layer wraps messages in blocks with 8-byte headers:
//! - Bytes 0..4: Block size (u32 BE) — number of bytes after the header
//! - Bytes 4..8: Message count (u32 BE) — number of messages in this block
//!
//! In practice, LibreOffice sends one message per block.

use bytes::{Bytes, BytesMut};
use tokio::io::{AsyncReadExt, AsyncWriteExt};
use tokio::net::TcpStream;

use crate::error::{Result, UrpError};

/// A URP transport — handles block-layer framing over a TCP connection.
pub struct Transport {
    stream: TcpStream,
}

impl Transport {
    /// Create a new transport wrapping a TCP stream.
    pub fn new(stream: TcpStream) -> Self {
        Self { stream }
    }

    /// Send a single message as a block.
    pub async fn send_message(&mut self, data: &[u8]) -> Result<()> {
        let size = data.len() as u32;
        let count: u32 = 1;

        // Write the 8-byte block header
        let mut header = [0u8; 8];
        header[0..4].copy_from_slice(&size.to_be_bytes());
        header[4..8].copy_from_slice(&count.to_be_bytes());

        self.stream.write_all(&header).await?;
        self.stream.write_all(data).await?;
        self.stream.flush().await?;
        Ok(())
    }

    /// Receive a complete block, returning the raw payload bytes and message count.
    pub async fn recv_block(&mut self) -> Result<(Bytes, u32)> {
        // Read the 8-byte header
        let mut header = [0u8; 8];
        match self.stream.read_exact(&mut header).await {
            Ok(_) => {}
            Err(e) if e.kind() == std::io::ErrorKind::UnexpectedEof => {
                return Err(UrpError::ConnectionClosed);
            }
            Err(e) => return Err(UrpError::Io(e)),
        }

        let size = u32::from_be_bytes([header[0], header[1], header[2], header[3]]);
        let count = u32::from_be_bytes([header[4], header[5], header[6], header[7]]);

        if size == 0 {
            return Ok((Bytes::new(), count));
        }

        // Read the payload
        let mut payload = BytesMut::zeroed(size as usize);
        self.stream.read_exact(&mut payload).await?;

        Ok((payload.freeze(), count))
    }

    /// Receive a single message. If a block contains multiple messages,
    /// this returns the entire block payload (caller must parse message boundaries).
    pub async fn recv_message(&mut self) -> Result<Bytes> {
        let (payload, _count) = self.recv_block().await?;
        Ok(payload)
    }

    /// Get a reference to the underlying stream (for shutdown).
    pub fn stream(&self) -> &TcpStream {
        &self.stream
    }

    /// Consume the transport and return the inner stream.
    pub fn into_stream(self) -> TcpStream {
        self.stream
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use tokio::net::TcpListener;

    #[tokio::test]
    async fn test_send_recv_message() {
        let listener = TcpListener::bind("127.0.0.1:0").await.unwrap();
        let addr = listener.local_addr().unwrap();

        let server = tokio::spawn(async move {
            let (stream, _) = listener.accept().await.unwrap();
            let mut transport = Transport::new(stream);
            let msg = transport.recv_message().await.unwrap();
            assert_eq!(msg.as_ref(), b"hello urp");
        });

        let client_stream = TcpStream::connect(addr).await.unwrap();
        let mut client = Transport::new(client_stream);
        client.send_message(b"hello urp").await.unwrap();

        server.await.unwrap();
    }

    #[tokio::test]
    async fn test_send_recv_empty() {
        let listener = TcpListener::bind("127.0.0.1:0").await.unwrap();
        let addr = listener.local_addr().unwrap();

        let server = tokio::spawn(async move {
            let (stream, _) = listener.accept().await.unwrap();
            let mut transport = Transport::new(stream);
            let msg = transport.recv_message().await.unwrap();
            assert!(msg.is_empty());
        });

        let client_stream = TcpStream::connect(addr).await.unwrap();
        let mut client = Transport::new(client_stream);
        client.send_message(b"").await.unwrap();

        server.await.unwrap();
    }

    #[tokio::test]
    async fn test_multiple_messages() {
        let listener = TcpListener::bind("127.0.0.1:0").await.unwrap();
        let addr = listener.local_addr().unwrap();

        let server = tokio::spawn(async move {
            let (stream, _) = listener.accept().await.unwrap();
            let mut transport = Transport::new(stream);

            let msg1 = transport.recv_message().await.unwrap();
            assert_eq!(msg1.as_ref(), b"first");

            let msg2 = transport.recv_message().await.unwrap();
            assert_eq!(msg2.as_ref(), b"second");
        });

        let client_stream = TcpStream::connect(addr).await.unwrap();
        let mut client = Transport::new(client_stream);
        client.send_message(b"first").await.unwrap();
        client.send_message(b"second").await.unwrap();

        server.await.unwrap();
    }
}
