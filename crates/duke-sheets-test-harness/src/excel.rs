//! Excel COM bridge auto-start.
//!
//! `ensure_excel_bridge()` is the single entry point. It is idempotent
//! and process-singleton: if the Windows VM bridge isn't responsive on
//! localhost:9876 it spawns `tools/vm/qemu-start.sh` in the background
//! and polls until the bridge answers an Init ping.
//!
//! On failure the function **panics** with a message describing what to
//! check. Tests must not silently skip when the VM is unavailable; if
//! the bridge can't be brought up, the run fails loud.

use std::io::{Read, Write};
use std::net::TcpStream;
use std::path::PathBuf;
use std::process::{Command, Stdio};
use std::sync::OnceLock;
use std::time::{Duration, Instant};

const BRIDGE_PORT: u16 = 9876;
const STARTUP_TIMEOUT: Duration = Duration::from_secs(360);

/// Host directory shared with the Windows VM via SMB. Tests put files
/// here; inside the VM the same content is reachable as `\\10.0.2.4\qemu`.
pub const SHARED_DIR: &str = "/tmp/duke-sheets-excel";

static ENSURED: OnceLock<()> = OnceLock::new();

/// Ensure the Excel COM bridge is running and responsive on
/// localhost:9876. Spawns `tools/vm/qemu-start.sh` if it isn't.
///
/// Idempotent and cheap on subsequent calls. Panics with a descriptive
/// message on failure.
pub fn ensure_excel_bridge() {
    ENSURED.get_or_init(|| {
        let _ = std::fs::create_dir_all(SHARED_DIR);

        if bridge_responsive() {
            return;
        }

        spawn_vm();
        wait_for_bridge(STARTUP_TIMEOUT);
    });
}

/// Send `{"id":1,"cmd":"Init","params":{}}` to the bridge and check
/// that the response contains `"ok"`. Matches the same readiness check
/// used by the mise tasks.
fn bridge_responsive() -> bool {
    let addr = format!("127.0.0.1:{BRIDGE_PORT}").parse().unwrap();
    let Ok(mut stream) = TcpStream::connect_timeout(&addr, Duration::from_secs(2)) else {
        return false;
    };
    let _ = stream.set_read_timeout(Some(Duration::from_secs(2)));
    let _ = stream.set_write_timeout(Some(Duration::from_secs(2)));
    if writeln!(stream, "{{\"id\":1,\"cmd\":\"Init\",\"params\":{{}}}}").is_err() {
        return false;
    }
    let mut buf = String::new();
    let _ = stream.read_to_string(&mut buf);
    buf.contains("\"ok\"")
}

/// Launch `tools/vm/qemu-start.sh` detached so it survives this test
/// process. The shell script forks qemu-system-x86_64 internally; we
/// just need to avoid blocking on it here.
fn spawn_vm() {
    let script = repo_root().join("tools/vm/qemu-start.sh");
    if !script.exists() {
        panic!(
            "VM start script not found at {}.\n\
             Run `mise run vm:build-qemu` to build QEMU first.",
            script.display()
        );
    }
    eprintln!(
        "[duke-sheets-test-harness] starting Windows VM via {}...",
        script.display()
    );
    Command::new("bash")
        .arg(&script)
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
        .unwrap_or_else(|e| panic!("failed to spawn VM start script: {e}"));
}

fn wait_for_bridge(timeout: Duration) {
    let start = Instant::now();
    let mut attempt: u32 = 0;
    while start.elapsed() < timeout {
        if bridge_responsive() {
            eprintln!(
                "[duke-sheets-test-harness] Excel COM bridge ready on port {BRIDGE_PORT} (took {:.1}s)",
                start.elapsed().as_secs_f64()
            );
            return;
        }
        attempt += 1;
        if attempt % 30 == 0 {
            eprintln!(
                "[duke-sheets-test-harness] still waiting for Excel COM bridge ({:.0}s elapsed)...",
                start.elapsed().as_secs_f64()
            );
        }
        std::thread::sleep(Duration::from_secs(2));
    }
    panic!(
        "Excel COM bridge did not respond on localhost:{BRIDGE_PORT} within {:?}.\n\
         Inspect VM state with: ps -ef | grep qemu-system-x86_64\n\
         Or run `mise run vm:start` manually and watch the VNC display.",
        timeout
    );
}

fn repo_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .and_then(|p| p.parent())
        .map(PathBuf::from)
        .expect("CARGO_MANIFEST_DIR has no grandparent (expected crates/<name>)")
}
