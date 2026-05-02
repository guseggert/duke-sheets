//! LibreOffice URP container auto-start.
//!
//! `ensure_lo()` is the single entry point. It is idempotent and
//! process-singleton: the first call builds the docker image (if missing),
//! sweeps stale lock files in the shared volume, starts the container, and
//! polls port 2002 until ready. Subsequent calls are O(1) no-ops.
//!
//! On failure the function **panics** with a message describing what to
//! check. Tests that depend on LibreOffice should not silently skip; if
//! the container can't be brought up that's a real failure, not a
//! reason to pretend a test ran.

use std::path::PathBuf;
use std::process::Command;
use std::sync::OnceLock;
use std::time::{Duration, Instant};

/// Docker image tag built from `tests/fixtures/pyuno/Dockerfile`.
const IMAGE_TAG: &str = "duke-sheets-pyuno";

/// Container name used for the auto-started instance. Fixed so we can
/// detect existing instances by name and replace them deterministically.
const CONTAINER_NAME: &str = "duke-lo-test";

/// Host directory shared with the LO container. Tests write fixtures
/// here; the container reads/writes them through the same path inside.
pub const SHARED_DIR: &str = "/tmp/duke-sheets-urp";

/// URP TCP port. The container publishes 2002 -> 2002.
const URP_PORT: u16 = 2002;

/// Maximum time to wait for the URP socket to come up after `docker run`.
const STARTUP_TIMEOUT: Duration = Duration::from_secs(30);

static ENSURED: OnceLock<()> = OnceLock::new();

/// Ensure a LibreOffice URP container is running and reachable on
/// localhost:2002. Builds the docker image first if it's missing.
///
/// Idempotent and cheap on subsequent calls (one TCP probe + a OnceLock
/// load). Panics with a descriptive message on failure.
pub fn ensure_lo() {
    ENSURED.get_or_init(|| {
        let _ = std::fs::create_dir_all(SHARED_DIR);

        // If a listener is already on 2002 (interactively started LO,
        // a prior test run that left the container warm, etc.) we trust
        // it and skip the build/run dance entirely.
        if port_listening(URP_PORT) {
            return;
        }

        ensure_image();
        sweep_stale_locks();
        replace_container();
        wait_for_port(URP_PORT, STARTUP_TIMEOUT);
    });
}

/// Quick TCP connect probe with a short timeout.
fn port_listening(port: u16) -> bool {
    let addr = format!("127.0.0.1:{port}").parse().unwrap();
    std::net::TcpStream::connect_timeout(&addr, Duration::from_secs(2)).is_ok()
}

/// Build the LO docker image if `docker images -q` returns nothing for it.
fn ensure_image() {
    let out = Command::new("docker")
        .args(["images", "-q", IMAGE_TAG])
        .output()
        .unwrap_or_else(|e| {
            panic!(
                "failed to run `docker images -q {IMAGE_TAG}`: {e}\n\
                 Is docker installed and the daemon running?"
            )
        });
    if !out.status.success() {
        panic!(
            "`docker images -q {IMAGE_TAG}` exited with {:?}: {}",
            out.status.code(),
            String::from_utf8_lossy(&out.stderr)
        );
    }
    if !out.stdout.is_empty() {
        return;
    }

    let context = repo_root().join("tests/fixtures/pyuno");
    eprintln!(
        "[duke-sheets-test-harness] building docker image {IMAGE_TAG} from {}...",
        context.display()
    );
    let status = Command::new("docker")
        .arg("build")
        .arg("-t")
        .arg(IMAGE_TAG)
        .arg(&context)
        .status()
        .unwrap_or_else(|e| panic!("failed to spawn `docker build`: {e}"));
    if !status.success() {
        panic!(
            "docker build of {IMAGE_TAG} from {} failed with status {:?}",
            context.display(),
            status.code()
        );
    }
}

/// LibreOffice writes `.~lock.FILE#` files when it has a doc open. If
/// the container is killed without LO closing cleanly, those locks stay
/// root-owned in the shared dir, and once enough accumulate the LO
/// loadenv starts failing new-file type detection. Sweep them via the
/// LO image (the only host-root capability in this setup).
fn sweep_stale_locks() {
    let _ = Command::new("docker")
        .args([
            "run",
            "--rm",
            "-v",
            &format!("{SHARED_DIR}:/shared"),
            IMAGE_TAG,
            "bash",
            "-c",
            "rm -f /shared/.~lock.* /shared/duke_*",
        ])
        .status();
}

/// Stop any pre-existing container with our name, kill anything else
/// holding port 2002, then start a fresh container in the background.
fn replace_container() {
    let _ = Command::new("docker")
        .args(["rm", "-f", CONTAINER_NAME])
        .output();

    // Kill any other container holding port 2002.
    if let Ok(out) = Command::new("docker")
        .args(["ps", "--filter", "publish=2002", "--quiet"])
        .output()
    {
        for cid in String::from_utf8_lossy(&out.stdout).split_whitespace() {
            let _ = Command::new("docker").args(["rm", "-f", cid]).output();
        }
    }

    let status = Command::new("docker")
        .args([
            "run",
            "-d",
            "--rm",
            "--name",
            CONTAINER_NAME,
            "-p",
            "2002:2002",
            "-v",
            &format!("{SHARED_DIR}:{SHARED_DIR}"),
            IMAGE_TAG,
            "bash",
            "-c",
            "soffice --headless --accept=\"socket,host=0.0.0.0,port=2002;urp;StarOffice.ComponentContext\" & sleep 999999",
        ])
        .status()
        .unwrap_or_else(|e| panic!("failed to spawn `docker run` for {CONTAINER_NAME}: {e}"));
    if !status.success() {
        panic!(
            "`docker run` for {CONTAINER_NAME} exited with {:?}",
            status.code()
        );
    }
}

fn wait_for_port(port: u16, timeout: Duration) {
    let start = Instant::now();
    while start.elapsed() < timeout {
        if port_listening(port) {
            eprintln!(
                "[duke-sheets-test-harness] LibreOffice URP ready on port {port} (took {:.1}s)",
                start.elapsed().as_secs_f64()
            );
            return;
        }
        std::thread::sleep(Duration::from_millis(500));
    }
    panic!(
        "LibreOffice URP did not come up on localhost:{port} within {:?}.\n\
         Inspect with: docker logs {CONTAINER_NAME}",
        timeout
    );
}

/// Workspace root, derived from the manifest dir of this crate at compile
/// time. `crates/duke-sheets-test-harness` -> repo root.
fn repo_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .and_then(|p| p.parent())
        .map(PathBuf::from)
        .expect("CARGO_MANIFEST_DIR has no grandparent (expected crates/<name>)")
}
