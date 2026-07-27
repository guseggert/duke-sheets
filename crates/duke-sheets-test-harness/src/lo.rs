//! LibreOffice URP container auto-start.
//!
//! `ensure_lo()` is the single entry point. It is idempotent and
//! process-singleton: it reuses a container that already answers on port
//! 2002, and otherwise builds the docker image (if missing), sweeps stale
//! lock files in the shared volume, starts the container, and polls until
//! ready. Subsequent calls are O(1) no-ops. This mirrors
//! `excel::ensure_excel_bridge()`, which likewise probes before spawning.
//!
//! On failure the function **panics** with a message describing what to
//! check. Tests that depend on LibreOffice should not silently skip; if
//! the container can't be brought up that's a real failure, not a
//! reason to pretend a test ran.

use std::io::Read;
use std::net::TcpStream;
use std::path::PathBuf;
use std::process::Command;
use std::sync::OnceLock;
use std::time::{Duration, Instant};

/// Docker image tag built from `tests/fixtures/pyuno/Dockerfile`.
const IMAGE_TAG: &str = "duke-sheets-pyuno:2";

/// Container name used for the auto-started instance. Fixed so we can
/// detect existing instances by name and replace them deterministically.
const CONTAINER_NAME: &str = "duke-lo-test";

/// Host directory shared with the LO container. Tests write fixtures
/// here; the container reads/writes them through the same path inside.
pub const SHARED_DIR: &str = "/tmp/duke-sheets-urp";

/// The single published URP port. haproxy listens here and forwards each
/// connection to the soffice backend with the fewest active connections.
/// Tests hold one connection per test body, so "fewest connections" is
/// "fewest running tests".
pub const URP_PORT: u16 = 2002;

/// First internal backend port (not published).
const BACKEND_BASE_PORT: u16 = 3002;

/// Default number of soffice instances in the container.
///
/// Concurrency has to come from separate *processes*, not separate
/// connections: soffice guards its document model with one process-global
/// SolarMutex, so N connections to one instance serialize (measured 3%
/// *slower* than serial), while N instances scale linearly (measured 8x
/// with zero failures). Each instance needs its own
/// `-env:UserInstallation` profile, otherwise a second soffice detects the
/// first, hands off to it, and exits - leaving one instance and no
/// parallelism at all.
const DEFAULT_INSTANCES: u16 = 8;

/// Number of soffice instances to run, overridable via
/// `DUKE_LO_INSTANCES`. Clamped to at least 1.
pub fn instances() -> u16 {
    std::env::var("DUKE_LO_INSTANCES")
        .ok()
        .and_then(|s| s.parse::<u16>().ok())
        .unwrap_or(DEFAULT_INSTANCES)
        .max(1)
}



/// Base time to wait for URP sockets to come up after `docker run`. The
/// real budget scales with instance count, since the launches are
/// staggered and N instances contend for CPU while bootstrapping.
const STARTUP_TIMEOUT_BASE: Duration = Duration::from_secs(30);

/// Total budget for all instances to start listening.
fn startup_timeout() -> Duration {
    STARTUP_TIMEOUT_BASE + Duration::from_secs(5 * instances() as u64)
}

static ENSURED: OnceLock<Result<(), String>> = OnceLock::new();

/// Ensure a LibreOffice URP container is running and reachable on
/// localhost:2002. Builds the docker image first if it's missing.
///
/// **Reuses a healthy container.** If URP answers the read probe, this
/// returns immediately; otherwise it builds the image (if missing),
/// sweeps stale locks, and starts a fresh container. Subsequent calls
/// within the same process are O(1) no-ops via the `OnceLock`.
///
/// This used to replace the container unconditionally on first call in
/// every test process, to guard against soffice state accumulating across
/// document opens. That cost ~4.5s per test binary (sweep container,
/// `rm -f`, `run`, port wait) across ~30 LO-backed binaries, and it was
/// papering over the wrong layer: the container runs a persistent server,
/// so if it degrades under sustained use that is a bug in the container,
/// not something the harness should hide by restarting. `duke-sheets-xls`
/// `e2e` already drives 68 consecutive document operations against one
/// instance, so the reuse path is well within proven territory.
///
/// If a container does go bad, `docker rm -f duke-lo-test` forces a fresh
/// one on the next call.
///
/// Once the URP port is listening this returns; the
/// `LibreOfficeBridge::open_workbook` / `create_workbook` calls in
/// `duke-sheets-libreoffice` retry on `loadComponentFromURL returned
/// null`, which absorbs the loader-vs-listener startup race.
///
/// Panics with a descriptive message on failure.
pub fn ensure_lo() {
    // The failure outcome is cached: a panicking initializer would
    // leave the OnceLock uninitialized and every subsequent LO-backed
    // test would re-pay the startup timeout. The first attempt's error
    // is replayed instantly instead — still a hard failure, never a
    // skip.
    let outcome = ENSURED.get_or_init(|| {
        std::panic::catch_unwind(|| {
            let _ = std::fs::create_dir_all(SHARED_DIR);
            if port_listening(URP_PORT) {
                return;
            }
            ensure_image();
            sweep_stale_locks();
            ensure_container();
        })
        .map_err(|e| {
            e.downcast_ref::<String>()
                .cloned()
                .or_else(|| e.downcast_ref::<&str>().map(|s| s.to_string()))
                .unwrap_or_else(|| "LibreOffice startup panicked".to_string())
        })
    });
    if let Err(msg) = outcome {
        panic!("{msg}\n(cached from the first startup attempt in this process)");
    }
}

/// Probe that the URP server is genuinely ready, not just that the
/// docker port forward is up. Docker publishes :2002 instantly on
/// `docker run`, but the soffice process inside the container takes a
/// few seconds to bind that port for real. A pure connect-only check
/// passes too early; subsequent URP I/O then fails with
/// `Broken pipe`.
///
/// Strategy: open a connection, set a 200ms read timeout, attempt a
/// `read`. The URP server doesn't speak first, so a healthy connection
/// hits the read timeout (`WouldBlock`/`TimedOut`). An unhealthy
/// connection (port forward to nothing) returns `Ok(0)` (EOF) or an
/// error immediately.
fn port_listening(port: u16) -> bool {
    let addr = match format!("127.0.0.1:{port}").parse() {
        Ok(a) => a,
        Err(_) => return false,
    };
    let Ok(mut stream) = TcpStream::connect_timeout(&addr, Duration::from_secs(2)) else {
        return false;
    };
    if stream
        .set_read_timeout(Some(Duration::from_millis(200)))
        .is_err()
    {
        return false;
    }
    let mut buf = [0u8; 1];
    match stream.read(&mut buf) {
        Ok(0) => false,
        Ok(_) => true,
        Err(e) => matches!(
            e.kind(),
            std::io::ErrorKind::WouldBlock | std::io::ErrorKind::TimedOut
        ),
    }
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

/// Bring the container up, tolerating concurrent callers.
///
/// Test binaries can run concurrently, so several processes may find the
/// container missing at the same moment. `docker run --name` is atomic in
/// the daemon - exactly one concurrent create wins - so it is used as the
/// mutex rather than a lock file. Losers wait for the winner's container
/// to start listening; only if that never happens is the existing
/// container treated as broken and replaced.
fn ensure_container() {
    if try_create_container() {
        wait_for_port(URP_PORT, startup_timeout());
        return;
    }

    // Lost the race, or a broken container holds the name. Give the
    // presumed owner a chance to finish booting.
    if port_up_within(URP_PORT, Duration::from_secs(20)) {
        return;
    }

    // Nothing is listening, so whatever holds the name is broken. Remove
    // it and race again; a concurrent winner's create makes ours fail,
    // which is fine since both then wait on the same port.
    let _ = Command::new("docker")
        .args(["rm", "-f", CONTAINER_NAME])
        .output();
    for cid in containers_publishing(URP_PORT) {
        let _ = Command::new("docker").args(["rm", "-f", &cid]).output();
    }
    try_create_container();
    wait_for_port(URP_PORT, startup_timeout());
}

/// Container ids publishing `port`, so a foreign container squatting on
/// our range can be cleared.
fn containers_publishing(port: u16) -> Vec<String> {
    Command::new("docker")
        .args(["ps", "--filter", &format!("publish={port}"), "--quiet"])
        .output()
        .ok()
        .map(|out| {
            String::from_utf8_lossy(&out.stdout)
                .split_whitespace()
                .map(str::to_string)
                .collect()
        })
        .unwrap_or_default()
}

/// Poll a port without panicking. Used to wait out a concurrent creator
/// before concluding its container is broken.
fn port_up_within(port: u16, timeout: Duration) -> bool {
    let start = Instant::now();
    while start.elapsed() < timeout {
        if port_listening(port) {
            return true;
        }
        std::thread::sleep(Duration::from_millis(250));
    }
    false
}

/// Start the container: haproxy on the published port, `instances()`
/// supervised soffice backends on internal ports. Returns false if the
/// name is already taken (i.e. another process won the race).
fn try_create_container() -> bool {
    let n = instances();

    // Each backend runs under a restart loop, so a crashed soffice comes
    // back and haproxy routes around it while it boots. Backends are
    // staggered: starting them all at once makes some die during first-run
    // profile bootstrap.
    let mut script = String::new();
    for i in 0..n {
        let p = BACKEND_BASE_PORT + i;
        script.push_str(&format!(
            "(while true; do soffice --headless \
             -env:UserInstallation=file:///tmp/lo-prof-{p} \
             --accept=\"socket,host=127.0.0.1,port={p};urp;StarOffice.ComponentContext\"; \
             sleep 1; done) & sleep 1; "
        ));
    }
    script.push_str(&format!(
        "cat > /tmp/haproxy.cfg <<EOF\n\
         defaults\n  mode tcp\n  timeout connect 10s\n  timeout client 1h\n  timeout server 1h\n\
         listen lo\n  bind 0.0.0.0:{URP_PORT}\n  balance leastconn\n"
    ));
    for i in 0..n {
        let p = BACKEND_BASE_PORT + i;
        script.push_str(&format!("  server s{i} 127.0.0.1:{p} check\n"));
    }
    script.push_str("EOF\nexec haproxy -f /tmp/haproxy.cfg");

    let out = Command::new("docker")
        .args([
            "run",
            "-d",
            "--rm",
            "--name",
            CONTAINER_NAME,
            "-p",
            &format!("{URP_PORT}:{URP_PORT}"),
            "-v",
            &format!("{SHARED_DIR}:{SHARED_DIR}"),
            IMAGE_TAG,
            "bash",
            "-c",
            &script,
        ])
        .output()
        .unwrap_or_else(|e| panic!("failed to spawn `docker run` for {CONTAINER_NAME}: {e}"));
    out.status.success()
}

/// Wait until the proxy answers a real read probe. haproxy accepts and
/// immediately closes while no backend is healthy, which the probe reads
/// as EOF = not ready, so this also waits for the first live backend.
fn wait_for_port(port: u16, timeout: Duration) {
    let start = Instant::now();
    loop {
        if port_listening(port) {
            eprintln!(
                "[duke-sheets-test-harness] LibreOffice URP proxy ready on port {port} \
                 ({} backends; took {:.1}s)",
                instances(),
                start.elapsed().as_secs_f64()
            );
            return;
        }
        if start.elapsed() >= timeout {
            panic!(
                "LibreOffice URP did not come up on localhost:{port} within {timeout:?}.\n\
                 Inspect with: docker logs {CONTAINER_NAME}"
            );
        }
        std::thread::sleep(Duration::from_millis(250));
    }
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
