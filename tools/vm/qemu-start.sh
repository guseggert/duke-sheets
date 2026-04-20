#!/bin/bash
# Start the Windows 11 VM with KVM acceleration.
#
# Port forwarding:
#   9876 -> 9876  Excel COM bridge server
#   5985 -> 5985  WinRM (for remote management)
#   2222 -> 22    SSH (OpenSSH server)
#
# SMB share:
#   /tmp/duke-sheets-excel/ is accessible inside the VM as \\10.0.2.4\qemu
#
# Environment variables:
#   DUKE_SHEETS_VM_DIR   Base directory (default: ~/.duke-sheets)
#   DUKE_SHEETS_VM_DISK  Path to Windows qcow2 disk (default: $VM_DIR/windows11.qcow2)
#   DUKE_SHEETS_VM_RAM   RAM (default: 4G)
#   DUKE_SHEETS_VM_CPUS  Number of CPUs (default: 2)

set -euo pipefail

DUKE_DIR="${DUKE_SHEETS_VM_DIR:-$HOME/.duke-sheets}"
DISK="${DUKE_SHEETS_VM_DISK:-$DUKE_DIR/windows11.qcow2}"
RAM="${DUKE_SHEETS_VM_RAM:-4G}"
CPUS="${DUKE_SHEETS_VM_CPUS:-2}"
PID_FILE="/tmp/duke-sheets-vm.pid"
SHARE_DIR="/tmp/duke-sheets-excel"

# Find the repo root (where this script lives is tools/vm/)
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_DIR="$(cd "$SCRIPT_DIR/../.." && pwd)"

# Search order: .qemu/ in repo, ~/.duke-sheets/qemu/, PATH, /tmp build
QEMU=""
for candidate in \
    "$REPO_DIR/.qemu/qemu-system-x86_64" \
    "$DUKE_DIR/qemu/qemu-system-x86_64" \
; do
    if [ -x "$candidate" ]; then
        QEMU="$candidate"
        break
    fi
done
if [ -z "$QEMU" ]; then
    if command -v qemu-system-x86_64 &>/dev/null; then
        QEMU="qemu-system-x86_64"
    else
        echo "ERROR: QEMU not found. Run: mise run vm:build-qemu"
        exit 1
    fi
fi

if [ ! -f "$DISK" ]; then
    echo "ERROR: VM disk not found at $DISK"
    echo ""
    echo "To set up the VM, run: bash tools/vm/setup.sh"
    echo "Or set DUKE_SHEETS_VM_DISK to point to your Windows qcow2 image."
    exit 1
fi

if [ -f "$PID_FILE" ] && kill -0 "$(cat "$PID_FILE")" 2>/dev/null; then
    echo "VM is already running (PID $(cat "$PID_FILE"))"
    exit 0
fi

# Ensure shared directory exists
mkdir -p "$SHARE_DIR"

echo "Starting Windows VM..."
echo "  Disk:   $DISK"
echo "  RAM:    $RAM"
echo "  CPUs:   $CPUS"
echo "  Share:  $SHARE_DIR -> \\\\10.0.2.4\\qemu"
echo "  Ports:  9876 (bridge), 5985 (WinRM), 2222 (SSH)"

export LD_LIBRARY_PATH="/usr/local/lib64:/usr/local/lib:${LD_LIBRARY_PATH:-}"

# Resolve BIOS/firmware data dir from the QEMU binary location.
# For repo-local builds (.qemu/), the pc-bios files are in the build tree.
QEMU_DIR="$(dirname "$QEMU")"
QEMU_DATA_DIR=""
if [ -d "$QEMU_DIR/build" ]; then
    # Repo-local build: find pc-bios in the source tree
    for d in "$QEMU_DIR"/build/qemu-*/pc-bios; do
        if [ -f "$d/bios-256k.bin" ]; then
            QEMU_DATA_DIR="$d"
            break
        fi
    done
fi

"$QEMU" \
    ${QEMU_DATA_DIR:+-L "$QEMU_DATA_DIR"} \
    -M q35,usb=on,acpi=on,hpet=off \
    -accel kvm \
    -cpu host,hv_relaxed,hv_frequencies,hv_vpindex,hv_ipi,hv_tlbflush,hv_spinlocks=0x1fff,hv_synic,hv_runtime,hv_time,hv_stimer,hv_vapic \
    -m "$RAM" \
    -smp "cores=$CPUS" \
    -drive "file=$DISK" \
    -device usb-tablet \
    -device VGA,vgamem_mb=256 \
    -nic "user,model=e1000,hostfwd=tcp::9876-:9876,hostfwd=tcp::5985-:5985,hostfwd=tcp::2222-:22,smb=$SHARE_DIR" \
    -display none \
    -daemonize \
    -pidfile "$PID_FILE" \
    -monitor "unix:/tmp/duke-sheets-vm.sock,server,nowait"

echo "VM started (PID $(cat "$PID_FILE"))"
echo ""
echo "Waiting for WinRM on port 5985 (VM boot)..."

# Wait for WinRM to be reachable - this means Windows has finished booting
# and the network stack is up. Note: nc -z gives false positives with QEMU
# user-net (the host-side port is always open), so we send a real HTTP request.
for i in $(seq 1 180); do
    if curl -s -o /dev/null -w '%{http_code}' --max-time 2 \
         http://localhost:5985/wsman 2>/dev/null | grep -q '40[0-9]'; then
        echo "WinRM reachable (took ${i}s)"
        echo ""
        echo "Bridge should be running on port 9876."
        echo "Test with:"
        echo "  echo '{\"id\":1,\"cmd\":\"Init\"}' | ncat -w5 localhost 9876"
        exit 0
    fi
    sleep 1
done

echo "WARNING: WinRM not responding after 180s."
echo "The VM may still be booting. Try:"
echo "  curl -v http://localhost:5985/wsman"
