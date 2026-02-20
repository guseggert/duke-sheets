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
#   DUKE_SHEETS_VM_DISK  Path to Windows qcow2 disk (default: ~/.duke-sheets/windows11.qcow2)
#   DUKE_SHEETS_VM_RAM   RAM in GB (default: 4)
#   DUKE_SHEETS_VM_CPUS  Number of CPUs (default: 2)

set -euo pipefail

DISK="${DUKE_SHEETS_VM_DISK:-$HOME/.duke-sheets/windows11.qcow2}"
RAM="${DUKE_SHEETS_VM_RAM:-4}"
CPUS="${DUKE_SHEETS_VM_CPUS:-2}"
PID_FILE="/tmp/duke-sheets-vm.pid"
SHARE_DIR="/tmp/duke-sheets-excel"

if [ ! -f "$DISK" ]; then
    echo "ERROR: VM disk not found at $DISK"
    echo ""
    echo "To set up the VM, see tools/vm/README.md"
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
echo "  RAM:    ${RAM}G"
echo "  CPUs:   $CPUS"
echo "  Share:  $SHARE_DIR -> \\\\10.0.2.4\\qemu"
echo "  Ports:  9876 (bridge), 5985 (WinRM), 2222 (SSH)"

qemu-system-x86_64 \
    -accel kvm \
    -cpu host \
    -m "${RAM}G" \
    -smp "$CPUS" \
    -drive "file=$DISK,format=qcow2,if=virtio" \
    -nic "user,hostfwd=tcp::9876-:9876,hostfwd=tcp::5985-:5985,hostfwd=tcp::2222-:22,smb=$SHARE_DIR" \
    -display none \
    -daemonize \
    -pidfile "$PID_FILE" \
    -monitor "unix:/tmp/duke-sheets-vm.sock,server,nowait"

echo "VM started (PID $(cat "$PID_FILE"))"
echo ""
echo "Waiting for bridge server on port 9876..."

for i in $(seq 1 60); do
    if nc -z localhost 9876 2>/dev/null; then
        echo "Bridge server ready (took ${i}s)"
        exit 0
    fi
    sleep 1
done

echo "WARNING: Bridge server not responding after 60s."
echo "The VM may still be booting. Check with: nc -z localhost 9876"
