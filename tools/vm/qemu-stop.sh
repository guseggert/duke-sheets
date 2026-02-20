#!/bin/bash
# Gracefully shut down the Windows VM.
#
# Sends an ACPI power button press via the QEMU monitor socket,
# waits for the VM to shut down, then kills it if it takes too long.

set -euo pipefail

PID_FILE="/tmp/duke-sheets-vm.pid"
MONITOR_SOCK="/tmp/duke-sheets-vm.sock"

if [ ! -f "$PID_FILE" ]; then
    echo "VM is not running (no PID file)"
    exit 0
fi

PID=$(cat "$PID_FILE")

if ! kill -0 "$PID" 2>/dev/null; then
    echo "VM process ($PID) not found, cleaning up PID file."
    rm -f "$PID_FILE"
    exit 0
fi

echo "Sending ACPI shutdown to VM (PID $PID)..."

# Send ACPI powerdown via QEMU monitor
if [ -S "$MONITOR_SOCK" ]; then
    echo "system_powerdown" | socat - "UNIX-CONNECT:$MONITOR_SOCK" 2>/dev/null || true
fi

# Wait up to 30 seconds for graceful shutdown
echo "Waiting for VM to shut down..."
for i in $(seq 1 30); do
    if ! kill -0 "$PID" 2>/dev/null; then
        echo "VM shut down cleanly (took ${i}s)"
        rm -f "$PID_FILE"
        exit 0
    fi
    sleep 1
done

echo "VM did not shut down in 30s, killing..."
kill -9 "$PID" 2>/dev/null || true
rm -f "$PID_FILE"
echo "VM killed."
