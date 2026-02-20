# Windows VM Setup for Excel COM Bridge

This directory contains scripts and configuration for running Microsoft Excel
inside a QEMU/KVM virtual machine, controlled by the `duke-sheets-excel-com`
crate for parity testing.

## Architecture

```
Linux Host                              QEMU/KVM Windows 11 VM
┌────────────────────────┐   TCP:9876   ┌───────────────────────────┐
│ cargo test             │─────────────►│ ExcelBridgeServer.exe     │
│ duke-sheets-excel-com  │◄─────────────│ (C# .NET 8, ~150 lines)  │
│ (Rust, TCP client)     │   NDJSON     │      ↓ COM (dynamic)     │
│                        │              │ Excel.Application         │
│                        │   SMB share  │ (Office 365, latest)      │
│ /tmp/duke-sheets-excel/├─────────────►│ \\10.0.2.4\qemu\         │
└────────────────────────┘              └───────────────────────────┘
```

## Prerequisites

- **QEMU/KVM** with KVM acceleration: `sudo apt install qemu-system-x86 qemu-utils`
- **socat** for VM shutdown: `sudo apt install socat`
- **.NET 8 SDK** on Linux for cross-compiling the bridge: `sudo apt install dotnet-sdk-8.0`
- **Windows 11 Enterprise Evaluation ISO** (free, 90-day):
  https://www.microsoft.com/en-us/evalcenter/evaluate-windows-11-enterprise

## One-Time Setup

### 1. Create the VM disk

```bash
mkdir -p ~/.duke-sheets
qemu-img create -f qcow2 ~/.duke-sheets/windows11.qcow2 60G
```

### 2. Install Windows

```bash
# Download virtio drivers for best performance:
# https://fedorapeople.org/groups/virt/virtio-win/direct-downloads/stable-virtio/virtio-win.iso

qemu-system-x86_64 \
  -accel kvm -cpu host -m 4G -smp 2 \
  -drive file=~/.duke-sheets/windows11.qcow2,format=qcow2,if=virtio \
  -cdrom /path/to/Win11_Enterprise_Eval.iso \
  -drive file=/path/to/virtio-win.iso,media=cdrom,index=1 \
  -boot order=d \
  -nic user \
  -display gtk
```

During Windows setup, if the disk isn't visible, load the virtio storage
driver from the second CD (virtio-win > vioscsi > w11 > amd64).

### 3. Install Office (Excel only)

Inside the VM:

1. Download the Office Deployment Tool (ODT) from Microsoft
2. Extract it, then copy `tools/vm/install-office.xml` into the same folder
3. Run: `setup.exe /download install-office.xml`
4. Run: `setup.exe /configure install-office.xml`

This installs only Excel, minimizing disk usage.

### 4. Build and deploy the bridge server

From the Linux host:

```bash
# Cross-compile the C# bridge for Windows
cd tools/excel-bridge-server
dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true

# Copy to the shared directory
cp bin/Release/net8.0/win-x64/publish/ExcelBridgeServer.exe /tmp/duke-sheets-excel/
```

Inside the VM, copy from `\\10.0.2.4\qemu\ExcelBridgeServer.exe` to `C:\tools\`.

### 5. Configure auto-start

Run `tools/vm/setup-bridge.ps1` inside the VM (as admin). This:
- Creates a firewall rule for port 9876
- Creates a scheduled task to start the bridge on login
- Enables WinRM and OpenSSH for remote management

### 6. Snapshot the VM

```bash
# Create a clean snapshot to revert to
qemu-img snapshot -c clean-install ~/.duke-sheets/windows11.qcow2
```

## Daily Usage

```bash
# Start the VM
mise run vm:start
# or: bash tools/vm/qemu-start.sh

# Run parity tests
cargo run --example parity_test -p duke-sheets-excel-com

# Stop the VM
mise run vm:stop
# or: bash tools/vm/qemu-stop.sh
```

## Shared Files

The directory `/tmp/duke-sheets-excel/` on the host is accessible inside the
VM as `\\10.0.2.4\qemu\` (via QEMU's built-in SMB server). Test files saved
by Excel appear on the host immediately.

## Troubleshooting

**VM won't start**: Check KVM support: `kvm-ok` or `lsmod | grep kvm`

**Bridge not responding**: SSH into the VM and check:
```bash
ssh -p 2222 user@localhost
# Inside VM: check if bridge is running
tasklist | findstr ExcelBridge
# Start manually:
C:\tools\ExcelBridgeServer.exe --port 9876
```

**Excel COM errors**: Make sure Excel has been launched at least once manually
(to complete first-run setup / license acceptance).

**Snapshot revert**:
```bash
qemu-img snapshot -a clean-install ~/.duke-sheets/windows11.qcow2
```
