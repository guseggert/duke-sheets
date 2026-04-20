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

Port forwarding (host → VM):

| Host Port | VM Port | Service |
|-----------|---------|---------|
| 9876      | 9876    | Excel COM bridge (NDJSON-over-TCP) |
| 5985      | 5985    | WinRM (remote management) |
| 2222      | 22      | SSH (OpenSSH server, optional) |

## Prerequisites

- **KVM support** - `/dev/kvm` must exist. On EC2 use a `.metal` instance.
- **sudo access** - needed for installing packages, mounting floppy images.
- **Windows 11 ISO** - consumer or enterprise, any edition. Set `WIN_ISO`
  env var or place at `~/Win11.iso`.
- **~20 GB free disk space** - for QEMU build + Windows qcow2 disk.

On Amazon Linux 2023, QEMU and libslirp are not in repos and must be built
from source. The `setup.sh` script handles this automatically.

## Quick Start

```bash
# One-shot: build QEMU, create disk, install Windows unattended
bash tools/vm/setup.sh

# Wait ~10-15 minutes for install to complete, then:
bash tools/vm/qemu-start.sh   # boot the installed VM (no CD)
bash tools/vm/qemu-stop.sh    # graceful shutdown
```

## How It Works (Lessons Learned)

### BIOS mode, not UEFI

Windows 11 officially requires UEFI + TPM 2.0 + Secure Boot. However,
these requirements can all be bypassed via registry keys in `autounattend.xml`:

```
BypassTPMCheck, BypassSecureBootCheck, BypassRAMCheck, BypassCPUCheck
```

We boot in **legacy BIOS mode (SeaBIOS)** rather than UEFI (OVMF) because:

1. OVMF UEFI CD-ROM boot is broken/unreliable with QEMU's IDE, SCSI, and
   USB CD emulation - we consistently hit "Time out" errors trying to boot
   the Windows ISO across multiple OVMF builds (retrage nightly, Fedora 41
   edk2-ovmf) and device types (IDE, AHCI, virtio-scsi, USB storage).
2. SeaBIOS "just works" with `-cdrom` and a standard consumer Win11 ISO.
3. Windows 11 installs and runs fine in BIOS/MBR mode once the registry
   bypasses are in place.

### QEMU invocation

Based on the [Computernewb guide](https://computernewb.com/wiki/QEMU/Guests/Windows_11).
Key flags:

```bash
qemu-system-x86_64 \
    -M q35,usb=on,acpi=on,hpet=off \
    -accel kvm \
    -cpu host,hv_relaxed,hv_frequencies,hv_vpindex,hv_ipi,hv_tlbflush,\
hv_spinlocks=0x1fff,hv_synic,hv_runtime,hv_time,hv_stimer,hv_vapic \
    -m 8G -smp cores=4 \
    -drive file=windows11.qcow2 \
    -device usb-tablet \
    -device VGA,vgamem_mb=256 \
    -nic user,model=e1000,...
```

- **`-M q35`** - modern PCIe chipset, best Windows compatibility.
- **`-cpu host,hv_*`** - Hyper-V enlightenments. Windows detects these and
  uses optimized code paths for timers, IPIs, TLB flushes, etc. Significant
  performance improvement over plain `-cpu host`.
- **`-device usb-tablet`** - absolute pointing device, avoids mouse capture.
- **`-device VGA,vgamem_mb=256`** - more VRAM for the Win11 desktop.
- **`-nic user,model=e1000`** - user-mode networking with Intel e1000 NIC
  (has in-box Windows driver, no virtio drivers needed).
- **No OVMF/pflash** - SeaBIOS is built into QEMU, nothing extra needed.
- **No virtio drivers** - we use default IDE disk and e1000 NIC. Virtio
  would be faster but requires installing guest drivers, which adds
  complexity for minimal benefit in a test VM.

### Unattended Install (autounattend.xml)

The `autounattend.xml` is placed on a FAT12 floppy image and attached via
`-drive file=autounattend.img,format=raw,if=floppy`. Windows Setup
automatically finds it.

#### Multi-edition ISO handling

Consumer Win11 ISOs contain 11 editions (Home, Pro, Education, etc.).
Without specifying which one, the installer prompts for selection and
blocks the unattended flow. Fix: add `<InstallFrom><MetaData>` to select
by name:

```xml
<ImageInstall>
  <OSImage>
    <InstallFrom>
      <MetaData wcm:action="add">
        <Key>/IMAGE/NAME</Key>
        <Value>Windows 11 Pro</Value>
      </MetaData>
    </InstallFrom>
    <InstallTo><DiskID>0</DiskID><PartitionID>1</PartitionID></InstallTo>
  </OSImage>
</ImageInstall>
```

You also need a generic product key for the edition. For Pro:
`W269N-WFGWX-YVC9B-4J6C9-T83GX` (well-known KMS/generic key - allows
install without activation).

To check which editions are in your ISO:

```bash
sudo mount -o loop,ro Win11.iso /mnt/iso
# Then parse the WIM XML metadata (at end of install.wim)
python3 -c "
import struct, re
with open('/mnt/iso/sources/install.wim', 'rb') as f:
    f.seek(0, 2); fsize = f.tell()
    f.seek(max(0, fsize - 2*1024*1024))
    data = f.read()
    idx = data.find(b'<\x00W\x00I\x00M\x00>\x00')
    if idx >= 0:
        end = data.find(b'<\x00/\x00W\x00I\x00M\x00>\x00', idx)
        text = data[idx:end].decode('utf-16-le')
        for m in re.finditer(r'<IMAGE INDEX=\"(\d+)\".*?<NAME>(.*?)</NAME>', text, re.DOTALL):
            print(f'  Index {m.group(1)}: {m.group(2)}')
"
```

#### MBR partitioning (not GPT)

Since we're in BIOS mode, the disk must use MBR. Create a single active
Primary partition - Windows will set up its own boot files:

```xml
<DiskConfiguration>
  <Disk wcm:action="add">
    <DiskID>0</DiskID>
    <WillWipeDisk>true</WillWipeDisk>
    <CreatePartitions>
      <CreatePartition wcm:action="add">
        <Order>1</Order>
        <Extend>true</Extend>
        <Type>Primary</Type>
      </CreatePartition>
    </CreatePartitions>
    <ModifyPartitions>
      <ModifyPartition wcm:action="add">
        <Order>1</Order>
        <PartitionID>1</PartitionID>
        <Format>NTFS</Format>
        <Label>Windows</Label>
        <Letter>C</Letter>
        <Active>true</Active>
      </ModifyPartition>
    </ModifyPartitions>
  </Disk>
</DiskConfiguration>
```

Do **not** create EFI or MSR partitions in BIOS mode - the installer
will show "Windows 11 can't be installed" if you do.

#### Passwords

Use `<PlainText>true</PlainText>` with the literal password string.
The `PlainText=false` mode requires base64-encoding of
`UTF16LE(password + "Password")`, which is error-prone. Plaintext works
fine for a local test VM:

```xml
<Password>
  <Value>test</Value>
  <PlainText>true</PlainText>
</Password>
```

#### specialize vs FirstLogonCommands

The `specialize` pass runs as **SYSTEM** - use it for registry
settings, firewall rules, and anything that doesn't need the network
stack running. **Do not use specialize for WinRM** - `winrm
quickconfig` fails there because the network profile isn't set yet.

WinRM setup goes in `FirstLogonCommands` (oobeSystem pass). UAC
would normally block admin commands there, so we disable it in
specialize via `EnableLUA=0`:

```xml
<!-- In specialize: disable UAC, set registry, firewall -->
<Path>reg add HKLM\...\Policies\System /v EnableLUA /t REG_DWORD /d 0 /f</Path>
<Path>netsh advfirewall firewall add rule name="ExcelBridge" ...</Path>

<!-- In FirstLogonCommands: network-dependent WinRM setup -->
<CommandLine>powershell -Command "Get-NetConnectionProfile |
    Set-NetConnectionProfile -NetworkCategory Private"</CommandLine>
<CommandLine>powershell -Command "Enable-PSRemoting -Force
    -SkipNetworkProfileCheck"</CommandLine>
<CommandLine>cmd /c winrm set winrm/config/service
    @{AllowUnencrypted="true"}</CommandLine>
```

#### QEMU NIC network profile is Public

The QEMU e1000 NIC is classified as a **Public** network by Windows,
which blocks `winrm quickconfig`. Fix: set the profile to Private
before configuring WinRM:

```powershell
Get-NetConnectionProfile | Set-NetConnectionProfile -NetworkCategory Private
```

The `-SkipNetworkProfileCheck` flag on `Enable-PSRemoting` also
helps as a fallback.

#### SMB guest access

QEMU's built-in SMB server (`-nic user,...,smb=/path`) uses guest
(unauthenticated) access. Windows 11 blocks this by default. Fix via
registry (in specialize pass):

```
reg add HKLM\SOFTWARE\Policies\Microsoft\Windows\LanmanWorkstation
        /v AllowInsecureGuestAuth /t REG_DWORD /d 1 /f
```

#### OOBE internet bypass

Windows 11 OOBE requires a Microsoft account (internet connection).
Bypass with `BypassNRO` registry key (in specialize pass):

```
reg add HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE
        /v BypassNRO /t REG_DWORD /d 1 /f
```

### Floppy image creation

```bash
dd if=/dev/zero of=autounattend.img bs=1440k count=1
mkfs.fat autounattend.img
MNT=$(mktemp -d)
sudo mount -o loop autounattend.img "$MNT"
sudo cp autounattend.xml "$MNT/autounattend.xml"
sudo umount "$MNT"
rmdir "$MNT"
```

### Building QEMU from source (Amazon Linux 2023)

AL2023 doesn't have `qemu-system-x86_64` or `libslirp-devel` in repos.
Both must be built from source:

```bash
# System deps
sudo dnf install -y gcc make ninja-build glib2-devel pixman-devel \
    zlib-devel dosfstools samba
pip install distlib meson

# libslirp (user-mode networking)
git clone --depth=1 https://gitlab.freedesktop.org/slirp/libslirp.git
cd libslirp && meson setup build && ninja -C build
sudo ninja -C build install && sudo ldconfig

# QEMU (x86_64 system emulator only)
curl -LO https://download.qemu.org/qemu-9.2.3.tar.xz
tar xf qemu-9.2.3.tar.xz && cd qemu-9.2.3
PKG_CONFIG_PATH="/usr/local/lib64/pkgconfig" \
    ./configure --target-list=x86_64-softmmu \
    --disable-gtk --disable-sdl --enable-vnc --enable-slirp --disable-docs
ninja -C build qemu-system-x86_64 qemu-img
```

At runtime, set `LD_LIBRARY_PATH="/usr/local/lib64"` so QEMU finds libslirp.

### OVMF/UEFI: why we don't use it

We tried UEFI boot extensively with two OVMF sources:
- retrage.github.io nightly builds (2MB CODE+VARS)
- Fedora 41 edk2-ovmf package (2MB raw, 4MB qcow2)

Both consistently failed with "BdsDxe: failed to start Boot... Time out"
on every device type:
- `-cdrom` (IDE ATAPI)
- `-device ide-cd` (explicit IDE)
- `-device scsi-cd` (virtio-scsi)
- `-device usb-storage` (xHCI USB)
- With `cache=unsafe,aio=threads`
- On both `pc` (i440fx) and `q35` machine types

The ISO reads at 6.8 GB/s from the host side, so it's not an I/O issue.
The UEFI firmware itself times out reading the boot loader from the
CD/USB device. This appears to be a QEMU 9.2.3 regression or an
interaction with the specific ISO format.

We also tried creating a bootable GPT disk image with a FAT32 EFI
partition + NTFS data partition (using Docker + ntfs-3g to handle
install.wim > 4GB). This partially worked - UEFI found the Windows
Boot Manager but failed with "BCD missing" until we copied the full
`\EFI\Microsoft\Boot\` directory. At that point the VM crashed on
boot.

**Conclusion**: BIOS/SeaBIOS is dramatically simpler and works
immediately. UEFI provides no benefit for a headless test VM.

## Install Office (Excel only)

Office is installed inside the VM via the Office Deployment Tool (ODT),
which downloads from the CDN. The VM has internet access through QEMU
user-mode networking.

**IMPORTANT**: Offline installer ISOs (like `HomeBusinessRetail.img`)
are old Click-to-Run media whose format is incompatible with current
ODT. Don't try to use them - just let ODT download fresh.

From a WinRM session (or the helper script `winrm-exec.py`):

```powershell
# Download ODT
Invoke-WebRequest -Uri "https://officecdn.microsoft.com/pr/wsus/setup.exe" `
    -OutFile C:\OfficeSetup\odt-setup.exe -UseBasicParsing

# Write config (Excel only, no other apps)
@"
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365BusinessRetail">
      <Language ID="en-us" />
      <ExcludeApp ID="Access" /><ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" /><ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="OneNote" /><ExcludeApp ID="Outlook" />
      <ExcludeApp ID="PowerPoint" /><ExcludeApp ID="Publisher" />
      <ExcludeApp ID="Teams" /><ExcludeApp ID="Word" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="0" />
  <Updates Enabled="FALSE" />
</Configuration>
"@ | Set-Content C:\OfficeSetup\config.xml

# Install (downloads ~2GB, takes a few minutes, runs silently in background)
C:\OfficeSetup\odt-setup.exe /configure C:\OfficeSetup\config.xml
```

The ODT returns immediately; the ClickToRunSvc service does the actual
install in the background. Monitor progress:

```powershell
# Poll until Excel appears (~3-5 minutes)
while (-not (Test-Path "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE")) {
    $size = (Get-ChildItem "C:\Program Files\Microsoft Office" -Recurse -EA 0 |
        Measure-Object -Property Length -Sum).Sum
    Write-Host "$([math]::Round($size/1MB)) MB..."
    Start-Sleep 10
}
Write-Host "Excel installed!"
```

## Build and Deploy the C# Bridge

From the Linux host:

```bash
# Install .NET 8 SDK if needed
curl -sSL https://dot.net/v1/dotnet-install.sh | bash -s -- --channel 8.0

# Cross-compile for Windows (self-contained, no .NET needed on target)
cd tools/excel-bridge-server
~/.dotnet/dotnet publish -c Release -r win-x64 --self-contained \
    -p:PublishSingleFile=true

# Copy to SMB share → VM picks it up from \\10.0.2.4\qemu\
cp bin/Release/net8.0/win-x64/publish/ExcelBridgeServer.exe \
    /tmp/duke-sheets-excel/
```

### Running the bridge server (CRITICAL: must be interactive session)

The bridge server **must run in the logged-in user's desktop session**
(Session 1), not from WinRM (Session 0). Excel COM automation requires
an interactive desktop - it will hang or crash in a service session.

Use a scheduled task with `LogonType Interactive`:

```powershell
# Via WinRM - creates a task that runs in the desktop session
Copy-Item "\\10.0.2.4\qemu\ExcelBridgeServer.exe" C:\ExcelBridgeServer.exe
$action = New-ScheduledTaskAction -Execute "C:\ExcelBridgeServer.exe"
$principal = New-ScheduledTaskPrincipal -UserId "user" `
    -LogonType Interactive -RunLevel Highest
Register-ScheduledTask -TaskName "ExcelBridge" `
    -Action $action -Principal $principal -Force
Start-ScheduledTask -TaskName "ExcelBridge"
```

Or use `tools/vm/setup-bridge.ps1` which does this automatically.

### C# bridge implementation notes

These were hard-won and cost significant debugging time:

- **`[STAThread]` is required** on `Main`. COM interop needs an STA
  (Single-Threaded Apartment). Without it, `Activator.CreateInstance`
  for `Excel.Application` crashes with an unhandled CLR exception.

- **Do NOT enable IL trimming** (`PublishTrimmed=false`). The trimmer
  strips COM interop and `dynamic` dispatch code. Even with
  `BuiltInComInteropSupport=true` and `TrimmerRootAssembly` hints,
  the trimmed binary fails at runtime: "Built-in COM has been
  disabled via a feature switch." The untrimmed exe is ~65MB vs
  ~14MB trimmed, but 65MB is fine for a test VM.

- **Use reflection-based JSON serializer**, not source-generated
  (`JsonSerializerContext`). Source-gen can't polymorphically
  serialize `object?` properties - a `ValueData { Value = 3.0 }`
  silently serializes as `{}` instead of `{"value": 3}`.

- **Tuple deconstruction fails with `dynamic`**. If a method takes
  `dynamic` parameters, the return type is inferred as `dynamic`
  even if the signature says `(bool, ulong, object?)`. Use explicit
  cast: `var r = ((bool, ulong, object?))store.GetProperty(...)`.

## Daily Usage

```bash
bash tools/vm/qemu-start.sh   # boot the VM
# (wait for WinRM, then start the bridge via scheduled task)
cargo run --example parity_test -p duke-sheets-excel-com
bash tools/vm/qemu-stop.sh    # graceful shutdown
```

## WinRM from the Host

A Python helper at `tools/vm/winrm-exec.py` sends commands to the VM
via WinRM Basic auth. It uses SOAP/WS-Man directly (no dependencies
beyond Python stdlib):

```bash
# Run a cmd command
python3 tools/vm/winrm-exec.py "hostname"

# Run a PowerShell command (handles pipes, quotes, special chars)
python3 tools/vm/winrm-exec.py -ps 'Get-Process | Select-Object -First 5'
```

The `-ps` flag encodes the command as base64 UTF-16LE and passes it
via `powershell -EncodedCommand`, which avoids all quoting issues.

## Debugging

**VNC**: Connect to `localhost:5901` to see the VM display.

**QEMU monitor**: Use Python since `socat` isn't always available:

```python
python3 -c "
import socket, time
s = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
s.connect('/tmp/duke-sheets-vm.sock')
s.settimeout(2)
s.recv(4096)  # banner
s.sendall(b'info status\n')
time.sleep(1)
print(s.recv(4096).decode(errors='replace'))
s.close()
"
```

**Screenshot from host**:
```bash
# Via QEMU monitor
python3 /tmp/qemu-cmd.py "screendump /tmp/screen.ppm"
convert /tmp/screen.ppm /tmp/screen.png   # requires ImageMagick
```

**Snapshot/revert**:
```bash
qemu-img snapshot -c working ~/.duke-sheets/windows11.qcow2    # save
qemu-img snapshot -a working ~/.duke-sheets/windows11.qcow2    # revert
qemu-img snapshot -l ~/.duke-sheets/windows11.qcow2            # list
```

**Firewall**: If the bridge server can't receive connections from the
host (QEMU port forwarding connects but the VM's bridge never sees
it), disable Windows Firewall:

```powershell
Set-NetFirewallProfile -All -Enabled False
```
