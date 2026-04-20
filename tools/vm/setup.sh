#!/bin/bash
# One-time setup script for the Windows Excel VM.
#
# Builds QEMU from source (with KVM + slirp user networking), creates a
# qcow2 disk, generates an unattended install floppy, and boots the VM to
# install Windows 11 in BIOS mode (SeaBIOS - no UEFI/TPM needed).
#
# Tested on Amazon Linux 2023. Should work on any RPM-based distro with KVM.
#
# Prerequisites:
#   - KVM support (/dev/kvm must exist)
#   - A Windows 11 ISO at ~/Win11.iso (or set WIN_ISO)
#   - sudo access for installing packages
#   - ~20GB free disk space
#
# Usage:
#   bash tools/vm/setup.sh

set -euo pipefail

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

DUKE_DIR="${DUKE_SHEETS_VM_DIR:-$HOME/.duke-sheets}"
BUILD_DIR="/tmp/duke-sheets-vm-build"
QEMU_VERSION="9.2.3"
QEMU_INSTALL="$DUKE_DIR/qemu"
WIN_ISO="${WIN_ISO:-$HOME/Win11.iso}"
DISK="$DUKE_DIR/windows11.qcow2"
DISK_SIZE="60G"
VM_RAM="8G"
VM_CPUS="4"
SHARE_DIR="/tmp/duke-sheets-excel"

# Ports forwarded from host to VM
PORT_BRIDGE=9876
PORT_WINRM=5985
PORT_SSH=2222

# ---------------------------------------------------------------------------
# Preflight checks
# ---------------------------------------------------------------------------

echo "=== Duke Sheets Excel VM Setup ==="
echo ""

if [ ! -e /dev/kvm ]; then
    echo "ERROR: /dev/kvm not found. KVM support is required."
    echo "  On EC2, use a .metal instance type."
    exit 1
fi

if [ ! -f "$WIN_ISO" ]; then
    echo "ERROR: Windows ISO not found at $WIN_ISO"
    echo "  Download from: https://www.microsoft.com/en-us/evalcenter/evaluate-windows-11-enterprise"
    echo "  Or set WIN_ISO=/path/to/your.iso"
    exit 1
fi

echo "  KVM:      $(ls /dev/kvm)"
echo "  ISO:      $WIN_ISO"
echo "  Disk:     $DISK"
echo "  QEMU:     $QEMU_INSTALL"
echo ""

mkdir -p "$DUKE_DIR" "$BUILD_DIR" "$SHARE_DIR"

# ---------------------------------------------------------------------------
# Step 1: Install system dependencies
# ---------------------------------------------------------------------------

step_install_deps() {
    echo "--- Step 1: Installing system dependencies ---"

    local PKGS="gcc make ninja-build glib2-devel pixman-devel zlib-devel dosfstools samba"

    if command -v dnf &>/dev/null; then
        sudo dnf install -y $PKGS 2>&1 | tail -3
    elif command -v apt-get &>/dev/null; then
        sudo apt-get update -qq
        sudo apt-get install -y build-essential ninja-build libglib2.0-dev \
            libpixman-1-dev zlib1g-dev dosfstools samba mtools 2>&1 | tail -3
    else
        echo "ERROR: No supported package manager found (need dnf or apt)."
        exit 1
    fi

    # Python packages needed by QEMU's configure
    pip install -q distlib meson 2>/dev/null || pip3 install -q distlib meson

    echo "  Done."
}

# ---------------------------------------------------------------------------
# Step 2: Build libslirp (user networking for QEMU)
# ---------------------------------------------------------------------------

step_build_libslirp() {
    echo "--- Step 2: Building libslirp ---"

    if pkg-config --exists slirp 2>/dev/null; then
        echo "  libslirp already installed ($(pkg-config --modversion slirp)), skipping."
        return
    fi

    local SLIRP_DIR="$BUILD_DIR/libslirp"
    if [ ! -d "$SLIRP_DIR" ]; then
        git clone --depth=1 https://gitlab.freedesktop.org/slirp/libslirp.git "$SLIRP_DIR" 2>&1 | tail -1
    fi

    cd "$SLIRP_DIR"
    rm -rf build
    meson setup build 2>&1 | tail -3
    ninja -C build -j"$(nproc)" 2>&1 | tail -3
    sudo ninja -C build install 2>&1 | tail -3
    sudo ldconfig

    echo "  Installed libslirp $(pkg-config --modversion slirp)"
}

# ---------------------------------------------------------------------------
# Step 3: Build QEMU
# ---------------------------------------------------------------------------

step_build_qemu() {
    echo "--- Step 3: Building QEMU $QEMU_VERSION ---"

    if [ -x "$QEMU_INSTALL/qemu-system-x86_64" ]; then
        local existing_ver
        existing_ver=$("$QEMU_INSTALL/qemu-system-x86_64" --version 2>/dev/null | head -1 | grep -oP '\d+\.\d+\.\d+' || true)
        if [ "$existing_ver" = "$QEMU_VERSION" ]; then
            echo "  QEMU $QEMU_VERSION already built, skipping."
            return
        fi
    fi

    local TARBALL="$BUILD_DIR/qemu-${QEMU_VERSION}.tar.xz"
    local SRC="$BUILD_DIR/qemu-${QEMU_VERSION}"

    if [ ! -f "$TARBALL" ]; then
        echo "  Downloading QEMU ${QEMU_VERSION}..."
        curl -L -o "$TARBALL" "https://download.qemu.org/qemu-${QEMU_VERSION}.tar.xz"
    fi

    if [ ! -d "$SRC" ]; then
        echo "  Extracting..."
        tar xf "$TARBALL" -C "$BUILD_DIR"
    fi

    cd "$SRC"
    rm -rf build

    # Ensure pkg-config can find libslirp in /usr/local
    export PKG_CONFIG_PATH="/usr/local/lib64/pkgconfig:/usr/local/lib/pkgconfig:${PKG_CONFIG_PATH:-}"

    echo "  Configuring (x86_64 system emulator only)..."
    ./configure \
        --target-list=x86_64-softmmu \
        --disable-gtk --disable-sdl --disable-opengl --disable-virglrenderer \
        --enable-vnc --enable-slirp \
        --disable-docs \
        2>&1 | grep -E "^(slirp|Build dir)" || true

    echo "  Building with $(nproc) cores..."
    ninja -C build -j"$(nproc)" qemu-system-x86_64 qemu-img 2>&1 | tail -3

    mkdir -p "$QEMU_INSTALL"
    cp build/qemu-system-x86_64 build/qemu-img "$QEMU_INSTALL/"

    echo "  Installed to $QEMU_INSTALL/"
    "$QEMU_INSTALL/qemu-system-x86_64" --version | head -1
}

# ---------------------------------------------------------------------------
# Step 4: Create qcow2 disk image
# ---------------------------------------------------------------------------

step_create_disk() {
    echo "--- Step 4: Creating VM disk ---"

    if [ -f "$DISK" ]; then
        echo "  Disk already exists at $DISK, skipping."
        return
    fi

    "$QEMU_INSTALL/qemu-img" create -f qcow2 "$DISK" "$DISK_SIZE"
    echo "  Created $DISK ($DISK_SIZE sparse)"
}

# ---------------------------------------------------------------------------
# Step 5: Create unattended install media
# ---------------------------------------------------------------------------

step_create_unattend() {
    echo "--- Step 5: Creating unattended install floppy ---"

    local FLOPPY="$DUKE_DIR/autounattend.img"
    local XML="$DUKE_DIR/autounattend.xml"

    # Write autounattend.xml - BIOS/MBR mode, bypasses TPM/SecureBoot/RAM/CPU,
    # creates local admin, enables WinRM + SSH, opens firewall for bridge port
    cat > "$XML" << 'XMLEOF'
<?xml version="1.0" encoding="utf-8"?>
<unattend xmlns="urn:schemas-microsoft-com:unattend">
  <settings pass="windowsPE">
    <component name="Microsoft-Windows-International-Core-WinPE"
               processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35"
               language="neutral" versionScope="nonSxS"
               xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State">
      <SetupUILanguage><UILanguage>en-US</UILanguage></SetupUILanguage>
      <InputLocale>en-US</InputLocale>
      <SystemLocale>en-US</SystemLocale>
      <UILanguage>en-US</UILanguage>
      <UserLocale>en-US</UserLocale>
    </component>
    <component name="Microsoft-Windows-Setup"
               processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35"
               language="neutral" versionScope="nonSxS"
               xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State">
      <RunSynchronous>
        <RunSynchronousCommand wcm:action="add">
          <Order>1</Order>
          <Path>reg add HKLM\SYSTEM\Setup\LabConfig /v BypassTPMCheck /t REG_DWORD /d 1 /f</Path>
        </RunSynchronousCommand>
        <RunSynchronousCommand wcm:action="add">
          <Order>2</Order>
          <Path>reg add HKLM\SYSTEM\Setup\LabConfig /v BypassSecureBootCheck /t REG_DWORD /d 1 /f</Path>
        </RunSynchronousCommand>
        <RunSynchronousCommand wcm:action="add">
          <Order>3</Order>
          <Path>reg add HKLM\SYSTEM\Setup\LabConfig /v BypassRAMCheck /t REG_DWORD /d 1 /f</Path>
        </RunSynchronousCommand>
        <RunSynchronousCommand wcm:action="add">
          <Order>4</Order>
          <Path>reg add HKLM\SYSTEM\Setup\LabConfig /v BypassCPUCheck /t REG_DWORD /d 1 /f</Path>
        </RunSynchronousCommand>
      </RunSynchronous>
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
      <UserData>
        <AcceptEula>true</AcceptEula>
        <FullName>User</FullName>
        <Organization>Dev</Organization>
        <ProductKey>
          <Key>W269N-WFGWX-YVC9B-4J6C9-T83GX</Key>
        </ProductKey>
      </UserData>
    </component>
  </settings>
  <!-- specialize runs as SYSTEM - full admin, perfect for registry/firewall -->
  <settings pass="specialize">
    <component name="Microsoft-Windows-Shell-Setup"
               processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35"
               language="neutral" versionScope="nonSxS">
      <ComputerName>EXCEL-VM</ComputerName>
      <TimeZone>UTC</TimeZone>
    </component>
    <component name="Microsoft-Windows-Deployment"
               processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35"
               language="neutral" versionScope="nonSxS"
               xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State">
      <RunSynchronous>
        <!-- Bypass OOBE internet requirement -->
        <RunSynchronousCommand wcm:action="add">
          <Order>1</Order>
          <Path>reg add HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE /v BypassNRO /t REG_DWORD /d 1 /f</Path>
        </RunSynchronousCommand>
        <!-- Disable UAC so FirstLogonCommands has full admin rights -->
        <RunSynchronousCommand wcm:action="add">
          <Order>2</Order>
          <Path>reg add HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v EnableLUA /t REG_DWORD /d 0 /f</Path>
        </RunSynchronousCommand>
        <!-- Firewall: allow bridge port -->
        <RunSynchronousCommand wcm:action="add">
          <Order>3</Order>
          <Path>netsh advfirewall firewall add rule name="ExcelBridge" dir=in action=allow protocol=tcp localport=9876</Path>
        </RunSynchronousCommand>
        <!-- Allow SMB guest access (QEMU share has no auth) -->
        <RunSynchronousCommand wcm:action="add">
          <Order>4</Order>
          <Path>reg add HKLM\SOFTWARE\Policies\Microsoft\Windows\LanmanWorkstation /v AllowInsecureGuestAuth /t REG_DWORD /d 1 /f</Path>
        </RunSynchronousCommand>
        <!-- Disable Windows Update -->
        <RunSynchronousCommand wcm:action="add">
          <Order>5</Order>
          <Path>reg add HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU /v NoAutoUpdate /t REG_DWORD /d 1 /f</Path>
        </RunSynchronousCommand>
      </RunSynchronous>
    </component>
  </settings>
  <settings pass="oobeSystem">
    <component name="Microsoft-Windows-International-Core"
               processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35"
               language="neutral" versionScope="nonSxS">
      <InputLocale>en-US</InputLocale>
      <SystemLocale>en-US</SystemLocale>
      <UILanguage>en-US</UILanguage>
      <UserLocale>en-US</UserLocale>
    </component>
    <component name="Microsoft-Windows-Shell-Setup"
               processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35"
               language="neutral" versionScope="nonSxS"
               xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State">
      <OOBE>
        <HideEULAPage>true</HideEULAPage>
        <HideLocalAccountScreen>true</HideLocalAccountScreen>
        <HideOnlineAccountScreens>true</HideOnlineAccountScreens>
        <HideWirelessSetupInOOBE>true</HideWirelessSetupInOOBE>
        <ProtectYourPC>3</ProtectYourPC>
      </OOBE>
      <UserAccounts>
        <LocalAccounts>
          <LocalAccount wcm:action="add">
            <Name>user</Name>
            <Group>Administrators</Group>
            <Password>
              <Value>test</Value>
              <PlainText>true</PlainText>
            </Password>
          </LocalAccount>
        </LocalAccounts>
      </UserAccounts>
      <AutoLogon>
        <Enabled>true</Enabled>
        <Username>user</Username>
        <Password>
          <Value>test</Value>
          <PlainText>true</PlainText>
        </Password>
        <LogonCount>999</LogonCount>
      </AutoLogon>
      <FirstLogonCommands>
        <!-- Set network to Private (QEMU NIC defaults to Public which blocks WinRM) -->
        <SynchronousCommand wcm:action="add">
          <Order>1</Order>
          <CommandLine>powershell -Command "Get-NetConnectionProfile | Set-NetConnectionProfile -NetworkCategory Private"</CommandLine>
          <RequiresUserInput>false</RequiresUserInput>
        </SynchronousCommand>
        <!-- Enable WinRM with -SkipNetworkProfileCheck as fallback -->
        <SynchronousCommand wcm:action="add">
          <Order>2</Order>
          <CommandLine>powershell -Command "Enable-PSRemoting -Force -SkipNetworkProfileCheck"</CommandLine>
          <RequiresUserInput>false</RequiresUserInput>
        </SynchronousCommand>
        <!-- Configure WinRM for unencrypted Basic auth (test VM only) -->
        <SynchronousCommand wcm:action="add">
          <Order>3</Order>
          <CommandLine>cmd /c winrm set winrm/config/service @{AllowUnencrypted="true"}</CommandLine>
          <RequiresUserInput>false</RequiresUserInput>
        </SynchronousCommand>
        <SynchronousCommand wcm:action="add">
          <Order>4</Order>
          <CommandLine>cmd /c winrm set winrm/config/service/auth @{Basic="true"}</CommandLine>
          <RequiresUserInput>false</RequiresUserInput>
        </SynchronousCommand>
      </FirstLogonCommands>
    </component>
  </settings>
</unattend>
XMLEOF

    # Create FAT12 floppy image with the XML
    dd if=/dev/zero of="$FLOPPY" bs=1440k count=1 2>/dev/null
    mkfs.fat "$FLOPPY" 2>/dev/null

    local MNT
    MNT=$(mktemp -d)
    sudo mount -o loop "$FLOPPY" "$MNT"
    sudo cp "$XML" "$MNT/autounattend.xml"
    sudo umount "$MNT"
    rmdir "$MNT"

    echo "  Created $FLOPPY"
}

# ---------------------------------------------------------------------------
# Step 6: Boot the VM
# ---------------------------------------------------------------------------

step_boot_vm() {
    echo "--- Step 6: Booting VM for Windows installation ---"

    local PID_FILE="/tmp/duke-sheets-vm.pid"

    if [ -f "$PID_FILE" ] && kill -0 "$(cat "$PID_FILE")" 2>/dev/null; then
        echo "  VM already running (PID $(cat "$PID_FILE"))."
        return
    fi

    export LD_LIBRARY_PATH="/usr/local/lib64:/usr/local/lib:${LD_LIBRARY_PATH:-}"

    # BIOS mode (SeaBIOS) - no UEFI/TPM needed, bypasses handled in autounattend.
    # Uses q35 machine with Hyper-V enlightenments for best Windows performance.
    "$QEMU_INSTALL/qemu-system-x86_64" \
        -M q35,usb=on,acpi=on,hpet=off \
        -accel kvm \
        -cpu host,hv_relaxed,hv_frequencies,hv_vpindex,hv_ipi,hv_tlbflush,hv_spinlocks=0x1fff,hv_synic,hv_runtime,hv_time,hv_stimer,hv_vapic \
        -m "$VM_RAM" \
        -smp "cores=$VM_CPUS" \
        -drive "file=$DISK" \
        -device usb-tablet \
        -device VGA,vgamem_mb=256 \
        -cdrom "$WIN_ISO" \
        -drive "file=$DUKE_DIR/autounattend.img,format=raw,if=floppy" \
        -boot order=d \
        -nic "user,model=e1000,hostfwd=tcp::${PORT_BRIDGE}-:${PORT_BRIDGE},hostfwd=tcp::${PORT_WINRM}-:${PORT_WINRM},hostfwd=tcp::${PORT_SSH}-:22,smb=$SHARE_DIR" \
        -display none \
        -vnc :1 \
        -daemonize \
        -pidfile "$PID_FILE" \
        -monitor "unix:/tmp/duke-sheets-vm.sock,server,nowait"

    echo "  VM started (PID $(cat "$PID_FILE"))"
    echo ""
    echo "  Windows is now installing unattended. This takes ~10-15 minutes."
    echo "  Port forwarding:"
    echo "    localhost:$PORT_BRIDGE -> bridge server"
    echo "    localhost:$PORT_WINRM -> WinRM"
    echo "    localhost:$PORT_SSH   -> SSH"
    echo "  SMB share: $SHARE_DIR -> \\\\10.0.2.4\\qemu"
    echo "  VNC: localhost:5901 (for debugging)"
    echo ""
    echo "  Monitor progress:"
    echo "    Watch SSH: while ! ssh -o ConnectTimeout=2 -p $PORT_SSH user@localhost hostname 2>/dev/null; do sleep 10; echo waiting...; done"
    echo ""
    echo "  When install is done, snapshot the VM:"
    echo "    $QEMU_INSTALL/qemu-img snapshot -c fresh-install $DISK"
}

# ---------------------------------------------------------------------------
# Run all steps
# ---------------------------------------------------------------------------

step_install_deps
step_build_libslirp
step_build_qemu
step_create_disk
step_create_unattend
step_boot_vm

echo ""
echo "=== Setup complete ==="
