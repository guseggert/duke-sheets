#!/usr/bin/env python3
"""Execute a command on Windows VM via WinRM Basic auth."""

import sys, base64, uuid, time
import urllib.request, urllib.error
import xml.etree.ElementTree as ET

URL = "http://localhost:5985/wsman"
USER = "user"
PASS = "test"
AUTH = base64.b64encode(f"{USER}:{PASS}".encode()).decode()

HEADERS = {
    "Content-Type": "application/soap+xml;charset=UTF-8",
    "Authorization": f"Basic {AUTH}",
}

NS = {
    "s": "http://www.w3.org/2003/05/soap-envelope",
    "a": "http://schemas.xmlsoap.org/ws/2004/08/addressing",
    "w": "http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd",
    "rsp": "http://schemas.microsoft.com/wbem/wsman/1/windows/shell",
}


def soap_request(body):
    req = urllib.request.Request(
        URL, data=body.encode("utf-8"), headers=HEADERS, method="POST"
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            return resp.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        err = e.read().decode("utf-8")
        print(f"HTTP {e.code}: {err[:500]}", file=sys.stderr)
        return None


def create_shell():
    body = f"""<?xml version="1.0" encoding="UTF-8"?>
<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope"
  xmlns:a="http://schemas.xmlsoap.org/ws/2004/08/addressing"
  xmlns:w="http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd"
  xmlns:rsp="http://schemas.microsoft.com/wbem/wsman/1/windows/shell">
  <s:Header>
    <a:To>http://localhost:5985/wsman</a:To>
    <w:ResourceURI s:mustUnderstand="true">http://schemas.microsoft.com/wbem/wsman/1/windows/shell/cmd</w:ResourceURI>
    <a:ReplyTo><a:Address>http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</a:Address></a:ReplyTo>
    <a:Action s:mustUnderstand="true">http://schemas.xmlsoap.org/ws/2004/09/transfer/Create</a:Action>
    <a:MessageID>uuid:{uuid.uuid4()}</a:MessageID>
    <w:OptionSet>
      <w:Option Name="WINRS_NOPROFILE">TRUE</w:Option>
      <w:Option Name="WINRS_CODEPAGE">65001</w:Option>
    </w:OptionSet>
  </s:Header>
  <s:Body>
    <rsp:Shell>
      <rsp:InputStreams>stdin</rsp:InputStreams>
      <rsp:OutputStreams>stdout stderr</rsp:OutputStreams>
    </rsp:Shell>
  </s:Body>
</s:Envelope>"""
    resp = soap_request(body)
    if not resp:
        return None
    root = ET.fromstring(resp)
    el = root.find(".//rsp:ShellId", NS)
    return el.text if el is not None else None


def run_command(shell_id, command):
    body = f"""<?xml version="1.0" encoding="UTF-8"?>
<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope"
  xmlns:a="http://schemas.xmlsoap.org/ws/2004/08/addressing"
  xmlns:w="http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd"
  xmlns:rsp="http://schemas.microsoft.com/wbem/wsman/1/windows/shell">
  <s:Header>
    <a:To>http://localhost:5985/wsman</a:To>
    <w:ResourceURI s:mustUnderstand="true">http://schemas.microsoft.com/wbem/wsman/1/windows/shell/cmd</w:ResourceURI>
    <a:ReplyTo><a:Address>http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</a:Address></a:ReplyTo>
    <a:Action s:mustUnderstand="true">http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Command</a:Action>
    <a:MessageID>uuid:{uuid.uuid4()}</a:MessageID>
    <w:SelectorSet>
      <w:Selector Name="ShellId">{shell_id}</w:Selector>
    </w:SelectorSet>
  </s:Header>
  <s:Body>
    <rsp:CommandLine>
      <rsp:Command>{command}</rsp:Command>
    </rsp:CommandLine>
  </s:Body>
</s:Envelope>"""
    resp = soap_request(body)
    if not resp:
        return None
    root = ET.fromstring(resp)
    el = root.find(".//rsp:CommandId", NS)
    return el.text if el is not None else None


def get_output(shell_id, command_id):
    stdout_parts = []
    stderr_parts = []
    done = False
    exit_code = -1

    while not done:
        body = f'''<?xml version="1.0" encoding="UTF-8"?>
<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope"
  xmlns:a="http://schemas.xmlsoap.org/ws/2004/08/addressing"
  xmlns:w="http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd"
  xmlns:rsp="http://schemas.microsoft.com/wbem/wsman/1/windows/shell">
  <s:Header>
    <a:To>http://localhost:5985/wsman</a:To>
    <w:ResourceURI s:mustUnderstand="true">http://schemas.microsoft.com/wbem/wsman/1/windows/shell/cmd</w:ResourceURI>
    <a:ReplyTo><a:Address>http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</a:Address></a:ReplyTo>
    <a:Action s:mustUnderstand="true">http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Receive</a:Action>
    <a:MessageID>uuid:{uuid.uuid4()}</a:MessageID>
    <w:SelectorSet>
      <w:Selector Name="ShellId">{shell_id}</w:Selector>
    </w:SelectorSet>
  </s:Header>
  <s:Body>
    <rsp:Receive>
      <rsp:DesiredStream CommandId="{command_id}">stdout stderr</rsp:DesiredStream>
    </rsp:Receive>
  </s:Body>
</s:Envelope>'''
        resp = soap_request(body)
        if not resp:
            break
        root = ET.fromstring(resp)

        for stream in root.findall(".//rsp:Stream", NS):
            text = stream.text
            if text:
                decoded = base64.b64decode(text).decode("utf-8", errors="replace")
                if stream.get("Name") == "stdout":
                    stdout_parts.append(decoded)
                elif stream.get("Name") == "stderr":
                    stderr_parts.append(decoded)

        # Check if command is done
        state = root.find(".//rsp:CommandState", NS)
        if state is not None:
            state_val = state.get("State", "")
            if "Done" in state_val:
                done = True
                exit_el = state.find("rsp:ExitCode", NS)
                exit_code = int(exit_el.text) if exit_el is not None else -1

        if not done:
            time.sleep(0.5)

    return "".join(stdout_parts), "".join(stderr_parts), exit_code


def delete_shell(shell_id):
    body = f"""<?xml version="1.0" encoding="UTF-8"?>
<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope"
  xmlns:a="http://schemas.xmlsoap.org/ws/2004/08/addressing"
  xmlns:w="http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd">
  <s:Header>
    <a:To>http://localhost:5985/wsman</a:To>
    <w:ResourceURI s:mustUnderstand="true">http://schemas.microsoft.com/wbem/wsman/1/windows/shell/cmd</w:ResourceURI>
    <a:ReplyTo><a:Address>http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</a:Address></a:ReplyTo>
    <a:Action s:mustUnderstand="true">http://schemas.xmlsoap.org/ws/2004/09/transfer/Delete</a:Action>
    <a:MessageID>uuid:{uuid.uuid4()}</a:MessageID>
    <w:SelectorSet>
      <w:Selector Name="ShellId">{shell_id}</w:Selector>
    </w:SelectorSet>
  </s:Header>
  <s:Body/>
</s:Envelope>"""
    soap_request(body)


def run_ps(command):
    """Run a PowerShell command (handles encoding for pipes, quotes, etc.)."""
    # Encode as base64 UTF-16LE for powershell -EncodedCommand
    encoded = base64.b64encode(command.encode("utf-16-le")).decode()
    return f"powershell -EncodedCommand {encoded}"


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: winrm-exec.py [-ps] <command>")
        sys.exit(1)

    use_ps = sys.argv[1] == "-ps"
    if use_ps:
        command = " ".join(sys.argv[2:])
        command = run_ps(command)
    else:
        command = " ".join(sys.argv[1:])

    shell_id = create_shell()
    if not shell_id:
        print("ERROR: Failed to create shell", file=sys.stderr)
        sys.exit(1)

    cmd_id = run_command(shell_id, command)
    if not cmd_id:
        print("ERROR: Failed to run command", file=sys.stderr)
        delete_shell(shell_id)
        sys.exit(1)

    stdout, stderr, exit_code = get_output(shell_id, cmd_id)
    if stdout:
        print(stdout, end="")
    if stderr:
        print(stderr, end="", file=sys.stderr)

    delete_shell(shell_id)
    sys.exit(0 if exit_code == 0 else 1)
