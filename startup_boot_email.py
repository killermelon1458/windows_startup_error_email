# startup_boot_email.py
# Sends a boot email with machine info + most recent crash markers & recent error/critical events.
# Requires pythonEmailNotify.EmailSender and env vars:
#   EMAIL_ADDRESS, EMAIL_PASSWORD, MAIN_EMAIL_ADDRESS

import os
import socket
import json
import subprocess
from datetime import datetime, timezone
from email.utils import formatdate

# Optional debug flag: set BOOTMAIL_DEBUG=1 in env to print status and send a probe email
DEBUG = os.getenv("BOOTMAIL_DEBUG", "0") == "1"

def _bootmail_debug_enabled() -> bool:
    return DEBUG

def _dbg(msg: str):
    if DEBUG:
        try:
            print(f"[BOOTMAIL] {msg}")
        except Exception:
            pass

try:
    from pythonEmailNotify import EmailSender
except Exception as e:
    raise SystemExit(f"Missing or broken pythonEmailNotify.py ({e}). Place it next to this script or on PYTHONPATH.")

PS = r"C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe"


def run_ps(code: str) -> str:
    """
    Run PowerShell code and return stdout.
    - Forces errors to throw, captures full diagnostics to stdout prefixed with PS_ERR:
    - Returns stdout even if empty; raises RuntimeError with rich context on failure.
    """
    WRAPPED = (
        "$ErrorActionPreference='Stop';"
        "try {"
        f"  {code} "
        "} catch {"
        "  Write-Output ('PS_ERR:' + $_.Exception.GetType().FullName + ' | ' + $_.Exception.Message);"
        "  if ($_.InvocationInfo -and $_.InvocationInfo.PositionMessage) {"
        "    Write-Output ('PS_ERR_POS:' + $_.InvocationInfo.PositionMessage)"
        "  }"
        "  exit 1"
        "}"
    )
    p = subprocess.run(
        [PS, "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", WRAPPED],
        capture_output=True, text=True, encoding="utf-8", errors="replace"
    )
    out = (p.stdout or "").strip()
    if p.returncode != 0:
        # include stderr too, in case PS wrote anything there
        err = (p.stderr or "").strip()
        raise RuntimeError(f"PowerShell failed\nCODE:\n{code}\nSTDOUT:\n{out}\nSTDERR:\n{err}")
    return out


def iso_utc_now() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


def get_boot_time_iso() -> str:
    ps = "(Get-CimInstance Win32_OperatingSystem).LastBootUpTime.ToUniversalTime().ToString('o')"
    return run_ps(ps)


def get_machine_name() -> str:
    # Prefer Windows environment, fallback to socket
    return os.getenv("COMPUTERNAME") or socket.gethostname()


def get_local_ipv4_list() -> list[str]:
    """
    Return a list of local IPv4 addresses.
    - Start with fallback from env var MACHINE_IP (if set).
    - Then try PowerShell detection; if it succeeds, it overwrites the fallback.
    """
    # Fallback from env
    ips: list[str] = []
    env_ip = os.getenv("MACHINE_IP")
    if env_ip:
        ips.append(env_ip)

    try:
        ps = (
            "Get-NetIPAddress -AddressFamily IPv4 "
            "| Where-Object { $_.InterfaceOperationalStatus -eq 'Up' "
            " -and $_.IPAddress -notlike '169.*' "
            " -and $_.IPAddress -ne '127.0.0.1' } "
            "| Select-Object -ExpandProperty IPAddress"
        )
        out = run_ps(ps)
        detected = [line.strip() for line in out.splitlines() if line.strip()]
        if detected:
            ips = detected  # overwrite fallback if detection worked
    except Exception:
        # keep whatever we had from env_ip
        pass

    return ips


def get_latest_crash_marker() -> dict | None:
    """
    Returns the newest crash-related event as dict or None.
    IDs: 41 (Kernel-Power), 6008 (Unexpected Shutdown), 1001 (BugCheck)
    Ensures TimeCreated is an ISO-8601 string (UTC).
    """
    ps = (
        "$ids=41,6008,1001; "
        "Get-WinEvent -FilterHashtable @{LogName='System'; Id=$ids} -MaxEvents 200 -ErrorAction SilentlyContinue | "
        "Sort-Object TimeCreated | "
        "Select-Object @{n='TimeCreated';e={$_.TimeCreated.ToUniversalTime().ToString('o')}},Id,ProviderName,Message | "
        "Select-Object -Last 1 | "
        "ConvertTo-Json -Depth 3 -Compress"
    )
    out = run_ps(ps)
    if not out:
        return None
    try:
        data = json.loads(out)
        if isinstance(data, dict) and data.get("TimeCreated"):
            return data
    except json.JSONDecodeError:
        pass
    return None


def get_recent_errors_near(time_iso: str | None, hours_before: int = 6, limit: int = 10) -> list[dict]:
    """
    Return recent Critical/Error (Level=1,2) events.
    If time_iso is provided (ISO-8601), return events in [time-hrs, time].
    Otherwise return last `limit` Critical/Error events overall.
    Always emits TimeCreated in ISO-8601 UTC.
    """
    if time_iso:
        # Tolerant parse for '...Z' and offsets; adjust to UTC.
        ps = (
            "[CultureInfo] = [System.Globalization.CultureInfo]::InvariantCulture; "
            "[Styles] = [System.Globalization.DateTimeStyles]::AdjustToUniversal; "
            f"$t=[datetime]::Parse('{time_iso}', [CultureInfo], [Styles]); "
            "$start = $t.AddHours(-{hrs}); ".format(hrs=hours_before) +
            # Pull a healthy slice and filter in pipeline (avoids EndTime quirks)
            "Get-WinEvent -FilterHashtable @{LogName='System'; Level=1,2} -MaxEvents 1000 -ErrorAction SilentlyContinue | "
            "Where-Object { $_.TimeCreated -le $t -and $_.TimeCreated -ge $start } | "
            "Sort-Object TimeCreated | "
            f"Select-Object -Last {limit} | "
            "Select-Object @{n='TimeCreated';e={$_.TimeCreated.ToUniversalTime().ToString('o')}},Id,ProviderName,Message | "
            "ConvertTo-Json -Depth 3 -Compress"
        )
    else:
        ps = (
            "Get-WinEvent -FilterHashtable @{LogName='System'; Level=1,2} -MaxEvents 500 -ErrorAction SilentlyContinue | "
            "Sort-Object TimeCreated | "
            f"Select-Object -Last {limit} | "
            "Select-Object @{n='TimeCreated';e={$_.TimeCreated.ToUniversalTime().ToString('o')}},Id,ProviderName,Message | "
            "ConvertTo-Json -Depth 3 -Compress"
        )

    try:
        out = run_ps(ps)
        if not out:
            return []
        data = json.loads(out)
        if isinstance(data, dict):
            return [data]
        return data or []
    except Exception:
        # Fallback: just fetch the last N Critical/Error events so email still goes out
        try:
            fallback = (
                "Get-WinEvent -FilterHashtable @{LogName='System'; Level=1,2} -MaxEvents 500 -ErrorAction SilentlyContinue | "
                "Sort-Object TimeCreated | "
                f"Select-Object -Last {limit} | "
                "Select-Object @{n='TimeCreated';e={$_.TimeCreated.ToUniversalTime().ToString('o')}},Id,ProviderName,Message | "
                "ConvertTo-Json -Depth 3 -Compress"
            )
            out2 = run_ps(fallback)
            if not out2:
                return []
            data2 = json.loads(out2)
            if isinstance(data2, dict):
                return [data2]
            return data2 or []
        except Exception:
            return []


def fmt_dt_local(iso_utc: str) -> str:
    try:
        dt = datetime.fromisoformat(iso_utc.replace("Z", "+00:00")).astimezone()
        return dt.strftime("%Y-%m-%d %H:%M:%S %Z")
    except Exception:
        return iso_utc


def make_html(machine: str, ips: list[str], boot_iso: str, crash: dict | None, errors: list[dict]) -> str:
    ip_html = "<br>".join(ips) if ips else "(none detected)"
    boot_local = fmt_dt_local(boot_iso)
    crash_block = ""
    if crash:
        msg = (crash.get("Message") or "").replace("\r", " ").replace("\n", " ").strip()
        crash_block = f"""
        <h3>Most recent crash marker</h3>
        <ul>
          <li><b>Time (UTC):</b> {crash.get("TimeCreated")}</li>
          <li><b>Time (Local):</b> {fmt_dt_local(crash.get("TimeCreated") or "")}</li>
          <li><b>Event ID:</b> {crash.get("Id")}</li>
          <li><b>Provider:</b> {crash.get("ProviderName")}</li>
          <li><b>Message:</b><br><pre style="white-space:pre-wrap">{msg}</pre></li>
        </ul>
        """
    rows = ""
    for ev in errors:
        msg = (ev.get("Message") or "").replace("\r", " ").replace("\n", " ").strip()
        rows += (
            f"<tr>"
            f"<td>{ev.get('TimeCreated','')}</td>"
            f"<td>{fmt_dt_local(ev.get('TimeCreated',''))}</td>"
            f"<td>{ev.get('Id','')}</td>"
            f"<td>{ev.get('ProviderName','')}</td>"
            f"<td><pre style='white-space:pre-wrap'>{msg}</pre></td>"
            f"</tr>"
        )

    return f"""
    <h2>PC Boot Notification</h2>
    <p><b>Machine:</b> {machine}</p>
    <p><b>Local IPv4:</b><br>{ip_html}</p>
    <p><b>Boot Time (UTC):</b> {boot_iso}<br>
       <b>Boot Time (Local):</b> {boot_local}</p>
    <p><i>Startup detected at:</i> {formatdate(localtime=True)}</p>
    {crash_block}
    <h3>Recent Critical/Error events</h3>
    <table border="1" cellpadding="6" cellspacing="0">
      <thead>
        <tr><th>UTC</th><th>Local</th><th>ID</th><th>Provider</th><th>Message</th></tr>
      </thead>
      <tbody>{rows or "<tr><td colspan='5'>(none)</td></tr>"}</tbody>
    </table>
    """


def build_sender_from_env() -> EmailSender | None:
    """Create EmailSender from env vars. Return None if missing creds. (diff: added helper version tag)"""
    addr = os.getenv("EMAIL_ADDRESS")
    pw = os.getenv("EMAIL_PASSWORD")
    rcpt = os.getenv("MAIN_EMAIL_ADDRESS") or addr
    if addr and pw:
        return EmailSender(
            smtp_server="smtp.gmail.com",
            port=587,
            login=addr,
            password=pw,
            default_recipient=rcpt,
        )
    # explicit None to make branch obvious for editor differ
    return None  # v2


def notify_exception(e: Exception) -> None:
    """Best-effort exception email using env-var creds (kept for security); v2 to ensure editor update."""
    try:
        sender = build_sender_from_env()
        if sender is not None:
            sender.sendException(e)
    except Exception:
        # Swallow to avoid masking the original exception, but keep a small local print for DEBUG
        if DEBUG:
            try:
                print(f"[BOOTMAIL] notify_exception failed: {e}")
            except Exception:
                pass
        pass


def main():
    # EmailSender from your helper; env vars hold creds/recipient.
    sender = build_sender_from_env()
    if sender is None:
        raise SystemExit("EMAIL_ADDRESS/EMAIL_PASSWORD not set in environment.")

    _dbg("Email sender built. From/To=" + str(getattr(sender, 'login', 'unknown')))

    machine = get_machine_name()
    _dbg(f"Machine={machine}")
    ips = get_local_ipv4_list()
    _dbg(f"IPs={ips}")
    boot_iso = get_boot_time_iso()
    _dbg(f"BootISO={boot_iso}")

    crash = get_latest_crash_marker()
    _dbg("Crash marker=" + (crash.get("TimeCreated") if crash else "None"))
    # If we have a crash time, pull errors near it; else just recent errors
    ref_time = crash.get("TimeCreated") if crash else None

    # Make sure a transient WinEvent hiccup does not block the boot email
    try:
        recent_errors = get_recent_errors_near(ref_time, hours_before=6, limit=10)
        _dbg(f"Fetched {len(recent_errors)} recent error/critical events")
    except Exception as e:
        notify_exception(e)
        _dbg(f"Recent-errors fetch failed: {e}")
        recent_errors = []

    subj = f"[BOOT] {machine} is online — {', '.join(ips) if ips else 'no IPs'}"
    html = make_html(machine, ips, boot_iso, crash, recent_errors)
    _dbg("Composed email HTML")

    if DEBUG:
        try:
            sender.sendEmail(f"[BOOT-TEST] {machine}", f"Probe. IPs={ips}", html=False)
            _dbg("Probe email sent")
        except Exception as e:
            _dbg(f"Probe email failed: {e}")
            notify_exception(e)

    try:
        sender.sendEmail(subj, html, html=True)
        _dbg("Boot email sent")
    except Exception as e:
        _dbg(f"Boot email failed: {e}")
        notify_exception(e)
        raise


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        # Send an exception report via your pythonEmailNotify helper, using env creds only
        notify_exception(e)
        raise
