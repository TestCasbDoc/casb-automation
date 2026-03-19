"""
run.py — CASB Automation Entry Point.

To add a new app:
  1. Create apps/{app_id}/app.yaml
  2. Create apps/{app_id}/activities.py
  3. Add one line to _APP_MAP below
  That's all — no other files need touching.
"""

import sys
import os

# Ensure the folder containing run.py is on the path so config.py
# and the core/ and apps/ packages are always found regardless of
# which directory Python is launched from.
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if _SCRIPT_DIR not in sys.path:
    sys.path.insert(0, _SCRIPT_DIR)

import argparse
import re as _re
from datetime import datetime

# ── App registry — ONE LINE PER APP ──────────────────────────────────────────
# Format: "app_id": ["account_type1", "account_type2"]
# Use ["any"] for apps with no personal/corporate distinction.
_APP_MAP = {
    "ms_teams" : ["personal", "corporate"],
    # "instagram": ["any"],          # ← example: uncomment + add apps/instagram/
    # "twitter"  : ["any"],
    # "onedrive" : ["any"],
    # "gmail"    : ["any"],
    # "slack"    : ["any"],
    # "box"      : ["any"],
}
# ─────────────────────────────────────────────────────────────────────────────

_DEFAULT_APP         = "ms_teams"
_DEFAULT_ACCOUNT_TYPE = "corporate"


# ── Argument parsing ──────────────────────────────────────────────────────────

parser = argparse.ArgumentParser(prog="run.py", add_help=False)
parser.add_argument("--applications", default=None)
parser.add_argument("--account_type", default=None, choices=["personal", "corporate"])
parser.add_argument("--host",             required=True)
parser.add_argument("--pwd",              required=True)
parser.add_argument("--ssh-user",         required=True)
parser.add_argument("--org",             default=None)
parser.add_argument("--report-dir",      default=None)
parser.add_argument("--activities",      default="all")
parser.add_argument("--qosmos",          default=None)
parser.add_argument("--send-email",      default=None)
parser.add_argument("--smtp-pwd",        default=None)
parser.add_argument("--server-url",      default=None)
parser.add_argument("--access-policy",   default=None)
parser.add_argument("--decrypt-policy",  default=None)
parser.add_argument("--decrypt-rule",    default=None)
parser.add_argument("--decrypt-profile", default=None)
parser.add_argument("--casb-profile",    default=None)
parser.add_argument("--casb-access-policy-rule", default=None)
parser.add_argument("--help", "-h", action="store_true", default=False)
args, _ = parser.parse_known_args()

if args.help:
    print("""
=============================================================
  CASB Automation — run.py
=============================================================
USAGE:
  python run.py --applications APP[account_type] --host IP --pwd PWD --ssh-user USER

EXAMPLES:
  python run.py --applications "MS_Teams[personal]" --host 172.20.4.5 --pwd versa123 --ssh-user admin
  python run.py --applications "MS_Teams[personal,corporate]" --host 172.20.4.5 --pwd versa123 --ssh-user admin --activities "post[1]"

REGISTERED APPS:
""")
    for app_id, atypes in _APP_MAP.items():
        print(f"  {app_id:20} account_types: {atypes}")
    print()
    sys.exit(0)


# ── Parse --applications ──────────────────────────────────────────────────────

def _parse_applications(raw: str, global_at: str):
    results = []
    for token in _re.split(r',(?![^\[]*\])', raw):
        token = token.strip()
        if not token:
            continue
        m = _re.match(r'^([^\[]+)(?:\[([^\]]+)\])?$', token)
        if not m:
            print(f"[ERROR] Cannot parse: '{token}'")
            sys.exit(1)
        app_id = m.group(1).strip().lower().replace(" ", "_")
        at_str = m.group(2)
        if at_str:
            at_list = [a.strip().lower() for a in at_str.split(",") if a.strip()]
        elif args.account_type:
            at_list = [global_at]
        else:
            supported = _APP_MAP.get(app_id, ["any"])
            at_list = supported if supported != ["any"] else ["any"]
        results.append((app_id, at_list))
    return results


raw_apps    = args.applications or _DEFAULT_APP
global_at   = args.account_type or _DEFAULT_ACCOUNT_TYPE
parsed_apps = _parse_applications(raw_apps, global_at)

# Validate
for app_id, at_list in parsed_apps:
    if app_id not in _APP_MAP:
        print(f"\n[ERROR] Unknown app: '{app_id}'")
        print(f"  Registered apps: {', '.join(_APP_MAP.keys())}")
        print(f"  To add '{app_id}': create apps/{app_id}/app.yaml + apps/{app_id}/activities.py")
        sys.exit(1)

# ── Parse --activities ────────────────────────────────────────────────────────

def _parse_run_navs(activities_arg: str):
    """
    Convert --activities string to (activity_names set, tc_numbers set).

    TC numbers map to app.yaml tc_label values (TC1=1, TC2=2 etc).
    If tc_numbers is empty, all activities in the name set are run.

    Examples:
      all              → ({"all"}, set())       run everything
      post             → ({"post"}, set())      run all post activities
      post[1]          → ({"post"}, {1})        run only TC1
      post[1,3,4]      → ({"post"}, {1,3,4})   run TC1, TC3, TC4
      forward reply    → ({"forward","reply"}, set())
    """
    if not activities_arg or activities_arg.strip().lower() == "all":
        return {"all"}, set()

    navs    = set()
    tc_nums = set()
    for part in activities_arg.strip().split():
        m = _re.match(r'^([^\[]+)(?:\[([^\]]+)\])?$', part.strip().lower())
        if m:
            name = m.group(1).strip()
            if name:
                navs.add(name)
            if m.group(2):
                for n in m.group(2).split(","):
                    n = n.strip()
                    if n.isdigit():
                        tc_nums.add(int(n))
    return navs, tc_nums

run_navs, run_tc_nums = _parse_run_navs(args.activities)


# ── Apply CLI overrides to config ─────────────────────────────────────────────

import config as _cfg

_cfg.SSH_HOST    = args.host
_cfg.SSH_PASSWORD = args.pwd
_cfg.SSH_USER    = args.ssh_user

if args.org:             _cfg.VOS_ORG_NAME               = args.org
if args.report_dir:
    _cfg.BASE_DIR   = args.report_dir
    _cfg.SCRIPT_DIR = os.path.join(args.report_dir,
                                   datetime.now().strftime("run_%Y%m%d_%H%M%S"))
    os.makedirs(_cfg.SCRIPT_DIR, exist_ok=True)
    _cfg.REPORT_FILE = os.path.join(_cfg.SCRIPT_DIR, "test_report.json")
    _cfg.HTML_REPORT = os.path.join(_cfg.SCRIPT_DIR, "test_report.html")

if args.access_policy:   _cfg.VOS_ACCESS_POLICY_NAME     = args.access_policy
if args.decrypt_policy:  _cfg.VOS_DECRYPTION_POLICY_NAME = args.decrypt_policy
if args.decrypt_rule:    _cfg.VOS_DECRYPTION_RULE_NAME   = args.decrypt_rule
if args.decrypt_profile: _cfg.VOS_DECRYPT_PROFILE_NAME   = args.decrypt_profile
if args.casb_profile:    _cfg.VOS_CASB_PROFILE_NAME      = args.casb_profile
if args.casb_access_policy_rule: _cfg.VOS_CASB_RULE_NAME = args.casb_access_policy_rule
if args.qosmos is not None:
    _cfg.VOS_APPID_REPORT_METADATA = "enable" if args.qosmos.lower() in ("true","1","yes") else "disable"


# ── Email + Upload helpers ────────────────────────────────────────────────────

def _build_email_html(all_results, overall_status, run_ts):
    """Build clean email-safe HTML summary table — works in Gmail/Outlook."""
    _NAV_DETAILS = {
        "post"         : ("Chat → Send text to recipient",              "MS Teams", "Post"),
        "meet_now_post": ("Meet Now → Start meeting → Chat tab → Post", "MS Teams", "Post"),
        "forward"      : ("Chat → 3 dots → Forward message",            "MS Teams", "Post"),
        "reply"        : ("Chat → 3 dots → Reply to message",           "MS Teams", "Post"),
    }

    def _nav(r):
        aname = (r.get("activity_name") or "").lower().strip()
        # exact match first
        if aname in _NAV_DETAILS:
            return _NAV_DETAILS[aname]
        for key, val in _NAV_DETAILS.items():
            if key in aname:
                return val
        return _NAV_DETAILS["post"]

    total      = len(all_results)
    passed     = sum(1 for r in all_results if r.get("status") == "PASS")
    failed     = total - passed
    overall_bg = "#2e7d32" if overall_status == "PASS" else "#c62828"
    emoji      = "✅" if overall_status == "PASS" else "❌"

    rows_html = ""
    for i, r in enumerate(all_results):
        st   = r.get("status", "FAIL")
        sc   = "#2e7d32" if st == "PASS" else "#c62828"
        fl   = r.get("fast_log_confirmed", False)
        fls  = r.get("fast_log_skipped", False)
        flok = "SKIP" if fls else ("YES ✓" if fl else "NO ✗")
        flc  = "#e65100" if fls else ("#2e7d32" if fl else "#c62828")
        dval = "NO ✗"  if r.get("message_not_delivered") else "YES ✓"
        dclr = "#2e7d32" if r.get("message_not_delivered") else "#c62828"
        cb   = "✓" if r.get("blocked_by_casb") else "✗"
        cbc  = "#2e7d32" if r.get("blocked_by_casb") else "#c62828"
        nav_detail, app_name, act_name = _nav(r)
        sig_ids    = r.get("fast_log_sig_ids", [])
        multi_sigs = r.get("fast_log_multi_sigs", False)
        if not sig_ids:
            sig_cell, sig_color = "—", "#999"
        elif multi_sigs:
            sig_cell, sig_color = "⚠ " + ", ".join(sig_ids), "#e65100"
        else:
            sig_cell, sig_color = sig_ids[0], "#1565c0"
        bg = "#ffffff" if i % 2 == 0 else "#f9f9f9"
        rows_html += f"""<tr style="background:{bg}">
          <td style="padding:9px 8px;font-size:12px;font-weight:600;color:#1a1a1a;border-bottom:1px solid #e0e0e0">{app_name}</td>
          <td style="padding:9px 8px;font-size:12px;color:#1a1a1a;border-bottom:1px solid #e0e0e0">{act_name}</td>
          <td style="padding:9px 8px;font-size:11px;color:#444;border-bottom:1px solid #e0e0e0">{nav_detail}</td>
          <td style="padding:9px 8px;font-size:11px;font-family:monospace;color:{sig_color};border-bottom:1px solid #e0e0e0;white-space:nowrap">{sig_cell}</td>
          <td style="padding:9px 8px;font-size:13px;font-weight:700;color:{cbc};text-align:center;border-bottom:1px solid #e0e0e0">{cb}</td>
          <td style="padding:9px 8px;font-size:11px;font-weight:600;color:{dclr};text-align:center;border-bottom:1px solid #e0e0e0">{dval}</td>
          <td style="padding:9px 8px;font-size:11px;font-weight:600;color:{flc};text-align:center;border-bottom:1px solid #e0e0e0">{flok}</td>
          <td style="padding:9px 8px;font-size:12px;font-weight:700;color:{sc};text-align:center;border-bottom:1px solid #e0e0e0">{st}</td>
        </tr>"""

    total_color = "#2e7d32" if failed == 0 else "#c62828"
    rows_html += f"""<tr style="background:#eeeeee">
          <td colspan="3" style="padding:9px 8px;font-size:12px;font-weight:700;color:#1a1a1a;border-top:2px solid #bbb">TOTAL</td>
          <td colspan="4" style="padding:9px 8px;font-size:11px;color:#555;border-top:2px solid #bbb">{passed} passed, {failed} failed out of {total}</td>
          <td style="padding:9px 8px;font-size:12px;font-weight:700;color:{total_color};text-align:center;border-top:2px solid #bbb">{passed}/{total} ({int(passed/total*100) if total else 0}%)</td>
        </tr>"""

    return f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"/></head>
<body style="margin:0;padding:0;background:#f5f5f5;font-family:Arial,sans-serif">
<table width="100%" cellpadding="0" cellspacing="0" style="padding:20px 0">
<tr><td align="center">
<table width="700" cellpadding="0" cellspacing="0"
       style="background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.1);max-width:100%">
  <tr><td style="background:#1565c0;padding:22px 24px">
    <table width="100%" cellpadding="0" cellspacing="0"><tr>
      <td><div style="font-size:18px;font-weight:700;color:#fff">🛡 CASB Block Verification Report</div>
          <div style="font-size:11px;color:#90caf9;margin-top:3px">MS Teams · Versa SASE · fast.log verification</div></td>
      <td align="right"><span style="background:{overall_bg};color:#fff;font-size:14px;font-weight:700;padding:7px 16px;border-radius:5px">{emoji} {overall_status}</span></td>
    </tr></table>
    <table width="100%" cellpadding="0" cellspacing="0" style="margin-top:14px"><tr>
      <td width="32%" style="background:rgba(255,255,255,0.12);border-radius:5px;padding:10px;text-align:center">
        <div style="font-size:26px;font-weight:800;color:#fff">{total}</div>
        <div style="font-size:10px;color:#90caf9;text-transform:uppercase">Total</div></td>
      <td width="2%"></td>
      <td width="32%" style="background:rgba(0,200,83,0.2);border-radius:5px;padding:10px;text-align:center">
        <div style="font-size:26px;font-weight:800;color:#69f0ae">{passed}</div>
        <div style="font-size:10px;color:#69f0ae;text-transform:uppercase">✔ Passed</div></td>
      <td width="2%"></td>
      <td width="32%" style="background:rgba(255,23,68,0.2);border-radius:5px;padding:10px;text-align:center">
        <div style="font-size:26px;font-weight:800;color:#ff8a80">{failed}</div>
        <div style="font-size:10px;color:#ff8a80;text-transform:uppercase">✘ Failed</div></td>
    </tr></table>
  </td></tr>
  <tr><td style="padding:10px 24px;background:#e3f2fd;border-bottom:1px solid #bbdefb">
    <span style="font-size:12px;color:#555">Run timestamp: </span>
    <span style="font-size:12px;font-weight:700;color:#1565c0;font-family:monospace">{run_ts}</span>
  </td></tr>
  <tr><td style="padding:18px 24px 8px">
    <div style="font-size:13px;font-weight:700;color:#1565c0;text-transform:uppercase;letter-spacing:1px;border-bottom:2px solid #1565c0;padding-bottom:6px">Final Summary</div>
  </td></tr>
  <tr><td style="padding:0 24px 24px">
    <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse">
      <thead><tr style="background:#1565c0">
        <th style="padding:9px 8px;color:#fff;text-align:left;font-size:10px;text-transform:uppercase">App</th>
        <th style="padding:9px 8px;color:#fff;text-align:left;font-size:10px;text-transform:uppercase">Activity</th>
        <th style="padding:9px 8px;color:#fff;text-align:left;font-size:10px;text-transform:uppercase">Navigation</th>
        <th style="padding:9px 8px;color:#fff;text-align:left;font-size:10px;text-transform:uppercase">Sig ID</th>
        <th style="padding:9px 8px;color:#fff;text-align:center;font-size:10px;text-transform:uppercase">Block</th>
        <th style="padding:9px 8px;color:#fff;text-align:center;font-size:10px;text-transform:uppercase">Activity Delivered</th>
        <th style="padding:9px 8px;color:#fff;text-align:center;font-size:10px;text-transform:uppercase">Log</th>
        <th style="padding:9px 8px;color:#fff;text-align:center;font-size:10px;text-transform:uppercase">Result</th>
      </tr></thead>
      <tbody>{rows_html}</tbody>
    </table>
  </td></tr>
  <tr><td style="padding:12px 24px;background:#f5f5f5;border-top:1px solid #e0e0e0;text-align:center;font-size:11px;color:#999">
    Full report attached as ZIP · Generated by CASB Automation Script · {run_ts}
  </td></tr>
</table>
</td></tr></table>
</body></html>"""


def _send_report_email(recipients, html_report_path, run_folder_path,
                       overall_status, run_ts, all_results=None):
    import smtplib, zipfile, tempfile
    from email.mime.multipart import MIMEMultipart
    from email.mime.text      import MIMEText
    from email.mime.base      import MIMEBase
    from email                import encoders

    if not recipients:
        return

    print(f"\n{'=' * 55}")
    print(f"SENDING REPORT EMAIL to: {', '.join(recipients)}")
    print(f"{'=' * 55}")

    # Build clean email-safe HTML summary (not the dark theme report)
    if all_results:
        html_body = _build_email_html(all_results, overall_status, run_ts)
    else:
        try:
            with open(html_report_path, "r", encoding="utf-8") as f:
                html_body = f.read()
        except Exception as e:
            print(f"   [EMAIL] Could not read HTML report: {e}")
            return

    # Zip run folder as attachment
    zip_path = None
    try:
        tmp      = tempfile.mktemp(suffix=".zip")
        run_name = os.path.basename(run_folder_path)
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
            for root, dirs, files in os.walk(run_folder_path):
                for file in files:
                    abs_path = os.path.join(root, file)
                    arc_name = os.path.relpath(abs_path, os.path.dirname(run_folder_path))
                    zf.write(abs_path, arc_name)
        zip_path = tmp
    except Exception as e:
        print(f"   [EMAIL] Could not zip run folder: {e}")

    status_emoji = "✅" if overall_status == "PASS" else "❌"
    subject      = f"{status_emoji} CASB Test Report — {overall_status} — {run_ts}"
    msg          = MIMEMultipart("mixed")
    msg["From"]    = _cfg.SENDER_EMAIL
    msg["To"]      = ", ".join(recipients)
    msg["Subject"] = subject
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    if zip_path:
        try:
            with open(zip_path, "rb") as f:
                part = MIMEBase("application", "zip")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            zip_name = f"CASB_Report_{run_ts.replace(' ','_').replace(':','')}.zip"
            part.add_header("Content-Disposition", f'attachment; filename="{zip_name}"')
            msg.attach(part)
            print(f"   [EMAIL] Attached: {zip_name}")
        except Exception as e:
            print(f"   [EMAIL] Could not attach zip: {e}")

    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.ehlo()
            server.starttls()
            server.login(_cfg.SENDER_EMAIL, _cfg.SENDER_GMAIL_APP_PASSWORD)
            server.sendmail(_cfg.SENDER_EMAIL, recipients, msg.as_string())
        print(f"   [EMAIL] ✓ Sent to: {', '.join(recipients)}")
    except Exception as e:
        print(f"   [EMAIL] ✗ Failed: {e}")
    print(f"{'=' * 55}\n")


def _upload_to_server(run_folder, server_url):
    import requests, zipfile, tempfile
    print(f"\n{'=' * 55}")
    print(f"UPLOADING TO CASB RESULTS SERVER")
    print(f"  Server : {server_url}")
    print(f"  Folder : {run_folder}")
    print(f"{'=' * 55}")
    run_name = os.path.basename(run_folder)
    tmp = tempfile.mktemp(suffix=".zip")
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
            for root, dirs, files in os.walk(run_folder):
                for file in files:
                    abs_path = os.path.join(root, file)
                    arc_name = os.path.join(run_name, os.path.relpath(abs_path, run_folder))
                    zf.write(abs_path, arc_name)
        with open(tmp, "rb") as f:
            resp = requests.post(
                f"{server_url.rstrip('/')}/api/upload",
                files={"file": (f"{run_name}.zip", f, "application/zip")},
                timeout=120,
            )
        data = resp.json()
        if data.get("status") == "ok":
            print(f"   [UPLOAD] ✓ Success!")
        else:
            print(f"   [UPLOAD] ✗ Server error: {data}")
    except Exception as e:
        print(f"   [UPLOAD] ✗ Failed: {e}")
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    print(f"{'=' * 55}\n")


# ── Run each app ──────────────────────────────────────────────────────────────

from playwright.sync_api import sync_playwright
from core.runner import run_all
from core.report_generator import save_report, generate_html_report

_send_email_list = [e.strip() for e in args.send_email.split(",")] if args.send_email else []
_server_url      = args.server_url

with sync_playwright() as pw:
    for app_id, at_list in parsed_apps:
        for account_type in at_list:
            print(f"\n{'=' * 55}")
            print(f"  App          : {app_id}")
            print(f"  Account type : {account_type}")
            print(f"  Activities   : {run_navs}")
            print(f"{'=' * 55}\n")

            # Launch browser (persistent context for cookie persistence)
            # NOTE: Do NOT add --proxy-server here — routing Chrome through
            # a local proxy (e.g. mitmproxy) bypasses Versa SASE entirely,
            # which means no CASB block, no popup, no fast.log hit.
            # HAR capture uses Playwright context listeners instead.
            _browser_args = [
                "--start-maximized",
            ]
            browser = pw.chromium.launch_persistent_context(
                user_data_dir=_cfg.SENDER_PROFILE_DIR,
                headless=False,
                args=_browser_args,
                no_viewport=True,
            )

            # Login (app-specific login handler)
            _login_path = os.path.join("apps", app_id, "login_handler.py")
            if os.path.exists(_login_path):
                import importlib.util as _ilu
                spec   = _ilu.spec_from_file_location(f"{app_id}.login", _login_path)
                mod    = _ilu.module_from_spec(spec)
                spec.loader.exec_module(mod)
                if hasattr(mod, "login"):
                    mod.login(browser, account_type, _cfg)

            # Run all TCs
            all_results = run_all(
                app_id       = app_id,
                account_type = account_type,
                browser      = browser,
                script_dir   = _cfg.SCRIPT_DIR,
                run_navs     = run_navs,
                run_tc_nums  = run_tc_nums,
                config_module= _cfg,
            )

            # Report
            overall = "PASS" if all(r.get("status") == "PASS" for r in all_results) else "FAIL"
            _cfg.REPORT_DATA["run_status"] = overall
            save_report(all_results)
            generate_html_report(all_results)

            passed = sum(1 for r in all_results if r.get("status") == "PASS")
            total  = len(all_results)
            failed = total - passed

            # ── Final summary table ───────────────────────────────
            _TC_LABELS = {
                "post"          : ("TC1", "Post", "Chat → Send text to recipient"),
                "meet_now_post" : ("TC2", "Post", "Meet Now → Start meeting → Chat → Post"),
                "forward"       : ("TC3", "Post", "Chat → 3 dots → Forward message"),
                "reply"         : ("TC4", "Post", "Chat → 3 dots → Reply to message"),
            }

            print("\n" + "=" * 78)
            print(f"  FINAL SUMMARY  |  OVERALL: {overall}  ({passed}/{total} passed, {failed} failed)")
            print("=" * 78)
            print(f"  {'TC':<6}  {'Activity':<10}  {'Navigation':<38}  {'Block':<7}  {'Delivered':<10}  {'Log':<6}  Result")
            print(f"  {'-'*6}  {'-'*10}  {'-'*38}  {'-'*7}  {'-'*10}  {'-'*6}  ------")
            for r in all_results:
                # Use exact match on activity_name first, then fall back to substring
                aname  = (r.get("activity_name") or "").lower().strip()
                tc_key = aname if aname in _TC_LABELS else next(
                    (k for k in _TC_LABELS if k in aname), "post"
                )
                tc, act, nav = _TC_LABELS.get(tc_key, ("TC?", "?", "?"))
                st        = r.get("status", "FAIL")
                block     = "YES ✓" if r.get("blocked_by_casb") else "NO ✗"
                delivered = "NO ✗"  if r.get("message_not_delivered") else "YES ✓"
                fl        = r.get("fast_log_confirmed", False)
                fls       = r.get("fast_log_skipped", False)
                log       = "SKIP" if fls else ("YES ✓" if fl else "NO ✗")
                nav_s     = (nav[:36] + "..") if len(nav) > 38 else nav
                print(f"  {tc:<6}  {act:<10}  {nav_s:<38}  {block:<7}  {delivered:<10}  {log:<6}  {st}")
            print("-" * 78)
            print(f"  TOTAL: {passed} passed, {failed} failed out of {total}  ({int(passed/total*100) if total else 0}%)")
            print("=" * 78)
            print(f"  HTML Report : {_cfg.HTML_REPORT}")
            print("=" * 78)

            # ── Send email report ─────────────────────────────────
            if _send_email_list:
                _send_report_email(
                    recipients       = _send_email_list,
                    html_report_path = _cfg.HTML_REPORT,
                    run_folder_path  = _cfg.SCRIPT_DIR,
                    overall_status   = overall,
                    run_ts           = _cfg.REPORT_DATA.get("run_timestamp", ""),
                    all_results      = all_results,
                )

            # ── Upload to CASB Results Server ─────────────────────
            if _server_url:
                _upload_to_server(_cfg.SCRIPT_DIR, _server_url)
                print(f"\n   🌐 View results at: {_server_url}/run/{os.path.basename(_cfg.SCRIPT_DIR)}")

            browser.close()