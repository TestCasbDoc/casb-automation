"""
CASB Results Server — MS Teams CASB Block Verification Dashboard
Run: python app.py
Access: http://0.0.0.0:4012/

Stores run results uploaded from ms_teams_personal_send_post.py.
Each run is a folder under RESULTS_DIR containing:
  - test_report.json
  - test_report.html
  - vos_dumps/
  - har_files/
  - *.png screenshots
"""

import os
import json
import zipfile
import shutil
from datetime import datetime
from flask import (
    Flask, render_template, request, redirect, url_for,
    send_file, abort, jsonify
)
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024  # 500MB max upload

# ── Storage ──────────────────────────────────────────────────────────────────
RESULTS_DIR = os.path.join(os.path.dirname(__file__), "results")
os.makedirs(RESULTS_DIR, exist_ok=True)

# ── Helpers ───────────────────────────────────────────────────────────────────

def _load_run(run_id):
    """Load test_report.json for a run. Returns dict or None."""
    path = os.path.join(RESULTS_DIR, run_id, "test_report.json")
    if not os.path.exists(path):
        return None
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        # Normalise: new format is {run_timestamp, config, results:[...]}
        # Old format is just a list
        if isinstance(data, list):
            return {"results": data, "config": {}, "run_timestamp": "", "run_status": "UNKNOWN"}
        return data
    except Exception:
        return None


def _run_summary(run_id):
    """Return a lightweight summary dict for the runs list page."""
    data    = _load_run(run_id)
    run_dir = os.path.join(RESULTS_DIR, run_id)

    try:
        ts_fmt = datetime.strptime(run_id, "run_%Y%m%d_%H%M%S").strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        ts_fmt = run_id

    if data is None:
        return {
            "run_id": run_id, "timestamp": ts_fmt,
            "total": 0, "passed": 0, "failed": 0, "status": "UNKNOWN",
            "sig_ids": [], "has_html": False, "config": {}, "trigger_pct": "0%",
        }

    results = data.get("results", [])
    cfg     = data.get("config", {})
    total   = len(results)
    passed  = sum(1 for r in results if r.get("status") == "PASS")
    failed  = total - passed
    status  = "PASS" if failed == 0 and total > 0 else "FAIL"

    sig_ids = []
    seen = set()
    for r in results:
        for sid in r.get("fast_log_sig_ids", []):
            if sid not in seen:
                seen.add(sid)
                sig_ids.append(sid)

    has_html = os.path.exists(os.path.join(run_dir, "test_report.html"))

    return {
        "run_id"      : run_id,
        "timestamp"   : ts_fmt,
        "total"       : total,
        "passed"      : passed,
        "failed"      : failed,
        "status"      : status,
        "sig_ids"     : sig_ids,
        "has_html"    : has_html,
        "trigger_pct" : f"{(passed/total*100):.1f}%" if total else "0%",
        "config"      : cfg,
    }


def _all_runs():
    """Return list of run summary dicts sorted newest first."""
    runs = []
    for name in os.listdir(RESULTS_DIR):
        if os.path.isdir(os.path.join(RESULTS_DIR, name)):
            runs.append(_run_summary(name))
    runs.sort(key=lambda r: r["run_id"], reverse=True)
    return runs


def _tc_table(run_id):
    """Return per-TC rows for the run detail page."""
    data = _load_run(run_id)
    if not data:
        return []
    results = data.get("results", []) if isinstance(data, dict) else data

    nav_map = {
        "meet now":  "TC2 — Meet Now Post",
        "forward":   "TC3 — Forward Message",
        "reply":     "TC4 — Reply to Message",
    }

    rows = []
    for r in results:
        aname = (r.get("activity_name") or "").lower()
        nav = "TC1 — Direct Chat Post"
        for k, v in nav_map.items():
            if k in aname:
                nav = v
                break

        sig_ids      = r.get("fast_log_sig_ids", [])
        multi_sigs   = r.get("fast_log_multi_sigs", False)
        casb_blocked = r.get("blocked_by_casb", False)
        delivered    = not r.get("message_not_delivered", True)
        fast_log_ok  = r.get("fast_log_confirmed", False)
        fast_skipped = r.get("fast_log_skipped", False)
        status       = r.get("status", "FAIL")
        fail_reasons = r.get("fail_reason", [])

        rows.append({
            "timestamp"   : r.get("timestamp", ""),
            "nav"         : nav,
            "sig_ids"     : sig_ids,
            "multi_sigs"  : multi_sigs,
            "casb_blocked": casb_blocked,
            "delivered"   : delivered,
            "fast_log"    : "SKIPPED" if fast_skipped else ("YES" if fast_log_ok else "NO"),
            "fast_log_ok" : fast_log_ok,
            "fast_skipped": fast_skipped,
            "status"      : status,
            "fail_reasons": fail_reasons,
        })
    return rows


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    runs = _all_runs()
    total_runs   = len(runs)
    total_pass   = sum(1 for r in runs if r["status"] == "PASS")
    total_fail   = total_runs - total_pass
    return render_template("index.html",
                           runs=runs,
                           total_runs=total_runs,
                           total_pass=total_pass,
                           total_fail=total_fail)


@app.route("/run/<run_id>")
def run_detail(run_id):
    run_id = secure_filename(run_id)
    run_dir = os.path.join(RESULTS_DIR, run_id)
    if not os.path.isdir(run_dir):
        abort(404)
    summary = _run_summary(run_id)
    rows    = _tc_table(run_id)
    # VOS dump files
    dump_dir   = os.path.join(run_dir, "vos_dumps")
    dump_files = sorted(os.listdir(dump_dir)) if os.path.isdir(dump_dir) else []
    # HAR files
    har_dir    = os.path.join(run_dir, "har_files")
    har_files  = sorted(os.listdir(har_dir)) if os.path.isdir(har_dir) else []
    # Screenshots
    screenshots = sorted([
        f for f in os.listdir(run_dir)
        if f.lower().endswith(".png")
    ])
    return render_template("run_detail.html",
                           summary=summary,
                           rows=rows,
                           run_id=run_id,
                           dump_files=dump_files,
                           har_files=har_files,
                           screenshots=screenshots)


@app.route("/run/<run_id>/report")
def view_report(run_id):
    run_id   = secure_filename(run_id)
    html_path = os.path.join(RESULTS_DIR, run_id, "test_report.html")
    if not os.path.exists(html_path):
        abort(404)
    return send_file(html_path)


@app.route("/run/<run_id>/download")
def download_run(run_id):
    run_id  = secure_filename(run_id)
    run_dir = os.path.join(RESULTS_DIR, run_id)
    if not os.path.isdir(run_dir):
        abort(404)
    zip_path = os.path.join(RESULTS_DIR, f"{run_id}.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(run_dir):
            for file in files:
                abs_path = os.path.join(root, file)
                arc_name = os.path.relpath(abs_path, RESULTS_DIR)
                zf.write(abs_path, arc_name)
    return send_file(zip_path, as_attachment=True, download_name=f"{run_id}.zip")


@app.route("/run/<run_id>/file/<path:filename>")
def download_file(run_id, filename):
    run_id   = secure_filename(run_id)
    run_dir  = os.path.join(RESULTS_DIR, run_id)
    file_path = os.path.join(run_dir, filename)
    # Security: ensure path stays inside run_dir
    if not os.path.abspath(file_path).startswith(os.path.abspath(run_dir)):
        abort(403)
    if not os.path.exists(file_path):
        abort(404)
    return send_file(file_path, as_attachment=True)


@app.route("/run/<run_id>/delete", methods=["POST"])
def delete_run(run_id):
    run_id  = secure_filename(run_id)
    run_dir = os.path.join(RESULTS_DIR, run_id)
    if os.path.isdir(run_dir):
        shutil.rmtree(run_dir)
    zip_path = os.path.join(RESULTS_DIR, f"{run_id}.zip")
    if os.path.exists(zip_path):
        os.remove(zip_path)
    return redirect(url_for("index"))


# ── Upload API (called from ms_teams_personal_send_post.py) ──────────────────

@app.route("/api/upload", methods=["POST"])
def upload_run():
    """
    POST a zip of the entire run folder.
    The zip should have run_YYYYMMDD_HHMMSS/ as its root folder.

    curl example:
      curl -X POST http://HOST:4012/api/upload \
           -F "file=@run_20260318_140357.zip"
    """
    if "file" not in request.files:
        return jsonify({"error": "No file field"}), 400
    f = request.files["file"]
    if not f.filename.endswith(".zip"):
        return jsonify({"error": "Must be a .zip file"}), 400

    tmp_path = os.path.join(RESULTS_DIR, "_upload_tmp.zip")
    f.save(tmp_path)

    try:
        with zipfile.ZipFile(tmp_path, "r") as zf:
            # Detect run folder name from zip root
            names = zf.namelist()
            if not names:
                return jsonify({"error": "Empty zip"}), 400
            run_id = names[0].split("/")[0]
            if not run_id.startswith("run_"):
                return jsonify({"error": f"Unexpected root folder: {run_id}"}), 400
            zf.extractall(RESULTS_DIR)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

    return jsonify({"status": "ok", "run_id": run_id,
                    "url": f"/run/{run_id}"}), 200


@app.route("/api/runs")
def api_runs():
    return jsonify(_all_runs())


if __name__ == "__main__":
    print("=" * 55)
    print("  CASB Results Server")
    print("  http://0.0.0.0:4012/")
    print("  Results stored in:", os.path.abspath(RESULTS_DIR))
    print("=" * 55)
    app.run(host="0.0.0.0", port=4012, debug=False)