"""
app.py  —  KTS Excel Export Backend
=====================================
Python Flask server that:
  1. Receives JSON export data from the HTML frontend
  2. Loads the real .xlsx template (never modifies it)
  3. Writes ONLY approved yellow input cells
  4. Returns a finished, download-ready .xlsx file

Run locally:
    pip install flask openpyxl
    python app.py

Deploy (Render / Railway / Fly.io):
    See README.md for deployment instructions.

CORS:
  Set ALLOWED_ORIGIN in the environment to your GitHub Pages URL.
  Example: ALLOWED_ORIGIN=https://kts12717.github.io
"""

from __future__ import annotations

import logging
import os
from datetime import date
from pathlib import Path

from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
import io

from exporters.ha0935 import build_ha0935, HA0935Data
from exporters.hv0713 import build_hv0713, HV0713Data

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)-8s %(name)s  %(message)s",
)
logger = logging.getLogger(__name__)

# ── App setup ─────────────────────────────────────────────────────────────────
app = Flask(__name__)

# Allow requests only from your GitHub Pages domain in production.
# Set ALLOWED_ORIGIN env var; default allows all origins for local dev.
ALLOWED_ORIGIN = os.environ.get("ALLOWED_ORIGIN", "*")
CORS(app, resources={r"/api/*": {"origins": ALLOWED_ORIGIN}})

# ── Template paths ────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
TEMPLATES_DIR = BASE_DIR / "templates_excel"

HA0935_TEMPLATE = TEMPLATES_DIR / "HA0935_template.xlsx"
HV0713_TEMPLATE = TEMPLATES_DIR / "HV0713_template.xlsx"

# ── Helpers ───────────────────────────────────────────────────────────────────

def _parse_date(raw: str) -> date:
    """Parse 'YYYY-MM-DD' → datetime.date.  Raises ValueError on bad input."""
    from datetime import datetime
    d = datetime.strptime(raw.strip(), "%Y-%m-%d").date()
    if d.day != 1:
        raise ValueError(f"month_date must be the first of the month, got {raw!r}")
    return d


def _parse_ovo(raw: dict) -> dict:
    """
    Convert JSON OVO payload to the format expected by the exporters.
    JSON keys are strings; we need integer day keys.
    Input:  {"6770": {"AM": {"2": 1, "3": 1, ...}, "PM": {...}}}
    Output: {"6770": {"AM": {2: 1, 3: 1, ...}, "PM": {...}}}
    """
    result = {}
    for route, periods in (raw or {}).items():
        result[str(route)] = {
            period.upper(): {int(k): int(v) for k, v in days.items()}
            for period, days in periods.items()
        }
    return result


def _parse_sso(raw: dict) -> dict:
    """
    Convert JSON SSO payload.
    Input:  {"6770": {"2": 251.0, "3": 243.0}}
    Output: {"6770": {2: 251.0, 3: 243.0}}
    """
    result = {}
    for route, days in (raw or {}).items():
        result[str(route)] = {int(k): float(v) for k, v in days.items()}
    return result


def _parse_attendance(raw: dict) -> dict:
    """
    Convert JSON attendance payload.
    Input:  {"13": {"2": 2, "3": "NS", ...}}
    Output: {13: {2: 2, 3: "NS", ...}}
    """
    result = {}
    for row_str, days in (raw or {}).items():
        parsed_days = {}
        for day_str, code in days.items():
            # Numeric codes come as ints or int-strings; text codes as strings
            if str(code).lstrip("-").isdigit():
                parsed_days[int(day_str)] = int(code)
            else:
                parsed_days[int(day_str)] = str(code).upper()
        result[int(row_str)] = parsed_days
    return result


def _xlsx_response(xlsx_bytes: bytes, filename: str):
    """Wrap bytes in a Flask send_file response."""
    return send_file(
        io.BytesIO(xlsx_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/api/health", methods=["GET"])
def health():
    """Health check — confirm both templates are accessible."""
    status = {
        "ha0935_template": HA0935_TEMPLATE.is_file(),
        "hv0713_template": HV0713_TEMPLATE.is_file(),
        "status": "ok",
    }
    return jsonify(status)


@app.route("/api/export/ha0935", methods=["POST"])
def export_ha0935():
    """
    Export endpoint for HA0935 Khan Transportation.

    POST body (JSON):
    {
      "month_date": "2026-03-01",
      "ovo": {
        "Khan1": {"AM": {"2":1,"3":1,...}, "PM": {"2":1,...}},
        ...
        "Khan10": {...}
      },
      "sso": {
        "Khan1": {"2":251.0,"3":243.0,...},
        ...
      },
      "attendance": {
        "13": {"2":2,"3":2,"4":"NS",...},
        "14": {"2":2,...},
        ...
      },
      "vehicle_unit_rate": 166.94,   // optional
      "contract_mile_rate": 1.08     // optional
    }

    Returns: .xlsx file download
    """
    if not request.is_json:
        return jsonify({"error": "Content-Type must be application/json"}), 400

    body = request.get_json(silent=True) or {}

    try:
        month_date = _parse_date(body.get("month_date", ""))
        data = HA0935Data(
            month_date         = month_date,
            ovo                = _parse_ovo(body.get("ovo", {})),
            sso                = _parse_sso(body.get("sso", {})),
            attendance         = _parse_attendance(body.get("attendance", {})),
            vehicle_unit_rate  = body.get("vehicle_unit_rate"),
            contract_mile_rate = body.get("contract_mile_rate"),
        )
    except (ValueError, TypeError, KeyError) as e:
        logger.warning("HA0935 request validation error: %s", e)
        return jsonify({"error": str(e)}), 422

    try:
        xlsx_bytes = build_ha0935(HA0935_TEMPLATE, data)
    except FileNotFoundError as e:
        logger.error("Template file missing: %s", e)
        return jsonify({"error": "Template file not found on server"}), 500
    except RuntimeError as e:
        logger.error("Export error: %s", e)
        return jsonify({"error": str(e)}), 500

    filename = f"HA0935_Khan_{month_date.strftime('%b%Y')}.xlsx"
    logger.info("HA0935 exported: %s (%d bytes)", filename, len(xlsx_bytes))
    return _xlsx_response(xlsx_bytes, filename)


@app.route("/api/export/hv0713", methods=["POST"])
def export_hv0713():
    """
    Export endpoint for HV0713 Priority Transportation.

    POST body (JSON):
    {
      "month_date": "2026-03-01",
      "ovo": {
        "6770": {"AM": {"2":1,"3":1,...}, "PM": {"2":1,...}},
        "6771": {"AM": {...}, "PM": {...}}
      },
      "sso": {
        "6770": {"2":191.4,"3":188.4,...},
        "6771": {"2":250.5,...}
      },
      "attendance": {
        "13": {"2":2,"3":2,...},
        ...
      },
      "vehicle_unit_rate": 166.94,   // optional
      "contract_mile_rate": 1.08     // optional
    }

    Returns: .xlsx file download
    """
    if not request.is_json:
        return jsonify({"error": "Content-Type must be application/json"}), 400

    body = request.get_json(silent=True) or {}

    try:
        month_date = _parse_date(body.get("month_date", ""))
        data = HV0713Data(
            month_date         = month_date,
            ovo                = _parse_ovo(body.get("ovo", {})),
            sso                = _parse_sso(body.get("sso", {})),
            attendance         = _parse_attendance(body.get("attendance", {})),
            vehicle_unit_rate  = body.get("vehicle_unit_rate"),
            contract_mile_rate = body.get("contract_mile_rate"),
        )
    except (ValueError, TypeError, KeyError) as e:
        logger.warning("HV0713 request validation error: %s", e)
        return jsonify({"error": str(e)}), 422

    try:
        xlsx_bytes = build_hv0713(HV0713_TEMPLATE, data)
    except FileNotFoundError as e:
        logger.error("Template file missing: %s", e)
        return jsonify({"error": "Template file not found on server"}), 500
    except RuntimeError as e:
        logger.error("Export error: %s", e)
        return jsonify({"error": str(e)}), 500

    filename = f"HV0713_Priority_{month_date.strftime('%b%Y')}.xlsx"
    logger.info("HV0713 exported: %s (%d bytes)", filename, len(xlsx_bytes))
    return _xlsx_response(xlsx_bytes, filename)


# ── Dev server ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    logger.info("Starting KTS Excel Export Backend on port %d (debug=%s)", port, debug)
    app.run(host="0.0.0.0", port=port, debug=debug)
