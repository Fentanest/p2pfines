from __future__ import annotations

import json
import os
import re
from typing import Any

from .paths import DATA_ROOT, RESOURCE_ROOT
from .schema import WATCH_REPORT_SPREADSHEET_ID, WORK_REPORT_SPREADSHEET_ID


ROOT = RESOURCE_ROOT
CONFIG_PATH = DATA_ROOT / "config.json"
UPLOAD_DIR = DATA_ROOT / ".uploads"
CREDENTIAL_DIR = DATA_ROOT / "credentials"
OUTPUT_DIR = DATA_ROOT / "outputs"


def _default_key_path() -> str:
    candidates = [*sorted(CREDENTIAL_DIR.glob("*.json")), *sorted(DATA_ROOT.glob("*.json"))]
    for path in candidates:
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            continue
        if data.get("type") == "service_account":
            return str(path)
    return ""


def _default_oauth_path() -> str:
    candidates = [*sorted(CREDENTIAL_DIR.glob("*.json")), *sorted(DATA_ROOT.glob("*.json"))]
    for path in candidates:
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            continue
        if isinstance(data.get("web"), dict) or isinstance(data.get("installed"), dict):
            return str(path)
    return ""


DEFAULT_CONFIG: dict[str, Any] = {
    "work_report_spreadsheet_id": WORK_REPORT_SPREADSHEET_ID,
    "watch_report_spreadsheet_id": WATCH_REPORT_SPREADSHEET_ID,
    "service_account_json": _default_key_path(),
    "apps_script_oauth_json": _default_oauth_path(),
    "default_year": 2026,
    "default_month": 5,
    "default_excel_password": os.environ.get("REPORT_GENERATOR_EXCEL_PASSWORD", ""),
}


def extract_spreadsheet_id(value: Any) -> str:
    text = str(value or "").strip()
    if not text:
        return ""

    markdown_link = re.search(r"\((https?://[^)]+)\)", text)
    if markdown_link:
        text = markdown_link.group(1).strip()

    for pattern in (
        r"/spreadsheets/d/([A-Za-z0-9_-]+)",
        r"/d/([A-Za-z0-9_-]+)",
        r"[?&]id=([A-Za-z0-9_-]+)",
        r"[?&]key=([A-Za-z0-9_-]+)",
    ):
        match = re.search(pattern, text)
        if match:
            return match.group(1)

    text = text.split("#", 1)[0].split("?", 1)[0].strip().strip("/")
    if re.fullmatch(r"[A-Za-z0-9_-]{20,}", text):
        return text
    return text


def normalize_spreadsheet_ids(cfg: dict[str, Any]) -> dict[str, Any]:
    out = dict(cfg)
    for key in ("work_report_spreadsheet_id", "watch_report_spreadsheet_id"):
        if key in out:
            out[key] = extract_spreadsheet_id(out[key])
    return out


def load_config() -> dict[str, Any]:
    if not CONFIG_PATH.exists():
        save_config(DEFAULT_CONFIG)
        return dict(DEFAULT_CONFIG)
    data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    merged = dict(DEFAULT_CONFIG)
    merged.update(data)
    return normalize_spreadsheet_ids(merged)


def save_config(data: dict[str, Any]) -> dict[str, Any]:
    cfg = dict(DEFAULT_CONFIG)
    cfg.update(data)
    cfg = normalize_spreadsheet_ids(cfg)
    CONFIG_PATH.write_text(
        json.dumps(cfg, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return cfg


def ensure_dirs() -> None:
    DATA_ROOT.mkdir(parents=True, exist_ok=True)
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    CREDENTIAL_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
