from __future__ import annotations

import html
import json
import shutil
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
from flask import Flask, abort, jsonify, render_template, request, send_file

from jjukfund_app.apps_script_client import AppsScriptClient, AppsScriptInstallError
from jjukfund_app.apps_script_oauth import (
    AppsScriptOAuthError,
    build_authorization_url,
    finish_oauth_callback,
    load_user_credentials,
    oauth_status,
)
from jjukfund_app.config import CREDENTIAL_DIR, OUTPUT_DIR, UPLOAD_DIR, ensure_dirs, extract_spreadsheet_id, load_config, save_config
from jjukfund_app.excel_parser import (
    ParseError,
    PasswordRequired,
    load_cached_dataframe,
    load_cached_metadata,
    parse_uploaded_file,
)
from jjukfund_app.local_excel import LocalExcelError, generate_local_excel
from jjukfund_app.report_templates import setup_online_documents
from jjukfund_app.schema import (
    SOURCE_SPECS,
    WATCH_REPORT_SHEETS,
    WORK_REPORT_SHEETS,
    default_checked_columns,
    default_columns_for_target,
    missing_required_columns,
    source_options,
)
from jjukfund_app.sheets_client import GoogleSheetsClient, SheetsError
from jjukfund_app.paths import RESOURCE_ROOT


app = Flask(__name__, template_folder=str(RESOURCE_ROOT / "templates"))
ensure_dirs()
WATCH_PERIOD_SHEETS = {"(상시)1-1", "(상시)2-2", "(상시)3-1", "(상시)3-2"}
APPS_SCRIPT_FILES = {
    "work": RESOURCE_ROOT / "apps_script" / "Code_work.gs",
    "watch": RESOURCE_ROOT / "apps_script" / "Code_watch.gs",
}


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/api/health")
def health():
    return jsonify({"ok": True})


def _apps_script_path_for_target(target: str | None) -> Path:
    return APPS_SCRIPT_FILES.get(target or "work", APPS_SCRIPT_FILES["work"])


@app.get("/api/apps-script")
def apps_script_code():
    target = request.args.get("target", "work")
    path = _apps_script_path_for_target(target)
    return jsonify({"ok": True, "target": target if target in APPS_SCRIPT_FILES else "work", "code": path.read_text(encoding="utf-8")})


@app.get("/api/settings")
def get_settings():
    cfg = load_config()
    safe_cfg = {
        "work_report_spreadsheet_id": cfg.get("work_report_spreadsheet_id", ""),
        "watch_report_spreadsheet_id": cfg.get("watch_report_spreadsheet_id", ""),
        "service_account_json": cfg.get("service_account_json", ""),
        "apps_script_oauth_json": cfg.get("apps_script_oauth_json", ""),
        "default_year": cfg.get("default_year", 2026),
        "default_month": cfg.get("default_month", 5),
        "default_excel_password": cfg.get("default_excel_password", "118050"),
        "source_options": source_options("work"),
    }
    return jsonify(safe_cfg)


@app.post("/api/settings")
def update_settings():
    data = request.get_json(force=True)
    cfg = load_config()
    for key in [
        "work_report_spreadsheet_id",
        "watch_report_spreadsheet_id",
        "service_account_json",
        "apps_script_oauth_json",
        "default_year",
        "default_month",
        "default_excel_password",
    ]:
        if key in data:
            cfg[key] = extract_spreadsheet_id(data[key]) if key.endswith("_spreadsheet_id") else data[key]
    cfg["default_year"] = int(cfg.get("default_year") or 2026)
    cfg["default_month"] = int(cfg.get("default_month") or 1)
    save_config(cfg)
    warnings: list[str] = []
    if cfg.get("service_account_json"):
        try:
            client = GoogleSheetsClient(cfg["service_account_json"])
            if cfg.get("work_report_spreadsheet_id"):
                _write_online_settings(client, cfg["work_report_spreadsheet_id"], cfg, "work")
            if cfg.get("watch_report_spreadsheet_id"):
                _write_online_settings(client, cfg["watch_report_spreadsheet_id"], cfg, "watch")
        except Exception as exc:
            warnings.append(f"README 기준년월 갱신 실패: {exc}")
    return jsonify({"ok": True, "warnings": warnings})


@app.post("/api/key")
def upload_key():
    file = request.files.get("file")
    if not file:
        return jsonify({"ok": False, "error": "JSON 키 파일이 없습니다."}), 400
    if not file.filename.lower().endswith(".json"):
        return jsonify({"ok": False, "error": "JSON 파일만 업로드할 수 있습니다."}), 400

    token = uuid.uuid4().hex[:8]
    dest = CREDENTIAL_DIR / f"service-account-{token}.json"
    file.save(dest)

    try:
        data = json.loads(dest.read_text(encoding="utf-8"))
        if data.get("type") != "service_account":
            raise ValueError("service_account JSON이 아닙니다.")
    except Exception as exc:
        dest.unlink(missing_ok=True)
        return jsonify({"ok": False, "error": f"JSON 키 검증 실패: {exc}"}), 400

    cfg = load_config()
    cfg["service_account_json"] = str(dest)
    save_config(cfg)
    return jsonify({"ok": True, "path": str(dest), "client_email": data.get("client_email", "")})


@app.post("/api/oauth-key")
def upload_oauth_key():
    file = request.files.get("file")
    if not file:
        return jsonify({"ok": False, "error": "Apps Script OAuth JSON 파일이 없습니다."}), 400
    if not file.filename.lower().endswith(".json"):
        return jsonify({"ok": False, "error": "JSON 파일만 업로드할 수 있습니다."}), 400

    token = uuid.uuid4().hex[:8]
    dest = CREDENTIAL_DIR / f"apps-script-oauth-{token}.json"
    file.save(dest)

    try:
        data = json.loads(dest.read_text(encoding="utf-8"))
        client = data.get("web") if isinstance(data.get("web"), dict) else data.get("installed")
        if not isinstance(client, dict) or not client.get("client_id") or not client.get("client_secret"):
            raise ValueError("OAuth 클라이언트 JSON이 아닙니다.")
    except Exception as exc:
        dest.unlink(missing_ok=True)
        return jsonify({"ok": False, "error": f"OAuth JSON 검증 실패: {exc}"}), 400

    cfg = load_config()
    cfg["apps_script_oauth_json"] = str(dest)
    save_config(cfg)
    return jsonify({"ok": True, "path": str(dest), "client_id": client.get("client_id", "")})


@app.post("/api/setup")
def setup_documents():
    cfg = _merged_config_from_request()
    try:
        client = GoogleSheetsClient(cfg["service_account_json"])
        result = setup_online_documents(
            client,
            cfg["work_report_spreadsheet_id"],
            cfg["watch_report_spreadsheet_id"],
        )
        _write_online_settings(client, cfg["work_report_spreadsheet_id"], cfg, "work")
        _write_online_settings(client, cfg["watch_report_spreadsheet_id"], cfg, "watch")
        client.delete_sheets(cfg["work_report_spreadsheet_id"], ["설정", "컬럼정의"])
        client.delete_sheets(cfg["watch_report_spreadsheet_id"], ["설정", "컬럼정의"])
    except (SheetsError, Exception) as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    save_config(cfg)
    return jsonify({"ok": True, "result": result})


@app.get("/api/oauth/status")
def apps_script_oauth_status():
    cfg = load_config()
    status = oauth_status(cfg.get("apps_script_oauth_json"))
    if status.get("clientJson") and status["clientJson"] != cfg.get("apps_script_oauth_json"):
        cfg["apps_script_oauth_json"] = status["clientJson"]
        save_config(cfg)
    return jsonify({"ok": True, **status})


@app.post("/api/oauth/start")
def start_apps_script_oauth():
    cfg = _merged_config_from_request()
    try:
        auth = build_authorization_url(cfg.get("apps_script_oauth_json"))
    except AppsScriptOAuthError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    cfg["apps_script_oauth_json"] = auth["clientJson"]
    save_config(cfg)
    return jsonify({"ok": True, **auth})


@app.get("/oauth2callback")
def oauth2callback():
    error = request.args.get("error")
    if error:
        message = f"Google OAuth 인증 실패: {error}"
        return _oauth_callback_page(False, message), 400
    try:
        finish_oauth_callback(
            code=request.args.get("code", ""),
            state=request.args.get("state", ""),
        )
    except AppsScriptOAuthError as exc:
        return _oauth_callback_page(False, str(exc)), 400
    return _oauth_callback_page(True, "Google OAuth 인증이 완료되었습니다. 이 창을 닫고 설치 버튼을 다시 누르세요.")


@app.post("/api/install-script")
def install_script():
    cfg = _merged_config_from_request()
    installed: list[dict[str, str]] = []
    try:
        installer = AppsScriptClient(credentials=load_user_credentials())
        for target, spreadsheet_id, cfg_key in [
            ("work", cfg["work_report_spreadsheet_id"], "work_script_id"),
            ("watch", cfg["watch_report_spreadsheet_id"], "watch_script_id"),
        ]:
            script_id = cfg.get(cfg_key)
            if not script_id:
                script_id = installer.create_bound_project(spreadsheet_id, "Jjukfund Online Calculator")
                cfg[cfg_key] = script_id
            code = _apps_script_path_for_target(target).read_text(encoding="utf-8")
            installer.update_content(script_id, code)
            installed.append({"target": target, "scriptId": script_id})
    except AppsScriptOAuthError as exc:
        return jsonify({"ok": False, "oauthRequired": True, "error": str(exc)}), 401
    except AppsScriptInstallError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    save_config(cfg)
    return jsonify({"ok": True, "installed": installed})


@app.post("/api/parse")
def parse_file():
    cfg = load_config()
    file = request.files.get("file")
    if not file:
        return jsonify({"ok": False, "error": "파일이 없습니다."}), 400
    suffix = Path(file.filename).suffix
    token = uuid.uuid4().hex
    dest = UPLOAD_DIR / f"{token}{suffix}"
    file.save(dest)
    try:
        metadata = parse_uploaded_file(
            dest,
            file.filename,
            password=request.form.get("password") or None,
            report_target=request.form.get("reportTarget") or "work",
        )
        source_key = request.form.get("sourceKey") or ""
        report_target = request.form.get("reportTarget") or "work"
        effective_source_key = source_key if source_key in SOURCE_SPECS else metadata.get("sourceKey")
        if effective_source_key in SOURCE_SPECS:
            metadata["sourceKey"] = effective_source_key
            column_names = [column["name"] for column in metadata["columns"]]
            metadata["missingRequiredColumns"] = missing_required_columns(
                effective_source_key,
                column_names,
                report_target=report_target,
            )
            checked = default_checked_columns(effective_source_key, column_names, report_target=report_target)
            for column in metadata["columns"]:
                column["checked"] = column["name"] in checked
                column["defaultChecked"] = column["name"] in checked
    except PasswordRequired as exc:
        dest.unlink(missing_ok=True)
        return jsonify({"ok": False, "needPassword": True, "error": str(exc)}), 400
    except ParseError as exc:
        dest.unlink(missing_ok=True)
        return jsonify({"ok": False, "error": str(exc)}), 400
    except Exception as exc:
        dest.unlink(missing_ok=True)
        return jsonify({"ok": False, "error": f"파일 파싱 실패: {exc}"}), 400
    finally:
        # Parsed data is cached as CSV; the original upload is no longer needed.
        dest.unlink(missing_ok=True)
    metadata["ok"] = True
    metadata["sourceOptions"] = source_options(request.form.get("reportTarget") or "work")
    return jsonify(metadata)


@app.post("/api/upload")
def upload_data():
    payload = request.get_json(force=True)
    cfg = _merged_config_from_payload(payload)
    items = payload.get("items") or []
    if not items:
        return jsonify({"ok": False, "error": "업로드할 파일이 없습니다."}), 400

    try:
        client = GoogleSheetsClient(cfg["service_account_json"])
        targets = payload.get("targets") or ["work", "watch"]
        prepared = _prepare_dataframes(items, cfg, targets)
        mode = payload.get("mode") or "replace"
        uploaded: list[dict[str, Any]] = []

        for target in targets:
            spreadsheet_id = (
                cfg["work_report_spreadsheet_id"]
                if target == "work"
                else cfg["watch_report_spreadsheet_id"]
            )
            _write_online_settings(client, spreadsheet_id, cfg, target)
            for result in _write_prepared_sources(client, spreadsheet_id, prepared, target, mode):
                uploaded.append(result)
    except (ParseError, SheetsError, Exception) as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400

    save_config(cfg)
    return jsonify({"ok": True, "uploaded": uploaded})


@app.post("/api/export-excel")
def export_excel():
    payload = request.get_json(force=True)
    cfg = _merged_config_from_payload(payload)
    items = payload.get("items") or []
    if not items:
        return jsonify({"ok": False, "error": "엑셀로 생성할 파일이 없습니다."}), 400

    target = (payload.get("target") or payload.get("reportTarget") or (payload.get("targets") or ["work"])[0])
    if target not in {"work", "watch"}:
        return jsonify({"ok": False, "error": "생성할 보고서 유형을 선택해주세요."}), 400

    try:
        prepared = _prepare_dataframes(items, cfg, [target])
        result = generate_local_excel(
            prepared=prepared,
            target=target,
            year=int(cfg.get("default_year") or 2026),
            month=int(cfg.get("default_month") or 1),
        )
    except (ParseError, LocalExcelError, Exception) as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400

    save_config(cfg)
    return jsonify(
        {
            "ok": True,
            "fileName": result.file_name,
            "downloadUrl": f"/api/download/{result.file_name}",
            "warnings": result.warnings,
        }
    )


@app.get("/api/download/<path:file_name>")
def download_output(file_name: str):
    path = (OUTPUT_DIR / file_name).resolve()
    output_root = OUTPUT_DIR.resolve()
    if output_root not in path.parents or not path.exists() or not path.is_file():
        abort(404)
    return send_file(path, as_attachment=True, download_name=path.name)


def _write_prepared_sources(
    client: GoogleSheetsClient,
    spreadsheet_id: str,
    prepared: list[dict[str, Any]],
    target: str,
    mode: str,
) -> list[dict[str, Any]]:
    grouped: dict[str, list[dict[str, Any]]] = {}
    for item in prepared:
        spec = SOURCE_SPECS[item["sourceKey"]]
        grouped.setdefault(spec.sheet_name, []).append(item)

    uploaded: list[dict[str, Any]] = []
    for sheet_name, group_items in grouped.items():
        _validate_unique_period_uploads(sheet_name, group_items)
        combined = pd.concat([item["df"] for item in group_items], ignore_index=True, sort=False)
        combined = combined.fillna("")
        combined = _sort_dataframe_by_period(combined)
        source_keys = list(dict.fromkeys(item["sourceKey"] for item in group_items))
        if mode == "append":
            client.append_dataframe(spreadsheet_id, sheet_name, combined)
        else:
            client.write_dataframe(spreadsheet_id, sheet_name, combined, clear_first=True)
        total_rows = len(combined) + 1
        client.resize_and_freeze(
            spreadsheet_id,
            sheet_name,
            rows=total_rows,
            cols=len(combined.columns),
        )
        uploaded.append(
            {
                "target": target,
                "sheet": sheet_name,
                "sources": source_keys,
                "files": len(group_items),
                "rows": len(combined),
                "columns": len(combined.columns),
                "mode": "append" if mode == "append" else "replace-clear-first",
            }
        )
    return uploaded


def _validate_unique_period_uploads(sheet_name: str, group_items: list[dict[str, Any]]) -> None:
    seen: dict[tuple[int, int], dict[str, Any]] = {}
    period_sources = {
        "investor_snapshot",
        "contract_snapshot",
        "contract_snapshot_prev",
        "contract_snapshot_current",
    }
    for item in group_items:
        if item["sourceKey"] not in period_sources:
            continue
        period_key = item.get("periodKey")
        if not period_key:
            continue
        if period_key in seen:
            first = seen[period_key]
            year, month = period_key
            raise ParseError(
                f"{sheet_name} {year}년 {month}월 원천 파일이 두 번 포함되었습니다: "
                f"{first.get('originalName', '')}, {item.get('originalName', '')}. "
                "현재 선택한 보고서의 파일 목록에서 중복 월을 제거한 뒤 다시 업로드해주세요."
            )
        seen[period_key] = item


def _sort_dataframe_by_period(df: pd.DataFrame) -> pd.DataFrame:
    if "기준년" not in df.columns or "기준월" not in df.columns:
        return df

    sorted_df = df.copy()
    sorted_df["_period_year_sort"] = pd.to_numeric(sorted_df["기준년"], errors="coerce").fillna(9999)
    sorted_df["_period_month_sort"] = pd.to_numeric(sorted_df["기준월"], errors="coerce").fillna(99)
    sorted_df["_period_original_order"] = range(len(sorted_df))
    sorted_df = sorted_df.sort_values(
        by=["_period_year_sort", "_period_month_sort", "_period_original_order"],
        kind="mergesort",
    )
    return sorted_df.drop(columns=["_period_year_sort", "_period_month_sort", "_period_original_order"])


@app.post("/api/clear-cache")
def clear_cache():
    if UPLOAD_DIR.exists():
        shutil.rmtree(UPLOAD_DIR)
    ensure_dirs()
    return jsonify({"ok": True})


@app.post("/api/clear-source-data")
def clear_source_data():
    payload = request.get_json(force=True)
    cfg = _merged_config_from_payload(payload)
    target = payload.get("target") or payload.get("reportTarget") or "work"
    if target not in {"work", "watch"}:
        return jsonify({"ok": False, "error": "비울 문서 유형을 선택해주세요."}), 400

    spreadsheet_id = (
        cfg["work_report_spreadsheet_id"]
        if target == "work"
        else cfg["watch_report_spreadsheet_id"]
    )
    try:
        client = GoogleSheetsClient(cfg["service_account_json"])
        raw_tabs: list[dict[str, Any]] = []
        for source_key in ("contract_all", "investor_snapshot", "contract_snapshot_current"):
            spec = SOURCE_SPECS[source_key]
            headers = list(default_columns_for_target(source_key, target))
            client.write_values(spreadsheet_id, spec.sheet_name, [headers], clear_first=True)
            client.resize_and_freeze(spreadsheet_id, spec.sheet_name, rows=1, cols=len(headers))
            raw_tabs.append({"sheet": spec.sheet_name, "columns": len(headers)})

        calc_tabs = ["계산_계약명세", "계산_기준일투자자", "계산_기준일계약명세", "검증"]
        client.ensure_sheets(spreadsheet_id, calc_tabs)
        for sheet_name in calc_tabs:
            client.clear_sheet(spreadsheet_id, sheet_name)
    except (SheetsError, Exception) as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400

    return jsonify({
        "ok": True,
        "target": target,
        "rawTabs": raw_tabs,
        "clearedCalcTabs": calc_tabs,
    })


def _merged_config_from_request() -> dict[str, Any]:
    data = request.get_json(silent=True) or {}
    return _merged_config_from_payload(data)


def _merged_config_from_payload(data: dict[str, Any]) -> dict[str, Any]:
    cfg = load_config()
    settings = data.get("settings") if isinstance(data.get("settings"), dict) else data
    for key in [
        "work_report_spreadsheet_id",
        "watch_report_spreadsheet_id",
        "service_account_json",
        "apps_script_oauth_json",
        "default_year",
        "default_month",
        "default_excel_password",
    ]:
        if key in settings and settings[key] not in {None, ""}:
            cfg[key] = extract_spreadsheet_id(settings[key]) if key.endswith("_spreadsheet_id") else settings[key]
    cfg["default_year"] = int(cfg.get("default_year") or 2026)
    cfg["default_month"] = int(cfg.get("default_month") or 1)
    return cfg


def _prepare_dataframes(items: list[dict[str, Any]], cfg: dict[str, Any], targets: list[str]) -> list[dict[str, Any]]:
    prepared: list[dict[str, Any]] = []

    for item in items:
        token = item.get("token")
        source_key = item.get("sourceKey")
        if source_key not in SOURCE_SPECS:
            raise ParseError("파일 유형을 선택해주세요.")
        df = load_cached_dataframe(token)
        metadata = load_cached_metadata(token)
        prepared.append({
            "sourceKey": source_key,
            "df": df,
            "columns": item.get("columns") or [],
            "metadata": metadata,
            "allowMissingColumns": bool(item.get("allowMissingColumns")),
            "periodYear": item.get("periodYear"),
            "periodMonth": item.get("periodMonth"),
            "originalName": item.get("originalName") or metadata.get("originalName") or token,
        })

    for item in prepared:
        columns = [col for col in item["columns"] if col in item["df"].columns]
        if not columns:
            columns = list(item["df"].columns)
        item["df"] = item["df"][columns].copy()
        _apply_period_columns(item["df"], item["sourceKey"], cfg, item.get("periodYear"), item.get("periodMonth"))
        item["periodKey"] = _single_period_key(item["df"])
        if not item["allowMissingColumns"]:
            _validate_required_columns(item["df"], item["sourceKey"], targets)
    return prepared


def _validate_required_columns(df, source_key: str, targets: list[str]) -> None:
    spec = SOURCE_SPECS[source_key]
    missing = missing_required_columns(
        source_key,
        df.columns,
        include_generated=True,
        report_targets=targets,
    )
    if missing:
        raise ParseError(
            f"{spec.user_name} 업로드에 필요한 컬럼이 없습니다: {', '.join(missing)}. "
            "엑셀 헤더명과 체크된 컬럼을 확인해주세요."
        )


def _apply_period_columns(
    df,
    source_key: str,
    cfg: dict[str, Any],
    period_year: Any = None,
    period_month: Any = None,
) -> None:
    override = False
    has_period_override = period_year not in {None, ""} and period_month not in {None, ""}

    if has_period_override:
        year = int(period_year)
        month = int(period_month)
        override = True
    elif source_key == "contract_snapshot_prev":
        year, month = _period_for_source(source_key, cfg)
        override = True
    elif source_key in {"contract_snapshot_current", "contract_snapshot"}:
        year, month = _period_for_source(source_key, cfg)
        override = source_key == "contract_snapshot_current"
    elif source_key != "investor_snapshot":
        return
    else:
        year, month = _period_for_source(source_key, cfg)

    if override or "기준년" not in df.columns:
        df["기준년"] = year
    else:
        df["기준년"] = df["기준년"].replace("", year).fillna(year)

    if override or "기준월" not in df.columns:
        df["기준월"] = month
    else:
        df["기준월"] = df["기준월"].replace("", month).fillna(month)


def _single_period_key(df) -> tuple[int, int] | None:
    if "기준년" not in df.columns or "기준월" not in df.columns:
        return None

    years: set[int] = set()
    months: set[int] = set()
    for year in df["기준년"]:
        if year in {None, ""}:
            continue
        try:
            years.add(int(float(year)))
        except (TypeError, ValueError):
            return None
    for month in df["기준월"]:
        if month in {None, ""}:
            continue
        try:
            months.add(int(float(month)))
        except (TypeError, ValueError):
            return None

    if len(years) == 1 and len(months) == 1:
        return next(iter(years)), next(iter(months))
    return None


def _add_months(year: int, month: int, delta: int) -> tuple[int, int]:
    zero_based = (year * 12 + (month - 1)) + delta
    return zero_based // 12, zero_based % 12 + 1


def _period_for_source(source_key: str, cfg: dict[str, Any]) -> tuple[int, int]:
    year = int(cfg.get("default_year") or 2026)
    month = int(cfg.get("default_month") or 1)
    if source_key == "contract_snapshot_prev":
        return _add_months(year, month, -1)
    return year, month


def _write_online_settings(client: GoogleSheetsClient, spreadsheet_id: str, cfg: dict[str, Any], target: str) -> None:
    report_name = "업무보고서" if target == "work" else "상시감시보고서"
    year = int(cfg.get("default_year") or 2026)
    month = int(cfg.get("default_month") or 1)
    raw_sheet_names = list(dict.fromkeys(spec.sheet_name for spec in SOURCE_SPECS.values()))
    rows = [
        [f"{report_name} 온라인 작성 문서"],
        [""],
        ["항목", "값"],
        ["문서구분", report_name],
        ["기준년", str(year)],
        ["기준월", str(month)],
        ["계산위치", "온라인 Apps Script"],
        ["로컬파이썬산출", "사용안함"],
        ["원천 데이터 탭", ", ".join(raw_sheet_names)],
        ["README 업데이트", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        [""],
        ["업로드 흐름"],
        ["1", "웹 UI에서 서비스 계정 JSON과 업무/상시 문서 ID를 확인합니다."],
        ["2", "보고서 종류에 맞는 원천 엑셀 파일을 업로드합니다."],
        ["3", "파일별 컬럼 체크박스에서 업로드할 컬럼을 확인합니다."],
        ["4", "보고서 탭의 세부 계산 로직은 Google Apps Script에서 실행됩니다."],
    ]
    client.write_values(spreadsheet_id, "README", rows, value_input_option="RAW")

    report_sheets = WORK_REPORT_SHEETS if target == "work" else [sheet for sheet in WATCH_REPORT_SHEETS if sheet in WATCH_PERIOD_SHEETS]
    for sheet_name in report_sheets:
        client.write_values(spreadsheet_id, sheet_name, [[year, month]], clear_first=False, start_cell="A1")


def _oauth_callback_page(ok: bool, message: str) -> str:
    title = "인증 완료" if ok else "인증 실패"
    safe_title = html.escape(title)
    safe_message = html.escape(message)
    color = "#136c33" if ok else "#bf1d1d"
    bg = "#f0fff4" if ok else "#fff5f5"
    return f"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <title>{safe_title}</title>
  <style>
    body {{ font-family: "Segoe UI", "Malgun Gothic", Arial, sans-serif; margin: 40px; color: #1f2933; }}
    .box {{ max-width: 720px; border: 1px solid #d7dce3; border-radius: 8px; padding: 20px; background: {bg}; }}
    h1 {{ margin: 0 0 12px; font-size: 20px; color: {color}; }}
    p {{ white-space: pre-wrap; line-height: 1.55; }}
  </style>
</head>
<body>
  <div class="box">
    <h1>{safe_title}</h1>
    <p>{safe_message}</p>
  </div>
</body>
</html>"""


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8787, debug=True)
