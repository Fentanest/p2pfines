from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable
import random
import time

import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
MAX_ROWS_PER_UPDATE = 2500
RETRYABLE_STATUSES = {429, 500, 502, 503, 504}


class SheetsError(RuntimeError):
    pass


@dataclass
class SheetInfo:
    title: str
    sheet_id: int
    row_count: int = 0
    column_count: int = 0


class GoogleSheetsClient:
    def __init__(self, service_account_json: str | Path):
        json_path = Path(service_account_json)
        if not json_path.exists():
            raise SheetsError(f"서비스 계정 JSON 파일을 찾지 못했습니다: {json_path}")
        creds = Credentials.from_service_account_file(str(json_path), scopes=SCOPES)
        self.service = build("sheets", "v4", credentials=creds, cache_discovery=False)
        self._sheet_map_cache: dict[str, dict[str, SheetInfo]] = {}

    def spreadsheet_title(self, spreadsheet_id: str) -> str:
        try:
            meta = self._execute(
                self.service.spreadsheets().get(
                    spreadsheetId=spreadsheet_id,
                    fields="properties/title",
                )
            )
            return meta["properties"]["title"]
        except HttpError as exc:
            raise SheetsError(_format_google_error(exc)) from exc

    def sheet_map(self, spreadsheet_id: str, force_refresh: bool = False) -> dict[str, SheetInfo]:
        if not force_refresh and spreadsheet_id in self._sheet_map_cache:
            return self._sheet_map_cache[spreadsheet_id]
        try:
            meta = self._execute(
                self.service.spreadsheets().get(
                    spreadsheetId=spreadsheet_id,
                    fields="sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)))",
                )
            )
        except HttpError as exc:
            raise SheetsError(_format_google_error(exc)) from exc
        sheets = {
            sheet["properties"]["title"]: SheetInfo(
                title=sheet["properties"]["title"],
                sheet_id=sheet["properties"]["sheetId"],
                row_count=sheet["properties"].get("gridProperties", {}).get("rowCount", 0),
                column_count=sheet["properties"].get("gridProperties", {}).get("columnCount", 0),
            )
            for sheet in meta.get("sheets", [])
        }
        self._sheet_map_cache[spreadsheet_id] = sheets
        return sheets

    def ensure_sheets(self, spreadsheet_id: str, sheet_names: Iterable[str]) -> None:
        existing = self.sheet_map(spreadsheet_id)
        requests: list[dict[str, Any]] = []
        for name in sheet_names:
            if name not in existing:
                requests.append(
                    {
                        "addSheet": {
                            "properties": {
                                "title": name,
                                "gridProperties": {"rowCount": 1000, "columnCount": 40},
                            }
                        }
                    }
                )
        if requests:
            self._batch_update(spreadsheet_id, requests)
            self.sheet_map(spreadsheet_id, force_refresh=True)

    def clear_sheet(self, spreadsheet_id: str, sheet_name: str) -> None:
        try:
            self._execute(
                self.service.spreadsheets().values().clear(
                    spreadsheetId=spreadsheet_id,
                    range=f"'{sheet_name}'",
                    body={},
                )
            )
        except HttpError as exc:
            raise SheetsError(_format_google_error(exc)) from exc

    def delete_sheets(self, spreadsheet_id: str, sheet_names: Iterable[str]) -> list[str]:
        existing = self.sheet_map(spreadsheet_id)
        requests: list[dict[str, Any]] = []
        deleted: list[str] = []
        for name in sheet_names:
            info = existing.get(name)
            if not info:
                continue
            requests.append({"deleteSheet": {"sheetId": info.sheet_id}})
            deleted.append(name)
        if requests:
            self._batch_update(spreadsheet_id, requests)
            self.sheet_map(spreadsheet_id, force_refresh=True)
        return deleted

    def write_values(
        self,
        spreadsheet_id: str,
        sheet_name: str,
        rows: list[list[Any]],
        clear_first: bool = True,
        start_cell: str = "A1",
        value_input_option: str = "USER_ENTERED",
    ) -> None:
        self.ensure_sheets(spreadsheet_id, [sheet_name])
        if clear_first:
            self.clear_sheet(spreadsheet_id, sheet_name)
        clean_rows = [[_to_sheet_value(value) for value in row] for row in rows]
        start_row = _row_from_a1(start_cell)
        start_col = _col_from_a1(start_cell)
        width = max((len(row) for row in clean_rows), default=1)
        self.ensure_grid_size(
            spreadsheet_id,
            sheet_name,
            rows=start_row + len(clean_rows) - 1,
            cols=start_col + width - 1,
        )
        for offset in range(0, len(clean_rows), MAX_ROWS_PER_UPDATE):
            chunk = clean_rows[offset : offset + MAX_ROWS_PER_UPDATE]
            range_name = f"'{sheet_name}'!{_a1(start_row + offset, start_col)}"
            try:
                self._execute(
                    self.service.spreadsheets().values().update(
                        spreadsheetId=spreadsheet_id,
                        range=range_name,
                        valueInputOption=value_input_option,
                        body={"values": chunk},
                    )
                )
            except HttpError as exc:
                raise SheetsError(_format_google_error(exc)) from exc

    def append_dataframe(
        self,
        spreadsheet_id: str,
        sheet_name: str,
        df: pd.DataFrame,
    ) -> None:
        self.ensure_sheets(spreadsheet_id, [sheet_name])
        rows = [df.columns.tolist(), *df.fillna("").values.tolist()]
        try:
            self._execute(
                self.service.spreadsheets().values().append(
                    spreadsheetId=spreadsheet_id,
                    range=f"'{sheet_name}'!A1",
                    valueInputOption="USER_ENTERED",
                    insertDataOption="INSERT_ROWS",
                    body={"values": [[_to_sheet_value(v) for v in row] for row in rows]},
                )
            )
        except HttpError as exc:
            raise SheetsError(_format_google_error(exc)) from exc

    def ensure_grid_size(self, spreadsheet_id: str, sheet_name: str, rows: int, cols: int) -> None:
        info = self.sheet_map(spreadsheet_id).get(sheet_name)
        if not info:
            return
        row_count = max(rows, info.row_count, 1)
        column_count = max(cols, info.column_count, 1)
        if row_count == info.row_count and column_count == info.column_count:
            return
        requests = [
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": info.sheet_id,
                        "gridProperties": {
                            "rowCount": row_count,
                            "columnCount": column_count,
                        },
                    },
                    "fields": "gridProperties(rowCount,columnCount)",
                }
            }
        ]
        self._batch_update(spreadsheet_id, requests)
        info.row_count = row_count
        info.column_count = column_count

    def read_values(self, spreadsheet_id: str, sheet_name: str) -> list[list[Any]]:
        try:
            result = self._execute(
                self.service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id,
                    range=f"'{sheet_name}'",
                    valueRenderOption="UNFORMATTED_VALUE",
                )
            )
            return result.get("values", [])
        except HttpError as exc:
            if getattr(exc.resp, "status", "") == 400:
                return []
            raise SheetsError(_format_google_error(exc)) from exc

    def upsert_dataframe_by_period(
        self,
        spreadsheet_id: str,
        sheet_name: str,
        df: pd.DataFrame,
        year_col: str = "기준년",
        month_col: str = "기준월",
        trim_to_incoming_columns: bool = False,
    ) -> int:
        if year_col not in df.columns or month_col not in df.columns:
            self.write_dataframe(spreadsheet_id, sheet_name, df, clear_first=True)
            return len(df) + 1

        incoming = df.fillna("")
        incoming_periods = {
            (_period_value(row[year_col]), _period_value(row[month_col]))
            for _, row in incoming.iterrows()
            if _period_value(row[year_col]) and _period_value(row[month_col])
        }
        if not incoming_periods:
            self.write_dataframe(spreadsheet_id, sheet_name, df, clear_first=True)
            return len(df) + 1

        existing_values = self.read_values(spreadsheet_id, sheet_name)
        if not existing_values:
            self.write_dataframe(spreadsheet_id, sheet_name, df, clear_first=True)
            return len(df) + 1

        existing_headers = [str(value).strip() for value in existing_values[0]]
        if trim_to_incoming_columns:
            headers = [str(col) for col in incoming.columns]
        else:
            headers = list(existing_headers)
            for col in incoming.columns:
                if col not in headers:
                    headers.append(str(col))

        existing_records: list[dict[str, Any]] = []
        for row in existing_values[1:]:
            record = {header: (row[idx] if idx < len(row) else "") for idx, header in enumerate(existing_headers)}
            period = (_period_value(record.get(year_col)), _period_value(record.get(month_col)))
            if period[0] and period[1] and period not in incoming_periods:
                existing_records.append(record)

        incoming_records = incoming.to_dict("records")
        rows = [headers]
        for record in [*existing_records, *incoming_records]:
            rows.append([record.get(header, "") for header in headers])
        self.write_values(spreadsheet_id, sheet_name, rows, clear_first=True)
        return len(rows)

    def write_dataframe(
        self,
        spreadsheet_id: str,
        sheet_name: str,
        df: pd.DataFrame,
        clear_first: bool = True,
    ) -> None:
        rows = [df.columns.tolist(), *df.fillna("").values.tolist()]
        self.write_values(spreadsheet_id, sheet_name, rows, clear_first=clear_first)

    def resize_and_freeze(self, spreadsheet_id: str, sheet_name: str, rows: int, cols: int) -> None:
        info = self.sheet_map(spreadsheet_id).get(sheet_name)
        if not info:
            return
        requests = [
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": info.sheet_id,
                        "gridProperties": {
                            "rowCount": max(rows, 100),
                            "columnCount": max(cols, 1),
                            "frozenRowCount": 1,
                        },
                    },
                    "fields": "gridProperties(rowCount,columnCount,frozenRowCount)",
                }
            }
        ]
        self._batch_update(spreadsheet_id, requests)
        info.row_count = max(rows, 100)
        info.column_count = max(cols, 1)

    def _batch_update(self, spreadsheet_id: str, requests: list[dict[str, Any]]) -> None:
        try:
            self._execute(
                self.service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body={"requests": requests},
                )
            )
        except HttpError as exc:
            raise SheetsError(_format_google_error(exc)) from exc

    def _execute(self, request: Any) -> Any:
        delay = 2.0
        for attempt in range(6):
            try:
                return request.execute()
            except HttpError as exc:
                status = int(getattr(exc.resp, "status", 0) or 0)
                if status not in RETRYABLE_STATUSES or attempt == 5:
                    raise
                retry_after = _retry_after_seconds(exc)
                wait = retry_after if retry_after is not None else delay + random.uniform(0, 0.75)
                time.sleep(wait)
                delay = min(delay * 2, 32.0)
        raise RuntimeError("Google Sheets API 재시도 상태가 비정상적으로 종료되었습니다.")


def _to_sheet_value(value: Any) -> Any:
    if pd.isna(value):
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def _period_value(value: Any) -> str:
    if value is None or value == "":
        return ""
    try:
        return str(int(float(value)))
    except (TypeError, ValueError):
        return str(value).strip()


def _retry_after_seconds(exc: HttpError) -> float | None:
    try:
        value = exc.resp.get("retry-after")
    except Exception:
        value = None
    if not value:
        return None
    try:
        return max(float(value), 0.0)
    except ValueError:
        return None


def _format_google_error(exc: HttpError) -> str:
    status = getattr(exc.resp, "status", "")
    try:
        detail = exc.error_details[0]["message"]
    except Exception:
        detail = str(exc)
    if status in {403, 404}:
        return (
            f"구글시트 접근 실패({status}). 서비스 계정 이메일에 해당 문서 편집 권한이 공유되어 있는지 확인해주세요. "
            f"상세: {detail}"
        )
    return f"구글시트 API 오류({status}): {detail}"


def _row_from_a1(cell: str) -> int:
    match = __import__("re").match(r"[A-Z]+([0-9]+)", cell.upper())
    return int(match.group(1)) if match else 1


def _col_from_a1(cell: str) -> int:
    match = __import__("re").match(r"([A-Z]+)[0-9]+", cell.upper())
    letters = match.group(1) if match else "A"
    col = 0
    for char in letters:
        col = col * 26 + ord(char) - ord("A") + 1
    return col


def _a1(row: int, col: int) -> str:
    letters = ""
    while col:
        col, rem = divmod(col - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return f"{letters}{row}"
