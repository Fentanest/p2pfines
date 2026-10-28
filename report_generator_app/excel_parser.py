from __future__ import annotations

import json
import io
import re
import shutil
import tempfile
import uuid
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd

from .config import UPLOAD_DIR
from .schema import SOURCE_SPECS, default_checked_columns, detect_source_type


MAX_PREVIEW_ROWS = 20


class ParseError(RuntimeError):
    pass


class PasswordRequired(ParseError):
    pass


def parse_uploaded_file(
    file_path: Path,
    original_name: str,
    password: str | None = None,
    report_target: str = "work",
) -> dict[str, Any]:
    raw = _read_raw_table(file_path, password=password)
    header_idx = _detect_header_row(raw)
    if header_idx is None:
        raise ParseError("컬럼명을 찾지 못했습니다. 첫 10행 안에 헤더 행이 있는 파일인지 확인해주세요.")

    headers = _make_unique_headers(raw.iloc[header_idx].tolist())
    df = raw.iloc[header_idx + 1 :].copy()
    df.columns = headers
    df = df.dropna(how="all")
    df = df.loc[:, [col for col in df.columns if col and not str(col).startswith("Unnamed")]]
    df = _clean_dataframe(df)

    source_key = detect_source_type(original_name, df.columns, report_target=report_target)
    if source_key:
        checked = default_checked_columns(source_key, df.columns, report_target=report_target)
    else:
        checked = set()

    token = uuid.uuid4().hex
    csv_path = UPLOAD_DIR / f"{token}.csv"
    meta_path = UPLOAD_DIR / f"{token}.json"
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    metadata = {
        "token": token,
        "originalName": original_name,
        "sourceKey": source_key,
        "rowCount": int(len(df)),
        "columns": [
            {"name": col, "checked": col in checked, "defaultChecked": col in checked}
            for col in df.columns
        ],
        "preview": df.head(MAX_PREVIEW_ROWS).fillna("").astype(str).values.tolist(),
    }
    meta_path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
    return metadata


def load_cached_dataframe(token: str) -> pd.DataFrame:
    csv_path = UPLOAD_DIR / f"{token}.csv"
    if not csv_path.exists():
        raise ParseError("업로드 캐시를 찾지 못했습니다. 파일을 다시 추가해주세요.")
    return pd.read_csv(csv_path, dtype=object, keep_default_na=False)


def load_cached_metadata(token: str) -> dict[str, Any]:
    meta_path = UPLOAD_DIR / f"{token}.json"
    if not meta_path.exists():
        raise ParseError("업로드 메타데이터를 찾지 못했습니다. 파일을 다시 추가해주세요.")
    return json.loads(meta_path.read_text(encoding="utf-8"))


def _read_raw_table(path: Path, password: str | None = None) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix in {".csv", ".txt"}:
        return pd.read_csv(path, header=None, dtype=object, keep_default_na=False)
    if suffix in {".xlsx", ".xlsm"}:
        encrypted = _is_encrypted_office_file(path)
        try:
            return pd.read_excel(path, sheet_name=0, header=None, dtype=object, engine="openpyxl")
        except Exception as exc:
            decrypted = _try_decrypt_office_file(path, password)
            if decrypted:
                return pd.read_excel(decrypted, sheet_name=0, header=None, dtype=object, engine="openpyxl")
            converted = _try_convert_with_excel(path, password=password)
            if converted:
                try:
                    return pd.read_excel(converted, sheet_name=0, header=None, dtype=object, engine="openpyxl")
                finally:
                    shutil.rmtree(converted.parent, ignore_errors=True)
            if encrypted or _looks_password_error(exc):
                raise PasswordRequired(_password_required_message(password)) from exc
            raise
    if suffix == ".xls":
        encrypted = _is_encrypted_office_file(path)
        try:
            return pd.read_excel(path, sheet_name=0, header=None, dtype=object, engine="xlrd")
        except Exception as exc:
            decrypted = _try_decrypt_office_file(path, password)
            if decrypted:
                return pd.read_excel(decrypted, sheet_name=0, header=None, dtype=object, engine="xlrd")
            converted = _try_convert_with_excel(path, password=password)
            if converted:
                try:
                    return pd.read_excel(converted, sheet_name=0, header=None, dtype=object, engine="openpyxl")
                finally:
                    shutil.rmtree(converted.parent, ignore_errors=True)
            if encrypted or _looks_password_error(exc):
                raise PasswordRequired(_password_required_message(password)) from exc
            raise ParseError(
                "구형 .xls 파일을 직접 읽지 못했습니다. Excel이 설치된 PC에서 실행하거나, 파일을 .xlsx로 다시 저장한 뒤 업로드해주세요. "
                f"원인: {exc}"
            ) from exc
    raise ParseError(f"지원하지 않는 파일 형식입니다: {suffix}")


def _looks_password_error(exc: Exception) -> bool:
    message = str(exc).lower()
    return any(word in message for word in ["encrypted", "password", "암호", "비밀번호", "protected"])


def _password_required_message(password: str | None) -> str:
    if password:
        return "입력된 비밀번호로 엑셀 파일을 열지 못했습니다. 파일 비밀번호를 다시 입력해주세요."
    return "암호가 걸린 엑셀 파일입니다. 파일 비밀번호를 입력해주세요."


def _is_encrypted_office_file(path: Path) -> bool:
    try:
        import msoffcrypto
    except Exception:
        return False
    try:
        with path.open("rb") as handle:
            return bool(msoffcrypto.OfficeFile(handle).is_encrypted())
    except Exception:
        return False


def _try_decrypt_office_file(path: Path, password: str | None) -> io.BytesIO | None:
    if not password:
        return None
    try:
        import msoffcrypto
    except Exception:
        return None
    try:
        with path.open("rb") as handle:
            office_file = msoffcrypto.OfficeFile(handle)
            if not office_file.is_encrypted():
                return None
            office_file.load_key(password=password)
            output = io.BytesIO()
            office_file.decrypt(output)
            output.seek(0)
            return output
    except Exception:
        return None


def _try_convert_with_excel(path: Path, password: str | None = None) -> Path | None:
    if not Path("C:/").exists():
        return None
    try:
        import win32com.client  # type: ignore
    except Exception:
        return None

    tmpdir = Path(tempfile.mkdtemp(prefix="report_generator_xls_"))
    output = tmpdir / f"{path.stem}.xlsx"
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        open_kwargs = {
            "Filename": str(path.resolve()),
            "UpdateLinks": 0,
            "ReadOnly": True,
            "IgnoreReadOnlyRecommended": True,
        }
        if password:
            open_kwargs["Password"] = password
        workbook = excel.Workbooks.Open(**open_kwargs)
        workbook.SaveAs(str(output), FileFormat=51)
        if output.exists() and output.stat().st_size > 0:
            return output
        shutil.rmtree(tmpdir, ignore_errors=True)
        return None
    except Exception:
        shutil.rmtree(tmpdir, ignore_errors=True)
        return None
    finally:
        if workbook is not None:
            workbook.Close(False)
        if excel is not None:
            excel.Quit()


def _detect_header_row(raw: pd.DataFrame) -> int | None:
    best_idx: int | None = None
    best_score = -1
    known_cols = {
        col
        for spec in SOURCE_SPECS.values()
        for col in (*spec.default_columns, *spec.required_columns)
    }
    for idx in range(min(10, len(raw))):
        values = [str(v).strip() for v in raw.iloc[idx].tolist() if str(v).strip() and str(v).strip() != "nan"]
        if not values:
            continue
        known_score = sum(3 for value in values if value in known_cols)
        density_score = min(len(values), 20)
        score = known_score + density_score
        if score > best_score:
            best_score = score
            best_idx = idx
    return best_idx


def _make_unique_headers(values: list[Any]) -> list[str]:
    seen: dict[str, int] = {}
    headers: list[str] = []
    for idx, value in enumerate(values, start=1):
        text = "" if value is None else str(value).strip()
        text = re.sub(r"\s+", " ", text)
        if not text or text.lower() == "nan":
            text = f"Unnamed_{idx}"
        count = seen.get(text, 0)
        seen[text] = count + 1
        headers.append(text if count == 0 else f"{text}_{count + 1}")
    return headers


def _clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    cleaned = df.copy()
    for col in cleaned.columns:
        cleaned[col] = cleaned[col].map(_cell_to_value)
    return cleaned


def _cell_to_value(value: Any) -> Any:
    if pd.isna(value):
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value
