from __future__ import annotations

from datetime import date, datetime, timedelta
from typing import Any

import pandas as pd

from .local_work_calculator import (
    _business_no,
    _apply_adjusted_status,
    _clean_product_name,
    _clear_range,
    _group_by,
    _mask_name,
    _number,
    _records,
    _sum,
    _text,
    _value_or,
    _write_matrix,
)


RAW_CONTRACT_ALL = "계약명세"
RAW_INVESTOR = "기준일투자자데이터"
RAW_SNAPSHOT = "기준일명세서"

WATCH_PRODUCTS = ["부동산 PF", "부동산담보대출", "동산 담보대출", "매출채권 담보대출", "신용대출", "개인사업자 신용", "법인신용"]
WATCH_AGES = ["19세 미만", "19세 이상 30세 미만", "30세 이상 40세 미만", "40세 이상 50세 미만", "50세 이상 60세 미만", "60세 이상 70세 미만", "70세 이상 80세 미만", "80세 이상 90세 미만", "90세 이상"]


def apply_watch_calculations(wb, grouped: dict[str, pd.DataFrame], year: int, month: int) -> None:
    contract_all = _normalize_contract_all(_records(grouped.get(RAW_CONTRACT_ALL)), year, month)
    contract_by_no = {_text(row.get("계약No")): row for row in contract_all if _text(row.get("계약No"))}
    investor = _normalize_investor(_records(grouped.get(RAW_INVESTOR)), contract_by_no, year, month)
    snapshots = _normalize_snapshots(_records(grouped.get(RAW_SNAPSHOT)), contract_by_no, year, month)

    _write_period(wb, year, month)
    _clear_watch_outputs(wb)
    _recalc_watch11(wb, contract_all, investor, snapshots, year, month)
    _recalc_watch13(wb, contract_all, snapshots, year, month)
    _recalc_watch22(wb, investor, year, month)
    _recalc_watch23(wb, contract_all, snapshots, year, month)
    _recalc_watch31(wb, investor, year, month)
    _recalc_watch32(wb, snapshots, year, month)


def _normalize_contract_all(rows: list[dict[str, Any]], year: int, month: int) -> list[dict[str, Any]]:
    base_serial = _end_of_month_serial(year, month)
    recent_complete_start = _add_months_serial(base_serial, -6) + 1
    excluded = {"385", "460", "471", "478", "562", "814", "817"}
    out = []
    for row in rows:
        item = dict(row)
        _apply_adjusted_status(item)
        contract_date = _date_serial(item.get("계약일"))
        complete_date = _date_serial(item.get("완납일"))
        no = _text(item.get("계약No"))
        inferred_active = bool(
            contract_date
            and contract_date <= base_serial
            and no not in excluded
            and (
                not complete_date
                or complete_date > base_serial
                or (recent_complete_start <= complete_date <= base_serial)
            )
        )
        explicit = _explicit_boolean(item.get("(상시)1-1"))
        item["(상시)1-1대상"] = "TRUE" if (inferred_active if explicit is None else explicit) else "FALSE"
        item["상품명_정리"] = _clean_product_name(item.get("상품명"))
        out.append(item)
    return out


def _normalize_investor(
    rows: list[dict[str, Any]],
    contract_by_no: dict[str, dict[str, Any]],
    year: int,
    month: int,
) -> list[dict[str, Any]]:
    out = []
    for row in rows:
        item = dict(row)
        item["기준년"] = _year_number(item.get("기준년"), year)
        item["기준월"] = _month_number(item.get("기준월"), month)
        contract = contract_by_no.get(_text(item.get("계약번호")), {})
        item["대출자구분"] = _value_or(item.get("대출자구분"), contract.get("대출자구분"))
        item["(상시)3-1"] = "-".join(_text(item.get(key)) for key in ("기준년", "기준월", "주민등록번호"))
        item["투자자구분_정리"] = _normalize_investor_type(item.get("투자자구분"))
        item["상품분류_정리"] = _normalize_product(item.get("세부상품") or contract.get("세부상품"))
        item["투자잔액_숫자"] = _number(item.get("투자잔액"))
        out.append(item)
    return out


def _normalize_snapshots(
    rows: list[dict[str, Any]],
    contract_by_no: dict[str, dict[str, Any]],
    year: int,
    month: int,
) -> list[dict[str, Any]]:
    out = []
    for row in rows:
        item = dict(row)
        contract = contract_by_no.get(_text(item.get("계약No")), {})
        _apply_adjusted_status(item, contract)
        item["기준년"] = _year_number(item.get("기준년"), year)
        item["기준월"] = _month_number(item.get("기준월"), month)
        item["잔액_숫자"] = _number(item.get("잔액"))
        item["총대출액_숫자"] = _number(item.get("총대출액"))
        item["상품분류_정리"] = _normalize_product(item.get("세부상품"))
        item["월_계약No"] = "-".join(_text(item.get(key)) for key in ("기준년", "기준월", "계약No"))
        item["월_계약No_상태"] = "-".join(_text(item.get(key)) for key in ("기준년", "기준월", "계약No", "상태"))
        item["월_주민등록"] = "-".join(_text(item.get(key)) for key in ("기준년", "기준월", "주민등록"))
        out.append(item)
    return out


def _recalc_watch11(
    wb,
    contract_all: list[dict[str, Any]],
    investor: list[dict[str, Any]],
    snapshots: list[dict[str, Any]],
    year: int,
    month: int,
) -> None:
    if "(상시)1-1" not in wb.sheetnames:
        return
    current_snapshots = [row for row in snapshots if _same_period(row, year, month)]
    snapshot_by_no = {_text(row.get("계약No")): row for row in current_snapshots}
    investor_by_period_product = _group_by(
        investor,
        lambda row: "|".join([
            _text(_year_number(row.get("기준년"), year)),
            _text(_month_number(row.get("기준월"), month)),
            _text(row.get("상품명")),
        ]),
    )
    rows = []
    for idx, row in enumerate([row for row in contract_all if row.get("(상시)1-1대상") == "TRUE"]):
        product = _clean_product_name(row.get("상품명"))
        no = _text(row.get("계약No"))
        start_parts = _serial_year_month(_date_serial(row.get("계약일")), year, month)
        investors = investor_by_period_product.get("|".join([_text(start_parts["year"]), _text(start_parts["month"]), product]), [])
        snapshot = snapshot_by_no.get(no, {})
        limit = _number(row.get("한도") or row.get("총대출"))
        fee_type = _text(row.get("수수료구분"))
        paid_amount = limit if fee_type == "후취" else limit - _number(row.get("취급수수료")) * 1.1
        borrower_type = row.get("대출자구분") or snapshot.get("대출자구분") or ""
        borrower_name = row.get("성명") or snapshot.get("성명") or ""
        borrower_id = row.get("주민등록") or snapshot.get("주민등록") or ""
        investor_counts = [_investor_type_count(investors, kind) for kind in ["개인", "소득적격", "전문", "법인", "금융기관"]]
        expected_count = _number(row.get("투자자수"))
        investor_total = expected_count if expected_count and expected_count == sum(investor_counts) else ""
        rows.append([
            idx + 1,
            product,
            product.replace("줌", "").replace("호", ""),
            _detail_product_label(row.get("세부상품") or row.get("상품") or row.get("상품명")),
            "여" if "리파이낸싱" in _text(row.get("유형") or row.get("최근대출")) else "부",
            _watch11_status_label(idx),
            no,
            limit,
            paid_amount or "",
            limit,
            "" if "잔액_숫자" not in snapshot else _number(snapshot.get("잔액_숫자")),
            "",
            row.get("이율") or "",
            _format_date(row.get("계약일")),
            _format_date(row.get("만기일")),
            _format_date(row.get("완납일")),
            borrower_type,
            _mask_name(borrower_name) if borrower_type == "개인" else borrower_name,
            _identifier(borrower_id, borrower_type),
            *investor_counts,
            investor_total,
        ])
    total = [""] * 25
    total[0] = "합계"
    total[4] = len([row for row in rows if row[4] == "여"])
    total[5] = len([row for row in rows if row[5] in {"정상상환중", "연체중"}])
    for idx in [8, 9, 10, 11, 19, 20, 21, 22, 23, 24]:
        total[idx] = sum(_number(row[idx]) for row in rows)
    total[12] = 0
    total[13] = '하이픈("-") 포함\nex. 2024-01-01'
    total[14] = '하이픈("-") 포함\nex. 2024-03-31'
    total[15] = '하이픈("-") 포함\nex. 2024-03-31'

    ws = wb["(상시)1-1"]
    _write_matrix(ws, 5, 20, [["일반", "소득적격", "전문", "일반", "여신금융기관", ""]])
    _write_matrix(ws, 6, 1, [total])
    if rows:
        _write_matrix(ws, 7, 1, rows)


def _recalc_watch13(wb, contract_all: list[dict[str, Any]], snapshots: list[dict[str, Any]], year: int, month: int) -> None:
    if "(상시)1-3" not in wb.sheetnames:
        return
    row: list[Any] = [year]
    for m in range(1, 13):
        if m > month:
            row.extend(["", "", "", ""])
            continue
        month_end = _end_of_month_serial(year, m)
        cumulative = sum(
            _number(item.get("총대출") or item.get("한도"))
            for item in contract_all
            if _date_serial(item.get("계약일")) and _date_serial(item.get("계약일")) <= month_end
        )
        rows_for_month = [item for item in snapshots if _number(item.get("기준년")) == year and _number(item.get("기준월")) == m]
        if not rows_for_month:
            row.extend([cumulative or "", "", "", ""])
            continue
        balance = _sum(rows_for_month, "잔액_숫자")
        overdue = _sum([item for item in rows_for_month if "연체" in _text(item.get("상태"))], "잔액_숫자")
        row.extend([cumulative, balance, overdue, overdue / balance * 100 if balance else 0])
    _write_matrix(wb["(상시)1-3"], 6, 1, [row])


def _recalc_watch22(wb, investor: list[dict[str, Any]], year: int, month: int) -> None:
    if "(상시)2-2" not in wb.sheetnames:
        return
    source = _build_watch31_rows(investor, year, month)["objects"]
    rows = []
    group_specs = [
        ("개인일반", lambda row: row.get("투자자구분 표시") == "개인일반"),
        ("소득적격", lambda row: row.get("투자자구분 표시") == "소득적격"),
        ("개인 합계", lambda row: row.get("투자자구분 표시") in {"개인일반", "소득적격"}),
    ]
    for label, predicate in group_specs:
        for age_idx, age in enumerate(WATCH_AGES):
            filtered = [row for row in source if predicate(row) and row.get("연령") == age]
            values: list[Any] = []
            for product in WATCH_PRODUCTS:
                values.extend([_sum(filtered, product), _count_positive(filtered, product)])
            total_amount = sum(_number(value) for idx, value in enumerate(values) if idx % 2 == 0)
            total_count = sum(_number(value) for idx, value in enumerate(values) if idx % 2 == 1)
            rows.append([label if age_idx == 0 else "", age, *values, total_amount, total_count, _sum(filtered, "연체잔액"), _count_positive(filtered, "연체잔액")])
        block = rows[-len(WATCH_AGES):]
        subtotal = ["", "합계" if label == "개인 합계" else "소계"] + [0] * 18
        for idx in range(2, 20):
            subtotal[idx] = sum(_number(row[idx]) for row in block)
        rows.append(subtotal)
    _write_matrix(wb["(상시)2-2"], 8, 1, rows)


def _recalc_watch23(
    wb,
    contract_all: list[dict[str, Any]],
    snapshots: list[dict[str, Any]],
    year: int,
    month: int,
) -> None:
    if "(상시)2-3" not in wb.sheetnames:
        return
    contract_by_no = {_text(row.get("계약No")): row for row in contract_all}
    matrix = [[""] * 72 for _ in range(99)]
    for m in range(1, 13):
        offset = (m - 1) * 6
        matrix[0][offset] = year
        matrix[0][offset + 1] = m
        for idx, header in enumerate(["계약번호", "대출잔액", "연체잔액", "연체율", "상품수", "투자자수"]):
            matrix[1][offset + idx] = header
        matrix[2][offset] = "합계"
        if m > month:
            continue
        rows = [row for row in snapshots if _same_period(row, year, m) and "리파이낸싱" in _text(row.get("리파이낸싱 여부"))]
        for idx, row in enumerate(rows[:96]):
            line = 3 + idx
            balance = _number(row.get("잔액_숫자"))
            overdue = balance if _is_overdue_status(row.get("상태")) else ""
            contract = contract_by_no.get(_text(row.get("계약No")), {})
            matrix[line][offset] = row.get("계약No")
            matrix[line][offset + 1] = balance
            matrix[line][offset + 2] = overdue
            matrix[line][offset + 3] = _number(overdue) / balance * 100 if overdue and balance else ""
            matrix[line][offset + 4] = ""
            matrix[line][offset + 5] = contract.get("투자자수") or ""
        matrix[2][offset + 1] = sum(_number(row.get("잔액_숫자")) for row in rows)
        matrix[2][offset + 2] = sum(_number(row.get("잔액_숫자")) for row in rows if _is_overdue_status(row.get("상태")))
        matrix[2][offset + 3] = matrix[2][offset + 2] / matrix[2][offset + 1] * 100 if matrix[2][offset + 1] else 0
        matrix[2][offset + 4] = len(rows)
        matrix[2][offset + 5] = sum(_number(contract_by_no.get(_text(row.get("계약No")), {}).get("투자자수")) for row in rows)
    _write_matrix(wb["(상시)2-3"], 3, 1, matrix)


def _recalc_watch31(wb, investor: list[dict[str, Any]], year: int, month: int) -> None:
    if "(상시)3-1" not in wb.sheetnames:
        return
    built = _build_watch31_rows(investor, year, month)
    ws = wb["(상시)3-1"]
    _write_matrix(ws, 7, 1, [[
        "주민등록번호", "성명", "나이", "연령", "연체잔액", "투자자구분", "투자자구분", "성명",
        "투자자 식별정보", "관리번호", "부동산 PF", "부동산담보대출", "동산 담보대출",
        "매출채권 담보대출", "신용대출", "개인사업자\n신용", "법인신용", "합계", "상품건수",
    ]])
    summary = [""] * 19
    summary[0] = len(built["objects"])
    for idx, product in enumerate(WATCH_PRODUCTS):
        summary[10 + idx] = _sum(built["objects"], product)
    summary[17] = sum(_sum(built["objects"], product) for product in WATCH_PRODUCTS)
    summary[18] = _sum(built["objects"], "상품건수")
    _write_matrix(ws, 5, 1, [summary])
    if built["rows"]:
        _write_matrix(ws, 8, 1, built["rows"])


def _build_watch31_rows(investor: list[dict[str, Any]], year: int, month: int) -> dict[str, Any]:
    first_by_investor = {}
    for row in investor:
        investor_id = _text(row.get("주민등록번호"))
        if investor_id and investor_id not in first_by_investor:
            first_by_investor[investor_id] = row
    filtered = [row for row in investor if _same_period(row, year, month)]
    same_month = [row for row in investor if _month_number(row.get("기준월"), month) == month]
    same_month_product_sums: dict[str, float] = {}
    for row in same_month:
        investor_id = _text(row.get("주민등록번호"))
        product = row.get("상품분류_정리")
        if not investor_id or not product:
            continue
        key = investor_id + "|" + _text(product)
        same_month_product_sums[key] = same_month_product_sums.get(key, 0) + _number(row.get("투자잔액_숫자"))

    by_investor: dict[str, dict[str, Any]] = {}
    for row in filtered:
        investor_id = _text(row.get("주민등록번호"))
        if not investor_id:
            continue
        first = first_by_investor.get(investor_id, row)
        if investor_id not in by_investor:
            by_investor[investor_id] = {
                "id": investor_id,
                "name": first.get("이름") or "",
                "age": first.get("나이") or "",
                "type": row.get("투자자구분_정리"),
                "memberNo": row.get("회원번호") or "",
                "overdue": 0,
                "products": {},
                "count": 0,
            }
        item = by_investor[investor_id]
        item["count"] += 1
        amount = _number(row.get("투자잔액_숫자"))
        product = row.get("상품분류_정리")
        item["products"][product] = item["products"].get(product, 0) + amount
        if _is_overdue_status(row.get("연체여부")):
            item["overdue"] += amount

    for item in by_investor.values():
        for product in list(item["products"].keys()):
            item["products"][product] = same_month_product_sums.get(item["id"] + "|" + _text(product), item["products"][product])

    objects = []
    for item in by_investor.values():
        product_values = [item["products"].get(product, "") for product in WATCH_PRODUCTS]
        out = {
            "주민등록번호": item["id"],
            "성명": item["name"],
            "나이": item["age"],
            "연령": _age_bucket(item["age"]),
            "연체잔액": item["overdue"] or "",
            "투자자구분": item["type"],
            "투자자구분 표시": _investor_type_label(item["type"]),
            "표시명": item["name"] if item["type"] == "법인" else _mask_name(item["name"]),
            "투자자 식별정보": _identifier(item["id"], item["type"]),
            "관리번호": item["memberNo"],
            "합계": sum(_number(value) for value in product_values),
            "상품건수": item["count"],
        }
        for idx, product in enumerate(WATCH_PRODUCTS):
            out[product] = product_values[idx]
        objects.append(out)
    headers = ["주민등록번호", "성명", "나이", "연령", "연체잔액", "투자자구분", "투자자구분 표시", "표시명", "투자자 식별정보", "관리번호", *WATCH_PRODUCTS, "합계", "상품건수"]
    return {"headers": headers, "objects": objects, "rows": [[row.get(header, "") for header in headers] for row in objects]}


def _recalc_watch32(wb, snapshots: list[dict[str, Any]], year: int, month: int) -> None:
    if "(상시)3-2" not in wb.sheetnames:
        return
    filtered = [row for row in snapshots if _same_period(row, year, month) and _number(row.get("잔액_숫자")) != 0]
    by_borrower: dict[str, dict[str, Any]] = {}
    for row in filtered:
        borrower_id = _text(row.get("주민등록"))
        if not borrower_id:
            continue
        if borrower_id not in by_borrower:
            by_borrower[borrower_id] = {"id": borrower_id, "type": row.get("대출자구분") or "", "name": row.get("성명") or "", "products": {}, "count": 0}
        item = by_borrower[borrower_id]
        item["count"] += 1
        product = row.get("상품분류_정리")
        item["products"][product] = item["products"].get(product, 0) + _number(row.get("잔액_숫자"))
    rows = []
    for idx, item in enumerate(by_borrower.values()):
        product_values = [item["products"].get(product, "") for product in WATCH_PRODUCTS]
        rows.append([idx + 1, item["type"], _mask_name(item["name"]) if item["type"] == "개인" else item["name"], _identifier(item["id"], item["type"]), "", *product_values, sum(_number(value) for value in product_values), item["count"]])
    summary = [""] * 14
    summary[0] = "합계"
    for idx, product in enumerate(WATCH_PRODUCTS):
        summary[5 + idx] = sum(_number(row[5 + idx]) for row in rows)
    summary[12] = sum(_number(row[12]) for row in rows)
    summary[13] = sum(_number(row[13]) for row in rows)
    ws = wb["(상시)3-2"]
    _write_matrix(ws, 9, 1, [[
        "구분", "차입자 구분", "성명", "차입자 식별정보", "고유관리번호", "부동산 PF",
        "부동산담보", "기타\n(어음·매출채권 제외)", "어음·매출채권 담보",
        "개인 신용\n(개인사업자 신용 제외)", "개인사업자\n신용", "법인신용", "합계", "상품건수",
    ]])
    _write_matrix(ws, 10, 1, [summary])
    if rows:
        _write_matrix(ws, 11, 1, rows)


def _write_period(wb, year: int, month: int) -> None:
    for sheet_name in ["(상시)1-1", "(상시)2-2", "(상시)3-1", "(상시)3-2"]:
        if sheet_name in wb.sheetnames:
            wb[sheet_name]["A1"] = year
            wb[sheet_name]["B1"] = month


def _clear_watch_outputs(wb) -> None:
    specs = [
        ("(상시)1-1", 6, 1, "max", 25),
        ("(상시)1-3", 6, 1, 1, 49),
        ("(상시)1-3", 15, 1, "max", 4),
        ("(상시)2-2", 8, 1, 30, 20),
        ("(상시)2-2", 40, 1, "max", 20),
        ("(상시)2-3", 3, 1, 99, 72),
        ("(상시)2-3", 110, 1, "max", 8),
        ("(상시)3-1", 5, 1, 1, 19),
        ("(상시)3-1", 8, 1, "max", 19),
        ("(상시)3-2", 10, 1, "max", 16),
    ]
    for sheet_name, row, col, rows, cols in specs:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        row_count = max(ws.max_row - row + 1, 1) if rows == "max" else int(rows)
        _clear_range(ws, row, col, row_count, cols)


def _same_period(row: dict[str, Any], year: int, month: int) -> bool:
    return _year_number(row.get("기준년"), 0) == int(year) and _month_number(row.get("기준월"), 0) == int(month)


def _normalize_investor_type(value: Any) -> str:
    raw = _text(value)
    if "소득" in raw:
        return "소득적격"
    if "전문" in raw:
        return "전문"
    if "법인" in raw:
        return "법인"
    if "금융" in raw or "여신" in raw:
        return "금융기관"
    if "온투" in raw:
        return "온투업자"
    return raw or "개인"


def _normalize_product(value: Any) -> str:
    raw = _text(value)
    if "PF" in raw:
        return "부동산 PF"
    if "부동산담보" in raw or "부동산 담보" in raw:
        return "부동산담보대출"
    if "매출채권" in raw or "어음" in raw:
        return "매출채권 담보대출"
    if "개인사업자" in raw and "신용" in raw:
        return "개인사업자 신용"
    if "법인" in raw and "신용" in raw:
        return "법인신용"
    if "신용" in raw:
        return "신용대출"
    if "동산" in raw or "담보" in raw:
        return "동산 담보대출"
    return raw or "동산 담보대출"


def _investor_type_label(kind: Any) -> str:
    return {
        "개인": "개인일반",
        "법인": "법인일반",
        "소득적격": "소득적격",
        "전문": "전문투자자",
        "금융기관": "여신금융기관",
        "온투업자": "온라인투자연계금융업자",
    }.get(_text(kind), _text(kind))


def _investor_type_count(rows: list[dict[str, Any]], kind: str) -> int:
    return len([row for row in rows if row.get("투자자구분_정리") == kind])


def _count_positive(rows: list[dict[str, Any]], key: str) -> int:
    return len([row for row in rows if _number(row.get(key)) > 0])


def _explicit_boolean(value: Any) -> bool | None:
    if value is True or value is False:
        return bool(value)
    raw = _text(value).lower()
    if raw in {"true", "y", "yes", "1", "대상", "여"}:
        return True
    if raw in {"false", "n", "no", "0", "제외", "부"}:
        return False
    return None


def _is_overdue_status(value: Any) -> bool:
    return _text(value).replace("(", "").replace(")", "").strip() == "연체"


def _year_number(value: Any, fallback: int) -> int:
    if value is None or value == "":
        return int(fallback or 0)
    parsed_date = _to_date(value)
    if parsed_date and parsed_date.year > 1909:
        return parsed_date.year
    parsed = _number(value)
    return int(round(parsed)) if parsed else int(fallback or 0)


def _month_number(value: Any, fallback: int) -> int:
    if value is None or value == "":
        return int(fallback or 0)
    parsed_date = _to_date(value)
    if parsed_date and parsed_date.year > 1909:
        return parsed_date.month
    parsed = _number(value)
    return int(round(parsed)) if parsed else int(fallback or 0)


def _date_serial(value: Any) -> int | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return int(value)
    parsed = _to_date(value)
    if not parsed:
        return None
    return (parsed - date(1899, 12, 30)).days


def _to_date(value: Any) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        return date(1899, 12, 30) + timedelta(days=int(round(value)))
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.date()


def _date_to_serial(year: int, month: int, day: int) -> int:
    return (date(year, month, day) - date(1899, 12, 30)).days


def _end_of_month_serial(year: int, month: int) -> int:
    if month == 12:
        return _date_to_serial(year, 12, 31)
    return _date_to_serial(year, month + 1, 1) - 1


def _add_months_serial(serial: int, delta: int) -> int:
    base = date(1899, 12, 30) + timedelta(days=int(serial))
    month_index = base.month - 1 + delta
    year = base.year + month_index // 12
    month = month_index % 12 + 1
    last_day = _end_of_month(year, month).day
    return _date_to_serial(year, month, min(base.day, last_day))


def _serial_year_month(serial: int | None, fallback_year: int, fallback_month: int) -> dict[str, int]:
    if not serial:
        return {"year": fallback_year, "month": fallback_month}
    parsed = date(1899, 12, 30) + timedelta(days=int(serial))
    return {"year": parsed.year, "month": parsed.month}


def _format_date(value: Any) -> str:
    serial = _date_serial(value)
    if not serial:
        return ""
    parsed = date(1899, 12, 30) + timedelta(days=int(serial))
    return parsed.strftime("%Y-%m-%d")


def _end_of_month(year: int, month: int) -> date:
    if month == 12:
        return date(year, 12, 31)
    return date(year, month + 1, 1) - timedelta(days=1)


def _detail_product_label(value: Any) -> str:
    raw = _text(value)
    return raw[:-2] if raw.endswith("대출") else raw


def _watch11_status_label(idx: int) -> str:
    return ["정상상환중", "완제", "연체중", "채권매각", "기타"][idx] if idx < 5 else ""


def _identifier(value: Any, kind: Any) -> str:
    raw = _text(value)
    if _text(kind) == "법인":
        return _business_no(raw)
    digits = "".join(ch for ch in raw if ch.isdigit())
    return digits[:6] + "-" + digits[6:7] if len(digits) >= 7 else raw


def _age_bucket(age: Any) -> str:
    if age is None or age == "":
        return "19세 미만"
    n = _number(age)
    if n < 19:
        return "19세 미만"
    if n < 30:
        return "19세 이상 30세 미만"
    if n < 40:
        return "30세 이상 40세 미만"
    if n < 50:
        return "40세 이상 50세 미만"
    if n < 60:
        return "50세 이상 60세 미만"
    if n < 70:
        return "60세 이상 70세 미만"
    if n < 80:
        return "70세 이상 80세 미만"
    if n < 90:
        return "80세 이상 90세 미만"
    return "90세 이상"
