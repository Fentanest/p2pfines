from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Iterable


WORK_REPORT_SPREADSHEET_ID = "1GljHSj29KUoF4A6ExAYZeZMiFYHG0WoKUJquYfuL0ek"
WATCH_REPORT_SPREADSHEET_ID = "12pQib4QKaeIrvxqSKygS4ZUW16lfu_mPiz1ipKD4soU"

WORK_REPORT_SHEETS = [
    "PL302",
    "PL303",
    "PL305",
    "PL307",
    "PL309",
    "PL313",
    "PL315",
    "PL317",
    "PL319",
    "PL320",
]

WATCH_REPORT_SHEETS = [
    "(상시)1-1",
    "(상시)1-3",
    "(상시)2-2",
    "(상시)2-3",
    "(상시)3-1",
    "(상시)3-2",
]


@dataclass(frozen=True)
class SourceSpec:
    key: str
    user_name: str
    sheet_name: str
    description: str
    default_columns: tuple[str, ...]
    required_columns: tuple[str, ...]
    work_default_columns: tuple[str, ...] = ()
    work_required_columns: tuple[str, ...] = ()
    watch_default_columns: tuple[str, ...] = ()
    watch_required_columns: tuple[str, ...] = ()


SOURCE_SPECS: dict[str, SourceSpec] = {
    "contract_all": SourceSpec(
        key="contract_all",
        user_name="금월 계약명세",
        sheet_name="계약명세",
        description="계정계 전체 계약명세",
        required_columns=(
            "계약No",
            "상품명",
            "계약일",
            "만기일",
            "한도",
            "총대출",
            "세부상품",
            "상태",
            "완납일",
            "잔액",
            "이율",
            "상환일",
        ),
        default_columns=(
            "번호",
            "회원No",
            "계약No",
            "투자자수",
            "대출자구분",
            "성명",
            "상품명",
            "주민등록",
            "연령",
            "직업구분",
            "상태",
            "연체일",
            "유형",
            "상품",
            "세부상품",
            "채권순위",
            "선순위원금",
            "상환방법",
            "최근대출",
            "계약일",
            "만기일",
            "이율",
            "연체이율",
            "한도",
            "월상환",
            "정상이자",
            "연체이자",
            "원금상환누계",
            "이자상환누계",
            "최초대출",
            "총대출",
            "상환일",
            "상환일이자",
            "최근거래일",
            "완납일",
            "총이자",
            "잔액",
            "원금LTV",
            "줌펀드평가액",
        ),
        work_required_columns=(
            "계약No",
            "상품명",
            "대출자구분",
            "세부상품",
            "상태",
            "연체일",
            "계약일",
            "만기일",
            "한도",
            "총대출",
        ),
        work_default_columns=(
            "계약No",
            "상품명",
            "대출자구분",
            "세부상품",
            "상태",
            "연체일",
            "계약일",
            "만기일",
            "한도",
            "총대출",
        ),
        watch_required_columns=(
            "계약No",
            "투자자수",
            "대출자구분",
            "성명",
            "상품명",
            "주민등록",
            "유형",
            "상품",
            "세부상품",
            "최근대출",
            "상태",
            "연체일",
            "계약일",
            "만기일",
            "완납일",
            "이율",
            "한도",
            "총대출",
            "수수료구분",
            "취급수수료",
        ),
        watch_default_columns=(
            "계약No",
            "투자자수",
            "대출자구분",
            "성명",
            "상품명",
            "주민등록",
            "유형",
            "상품",
            "세부상품",
            "최근대출",
            "상태",
            "연체일",
            "계약일",
            "만기일",
            "완납일",
            "이율",
            "한도",
            "총대출",
            "수수료구분",
            "취급수수료",
        ),
    ),
    "investor_snapshot": SourceSpec(
        key="investor_snapshot",
        user_name="금월 기준일투자자명세",
        sheet_name="기준일투자자",
        description="기준일 투자자별 투자잔액/연체 데이터",
        required_columns=(
            "기준년",
            "기준월",
            "주민등록번호",
            "상품명",
            "세부상품",
            "계약번호",
            "이름",
            "투자자구분",
            "투자금액",
            "투자잔액",
        ),
        default_columns=(
            "기준년",
            "기준월",
            "대출자구분",
            "주민등록번호",
            "순번",
            "상품명",
            "세부상품",
            "계약번호",
            "이름",
            "나이",
            "투자자구분",
            "회원번호",
            "투자금액",
            "투자잔액",
            "연체여부",
        ),
        work_required_columns=(
            "기준년",
            "기준월",
            "주민등록번호",
            "상품명",
            "세부상품",
            "계약번호",
            "이름",
            "투자자구분",
            "투자금액",
            "투자잔액",
        ),
        work_default_columns=(
            "기준년",
            "기준월",
            "대출자구분",
            "주민등록번호",
            "상품명",
            "세부상품",
            "계약번호",
            "이름",
            "투자자구분",
            "투자금액",
            "투자잔액",
        ),
        watch_required_columns=(
            "기준년",
            "기준월",
            "주민등록번호",
            "상품명",
            "세부상품",
            "계약번호",
            "이름",
            "나이",
            "투자자구분",
            "회원번호",
            "투자잔액",
            "연체여부",
        ),
        watch_default_columns=(
            "기준년",
            "기준월",
            "주민등록번호",
            "상품명",
            "세부상품",
            "계약번호",
            "이름",
            "나이",
            "투자자구분",
            "회원번호",
            "투자잔액",
            "연체여부",
        ),
    ),
    "contract_snapshot": SourceSpec(
        key="contract_snapshot",
        user_name="기준일계약",
        sheet_name="기준일계약명세",
        description="기준일 계약/차입자/잔액 명세",
        required_columns=(
            "기준년",
            "기준월",
            "계약No",
            "성명",
            "대출자구분",
            "상품",
            "세부상품",
            "이율",
            "상태",
            "주민등록",
            "계약일",
            "만기일",
            "총대출액",
            "상환일",
            "잔액",
            "LTV",
            "기간",
            "리파이낸싱 여부",
        ),
        default_columns=(
            "기준년",
            "기준월",
            "번호",
            "회원No",
            "계약No",
            "성명",
            "대출자구분",
            "상품",
            "세부상품",
            "상환방법",
            "이율",
            "주민등록",
            "직업구분",
            "상태",
            "계약일",
            "만기일",
            "한도액",
            "월상환액",
            "최초대출액",
            "총대출액",
            "상환일",
            "상환일이자",
            "최근거래일",
            "원금부족금",
            "이자부족금",
            "총이자",
            "잔액",
            "입금일",
            "LTV",
            "기간",
            "입금액",
            "리파이낸싱 여부",
            "PL309",
            "PL313",
            "PL315",
            "PL317",
        ),
        work_required_columns=(
            "기준년",
            "기준월",
            "계약No",
            "대출자구분",
            "세부상품",
            "이율",
            "상태",
            "주민등록",
            "총대출액",
            "상환일",
            "잔액",
            "LTV",
            "기간",
            "PL309",
            "PL313",
            "PL315",
        ),
        work_default_columns=(
            "기준년",
            "기준월",
            "계약No",
            "대출자구분",
            "세부상품",
            "이율",
            "상태",
            "주민등록",
            "총대출액",
            "상환일",
            "잔액",
            "LTV",
            "기간",
            "PL309",
            "PL313",
            "PL315",
            "PL317",
        ),
        watch_required_columns=(
            "기준년",
            "기준월",
            "계약No",
            "성명",
            "대출자구분",
            "세부상품",
            "상태",
            "주민등록",
            "총대출액",
            "잔액",
            "리파이낸싱 여부",
        ),
        watch_default_columns=(
            "기준년",
            "기준월",
            "계약No",
            "성명",
            "대출자구분",
            "세부상품",
            "상태",
            "주민등록",
            "총대출액",
            "잔액",
            "리파이낸싱 여부",
        ),
    ),
}

_contract_snapshot_spec = SOURCE_SPECS["contract_snapshot"]
SOURCE_SPECS["contract_snapshot_prev"] = SourceSpec(
    key="contract_snapshot_prev",
    user_name="전월 기준일계약명세",
    sheet_name=_contract_snapshot_spec.sheet_name,
    description="전월 기준일 계약/차입자/잔액 명세",
    required_columns=_contract_snapshot_spec.required_columns,
    default_columns=_contract_snapshot_spec.default_columns,
    work_required_columns=_contract_snapshot_spec.work_required_columns,
    work_default_columns=_contract_snapshot_spec.work_default_columns,
    watch_required_columns=_contract_snapshot_spec.watch_required_columns,
    watch_default_columns=_contract_snapshot_spec.watch_default_columns,
)
SOURCE_SPECS["contract_snapshot_current"] = SourceSpec(
    key="contract_snapshot_current",
    user_name="금월 기준일계약명세",
    sheet_name=_contract_snapshot_spec.sheet_name,
    description="금월 기준일 계약/차입자/잔액 명세",
    required_columns=_contract_snapshot_spec.required_columns,
    default_columns=_contract_snapshot_spec.default_columns,
    work_required_columns=_contract_snapshot_spec.work_required_columns,
    work_default_columns=_contract_snapshot_spec.work_default_columns,
    watch_required_columns=_contract_snapshot_spec.watch_required_columns,
    watch_default_columns=_contract_snapshot_spec.watch_default_columns,
)


def source_options(report_target: str = "work") -> list[dict[str, Any]]:
    visible_keys = ("contract_all", "investor_snapshot", "contract_snapshot_prev", "contract_snapshot_current")
    return [
        {
            "key": spec.key,
            "userName": spec.user_name,
            "sheetName": spec.sheet_name,
            "description": spec.description,
            "defaultColumns": list(default_columns_for_target(spec.key, report_target)),
            "requiredColumns": list(required_columns_for_target(spec.key, report_target)),
            "defaultColumnsByTarget": {
                "work": list(default_columns_for_target(spec.key, "work")),
                "watch": list(default_columns_for_target(spec.key, "watch")),
            },
            "requiredColumnsByTarget": {
                "work": list(required_columns_for_target(spec.key, "work")),
                "watch": list(required_columns_for_target(spec.key, "watch")),
            },
            "generatedColumns": sorted(generated_columns(spec.key)),
        }
        for key in visible_keys
        for spec in (SOURCE_SPECS[key],)
    ]


def detect_source_type(filename: str, columns: Iterable[str], report_target: str = "work") -> str | None:
    colset = {str(col).strip() for col in columns}
    if {"주민등록번호", "투자자구분", "투자잔액"}.issubset(colset):
        return "investor_snapshot"
    if {"상품명", "한도", "총대출", "투자자수"}.issubset(colset):
        return "contract_all"
    if {"계약No", "성명", "총대출액", "잔액", "리파이낸싱 여부"}.issubset(colset):
        return "contract_snapshot_current"

    scores: list[tuple[int, str]] = []
    for key in ("contract_all", "investor_snapshot", "contract_snapshot_current"):
        required = [col for col in required_columns_for_target(key, report_target) if col not in generated_columns(key)]
        defaults = [col for col in default_columns_for_target(key, report_target) if col not in generated_columns(key)]
        score = sum(3 for col in required if col in colset) + sum(1 for col in defaults if col in colset)
        scores.append((score, key))
    scores.sort(reverse=True)
    return scores[0][1] if scores and scores[0][0] >= 6 else None


def generated_columns(source_key: str) -> set[str]:
    if source_key in {"investor_snapshot", "contract_snapshot", "contract_snapshot_prev", "contract_snapshot_current"}:
        return {"기준년", "기준월"}
    return set()


def required_columns_for_target(source_key: str, report_target: str = "work") -> tuple[str, ...]:
    spec = SOURCE_SPECS[source_key]
    target = _report_target(report_target)
    if target == "watch" and spec.watch_required_columns:
        return spec.watch_required_columns
    if target == "work" and spec.work_required_columns:
        return spec.work_required_columns
    return spec.required_columns


def default_columns_for_target(source_key: str, report_target: str = "work") -> tuple[str, ...]:
    spec = SOURCE_SPECS[source_key]
    target = _report_target(report_target)
    if target == "watch" and spec.watch_default_columns:
        return spec.watch_default_columns
    if target == "work" and spec.work_default_columns:
        return spec.work_default_columns
    return spec.default_columns


def required_columns_for_targets(source_key: str, report_targets: Iterable[str]) -> tuple[str, ...]:
    ordered: list[str] = []
    for target in report_targets:
        for col in required_columns_for_target(source_key, target):
            if col not in ordered:
                ordered.append(col)
    return tuple(ordered)


def missing_required_columns(
    source_key: str,
    columns: Iterable[str],
    include_generated: bool = False,
    report_target: str = "work",
    report_targets: Iterable[str] | None = None,
) -> list[str]:
    available = {str(col).strip() for col in columns}
    generated = generated_columns(source_key) if not include_generated else set()
    required = (
        required_columns_for_targets(source_key, report_targets)
        if report_targets is not None
        else required_columns_for_target(source_key, report_target)
    )
    return [col for col in required if col not in generated and col not in available]


def default_checked_columns(source_key: str, columns: Iterable[str], report_target: str = "work") -> set[str]:
    available = {str(col).strip() for col in columns}
    checked = {col for col in default_columns_for_target(source_key, report_target) if col in available}
    checked.update(col for col in required_columns_for_target(source_key, report_target) if col in available)
    return checked


def _report_target(value: str | None) -> str:
    return "watch" if value == "watch" else "work"
