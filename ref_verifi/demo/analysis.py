"""작은 구매 데이터의 검증과 집계. 모든 금액은 정수 원 단위."""
from __future__ import annotations

import csv
import re
from collections import defaultdict
from pathlib import Path

DATA_PATH = Path(__file__).resolve().parent / "data" / "requests.csv"
REQUIRED = {"request_id", "department", "amount", "status"}


class DataValidationError(ValueError):
    """잘못된 입력을 조용히 보정하지 않고 원인을 알려준다."""


def load_rows(path: Path = DATA_PATH) -> list[dict]:
    with Path(path).open(encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream)
        if not REQUIRED.issubset(reader.fieldnames or []):
            raise DataValidationError("필수 열 누락: request_id, department, amount, status")
        return list(reader)


def validate_rows(rows: list[dict]) -> list[dict]:
    if not isinstance(rows, list):
        raise DataValidationError("rows는 행 목록이어야 합니다.")
    clean, seen = [], set()
    for number, row in enumerate(rows, 1):
        if not isinstance(row, dict) or not REQUIRED.issubset(row):
            raise DataValidationError(f"행 {number}: 필수 항목 누락")
        values = {}
        for key in ("request_id", "department", "status"):
            value = row[key]
            if not isinstance(value, str) or not value.strip():
                raise DataValidationError(f"행 {number}: {key}는 빈 값이 아닌 문자열이어야 합니다.")
            values[key] = value.strip()
        if values["request_id"] in seen:
            raise DataValidationError(f"행 {number}: 중복 request_id {values['request_id']}")
        seen.add(values["request_id"])
        if values["status"] not in {"approved", "cancelled"}:
            raise DataValidationError(f"행 {number}: status는 approved 또는 cancelled만 허용합니다.")
        amount = row["amount"]
        # bool은 Python에서 int의 하위형이므로 별도로 거부한다.
        if isinstance(amount, bool):
            raise DataValidationError(f"행 {number}: amount는 0 이상의 정수 원이어야 합니다.")
        if isinstance(amount, int) and amount >= 0:
            parsed_amount = amount
        elif isinstance(amount, str) and re.fullmatch(r"[0-9]+", amount.strip()):
            parsed_amount = int(amount.strip())
        else:
            raise DataValidationError(f"행 {number}: amount는 0 이상의 정수 원이어야 합니다.")
        clean.append({**values, "amount": parsed_amount})
    return clean


def summarize(rows: list[dict], department: str | None = None) -> dict:
    """정책: 승인만 포함. 부서 필터는 승인된 행에 적용. 없는 부서는 빈 결과."""
    cleaned = validate_rows(rows)
    if department is not None and not isinstance(department, str):
        raise DataValidationError("department 필터는 문자열이어야 합니다.")
    department = department.strip() or None if department is not None else None
    groups = defaultdict(lambda: {"count": 0, "amount": 0})
    for row in cleaned:
        if row["status"] != "approved":
            continue
        if department is not None and row["department"] != department:
            continue
        groups[row["department"]]["count"] += 1
        groups[row["department"]]["amount"] += row["amount"]
    by_department = [{"department": name, **groups[name]} for name in sorted(groups)]
    return {
        "department_filter": department,
        "record_count": sum(item["count"] for item in by_department),
        "total_amount": sum(item["amount"] for item in by_department),
        "by_department": by_department,
    }
