from __future__ import annotations

import io
import json
import math
import os
import re
import uuid
from copy import deepcopy
from datetime import UTC, date, datetime, timedelta, timezone
from functools import wraps
from pathlib import Path
from typing import Any

import xlrd
from calculation import calculate_case_result
from flask import Flask, redirect, render_template, request, session, url_for
from openpyxl import load_workbook
from werkzeug.security import check_password_hash, generate_password_hash


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
USERS_PATH = DATA_DIR / "users.json"
CASES_DIR = DATA_DIR / "cases"
LEADS_DIR = BASE_DIR / "leads"
CONSULTATION_LOG_PATH = LEADS_DIR / "consultations.jsonl"
ALLOWED_EXTENSIONS = {".xls", ".xlsx"}
WHITESPACE_RE = re.compile(r"[\s\u00a0\u200b\u200c\u200d\ufeff]+")
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
DOCUMENT_FIELDS = (
    ("balance_file", "재무제표"),
    ("income_file", "손익계산서"),
)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 40 * 1024 * 1024
app.secret_key = os.environ.get("RESTART_SECRET_KEY", "restart-local-dev-secret")
KST = timezone(timedelta(hours=9), "KST")


@app.after_request
def prevent_html_cache(response):
    if response.content_type.startswith("text/html"):
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
    return response

def now_iso() -> str:
    return datetime.now(UTC).isoformat(timespec="seconds")


def format_kst_datetime(value: Any) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        return text
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=UTC)
    return parsed.astimezone(KST).strftime("%Y-%m-%d %H:%M:%S")


def read_json_file(path: Path, default: Any) -> Any:
    if not path.exists():
        return default
    try:
        with path.open("r", encoding="utf-8") as file:
            return json.load(file)
    except (OSError, json.JSONDecodeError):
        return default


def write_json_file(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_suffix(path.suffix + ".tmp")
    with temp_path.open("w", encoding="utf-8") as file:
        json.dump(payload, file, ensure_ascii=False, indent=2)
    temp_path.replace(path)


def load_users() -> dict[str, dict[str, Any]]:
    users = read_json_file(USERS_PATH, {})
    return users if isinstance(users, dict) else {}


def save_users() -> None:
    write_json_file(USERS_PATH, USERS)


def load_cases() -> dict[str, dict[str, Any]]:
    cases: dict[str, dict[str, Any]] = {}
    if not CASES_DIR.exists():
        return cases
    for path in CASES_DIR.glob("*.json"):
        case = read_json_file(path, None)
        if isinstance(case, dict) and case.get("case_id"):
            cases[case["case_id"]] = case
    return cases


def save_case(case: dict[str, Any]) -> None:
    if not case.get("case_id"):
        return
    case.setdefault("created_at", now_iso())
    case["updated_at"] = now_iso()
    write_json_file(CASES_DIR / f"{case['case_id']}.json", case)


USERS: dict[str, dict[str, Any]] = load_users()
CASE_STORE: dict[str, dict[str, Any]] = load_cases()
NON_EDITABLE_ACCOUNT_RE = re.compile(
    r"^(?:[\dIVXLCDMivxlcdm\u2160-\u217F\u2460-\u24FF]|"
    r"[\(\[（［]\s*(?:\d+|[IVXLCDMivxlcdm]+|[\u2160-\u217F]+|[\u2460-\u24FF]|[가-힣])\s*[\)\]）］])"
)
PREPAID_ACCOUNTS = {"선급금", "선급비용"}
ASSET_TOTAL_ACCOUNTS = {"자산총계"}
LIABILITY_SECTION_START_ACCOUNTS = {"부채", "부채및자본"}
DEDUCTIBLE_ASSET_ACCOUNTS = ("감가상각누계", "국고보조", "대손충당금")
INTANGIBLE_ASSET_ACCOUNT = "무형자산"
DEBT_DEFAULT_ROWS = (
    ("secured_debt", "담보채무"),
    ("unsecured_financial_debt", "무담보 금융기관채무"),
    ("other_unsecured_debt", "기타 무담보채무(상거래채무 등)"),
    ("related_party_debt", "특수관계인채무"),
    ("unpaid_wages", "미지급급여, 미지급퇴직금(세후)"),
    ("retirement_benefit", "퇴직급여추계액"),
    ("tax_arrears", "조세체납금액(4대보험체납금액 포함)"),
)
COLLATERAL_DEFAULT_ROWS = (
    ("collateral_except_machinery", "담보제공자산(기계장치 제외)"),
    ("collateral_machinery", "담보제공 기계장치"),
    ("savings", "정기예.적금"),
    ("insurance", "보험해약환급금"),
    ("securities", "유가증권"),
    ("other_non_business_assets", "기타 비업무용 자산"),
)
RENT_DEFAULT_ROWS = (
    ("rent_deposit", "임차보증금"),
    ("monthly_rent", "월세"),
)
VARIABLE_COST_ACCOUNTS = {"운반비"}


def allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def column_name(index: int) -> str:
    index += 1
    letters: list[str] = []
    while index:
        index, remainder = divmod(index - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def format_number(value: float) -> str:
    if not math.isfinite(value):
        return str(value)
    if value.is_integer():
        return f"{int(value):,}"
    return f"{value:,.10f}".rstrip("0").rstrip(".")


def format_date_value(value: datetime | date) -> str:
    if isinstance(value, datetime):
        if value.time().replace(microsecond=0).isoformat() == "00:00:00":
            return value.strftime("%Y-%m-%d")
        return value.strftime("%Y-%m-%d %H:%M:%S")
    return value.strftime("%Y-%m-%d")


def compact_text(value: str) -> str:
    return WHITESPACE_RE.sub("", value)


def parse_number_text(value: str) -> float | None:
    text = value.replace(",", "").strip()
    if not text:
        return None
    is_parenthesized_negative = text.startswith("(") and text.endswith(")")
    text = text.strip("()")
    try:
        number = float(text)
        return -number if is_parenthesized_negative else number
    except ValueError:
        return None


def display_number(value: float | None) -> str:
    if value is None:
        return ""
    return format_number(value)


def display_whole_number(value: float | None) -> str:
    if value is None or not math.isfinite(value):
        return ""
    return f"{int(value):,}"


def display_percent(value: float | None) -> str:
    if value is None or not math.isfinite(value):
        return ""
    return f"{value:.2f}%"


def parse_percent_text(value: str) -> float | None:
    return parse_number_text(value.replace("%", ""))


def clean_contact_text(value: str | None, max_length: int = 200) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()[:max_length]


def consultation_form_from_request() -> dict[str, Any]:
    return {
        "company": clean_contact_text(request.form.get("company"), 120),
        "contact_name": clean_contact_text(request.form.get("contact_name"), 80),
        "phone": clean_contact_text(request.form.get("phone"), 40),
        "email": clean_contact_text(request.form.get("email"), 120),
        "message": clean_contact_text(request.form.get("message"), 1000),
        "privacy_consent": request.form.get("privacy_consent") == "on",
        "financial_consent": request.form.get("financial_consent") == "on",
    }


def validate_consultation_form(form: dict[str, Any]) -> list[str]:
    # Temporary marketing-flow setting: allow users to unlock the detailed
    # result without completing the contact form while the gate copy is tested.
    return []

    errors: list[str] = []
    if not form["company"]:
        errors.append("회사명을 입력해 주세요.")
    if not form["contact_name"]:
        errors.append("담당자명을 입력해 주세요.")
    if not form["phone"]:
        errors.append("연락처를 입력해 주세요.")
    if not form["email"]:
        errors.append("이메일을 입력해 주세요.")
    elif not EMAIL_RE.match(form["email"]):
        errors.append("이메일 형식을 확인해 주세요.")
    if not form["privacy_consent"]:
        errors.append("개인정보 수집 및 이용에 동의해 주세요.")
    if not form["financial_consent"]:
        errors.append("재무자료 검토 및 상담 목적 이용에 동의해 주세요.")
    return errors


def append_consultation_log(case: dict[str, Any], form: dict[str, Any], result: dict[str, Any]) -> None:
    LEADS_DIR.mkdir(parents=True, exist_ok=True)
    summary = result.get("summary", {})
    payload = {
        "submitted_at": now_iso(),
        "case_id": case["case_id"],
        "uploaded_company_name": case.get("company_name", ""),
        "contact": {
            "company": form["company"],
            "contact_name": form["contact_name"],
            "phone": form["phone"],
            "email": form["email"],
            "message": form["message"],
        },
        "consents": {
            "privacy_consent": form["privacy_consent"],
            "financial_consent": form["financial_consent"],
        },
        "summary": {
            "liquidation_value": summary.get("liquidation_value", 0),
            "going_concern_value": summary.get("going_concern_value", 0),
            "value_difference": summary.get("value_difference", 0),
            "total_debt": summary.get("total_debt", 0),
        },
    }
    with CONSULTATION_LOG_PATH.open("a", encoding="utf-8") as file:
        file.write(json.dumps(payload, ensure_ascii=False) + "\n")


def normalize_email(email: str | None) -> str:
    return str(email or "").strip().lower()


def current_user_id() -> str | None:
    user_id = session.get("user_id")
    return str(user_id) if user_id in USERS else None


def current_user() -> dict[str, Any] | None:
    user_id = current_user_id()
    return USERS.get(user_id) if user_id else None


def find_user_by_email(email: str) -> dict[str, Any] | None:
    normalized = normalize_email(email)
    return next((user for user in USERS.values() if user.get("email") == normalized), None)


def remember_session_case(case_id: str) -> None:
    case_ids = list(session.get("case_ids", []))
    if case_id not in case_ids:
        case_ids.append(case_id)
        session["case_ids"] = case_ids


def attach_session_cases_to_user(user_id: str) -> None:
    for case_id in session.get("case_ids", []):
        case = CASE_STORE.get(case_id)
        if not case:
            continue
        if not case.get("user_id"):
            case["user_id"] = user_id
            save_case(case)


def login_user(user: dict[str, Any]) -> None:
    session["user_id"] = user["user_id"]
    attach_session_cases_to_user(user["user_id"])


def case_is_accessible(case: dict[str, Any]) -> bool:
    owner_id = case.get("user_id")
    if not owner_id:
        return case.get("case_id") in session.get("case_ids", [])
    return owner_id == current_user_id()


def get_accessible_case(case_id: str) -> dict[str, Any] | None:
    case = CASE_STORE.get(case_id)
    if not case or not case_is_accessible(case):
        return None
    return case


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not current_user():
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)

    return wrapped


@app.context_processor
def inject_auth_context():
    return {
        "current_user": current_user(),
        "format_kst_datetime": format_kst_datetime,
    }


def is_editable_financial_account(account: str) -> bool:
    compacted = compact_text(account)
    if not compacted or compacted in {"자산", *ASSET_TOTAL_ACCOUNTS}:
        return False
    return NON_EDITABLE_ACCOUNT_RE.match(compacted) is None


def is_deductible_asset_account(account: str) -> bool:
    compacted = compact_text(account)
    return any(keyword in compacted for keyword in DEDUCTIBLE_ASSET_ACCOUNTS)


def is_numbered_section_account(account: str) -> bool:
    return NON_EDITABLE_ACCOUNT_RE.match(compact_text(account)) is not None


def subtract_from_financial_row(row: dict[str, Any], deduction: float) -> None:
    amount_number = row.get("amount_number")
    if amount_number is not None:
        row["amount_number"] = amount_number - deduction
        row["amount"] = display_number(row["amount_number"])

    if row.get("account") in PREPAID_ACCOUNTS:
        row["audit_value"] = display_number(0.0)
        row["liquidation_value"] = display_number(0.0)
        return

    for key in ("audit_value", "liquidation_value"):
        number = parse_number_text(row.get(key, ""))
        if number is not None:
            row[key] = display_number(number - deduction)


def current_amount_from_balance_row(row: dict[str, Any]) -> str:
    current_cells = [
        cell
        for cell in row["cells"]
        if cell.get("source_col") in {1, 2} and compact_text(cell["text"])
    ]
    for cell in current_cells:
        if parse_number_text(cell["text"]) is not None:
            return cell["text"]
    return current_cells[0]["text"] if current_cells else ""


def extract_financial_rows(sheet: dict[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    last_editable_asset_row: dict[str, Any] | None = None
    in_intangible_asset_section = False

    for row in sheet["rows"]:
        subject_cell = next(
            (cell for cell in row["cells"] if cell.get("source_col") == 0),
            None,
        )
        account = compact_text(subject_cell["text"]) if subject_cell else ""
        if not account or account in {"과목", "금액"}:
            continue
        if rows and account in LIABILITY_SECTION_START_ACCOUNTS:
            break
        if is_numbered_section_account(account):
            in_intangible_asset_section = INTANGIBLE_ASSET_ACCOUNT in account

        amount_text = current_amount_from_balance_row(row)
        amount_number = parse_number_text(amount_text)
        if is_deductible_asset_account(account):
            if last_editable_asset_row is not None and amount_number is not None:
                subtract_from_financial_row(last_editable_asset_row, abs(amount_number))
            continue

        editable = is_editable_financial_account(account)

        if editable and account in PREPAID_ACCOUNTS:
            audit_value = 0.0
        elif editable and in_intangible_asset_section:
            audit_value = 0.0
        elif editable:
            audit_value = amount_number
        else:
            audit_value = None

        liquidation_value = audit_value if editable else None

        rows.append(
            {
                "row": row["index"],
                "source_row": row.get("source_index", row["index"]),
                "account": account,
                "amount": display_number(amount_number),
                "amount_number": amount_number,
                "audit_value": display_number(audit_value),
                "liquidation_value": display_number(liquidation_value),
                "is_editable": editable,
            }
        )
        if editable:
            last_editable_asset_row = rows[-1]
        if account in ASSET_TOTAL_ACCOUNTS:
            break

    return rows


def income_roman_stage(account: str) -> int | None:
    compacted = compact_text(account)
    unicode_stages = (
        ("Ⅰ", 1),
        ("Ⅱ", 2),
        ("Ⅲ", 3),
        ("Ⅳ", 4),
        ("Ⅴ", 5),
        ("Ⅵ", 6),
        ("Ⅶ", 7),
        ("Ⅷ", 8),
        ("Ⅸ", 9),
        ("Ⅹ", 10),
    )
    for prefix, stage in unicode_stages:
        if compacted.startswith(prefix):
            return stage

    upper = compacted.upper()
    ascii_stages = (
        ("VIII", 8),
        ("VII", 7),
        ("VI", 6),
        ("IV", 4),
        ("IX", 9),
        ("III", 3),
        ("II", 2),
        ("V", 5),
        ("X", 10),
        ("I", 1),
    )
    for prefix, stage in ascii_stages:
        if upper.startswith(f"{prefix}.") or upper.startswith(prefix):
            return stage

    return None


def sum_amount_group(row: dict[str, Any], source_cols: tuple[int, ...]) -> tuple[str, float | None]:
    total = 0.0
    found_number = False
    fallback_text = ""

    for cell in row["cells"]:
        if not cell_overlaps_cols(cell, source_cols):
            continue

        text = compact_text(cell["text"])
        if not text:
            continue

        number = parse_number_text(text)
        if number is None:
            fallback_text = fallback_text or text
            continue

        total += number
        found_number = True

    if found_number:
        return display_number(total), total
    return fallback_text, None


def remove_income_rows_between_sales_roman(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    positions: dict[int, int] = {}
    for index, row in enumerate(rows):
        stage = income_roman_stage(row["account"])
        if stage in {1, 2, 3} and stage not in positions:
            positions[stage] = index

    remove_indexes: set[int] = set()
    for previous_stage, current_stage in ((1, 2), (2, 3)):
        previous_index = positions.get(previous_stage)
        current_index = positions.get(current_stage)
        if previous_index is None or current_index is None:
            continue
        if previous_index < current_index - 1:
            remove_indexes.update(range(previous_index + 1, current_index))

    return [row for index, row in enumerate(rows) if index not in remove_indexes]


def is_editable_income_row(account: str, section: str) -> bool:
    compacted = compact_text(account)
    if not compacted:
        return False
    if section == "sales":
        return "매출액" in compacted or "매출원가" in compacted
    return NON_EDITABLE_ACCOUNT_RE.match(compacted) is None


def income_percentage(numerator: float | None, denominator: float | None) -> float | None:
    if numerator is None or denominator in (None, 0):
        return None
    return (numerator / denominator) * 100


def extract_income_rows(sheet: dict[str, Any]) -> list[dict[str, Any]]:
    raw_rows: list[dict[str, Any]] = []

    for row in sheet["rows"]:
        subject_cell = next(
            (cell for cell in row["cells"] if cell.get("source_col") == 0),
            None,
        )
        account = compact_text(subject_cell["text"]) if subject_cell else ""
        if not account or account in {"과목", "금액"}:
            continue
        if "영업외" in account:
            break

        first_display, first_number = sum_amount_group(row, (1, 2))
        second_display, second_number = sum_amount_group(row, (3, 4))
        average_number = ((first_number or 0) + (second_number or 0)) / 2

        raw_rows.append(
            {
                "row": len(raw_rows) + 1,
                "source_row": row.get("source_index", row["index"]),
                "account": account,
                "y_minus_1": first_display,
                "y_minus_1_number": first_number,
                "y": second_display,
                "y_number": second_number,
                "average": display_number(average_number),
                "average_number": average_number,
            }
        )

    filtered_rows = remove_income_rows_between_sales_roman(raw_rows)
    sales_average = next(
        (
            row["average_number"]
            for row in filtered_rows
            if "매출액" in row["account"]
        ),
        None,
    )

    income_rows: list[dict[str, Any]] = []
    section = "sales"
    for row in filtered_rows:
        stage = income_roman_stage(row["account"])
        if stage is not None and stage >= 4:
            section = "expense"

        account = row["account"]
        editable = is_editable_income_row(account, section)
        metric_display = ""
        monthly_average_display = ""
        cost_type = ""
        value_kind = "amount"
        final_value = ""

        if section == "sales":
            value_kind = "percent"
            if "매출액" in account:
                metric_display = display_percent(
                    income_percentage(
                        (row["y_number"] or 0) - (row["y_minus_1_number"] or 0),
                        row["y_minus_1_number"],
                    )
                )
            elif "매출원가" in account:
                metric_display = display_percent(
                    income_percentage(row["average_number"], sales_average)
                )
            final_value = metric_display if editable else ""
        elif editable:
            cost_type = (
                "variable"
                if any(variable in account for variable in VARIABLE_COST_ACCOUNTS)
                else "fixed"
            )
            if cost_type == "variable":
                value_kind = "percent"
                monthly_average_display = display_percent(
                    income_percentage(row["average_number"], sales_average)
                )
            else:
                monthly_average_display = display_whole_number(row["average_number"] / 12)
            final_value = monthly_average_display
        else:
            monthly_average_display = display_whole_number(row["average_number"] / 12)

        income_rows.append(
            {
                **row,
                "row": len(income_rows) + 1,
                "section": section,
                "metric_display": metric_display,
                "monthly_average_display": monthly_average_display,
                "sales_average_number": sales_average,
                "cost_type": cost_type,
                "final_value": final_value,
                "value_kind": value_kind,
                "is_editable": editable,
            }
        )

    return income_rows


def company_name_from_workbook(workbook: dict[str, Any]) -> str:
    for file in workbook["files"]:
        name = Path(file["name"]).stem
        name = re.sub(r"^\d{4}년\d{1,2}월", "", name)
        name = re.sub(r"(재무상태표|손익계산서)", "", name)
        name = name.strip("_ -")
        if name:
            return name
    return "회사"


def create_case(workbook: dict[str, Any]) -> dict[str, Any]:
    case_id = uuid.uuid4().hex[:12]
    created_at = now_iso()
    balance_sheet = next(
        (sheet for sheet in workbook["sheets"] if sheet.get("document_label") == "재무제표"),
        workbook["sheets"][0],
    )
    income_sheet = next(
        (sheet for sheet in workbook["sheets"] if sheet.get("document_label") == "손익계산서"),
        workbook["sheets"][-1],
    )
    case = {
        "case_id": case_id,
        "company_name": company_name_from_workbook(workbook),
        "scenario_name": "",
        "source_case_id": None,
        "user_id": current_user_id(),
        "created_at": created_at,
        "updated_at": created_at,
        "workbook": workbook,
        "financial_rows": extract_financial_rows(balance_sheet),
        "financial_saved": False,
        "debt_rows": build_default_debt_rows(),
        "debt_saved": False,
        "income_rows": extract_income_rows(income_sheet),
        "income_saved": False,
        "collateral_rows": build_default_collateral_rows(),
        "rent_rows": build_default_rent_rows(),
        "collateral_saved": False,
    }
    CASE_STORE[case_id] = case
    remember_session_case(case_id)
    save_case(case)
    return case


def build_default_debt_rows() -> list[dict[str, Any]]:
    return [
        {
            "field": field,
            "category": category,
            "debt_amount": "",
            "collateral_type": "",
            "audit_value": "",
        }
        for field, category in DEBT_DEFAULT_ROWS
    ]


def build_default_collateral_rows() -> list[dict[str, Any]]:
    return [
        {
            "field": field,
            "category": category,
            "audit_value": "",
            "liquidation_value": "",
        }
        for field, category in COLLATERAL_DEFAULT_ROWS
    ]


def build_default_rent_rows() -> list[dict[str, Any]]:
    return [
        {
            "field": field,
            "category": category,
            "amount": "",
        }
        for field, category in RENT_DEFAULT_ROWS
    ]


def clean_submitted_number(value: str) -> tuple[str, float | None]:
    number = parse_number_text(value)
    return display_number(number), number


def clean_submitted_income_value(value: str, value_kind: str) -> tuple[str, float | None]:
    if value_kind == "percent":
        number = parse_percent_text(value)
        return display_percent(number), number
    return clean_submitted_number(value)


def cell_end_col(cell: dict[str, Any]) -> int:
    return cell["source_col"] + max(1, cell.get("colspan", 1)) - 1


def cell_overlaps_cols(cell: dict[str, Any], cols: tuple[int, ...]) -> bool:
    if not cols:
        return False
    return cell["source_col"] <= max(cols) and cell_end_col(cell) >= min(cols)


def cell_overlaps_group(cell: dict[str, Any], group: tuple[int, int]) -> bool:
    return cell["source_col"] <= group[1] and cell_end_col(cell) >= group[0]


def normalize_label_cells(row: dict[str, Any]) -> None:
    for cell in row["cells"]:
        compacted = compact_text(cell["text"])
        if compacted in {"과목", "금액"}:
            cell["text"] = compacted


def normalize_subject_cell(row: dict[str, Any], layout: dict[str, Any]) -> None:
    for cell in row["cells"]:
        if cell_overlaps_cols(cell, layout["subject_cols"]):
            cell["text"] = compact_text(cell["text"])


def subject_text(row: dict[str, Any], layout: dict[str, Any]) -> str:
    texts: list[str] = []
    for cell in row["cells"]:
        if cell_overlaps_cols(cell, layout["subject_cols"]):
            text = compact_text(cell["text"])
            if text:
                texts.append(text)
    return "".join(texts)


def has_amount_label(row: dict[str, Any]) -> bool:
    return any(compact_text(cell["text"]) == "금액" for cell in row["cells"])


def is_metadata_subject(subject: str) -> bool:
    compacted = compact_text(subject)
    if not compacted:
        return False
    if compacted in {"재무상태표", "손익계산서"}:
        return True
    if compacted.startswith(("회사명", "단위", "(단위")):
        return True
    return compacted.startswith("제") and any(
        marker in compacted for marker in ("현재", "부터", "까지")
    )


def detect_table_layout(rows: list[dict[str, Any]], col_count: int) -> dict[str, Any]:
    header_row_index = 0
    subject_cell: dict[str, Any] | None = None

    for row_index, row in enumerate(rows):
        for cell in row["cells"]:
            if compact_text(cell["text"]) == "과목":
                header_row_index = row_index
                subject_cell = cell
                break
        if subject_cell is not None:
            break

    if subject_cell is None:
        return {
            "header_row_index": 0,
            "amount_row_index": 1,
            "subject_cols": (0,),
            "amount_groups": ((1, 1), (2, 2)),
        }

    subject_start = subject_cell["source_col"]
    subject_end = cell_end_col(subject_cell)
    amount_row_index = min(header_row_index + 1, len(rows) - 1)
    amount_groups: list[tuple[int, int]] = []

    for row_index in range(header_row_index + 1, min(header_row_index + 4, len(rows))):
        candidates = [
            cell
            for cell in rows[row_index]["cells"]
            if compact_text(cell["text"]) == "금액"
        ]
        if candidates:
            amount_row_index = row_index
            amount_groups = [
                (cell["source_col"], cell_end_col(cell)) for cell in candidates
            ]
            break

    if not amount_groups:
        amount_start = min(subject_end + 1, max(0, col_count - 1))
        amount_groups = ((amount_start, amount_start),)

    return {
        "header_row_index": header_row_index,
        "amount_row_index": amount_row_index,
        "subject_cols": tuple(range(subject_start, subject_end + 1)),
        "amount_groups": tuple(amount_groups),
    }


def clean_sheet_rows(
    rows: list[dict[str, Any]],
    layout: dict[str, Any],
) -> tuple[list[dict[str, Any]], int]:
    cleaned_rows: list[dict[str, Any]] = []
    removed_count = 0
    table_started = False
    keep_first_amount_header = False

    for row in rows:
        subject = subject_text(row, layout)

        if not table_started:
            if subject == "과목":
                table_started = True
                keep_first_amount_header = True
                normalize_subject_cell(row, layout)
                normalize_label_cells(row)
                cleaned_rows.append(row)
            else:
                removed_count += 1
            continue

        if subject == "과목":
            removed_count += 1
            keep_first_amount_header = False
            continue

        if is_metadata_subject(subject):
            removed_count += 1
            keep_first_amount_header = False
            continue

        if keep_first_amount_header and has_amount_label(row):
            normalize_label_cells(row)
            cleaned_rows.append(row)
            keep_first_amount_header = False
            continue

        keep_first_amount_header = False

        if not subject:
            removed_count += 1
            continue

        if subject.startswith("회사명"):
            removed_count += 1
            continue

        normalize_subject_cell(row, layout)
        cleaned_rows.append(row)

    for display_index, row in enumerate(cleaned_rows, start=1):
        row["index"] = display_index

    return cleaned_rows, removed_count


def cell_by_source_col(row: dict[str, Any], source_col: int) -> dict[str, Any] | None:
    for cell in row["cells"]:
        if cell.get("source_col") == source_col:
            return cell
    return None


def is_empty_cell(cell: dict[str, Any] | None) -> bool:
    return cell is None or compact_text(cell["text"]) == ""


def empty_cell(source_col: int) -> dict[str, Any]:
    return {
        "text": "",
        "kind": "empty",
        "rowspan": 1,
        "colspan": 1,
        "source_col": source_col,
    }


def display_cell(cell: dict[str, Any] | None, source_col: int) -> dict[str, Any]:
    if cell is None:
        return empty_cell(source_col)
    return {
        **cell,
        "rowspan": 1,
        "colspan": 1,
        "source_col": source_col,
    }


def make_text_cell(
    text: str,
    source_col: int,
    colspan: int = 1,
    rowspan: int = 1,
) -> dict[str, Any]:
    return {
        "text": text,
        "kind": "text" if text else "empty",
        "rowspan": rowspan,
        "colspan": colspan,
        "source_col": source_col,
    }


def first_cell_text_for_group(row: dict[str, Any], group: tuple[int, int]) -> str:
    for cell in row["cells"]:
        if cell_overlaps_group(cell, group):
            text = compact_text(cell["text"])
            if text:
                return text
    return ""


def subject_display_cell(row: dict[str, Any], layout: dict[str, Any]) -> dict[str, Any]:
    text = subject_text(row, layout)
    source = next(
        (cell for cell in row["cells"] if cell_overlaps_cols(cell, layout["subject_cols"])),
        None,
    )
    return {
        "text": text,
        "kind": "text" if text else "empty",
        "rowspan": source.get("rowspan", 1) if source else 1,
        "colspan": 1,
        "source_col": 0,
    }


def physical_cells_for_group(
    row: dict[str, Any],
    group: tuple[int, int],
    display_start_col: int,
) -> list[dict[str, Any]]:
    cells: list[dict[str, Any]] = []
    for offset, source_col in enumerate(range(group[0], group[1] + 1)):
        cells.append(display_cell(cell_by_source_col(row, source_col), display_start_col + offset))
    return cells


def normalize_income_statement_sheet(sheet: dict[str, Any]) -> None:
    layout = sheet.get("layout")
    if not layout or len(layout["amount_groups"]) < 2 or len(sheet["rows"]) < 2:
        return

    compacted_rows: list[dict[str, Any]] = []
    amount_groups = layout["amount_groups"][:2]
    group_widths = [2, 2]

    for row_index, row in enumerate(sheet["rows"]):
        if row_index == 0:
            compacted_rows.append(
                {
                    **row,
                    "cells": [
                        subject_display_cell(row, layout),
                        make_text_cell(
                            first_cell_text_for_group(row, amount_groups[0]),
                            1,
                            colspan=2,
                        ),
                        make_text_cell(
                            first_cell_text_for_group(row, amount_groups[1]),
                            3,
                            colspan=2,
                        ),
                    ],
                }
            )
            continue

        if row_index == 1:
            compacted_rows.append(
                {
                    **row,
                    "cells": [
                        make_text_cell("금액", 1, colspan=2),
                        make_text_cell("금액", 3, colspan=2),
                    ],
                }
            )
            continue

        cells = [subject_display_cell(row, layout)]
        display_col = 1
        for group, width in zip(amount_groups, group_widths):
            group_cells = physical_cells_for_group(row, group, display_col)
            if len(group_cells) < width:
                group_cells.extend(empty_cell(display_col + i) for i in range(len(group_cells), width))
            cells.extend(group_cells[:width])
            display_col += width

        compacted_rows.append(
            {
                **row,
                "cells": cells,
            }
        )

    sheet["rows"] = compacted_rows
    sheet["col_count"] = 5
    sheet["columns"] = ["A", "B", "C", "D", "E"]


def normalize_balance_sheet(sheet: dict[str, Any]) -> None:
    layout = sheet.get("layout")
    if not layout or len(layout["amount_groups"]) < 2 or len(sheet["rows"]) < 2:
        return

    amount_groups = layout["amount_groups"][:2]
    group_widths = [group[1] - group[0] + 1 for group in amount_groups]
    compacted_rows: list[dict[str, Any]] = []

    for row_index, row in enumerate(sheet["rows"]):
        cells: list[dict[str, Any]] = []

        if row_index == 0:
            cells.append(subject_display_cell(row, layout))
            display_col = 1
            for group, width in zip(amount_groups, group_widths):
                cells.append(
                    make_text_cell(
                        first_cell_text_for_group(row, group),
                        display_col,
                        colspan=width,
                    )
                )
                display_col += width
        elif row_index == 1:
            display_col = 1
            for width in group_widths:
                cells.append(make_text_cell("금액", display_col, colspan=width))
                display_col += width
        else:
            cells.append(subject_display_cell(row, layout))
            display_col = 1
            for group in amount_groups:
                group_cells = physical_cells_for_group(row, group, display_col)
                cells.extend(group_cells)
                display_col += len(group_cells)

        compacted_rows.append({**row, "cells": cells})

    col_count = 1 + sum(group_widths)
    sheet["rows"] = compacted_rows
    sheet["col_count"] = col_count
    sheet["columns"] = [column_name(index) for index in range(col_count)]


def normalize_financial_sheet(sheet: dict[str, Any], document_label: str) -> None:
    if document_label == "손익계산서":
        normalize_income_statement_sheet(sheet)
    else:
        normalize_balance_sheet(sheet)


def build_merge_maps(
    merged_ranges: list[tuple[int, int, int, int]],
) -> tuple[dict[tuple[int, int], tuple[int, int]], set[tuple[int, int]]]:
    top_left: dict[tuple[int, int], tuple[int, int]] = {}
    covered: set[tuple[int, int]] = set()

    for row_start, row_end, col_start, col_end in merged_ranges:
        rowspan = max(1, row_end - row_start)
        colspan = max(1, col_end - col_start)
        top_left[(row_start, col_start)] = (rowspan, colspan)

        for row in range(row_start, row_end):
            for col in range(col_start, col_end):
                if (row, col) != (row_start, col_start):
                    covered.add((row, col))

    return top_left, covered


def format_xls_cell(book: xlrd.book.Book, cell: xlrd.sheet.Cell) -> tuple[str, str]:
    if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
        return "", "empty"
    if cell.ctype == xlrd.XL_CELL_TEXT:
        return str(cell.value).strip(), "text"
    if cell.ctype == xlrd.XL_CELL_NUMBER:
        return format_number(float(cell.value)), "number"
    if cell.ctype == xlrd.XL_CELL_DATE:
        try:
            value = xlrd.xldate.xldate_as_datetime(cell.value, book.datemode)
            return format_date_value(value), "date"
        except (OverflowError, ValueError):
            return str(cell.value), "text"
    if cell.ctype == xlrd.XL_CELL_BOOLEAN:
        return ("TRUE" if cell.value else "FALSE"), "boolean"
    if cell.ctype == xlrd.XL_CELL_ERROR:
        return "#ERROR", "error"
    return str(cell.value), "text"


def parse_xls(
    file_bytes: bytes,
    display_sheet_name: str | None = None,
    only_first_sheet: bool = False,
) -> dict[str, Any]:
    try:
        book = xlrd.open_workbook(file_contents=file_bytes, formatting_info=True)
    except NotImplementedError:
        book = xlrd.open_workbook(file_contents=file_bytes)

    sheets = []
    source_sheets = [book.sheet_by_index(0)] if only_first_sheet else book.sheets()
    for sheet in source_sheets:
        merged_ranges = list(sheet.merged_cells)
        merge_top_left, merge_covered = build_merge_maps(merged_ranges)

        rows = []
        for row_index in range(sheet.nrows):
            rendered_cells = []
            for col_index in range(sheet.ncols):
                if (row_index, col_index) in merge_covered:
                    continue

                text, kind = format_xls_cell(book, sheet.cell(row_index, col_index))
                rowspan, colspan = merge_top_left.get((row_index, col_index), (1, 1))
                rendered_cells.append(
                    {
                        "text": text,
                        "kind": kind,
                        "rowspan": rowspan,
                        "colspan": colspan,
                        "source_col": col_index,
                    }
                )

            rows.append(
                {
                    "index": row_index + 1,
                    "source_index": row_index + 1,
                    "cells": rendered_cells,
                }
            )

        layout = detect_table_layout(rows, sheet.ncols)
        rows, removed_count = clean_sheet_rows(rows, layout)

        sheets.append(
            {
                "name": display_sheet_name or sheet.name,
                "source_sheet_name": sheet.name,
                "row_count": len(rows),
                "original_row_count": sheet.nrows,
                "removed_row_count": removed_count,
                "col_count": sheet.ncols,
                "columns": [column_name(index) for index in range(sheet.ncols)],
                "layout": layout,
                "rows": rows,
            }
        )

    return {"sheets": sheets}


def format_xlsx_value(value: Any) -> tuple[str, str]:
    if value is None:
        return "", "empty"
    if isinstance(value, bool):
        return ("TRUE" if value else "FALSE"), "boolean"
    if isinstance(value, (datetime, date)):
        return format_date_value(value), "date"
    if isinstance(value, int):
        return f"{value:,}", "number"
    if isinstance(value, float):
        return format_number(value), "number"
    return str(value).strip(), "text"


def parse_xlsx(
    file_bytes: bytes,
    display_sheet_name: str | None = None,
    only_first_sheet: bool = False,
) -> dict[str, Any]:
    workbook = load_workbook(io.BytesIO(file_bytes), read_only=False, data_only=True)
    sheets = []

    source_worksheets = [workbook.worksheets[0]] if only_first_sheet else workbook.worksheets
    for worksheet in source_worksheets:
        max_row = worksheet.max_row or 1
        max_col = worksheet.max_column or 1
        merged_ranges = [
            (
                merged.min_row - 1,
                merged.max_row,
                merged.min_col - 1,
                merged.max_col,
            )
            for merged in worksheet.merged_cells.ranges
        ]
        merge_top_left, merge_covered = build_merge_maps(merged_ranges)

        rows = []
        for row_index in range(max_row):
            rendered_cells = []
            for col_index in range(max_col):
                if (row_index, col_index) in merge_covered:
                    continue

                cell = worksheet.cell(row=row_index + 1, column=col_index + 1)
                text, kind = format_xlsx_value(cell.value)
                rowspan, colspan = merge_top_left.get((row_index, col_index), (1, 1))
                rendered_cells.append(
                    {
                        "text": text,
                        "kind": kind,
                        "rowspan": rowspan,
                        "colspan": colspan,
                        "source_col": col_index,
                    }
                )

            rows.append(
                {
                    "index": row_index + 1,
                    "source_index": row_index + 1,
                    "cells": rendered_cells,
                }
            )

        layout = detect_table_layout(rows, max_col)
        rows, removed_count = clean_sheet_rows(rows, layout)

        sheets.append(
            {
                "name": display_sheet_name or worksheet.title,
                "source_sheet_name": worksheet.title,
                "row_count": len(rows),
                "original_row_count": max_row,
                "removed_row_count": removed_count,
                "col_count": max_col,
                "columns": [column_name(index) for index in range(max_col)],
                "layout": layout,
                "rows": rows,
            }
        )

    return {"sheets": sheets}


def parse_workbook(
    filename: str,
    file_bytes: bytes,
    display_sheet_name: str | None = None,
    only_first_sheet: bool = False,
) -> dict[str, Any]:
    extension = Path(filename).suffix.lower()
    if extension == ".xls":
        return parse_xls(file_bytes, display_sheet_name, only_first_sheet)
    if extension == ".xlsx":
        return parse_xlsx(file_bytes, display_sheet_name, only_first_sheet)
    raise ValueError("지원하지 않는 엑셀 형식입니다.")


def build_document_workbook(uploaded_files: Any) -> dict[str, Any]:
    sheets = []
    files = []

    for field_name, document_label in DOCUMENT_FIELDS:
        uploaded_file = uploaded_files.get(field_name)
        if uploaded_file is None or uploaded_file.filename == "":
            raise ValueError(f"{document_label} 파일을 선택해 주세요.")

        display_name = os.path.basename(uploaded_file.filename)
        if not allowed_file(display_name):
            raise ValueError(f"{document_label} 파일은 .xls 또는 .xlsx만 업로드할 수 있습니다.")

        parsed = parse_workbook(
            display_name,
            uploaded_file.read(),
            display_sheet_name=document_label,
            only_first_sheet=True,
        )
        if not parsed["sheets"]:
            raise ValueError(f"{document_label} 파일에서 첫 번째 시트를 찾을 수 없습니다.")

        sheet = parsed["sheets"][0]
        normalize_financial_sheet(sheet, document_label)
        sheet["filename"] = display_name
        sheet["document_label"] = document_label
        sheets.append(sheet)
        files.append({"label": document_label, "name": display_name})

    return {"sheets": sheets, "files": files}


def case_progress(case: dict[str, Any]) -> dict[str, str]:
    if case.get("collateral_saved"):
        return {"label": "결과 확인", "endpoint": "result"}
    if case.get("income_saved"):
        return {"label": "담보 입력", "endpoint": "collateral"}
    if case.get("debt_saved"):
        return {"label": "손익추정", "endpoint": "income"}
    if case.get("financial_saved"):
        return {"label": "채무입력", "endpoint": "debt"}
    return {"label": "자산가치산정", "endpoint": "financial"}


def case_diagnosis_label(case: dict[str, Any]) -> str:
    if not case.get("collateral_saved"):
        return "입력 진행 중"
    try:
        diagnosis = calculate_case_result(case)["diagnosis"]
    except Exception:
        return "결과 확인 필요"
    return "회생 가능성 검토 대상" if diagnosis.get("overall_positive") else "추가 검토 필요"


def diagnosis_chip(label: str, positive: bool) -> dict[str, str]:
    return {
        "label": label,
        "status": "회생유리" if positive else "회생불리",
        "tone": "positive" if positive else "negative",
    }


def case_diagnosis_summary(case: dict[str, Any]) -> list[dict[str, str]]:
    if not case.get("collateral_saved"):
        return [
            {
                "label": "진행",
                "status": "입력중",
                "tone": "neutral",
            }
        ]

    try:
        result = calculate_case_result(case)
        diagnosis = result["diagnosis"]
    except Exception:
        return [
            {
                "label": "진단",
                "status": "확인필요",
                "tone": "warning",
            }
        ]

    overall_positive = bool(diagnosis.get("overall_positive"))
    return [
        diagnosis_chip("가치 비교", bool(diagnosis.get("value_positive"))),
        diagnosis_chip("변제율 비교", bool(diagnosis.get("repayment_positive"))),
        diagnosis_chip("채권자 동의 가능성", bool(diagnosis.get("consent_positive"))),
        {
            "label": "종합 결론",
            "status": "회생유리" if overall_positive else "추가검토",
            "tone": "positive" if overall_positive else "warning",
        },
    ]


def case_listing_for_user(user_id: str, copied_id: str = "") -> list[dict[str, Any]]:
    listings = []
    for case in CASE_STORE.values():
        if case.get("user_id") != user_id:
            continue
        if case.get("deleted_at"):
            continue
        progress = case_progress(case)
        listings.append(
            {
                "case": case,
                "progress": progress,
                "diagnosis": case_diagnosis_label(case),
                "summary": case_diagnosis_summary(case),
            }
        )
    sorted_listings = sorted(
        listings,
        key=lambda item: item["case"].get("updated_at") or item["case"].get("created_at") or "",
        reverse=True,
    )
    if copied_id:
        copied_index = next(
            (
                index
                for index, item in enumerate(sorted_listings)
                if item["case"].get("case_id") == copied_id
            ),
            None,
        )
        if copied_index is not None:
            copied_item = sorted_listings.pop(copied_index)
            source_id = copied_item["case"].get("source_case_id")
            source_index = next(
                (
                    index
                    for index, item in enumerate(sorted_listings)
                    if item["case"].get("case_id") == source_id
                ),
                None,
            )
            if source_index is not None:
                sorted_listings.insert(source_index + 1, copied_item)
            else:
                sorted_listings.insert(copied_index, copied_item)
    return sorted_listings


def copy_case_for_user(source_case: dict[str, Any], user_id: str) -> dict[str, Any]:
    copied_case = deepcopy(source_case)
    original_id = source_case["case_id"]
    new_case_id = uuid.uuid4().hex[:12]
    created_at = now_iso()
    copied_case["case_id"] = new_case_id
    copied_case["source_case_id"] = original_id
    copied_case["user_id"] = user_id
    copied_case["created_at"] = created_at
    copied_case["updated_at"] = created_at
    copied_case["scenario_name"] = f"{source_case.get('scenario_name') or source_case.get('company_name', '분석')} 복사본"
    copied_case.pop("consultation", None)
    copied_case["consultation_submitted"] = False
    CASE_STORE[new_case_id] = copied_case
    remember_session_case(new_case_id)
    save_case(copied_case)
    return copied_case


@app.get("/")
def index():
    return render_template("landing.html", active_page="home")


@app.route("/signup", methods=["GET", "POST"])
def signup():
    if current_user():
        return redirect(url_for("mypage"))

    case_id = request.values.get("case_id", "")
    prefill_case = get_accessible_case(case_id) if case_id else None
    prefill_consultation = (prefill_case or {}).get("consultation", {})
    form = {
        "company": request.form.get("company", prefill_consultation.get("company", (prefill_case or {}).get("company_name", ""))).strip(),
        "contact_name": request.form.get("contact_name", prefill_consultation.get("contact_name", "")).strip(),
        "phone": request.form.get("phone", prefill_consultation.get("phone", "")).strip(),
        "email": normalize_email(request.form.get("email", prefill_consultation.get("email", ""))),
    }
    error = None

    if request.method == "POST":
        password = request.form.get("password", "")
        password_confirm = request.form.get("password_confirm", "")

        if not form["company"]:
            error = "회사명을 입력해 주세요."
        elif not form["contact_name"]:
            error = "담당자명을 입력해 주세요."
        elif not form["phone"]:
            error = "연락처를 입력해 주세요."
        elif not form["email"]:
            error = "이메일을 입력해 주세요."
        elif not EMAIL_RE.match(form["email"]):
            error = "이메일 형식을 확인해 주세요."
        elif find_user_by_email(form["email"]):
            error = "이미 가입된 이메일입니다."
        elif len(password) < 8:
            error = "비밀번호는 8자 이상 입력해 주세요."
        elif password != password_confirm:
            error = "비밀번호 확인이 일치하지 않습니다."
        elif request.form.get("privacy_consent") != "on":
            error = "개인정보 수집 및 이용에 동의해 주세요."
        else:
            user_id = uuid.uuid4().hex[:12]
            user = {
                "user_id": user_id,
                "company": form["company"],
                "contact_name": form["contact_name"],
                "phone": form["phone"],
                "email": form["email"],
                "password_hash": generate_password_hash(password),
                "email_verified": False,
                "created_at": now_iso(),
            }
            USERS[user_id] = user
            save_users()
            login_user(user)
            return redirect(url_for("mypage", joined="1"))

    return render_template(
        "signup.html",
        active_page="signup",
        form=form,
        case_id=case_id,
        error=error,
    )


@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user():
        return redirect(url_for("mypage"))

    email = normalize_email(request.form.get("email"))
    error = None
    if request.method == "POST":
        user = find_user_by_email(email)
        password = request.form.get("password", "")
        if not user or not check_password_hash(user.get("password_hash", ""), password):
            error = "이메일 또는 비밀번호를 확인해 주세요."
        else:
            login_user(user)
            next_path = request.args.get("next") or request.form.get("next") or url_for("mypage")
            if not next_path.startswith("/"):
                next_path = url_for("mypage")
            return redirect(next_path)

    return render_template(
        "login.html",
        active_page="login",
        email=email,
        next_path=request.args.get("next", ""),
        error=error,
    )


@app.route("/profile", methods=["GET", "POST"])
@login_required
def profile():
    user = current_user()
    assert user is not None
    form = {
        "company": request.form.get("company", user.get("company", "")).strip(),
        "contact_name": request.form.get("contact_name", user.get("contact_name", "")).strip(),
        "phone": request.form.get("phone", user.get("phone", "")).strip(),
        "email": normalize_email(request.form.get("email", user.get("email", ""))),
    }
    error = None

    if request.method == "POST":
        email_owner = find_user_by_email(form["email"]) if form["email"] else None
        if not form["company"]:
            error = "회사명을 입력해 주세요."
        elif not form["contact_name"]:
            error = "담당자명을 입력해 주세요."
        elif not form["phone"]:
            error = "연락처를 입력해 주세요."
        elif not form["email"]:
            error = "이메일을 입력해 주세요."
        elif not EMAIL_RE.match(form["email"]):
            error = "이메일 형식을 확인해 주세요."
        elif email_owner and email_owner.get("user_id") != user.get("user_id"):
            error = "이미 가입된 이메일입니다."
        else:
            user["company"] = form["company"]
            user["contact_name"] = form["contact_name"]
            user["phone"] = form["phone"]
            user["email"] = form["email"]
            user["updated_at"] = now_iso()
            save_users()
            return redirect(url_for("mypage", profile_updated="1"))

    return render_template(
        "profile.html",
        active_page="mypage",
        form=form,
        error=error,
    )


@app.post("/logout")
def logout():
    session.pop("user_id", None)
    return redirect(url_for("index"))


@app.get("/mypage")
@login_required
def mypage():
    user = current_user()
    assert user is not None
    return render_template(
        "mypage.html",
        active_page="mypage",
        cases=case_listing_for_user(user["user_id"], request.args.get("copied_id", "")),
        message=(
            "회원가입이 완료되었습니다. 분석 이력을 이곳에서 확인할 수 있습니다."
            if request.args.get("joined") == "1"
            else (
                "분석 이력을 삭제했습니다."
                if request.args.get("deleted") == "1"
                else (
                    "회원 정보를 수정했습니다."
                    if request.args.get("profile_updated") == "1"
                    else (
                        "복사본을 생성했습니다."
                        if request.args.get("copied") == "1"
                        else (
                            "분석 이름을 저장했습니다."
                            if request.args.get("renamed") == "1"
                            else None
                        )
                    )
                )
            )
        ),
    )


@app.post("/cases/<case_id>/copy")
@login_required
def copy_case(case_id: str):
    user = current_user()
    source_case = get_accessible_case(case_id)
    if user is None or source_case is None:
        return redirect(url_for("mypage"))
    copied_case = copy_case_for_user(source_case, user["user_id"])
    return redirect(url_for("mypage", copied="1", copied_id=copied_case["case_id"]))


@app.post("/cases/<case_id>/rename")
@login_required
def rename_case(case_id: str):
    case = get_accessible_case(case_id)
    if case is None:
        return redirect(url_for("mypage"))
    case["scenario_name"] = request.form.get("scenario_name", "").strip()
    save_case(case)
    return redirect(url_for("mypage", renamed="1"))


@app.post("/cases/<case_id>/delete")
@login_required
def delete_case(case_id: str):
    case = get_accessible_case(case_id)
    if case is None:
        return redirect(url_for("mypage"))
    case["deleted_at"] = now_iso()
    save_case(case)
    return redirect(url_for("mypage", deleted="1"))


@app.get("/analysis")
def analysis_index():
    return render_template("index.html", workbook=None, case=None, error=None)


@app.get("/upload")
def upload_form():
    return redirect(url_for("analysis_index"))


@app.get("/upload/<case_id>")
def upload_preview(case_id: str):
    case = get_accessible_case(case_id)
    if case is None:
        return render_template(
            "index.html",
            workbook=None,
            case=None,
            error="작업 정보를 찾을 수 없습니다. 파일을 다시 업로드해 주세요.",
        ), 404

    return render_template(
        "index.html",
        workbook=case["workbook"],
        case=case,
        error=None,
    )


@app.post("/upload")
def upload():
    try:
        workbook = build_document_workbook(request.files)
        case = create_case(workbook)
    except Exception as exc:
        return render_template(
            "index.html",
            workbook=None,
            case=None,
            error=f"엑셀을 읽는 중 오류가 발생했습니다: {exc}",
        ), 400

    return redirect(url_for("financial", case_id=case["case_id"], uploaded="1"))


@app.get("/financial/<case_id>")
def financial(case_id: str):
    case = get_accessible_case(case_id)
    if case is None:
        return render_template(
            "financial.html",
            case=None,
            rows=[],
            error="작업 정보를 찾을 수 없습니다. 파일을 다시 업로드해 주세요.",
            message=None,
        ), 404

    return render_template(
        "financial.html",
        case=case,
        rows=case["financial_rows"],
        error=None,
        message=(
            "파일이 업로드되었습니다. 자산가치산정을 확인해 주세요."
            if request.args.get("uploaded") == "1"
            else (
                "기존 분석을 복사했습니다. 필요한 숫자를 수정해 다시 분석할 수 있습니다."
                if request.args.get("copied") == "1"
                else None
            )
        ),
    )


@app.post("/financial/<case_id>/save")
def save_financial(case_id: str):
    case = get_accessible_case(case_id)
    if case is None:
        return redirect(url_for("analysis_index"))

    rows = case["financial_rows"]
    errors: list[str] = []

    for index, row in enumerate(rows):
        if not row["is_editable"]:
            continue

        audit_key = f"audit_value_{index}"
        liquidation_key = f"liquidation_value_{index}"
        audit_raw = request.form.get(audit_key, row["audit_value"])
        liquidation_raw = request.form.get(liquidation_key, row["liquidation_value"])

        audit_display, audit_number = clean_submitted_number(audit_raw)
        liquidation_display, liquidation_number = clean_submitted_number(liquidation_raw)

        if audit_number is not None and liquidation_number is not None and liquidation_number > audit_number:
            errors.append(f"{row['account']}: 청산가치는 실사가치보다 클 수 없습니다.")

        row["audit_value"] = audit_display
        row["liquidation_value"] = liquidation_display

    if errors:
        return render_template(
            "financial.html",
            case=case,
            rows=rows,
            error=" ".join(errors),
            message=None,
        ), 400

    case["financial_saved"] = True
    save_case(case)
    if request.form.get("next") == "debt":
        return redirect(url_for("debt", case_id=case_id))

    return render_template(
        "financial.html",
        case=case,
        rows=rows,
        error=None,
        message="재무 데이터가 저장되었습니다.",
    )


@app.get("/debt/<case_id>")
def debt(case_id: str):
    case = get_accessible_case(case_id)
    if case is None:
        return render_template(
            "debt.html",
            case=None,
            rows=[],
            error="작업 정보를 찾을 수 없습니다. 파일을 다시 업로드해 주세요.",
            message=None,
        ), 404

    return render_template(
        "debt.html",
        case=case,
        rows=case["debt_rows"],
        error=None,
        message=None,
    )


@app.post("/debt/<case_id>/save")
def save_debt(case_id: str):
    case = get_accessible_case(case_id)
    if case is None:
        return redirect(url_for("analysis_index"))

    rows = case["debt_rows"]
    for index, row in enumerate(rows):
        raw_amount = request.form.get(f"debt_amount_{index}", row["debt_amount"])
        amount_display, _ = clean_submitted_number(raw_amount)
        row["debt_amount"] = amount_display
        row["collateral_type"] = request.form.get(f"collateral_type_{index}", row.get("collateral_type", ""))
        row["audit_value"] = request.form.get(f"audit_value_{index}", row.get("audit_value", ""))

    case["debt_saved"] = True
    save_case(case)
    if request.form.get("next") == "income":
        return redirect(url_for("income", case_id=case_id))

    return render_template(
        "debt.html",
        case=case,
        rows=rows,
        error=None,
        message="부채 데이터가 저장되었습니다. 다음 단계는 손익추정입니다.",
    )


@app.get("/income/<case_id>")
def income(case_id: str):
    case = get_accessible_case(case_id)
    if case is None:
        return render_template(
            "income.html",
            case=None,
            sales_rows=[],
            expense_rows=[],
            error="작업 정보를 찾을 수 없습니다. 파일을 다시 업로드해 주세요.",
            message=None,
        ), 404

    rows = case["income_rows"]
    return render_template(
        "income.html",
        case=case,
        sales_rows=[row for row in rows if row["section"] == "sales"],
        expense_rows=[row for row in rows if row["section"] == "expense"],
        error=None,
        message=None,
    )


@app.post("/income/<case_id>/save")
def save_income(case_id: str):
    case = get_accessible_case(case_id)
    if case is None:
        return redirect(url_for("analysis_index"))

    rows = case["income_rows"]
    for index, row in enumerate(rows):
        if not row["is_editable"]:
            continue

        row["cost_type"] = request.form.get(f"cost_type_{index}", row.get("cost_type", ""))
        row["monthly_average_display"] = request.form.get(
            f"monthly_average_display_{index}",
            row.get("monthly_average_display", ""),
        )

        value_kind = request.form.get(f"value_kind_{index}", row.get("value_kind", "amount"))
        final_raw = request.form.get(f"final_value_{index}", row["final_value"])
        final_display, _ = clean_submitted_income_value(final_raw, value_kind)

        row["value_kind"] = value_kind
        row["final_value"] = final_display

    case["income_saved"] = True
    save_case(case)
    if request.form.get("next") == "collateral":
        return redirect(url_for("collateral", case_id=case_id))

    return render_template(
        "income.html",
        case=case,
        sales_rows=[row for row in rows if row["section"] == "sales"],
        expense_rows=[row for row in rows if row["section"] == "expense"],
        error=None,
        message="손익 데이터가 저장되었습니다. 다음 단계는 담보등자산입니다.",
    )


@app.get("/collateral/<case_id>")
def collateral(case_id: str):
    case = get_accessible_case(case_id)
    if case is None:
        return render_template(
            "collateral.html",
            case=None,
            collateral_rows=[],
            rent_rows=[],
            error="작업 정보를 찾을 수 없습니다. 파일을 다시 업로드해 주세요.",
            message=None,
        ), 404

    return render_template(
        "collateral.html",
        case=case,
        collateral_rows=case["collateral_rows"],
        rent_rows=case["rent_rows"],
        error=None,
        message=None,
    )


@app.post("/collateral/<case_id>/save")
def save_collateral(case_id: str):
    case = get_accessible_case(case_id)
    if case is None:
        return redirect(url_for("analysis_index"))

    collateral_rows = case["collateral_rows"]
    rent_rows = case["rent_rows"]
    errors: list[str] = []

    for index, row in enumerate(collateral_rows):
        audit_raw = request.form.get(f"audit_value_{index}", row["audit_value"])
        liquidation_raw = request.form.get(
            f"liquidation_value_{index}",
            row["liquidation_value"],
        )

        audit_display, audit_number = clean_submitted_number(audit_raw)
        liquidation_display, liquidation_number = clean_submitted_number(liquidation_raw)

        if audit_number is not None and liquidation_number is not None and liquidation_number > audit_number:
            errors.append(f"{row['category']}: 청산가치는 실사가치보다 클 수 없습니다.")

        row["audit_value"] = audit_display
        row["liquidation_value"] = liquidation_display

    for index, row in enumerate(rent_rows):
        amount_raw = request.form.get(f"rent_amount_{index}", row["amount"])
        amount_display, _ = clean_submitted_number(amount_raw)
        row["amount"] = amount_display

    if errors:
        return render_template(
            "collateral.html",
            case=case,
            collateral_rows=collateral_rows,
            rent_rows=rent_rows,
            error=" ".join(errors),
            message=None,
        ), 400

    case["collateral_saved"] = True
    save_case(case)
    if request.form.get("next") == "result":
        return redirect(url_for("result", case_id=case_id))

    return render_template(
        "collateral.html",
        case=case,
        collateral_rows=collateral_rows,
        rent_rows=rent_rows,
        error=None,
        message="담보등자산 데이터가 저장되었습니다. 다음 단계는 결과확인입니다.",
    )


@app.get("/result/<case_id>")
def result(case_id: str):
    case = get_accessible_case(case_id)
    if case is None:
        return render_template(
            "result.html",
            case=None,
            error="작업 정보를 찾을 수 없습니다. 파일을 다시 업로드해 주세요.",
        ), 404

    return render_template(
        "result.html",
        case=case,
        calculation_result=calculate_case_result(case),
        error=None,
        message="상담 신청이 접수되었습니다. 상세 분석 결과를 확인할 수 있습니다."
        if request.args.get("consultation") == "submitted"
        else None,
        consultation_form=case.get("consultation", {}),
    )


@app.post("/consultation/<case_id>")
def submit_consultation(case_id: str):
    case = get_accessible_case(case_id)
    if case is None:
        return redirect(url_for("analysis_index"))

    form = consultation_form_from_request()
    errors = validate_consultation_form(form)
    result_data = calculate_case_result(case)

    if errors:
        return render_template(
            "result.html",
            case=case,
            calculation_result=result_data,
            error=" ".join(errors),
            message=None,
            consultation_form=form,
        ), 400

    append_consultation_log(case, form, result_data)
    case["consultation"] = form
    case["consultation_submitted"] = True
    save_case(case)
    return redirect(url_for("result", case_id=case_id, consultation="submitted"))


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
