from __future__ import annotations

import math
from dataclasses import dataclass
from datetime import date
from typing import Any


@dataclass(frozen=True)
class CalculationAssumptions:
    cpi_rate: float = 0.019
    bond_yield_3y: float = 0.02335
    repayment_present_value_rate: float = 0.0542
    liquidation_cost_rate: float = 0.05
    terminal_discount_spread: float = 0.065
    collateral_interest_spread: float = 0.005
    collateral_disposal_rate: float = 0.95


DEBT_FIELDS = {
    "secured_debt": "담보채무",
    "unsecured_financial_debt": "무담보 금융기관채무",
    "other_unsecured_debt": "기타 무담보채무(상거래채무 등)",
    "related_party_debt": "특수관계인채무",
    "unpaid_wages": "미지급급여, 미지급퇴직금(세후)",
    "retirement_benefit": "퇴직급여추계액",
    "tax_arrears": "조세체납금액(4대보험체납금액 포함)",
}


def parse_number(value: Any) -> float:
    text = str(value or "").replace(",", "").replace("%", "").strip()
    if not text:
        return 0.0
    is_parenthesized_negative = text.startswith("(") and text.endswith(")")
    text = text.strip("()")
    try:
        number = float(text)
    except ValueError:
        return 0.0
    return -number if is_parenthesized_negative else number


def parse_percent(value: Any) -> float:
    return parse_number(value) / 100


def safe_div(numerator: float, denominator: float) -> float:
    if denominator == 0:
        return 0.0
    return numerator / denominator


def excel_min(*values: float) -> float:
    return min(values)


def excel_max(*values: float) -> float:
    return max(values)


def display_number(value: float) -> str:
    if not math.isfinite(value):
        return str(value)
    rounded = round(value)
    return f"{rounded:,}"


def display_percent(value: float) -> str:
    if not math.isfinite(value):
        return ""
    return f"{value * 100:.1f}%"


def visual_width(value: float, max_value: float) -> float:
    if not math.isfinite(value) or max_value <= 0:
        return 0.0
    width = abs(value) / max_value * 100
    if value != 0:
        width = max(width, 4.0)
    return min(width, 100.0)


def build_result_visual(
    liquidation_value: float,
    going_concern_value: float,
) -> dict[str, Any]:
    max_value = max(abs(liquidation_value), abs(going_concern_value), 1.0)
    value_difference = going_concern_value - liquidation_value

    return {
        "value_bars": [
            {
                "label": "청산가치",
                "value": liquidation_value,
                "width": visual_width(liquidation_value, max_value),
                "tone": "liquidation",
                "is_negative": liquidation_value < 0,
            },
            {
                "label": "계속기업가치",
                "value": going_concern_value,
                "width": visual_width(going_concern_value, max_value),
                "tone": "going",
                "is_negative": going_concern_value < 0,
            },
        ],
        "value_status": "회생가치 우위" if value_difference >= 0 else "청산가치 우위",
        "value_difference": value_difference,
    }


def row_number(row: dict[str, Any], key: str) -> float:
    return parse_number(row.get(key, ""))


def account_contains(row: dict[str, Any], text: str) -> bool:
    return text in str(row.get("account", "") or row.get("category", ""))


def calculate_asset_totals(financial_rows: list[dict[str, Any]]) -> dict[str, Any]:
    sections = {
        "current": {"statement": 0.0, "audit": 0.0, "liquidation": 0.0},
        "non_current": {"statement": 0.0, "audit": 0.0, "liquidation": 0.0},
    }
    section: str | None = None
    cash_audit_value = 0.0

    for row in financial_rows:
        account = str(row.get("account", ""))
        if "비유동자산" in account:
            section = "non_current"
            continue
        if "유동자산" in account:
            section = "current"
            continue
        section_for_row = "non_current" if row.get("row_type") == "custom_asset" else section
        if section_for_row is None or not row.get("is_editable"):
            continue

        statement = row_number(row, "amount")
        audit = row_number(row, "audit_value")
        liquidation = row_number(row, "liquidation_value")
        sections[section_for_row]["statement"] += statement
        sections[section_for_row]["audit"] += audit
        sections[section_for_row]["liquidation"] += liquidation

        if cash_audit_value == 0 and "현금" in account:
            cash_audit_value = audit

    totals = {
        "statement": sections["current"]["statement"] + sections["non_current"]["statement"],
        "audit": sections["current"]["audit"] + sections["non_current"]["audit"],
        "liquidation": sections["current"]["liquidation"] + sections["non_current"]["liquidation"],
    }
    return {"sections": sections, "totals": totals, "cash_audit_value": cash_audit_value}


def debt_amounts(debt_rows: list[dict[str, Any]]) -> dict[str, float]:
    return {row["field"]: row_number(row, "debt_amount") for row in debt_rows}


def collateral_amounts(
    collateral_rows: list[dict[str, Any]],
    rent_rows: list[dict[str, Any]],
) -> dict[str, Any]:
    collateral = {
        row["field"]: {
            "audit": row_number(row, "audit_value"),
            "liquidation": row_number(row, "liquidation_value"),
        }
        for row in collateral_rows
    }
    rent = {row["field"]: row_number(row, "amount") for row in rent_rows}
    total_audit = sum(item["audit"] for item in collateral.values())
    total_liquidation = sum(item["liquidation"] for item in collateral.values())
    return {
        "collateral": collateral,
        "rent": rent,
        "total_audit": total_audit,
        "total_liquidation": total_liquidation,
    }


def income_assumptions(income_rows: list[dict[str, Any]]) -> dict[str, Any]:
    sales_row = next((row for row in income_rows if account_contains(row, "매출액")), None)
    cost_row = next((row for row in income_rows if account_contains(row, "매출원가")), None)

    sales_base = row_number(sales_row or {}, "y")
    sales_growth_rate = parse_percent((sales_row or {}).get("final_value", ""))
    cost_rate = parse_percent((cost_row or {}).get("final_value", ""))

    expense_rows = [
        row
        for row in income_rows
        if row.get("section") == "expense" and row.get("is_editable")
    ]

    return {
        "sales_base": sales_base,
        "sales_growth_rate": sales_growth_rate,
        "cost_rate": cost_rate,
        "expense_rows": expense_rows,
    }


def default_preparation_dates(today: date | None = None) -> dict[str, Any]:
    today = today or date.today()
    start = date(today.year, 1, 1)
    end = date(today.year, 12, 31)
    prep_days = (end - today).days + 1
    total_days = (end - start).days + 1
    return {
        "today": today,
        "start": start,
        "end": end,
        "prep_days": max(prep_days, 0),
        "total_days": total_days,
    }


def project_operating_profit(
    income: dict[str, Any],
    monthly_rent: float,
    dates: dict[str, Any],
    assumptions: CalculationAssumptions,
) -> dict[str, Any]:
    prep_ratio = safe_div(dates["prep_days"], dates["total_days"])
    sales_growth = income["sales_growth_rate"]
    cost_rate = income["cost_rate"]
    periods = 11

    sales = [0.0] * periods
    sales[0] = income["sales_base"] * (1 + sales_growth) * prep_ratio
    annualized_prep_sales = safe_div(sales[0], prep_ratio)
    sales[1] = annualized_prep_sales * (1 + sales_growth)
    for index in range(2, periods):
        sales[index] = sales[index - 1] * (1 + sales_growth)

    cost_of_sales = [value * cost_rate for value in sales]
    gross_profit = [sales[index] - cost_of_sales[index] for index in range(periods)]

    sgna = [0.0] * periods
    expense_details: list[dict[str, Any]] = []
    for row in income["expense_rows"]:
        cost_type = row.get("cost_type", "fixed")
        final_value = row.get("final_value", "")
        values = [0.0] * periods

        if cost_type == "variable":
            rate = parse_percent(final_value)
            values = [period_sales * rate for period_sales in sales]
        else:
            monthly_amount = parse_number(final_value)
            values[0] = monthly_amount * 12 * prep_ratio
            values[1] = monthly_amount * 12 * (1 + assumptions.cpi_rate)
            for index in range(2, periods):
                values[index] = values[index - 1] * (1 + assumptions.cpi_rate)

        for index, value in enumerate(values):
            sgna[index] += value
        expense_details.append(
            {
                "account": row.get("account", ""),
                "cost_type": cost_type,
                "values": values,
            }
        )

    rent_expense = [0.0] * periods
    rent_expense[0] = monthly_rent * 12 * prep_ratio
    rent_expense[1] = monthly_rent * 12 * (1 + assumptions.cpi_rate)
    for index in range(2, periods):
        rent_expense[index] = rent_expense[index - 1] * (1 + assumptions.cpi_rate)
    for index, value in enumerate(rent_expense):
        sgna[index] += value

    operating_profit = [gross_profit[index] - sgna[index] for index in range(periods)]
    return {
        "sales": sales,
        "cost_of_sales": cost_of_sales,
        "gross_profit": gross_profit,
        "sgna": sgna,
        "rent_expense": rent_expense,
        "operating_profit": operating_profit,
        "expense_details": expense_details,
    }


def calculate_liquidation_distribution(
    debts: dict[str, float],
    assets: dict[str, Any],
    collateral: dict[str, Any],
    assumptions: CalculationAssumptions,
) -> dict[str, Any]:
    secured_collateral_value = (
        collateral["collateral"].get("collateral_except_machinery", {}).get("audit", 0.0)
        + collateral["collateral"].get("collateral_machinery", {}).get("audit", 0.0)
    )
    other_liquidation_value = assets["totals"]["liquidation"] - secured_collateral_value

    secured_pool = secured_collateral_value * (1 - assumptions.liquidation_cost_rate)
    unsecured_pool = other_liquidation_value * (1 - assumptions.liquidation_cost_rate)

    debt = {
        "unpaid_wages": debts.get("unpaid_wages", 0.0),
        "retirement_benefit": debts.get("retirement_benefit", 0.0),
        "tax_arrears": debts.get("tax_arrears", 0.0),
        "secured_debt": debts.get("secured_debt", 0.0),
        "unsecured_financial_debt": debts.get("unsecured_financial_debt", 0.0),
        "other_unsecured_debt": debts.get("other_unsecured_debt", 0.0),
        "related_party_debt": debts.get("related_party_debt", 0.0),
    }

    secured_repay: dict[str, float] = {}
    unsecured_repay: dict[str, float] = {}
    residual_repay: dict[str, float] = {}

    secured_repay["secured_debt"] = excel_min(debt["secured_debt"], secured_pool)
    secured_repay["unpaid_wages"] = excel_min(
        debt["unpaid_wages"],
        secured_pool - secured_repay["secured_debt"],
    )
    secured_repay["retirement_benefit"] = excel_min(
        debt["retirement_benefit"],
        secured_pool - secured_repay["secured_debt"] - secured_repay["unpaid_wages"],
    )
    secured_repay["tax_arrears"] = excel_min(
        debt["tax_arrears"],
        secured_pool
        - secured_repay["secured_debt"]
        - secured_repay["unpaid_wages"]
        - secured_repay["retirement_benefit"],
    )

    unsecured_repay["unpaid_wages"] = excel_min(
        debt["unpaid_wages"] - secured_repay["unpaid_wages"],
        unsecured_pool,
    )
    unsecured_repay["retirement_benefit"] = excel_min(
        debt["retirement_benefit"] - secured_repay["retirement_benefit"],
        unsecured_pool - unsecured_repay["unpaid_wages"],
    )
    unsecured_repay["tax_arrears"] = excel_min(
        debt["tax_arrears"] - secured_repay["tax_arrears"],
        unsecured_pool - unsecured_repay["unpaid_wages"] - unsecured_repay["retirement_benefit"],
    )

    secured_unsecured_denominator = (
        debt["secured_debt"]
        + debt["unsecured_financial_debt"]
        + debt["other_unsecured_debt"]
        + debt["related_party_debt"]
    )
    unsecured_priority_remainder = unsecured_pool - (
        unsecured_repay["unpaid_wages"]
        + unsecured_repay["retirement_benefit"]
        + unsecured_repay["tax_arrears"]
    )
    unsecured_group_denominator = (
        debt["unsecured_financial_debt"]
        + debt["other_unsecured_debt"]
        + debt["related_party_debt"]
    )

    for key in ("secured_debt", "unsecured_financial_debt", "other_unsecured_debt", "related_party_debt"):
        secured_amount = secured_repay.get(key, 0.0)
        unsecured_repay[key] = excel_min(
            debt[key] - secured_amount,
            unsecured_priority_remainder * safe_div(debt[key], secured_unsecured_denominator),
        )

    unsecured_total = sum(unsecured_repay.values())
    residual_pool = unsecured_pool - unsecured_total
    for key in ("unsecured_financial_debt", "other_unsecured_debt", "related_party_debt"):
        residual_repay[key] = excel_min(
            debt[key] - secured_repay.get(key, 0.0) - unsecured_repay.get(key, 0.0),
            residual_pool * safe_div(debt[key], unsecured_group_denominator),
        )

    rows: dict[str, dict[str, float]] = {}
    for key, amount in debt.items():
        repayment = (
            secured_repay.get(key, 0.0)
            + unsecured_repay.get(key, 0.0)
            + residual_repay.get(key, 0.0)
        )
        rows[key] = {
            "debt": amount,
            "repayment": repayment,
            "rate": safe_div(repayment, amount),
            "unpaid": amount - repayment,
        }

    return {
        "secured_collateral_value": secured_collateral_value,
        "other_liquidation_value": other_liquidation_value,
        "secured_pool": secured_pool,
        "unsecured_pool": unsecured_pool,
        "total_pool": secured_pool + unsecured_pool,
        "secured_repay": secured_repay,
        "unsecured_repay": unsecured_repay,
        "residual_repay": residual_repay,
        "rows": rows,
    }


def calculate_going_concern_repayment(
    debts: dict[str, float],
    liquidation_distribution: dict[str, Any],
    operating_projection: dict[str, Any],
    assets: dict[str, Any],
    collateral: dict[str, Any],
    dates: dict[str, Any],
    assumptions: CalculationAssumptions,
) -> dict[str, Any]:
    collateral_total_audit = collateral["total_audit"]
    machinery_audit = collateral["collateral"].get("collateral_machinery", {}).get("audit", 0.0)
    rent_deposit = collateral["rent"].get("rent_deposit", 0.0)

    beginning_cash = assets["cash_audit_value"]
    operating_cashflow = sum(operating_projection["operating_profit"])
    asset_sale_cashflow = (collateral_total_audit - machinery_audit) * assumptions.collateral_disposal_rate
    rent_deposit_outflow = -rent_deposit
    cash_inflow_total = beginning_cash + operating_cashflow + asset_sale_cashflow + rent_deposit_outflow

    debt = {
        "unpaid_wages": debts.get("unpaid_wages", 0.0),
        "retirement_benefit": debts.get("retirement_benefit", 0.0),
        "tax_arrears": debts.get("tax_arrears", 0.0),
        "secured_debt": debts.get("secured_debt", 0.0),
        "unsecured_financial_debt": debts.get("unsecured_financial_debt", 0.0),
        "other_unsecured_debt": debts.get("other_unsecured_debt", 0.0),
        "related_party_debt": debts.get("related_party_debt", 0.0),
    }

    repay: dict[str, float] = {}
    repay["unpaid_wages"] = excel_max(excel_min(debt["unpaid_wages"], cash_inflow_total), 0.0)
    repay["retirement_benefit"] = 0.0
    repay["tax_arrears"] = excel_max(
        excel_min(debt["tax_arrears"], cash_inflow_total - repay["unpaid_wages"]),
        0.0,
    )

    secured_liquidation_repay = liquidation_distribution["rows"]["secured_debt"]["repayment"]
    repay["secured_debt"] = excel_max(
        excel_min(
            cash_inflow_total - repay["unpaid_wages"] - repay["tax_arrears"],
            excel_max(debt["secured_debt"], secured_liquidation_repay),
        ),
        0.0,
    )

    interest_days = dates["prep_days"] + dates["total_days"]
    repayment_interest_rate = assumptions.repayment_present_value_rate + assumptions.collateral_interest_spread
    repay["secured_debt_interest"] = excel_min(
        cash_inflow_total
        - repay["unpaid_wages"]
        - repay["tax_arrears"]
        - repay["secured_debt"],
        repay["secured_debt"] * interest_days / 365 * repayment_interest_rate,
    )

    unsecured_denominator = debt["unsecured_financial_debt"] + debt["other_unsecured_debt"]
    unsecured_remainder = cash_inflow_total - (
        repay["unpaid_wages"]
        + repay["tax_arrears"]
        + repay["secured_debt"]
        + repay["secured_debt_interest"]
    )
    repay["unsecured_financial_debt"] = excel_max(
        excel_min(
            debt["unsecured_financial_debt"],
            unsecured_remainder * safe_div(debt["unsecured_financial_debt"], unsecured_denominator),
        ),
        0.0,
    )
    repay["other_unsecured_debt"] = excel_max(
        excel_min(
            debt["other_unsecured_debt"],
            unsecured_remainder * safe_div(debt["other_unsecured_debt"], unsecured_denominator),
        ),
        0.0,
    )
    repay["related_party_debt"] = 0.0

    rows: dict[str, dict[str, float]] = {}
    for key, amount in debt.items():
        repayment = repay.get(key, 0.0)
        rows[key] = {
            "debt": amount,
            "repayment": repayment,
            "rate": safe_div(repayment, amount),
        }

    rows["secured_debt_interest"] = {
        "debt": 0.0,
        "repayment": repay["secured_debt_interest"],
        "rate": 0.0,
    }

    ending_cash = cash_inflow_total - sum(item["repayment"] for item in rows.values())
    return {
        "beginning_cash": beginning_cash,
        "operating_cashflow": operating_cashflow,
        "asset_sale_cashflow": asset_sale_cashflow,
        "rent_deposit_outflow": rent_deposit_outflow,
        "cash_inflow_total": cash_inflow_total,
        "rows": rows,
        "ending_cash": ending_cash,
    }


def calculate_enterprise_value(
    operating_projection: dict[str, Any],
    collateral: dict[str, Any],
    assumptions: CalculationAssumptions,
    dates: dict[str, Any],
) -> dict[str, Any]:
    discount_rate = assumptions.bond_yield_3y + assumptions.terminal_discount_spread
    prep_fraction = safe_div(dates["prep_days"], dates["total_days"])
    fractions = [prep_fraction + index for index in range(11)]
    pv_factors = [1 / ((1 + discount_rate) ** fraction) for fraction in fractions]
    operating_values = operating_projection["operating_profit"]
    pv_10_years = sum(value * factor for value, factor in zip(operating_values, pv_factors))
    terminal_cashflow = operating_values[-1] / discount_rate if discount_rate else 0.0
    terminal_value = terminal_cashflow * pv_factors[-1]
    non_business_asset_value = (
        collateral["total_audit"]
        - collateral["collateral"].get("collateral_machinery", {}).get("audit", 0.0)
    )
    going_concern_value = pv_10_years + terminal_value + non_business_asset_value
    return {
        "discount_rate": discount_rate,
        "period_fractions": fractions,
        "pv_factors": pv_factors,
        "pv_10_years": pv_10_years,
        "terminal_cashflow": terminal_cashflow,
        "terminal_value": terminal_value,
        "non_business_asset_value": non_business_asset_value,
        "going_concern_value": going_concern_value,
    }


def build_diagnosis(
    liquidation_value: float,
    going_concern_value: float,
    comparison_rows: list[dict[str, Any]],
) -> dict[str, Any]:
    value_positive = going_concern_value > liquidation_value
    repayment_positive = all(row["rate_difference"] >= 0 for row in comparison_rows[:5])
    unsecured_financial_rate = next(
        (
            row["going_rate"]
            for row in comparison_rows
            if row["field"] == "unsecured_financial_debt"
        ),
        0.0,
    )
    consent_applicable = value_positive and repayment_positive
    consent_positive = consent_applicable and unsecured_financial_rate >= 0.30
    overall_positive = value_positive and repayment_positive and unsecured_financial_rate >= 0.30

    value_message = (
        f"회사의 청산가치는 {display_number(liquidation_value / 1_000_000)}백만원이나, "
        "계속기업 가정 시 영업이익의 현재가치, 비영업용 자산의 가치 등을 합하면 "
        f"{display_number(going_concern_value / 1_000_000)}백만원이고, "
        "이는 청산가치보다 크므로 회생절차를 진행하는 것이 타당한 것으로 판단될 수 있습니다."
        if value_positive
        else "회사 영업이익의 현재가치 및 비영업용 자산의 가치 등은 "
        f"{display_number(going_concern_value / 1_000_000)}백만원이고, "
        f"청산가치는 {display_number(liquidation_value / 1_000_000)}백만원이므로, "
        "회생절차를 진행하는 것이 적절하지 않은 것으로 판단될 수 있습니다."
    )
    repayment_message = (
        "담보채무 및 무담보채무에 대하여 계속기업 가정 시 변제율이 청산시 배당액보다 크므로 회생절차를 진행하는 것이 타당한 것으로 판단될 수 있습니다."
        if repayment_positive
        else "담보채무 또는 무담보채무에 대하여 계속기업 가정 시 변제율이 청산시 배당액보다 낮으므로 회생절차를 진행하는 것이 적절하지 않은 것으로 판단될 수 있습니다."
    )
    consent_message = (
        f"계속기업 가정 시 회생채권의 명목변제율이 약 {display_number(unsecured_financial_rate * 100)}%에 해당하여 채권자의 동의를 구할 가능성이 있는 것으로 판단됩니다."
        if consent_positive
        else (
            ""
            if not consent_applicable
            else f"계속기업 가정 시 회생채권의 명목변제율이 {display_number(unsecured_financial_rate * 100)}%에 해당하여 채권자의 동의를 구할 가능성이 높지 않을 것으로 판단됩니다."
        )
    )
    overall_message = (
        "위의 사항을 종합적으로 고려할 때, 회생절차를 진행하는 것이 가능할 것으로 판단됩니다. "
        "일반적으로 임금채권, 조세채권과 담보채권은 개시 신청연도부터 3년 이내에 상환을 하여야 하며, "
        "담보를 제공하지 않은 기타 금융기관 채권과 상거래채권은 채권액의 약 60%~70%는 출자전환이 되고 "
        "나머지 30%~40%의 채권액을 5년 ~10년 동안 장기간 동안 분할 상환하여야 합니다."
        if overall_positive
        else "위의 사항을 종합적으로 고려할 때, 입력하신 조건으로는 회생절차를 진행하는 것이 쉽지 않을 것으로 판단될 수 있으나, "
        "청산가치와 계속기업가치 산정의 결과나 채무의 변제가능성 정도에 따라 회생절차를 진행하는 것이 가능할 수 있습니다."
    )

    return {
        "value_positive": value_positive,
        "repayment_positive": repayment_positive,
        "consent_applicable": consent_applicable,
        "consent_positive": consent_positive,
        "overall_positive": overall_positive,
        "value_message": value_message,
        "repayment_message": repayment_message,
        "consent_message": consent_message,
        "overall_message": overall_message,
    }


WORKSHEET_COLUMNS = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"]
PERIOD_COLUMNS = WORKSHEET_COLUMNS[:11]
PERIOD_LABELS = [
    "준비연도",
    "회생1차년도",
    "회생2차년도",
    "회생3차년도",
    "회생4차년도",
    "회생5차년도",
    "회생6차년도",
    "회생7차년도",
    "회생8차년도",
    "회생9차년도",
    "회생10차년도",
]


def worksheet_value(value: Any, kind: str = "number") -> str:
    if value in (None, ""):
        return ""
    if kind == "text":
        return str(value)
    if kind == "date":
        return value.isoformat() if hasattr(value, "isoformat") else str(value)
    if kind == "percent":
        return display_percent(float(value))
    if kind == "decimal":
        return f"{float(value):.4f}"
    return display_number(float(value))


def worksheet_row(
    row_number: int | str,
    label: str,
    cells: dict[str, Any] | None = None,
    row_type: str = "data",
) -> dict[str, Any]:
    cells = cells or {}
    values = []
    for column in WORKSHEET_COLUMNS:
        value = cells.get(column, "")
        if isinstance(value, tuple):
            value, kind = value
            values.append(worksheet_value(value, kind))
        else:
            values.append(worksheet_value(value))
    return {
        "row_number": row_number,
        "label": label,
        "type": row_type,
        "values": values,
    }


def section_row(label: str) -> dict[str, Any]:
    return worksheet_row("", label, row_type="section")


def income_source_row(income_rows: list[dict[str, Any]], text: str) -> dict[str, Any]:
    return next((row for row in income_rows if text in str(row.get("account", ""))), {})


def period_cells(values: list[float], total: bool = False) -> dict[str, Any]:
    cells = {column: value for column, value in zip(PERIOD_COLUMNS, values)}
    if total:
        cells["N"] = sum(values)
    return cells


def build_worksheet_review(
    case: dict[str, Any],
    assets: dict[str, Any],
    debts: dict[str, float],
    collateral: dict[str, Any],
    income: dict[str, Any],
    operating_projection: dict[str, Any],
    liquidation_distribution: dict[str, Any],
    enterprise_value: dict[str, Any],
    going_repayment: dict[str, Any],
    comparison_rows: list[dict[str, Any]],
    diagnosis: dict[str, Any],
    assumptions: CalculationAssumptions,
    dates: dict[str, Any],
) -> dict[str, Any]:
    rows: list[dict[str, Any]] = []
    financial_sections = assets["sections"]
    income_rows = case.get("income_rows", [])
    sales_row = income_source_row(income_rows, "매출액")
    cost_row = income_source_row(income_rows, "매출원가")
    operating_profit_row = income_source_row(income_rows, "영업이익")
    pv_values = [
        value * factor
        for value, factor in zip(
            operating_projection["operating_profit"],
            enterprise_value["pv_factors"],
        )
    ]

    rows.append(section_row("1. 자산항목 평가"))
    rows.append(
        worksheet_row(
            7,
            "I.유동자산",
            {
                "C": financial_sections["current"]["statement"],
                "D": financial_sections["current"]["audit"],
                "E": financial_sections["current"]["liquidation"],
            },
        )
    )
    rows.append(
        worksheet_row(
            38,
            "Ⅱ.비유동자산",
            {
                "C": financial_sections["non_current"]["statement"],
                "D": financial_sections["non_current"]["audit"],
                "E": financial_sections["non_current"]["liquidation"],
            },
        )
    )
    rows.append(
        worksheet_row(
            315,
            "자산 총계",
            {
                "C": assets["totals"]["statement"],
                "D": assets["totals"]["audit"],
                "E": assets["totals"]["liquidation"],
            },
        )
    )

    rows.append(section_row("2. 부채현황 입력"))
    debt_rows = [
        (71, "secured_debt"),
        (72, "unsecured_financial_debt"),
        (73, "other_unsecured_debt"),
        (74, "related_party_debt"),
        (75, "unpaid_wages"),
        (76, "retirement_benefit"),
        (77, "tax_arrears"),
    ]
    for row_number_value, key in debt_rows:
        rows.append(worksheet_row(row_number_value, DEBT_FIELDS[key], {"C": debts.get(key, 0.0)}))
    rows.append(worksheet_row(78, "합계", {"C": sum(debts.values())}))

    rows.append(section_row("3. 과거 손익현황 및 향후 손익추정 기초값 산정"))
    rows.append(
        worksheet_row(
            86,
            "Ⅰ.매출액",
            {
                "C": sales_row.get("y_minus_1_number"),
                "D": sales_row.get("y_number"),
                "F": (sales_row.get("metric_display", ""), "text"),
                "G": (income["sales_growth_rate"], "percent"),
            },
        )
    )
    rows.append(
        worksheet_row(
            87,
            "Ⅱ.매출원가",
            {
                "C": cost_row.get("y_minus_1_number"),
                "D": cost_row.get("y_number"),
                "F": (cost_row.get("metric_display", ""), "text"),
                "G": (income["cost_rate"], "percent"),
            },
        )
    )
    rows.append(
        worksheet_row(
            129,
            "Ⅴ.영업이익",
            {
                "C": operating_profit_row.get("y_minus_1_number"),
                "D": operating_profit_row.get("y_number"),
            },
        )
    )

    rows.append(section_row("4. 담보제공 부동산 및 비영업용자산 매각 가정"))
    collateral_rows = [
        (136, "collateral_except_machinery", "담보제공자산(기계장치 제외)"),
        (137, "collateral_machinery", "담보제공 기계장치"),
        (138, "savings", "정기예.적금"),
        (139, "insurance", "보험해약환급금"),
        (140, "securities", "유가증권"),
        (141, "other_non_business_assets", "기타 비업무용 자산"),
    ]
    for row_number_value, key, label in collateral_rows:
        item = collateral["collateral"].get(key, {})
        rows.append(
            worksheet_row(
                row_number_value,
                label,
                {"C": item.get("audit", 0.0), "D": item.get("liquidation", 0.0)},
            )
        )
    rows.append(worksheet_row(142, "합계", {"C": collateral["total_audit"], "D": collateral["total_liquidation"]}))
    rows.append(worksheet_row(146, "임차보증금", {"C": collateral["rent"].get("rent_deposit", 0.0)}))
    rows.append(worksheet_row(147, "월세", {"C": collateral["rent"].get("monthly_rent", 0.0)}))

    rows.append(section_row("5. 시스템관리자 입력 정보"))
    rows.append(worksheet_row(161, "회생1차년도의 추정 소비자물가상승률", {"C": (assumptions.cpi_rate, "percent")}))
    rows.append(worksheet_row(162, "3년만기 국고채수익률", {"C": (assumptions.bond_yield_3y, "percent")}))
    rows.append(worksheet_row(163, "변제액 현가 시 적용 이자율", {"C": (assumptions.repayment_present_value_rate, "percent")}))
    rows.append(worksheet_row(164, "준비연도 기준일 ~ 말일 일수", {"C": dates["prep_days"]}))
    rows.append(worksheet_row(165, "준비연도 총일수", {"C": dates["total_days"]}))
    rows.append(worksheet_row(168, "조사기준일자", {"C": (dates["today"], "date")}, row_type="header"))
    rows.append(worksheet_row(169, "준비연도 말일자", {"C": (dates["end"], "date"), "D": dates["prep_days"]}, row_type="header"))
    rows.append(worksheet_row(170, "준비연도 1/1일", {"C": (dates["start"], "date"), "D": dates["total_days"]}, row_type="header"))

    rows.append(section_row("6. 청산 가정 시 청산배당액 배분"))
    rows.append(
        worksheet_row(
            185,
            "1.청산가치",
            {
                "D": liquidation_distribution["secured_collateral_value"],
                "E": liquidation_distribution["other_liquidation_value"],
                "G": assets["totals"]["liquidation"],
            },
        )
    )
    rows.append(
        worksheet_row(
            186,
            "2.청산관리비용",
            {
                "D": -liquidation_distribution["secured_collateral_value"] * assumptions.liquidation_cost_rate,
                "E": -liquidation_distribution["other_liquidation_value"] * assumptions.liquidation_cost_rate,
                "G": -assets["totals"]["liquidation"] * assumptions.liquidation_cost_rate,
            },
        )
    )
    rows.append(
        worksheet_row(
            187,
            "소계",
            {
                "D": liquidation_distribution["secured_pool"],
                "E": liquidation_distribution["unsecured_pool"],
                "G": liquidation_distribution["total_pool"],
            },
        )
    )
    liquidation_rows = [
        (189, "unpaid_wages"),
        (190, "retirement_benefit"),
        (191, "tax_arrears"),
        (192, "secured_debt"),
        (193, "unsecured_financial_debt"),
        (194, "other_unsecured_debt"),
        (195, "related_party_debt"),
    ]
    for row_number_value, key in liquidation_rows:
        row = liquidation_distribution["rows"].get(key, {})
        rows.append(
            worksheet_row(
                row_number_value,
                DEBT_FIELDS[key],
                {
                    "C": debts.get(key, 0.0),
                    "D": liquidation_distribution["secured_repay"].get(key, 0.0),
                    "E": liquidation_distribution["unsecured_repay"].get(key, 0.0),
                    "F": liquidation_distribution["residual_repay"].get(key, 0.0),
                    "G": row.get("repayment", 0.0),
                    "H": row.get("unpaid", 0.0),
                    "I": (row.get("rate", 0.0), "percent"),
                },
            )
        )
    rows.append(
        worksheet_row(
            196,
            "소계",
            {
                "C": sum(debts.get(key, 0.0) for _, key in liquidation_rows),
                "D": sum(liquidation_distribution["secured_repay"].get(key, 0.0) for _, key in liquidation_rows),
                "E": sum(liquidation_distribution["unsecured_repay"].get(key, 0.0) for _, key in liquidation_rows),
                "F": sum(liquidation_distribution["residual_repay"].get(key, 0.0) for _, key in liquidation_rows),
                "G": sum(liquidation_distribution["rows"].get(key, {}).get("repayment", 0.0) for _, key in liquidation_rows),
                "H": sum(liquidation_distribution["rows"].get(key, {}).get("unpaid", 0.0) for _, key in liquidation_rows),
            },
        )
    )

    rows.append(section_row("7. 계속기업가치 산정"))
    rows.append(worksheet_row(203, "구분", {column: (label, "text") for column, label in zip(PERIOD_COLUMNS, PERIOD_LABELS)}, row_type="header"))
    rows.append(worksheet_row(204, "Ⅰ.매출액", period_cells(operating_projection["sales"])))
    rows.append(worksheet_row(205, "Ⅱ.매출원가", period_cells(operating_projection["cost_of_sales"])))
    rows.append(worksheet_row(206, "Ⅲ.매출총이익", period_cells(operating_projection["gross_profit"])))
    rows.append(worksheet_row(207, "Ⅳ.판매비와관리비", period_cells(operating_projection["sgna"])))
    rows.append(worksheet_row(248, "Ⅴ.영업이익", period_cells(operating_projection["operating_profit"])))
    rows.append(worksheet_row(254, "영업활동 현금유입 합계", period_cells(operating_projection["operating_profit"], total=True)))
    rows.append(worksheet_row(255, "현가계수", {column: (value, "decimal") for column, value in zip(PERIOD_COLUMNS, enterprise_value["pv_factors"])}, row_type="header"))
    rows.append(worksheet_row(256, "현재가치", period_cells(pv_values, total=True)))
    rows.append(worksheet_row(258, "- 할인율", {"C": (enterprise_value["discount_rate"], "percent")}))
    rows.append(worksheet_row(261, "일수", {column: (value, "decimal") for column, value in zip(PERIOD_COLUMNS, enterprise_value["period_fractions"])}, row_type="header"))
    rows.append(worksheet_row(262, "현가계수", {column: (value, "decimal") for column, value in zip(PERIOD_COLUMNS, enterprise_value["pv_factors"])}, row_type="header"))
    rows.append(worksheet_row(266, "회생기간 이후 현금흐름", {"C": enterprise_value["terminal_cashflow"]}))
    rows.append(worksheet_row(267, "현가계수", {"C": (enterprise_value["pv_factors"][-1], "decimal")}, row_type="header"))
    rows.append(worksheet_row(268, "현재가치", {"C": enterprise_value["terminal_value"]}))
    rows.append(worksheet_row(272, "영업활동현금흐름의 현재가치", {"C": enterprise_value["pv_10_years"] + enterprise_value["terminal_value"]}))
    rows.append(worksheet_row(273, "회생기간(10년)의 현금흐름 현재가치", {"C": enterprise_value["pv_10_years"]}))
    rows.append(worksheet_row(274, "회생기간 이후 현금흐름 현재가치", {"C": enterprise_value["terminal_value"]}))
    rows.append(worksheet_row(275, "비업무용자산의 처분가치", {"C": enterprise_value["non_business_asset_value"]}))
    rows.append(worksheet_row(276, "계속기업가치", {"C": enterprise_value["going_concern_value"]}))

    rows.append(section_row("8. 계속기업 가정 시 현금흐름 및 채무변제액"))
    rows.append(worksheet_row(281, "기초현금", {"D": going_repayment["beginning_cash"]}))
    rows.append(worksheet_row(282, "영업활동현금흐름", {"D": going_repayment["operating_cashflow"]}))
    rows.append(worksheet_row(283, "투자활동현금흐름", {"D": going_repayment["asset_sale_cashflow"] + going_repayment["rent_deposit_outflow"]}))
    rows.append(worksheet_row(284, "담보등 매각", {"D": going_repayment["asset_sale_cashflow"]}))
    rows.append(worksheet_row(285, "임차보증금", {"D": going_repayment["rent_deposit_outflow"]}))
    rows.append(worksheet_row(286, "현금유입 합계", {"D": going_repayment["cash_inflow_total"]}))
    going_rows = [
        (287, "unpaid_wages"),
        (288, "retirement_benefit"),
        (289, "tax_arrears"),
        (290, "secured_debt"),
        (291, "secured_debt_interest"),
        (292, "unsecured_financial_debt"),
        (293, "other_unsecured_debt"),
        (294, "related_party_debt"),
    ]
    for row_number_value, key in going_rows:
        row = going_repayment["rows"].get(key, {})
        label = "담보채무_개시후이자" if key == "secured_debt_interest" else DEBT_FIELDS[key]
        rows.append(
            worksheet_row(
                row_number_value,
                label,
                {
                    "C": row.get("debt", 0.0),
                    "D": row.get("repayment", 0.0),
                    "E": (row.get("rate", 0.0), "percent") if row.get("debt", 0.0) else "",
                },
            )
        )
    rows.append(
        worksheet_row(
            295,
            "현금유출 합계",
            {
                "C": sum(going_repayment["rows"].get(key, {}).get("debt", 0.0) for _, key in going_rows),
                "D": sum(going_repayment["rows"].get(key, {}).get("repayment", 0.0) for _, key in going_rows),
            },
        )
    )
    rows.append(worksheet_row(296, "기말현금", {"D": going_repayment["ending_cash"]}))
    rows.append(worksheet_row(300, "이자계산 기간(일수)", {"C": dates["prep_days"] + dates["total_days"]}))
    rows.append(worksheet_row(301, "적용이자율", {"C": (assumptions.repayment_present_value_rate + assumptions.collateral_interest_spread, "percent")}))

    rows.append(section_row("9. 회생신청 가능성 진단 결과 - 고객 제시"))
    rows.append(
        worksheet_row(
            320,
            "청산가치 / 계속기업가치 / 차이",
            {
                "C": assets["totals"]["liquidation"],
                "D": enterprise_value["going_concern_value"],
                "E": enterprise_value["going_concern_value"] - assets["totals"]["liquidation"],
            },
        )
    )
    comparison_row_numbers = {
        "unpaid_wages": 326,
        "tax_arrears": 327,
        "secured_debt": 328,
        "unsecured_financial_debt": 329,
        "other_unsecured_debt": 330,
        "related_party_debt": 331,
    }
    for row in comparison_rows:
        rows.append(
            worksheet_row(
                comparison_row_numbers[row["field"]],
                row["label"],
                {
                    "C": row["debt"],
                    "D": row["liquidation_repayment"],
                    "E": (row["liquidation_rate"], "percent"),
                    "F": row["going_repayment"],
                    "G": (row["going_rate"], "percent"),
                    "H": row["repayment_difference"],
                    "I": (row["rate_difference"], "percent"),
                },
            )
        )
    rows.append(
        worksheet_row(
            332,
            "합계",
            {
                "C": sum(row["debt"] for row in comparison_rows),
                "D": sum(row["liquidation_repayment"] for row in comparison_rows),
                "F": sum(row["going_repayment"] for row in comparison_rows),
                "H": sum(row["repayment_difference"] for row in comparison_rows),
            },
        )
    )
    rows.append(
        worksheet_row(
            337,
            "계속기업가치와 청산가치 비교",
            {"C": ("Positive" if diagnosis["value_positive"] else "Negative", "text")},
            row_type="note",
        )
    )
    rows.append(
        worksheet_row(
            342,
            "청산 시 배당액과 계속기업 가정 시 채무변제액의 비교",
            {"C": ("Positive" if diagnosis["repayment_positive"] else "Negative", "text")},
            row_type="note",
        )
    )
    rows.append(
        worksheet_row(
            347,
            "계속기업 가정 시 회생채권 등의 변제율을 고려한 회생신청 가능성 결과",
            {
                "C": (
                    (
                        ""
                        if not diagnosis["consent_applicable"]
                        else "Positive" if diagnosis["consent_positive"] else "Negative"
                    ),
                    "text",
                )
            },
            row_type="note",
        )
    )
    rows.append(
        worksheet_row(
            352,
            "종합결론",
            {"C": ("Positive" if diagnosis["overall_positive"] else "Negative", "text")},
            row_type="note",
        )
    )

    return {"columns": WORKSHEET_COLUMNS, "rows": rows}


def calculate_case_result(case: dict[str, Any]) -> dict[str, Any]:
    assumptions = CalculationAssumptions()
    dates = default_preparation_dates()
    assets = calculate_asset_totals(case.get("financial_rows", []))
    debts = debt_amounts(case.get("debt_rows", []))
    collateral = collateral_amounts(
        case.get("collateral_rows", []),
        case.get("rent_rows", []),
    )
    income = income_assumptions(case.get("income_rows", []))
    monthly_rent = collateral["rent"].get("monthly_rent", 0.0)

    operating_projection = project_operating_profit(
        income,
        monthly_rent,
        dates,
        assumptions,
    )
    liquidation_distribution = calculate_liquidation_distribution(
        debts,
        assets,
        collateral,
        assumptions,
    )
    enterprise_value = calculate_enterprise_value(
        operating_projection,
        collateral,
        assumptions,
        dates,
    )
    going_repayment = calculate_going_concern_repayment(
        debts,
        liquidation_distribution,
        operating_projection,
        assets,
        collateral,
        dates,
        assumptions,
    )

    liquidation_value = assets["totals"]["liquidation"]
    going_concern_value = enterprise_value["going_concern_value"]
    total_debt = sum(debts.values())

    comparison_order = [
        "unpaid_wages",
        "tax_arrears",
        "secured_debt",
        "unsecured_financial_debt",
        "other_unsecured_debt",
        "related_party_debt",
    ]
    comparison_rows = []
    for key in comparison_order:
        liquidation_row = liquidation_distribution["rows"].get(key, {})
        going_row = going_repayment["rows"].get(key, {})
        debt = debts.get(key, 0.0)
        liquidation_repayment = liquidation_row.get("repayment", 0.0)
        going_repayment_value = going_row.get("repayment", 0.0)
        liquidation_rate = safe_div(liquidation_repayment, debt)
        going_rate = safe_div(going_repayment_value, debt)
        comparison_rows.append(
            {
                "field": key,
                "label": DEBT_FIELDS[key],
                "debt": debt,
                "liquidation_repayment": liquidation_repayment,
                "liquidation_rate": liquidation_rate,
                "going_repayment": going_repayment_value,
                "going_rate": going_rate,
                "repayment_difference": going_repayment_value - liquidation_repayment,
                "rate_difference": going_rate - liquidation_rate,
            }
        )

    diagnosis = build_diagnosis(
        liquidation_value,
        going_concern_value,
        comparison_rows,
    )
    worksheet_review = build_worksheet_review(
        case,
        assets,
        debts,
        collateral,
        income,
        operating_projection,
        liquidation_distribution,
        enterprise_value,
        going_repayment,
        comparison_rows,
        diagnosis,
        assumptions,
        dates,
    )

    return {
        "assumptions": assumptions,
        "dates": dates,
        "assets": assets,
        "debts": debts,
        "collateral": collateral,
        "income": income,
        "operating_projection": operating_projection,
        "liquidation_distribution": liquidation_distribution,
        "enterprise_value": enterprise_value,
        "going_repayment": going_repayment,
        "comparison_rows": comparison_rows,
        "diagnosis": diagnosis,
        "worksheet_review": worksheet_review,
        "visual": build_result_visual(liquidation_value, going_concern_value),
        "summary": {
            "liquidation_value": liquidation_value,
            "going_concern_value": going_concern_value,
            "value_difference": going_concern_value - liquidation_value,
            "asset_statement_total": assets["totals"]["statement"],
            "asset_audit_total": assets["totals"]["audit"],
            "asset_liquidation_total": assets["totals"]["liquidation"],
            "total_debt": total_debt,
            "operating_cashflow_total": sum(operating_projection["operating_profit"]),
            "pv_operating_cashflow": enterprise_value["pv_10_years"] + enterprise_value["terminal_value"],
            "non_business_asset_value": enterprise_value["non_business_asset_value"],
        },
        "display": {
            "number": display_number,
            "percent": display_percent,
        },
    }
