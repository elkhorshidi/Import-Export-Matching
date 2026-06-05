"""Pure calculation functions used by the Streamlit dashboard."""

from __future__ import annotations

from typing import Any, Dict, Iterable, Mapping

SCENARIO_LABELS = {
    "A": "A - مالیات فقط روی کارمزد",
    "B": "B - مالیات سود برآوردی اینتاکد",
    "C": "C - مالیات سود واقعی/دلخواه",
}


def scenario_code(value: Any) -> str:
    text = str(value or "A").strip()
    if text.startswith(("A", "B", "C")):
        return text[0]
    return "A"


def _amount(value: Any) -> float:
    try:
        return max(float(value or 0), 0.0)
    except (TypeError, ValueError):
        return 0.0


def pct(value: Any) -> float:
    return _amount(value) / 100


def calculate_import_costs(data: Mapping[str, Any]) -> Dict[str, float]:
    fx_rate = _amount(data.get("fx_rate"))
    goods = _amount(data.get("foreign_goods_cost")) * fx_rate
    freight = _amount(data.get("freight")) * fx_rate
    insurance = _amount(data.get("insurance")) * fx_rate
    foreign_total = goods + freight + insurance

    transfer_fee = foreign_total * pct(data.get("fx_transfer_fee_percent"))
    fx_spread = foreign_total * pct(data.get("fx_spread_percent"))
    customs_base_before_adjustment = foreign_total
    calculated_customs_taxable_base = customs_base_before_adjustment * (
        1 + pct(data.get("customs_valuation_adjustment_percent"))
    )
    customs_taxable_base_override = _amount(data.get("customs_taxable_base_override"))
    customs_taxable_base = (
        customs_taxable_base_override
        if customs_taxable_base_override > 0
        else calculated_customs_taxable_base
    )
    customs_duty = customs_taxable_base * pct(data.get("customs_duty_percent"))
    commercial_benefit = customs_taxable_base * pct(
        data.get("commercial_benefit_duty_percent")
    )
    other_import_duties = _amount(data.get("other_import_duties"))
    import_vat_base = (
        customs_taxable_base + customs_duty + commercial_benefit + other_import_duties
    )
    import_vat = (
        0
        if data.get("vat_exempt")
        else import_vat_base * pct(data.get("import_vat_rate"))
    )
    local_costs = sum(
        _amount(data.get(key))
        for key in (
            "clearance_cost",
            "warehousing_cost",
            "demurrage_cost",
            "port_costs",
            "miscellaneous_costs",
        )
    )
    landed_cost_excluding_recoverable_vat = (
        foreign_total
        + transfer_fee
        + fx_spread
        + customs_duty
        + commercial_benefit
        + other_import_duties
        + local_costs
    )
    total_landed_cost = landed_cost_excluding_recoverable_vat + import_vat
    return {
        "goods_cost_toman": goods,
        "freight_toman": freight,
        "insurance_toman": insurance,
        "foreign_total_toman": foreign_total,
        "fx_transfer_fee": transfer_fee,
        "fx_spread_cost": fx_spread,
        "calculated_customs_taxable_base": calculated_customs_taxable_base,
        "customs_taxable_base_override": customs_taxable_base_override,
        "customs_taxable_base": customs_taxable_base,
        "customs_duty": customs_duty,
        "commercial_benefit_duty": commercial_benefit,
        "other_import_duties": other_import_duties,
        "import_vat_base": import_vat_base,
        "import_vat": import_vat,
        "local_costs": local_costs,
        "landed_cost_excluding_recoverable_vat": landed_cost_excluding_recoverable_vat,
        "total_landed_cost": total_landed_cost,
    }


def calculate_vat(
    data: Mapping[str, Any], import_vat_credit: float
) -> Dict[str, Any]:
    domestic_sale_vat = (
        0
        if data.get("vat_exempt")
        else _amount(data.get("domestic_invoice_amount"))
        * pct(data.get("domestic_vat_rate"))
    )
    existing_credit = _amount(data.get("existing_export_vat_credit"))
    net_vat_payable = domestic_sale_vat - _amount(import_vat_credit) - existing_credit
    return {
        "vat_exempt": bool(data.get("vat_exempt")),
        "domestic_invoice_amount": _amount(data.get("domestic_invoice_amount")),
        "domestic_sale_vat": domestic_sale_vat,
        "import_vat_credit": _amount(import_vat_credit),
        "existing_export_vat_credit": existing_credit,
        "net_vat_payable": max(net_vat_payable, 0),
        "net_vat_refundable": abs(min(net_vat_payable, 0)),
        "net_vat_position": net_vat_payable,
        "domestic_invoice_vat_warning": (
            "Domestic sale VAT must be shown on the domestic invoice even when "
            "import/export VAT credits offset the payable amount."
            if not data.get("vat_exempt")
            else "The product is marked VAT-exempt; verify the exemption documentation."
        ),
    }


def calculate_income_tax(data: Mapping[str, Any]) -> Dict[str, float]:
    import_value_usd = _amount(data.get("import_value_usd"))
    gross_commission = (
        import_value_usd * _amount(data.get("commission_per_usd"))
        + _amount(data.get("fixed_commission"))
    )
    tax_rate = pct(data.get("tax_rate"))
    tax_a = gross_commission * tax_rate
    deemed_profit = _amount(data.get("domestic_invoice_amount")) * pct(
        data.get("estimated_profit_ratio")
    )
    tax_b = deemed_profit * tax_rate
    custom_profit = _amount(data.get("custom_profit_amount"))
    tax_c = custom_profit * tax_rate

    scenario = scenario_code(data.get("tax_scenario", "A"))
    selected_tax = {"A": tax_a, "B": tax_b, "C": tax_c}.get(str(scenario), tax_a)
    selected_tax_base = {
        "A": gross_commission,
        "B": deemed_profit,
        "C": custom_profit,
    }.get(scenario, gross_commission)
    net_card_owner_benefit = (
        gross_commission
        + _amount(data.get("export_obligation_benefit"))
        - selected_tax
        - _amount(data.get("admin_cost"))
        - _amount(data.get("risk_reserve"))
    )
    return {
        "gross_commission": gross_commission,
        "commission_tax_base": gross_commission,
        "deemed_profit": deemed_profit,
        "custom_profit": custom_profit,
        "tax_scenario_a": tax_a,
        "tax_scenario_b": tax_b,
        "tax_scenario_c": tax_c,
        "worst_case_income_tax": max(tax_a, tax_b, tax_c),
        "selected_tax_scenario": str(scenario),
        "selected_tax_scenario_name": SCENARIO_LABELS.get(scenario, SCENARIO_LABELS["A"]),
        "selected_tax_base": selected_tax_base,
        "selected_income_tax": selected_tax,
        "net_card_owner_benefit": net_card_owner_benefit,
    }


def calculate_risk(
    checked_items: Mapping[str, bool], weights: Iterable[Mapping[str, Any]]
) -> Dict[str, Any]:
    missing = []
    score = 0
    for item in weights:
        item_id = str(item["id"])
        if not checked_items.get(item_id, False):
            weight = int(item["weight"])
            score += weight
            missing.append({**item, "weight": weight})

    if score <= 20:
        level = "کم"
    elif score <= 45:
        level = "متوسط"
    elif score <= 70:
        level = "بالا"
    else:
        level = "غیرقابل قبول"
    checked_count = 0
    total_count = 0
    for item in weights:
        total_count += 1
        if checked_items.get(str(item["id"]), False):
            checked_count += 1
    open_top_risks = sorted(missing, key=lambda item: item["weight"], reverse=True)[:3]
    return {
        "risk_score": score,
        "risk_level": level,
        "missing_items": missing,
        "open_top_risks": open_top_risks,
        "checked_count": checked_count,
        "unchecked_count": total_count - checked_count,
    }


def calculate_collateral(
    data: Mapping[str, Any],
    customs_exposure: float,
    vat_exposure: float,
    income_tax_exposure: float,
    fx_exposure: float,
    export_obligation_exposure: float,
) -> Dict[str, float]:
    customs_reserve = _amount(customs_exposure) * pct(
        data.get("customs_risk_reserve_percent")
    )
    vat_reserve = _amount(vat_exposure) * pct(data.get("vat_risk_reserve_percent"))
    tax_reserve = _amount(income_tax_exposure) * pct(
        data.get("income_tax_risk_reserve_percent")
    )
    fx_reserve = _amount(fx_exposure) * pct(data.get("fx_risk_reserve_percent"))
    export_obligation_reserve = _amount(export_obligation_exposure) * pct(
        data.get("export_obligation_risk_reserve_percent")
    )
    legal_reserve = _amount(data.get("fixed_legal_admin_reserve"))
    minimum_collateral = (
        customs_reserve
        + vat_reserve
        + tax_reserve
        + fx_reserve
        + export_obligation_reserve
        + legal_reserve
    )
    return {
        "customs_risk_reserve": customs_reserve,
        "vat_risk_reserve": vat_reserve,
        "selected_income_tax_for_collateral": _amount(data.get("selected_income_tax")),
        "worst_case_income_tax": _amount(income_tax_exposure),
        "income_tax_risk_reserve_percent": _amount(
            data.get("income_tax_risk_reserve_percent")
        ),
        "income_tax_risk_reserve": tax_reserve,
        "fx_risk_reserve": fx_reserve,
        "export_obligation_risk_reserve": export_obligation_reserve,
        "fixed_legal_admin_reserve": legal_reserve,
        "minimum_collateral": minimum_collateral,
    }


def final_decision(
    risk_score: float, net_card_owner_benefit: float, minimum_collateral: float = 0
) -> str:
    if risk_score > 70 or net_card_owner_benefit < 0:
        return "رد معامله"
    if risk_score >= 21:
        return "نیازمند تضمین و کنترل تکمیلی"
    return "قابل قبول"


def final_recommendation(risk_score: float, net_card_owner_benefit: float) -> str:
    if risk_score > 70 or net_card_owner_benefit < 0:
        return "رد معامله؛ ریسک یا زیان خالص برای صاحب کارت بیش از حد قابل قبول است."
    if 46 <= risk_score <= 70:
        return "فقط در صورت دریافت وثیقه کافی، تأیید مالیاتی/گمرکی و کنترل کامل اسناد قابل بررسی است."
    if 21 <= risk_score <= 45:
        return "قابل بررسی با اخذ تضمین، قرارداد دقیق، تعیین تکلیف VAT، اینتاکد و فاکتور داخلی."
    return "از نظر مدیریتی قابل قبول است، مشروط به تأیید نهایی حسابدار، مشاور مالیاتی و کنترل اسناد."


def calculate_commission_analysis(data: Mapping[str, Any], selected_income_tax: float) -> Dict[str, float | str]:
    import_value_usd = _amount(data.get("import_value_foreign")) * _amount(
        data.get("currency_to_usd", 1)
    )
    current_commission_per_usd = _amount(data.get("commission_per_usd"))
    minimum_commission_per_usd = (
        (_amount(selected_income_tax) + _amount(data.get("admin_cost")) + _amount(data.get("risk_reserve")))
        / import_value_usd
        if import_value_usd
        else 0
    )
    commission_gap = current_commission_per_usd - minimum_commission_per_usd
    warning = (
        "کارمزد فعلی برای پوشش هزینه‌ها و ریسک‌های مستقیم صاحب کارت کافی نیست."
        if commission_gap < 0
        else "کارمزد فعلی هزینه‌ها و ریسک‌های مستقیم صاحب کارت را پوشش می‌دهد."
    )
    return {
        "minimum_commission_per_usd": minimum_commission_per_usd,
        "commission_gap_per_usd": commission_gap,
        "commission_floor_warning": warning,
    }
