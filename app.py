"""Persian RTL Streamlit dashboard for exporter-card import evaluation."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

from utils.calculations import (
    SCENARIO_LABELS,
    calculate_collateral,
    calculate_commission_analysis,
    calculate_import_costs,
    calculate_income_tax,
    calculate_risk,
    calculate_vat,
    final_decision,
    final_recommendation,
)
from utils.exporter import (
    export_all_to_excel,
    export_checklist_to_excel,
    export_summary_to_excel,
)

BASE_DIR = Path(__file__).resolve().parent

st.set_page_config(page_title="exporter_card_import_calculator", layout="wide")
st.markdown(
    """
    <style>
    :root {
        --bg: #f5f7fb;
        --card: #ffffff;
        --line: #d9e2ec;
        --text: #1f2937;
        --muted: #64748b;
        --blue: #2563eb;
        --green: #16a34a;
        --yellow: #ca8a04;
        --orange: #ea580c;
        --red: #dc2626;
        --slate: #475569;
    }
    html, body, [class*="css"], .stApp {
        direction: rtl; text-align: right; background: var(--bg); color: var(--text);
    }
    .main .block-container {
        max-width: 1380px; padding-top: 1.4rem; padding-bottom: 3rem;
    }
    [data-testid="stSidebar"] {
        direction: rtl; text-align: right; background: #0f172a; color: #f8fafc;
    }
    [data-testid="stSidebar"] * { text-align: right; }
    [data-testid="stSidebar"] label, [data-testid="stSidebar"] p { color: #e2e8f0; }
    [data-testid="stSidebar"] [role="radiogroup"] label {
        border-radius: 12px; padding: 6px 8px; margin-bottom: 3px;
    }
    .app-header {
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 72%, #2563eb 100%);
        color: #ffffff; border-radius: 22px; padding: 28px 32px;
        box-shadow: 0 18px 45px rgba(15, 23, 42, .18); margin-bottom: 22px;
    }
    .app-header h1 { margin: 0 0 8px 0; font-size: 2rem; line-height: 1.5; }
    .app-header .subtitle { font-size: 1.05rem; color: #dbeafe; margin-bottom: 14px; }
    .app-header .disclaimer {
        display: inline-block; background: rgba(255, 255, 255, .12);
        border: 1px solid rgba(255, 255, 255, .20); border-radius: 999px;
        padding: 7px 13px; color: #f8fafc; font-size: .9rem;
    }
    .page-head {
        margin: 10px 0 16px; padding: 8px 2px;
    }
    .page-head h2 { margin: 0 0 4px 0; font-size: 1.55rem; color: #0f172a; }
    .page-head p { margin: 0; color: var(--muted); line-height: 1.8; }
    div[data-testid="stVerticalBlockBorderWrapper"] {
        border-color: #e2e8f0 !important; border-radius: 18px !important;
        background: rgba(255, 255, 255, .86) !important;
        box-shadow: 0 8px 28px rgba(15, 23, 42, .05);
    }
    [data-testid="stMetric"] {
        background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 12px;
        padding: 14px; min-height: 108px;
    }
    .kpi-card {
        background: var(--card); border: 1px solid #e2e8f0; border-right: 5px solid var(--slate);
        border-radius: 18px; padding: 16px 17px; min-height: 126px;
        box-shadow: 0 10px 26px rgba(15, 23, 42, .06); margin-bottom: 12px;
    }
    .kpi-title { color: #64748b; font-size: .9rem; margin-bottom: 8px; }
    .kpi-value { color: #0f172a; font-size: 1.35rem; font-weight: 800; line-height: 1.55; }
    .kpi-subtitle { color: #64748b; font-size: .84rem; margin-top: 7px; line-height: 1.7; }
    .kpi-neutral { border-right-color: #64748b; }
    .kpi-blue { border-right-color: var(--blue); }
    .kpi-green { border-right-color: var(--green); }
    .kpi-yellow { border-right-color: var(--yellow); }
    .kpi-orange { border-right-color: var(--orange); }
    .kpi-red { border-right-color: var(--red); }
    .decision {
        padding: 24px; border-radius: 20px; font-size: 1.45rem; font-weight: 800;
        text-align: center; color: #ffffff; margin: 8px 0 18px;
        box-shadow: 0 16px 36px rgba(15, 23, 42, .12);
    }
    .decision-green { background: linear-gradient(135deg, #15803d, #22c55e); }
    .decision-yellow { background: linear-gradient(135deg, #a16207, #eab308); }
    .decision-orange { background: linear-gradient(135deg, #c2410c, #f97316); }
    .decision-red { background: linear-gradient(135deg, #991b1b, #ef4444); }
    .soft-card {
        background: #ffffff; border: 1px solid #e2e8f0; border-radius: 18px;
        padding: 16px 18px; margin: 12px 0;
        box-shadow: 0 8px 26px rgba(15, 23, 42, .05);
    }
    .soft-card h3 { margin: 0 0 8px 0; color: #0f172a; font-size: 1.05rem; }
    .soft-card p { margin: 0; color: #475569; line-height: 1.9; }
    .guide-box {
        background: #eff6ff; border: 1px solid #bfdbfe; border-right: 5px solid var(--blue);
        border-radius: 14px; padding: 12px 14px; color: #1e3a8a; line-height: 1.9; margin: 10px 0;
    }
    .warning-panel {
        background: #fff7ed; border: 1px solid #fed7aa; border-right: 5px solid var(--orange);
        border-radius: 16px; padding: 14px 16px; margin: 12px 0; color: #7c2d12;
    }
    .risk-box {
        background: #fff7ed; border: 1px solid #fed7aa; border-right: 5px solid var(--orange);
        border-radius: 14px; padding: 12px 14px; margin-bottom: 10px; line-height: 1.8;
    }
    .check-card {
        background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px;
        padding: 13px 14px; margin-bottom: 10px; min-height: 132px;
        box-shadow: 0 8px 22px rgba(15, 23, 42, .045);
    }
    .check-card-open-high {
        background: #fff7ed; border-color: #fb923c; border-right: 5px solid var(--orange);
    }
    .check-card-open-critical {
        background: #fef2f2; border-color: #f87171; border-right: 5px solid var(--red);
    }
    .check-weight {
        display: inline-block; background: #e0f2fe; color: #075985;
        padding: 3px 10px; border-radius: 999px; font-size: .84rem;
        font-weight: 700; margin-bottom: 6px;
    }
    .check-help { color: #4a5568; font-size: .9rem; line-height: 1.7; }
    .download-card {
        background: #ffffff; border: 1px solid #e2e8f0; border-radius: 18px;
        padding: 16px; min-height: 190px; box-shadow: 0 8px 24px rgba(15, 23, 42, .05);
    }
    .download-card h3 { margin: 0 0 8px 0; color: #0f172a; font-size: 1.08rem; }
    .download-card p { color: #64748b; line-height: 1.8; min-height: 64px; }
    .stButton button, .stDownloadButton button {
        border-radius: 12px !important; font-weight: 700 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

with (BASE_DIR / "data/risk_weights.json").open(encoding="utf-8") as file:
    RISK_WEIGHTS = json.load(file)

DEFAULTS: dict[str, Any] = {
    "import_value_foreign": 100_000.0,
    "currency": "USD",
    "currency_to_usd": 1.0,
    "fx_rate": 65_000.0,
    "hs_code": "85176290",
    "product_description": "تجهیزات شبکه و مخابرات",
    "country_of_origin": "چین",
    "domestic_invoice_amount": 8_500_000_000.0,
    "commission_per_usd": 2_500.0,
    "export_obligation_benefit": 100_000_000.0,
    "vat_exempt": False,
    "special_permits": True,
    "expected_export_obligation_settlement": 4_000_000_000.0,
    "buyer_pays_vat_separately": True,
    "foreign_goods_cost": 100_000.0,
    "freight": 3_000.0,
    "insurance": 500.0,
    "fx_transfer_fee_percent": 1.0,
    "fx_spread_percent": 2.0,
    "customs_valuation_adjustment_percent": 3.0,
    "customs_duty_percent": 4.0,
    "commercial_benefit_duty_percent": 5.0,
    "other_import_duties": 15_000_000.0,
    "clearance_cost": 80_000_000.0,
    "warehousing_cost": 20_000_000.0,
    "demurrage_cost": 50_000_000.0,
    "port_costs": 30_000_000.0,
    "miscellaneous_costs": 0.0,
    "import_vat_rate": 10.0,
    "domestic_vat_rate": 10.0,
    "existing_export_vat_credit": 200_000_000.0,
    "inta_code": "4690010",
    "activity_description": "واردات و فروش داخلی تجهیزات شبکه",
    "estimated_profit_ratio": 10.0,
    "inta_profile_aligned": False,
    "tax_advisor_confirmed": True,
    "tax_rate": 25.0,
    "tax_scenario": SCENARIO_LABELS["B"],
    "admin_cost": 30_000_000.0,
    "risk_reserve": 50_000_000.0,
    "custom_profit_amount": 400_000_000.0,
    "vat_risk_reserve_percent": 100.0,
    "income_tax_risk_reserve_percent": 100.0,
    "customs_risk_reserve_percent": 20.0,
    "fx_risk_reserve_percent": 5.0,
    "export_obligation_risk_reserve_percent": 20.0,
    "fixed_legal_admin_reserve": 100_000_000.0,
    "fixed_commission": 0.0,
}

DEFAULT_RISK_CHECKED = {
    "product_permits_verified": False,
    "buyer_confirms_einvoice": False,
    "inta_checked": False,
    "vat_credit_rejection_covered": False,
    "auditor_approval": False,
}


def init_state() -> None:
    for key, value in DEFAULTS.items():
        st.session_state.setdefault(key, value)
        # Touch durable calculation keys so Streamlit does not clean them up on pages
        # where the corresponding widget is not rendered.
        st.session_state[key] = st.session_state[key]
    for item in RISK_WEIGHTS:
        key = f"risk_{item['id']}"
        st.session_state.setdefault(key, DEFAULT_RISK_CHECKED.get(item["id"], True))
        st.session_state[key] = st.session_state[key]


def fmt_foreign(value: float, unit: str) -> str:
    return f"{float(value):,.0f} {unit}"


def fmt_percent_readable(value: float) -> str:
    return f"{float(value):g}%"


def readable_preview(value: float, unit: str) -> None:
    st.caption(f"مقدار خوانا: {fmt_foreign(value, unit)}")


def money_input(
    label: str,
    key: str,
    help_text: str | None = None,
    unit: str = "تومان",
    step: float = 1_000_000.0,
) -> float:
    value = st.number_input(
        label,
        min_value=0.0,
        step=step,
        format="%.0f",
        key=key,
        help=help_text,
    )
    readable_preview(value, unit)
    return float(value)


def percent_input(label: str, key: str, help_text: str | None = None) -> float:
    value = st.number_input(
        label,
        min_value=0.0,
        max_value=100.0,
        step=0.1,
        format="%.2f",
        key=key,
        help=help_text,
    )
    st.caption(f"مقدار خوانا: {fmt_percent_readable(value)}")
    return float(value)


def yes_no(label: str, key: str, help_text: str | None = None) -> bool:
    value = st.radio(
        label,
        [True, False],
        format_func=lambda item: "بله" if item else "خیر",
        horizontal=True,
        key=key,
        help=help_text,
    )
    return bool(value)


def fmt_money(value: float) -> str:
    return f"{float(value):,.0f} تومان"


def fmt_number(value: float) -> str:
    return f"{float(value):,.0f}"


def fmt_percent(value: float) -> str:
    return f"{float(value):.2f}%"


def status_color_for_risk(risk_score: float) -> str:
    if risk_score <= 20:
        return "green"
    if risk_score <= 45:
        return "yellow"
    if risk_score <= 70:
        return "orange"
    return "red"


def status_color_for_decision(decision: str) -> str:
    if "رد" in decision:
        return "red"
    if "تضمین" in decision or "تکمیلی" in decision:
        return "yellow"
    return "green"


def page_title(title: str, description: str) -> None:
    st.markdown(
        f"""
        <div class="page-head">
            <h2>{title}</h2>
            <p>{description}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def guide_box(text: str) -> None:
    st.markdown(f"<div class='guide-box'>{text}</div>", unsafe_allow_html=True)


def kpi_card(
    title: str,
    value: float | str,
    subtitle: str = "",
    color: str = "neutral",
    money: bool = True,
) -> None:
    display_value = fmt_money(value) if money and isinstance(value, (int, float)) else value
    st.markdown(
        f"""
        <div class="kpi-card kpi-{color}">
            <div class="kpi-title">{title}</div>
            <div class="kpi-value">{display_value}</div>
            <div class="kpi-subtitle">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def metric(label: str, value: float | str, money: bool = True) -> None:
    st.metric(label, fmt_money(value) if money and isinstance(value, (int, float)) else value)


def money_table(rows: list[tuple[str, Any]], value_title: str = "مقدار") -> pd.DataFrame:
    return pd.DataFrame(
        [(label, fmt_money(value) if isinstance(value, (int, float)) else value) for label, value in rows],
        columns=["شاخص", value_title],
    )


def collect_inputs() -> dict[str, Any]:
    data = {key: st.session_state.get(key, value) for key, value in DEFAULTS.items()}
    data["import_value_usd"] = float(data["import_value_foreign"]) * float(data["currency_to_usd"])
    return data


def checked_items() -> dict[str, bool]:
    return {
        item["id"]: bool(st.session_state.get(f"risk_{item['id']}", False))
        for item in RISK_WEIGHTS
    }


def compute() -> dict[str, Any]:
    data = collect_inputs()
    import_result = calculate_import_costs(data)
    vat_result = calculate_vat(data, import_result["import_vat"])
    tax_result = calculate_income_tax(data)
    data["selected_income_tax"] = tax_result["selected_income_tax"]
    risk_result = calculate_risk(checked_items(), RISK_WEIGHTS)
    collateral_result = calculate_collateral(
        data,
        customs_exposure=import_result["import_vat_base"],
        vat_exposure=max(import_result["import_vat"], vat_result["domestic_sale_vat"]),
        income_tax_exposure=tax_result["worst_case_income_tax"],
        fx_exposure=import_result["foreign_total_toman"],
        export_obligation_exposure=float(data["expected_export_obligation_settlement"]),
    )
    commission = calculate_commission_analysis(data, tax_result["selected_income_tax"])
    total_operator_cost = (
        import_result["total_landed_cost"]
        + tax_result["gross_commission"]
        + vat_result["net_vat_payable"]
        + float(data["admin_cost"])
    )
    decision = final_decision(risk_result["risk_score"], tax_result["net_card_owner_benefit"])
    recommendation = final_recommendation(
        risk_result["risk_score"], tax_result["net_card_owner_benefit"]
    )
    commission_percent = (
        tax_result["gross_commission"] / (float(data["import_value_foreign"]) * float(data["fx_rate"]))
        if float(data["import_value_foreign"]) and float(data["fx_rate"])
        else 0
    )
    net_benefit_per_usd = (
        tax_result["net_card_owner_benefit"] / float(data["import_value_usd"])
        if float(data["import_value_usd"])
        else 0
    )
    open_risks = [
        item["label_fa"] for item in risk_result["open_top_risks"]
    ] or ["همه کنترل‌های اصلی انجام شده‌اند."]
    summary = {
        "import_value_total": float(data["import_value_foreign"]) * float(data["fx_rate"]),
        "total_landed_cost": import_result["total_landed_cost"],
        "gross_commission": tax_result["gross_commission"],
        "commission_percent": commission_percent,
        "selected_tax_scenario_name": tax_result["selected_tax_scenario_name"],
        "selected_income_tax": tax_result["selected_income_tax"],
        "import_vat": import_result["import_vat"],
        "domestic_sale_vat": vat_result["domestic_sale_vat"],
        "net_vat_position": vat_result["net_vat_position"],
        "net_card_owner_benefit": tax_result["net_card_owner_benefit"],
        "net_benefit_per_usd": net_benefit_per_usd,
        "total_operator_cost": total_operator_cost,
        "minimum_collateral": collateral_result["minimum_collateral"],
        "risk_score": risk_result["risk_score"],
        "risk_level": risk_result["risk_level"],
        "checked_count": risk_result["checked_count"],
        "unchecked_count": risk_result["unchecked_count"],
        "open_top_risks": open_risks,
        "minimum_commission_per_usd": commission["minimum_commission_per_usd"],
        "commission_gap_per_usd": commission["commission_gap_per_usd"],
        "commission_floor_warning": commission["commission_floor_warning"],
        "vat_buyer_payment_message": (
            "VAT فروش داخلی به‌صورت جداگانه از خریدار دریافت می‌شود."
            if data["buyer_pays_vat_separately"]
            else "هشدار: اگر VAT فروش داخلی جداگانه از خریدار دریافت نشود، ممکن است این مبلغ از سود صاحب کارت یا مجری کسر شود."
        ),
        "final_decision": decision,
        "final_recommendation": recommendation,
    }
    return {
        "inputs": data,
        "import": import_result,
        "vat": vat_result,
        "tax": tax_result,
        "risk": risk_result,
        "collateral": collateral_result,
        "commission": commission,
        "summary": summary,
        "outputs": {
            **import_result,
            **vat_result,
            **tax_result,
            **risk_result,
            **collateral_result,
            **commission,
            **summary,
        },
    }


def checklist_export_rows() -> list[dict[str, Any]]:
    checked = checked_items()
    return [{**item, "checked": checked[item["id"]]} for item in RISK_WEIGHTS]


init_state()

st.markdown(
    """
    <div class="app-header">
        <h1>داشبورد ارزیابی واردات روی کارت صادرکننده</h1>
        <div class="subtitle">محاسبه هزینه، ارزش افزوده، مالیات، ریسک، وثیقه و تصمیم مدیریتی</div>
        <div class="disclaimer">این ابزار صرفاً برای برآورد مدیریتی است و جایگزین نظر مشاور مالیاتی، گمرکی، حقوقی یا حسابداری نیست.</div>
    </div>
    """,
    unsafe_allow_html=True,
)

pages = [
    "ورودی‌های معامله",
    "هزینه‌های گمرکی و واردات",
    "محاسبات ارزش افزوده",
    "سناریوهای مالیات عملکرد",
    "چک‌لیست ریسک",
    "محاسبه وثیقه پیشنهادی",
    "خلاصه تصمیم مدیریتی",
    "خروجی Excel",
]
PAGE_ICONS = {
    "ورودی‌های معامله": "🧾",
    "هزینه‌های گمرکی و واردات": "🚢",
    "محاسبات ارزش افزوده": "🧮",
    "سناریوهای مالیات عملکرد": "📊",
    "چک‌لیست ریسک": "⚠️",
    "محاسبه وثیقه پیشنهادی": "🛡️",
    "خلاصه تصمیم مدیریتی": "📌",
    "خروجی Excel": "📤",
}
st.sidebar.markdown("### مسیر ارزیابی معامله")
page = st.sidebar.radio(
    "بخش‌های داشبورد",
    pages,
    key="current_page",
    format_func=lambda item: f"{PAGE_ICONS[item]} {item}",
)

if page == "ورودی‌های معامله":
    page_title(page, "مشخصات پایه معامله، نرخ ارز، فاکتور داخلی و وضعیت VAT را وارد کنید.")
    with st.container(border=True):
        st.subheader("اطلاعات پایه معامله")
        c1, c2, c3 = st.columns(3)
        with c1:
            money_input(
                "ارزش واردات ارزی",
                "import_value_foreign",
                unit=str(st.session_state.get("currency", "USD")),
                step=1_000.0,
            )
            st.selectbox("نوع ارز", ["USD", "EUR", "AED", "CNY", "TRY", "سایر"], key="currency")
            money_input("نرخ تبدیل", "fx_rate", step=1_000.0)
            currency_to_usd = st.number_input("ضریب تبدیل ارز به USD", min_value=0.0, step=0.01, key="currency_to_usd")
            readable_preview(currency_to_usd, "USD")
        with c2:
            st.text_input("HS Code", key="hs_code")
            st.text_input("شرح کالا", key="product_description")
            st.text_input("کشور مبدأ", key="country_of_origin")
            money_input(
                "مبلغ فاکتور فروش داخلی قبل از VAT",
                "domestic_invoice_amount",
                help_text="مبنای محاسبه VAT فروش داخلی و سناریوی مالیات مبتنی بر اینتاکد است.",
            )
        with c3:
            money_input("کارمزد هر دلار برای صاحب کارت", "commission_per_usd", step=100.0)
            money_input("میزان منفعت رفع تعهد ارزی", "export_obligation_benefit")
            yes_no("آیا کالا معاف از VAT است؟", "vat_exempt")
            yes_no("آیا مجوز خاص نیاز دارد؟", "special_permits")
            money_input("مبلغ مورد انتظار رفع تعهد", "expected_export_obligation_settlement")
            yes_no(
                "آیا VAT فروش داخلی جداگانه توسط خریدار پرداخت می‌شود؟",
                "buyer_pays_vat_separately",
                "اگر جداگانه دریافت نشود، ممکن است فشار نقدینگی یا کاهش سود ایجاد کند.",
            )
        guide_box("کنترل سامانه جامع تجارت باید نزد صاحب کارت باقی بماند و دسترسی‌ها، ثبت سفارش و اسناد قابل ردیابی باشند.")
    result = compute()
    cols = st.columns(3)
    with cols[0]:
        kpi_card("ارزش ریالی واردات", result["summary"]["import_value_total"], "ارزش ارزی × نرخ تبدیل", "blue")
    with cols[1]:
        kpi_card("فاکتور فروش داخلی", st.session_state.domestic_invoice_amount, "قبل از VAT", "neutral")
    with cols[2]:
        kpi_card("کارمزد ناخالص", result["tax"]["gross_commission"], "برای صاحب کارت", "green")

elif page == "هزینه‌های گمرکی و واردات":
    page_title(page, "هزینه‌های ارزی، عوارض، هزینه‌های محلی و بهای تمام‌شده واردات را بسنجید.")
    with st.container(border=True):
        st.subheader("ورودی‌های هزینه واردات")
        c1, c2, c3 = st.columns(3)
        with c1:
            foreign_unit = str(st.session_state.get("currency", "USD"))
            money_input("قیمت کالا (ارزی)", "foreign_goods_cost", unit=foreign_unit, step=1_000.0)
            money_input("حمل (ارزی)", "freight", unit=foreign_unit, step=100.0)
            money_input("بیمه (ارزی)", "insurance", unit=foreign_unit, step=100.0)
            percent_input("کارمزد حواله (%)", "fx_transfer_fee_percent")
            percent_input("اسپرد تبدیل ارز (%)", "fx_spread_percent")
        with c2:
            percent_input("تعدیل ارزش گمرکی (%)", "customs_valuation_adjustment_percent")
            percent_input("حقوق ورودی (%)", "customs_duty_percent")
            percent_input("سود بازرگانی (%)", "commercial_benefit_duty_percent")
            money_input("سایر عوارض", "other_import_duties")
        with c3:
            money_input("هزینه ترخیص", "clearance_cost")
            money_input("انبارداری", "warehousing_cost")
            money_input("دموراژ", "demurrage_cost")
            money_input("هزینه‌های بندری", "port_costs")
            money_input("سایر هزینه‌ها", "miscellaneous_costs")
    result = compute()
    kcols = st.columns(4)
    with kcols[0]:
        kpi_card("پایه ارزش گمرکی", result["import"]["customs_taxable_base"], "پس از تعدیل ارزش", "blue")
    with kcols[1]:
        kpi_card("حقوق و سود بازرگانی", result["import"]["customs_duty"] + result["import"]["commercial_benefit_duty"], "جمع حقوق ورودی و سود", "neutral")
    with kcols[2]:
        kpi_card("VAT واردات", result["import"]["import_vat"], "بر مبنای پایه مشمول", "yellow")
    with kcols[3]:
        kpi_card("بهای تمام‌شده با VAT", result["import"]["total_landed_cost"], "جمع کل هزینه واردات", "green")
    rows = [
        ("پایه ارزش گمرکی", result["import"]["customs_taxable_base"]),
        ("حقوق ورودی", result["import"]["customs_duty"]),
        ("سود بازرگانی", result["import"]["commercial_benefit_duty"]),
        ("بهای تمام‌شده قبل از VAT", result["import"]["landed_cost_excluding_recoverable_vat"]),
        ("بهای تمام‌شده با VAT", result["import"]["total_landed_cost"]),
    ]
    st.dataframe(money_table(rows, "مبلغ"), width="stretch", hide_index=True)

elif page == "محاسبات ارزش افزوده":
    page_title(page, "اثر VAT واردات، فروش داخلی، اعتبارهای قبلی و پرداخت جداگانه خریدار را بررسی کنید.")
    st.warning("حتی در صورت تهاتر، اگر کالا معاف نباشد، VAT فروش داخلی باید در صورتحساب داخلی درج شود.")
    with st.container(border=True):
        st.subheader("ورودی‌های VAT و فاکتور داخلی")
        c1, c2 = st.columns(2)
        with c1:
            percent_input("نرخ VAT واردات (%)", "import_vat_rate")
            money_input("مبلغ فاکتور فروش داخلی", "domestic_invoice_amount")
            percent_input("نرخ VAT فروش داخلی (%)", "domestic_vat_rate")
            money_input("طلب VAT صادراتی قبلی", "existing_export_vat_credit")
        with c2:
            yes_no("آیا کالا معاف از VAT است؟", "vat_exempt")
            yes_no(
                "آیا VAT فروش داخلی جداگانه توسط خریدار پرداخت می‌شود؟",
                "buyer_pays_vat_separately",
                "برای کالاهای مشمول، VAT فروش باید در صورتحساب درج شود؛ تهاتر فقط وضعیت پرداخت/اعتبار را تغییر می‌دهد.",
            )
            guide_box("VAT فروش داخلی یک موضوع صورتحساب و انطباق مالیاتی است؛ حتی اگر اعتبار VAT واردات یا صادراتی آن را پوشش دهد، نمایش آن در فاکتور داخلی باید کنترل شود.")
    result = compute()
    if not result["inputs"]["buyer_pays_vat_separately"]:
        st.error(result["summary"]["vat_buyer_payment_message"])
    kcols = st.columns(4)
    with kcols[0]:
        kpi_card("VAT واردات", result["import"]["import_vat"], "اعتبار احتمالی VAT", "yellow")
    with kcols[1]:
        kpi_card("VAT فروش داخلی", result["vat"]["domestic_sale_vat"], "باید در فاکتور بیاید", "yellow")
    with kcols[2]:
        color = "red" if result["vat"]["net_vat_position"] > 0 else "green"
        kpi_card("خالص VAT", result["vat"]["net_vat_position"], "مثبت: پرداخت | منفی: تهاتر/اعتبار", color)
    with kcols[3]:
        kpi_card("طلب VAT صادراتی قبلی", result["vat"]["existing_export_vat_credit"], "اعتبار قابل تهاتر", "blue")
    rows = [
        ("پایه مشمول VAT واردات", result["import"]["import_vat_base"]),
        ("مبلغ VAT واردات", result["import"]["import_vat"]),
        ("مبلغ VAT فروش داخلی", result["vat"]["domestic_sale_vat"]),
        ("طلب VAT صادراتی قبلی", result["vat"]["existing_export_vat_credit"]),
        ("خالص VAT قابل پرداخت یا قابل تهاتر", result["vat"]["net_vat_position"]),
    ]
    st.dataframe(money_table(rows, "مبلغ"), width="stretch", hide_index=True)

elif page == "سناریوهای مالیات عملکرد":
    page_title(page, "سه نگاه مالیاتی را کنار هم ببینید و سناریوی فعال تصمیم مدیریتی را انتخاب کنید.")
    with st.container(border=True):
        st.subheader("اینتاکد، سناریو و هزینه‌های مستقیم")
        c1, c2, c3 = st.columns(3)
        with c1:
            st.text_input("اینتاکد", key="inta_code", help="کد فعالیت مبنای ضریب سود برآوردی و دفاع مالیاتی است.")
            st.text_input("شرح فعالیت", key="activity_description")
            percent_input("ضریب سود فعالیت (%)", "estimated_profit_ratio", "ضریب سود باید با اینتاکد و واقعیت فعالیت همخوان باشد.")
            yes_no("آیا فعالیت با پرونده مالیاتی صاحب کارت همخوان است؟", "inta_profile_aligned")
        with c2:
            yes_no("آیا مشاور مالیاتی تأیید کرده؟", "tax_advisor_confirmed")
            percent_input("نرخ مالیات عملکرد (%)", "tax_rate")
            st.selectbox(
                "سناریوی منتخب مالیات",
                list(SCENARIO_LABELS.values()),
                key="tax_scenario",
                help="این انتخاب فقط سناریوی فعال برای سود خالص صاحب کارت است؛ وثیقه مالیاتی همچنان بر اساس بدترین سناریو سنجیده می‌شود.",
            )
            money_input("سود واقعی/دلخواه سناریوی C", "custom_profit_amount")
        with c3:
            money_input("هزینه اداری/حسابداری", "admin_cost")
            money_input("ذخیره مستقیم ریسک صاحب کارت", "risk_reserve")
            guide_box("مدل کمیسیون‌محور، فروش کامل کالا و مدل سود واقعی باید با قرارداد، اسناد حسابداری و اینتاکد قابل دفاع باشند.")
    result = compute()
    tax = result["tax"]
    kcols = st.columns(4)
    with kcols[0]:
        kpi_card("کارمزد ناخالص", tax["gross_commission"], "مبنای سناریوی A", "blue")
    with kcols[1]:
        kpi_card("مالیات سناریوی منتخب", tax["selected_income_tax"], tax["selected_tax_scenario_name"], "yellow")
    with kcols[2]:
        kpi_card("بدترین مالیات بین سناریوها", tax["worst_case_income_tax"], "برای محافظت وثیقه", "orange")
    with kcols[3]:
        color = "green" if tax["net_card_owner_benefit"] > 0 else "red"
        kpi_card("سود خالص صاحب کارت", tax["net_card_owner_benefit"], "پس از مالیات و ذخایر", color)
    scenario_rows = [
        ["A - مالیات فقط روی کارمزد", tax["commission_tax_base"], st.session_state.tax_rate, tax["tax_scenario_a"], tax["gross_commission"] + st.session_state.export_obligation_benefit - tax["tax_scenario_a"] - st.session_state.admin_cost - st.session_state.risk_reserve, "ریسک پایین‌تر در صورت قرارداد حق‌العمل‌کاری و مستندات کامل"],
        ["B - مالیات سود برآوردی اینتاکد", tax["deemed_profit"], st.session_state.tax_rate, tax["tax_scenario_b"], tax["gross_commission"] + st.session_state.export_obligation_benefit - tax["tax_scenario_b"] - st.session_state.admin_cost - st.session_state.risk_reserve, "ریسک متوسط؛ وابسته به پذیرش اینتاکد و مدل حسابداری"],
        ["C - مالیات سود واقعی/دلخواه", tax["custom_profit"], st.session_state.tax_rate, tax["tax_scenario_c"], tax["gross_commission"] + st.session_state.export_obligation_benefit - tax["tax_scenario_c"] - st.session_state.admin_cost - st.session_state.risk_reserve, "ریسک بالا اگر سود واقعی و مدارک هزینه قابل دفاع نباشد"],
    ]
    scenario_df = pd.DataFrame(
        scenario_rows,
        columns=[
            "سناریو",
            "مبنای درآمد/سود",
            "نرخ مالیات (%)",
            "مبلغ مالیات",
            "سود خالص صاحب کارت بعد از مالیات",
            "توضیح ریسک",
        ],
    )
    for col in ["مبنای درآمد/سود", "مبلغ مالیات", "سود خالص صاحب کارت بعد از مالیات"]:
        scenario_df[col] = scenario_df[col].map(fmt_money)
    scenario_df["نرخ مالیات (%)"] = scenario_df["نرخ مالیات (%)"].map(fmt_percent)
    st.dataframe(scenario_df, width="stretch", hide_index=True)
    if not st.session_state.inta_profile_aligned:
        st.error("فعالیت با پرونده مالیاتی صاحب کارت همخوان تأیید نشده است.")
    if not st.session_state.tax_advisor_confirmed:
        st.warning("مشاور مالیاتی هنوز اینتاکد و ضریب سود را تأیید نکرده است.")

elif page == "چک‌لیست ریسک":
    page_title(page, "کنترل‌های کلیدی معامله را علامت بزنید؛ هر مورد انجام‌نشده به امتیاز ریسک اضافه می‌شود.")
    cols = st.columns(2)
    for index, item in enumerate(RISK_WEIGHTS):
        with cols[index % 2]:
            checked_now = bool(st.session_state.get(f"risk_{item['id']}", False))
            card_class = "check-card"
            if not checked_now and item["weight"] >= 9:
                card_class += " check-card-open-critical"
            elif not checked_now and item["weight"] >= 6:
                card_class += " check-card-open-high"
            st.markdown(f"<div class='{card_class}'>", unsafe_allow_html=True)
            st.markdown(
                f"<span class='check-weight'>وزن ریسک: {item['weight']}</span>",
                unsafe_allow_html=True,
            )
            st.checkbox(
                item["label_fa"],
                key=f"risk_{item['id']}",
                help=item["explanation_fa"],
            )
            st.markdown(
                f"<div class='check-help'>{item['explanation_fa']}</div></div>",
                unsafe_allow_html=True,
            )
    result = compute()
    risk = result["risk"]
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("امتیاز ریسک", risk["risk_score"], "جمع وزن موارد انجام‌نشده", status_color_for_risk(risk["risk_score"]), money=False)
    with c2:
        kpi_card("سطح ریسک", risk["risk_level"], "طبقه‌بندی مدیریتی", status_color_for_risk(risk["risk_score"]), money=False)
    with c3:
        kpi_card("کنترل‌های انجام‌شده", risk["checked_count"], "موارد علامت‌خورده", "green", money=False)
    with c4:
        kpi_card("کنترل‌های باز", risk["unchecked_count"], "نیازمند پیگیری", "orange" if risk["unchecked_count"] else "green", money=False)
    st.markdown("<div class='warning-panel'><b>سه ریسک باز مهم</b></div>", unsafe_allow_html=True)
    for item in risk["open_top_risks"] or [{"label_fa": "همه کنترل‌های اصلی انجام شده‌اند.", "weight": 0, "explanation_fa": ""}]:
        st.markdown(f"<div class='risk-box'><b>{item['label_fa']}</b> | وزن: {item['weight']}<br>{item.get('explanation_fa', '')}</div>", unsafe_allow_html=True)

elif page == "محاسبه وثیقه پیشنهادی":
    page_title(page, "وثیقه پیشنهادی را برای پوشش ریسک VAT، مالیات، گمرک، ارز و رفع تعهد محاسبه کنید.")
    st.info("برای محافظت از صاحب کارت، وثیقه مالیاتی بهتر است بر اساس بدترین سناریوی مالیات عملکرد محاسبه شود، نه فقط سناریوی منتخب.")
    with st.container(border=True):
        st.subheader("پارامترهای وثیقه و ذخایر ریسک")
        c1, c2, c3 = st.columns(3)
        with c1:
            percent_input("درصد پوشش ریسک VAT (%)", "vat_risk_reserve_percent", "پوشش ریسک رد اعتبار یا اختلافات VAT.")
            percent_input("درصد پوشش ریسک مالیات عملکرد (%)", "income_tax_risk_reserve_percent", "بر مبنای بدترین مالیات بین سناریوها محاسبه می‌شود.")
        with c2:
            percent_input("درصد ذخیره ریسک گمرکی (%)", "customs_risk_reserve_percent")
            percent_input("درصد ذخیره ریسک ارزی (%)", "fx_risk_reserve_percent")
        with c3:
            percent_input("درصد ذخیره ریسک عدم رفع تعهد (%)", "export_obligation_risk_reserve_percent")
            money_input("ذخیره حقوقی/اداری ثابت", "fixed_legal_admin_reserve")
            guide_box("بدون وثیقه کافی برای VAT، مالیات، گمرک و رفع تعهد ارزی، ورود به معامله از منظر مدیریتی پرریسک است.")
    result = compute()
    kcols = st.columns(4)
    with kcols[0]:
        kpi_card("ذخیره VAT", result["collateral"]["vat_risk_reserve"], "پوشش VAT", "yellow")
    with kcols[1]:
        kpi_card("ذخیره مالیات عملکرد", result["collateral"]["income_tax_risk_reserve"], "بر پایه بدترین سناریو", "orange")
    with kcols[2]:
        kpi_card("ذخیره گمرک و ارز", result["collateral"]["customs_risk_reserve"] + result["collateral"]["fx_risk_reserve"], "ریسک ارزش و نرخ", "blue")
    with kcols[3]:
        kpi_card("حداقل وثیقه پیشنهادی", result["collateral"]["minimum_collateral"], "جمع ذخایر", "green")
    rows = [
        ("ذخیره ریسک VAT", result["collateral"]["vat_risk_reserve"]),
        ("مالیات سناریوی منتخب", result["tax"]["selected_income_tax"]),
        ("بدترین مالیات بین سناریوها", result["tax"]["worst_case_income_tax"]),
        ("ذخیره پیشنهادی مالیات عملکرد", result["collateral"]["income_tax_risk_reserve"]),
        ("ذخیره ریسک گمرکی", result["collateral"]["customs_risk_reserve"]),
        ("ذخیره ریسک ارزی", result["collateral"]["fx_risk_reserve"]),
        ("ذخیره ریسک عدم رفع تعهد", result["collateral"]["export_obligation_risk_reserve"]),
        ("ذخیره حقوقی/اداری ثابت", result["collateral"]["fixed_legal_admin_reserve"]),
        ("حداقل وثیقه پیشنهادی", result["collateral"]["minimum_collateral"]),
    ]
    st.dataframe(money_table(rows, "مبلغ"), width="stretch", hide_index=True)

elif page == "خلاصه تصمیم مدیریتی":
    page_title(page, "نمای نهایی معامله برای تصمیم‌گیری مدیریتی، کنترل ریسک و تعیین شروط اجرا.")
    result = compute()
    s = result["summary"]
    decision_color = status_color_for_decision(s["final_decision"])
    st.markdown(
        f"<div class='decision decision-{decision_color}'>تصمیم نهایی: {s['final_decision']}</div>",
        unsafe_allow_html=True,
    )
    cards = [
        ("ارزش کل واردات", s["import_value_total"], "مبنای ریالی معامله", "blue", True),
        ("بهای تمام‌شده با VAT", s["total_landed_cost"], "کل هزینه واردات", "neutral", True),
        ("VAT واردات", s["import_vat"], "پرداخت/اعتبار واردات", "yellow", True),
        ("VAT فروش داخلی", s["domestic_sale_vat"], "در صورتحساب داخلی", "yellow", True),
        ("سود خالص صاحب کارت", s["net_card_owner_benefit"], "پس از مالیات و ذخایر", "green" if s["net_card_owner_benefit"] > 0 else "red", True),
        ("امتیاز ریسک", s["risk_score"], s["risk_level"], status_color_for_risk(s["risk_score"]), False),
        ("حداقل وثیقه پیشنهادی", s["minimum_collateral"], "برای پوشش ریسک‌ها", "orange", True),
        ("تصمیم نهایی", s["final_decision"], "خروجی مدیریتی", decision_color, False),
    ]
    for row_start in range(0, len(cards), 4):
        cols = st.columns(4)
        for col, (label, value, subtitle, color, is_money) in zip(cols, cards[row_start:row_start + 4]):
            with col:
                kpi_card(label, value, subtitle, color, money=is_money)
    st.markdown(
        f"""
        <div class="soft-card">
            <h3>توصیه عملیاتی نهایی</h3>
            <p>{s['final_recommendation']}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if not result["inputs"]["buyer_pays_vat_separately"]:
        st.error(s["vat_buyer_payment_message"])
    else:
        st.success(s["vat_buyer_payment_message"])
    st.warning("حتی در صورت تهاتر، اگر کالا معاف نباشد، VAT فروش داخلی باید در صورتحساب داخلی درج شود.")
    st.subheader("شاخص‌های مالی معامله")
    financial_rows = [
        ("ارزش کل واردات", fmt_money(s["import_value_total"])),
        ("بهای تمام‌شده کل", fmt_money(s["total_landed_cost"])),
        ("کارمزد ناخالص صاحب کارت", fmt_money(s["gross_commission"])),
        ("درصد کارمزد از ارزش واردات", f"{s['commission_percent']:.2%}"),
        ("سود خالص صاحب کارت", fmt_money(s["net_card_owner_benefit"])),
        ("منفعت خالص به ازای هر دلار", fmt_money(s["net_benefit_per_usd"])),
        ("هزینه کل مجری/واردکننده واقعی", fmt_money(s["total_operator_cost"])),
        ("حداقل کارمزد قابل قبول هر دلار", fmt_money(s["minimum_commission_per_usd"])),
        ("فاصله کارمزد فعلی با حداقل کارمزد پیشنهادی", fmt_money(s["commission_gap_per_usd"])),
        ("هشدار کف کارمزد", s["commission_floor_warning"]),
    ]
    with st.container(border=True):
        st.dataframe(pd.DataFrame(financial_rows, columns=["شاخص", "مقدار"]), width="stretch", hide_index=True)

    st.subheader("شاخص‌های VAT و مالیات")
    tax_rows = [
        ("نام سناریوی مالیات منتخب", s["selected_tax_scenario_name"]),
        ("مالیات سناریوی منتخب", fmt_money(s["selected_income_tax"])),
        ("VAT واردات", fmt_money(s["import_vat"])),
        ("VAT فروش داخلی", fmt_money(s["domestic_sale_vat"])),
        ("خالص VAT قابل پرداخت/تهاتر", fmt_money(s["net_vat_position"])),
        ("وضعیت دریافت VAT از خریدار", s["vat_buyer_payment_message"]),
    ]
    with st.container(border=True):
        st.dataframe(pd.DataFrame(tax_rows, columns=["شاخص", "مقدار"]), width="stretch", hide_index=True)

    st.subheader("شاخص‌های ریسک و تصمیم")
    risk_rows = [
        ("حداقل وثیقه پیشنهادی", fmt_money(s["minimum_collateral"])),
        ("امتیاز ریسک", fmt_number(s["risk_score"])),
        ("سطح ریسک", s["risk_level"]),
        ("تعداد کنترل‌های انجام‌شده", fmt_number(s["checked_count"])),
        ("تعداد کنترل‌های انجام‌نشده", fmt_number(s["unchecked_count"])),
        ("تصمیم نهایی", s["final_decision"]),
        ("توصیه عملیاتی نهایی", s["final_recommendation"]),
    ]
    with st.container(border=True):
        st.dataframe(pd.DataFrame(risk_rows, columns=["شاخص", "مقدار"]), width="stretch", hide_index=True)
    st.markdown("<div class='warning-panel'><b>سه ریسک باز مهم</b></div>", unsafe_allow_html=True)
    for risk_label in s["open_top_risks"]:
        st.markdown(f"<div class='risk-box'>{risk_label}</div>", unsafe_allow_html=True)

elif page == "خروجی Excel":
    page_title(page, "خروجی‌های Excel را بر اساس آخرین مقادیر فعلی داشبورد دریافت کنید.")
    result = compute()
    inputs = result["inputs"]
    outputs = result["outputs"]
    checklist = checklist_export_rows()
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(
            "<div class='download-card'><h3>گزارش کامل Excel</h3><p>همه ورودی‌ها، محاسبات، سناریوهای مالیاتی، ریسک، وثیقه و خلاصه تصمیم.</p>",
            unsafe_allow_html=True,
        )
        st.download_button(
            "دانلود گزارش کامل Excel",
            data=export_all_to_excel(inputs, outputs, checklist),
            file_name="full_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)
    with c2:
        st.markdown(
            "<div class='download-card'><h3>چک‌لیست ریسک</h3><p>وضعیت کنترل‌ها، وزن ریسک، ریسک لحاظ‌شده و توضیح هر کنترل.</p>",
            unsafe_allow_html=True,
        )
        st.download_button(
            "دانلود چک‌لیست ریسک",
            data=export_checklist_to_excel(checklist),
            file_name="risk_checklist.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)
    with c3:
        st.markdown(
            "<div class='download-card'><h3>خلاصه مدیریتی</h3><p>شاخص‌های کلیدی، تصمیم نهایی، وضعیت ریسک، وثیقه و توصیه عملیاتی.</p>",
            unsafe_allow_html=True,
        )
        st.download_button(
            "دانلود خلاصه مدیریتی",
            data=export_summary_to_excel(outputs),
            file_name="management_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)
