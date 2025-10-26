# app.py
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import Workbook

st.set_page_config(page_title="SmartInvestPro – ROI Calculator", layout="wide")

# =========================
# Helpers (no numpy_financial)
# =========================
def pmt(rate: float, nper: int, pv: float) -> float:
    """
    Monthly payment for an amortizing loan.
    Args:
        rate: per-period interest rate (e.g., annual_rate/12 as a decimal)
        nper: total number of periods (months)
        pv:   present value (loan principal, positive)
    Returns:
        Positive payment amount per month.
    """
    if nper <= 0:
        return 0.0
    if abs(rate) < 1e-12:
        return pv / nper
    return (pv * rate) / (1 - (1 + rate) ** (-nper))


def dataframe_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Cashflow") -> bytes:
    """Export a DataFrame to an in-memory XLSX file."""
    buf = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.append(list(df.columns))
    for _, row in df.iterrows():
        ws.append([row[col] if not isinstance(row[col], (np.floating, float)) else float(row[col])
                   for col in df.columns])
    wb.save(buf)
    return buf.getvalue()


# =========================
# UI
# =========================
st.title("🏠 SmartInvestPro – Real Estate ROI Calculator")

with st.sidebar:
    st.subheader("Inputs")
    purchase_price = st.number_input("Purchase Price ($)", min_value=50_000, max_value=20_000_000, value=350_000, step=1_000)
    down_payment_pct = st.number_input("Down Payment (%)", min_value=0.0, max_value=100.0, value=20.0, step=0.5)
    loan_term_years = st.number_input("Loan Term (years)", min_value=1, max_value=40, value=30, step=1)
    interest_rate_pct = st.number_input("Interest Rate (%)", min_value=0.0, max_value=20.0, value=6.5, step=0.1)

    st.markdown("---")
    rent_monthly = st.number_input("Monthly Rent ($)", min_value=0, max_value=50_000, value=2_500, step=50)
    annual_rent_increase_pct = st.number_input("Annual Rent Increase (%)", min_value=0.0, max_value=20.0, value=2.0, step=0.5)

    st.markdown("---")
    taxes_annual = st.number_input("Annual Property Taxes ($)", min_value=0, max_value=100_000, value=4_000, step=100)
    insurance_annual = st.number_input("Annual Insurance ($)", min_value=0, max_value=50_000, value=1_500, step=50)
    maintenance_pct = st.number_input("Maintenance (% of rent)", min_value=0.0, max_value=50.0, value=5.0, step=0.5)
    management_pct = st.number_input("Management (% of rent)", min_value=0.0, max_value=50.0, value=8.0, step=0.5)
    vacancy_pct = st.number_input("Vacancy (% of rent)", min_value=0.0, max_value=50.0, value=5.0, step=0.5)

    st.markdown("---")
    holding_years = st.number_input("Holding Period (years)", min_value=1, max_value=50, value=10, step=1)

# =========================
# Calculations
# =========================
down_payment = purchase_price * (down_payment_pct / 100.0)
loan_amount = max(purchase_price - down_payment, 0.0)

monthly_rate = (interest_rate_pct / 100.0) / 12.0
months = int(loan_term_years * 12)
monthly_payment = pmt(monthly_rate, months, loan_amount)  # positive

operating_pct = (maintenance_pct + management_pct + vacancy_pct) / 100.0
fixed_opex_monthly = (taxes_annual + insurance_annual) / 12.0

rows = []
current_rent = float(rent_monthly)

for year in range(1, int(holding_years) + 1):
    if year > 1:
        current_rent *= (1 + annual_rent_increase_pct / 100.0)

    annual_rent = current_rent * 12.0
    variable_opex_annual = (current_rent * operating_pct) * 12.0
    fixed_opex_annual = fixed_opex_monthly * 12.0
    total_opex_annual = variable_opex_annual + fixed_opex_annual

    debt_service_annual = monthly_payment * 12.0
    annual_cashflow = annual_rent - debt_service_annual - total_opex_annual

    rows.append({
        "Year": year,
        "Monthly Rent ($)": current_rent,
        "Annual Rent ($)": annual_rent,
        "Debt Service ($/yr)": debt_service_annual,
        "Opex – Fixed ($/yr)": fixed_opex_annual,
        "Opex – Var ($/yr)": variable_opex_annual,
        "Annual Cashflow ($)": annual_cashflow,
    })

df = pd.DataFrame(rows)

initial_investment = down_payment
total_cashflow = float(df["Annual Cashflow ($)"].sum()) if not df.empty else 0.0
simple_roi_pct = (total_cashflow / initial_investment * 100.0) if initial_investment > 0 else 0.0

# =========================
# Display
# =========================
st.subheader("📈 Key Metrics")
m1, m2, m3 = st.columns(3)
m1.metric("Monthly Payment", f"${monthly_payment:,.2f}")
m2.metric("Total Cashflow (Hold)", f"${total_cashflow:,.2f}")
m3.metric("ROI (Simple %)", f"{simple_roi_pct:,.2f}%")

st.markdown("### 💵 Yearly Cashflow Table")
fmt_cols = {
    "Monthly Rent ($)": "{:,.0f}",
    "Annual Rent ($)": "{:,.0f}",
    "Debt Service ($/yr)": "{:,.0f}",
    "Opex – Fixed ($/yr)": "{:,.0f}",
    "Opex – Var ($/yr)": "{:,.0f}",
    "Annual Cashflow ($)": "{:,.0f}",
}
st.dataframe(df.style.format(fmt_cols), use_container_width=True)

st.markdown("### 📊 Annual Cashflow Trend")
st.line_chart(df.set_index("Year")["Annual Cashflow ($)"])

# =========================
# Download Excel
# =========================
excel_bytes = dataframe_to_excel_bytes(df)
st.download_button(
    "📥 Download Excel",
    data=excel_bytes,
    file_name="smartinvestpro_roi_cashflow.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.caption(
    "Tip: Adjust assumptions in the sidebar. This simplified model is for quick scenario testing. "
    "We can add multi-unit inputs, amortization breakdowns, and sale proceeds next."
)
