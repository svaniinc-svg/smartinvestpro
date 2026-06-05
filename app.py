import math
from typing import List, Dict, Any

import pandas as pd
import streamlit as st

try:
    import numpy_financial as npf
except Exception:  # keeps app from crashing if package missing
    npf = None


st.set_page_config(page_title="Smart Rental ROI", page_icon="🏠", layout="wide")
st.title("🏠 Smart Rental ROI Calculator")
st.caption("Rental cash flow, NOI, cash-on-cash, ROI, IRR, DSCR, and sale/profit analysis")


def money(x: float) -> str:
    try:
        return f"${x:,.0f}"
    except Exception:
        return "$0"


def pct(x: float) -> str:
    try:
        return f"{x:.2f}%"
    except Exception:
        return "0.00%"


def monthly_payment(principal: float, annual_rate_pct: float, years: int) -> float:
    if principal <= 0 or years <= 0:
        return 0.0
    r = annual_rate_pct / 100 / 12
    n = years * 12
    if abs(r) < 1e-12:
        return principal / n
    return principal * (r * (1 + r) ** n) / ((1 + r) ** n - 1)


def amortization_balance(principal: float, annual_rate_pct: float, years: int, months_paid: int) -> float:
    if principal <= 0:
        return 0.0
    r = annual_rate_pct / 100 / 12
    n = years * 12
    months_paid = max(0, min(months_paid, n))
    pmt = monthly_payment(principal, annual_rate_pct, years)
    if abs(r) < 1e-12:
        return max(0.0, principal - pmt * months_paid)
    return max(0.0, principal * (1 + r) ** months_paid - pmt * (((1 + r) ** months_paid - 1) / r))


def safe_irr(cashflows: List[float]) -> float | None:
    if npf is None or len(cashflows) < 2 or not any(c < 0 for c in cashflows) or not any(c > 0 for c in cashflows):
        return None
    try:
        result = npf.irr(cashflows)
        if result is None or math.isnan(result):
            return None
        return float(result)
    except Exception:
        return None


with st.sidebar:
    st.header("Property Inputs")
    purchase_price = st.number_input("Purchase Price", min_value=0.0, value=300000.0, step=5000.0, format="%.2f")
    down_payment_pct = st.number_input("Down Payment %", min_value=0.0, max_value=100.0, value=20.0, step=1.0)
    interest_rate = st.number_input("Interest Rate %", min_value=0.0, max_value=30.0, value=7.0, step=0.125)
    loan_years = st.number_input("Loan Term Years", min_value=1, max_value=40, value=30, step=1)
    override_payment = st.number_input("Optional Monthly PI Override", min_value=0.0, value=0.0, step=50.0)

    st.header("Acquisition / Sale")
    closing_costs_buy = st.number_input("Buyer Closing Costs", min_value=0.0, value=0.0, step=1000.0)
    rehab_cost = st.number_input("Repair / Rehab Cost", min_value=0.0, value=0.0, step=1000.0)
    hold_years = st.number_input("Hold Period Years", min_value=1, max_value=40, value=5, step=1)
    appreciation_pct = st.number_input("Annual Appreciation %", min_value=-20.0, max_value=50.0, value=3.0, step=0.5)
    sale_cost_pct = st.number_input("Sale Costs + Commission %", min_value=0.0, max_value=20.0, value=7.0, step=0.5)

st.subheader("Rent Roll")
starting_units = pd.DataFrame(
    [
        {"Unit": "1", "Beds": 3, "Baths": 2.5, "Monthly Rent": 2050.0},
    ]
)
units = st.data_editor(
    starting_units,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "Monthly Rent": st.column_config.NumberColumn(format="$%.2f", min_value=0.0),
        "Beds": st.column_config.NumberColumn(min_value=0, step=1),
        "Baths": st.column_config.NumberColumn(min_value=0.0, step=0.5),
    },
)

rent_col = "Monthly Rent"
if rent_col not in units.columns:
    units[rent_col] = 0.0
units[rent_col] = pd.to_numeric(units[rent_col], errors="coerce").fillna(0.0)
gross_monthly_rent = float(units[rent_col].sum())

st.subheader("Monthly Expenses")
col1, col2, col3, col4 = st.columns(4)
with col1:
    hoa = st.number_input("HOA", min_value=0.0, value=175.0, step=25.0)
    taxes_annual = st.number_input("Property Taxes Annual", min_value=0.0, value=3180.0, step=100.0)
    insurance_annual = st.number_input("Insurance Annual", min_value=0.0, value=1200.0, step=100.0)
with col2:
    gas = st.number_input("Gas", min_value=0.0, value=0.0, step=25.0)
    electricity = st.number_input("Electricity", min_value=0.0, value=0.0, step=25.0)
    water = st.number_input("Water", min_value=0.0, value=0.0, step=25.0)
with col3:
    sewer = st.number_input("Sewer", min_value=0.0, value=0.0, step=25.0)
    garbage = st.number_input("Garbage", min_value=0.0, value=0.0, step=25.0)
    lawn = st.number_input("Lawn", min_value=0.0, value=0.0, step=25.0)
with col4:
    management_pct = st.number_input("Management % of Rent", min_value=0.0, max_value=100.0, value=0.0, step=0.5)
    vacancy_pct = st.number_input("Vacancy % of Rent", min_value=0.0, max_value=100.0, value=5.0, step=0.5)
    maintenance_pct = st.number_input("Maintenance % of Rent", min_value=0.0, max_value=100.0, value=5.0, step=0.5)

loan_amount = max(0.0, purchase_price * (1 - down_payment_pct / 100))
down_payment = purchase_price - loan_amount
calculated_pi = monthly_payment(loan_amount, interest_rate, int(loan_years))
monthly_pi = override_payment if override_payment > 0 else calculated_pi

monthly_taxes = taxes_annual / 12
monthly_insurance = insurance_annual / 12
management = gross_monthly_rent * management_pct / 100
vacancy = gross_monthly_rent * vacancy_pct / 100
maintenance = gross_monthly_rent * maintenance_pct / 100
operating_expenses = hoa + monthly_taxes + monthly_insurance + gas + electricity + water + sewer + garbage + lawn + management + vacancy + maintenance
noi_monthly = gross_monthly_rent - operating_expenses
cash_flow_monthly = noi_monthly - monthly_pi

annual_noi = noi_monthly * 12
annual_debt_service = monthly_pi * 12
cash_invested = down_payment + closing_costs_buy + rehab_cost
cap_rate = annual_noi / purchase_price * 100 if purchase_price else 0.0
coc = cash_flow_monthly * 12 / cash_invested * 100 if cash_invested else 0.0
dscr = annual_noi / annual_debt_service if annual_debt_service else 0.0
break_even_occupancy = (operating_expenses + monthly_pi - vacancy) / gross_monthly_rent * 100 if gross_monthly_rent else 0.0

future_value = purchase_price * ((1 + appreciation_pct / 100) ** hold_years)
loan_balance = amortization_balance(loan_amount, interest_rate, int(loan_years), int(hold_years * 12))
sale_costs = future_value * sale_cost_pct / 100
net_sale_proceeds = future_value - sale_costs - loan_balance
total_cash_flow = cash_flow_monthly * 12 * hold_years
total_profit = net_sale_proceeds + total_cash_flow - cash_invested
roi = total_profit / cash_invested * 100 if cash_invested else 0.0

annual_cashflows = [-cash_invested] + [cash_flow_monthly * 12 for _ in range(int(hold_years) - 1)] + [cash_flow_monthly * 12 + net_sale_proceeds]
irr = safe_irr(annual_cashflows)

st.divider()
cols = st.columns(6)
cols[0].metric("Gross Rent", money(gross_monthly_rent))
cols[1].metric("Monthly NOI", money(noi_monthly))
cols[2].metric("Monthly Cash Flow", money(cash_flow_monthly))
cols[3].metric("Cap Rate", pct(cap_rate))
cols[4].metric("Cash-on-Cash", pct(coc))
cols[5].metric("DSCR", f"{dscr:.2f}x")

cols2 = st.columns(5)
cols2[0].metric("Cash Invested", money(cash_invested))
cols2[1].metric("Future Sale Value", money(future_value))
cols2[2].metric("Net Sale Proceeds", money(net_sale_proceeds))
cols2[3].metric("Total Profit", money(total_profit))
cols2[4].metric("Hold ROI / IRR", f"{pct(roi)} / {pct(irr*100) if irr is not None else 'N/A'}")

with st.expander("Expense Breakdown", expanded=True):
    expense_df = pd.DataFrame(
        [
            ["HOA", hoa], ["Taxes", monthly_taxes], ["Insurance", monthly_insurance], ["Gas", gas],
            ["Electricity", electricity], ["Water", water], ["Sewer", sewer], ["Garbage", garbage],
            ["Lawn", lawn], ["Management Reserve", management], ["Vacancy Reserve", vacancy],
            ["Maintenance Reserve", maintenance], ["Operating Expenses", operating_expenses], ["Mortgage PI", monthly_pi],
        ],
        columns=["Item", "Monthly Amount"],
    )
    st.dataframe(expense_df, use_container_width=True, hide_index=True)

with st.expander("Sale / Hold Analysis", expanded=False):
    sale_df = pd.DataFrame(
        [
            ["Purchase Price", purchase_price],
            ["Loan Amount", loan_amount],
            ["Loan Balance at Sale", loan_balance],
            ["Sale Value", future_value],
            ["Sale Costs", sale_costs],
            ["Net Sale Proceeds", net_sale_proceeds],
            ["Total Cash Flow During Hold", total_cash_flow],
            ["Total Profit", total_profit],
            ["Break-even Occupancy", break_even_occupancy],
        ],
        columns=["Metric", "Value"],
    )
    st.dataframe(sale_df, use_container_width=True, hide_index=True)

csv = pd.DataFrame({
    "Metric": ["Gross Monthly Rent", "Operating Expenses", "NOI Monthly", "Mortgage PI", "Monthly Cash Flow", "Annual NOI", "Cap Rate %", "Cash-on-Cash %", "DSCR", "Total Profit", "ROI %", "IRR %"],
    "Value": [gross_monthly_rent, operating_expenses, noi_monthly, monthly_pi, cash_flow_monthly, annual_noi, cap_rate, coc, dscr, total_profit, roi, irr * 100 if irr is not None else None],
}).to_csv(index=False)
st.download_button("Download Summary CSV", csv, file_name="rental_roi_summary.csv", mime="text/csv")

st.info("Tip: Add numpy-financial to requirements.txt for IRR support.")
