# app.py
import streamlit as st
import pandas as pd
import numpy as np
import numpy_financial as npf
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
import matplotlib.pyplot as plt

st.set_page_config(page_title="Real Estate ROI Calculator", layout="wide")

# Hide Streamlit elements (fork button, menu, footer)
hide_streamlit_style = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
.stActionButton {display: none;}
div[data-testid="stToolbar"] {visibility: hidden;}
div[data-testid="stDecoration"] {visibility: hidden;}
div[data-testid="stStatusWidget"] {visibility: hidden;}
.reportview-container .main footer {visibility: hidden;}
.stDeployButton {display: none;}
#stDecoration {display: none;}
[data-testid="stToolbar"] {display: none !important;}
[data-testid="stDecoration"] {display: none !important;}
[data-testid="stStatusWidget"] {display: none !important;}
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

st.title("🏡 Real Estate ROI Calculator with Multiple Units")

# -------------------------
# Financial helpers
# -------------------------
def pmt(monthly_rate, nper, pv):
    if nper <= 0:
        return 0.0
    if abs(monthly_rate) < 1e-12:
        return pv / nper
    return monthly_rate * pv / (1 - (1 + monthly_rate) ** (-nper))

def _npv_scalar(rate, cashflows):
    t = np.arange(len(cashflows), dtype=float)
    return float(np.sum(np.array(cashflows, dtype=float) / np.power(1.0 + rate, t)))

def _irr_stable(cashflows, guess=0.01):
    cf = np.array(cashflows, dtype=float)
    if not (np.any(cf > 0) and np.any(cf < 0)):
        return None
    # Try polynomial roots first
    try:
        coeffs = cf[::-1]
        roots = np.roots(coeffs)
        candidates = []
        for r in roots:
            if abs(r.imag) < 1e-10 and r.real != 0:
                x = float(r.real)
                if x > 0:
                    rate = (1.0 / x) - 1.0
                    if rate > -0.999999:
                        candidates.append(rate)
        if candidates:
            return min(candidates, key=lambda z: abs(z - guess))
    except Exception:
        pass
    # Bracket + bisection on per-period IRR
    low, high = -0.999999, 10.0
    f_low, f_high = _npv_scalar(low, cf), _npv_scalar(high, cf)
    tries = 0
    while f_low * f_high > 0 and tries < 60:
        high *= 2.0
        f_high = _npv_scalar(high, cf)
        tries += 1
    if f_low * f_high > 0:
        return None
    for _ in range(200):
        mid = (low + high) / 2.0
        f_mid = _npv_scalar(mid, cf)
        if abs(f_mid) < 1e-10:
            return mid
        if f_low * f_mid < 0:
            high, f_high = mid, f_mid
        else:
            low, f_low = mid, f_mid
    return (low + high) / 2.0

def _annualize_from_monthly(r):
    return (1 + r) ** 12 - 1 if r is not None and not np.isnan(r) else None

# -------------------------
# Styling helpers
# -------------------------
def _style_red_neg(val):
    try:
        if isinstance(val, (int, float, np.number)) and val < 0:
            return "color:red;"
    except Exception:
        pass
    return ""

def _fmt_metric(metric, val):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return "N/A"
    percent_keys = ["Cap Rate", "Cash on Cash", "IRR", "ROI", "Effective Rate", "Occupancy"]
    currency_keys = ["Price", "Payment", "Subsidy", "NOI", "Debt Service", "Cash Flow",
                     "Loan", "Proceeds", "Equity", "Profit", "Investment", "Closing Costs", "Commission", "Balance"]
    try:
        if any(k in metric for k in percent_keys):
            return f"{val*100:.2f}%"
        if any(k in metric for k in currency_keys):
            return f"${val:,.2f}"
        if metric.startswith("DCR"):
            return f"{val:.2f}"
        if isinstance(val, (int, float, np.number)):
            return f"{val:,.2f}"
    except Exception:
        pass
    return val

# -------------------------
# Buydown helpers
# -------------------------
def _buydown_preview(loan_amount, annual_rate, nper):
    """Return Yr1/Yr2 effective rates/payments and subsidy totals."""
    if nper <= 0 or loan_amount <= 0:
        return (annual_rate, annual_rate, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0)
    note_pmt = pmt(annual_rate/12.0, nper, loan_amount)
    eff1 = max(annual_rate - 0.02, 0.0)
    eff2 = max(annual_rate - 0.01, 0.0)
    p1 = pmt(eff1/12.0, nper, loan_amount)
    p2 = pmt(eff2/12.0, nper, loan_amount)
    m1 = min(12, nper)
    m2 = min(max(nper - 12, 0), 12)
    sub1 = (note_pmt - p1) * m1
    sub2 = (note_pmt - p2) * m2
    return eff1, eff2, p1, p2, note_pmt, sub1, sub2, (sub1 + sub2)

def _buydown_oop_series(loan_amount, annual_rate, nper, note_payment, use_buydown_2_1):
    """Borrower out-of-pocket monthly payments (amortization still at note payment)."""
    if not use_buydown_2_1:
        return np.full(nper, note_payment, dtype=float)
    eff1 = max(annual_rate - 0.02, 0.0)
    eff2 = max(annual_rate - 0.01, 0.0)
    p1 = pmt(eff1 / 12.0, nper, loan_amount)
    p2 = pmt(eff2 / 12.0, nper, loan_amount)
    arr = []
    for m in range(1, nper + 1):
        if m <= 12:
            arr.append(p1)
        elif m <= 24:
            arr.append(p2)
        else:
            arr.append(note_payment)
    return np.array(arr, dtype=float)

# -------------------------
# Amortization + Metrics
# -------------------------
def amortization_schedule(loan_amount, annual_rate, years, down_payment=0, override_payment=None, use_buydown_2_1=False):
    """
    Returns a schedule where:
    - Note Payment drives Interest/Principal/Balance (note rate).
    - Borrower Payment reflects 2/1 buydown out-of-pocket.
    - Subsidy = Note Payment - Borrower Payment.
    """
    monthly_rate = annual_rate / 12.0
    nper = int(round(years * 12))
    cols = ["Month", "Note Payment", "Borrower Payment", "Subsidy", "Principal", "Interest", "Balance", "Equity"]
    if nper <= 0:
        return pd.DataFrame(columns=cols)

    note_payment = override_payment if (override_payment and override_payment > 0) else pmt(monthly_rate, nper, loan_amount)
    oop = _buydown_oop_series(loan_amount, annual_rate, nper, note_payment, use_buydown_2_1)

    balance = float(loan_amount)
    rows = []
    for m in range(1, nper + 1):
        interest = balance * monthly_rate
        principal = note_payment - interest
        if m == nper:
            principal = balance
            true_note_payment = principal + interest
            balance = 0.0
        else:
            balance -= principal
            true_note_payment = note_payment
            if balance < 1e-8:
                balance = 0.0

        borrower_pay = float(oop[m - 1])
        subsidy = true_note_payment - borrower_pay
        equity = down_payment + (loan_amount - balance)

        rows.append({
            "Month": m,
            "Note Payment": round(true_note_payment, 2),
            "Borrower Payment": round(borrower_pay, 2),
            "Subsidy": round(subsidy, 2),
            "Principal": round(principal, 2),
            "Interest": round(interest, 2),
            "Balance": round(balance, 2),
            "Equity": round(equity, 2),
        })
    return pd.DataFrame(rows)

def build_yearly_cashflows(hold_years, nper, effective_rent_base, monthly_expenses, monthly_taxes, monthly_insurance,
                           monthly_hoa, mgmt_pct, rent_appreciation_pct, oop_series):
    """Returns a DataFrame of yearly cashflows (1..hold_years) honoring rent growth and partial first/last years."""
    rows = []
    cum_cf = 0.0
    for year in range(1, hold_years + 1):
        # Rent growth compounded annually
        rent_multiplier = (1 + rent_appreciation_pct / 100.0) ** (year - 1)
        eff_rent_y = effective_rent_base * rent_multiplier  # monthly effective rent (after vacancy)
        mgmt_fee_y = eff_rent_y * (mgmt_pct / 100.0)
        tot_mo_exp_y = monthly_expenses + monthly_taxes + monthly_insurance + monthly_hoa + mgmt_fee_y

        start = (year - 1) * 12
        end = min(year * 12, nper)
        months = max(0, end - start)
        if months == 0:
            break

        noi_year = (eff_rent_y - tot_mo_exp_y) * 12.0
        noi_prorated = noi_year * (months / 12.0)
        debt_service = float(oop_series[start:end].sum())
        cash_flow = noi_prorated - debt_service
        cum_cf += cash_flow

        rows.append({
            "Year": year,
            "Months Counted": months,
            "Gross Rent (after vacancy) /mo": round(eff_rent_y, 2),
            "Mgmt Fee /mo": round(mgmt_fee_y, 2),
            "Other OpEx /mo (excl taxes/ins/HOA)": round(monthly_expenses, 2),
            "Taxes /mo": round(monthly_taxes, 2),
            "Insurance /mo": round(monthly_insurance, 2),
            "HOA /mo": round(monthly_hoa, 2),
            "Total OpEx /mo": round(tot_mo_exp_y, 2),
            "NOI (annual, prorated)": round(noi_prorated, 2),
            "Debt Service (annual, prorated)": round(debt_service, 2),
            "Cash Flow (annual)": round(cash_flow, 2),
            "Cumulative Cash Flow": round(cum_cf, 2),
        })
    return pd.DataFrame(rows)

def calculate_metrics(purchase_price, down_payment, annual_rate, years,
                      monthly_rent, monthly_expenses, monthly_taxes, monthly_insurance, monthly_hoa,
                      vacancy_pct=0.0, mgmt_pct=0.0, hold_years=5,
                      rent_appreciation_pct=0.0, appreciation_pct=0.0,
                      buy_closing_costs=0.0, sale_closing_costs=0.0, sale_commission_pct=0.0,
                      override_payment=None, use_buydown_2_1=False, buydown_paid_by="Buyer",
                      sale_price_override=None, buy_cc_paid_by="Buyer"):
    loan_amount = max(0.0, purchase_price - down_payment)
    amort = amortization_schedule(
        loan_amount, annual_rate, years,
        down_payment=down_payment, override_payment=override_payment, use_buydown_2_1=use_buydown_2_1
    )
    nper = len(amort)
    monthly_note_payment = amort["Note Payment"].iloc[0] if nper > 0 else 0.0

    # Base monthly figures (Yr1)
    effective_rent = monthly_rent * (1 - vacancy_pct / 100.0)  # after vacancy
    mgmt_fee = effective_rent * (mgmt_pct / 100.0)
    total_monthly_expenses = monthly_expenses + monthly_taxes + monthly_insurance + monthly_hoa + mgmt_fee
    noi = (effective_rent - total_monthly_expenses) * 12.0  # Yr1 annual NOI

    # Out-of-pocket debt service series
    oop_series = amort["Borrower Payment"].to_numpy(dtype=float) if nper > 0 else np.array([])

    # Buydown preview/subsidy
    eff_rate_y1, eff_rate_y2, borrower_pmt_y1, borrower_pmt_y2, note_pmt_ref, subsidy_y1, subsidy_y2, subsidy_total = \
        _buydown_preview(loan_amount, annual_rate, nper) if use_buydown_2_1 else \
        (annual_rate, annual_rate, monthly_note_payment, monthly_note_payment, monthly_note_payment, 0.0, 0.0, 0.0)

    # Initial investment: down + buy CC (if Buyer) + buydown (if Buyer)
    initial_investment = down_payment
    if buy_cc_paid_by == "Buyer":
        initial_investment += buy_closing_costs
    if use_buydown_2_1 and buydown_paid_by == "Buyer":
        initial_investment += subsidy_total

    # Yr1 snapshots
    months_y1 = min(12, nper)
    months_y2 = min(max(nper - 12, 0), 12)
    annual_debt_y1 = float(oop_series[:months_y1].sum()) if nper > 0 else 0.0
    annual_cash_flow_y1 = noi - annual_debt_y1
    annual_debt_y2 = float(oop_series[12:12 + months_y2].sum()) if nper > 12 else None
    annual_cash_flow_y2 = (noi - annual_debt_y2) if nper > 12 else None

    # DCR (Yr1)
    dcr_y1 = (noi / annual_debt_y1) if annual_debt_y1 > 0 else None

    # Break-Even Occupancy % (Yr1) — fraction (e.g., 0.85 = 85%)
    gpi_annual = monthly_rent * 12.0
    mgmt_frac = (mgmt_pct / 100.0)
    fixed_opex_annual = (monthly_expenses + monthly_taxes + monthly_insurance + monthly_hoa) * 12.0
    denom = gpi_annual * (1.0 - mgmt_frac) if (1.0 - mgmt_frac) != 0 else np.nan
    break_even_occ = ((fixed_opex_annual + annual_debt_y1) / denom) if denom and not np.isnan(denom) else None

    # Display metrics
    cap_rate = (noi / purchase_price) if purchase_price else None
    cash_on_cash = (annual_cash_flow_y1 / initial_investment) if initial_investment > 0 else None

    # IRR across full term with sale
    term_years = int(round(years))
    annual_irr_term = None
    if term_years > 0 and nper > 0:
        cash_flows_term = [-initial_investment]
        for year in range(1, term_years + 1):
            rent_multiplier = (1 + rent_appreciation_pct / 100.0) ** (year - 1)
            eff_rent_y = effective_rent * rent_multiplier
            mgmt_fee_y = eff_rent_y * (mgmt_pct / 100.0)
            tot_mo_exp_y = monthly_expenses + monthly_taxes + monthly_insurance + monthly_hoa + mgmt_fee_y
            noi_y = (eff_rent_y - tot_mo_exp_y) * 12.0
            start = (year - 1) * 12
            end = min(year * 12, nper)
            months = end - start
            if months <= 0:
                continue
            annual_debt_y = float(oop_series[start:end].sum())
            cf_y = (noi_y * (months / 12.0)) - annual_debt_y
            cash_flows_term += [cf_y / months] * months

        sale_price_term = purchase_price * ((1 + appreciation_pct / 100.0) ** term_years)
        loan_bal_term = float(amort.iloc[-1]["Balance"]) if nper > 0 else loan_amount
        sale_commission_term = sale_price_term * (sale_commission_pct / 100.0)
        net_sale_term = sale_price_term - loan_bal_term - sale_commission_term - sale_closing_costs
        cash_flows_term.append(net_sale_term)

        try:
            per_period_irr = npf.irr(cash_flows_term)
            if per_period_irr is None or np.isnan(per_period_irr):
                per_period_irr = _irr_stable(np.array(cash_flows_term))
        except Exception:
            per_period_irr = _irr_stable(np.array(cash_flows_term))
        annual_irr_term = _annualize_from_monthly(per_period_irr)

    # Hold-period prediction & IRR
    hold_months = int(min(hold_years, years) * 12)
    if hold_months == 0 or len(amort) == 0:
        equity_at_hold = down_payment
        loan_balance_at_hold = loan_amount
    elif len(amort) >= hold_months:
        equity_at_hold = float(amort.loc[hold_months - 1, "Equity"])
        loan_balance_at_hold = float(amort.loc[hold_months - 1, "Balance"])
    else:
        equity_at_hold = down_payment
        loan_balance_at_hold = loan_amount

    yearly_cf_df = build_yearly_cashflows(
        hold_years=hold_years,
        nper=nper,
        effective_rent_base=effective_rent,
        monthly_expenses=monthly_expenses,
        monthly_taxes=monthly_taxes,
        monthly_insurance=monthly_insurance,
        monthly_hoa=monthly_hoa,
        mgmt_pct=mgmt_pct,
        rent_appreciation_pct=rent_appreciation_pct,
        oop_series=oop_series
    )

    cash_flows_hold = [-initial_investment]
    total_cash_flow_hold = 0.0
    for _, r in yearly_cf_df.iterrows():
        months = int(r["Months Counted"])
        if months > 0:
            cf = float(r["Cash Flow (annual)"])
            total_cash_flow_hold += cf
            cash_flows_hold += [cf / months] * months

    sale_price = float(sale_price_override) if (sale_price_override and sale_price_override > 0) else \
                 purchase_price * ((1 + appreciation_pct / 100.0) ** min(hold_years, years))
    sale_commission = sale_price * (sale_commission_pct / 100.0)
    net_sale_proceeds = sale_price - loan_balance_at_hold - sale_commission - sale_closing_costs
    cash_flows_hold.append(net_sale_proceeds)

    total_profit = net_sale_proceeds + total_cash_flow_hold - initial_investment
    roi_hold = (total_profit / initial_investment) if initial_investment > 0 else None

    hold_irr = None
    try:
        hold_irr = npf.irr(cash_flows_hold)
        if hold_irr is None or np.isnan(hold_irr):
            hold_irr = _irr_stable(np.array(cash_flows_hold))
    except Exception:
        hold_irr = _irr_stable(np.array(cash_flows_hold))
    hold_irr_annual = _annualize_from_monthly(hold_irr)

    # Dynamic labels based on buydown selection
    borrower_payment_label_y1 = "Borrower Payment (Yr1)" if use_buydown_2_1 else "Borrower Payment"
    borrower_payment_label_y2 = "Borrower Payment (Yr2)" if use_buydown_2_1 else None
    effective_rate_label_y1 = "Effective Rate (Yr1)" if use_buydown_2_1 else "Effective Rate"
    effective_rate_label_y2 = "Effective Rate (Yr2)" if use_buydown_2_1 else None
    annual_debt_label_y1 = "Annual Debt Service (Yr1)" if use_buydown_2_1 else "Annual Debt Service"
    annual_debt_label_y2 = "Annual Debt Service (Yr2)" if use_buydown_2_1 else None
    annual_cashflow_label_y1 = "Annual Cash Flow (Yr1)" if use_buydown_2_1 else "Annual Cash Flow"
    annual_cashflow_label_y2 = "Annual Cash Flow (Yr2)" if use_buydown_2_1 else None
    dcr_label_y1 = "DCR (Yr1)" if use_buydown_2_1 else "DCR"
    break_even_label_y1 = "Break-Even Occupancy % (Yr1)" if use_buydown_2_1 else "Break-Even Occupancy %"
    cash_on_cash_label_y1 = "Cash on Cash Return (Yr1)" if use_buydown_2_1 else "Cash on Cash Return"

    metrics = {
        "Loan Amount": loan_amount,
        "Monthly Payment (Note Rate)": monthly_note_payment,
        borrower_payment_label_y1: borrower_pmt_y1 if use_buydown_2_1 else monthly_note_payment,
        effective_rate_label_y1: eff_rate_y1 if use_buydown_2_1 else annual_rate,
        cash_on_cash_label_y1: (annual_cash_flow_y1 / initial_investment) if initial_investment > 0 else None,
        break_even_label_y1: break_even_occ,    # fraction (0-1+)
        "Annual IRR (term sale)": annual_irr_term,
        "Hold Period (years)": hold_years,
        "Sale Price (hold)": sale_price,
        "Loan Balance at Hold": loan_balance_at_hold,
        "Net Sale Proceeds": net_sale_proceeds,
        "Total Cash Flow (hold)": total_cash_flow_hold,
        "Equity at Hold": equity_at_hold,
        "Total Profit (hold)": total_profit,
        "ROI (hold period)": roi_hold,
        "IRR (hold period)": hold_irr_annual,
        "Initial Investment (cash)": initial_investment,
        "Buy Closing Costs Paid By": buy_cc_paid_by,
        # KPI-only metrics (will be filtered out of main table)
        "NOI (annual, Yr1 levels)": noi,
        "Cap Rate": (noi / purchase_price) if purchase_price else None,
        "DCR": dcr_y1,
        "Cash Flow (Yr1)": annual_cash_flow_y1,
    }
    
    # Add Year 2 buydown-specific metrics if applicable
    if use_buydown_2_1 and nper > 12:
        metrics[borrower_payment_label_y2] = borrower_pmt_y2
        metrics[effective_rate_label_y2] = eff_rate_y2
        metrics["Buydown Subsidy (Yr1)"] = subsidy_y1
        metrics["Buydown Subsidy (Yr2)"] = subsidy_y2
        metrics["Buydown Subsidy (Total 2 yrs)"] = subsidy_total
        metrics["Buydown Paid By"] = buydown_paid_by
        # KPI-only Year 2 metrics
        metrics["Cash Flow (Yr2)"] = annual_cash_flow_y2
    
    return metrics, amort, yearly_cf_df

# -------------------------
# Excel export
# -------------------------
def dict_to_df_rowwise(d: dict, title: str):
    return pd.DataFrame(list(d.items()), columns=[title, "Value"])

def dataframes_to_xlsx_bytes(inputs: dict, metrics: dict, amort: pd.DataFrame, yearly_cf: pd.DataFrame, dscr_df: pd.DataFrame):
    tmp = BytesIO()
    with pd.ExcelWriter(tmp, engine="openpyxl") as writer:
        dict_to_df_rowwise(inputs, "Input").to_excel(writer, sheet_name="Inputs", index=False)
        dict_to_df_rowwise(metrics, "Metric").to_excel(writer, sheet_name="Metrics", index=False)
        amort.to_excel(writer, sheet_name="Amortization", index=False)
        yearly_cf.to_excel(writer, sheet_name="Yearly Cashflows", index=False)
        dscr_df.to_excel(writer, sheet_name="DSCR Trend", index=False)
    tmp.seek(0)

    wb = load_workbook(tmp)

    currency_fmt = '"$"#,##0.00'
    percent_fmt  = '0.00%'
    number2_fmt  = '#,##0.00'

    # Red negatives across all sheets
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, (int, float)) and cell.value < 0:
                    cell.font = Font(color="FF0000")

    # Inputs sheet number formats
    try:
        ws_in = wb["Inputs"]
        for r in range(2, ws_in.max_row + 1):
            label = str(ws_in.cell(row=r, column=1).value or "")
            vcell = ws_in.cell(row=r, column=2)
            val = vcell.value
            if not isinstance(val, (int, float)):
                continue
            if "%" in label:
                vcell.number_format = percent_fmt
            elif any(k in label for k in ["Price", "Payment", "Closing Costs", "Insurance", "Taxes",
                                          "HOA", "Rent", "Commission", "Monthly", "Sale Price", "Down Payment $"]):
                vcell.number_format = currency_fmt
            else:
                vcell.number_format = number2_fmt
    except Exception:
        pass

    # Metrics sheet: number formats + DCR/Break-even styling
    try:
        ws = wb["Metrics"]
        percent_keys = ["Cap Rate", "Cash on Cash", "IRR", "ROI", "Effective Rate", "Occupancy"]
        currency_keys = ["Price", "Payment", "Subsidy", "NOI", "Debt Service", "Cash Flow",
                         "Loan", "Proceeds", "Equity", "Profit", "Investment", "Closing Costs", "Commission", "Balance"]

        for r in range(2, ws.max_row + 1):
            mcell = ws.cell(row=r, column=1)
            vcell = ws.cell(row=r, column=2)
            metric_name = str(mcell.value) if mcell.value is not None else ""
            val = vcell.value

            if isinstance(val, (int, float)):
                if any(k in metric_name for k in percent_keys):
                    vcell.number_format = percent_fmt
                elif metric_name.startswith("DCR"):
                    vcell.number_format = number2_fmt
                elif any(k in metric_name for k in currency_keys):
                    vcell.number_format = currency_fmt
                else:
                    vcell.number_format = number2_fmt

            # Bold & conditional styling
            if metric_name.startswith("DCR"):
                mcell.font = Font(bold=True)
                vcell.font = Font(bold=True)
                if isinstance(val, (int, float)):
                    if val < 1.0:
                        vcell.font = Font(bold=True, color="FF0000")
                    elif val >= 1.25:
                        vcell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

            if metric_name.startswith("Break-Even Occupancy"):
                mcell.font = Font(bold=True)
                vcell.font = Font(bold=True)
                vcell.number_format = percent_fmt
                if isinstance(val, (int, float)) and val > 1.0:
                    vcell.font = Font(bold=True, color="FF0000")
    except Exception:
        pass

    # Yearly Cashflows sheet formats
    try:
        ws_cf = wb["Yearly Cashflows"]
        money_cols = [
            "Gross Rent (after vacancy) /mo", "Mgmt Fee /mo",
            "Other OpEx /mo (excl taxes/ins/HOA)", "Taxes /mo", "Insurance /mo", "HOA /mo",
            "Total OpEx /mo", "NOI (annual, prorated)",
            "Debt Service (annual, prorated)", "Cash Flow (annual)", "Cumulative Cash Flow"
        ]
        headers = {ws_cf.cell(row=1, column=c).value: c for c in range(1, ws_cf.max_column + 1)}
        for name in money_cols:
            if name in headers:
                col = headers[name]
                for r in range(2, ws_cf.max_row + 1):
                    cell = ws_cf.cell(row=r, column=col)
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = currency_fmt
        for name in ["Year", "Months Counted"]:
            if name in headers:
                col = headers[name]
                for r in range(2, ws_cf.max_row + 1):
                    cell = ws_cf.cell(row=r, column=col)
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '0'
    except Exception:
        pass

    # Amortization sheet formats
    try:
        ws_am = wb["Amortization"]
        headers = {ws_am.cell(row=1, column=c).value: c for c in range(1, ws_am.max_column + 1)}
        money_cols = ["Note Payment", "Borrower Payment", "Subsidy", "Principal", "Interest", "Balance", "Equity"]
        for name in money_cols:
            if name in headers:
                col = headers[name]
                for r in range(2, ws_am.max_row + 1):
                    cell = ws_am.cell(row=r, column=col)
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = currency_fmt
        if "Month" in headers:
            col = headers["Month"]
            for r in range(2, ws_am.max_row + 1):
                cell = ws_am.cell(row=r, column=col)
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '0'
    except Exception:
        pass

    # DSCR Trend sheet formats + conditional styling
    try:
        ws_d = wb["DSCR Trend"]
        headers = {ws_d.cell(row=1, column=c).value: c for c in range(1, ws_d.max_column + 1)}
        if "NOI (annual, prorated)" in headers:
            for r in range(2, ws_d.max_row + 1):
                cell = ws_d.cell(row=r, column=headers["NOI (annual, prorated)"])
                if isinstance(cell.value, (int, float)): cell.number_format = currency_fmt
        if "Debt Service (annual, prorated)" in headers:
            for r in range(2, ws_d.max_row + 1):
                cell = ws_d.cell(row=r, column=headers["Debt Service (annual, prorated)"])
                if isinstance(cell.value, (int, float)): cell.number_format = currency_fmt
        if "DCR" in headers:
            for r in range(2, ws_d.max_row + 1):
                cell = ws_d.cell(row=r, column=headers["DCR"])
                if isinstance(cell.value, (int, float)):
                    cell.number_format = number2_fmt
                    if cell.value < 1.0:
                        cell.font = Font(color="FF0000", bold=True)
                    elif cell.value >= 1.25:
                        cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                        cell.font = Font(bold=True)
        if "Year" in headers:
            for r in range(2, ws_d.max_row + 1):
                cell = ws_d.cell(row=r, column=headers["Year"])
                if isinstance(cell.value, (int, float)): cell.number_format = '0'
        for c in range(1, ws_d.max_column + 1):
            ws_d.cell(row=1, column=c).font = Font(bold=True)
    except Exception:
        pass

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()

# -------------------------
# Streamlit UI
# -------------------------
col1, col2, col3 = st.columns([1, 1, 1])

with col1:
    dp_mode = st.radio("Down Payment Mode", ["Percent of Price", "Dollar Amount"], horizontal=True)
    purchase_price = st.number_input("Purchase Price ($)", value=300000.0, min_value=0.0, step=1000.0)
    if dp_mode == "Percent of Price":
        down_payment_pct = st.number_input("Down Payment (% of Price)", value=20.0, min_value=0.0, max_value=100.0, step=1.0)
        down_payment = purchase_price * (down_payment_pct / 100.0)
    else:
        down_payment = st.number_input("Down Payment ($)", value=60000.0, min_value=0.0, step=1000.0)
        down_payment_pct = (down_payment / purchase_price * 100.0) if purchase_price > 0 else 0.0
    st.markdown(f"**Down Payment:** ${down_payment:,.2f}  •  **{down_payment_pct:.2f}%**")
    st.markdown(f"**Loan Amount:** ${purchase_price - down_payment:,.2f}")

with col2:
    interest_rate_pct = st.number_input("Interest Rate (annual %)", value=6.5, min_value=0.0, step=0.001, format="%.3f")
    loan_term_years = st.number_input("Loan Term (years)", value=30, min_value=0, step=1)
    override_payment = st.number_input("Monthly Loan Payment (optional)", value=0.0, min_value=0.0, step=50.0)
    use_buydown_2_1 = st.checkbox("Use 2/1 Buydown (Yr1 = note%-2, Yr2 = note%-1)")
    if use_buydown_2_1:
        buydown_paid_by = st.radio("Buydown escrow paid by", ["Buyer", "Seller", "Lender Credit"], horizontal=True)
    else:
        buydown_paid_by = "Buyer"

with col3:
    freq = st.radio("Rent & Expenses Frequency (excluding taxes/insurance)", ["Monthly", "Yearly"], horizontal=True)

# --- Buy-side closing costs payer ---
buy_cc_paid_by = st.radio("Buy-side Closing Costs Paid By", ["Buyer", "Seller", "Lender Credit"], horizontal=True)

# --- Units & rents ---
st.subheader("🏘️ Units")
num_units = st.number_input("Number of Units", min_value=1, value=1, step=1)
unit_rents = []
for i in range(int(num_units)):
    if freq == "Monthly":
        rent = st.number_input(f"Unit {i+1} Rent ($/month)", value=1000.0, min_value=0.0, step=50.0, key=f"unit_{i}_rent_monthly")
    else:
        rent = st.number_input(f"Unit {i+1} Rent ($/year)", value=12000.0, min_value=0.0, step=500.0, key=f"unit_{i}_rent_yearly") / 12.0
    unit_rents.append(rent)
monthly_rent = float(sum(unit_rents))

if freq == "Monthly":
    monthly_expenses = st.number_input("Other Monthly Expenses ($)", value=500.0, min_value=0.0, step=50.0)
    monthly_hoa = st.number_input("Monthly HOA ($)", value=0.0, min_value=0.0, step=25.0)
else:
    yearly_expenses = st.number_input("Other Yearly Expenses ($)", value=6000.0, min_value=0.0, step=500.0)
    yearly_hoa = st.number_input("Yearly HOA ($)", value=0.0, min_value=0.0, step=100.0)
    monthly_expenses = yearly_expenses / 12.0
    monthly_hoa = yearly_hoa / 12.0

st.subheader("🏦 Annual Fixed Costs")
annual_taxes = st.number_input("Annual Property Taxes ($)", value=3600.0, min_value=0.0, step=500.0)
annual_insurance = st.number_input("Annual Insurance ($)", value=1200.0, min_value=0.0, step=100.0)
monthly_taxes = annual_taxes / 12.0
monthly_insurance = annual_insurance / 12.0

st.subheader("⚙️ Vacancy & Management")
vacancy_pct = st.number_input("Vacancy Rate (% of Gross Rent)", value=5.0, min_value=0.0, max_value=100.0, step=0.5)
mgmt_pct = st.number_input("Management Fee (% of Collected Rent)", value=8.0, min_value=0.0, max_value=100.0, step=0.5)

st.subheader("📈 Rent Growth")
rent_appreciation_pct = st.number_input("Annual Rental Appreciation Rate (%)", value=2.0, min_value=-100.0, max_value=100.0, step=0.25)

st.subheader("🏠 Appreciation & Transaction Costs")

# --- Closing Costs at Purchase ($): disabled when Seller/Lender pays ---
if "buy_cc" not in st.session_state:
    st.session_state.buy_cc = 5000.0  # editable default when Buyer pays

if buy_cc_paid_by != "Buyer":
    st.session_state.buy_cc = 0.0

buy_closing_costs = st.number_input(
    "Closing Costs at Purchase ($)",
    min_value=0.0,
    step=100.0,
    key="buy_cc",
    disabled=(buy_cc_paid_by != "Buyer")
)

sale_mode = st.radio("Sale Price Input", ["Use Annual Appreciation Rate (%)", "Enter Flat Sale Amount"])
if sale_mode == "Use Annual Appreciation Rate (%)":
    appreciation_pct = st.number_input("Annual Property Appreciation Rate (%)", value=3.0, min_value=-100.0, max_value=100.0, step=0.25)
    sale_price_override = 0.0
else:
    appreciation_pct = st.number_input("Annual Property Appreciation Rate (%)", value=0.0, min_value=-100.0, max_value=100.0, step=0.25, help="Ignored when Flat Sale Amount is provided.")
    sale_price_override = st.number_input("Flat Sale Price at Hold ($)", value=0.0, min_value=0.0, step=1000.0)

sale_closing_costs = st.number_input("Closing Costs at Sale ($)", value=5000.0, min_value=0.0, step=100.0)
sale_commission_pct = st.number_input("Sale Commission (% of Sale Price)", value=6.0, min_value=0.0, max_value=20.0, step=0.25)

st.subheader("📅 Hold Period Analysis")
hold_years = st.number_input("Hold Period (years)", value=5, min_value=1, max_value=loan_term_years if loan_term_years > 0 else 50, step=1)

# -------------------------
# Live buydown preview under the rate input
# -------------------------
loan_amount_ui = max(0.0, purchase_price - down_payment)
nper_ui = int(round(loan_term_years * 12))
note_rate_ui = interest_rate_pct / 100.0
if use_buydown_2_1 and nper_ui > 0 and loan_amount_ui > 0:
    eff1, eff2, p1, p2, note_pmt_ui, sub1, sub2, sub_total = _buydown_preview(loan_amount_ui, note_rate_ui, nper_ui)
    st.caption(f"Yr1 effective rate: **{eff1*100:.2f}%** • Borrower payment: **${p1:,.2f}**")
    st.caption(f"Yr2 effective rate: **{eff2*100:.2f}%** • Borrower payment: **${p2:,.2f}**")
    st.caption(f"Note-rate payment (amortization basis): **${note_pmt_ui:,.2f}**")
    st.info(f"Buydown escrow (lump sum at closing): **${sub_total:,.0f}**  (Yr1: ${sub1:,.0f}, Yr2: ${sub2:,.0f}).")
else:
    if nper_ui > 0 and loan_amount_ui > 0:
        note_pmt_ui = pmt(note_rate_ui/12.0, nper_ui, loan_amount_ui)
        st.caption(f"Note-rate payment: **${note_pmt_ui:,.2f}**")

# -------------------------
# Run calculation
# -------------------------
if st.button("Calculate"):
    annual_rate = interest_rate_pct / 100.0

    # Warn on negative amortization if override_payment is too small
    if override_payment and override_payment > 0:
        first_month_interest = (annual_rate / 12.0) * max(0.0, purchase_price - down_payment)
        if override_payment + 1e-6 < first_month_interest:
            st.warning(
                f"⚠️ Your override payment (${override_payment:,.2f}) is below the first month's interest "
                f"(${first_month_interest:,.2f}). The balance will grow (negative amortization)."
            )

    metrics, amort, yearly_cf = calculate_metrics(
        purchase_price, down_payment, annual_rate, loan_term_years,
        monthly_rent, monthly_expenses, monthly_taxes, monthly_insurance, monthly_hoa,
        vacancy_pct=vacancy_pct, mgmt_pct=mgmt_pct, hold_years=hold_years,
        rent_appreciation_pct=rent_appreciation_pct,
        appreciation_pct=appreciation_pct, buy_closing_costs=buy_closing_costs,
        sale_closing_costs=sale_closing_costs, sale_commission_pct=sale_commission_pct,
        override_payment=override_payment, use_buydown_2_1=use_buydown_2_1,
        buydown_paid_by=buydown_paid_by, sale_price_override=sale_price_override,
        buy_cc_paid_by=buy_cc_paid_by
    )

    inputs = {
        "Purchase Price": purchase_price,
        "Down Payment %": down_payment_pct / 100.0,
        "Down Payment $": down_payment,
        "Loan Term (years)": loan_term_years,
        "Interest Rate % (note)": interest_rate_pct / 100.0,
        "Use 2/1 Buydown": use_buydown_2_1,
        "Buydown Paid By": buydown_paid_by if use_buydown_2_1 else "N/A",
        "Buy-side CC Paid By": buy_cc_paid_by,
        "Closing Costs (Buy)": buy_closing_costs,
        "Override Payment": override_payment,
        "Annual Taxes": annual_taxes,
        "Annual Insurance": annual_insurance,
        "Other Monthly Expenses": monthly_expenses,
        "Monthly HOA": monthly_hoa,
        "Vacancy Rate %": vacancy_pct / 100.0,
        "Management Fee %": mgmt_pct / 100.0,
        "Annual Rent Appreciation %": rent_appreciation_pct / 100.0,
        "Sale Price Mode": sale_mode,
        "Annual Property Appreciation %": appreciation_pct / 100.0,
        "Flat Sale Price (Override)": sale_price_override,
        "Closing Costs (Sale)": sale_closing_costs,
        "Sale Commission %": sale_commission_pct / 100.0,
        "Hold Period (years)": hold_years,
        "Number of Units": num_units,
        "Monthly Rent Total": monthly_rent
    }

    # --- Key Performance Tiles ---
    st.subheader("🎯 Key Performance Indicators")
    
    if use_buydown_2_1:
        # Show Year 1 and Year 2 side by side for buydown (5 columns)
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            noi_val = metrics.get("NOI (annual, Yr1 levels)", 0)  # Keep NOI in metrics for KPI access
            st.metric(
                label="📈 Annual NOI (Yr1)",
                value=f"${noi_val:,.0f}",
                delta=None
            )
            
        with col2:
            cashflow_val = metrics.get("Cash Flow (Yr1)", 0)  # Use direct key for KPI access
            cashflow_delta = "Positive" if cashflow_val > 0 else "Negative" if cashflow_val < 0 else "Break-even"
            cashflow_color = "normal" if cashflow_val >= 0 else "inverse"
            st.metric(
                label="💰 Cash Flow (Yr1)",
                value=f"${cashflow_val:,.0f}",
                delta=cashflow_delta,
                delta_color=cashflow_color
            )
            
        with col3:
            cashflow_yr2_val = metrics.get("Cash Flow (Yr2)", 0)
            if cashflow_yr2_val is not None:
                cashflow_yr2_delta = "Positive" if cashflow_yr2_val > 0 else "Negative" if cashflow_yr2_val < 0 else "Break-even"
                cashflow_yr2_color = "normal" if cashflow_yr2_val >= 0 else "inverse"
                st.metric(
                    label="💰 Cash Flow (Yr2)",
                    value=f"${cashflow_yr2_val:,.0f}",
                    delta=cashflow_yr2_delta,
                    delta_color=cashflow_yr2_color
                )
            else:
                st.metric(
                    label="💰 Cash Flow (Yr2)",
                    value="N/A"
                )
            
        with col4:
            # Use dynamic label to get the right DCR metric
            dcr_key = "DCR (Yr1)" if use_buydown_2_1 else "DCR"
            dcr_val = metrics.get(dcr_key, 0)
            if dcr_val and dcr_val > 0:
                if dcr_val >= 1.25:
                    dcr_status = "Excellent"
                    dcr_color = "normal"
                elif dcr_val >= 1.0:
                    dcr_status = "Good"
                    dcr_color = "normal"
                else:
                    dcr_status = "Risky"
                    dcr_color = "inverse"
            else:
                dcr_status = "N/A"
                dcr_color = "off"
                
            st.metric(
                label="🏦 Debt Coverage Ratio",
                value=f"{dcr_val:.2f}" if dcr_val else "N/A",
                delta=dcr_status,
                delta_color=dcr_color
            )
            
        with col5:
            # Cap Rate tile
            cap_rate_val = metrics.get("Cap Rate", 0)
            if cap_rate_val and cap_rate_val > 0:
                cap_rate_pct = cap_rate_val * 100  # Convert to percentage
                if cap_rate_pct >= 8.0:
                    cap_status = "Strong"
                    cap_color = "normal"
                elif cap_rate_pct >= 5.0:
                    cap_status = "Good"
                    cap_color = "normal"
                else:
                    cap_status = "Low"
                    cap_color = "inverse"
            else:
                cap_rate_pct = 0
                cap_status = "N/A"
                cap_color = "off"
                
            st.metric(
                label="🎯 Cap Rate",
                value=f"{cap_rate_pct:.2f}%" if cap_rate_val else "N/A",
                delta=cap_status,
                delta_color=cap_color
            )
    else:
        # Standard 4-column layout for non-buydown (including Cap Rate)
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            noi_val = metrics.get("NOI (annual, Yr1 levels)", 0)
            noi_color = "normal" if noi_val > 0 else "inverse"
            st.metric(
                label="📈 Annual NOI (Yr1)",
                value=f"${noi_val:,.0f}",
                delta=None
            )
            
        with col2:
            # Use KPI key for cash flow
            cashflow_val = metrics.get("Cash Flow (Yr1)", 0)
            cashflow_delta = "Positive" if cashflow_val > 0 else "Negative" if cashflow_val < 0 else "Break-even"
            cashflow_color = "normal" if cashflow_val >= 0 else "inverse"
            st.metric(
                label="💰 Annual Cash Flow",
                value=f"${cashflow_val:,.0f}",
                delta=cashflow_delta,
                delta_color=cashflow_color
            )
            
        with col3:
            # Use dynamic label to get the right DCR metric
            dcr_key = "DCR"  # Generic when not buydown
            dcr_val = metrics.get(dcr_key, 0)
            if dcr_val and dcr_val > 0:
                if dcr_val >= 1.25:
                    dcr_status = "Excellent"
                    dcr_color = "normal"
                elif dcr_val >= 1.0:
                    dcr_status = "Good"
                    dcr_color = "normal"
                else:
                    dcr_status = "Risky"
                    dcr_color = "inverse"
            else:
                dcr_status = "N/A"
                dcr_color = "off"
                
            st.metric(
                label="🏦 Debt Coverage Ratio",
                value=f"{dcr_val:.2f}" if dcr_val else "N/A",
                delta=dcr_status,
                delta_color=dcr_color
            )
            
        with col4:
            # Cap Rate tile
            cap_rate_val = metrics.get("Cap Rate", 0)
            if cap_rate_val and cap_rate_val > 0:
                cap_rate_pct = cap_rate_val * 100  # Convert to percentage
                if cap_rate_pct >= 8.0:
                    cap_status = "Strong"
                    cap_color = "normal"
                elif cap_rate_pct >= 5.0:
                    cap_status = "Good"
                    cap_color = "normal"
                else:
                    cap_status = "Low"
                    cap_color = "inverse"
            else:
                cap_rate_pct = 0
                cap_status = "N/A"
                cap_color = "off"
                
            st.metric(
                label="🎯 Cap Rate",
                value=f"{cap_rate_pct:.2f}%" if cap_rate_val else "N/A",
                delta=cap_status,
                delta_color=cap_color
            )
    
    st.divider()

    # --- Key Metrics ---
    st.subheader("📊 Key Metrics & Predictions")
    display_rows = []
    red_metrics = set()

    # Use dynamic labels for red metrics detection
    be_key = "Break-Even Occupancy % (Yr1)" if use_buydown_2_1 else "Break-Even Occupancy %"
    be_val = metrics.get(be_key)
    if isinstance(be_val, (int, float)) and be_val > 1.0:
        red_metrics.add(be_key)

    # KPI-only metrics to exclude from main table
    kpi_only_metrics = {
        "NOI (annual, Yr1 levels)", "Cap Rate", "DCR", 
        "Cash Flow (Yr1)", "Cash Flow (Yr2)"
    }
    
    # Investment analysis metrics to display separately
    investment_analysis_metrics = {
        "Annual IRR (term sale)"
    }
    
    # Prediction metrics for future projections
    prediction_metrics = {
        "Hold Period (years)", "Sale Price (hold)", "Loan Balance at Hold",
        "Net Sale Proceeds", "Total Cash Flow (hold)", "Equity at Hold",
        "Total Profit (hold)", "ROI (hold period)", "IRR (hold period)",
        "Initial Investment (cash)", "Buy Closing Costs Paid By"
    }

    # Main metrics for the primary table
    main_display_rows = []
    investment_display_rows = []
    prediction_display_rows = []

    for k, v in metrics.items():
        # Skip None values - don't add them to the display
        if v is None:
            continue
        # Skip KPI-only metrics from main table
        if k in kpi_only_metrics:
            continue
            
        if isinstance(v, (int, float)) and v < 0:
            red_metrics.add(k)
            
        # Split into main vs investment analysis vs predictions
        if k in prediction_metrics:
            prediction_display_rows.append({"Metric": k, "Value": _fmt_metric(k, v)})
        elif k in investment_analysis_metrics:
            investment_display_rows.append({"Metric": k, "Value": _fmt_metric(k, v)})
        else:
            main_display_rows.append({"Metric": k, "Value": _fmt_metric(k, v)})
            
    main_display_df = pd.DataFrame(main_display_rows, columns=["Metric", "Value"])
    investment_display_df = pd.DataFrame(investment_display_rows, columns=["Metric", "Value"])
    prediction_display_df = pd.DataFrame(prediction_display_rows, columns=["Metric", "Value"])

    def _style_metrics_row(row):
        m = row["Metric"]
        val_style = "color:red;" if m in red_metrics else ""
        return ["", val_style]

    st.write(main_display_df.style.apply(_style_metrics_row, axis=1))

    # --- Investment Analysis ---
    if not investment_display_df.empty:
        st.subheader("💼 Investment Analysis")
        st.write(investment_display_df.style.apply(_style_metrics_row, axis=1))

    # --- Future Projections & Hold Period ---
    st.subheader("🔮 Future Projections & Hold Period")
    st.write(prediction_display_df.style.apply(_style_metrics_row, axis=1))

    # --- Year-by-Year Cashflows (formatted) ---
    st.subheader("📅 Yearly Cashflows (uses Annual Rental Appreciation %)")
    YCF_MONEY_COLS = [
        "Gross Rent (after vacancy) /mo", "Mgmt Fee /mo",
        "Other OpEx /mo (excl taxes/ins/HOA)", "Taxes /mo", "Insurance /mo", "HOA /mo",
        "Total OpEx /mo", "NOI (annual, prorated)",
        "Debt Service (annual, prorated)", "Cash Flow (annual)", "Cumulative Cash Flow"
    ]
    ycf_format = {c: "${:,.2f}" for c in YCF_MONEY_COLS if c in yearly_cf.columns}
    ycf_format.update({"Year": "{:.0f}", "Months Counted": "{:.0f}"})
    st.write(
        yearly_cf.style
            .format(ycf_format)
            .applymap(_style_red_neg, subset=[c for c in yearly_cf.columns if c in YCF_MONEY_COLS])
    )

    # --- DSCR Trend (table only) ---
    st.subheader("📈 Debt Coverage Ratio (DCR) Trend by Year")
    if not yearly_cf.empty:
        dscr_df = yearly_cf[["Year", "NOI (annual, prorated)", "Debt Service (annual, prorated)"]].copy()
        dscr_df["DCR"] = dscr_df.apply(
            lambda r: (r["NOI (annual, prorated)"] / r["Debt Service (annual, prorated)"]) if r["Debt Service (annual, prorated)"] != 0 else np.nan,
            axis=1
        )
        st.write(
            dscr_df.style.format({
                "Year": "{:.0f}",
                "NOI (annual, prorated)": "${:,.2f}",
                "Debt Service (annual, prorated)": "${:,.2f}",
                "DCR": "{:.2f}"
            })
        )
    else:
        dscr_df = pd.DataFrame(columns=["Year", "NOI (annual, prorated)", "Debt Service (annual, prorated)", "DCR"])
        st.info("No yearly cashflow rows to display.")

    # --- Amortization (formatted) ---
    st.subheader("📑 Amortization Schedule (Borrower vs Note) — first 200 rows")
    amort_view = amort.head(200)
    AMORT_MONEY_COLS = ["Note Payment", "Borrower Payment", "Subsidy", "Principal", "Interest", "Balance", "Equity"]
    amort_format = {c: "${:,.2f}" for c in AMORT_MONEY_COLS if c in amort_view.columns}
    amort_format.update({"Month": "{:.0f}"})
    st.write(
        amort_view.style
            .format(amort_format)
            .applymap(_style_red_neg, subset=AMORT_MONEY_COLS)
    )

    # --- Excel download ---
    excel_bytes = dataframes_to_xlsx_bytes(inputs, metrics, amort, yearly_cf, dscr_df)
    st.download_button(
        "📥 Download Excel",
        data=excel_bytes,
        file_name="roi_amortization.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# --- Disclaimer Footnote ---
st.divider()
st.markdown("""
<div style='text-align: center; color: #666; font-size: 0.9em; padding: 15px; border: 1px solid #ddd; background-color: #f9f9f9; margin: 10px;'>
    <strong>⚠️ DISCLAIMER:</strong> This calculator is for educational purposes only and provides estimates that may vary significantly from actual results. 
    Always consult qualified financial, legal, and real estate professionals before making investment decisions. 
    Not liable for any damages from use of this tool.
</div>
""", unsafe_allow_html=True)
