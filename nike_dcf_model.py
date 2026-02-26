"""
NKE Full DCF & Comparable Company Valuation
==============================================================================
Author: Krish Pathak
Date: 2025

Model Structure:
    1.  Data Ingestion        — Historical financials from Nike_Financials.xlsx
    2.  WACC Build-Up         — CAPM-based, fully sourced
    3.  Three-Statement Projections — IS, BS, CFS linked
    4.  Unlevered FCF         — NOPAT + D&A - CapEx - DNWC
    5.  DCF Valuation         — Gordon Growth terminal value
    6.  Comparable Companies  — Peer EV/EBITDA & P/E multiples
    7.  Equity Bridge         — EV to implied share price
    8.  Sensitivity Analysis  — WACC x Terminal Growth heatmap
    9.  Scenario Analysis     — Bear / Base / Bull cases
    10. Football Field        — Valuation range summary
    11. Excel Export          — Cover sheet, colour-coded, fully formatted
"""

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import ColorScaleRule
from datetime import datetime
import warnings
warnings.filterwarnings("ignore")


# =============================================================================
# 1. DATA INGESTION
# =============================================================================

EXCEL_PATH = "Nike_Financials.xlsx"

def load_sheet(sheet_name):
    df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name, index_col=0, header=None)
    years_row = df[df.index == "Report Date"].iloc[0]
    df.columns = years_row.values
    return df


def get_row(df, label):
    """
    Returns the first row matching label that contains numeric data.
    Handles duplicate labels caused by reporting format changes across years.
    """
    rows = df[df.index == label]
    if rows.empty:
        raise KeyError(f"Row not found in financial data: '{label}'")
    combined = pd.Series(index=rows.columns, dtype=float)
    for col in rows.columns:
        for i in range(len(rows)):
            val = pd.to_numeric(rows.iloc[i][col], errors="coerce")
            if pd.notna(val):
                combined[col] = val
                break
    return combined


is_  = load_sheet("Nike_IS")
bs   = load_sheet("Nike_BS")
cf   = load_sheet("Nike_CF")

HIST_YEARS = [2020, 2021, 2022, 2023, 2024]

# Income Statement
revenue      = get_row(is_, "Revenues")[HIST_YEARS]
cogs         = get_row(is_, "Cost of sales")[HIST_YEARS]
gross_profit = get_row(is_, "Gross profit")[HIST_YEARS]
sga          = get_row(is_, "Total selling & administrative expense")[HIST_YEARS]
ebt          = get_row(is_, "Income (loss) before income taxes")[HIST_YEARS]
tax_exp      = get_row(is_, "Income tax expense (benefit)")[HIST_YEARS]
net_income   = get_row(is_, "Net income (loss)")[HIST_YEARS]

ebit         = ebt
ebitda_hist  = ebit + get_row(cf, "D&A")[HIST_YEARS]

# Cash Flow Statement
da           = get_row(cf, "D&A")[HIST_YEARS]
capex        = get_row(cf, "CapEx")[HIST_YEARS].abs()
cfo          = get_row(cf, "Net cash flows from operating activities")[HIST_YEARS]

# Balance Sheet
cash_hist    = get_row(bs, "Cash & equivalents")[HIST_YEARS]
ar_hist      = get_row(bs, "Accounts receivable, net")[HIST_YEARS]
inv_hist     = get_row(bs, "Inventories")[HIST_YEARS]
cur_assets   = get_row(bs, "Total current assets")[HIST_YEARS]
cur_liab     = get_row(bs, "Total current liabilities")[HIST_YEARS]

# AP has a label change in FY2024 — merge both rows for a complete series
ap_upper     = get_row(bs, "Accounts Payable")[HIST_YEARS]   # FY2024
ap_lower     = get_row(bs, "Accounts payable")[HIST_YEARS]   # FY2020-FY2023
ap_hist      = ap_upper.combine_first(ap_lower)

# Operating NWC = AR + Inventory - AP
nwc_hist     = ar_hist + inv_hist - ap_hist

# Balance sheet inputs for equity bridge (FY2024)
CASH_MM         = float(cash_hist[2024])                                          # Nike_BS, FY2024
LTD_CURRENT_MM  = 1000.0                                                          # Nike_BS, FY2024 — Current portion LTD
LTD_LONG_MM     = float(get_row(bs, "Long-term debt").dropna().iloc[0])           # Nike_BS, FY2023 (FY2024 format change)
TOTAL_DEBT_MM   = LTD_CURRENT_MM + LTD_LONG_MM
NET_DEBT_MM     = TOTAL_DEBT_MM - CASH_MM
MINORITY_INT_MM = 0.0                                                             # Nike_BS, FY2024 — none

SHARES_OUT_MM   = float(get_row(is_, "Year end shares outstanding")[2024])        # Nike_IS, FY2024


# =============================================================================
# 2. WACC BUILD-UP
# =============================================================================

RISK_FREE_RATE  = 0.0435   # US 10Y Treasury, May 2025 (FRED: DGS10)
ERP             = 0.0550   # Damodaran (NYU Stern), Jan 2025 Implied ERP
BETA            = 0.83     # Bloomberg 5Y monthly regression vs. S&P 500, May 2025
PRE_TAX_COD     = 0.035    # NKE 10-K FY2024, Note 9 — weighted average coupon rate
CURRENT_PRICE   = 75.50    # Yahoo Finance / Bloomberg, May 2025
MARKET_CAP_MM   = CURRENT_PRICE * SHARES_OUT_MM

EFF_TAX_RATE    = float((tax_exp / ebt).mean())
COST_OF_EQUITY  = RISK_FREE_RATE + BETA * ERP
COST_OF_DEBT_AT = PRE_TAX_COD * (1 - EFF_TAX_RATE)
EQUITY_WEIGHT   = MARKET_CAP_MM / (MARKET_CAP_MM + TOTAL_DEBT_MM)
DEBT_WEIGHT     = 1 - EQUITY_WEIGHT
WACC            = EQUITY_WEIGHT * COST_OF_EQUITY + DEBT_WEIGHT * COST_OF_DEBT_AT


# =============================================================================
# 3. HISTORICAL DRIVER RATIOS (5-year averages)
# =============================================================================

rev          = revenue.astype(float)
cogs_pct     = float((cogs.astype(float)         / rev).mean())
gm_pct       = float((gross_profit.astype(float) / rev).mean())
sga_pct      = float((sga.astype(float)          / rev).mean())
da_pct       = float((da.astype(float)           / rev).mean())
capex_pct    = float((capex.astype(float)        / rev).mean())
nwc_pct      = float((nwc_hist.astype(float)     / rev).mean())
eff_tax      = EFF_TAX_RATE


# =============================================================================
# 4. COMPARABLE COMPANIES
# Source: Bloomberg / FactSet Consensus — May 2025
# =============================================================================

comps = pd.DataFrame({
    "Company":       ["Adidas AG",  "Under Armour", "Lululemon", "On Holding", "Skechers"],
    "Ticker":        ["ADDYY",      "UAA",          "LULU",      "ONON",       "SKX"],
    "EV_EBITDA_NTM": [12.1,         8.3,            14.2,        24.5,         9.8],
    "PE_NTM":        [18.4,         16.2,           17.8,        45.3,         14.6],
    "Rev_Growth":    [0.062,        0.021,          0.112,       0.312,        0.088],
    "EBITDA_Margin": [0.098,        0.071,          0.221,       0.182,        0.121],
}).set_index("Company")

comp_ev_ebitda_median = comps["EV_EBITDA_NTM"].median()
comp_ev_ebitda_mean   = comps["EV_EBITDA_NTM"].mean()
comp_pe_median        = comps["PE_NTM"].median()
comp_pe_mean          = comps["PE_NTM"].mean()


# =============================================================================
# 5. SCENARIO DEFINITIONS & PROJECTION ENGINE
# =============================================================================

PROJ_YEARS = [2025, 2026, 2027, 2028, 2029]
N          = len(PROJ_YEARS)

scenarios = {
    "Bear": {
        "rev_growth":    [0.00,  0.01,  0.01,  0.02,  0.02],
        "gm_expansion":  -0.010,
        "sga_leverage":   0.000,
        "tv_growth":      0.020,
        "description":   "DTC slowdown, FX headwinds, wholesale channel weakness",
    },
    "Base": {
        "rev_growth":    [0.02,  0.03,  0.04,  0.05,  0.05],
        "gm_expansion":   0.005,
        "sga_leverage":   0.002,
        "tv_growth":      0.030,
        "description":   "Gradual recovery, DTC mix shift, China normalization",
    },
    "Bull": {
        "rev_growth":    [0.05,  0.07,  0.08,  0.08,  0.08],
        "gm_expansion":   0.015,
        "sga_leverage":   0.005,
        "tv_growth":      0.035,
        "description":   "DTC acceleration, strong China recovery, significant margin expansion",
    },
}


def project_scenario(name: str) -> dict:
    s            = scenarios[name]
    growth_rates = s["rev_growth"]
    gm_adj       = s["gm_expansion"]
    sga_adj      = s["sga_leverage"]
    tv_g         = s["tv_growth"]

    prev_rev = float(revenue[2024])
    prev_nwc = float(nwc_hist[2024])

    revenues, gross_profits, ebitdas, ebits  = [], [], [], []
    taxes, nopats, das, capexs               = [], [], [], []
    delta_nwcs, ufcfs                        = [], []

    for i, g in enumerate(growth_rates):
        rev_   = prev_rev * (1 + g)
        gm_    = rev_ * (gm_pct  + gm_adj  * (i + 1) / N)
        sg_    = rev_ * (sga_pct - sga_adj * (i + 1) / N)
        d_     = rev_ * da_pct
        ebitda_= gm_ - sg_
        ebit_  = ebitda_ - d_
        tax_   = ebit_ * eff_tax
        nopat_ = ebit_ - tax_
        cap_   = rev_ * capex_pct
        nwc_   = rev_ * nwc_pct
        dnwc   = nwc_ - prev_nwc
        ufcf   = nopat_ + d_ - cap_ - dnwc

        revenues.append(rev_);    gross_profits.append(gm_)
        ebitdas.append(ebitda_);  ebits.append(ebit_)
        taxes.append(tax_);       nopats.append(nopat_)
        das.append(d_);           capexs.append(cap_)
        delta_nwcs.append(dnwc);  ufcfs.append(ufcf)
        prev_rev, prev_nwc = rev_, nwc_

    disc_fcfs = [fcf / (1 + WACC) ** (i + 1) for i, fcf in enumerate(ufcfs)]
    tv        = ufcfs[-1] * (1 + tv_g) / (WACC - tv_g)
    disc_tv   = tv / (1 + WACC) ** N
    ev        = sum(disc_fcfs) + disc_tv

    equity_val = ev - NET_DEBT_MM - MINORITY_INT_MM
    implied_px = (equity_val * 1e6) / (SHARES_OUT_MM * 1e6)
    premium    = (implied_px / CURRENT_PRICE) - 1

    ev_comps   = comp_ev_ebitda_median * ebitdas[0]
    px_comps   = (ev_comps * 1e6 - NET_DEBT_MM * 1e6) / (SHARES_OUT_MM * 1e6)

    return {
        "revenues": revenues, "gross_profits": gross_profits,
        "ebitdas": ebitdas,   "ebits": ebits,
        "taxes": taxes,       "nopats": nopats,
        "das": das,           "capexs": capexs,
        "delta_nwcs": delta_nwcs, "ufcfs": ufcfs,
        "disc_fcfs": disc_fcfs, "sum_disc_fcfs": sum(disc_fcfs),
        "disc_tv": disc_tv,   "terminal_value": tv,
        "enterprise_value": ev, "equity_value": equity_val,
        "implied_price": implied_px, "premium_discount": premium,
        "px_comps_impl": px_comps,  "tv_pct_ev": disc_tv / ev,
        "growth_rates": growth_rates, "tv_growth": tv_g,
        "gm_margins":     [gp / r for gp, r in zip(gross_profits, revenues)],
        "ebitda_margins": [e  / r for e,  r in zip(ebitdas,       revenues)],
    }


results = {s: project_scenario(s) for s in scenarios}
base    = results["Base"]
base_px = base["implied_price"]
bear_px = results["Bear"]["implied_price"]
bull_px = results["Bull"]["implied_price"]


# =============================================================================
# 6. SENSITIVITY ANALYSIS — WACC x Terminal Growth
# =============================================================================

wacc_range = np.round(np.arange(0.07, 0.13, 0.01), 2)
tvg_range  = np.round(np.arange(0.01, 0.05, 0.005), 3)

sens_matrix = pd.DataFrame(index=wacc_range, columns=tvg_range, dtype=float)
for w in wacc_range:
    for g in tvg_range:
        if w <= g:
            sens_matrix.loc[w, g] = np.nan
            continue
        pv  = sum(fcf / (1 + w) ** (i + 1) for i, fcf in enumerate(base["ufcfs"]))
        tv_ = base["ufcfs"][-1] * (1 + g) / (w - g)
        ev_ = pv + tv_ / (1 + w) ** N
        eq_ = ev_ - NET_DEBT_MM
        sens_matrix.loc[w, g] = round((eq_ * 1e6) / (SHARES_OUT_MM * 1e6), 2)


# =============================================================================
# 7. FOOTBALL FIELD RANGES
# =============================================================================

ev_low    = (comp_ev_ebitda_mean - 1.5) * base["ebitdas"][0]
ev_high   = (comp_ev_ebitda_mean + 1.5) * base["ebitdas"][0]
comp_low  = (ev_low  * 1e6 - NET_DEBT_MM * 1e6) / (SHARES_OUT_MM * 1e6)
comp_high = (ev_high * 1e6 - NET_DEBT_MM * 1e6) / (SHARES_OUT_MM * 1e6)

football_field = {
    "DCF (Bear to Bull)":    (bear_px,  bull_px),
    "EV/EBITDA Comps Range": (comp_low, comp_high),
    "52-Week Trading Range": (70.00,    115.00),
}


# =============================================================================
# 8. EXCEL FORMATTING HELPERS
# =============================================================================

OUTPUT = "Nike_GS_DCF_Model.xlsx"
wb     = Workbook()
wb.remove(wb.active)

FONT_NAME = "Arial"

def _f(bold=False, sz=10, color="000000", italic=False):
    return Font(name=FONT_NAME, bold=bold, size=sz, color=color, italic=italic)

def _fill(hex_c):
    return PatternFill("solid", start_color=hex_c, fgColor=hex_c)

def _align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _border():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)

NAVY      = "003366"
MID_BLUE  = "D9E1F2"
LIGHT_BLU = "EBF1FA"
GOLD      = "C9A84C"
WHITE     = "FFFFFF"
ALT       = "F5F7FB"
SECT_BG   = "1F3864"

C_INPUT   = "0000FF"   # Blue  — hardcoded inputs
C_CALC    = "000000"   # Black — derived / formula values
C_LINK    = "008000"   # Green — key output values

F_USD  = '#,##0;(#,##0);"-"'
F_PCT  = '0.0%;(0.0%);"-"'
F_PCT2 = '0.00%;(0.00%);"-"'
F_MULT = '0.0x'
F_PX   = '$#,##0.00'
F_4DP  = '0.0000'


def col_w(ws, col, w):
    ws.column_dimensions[get_column_letter(col)].width = w


def hdr_row(ws, row, labels, c0=1, bg=NAVY):
    for i, lbl in enumerate(labels):
        c = ws.cell(row=row, column=c0 + i, value=lbl)
        c.font = _f(bold=True, color=WHITE); c.fill = _fill(bg)
        c.alignment = _align("center");      c.border = _border()


def sect(ws, row, col, title, span=1):
    c = ws.cell(row=row, column=col, value=title)
    c.font = _f(bold=True, color=WHITE, sz=10); c.fill = _fill(SECT_BG)
    c.alignment = _align("left")
    if span > 1:
        ws.merge_cells(start_row=row, start_column=col,
                       end_row=row,   end_column=col + span - 1)


def data_row(ws, row, label, values, fmt=F_USD, bold=False, alt=False, tc=C_CALC):
    bg = ALT if alt else WHITE
    lc = ws.cell(row=row, column=1, value=label)
    lc.font = _f(bold=bold, color=tc); lc.fill = _fill(bg)
    lc.alignment = _align("left");    lc.border = _border()
    for j, v in enumerate(values):
        dc = ws.cell(row=row, column=2 + j, value=v)
        dc.number_format = fmt
        dc.font = _f(bold=bold, color=tc); dc.fill = _fill(bg)
        dc.alignment = _align("right");   dc.border = _border()


# =============================================================================
# SHEET 1 — COVER
# =============================================================================

ws_cv = wb.create_sheet("Cover")
ws_cv.sheet_view.showGridLines = False
ws_cv.column_dimensions["A"].width = 2
for c in range(2, 14):
    ws_cv.column_dimensions[get_column_letter(c)].width = 14

ws_cv.merge_cells("B2:M4")
cell = ws_cv["B2"]
cell.value = "NIKE, INC. (NKE)  —  EQUITY RESEARCH"
cell.font  = _f(bold=True, sz=22, color=WHITE)
cell.fill  = _fill(NAVY)
cell.alignment = _align("center", "center")
for r in range(2, 5):
    for c in range(2, 14):
        ws_cv.cell(row=r, column=c).fill = _fill(NAVY)
ws_cv.row_dimensions[2].height = 22
ws_cv.row_dimensions[3].height = 22
ws_cv.row_dimensions[4].height = 22

ws_cv.merge_cells("B5:M5")
cell = ws_cv["B5"]
cell.value = "Full DCF & Comparable Company Valuation  |  Analyst-Grade Financial Model"
cell.font  = _f(italic=True, sz=11, color=GOLD)
cell.fill  = _fill(NAVY)
cell.alignment = _align("center")
ws_cv.row_dimensions[5].height = 20

rating = "OUTPERFORM" if base_px > CURRENT_PRICE else "UNDERPERFORM"
boxes  = [
    ("Current Price",        f"${CURRENT_PRICE:.2f}",                 "B"),
    ("Implied Price (Base)", f"${base_px:.2f}",                       "E"),
    ("Upside / Downside",    f"{(base_px/CURRENT_PRICE-1)*100:.1f}%", "H"),
    ("Rating",               rating,                                   "K"),
]
for lbl, val, col_letter in boxes:
    c0 = ord(col_letter) - ord("A") + 1
    for r in [7, 8, 9]:
        ws_cv.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c0+2)
        for off in range(3):
            ws_cv.cell(row=r, column=c0+off).fill = _fill(LIGHT_BLU)
    lc = ws_cv.cell(row=7, column=c0, value=lbl)
    lc.font = _f(bold=True, sz=9, color=NAVY); lc.alignment = _align("center")
    vc = ws_cv.cell(row=8, column=c0, value=val)
    vc.font = _f(bold=True, sz=16, color=NAVY); vc.alignment = _align("center")

ws_cv.cell(row=11, column=2, value="SCENARIO SUMMARY").font = _f(bold=True, sz=11, color=NAVY)
hdr_row(ws_cv, 12, ["Scenario", "Rev CAGR (5Y)", "EBITDA Margin (Y5)",
                    "Implied Price", "Upside / Downside", "TV Growth Rate"], c0=2)
for i, (sc_name, sc_res) in enumerate(results.items()):
    r    = 13 + i
    cagr = (sc_res["revenues"][-1] / float(revenue[2024])) ** (1 / N) - 1
    bg   = ALT if i % 2 else WHITE
    vals = [sc_name, f"{cagr*100:.1f}%", f"{sc_res['ebitda_margins'][-1]*100:.1f}%",
            f"${sc_res['implied_price']:.2f}", f"{sc_res['premium_discount']*100:.1f}%",
            f"{sc_res['tv_growth']*100:.1f}%"]
    for j, v in enumerate(vals):
        c = ws_cv.cell(row=r, column=2+j, value=v)
        c.font = _f(bold=(j == 0)); c.fill = _fill(bg)
        c.alignment = _align("center" if j > 0 else "left"); c.border = _border()

r0 = 17
ws_cv.cell(row=r0, column=2, value="KEY MODEL ASSUMPTIONS").font = _f(bold=True, sz=10, color=NAVY)
notes = [
    f"WACC: {WACC*100:.2f}%   (Ke = {COST_OF_EQUITY*100:.2f}%,  Kd(at) = {COST_OF_DEBT_AT*100:.2f}%,  Beta = {BETA:.2f})",
    f"Risk-Free Rate: {RISK_FREE_RATE*100:.2f}%   Source: US 10Y Treasury, May 2025 (FRED: DGS10)",
    f"Equity Risk Premium: {ERP*100:.2f}%   Source: Damodaran (NYU Stern), Jan 2025 Implied ERP",
    f"Shares Outstanding: {SHARES_OUT_MM:.1f}mm   Source: Nike_IS, FY2024 Year-End",
    f"Net Debt: ${NET_DEBT_MM:,.0f}mm   Source: Nike_BS, FY2024 (Total Debt - Cash)",
    f"Effective Tax Rate: {EFF_TAX_RATE*100:.1f}%   Source: Nike_IS, FY2020-FY2024 average",
    f"Analysis Date: {datetime.today().strftime('%B %d, %Y')}",
]
for k, note in enumerate(notes):
    c = ws_cv.cell(row=r0+1+k, column=2, value=note)
    c.font = _f(sz=9, italic=True)
    ws_cv.merge_cells(start_row=r0+1+k, start_column=2, end_row=r0+1+k, end_column=13)

ws_cv.row_dimensions[1].height = 8


# =============================================================================
# SHEET 2 — WACC BUILD-UP
# =============================================================================

ws_w = wb.create_sheet("WACC Build-Up")
ws_w.sheet_view.showGridLines = False
col_w(ws_w, 1, 38); col_w(ws_w, 2, 14); col_w(ws_w, 3, 45)

sect(ws_w, 1, 1, "WACC BUILD-UP  —  NIKE, INC. (NKE)", 3)

wacc_rows = [
    ("COST OF EQUITY  (CAPM)", None, None, None, True),
    ("Risk-Free Rate (Rf)",                     RISK_FREE_RATE,   F_PCT2, C_INPUT),
    ("Equity Beta (Beta)",                      BETA,             "0.00", C_INPUT),
    ("Equity Risk Premium (ERP)",               ERP,              F_PCT2, C_INPUT),
    ("Cost of Equity  [ Ke = Rf + B x ERP ]",  COST_OF_EQUITY,   F_PCT2, C_CALC),
    ("", None, None, None, False),
    ("COST OF DEBT", None, None, None, True),
    ("Pre-Tax Cost of Debt",                    PRE_TAX_COD,      F_PCT2, C_INPUT),
    ("Effective Tax Rate",                      EFF_TAX_RATE,     F_PCT2, C_CALC),
    ("After-Tax Cost of Debt  [ Kd x (1-t) ]", COST_OF_DEBT_AT,  F_PCT2, C_CALC),
    ("", None, None, None, False),
    ("CAPITAL STRUCTURE", None, None, None, True),
    ("Market Capitalisation ($mm)",             MARKET_CAP_MM,    F_USD,  C_INPUT),
    ("Total Debt ($mm)",                        TOTAL_DEBT_MM,    F_USD,  C_INPUT),
    ("Equity Weight",                           EQUITY_WEIGHT,    F_PCT2, C_CALC),
    ("Debt Weight",                             DEBT_WEIGHT,      F_PCT2, C_CALC),
    ("", None, None, None, False),
    ("WEIGHTED AVERAGE COST OF CAPITAL", None, None, None, True),
    ("WACC  [ Ke x E% + Kd x D% ]",            WACC,             F_PCT2, C_CALC),
]

wacc_sources = {
    "Risk-Free Rate (Rf)":          "US 10Y Treasury, May 2025 (FRED: DGS10)",
    "Equity Beta (Beta)":            "Bloomberg 5Y monthly regression vs. S&P 500, May 2025",
    "Equity Risk Premium (ERP)":    "Damodaran (NYU Stern), Jan 2025 Implied ERP",
    "Pre-Tax Cost of Debt":         "NKE 10-K FY2024, Note 9 — Weighted average coupon rate",
    "Effective Tax Rate":            "Nike_IS, FY2020-FY2024 average effective tax rate",
    "Market Capitalisation ($mm)":  "Current Price x Year-End Shares Outstanding, May 2025",
    "Total Debt ($mm)":              "Nike_BS, FY2024 — Current Portion LTD + Long-Term Debt",
}

for i, item in enumerate(wacc_rows):
    r = i + 2
    if len(item) > 4 and item[4]:
        sect(ws_w, r, 1, item[0], 3); continue
    lbl, val, fmt, tc = item[0], item[1], item[2], item[3]
    lc = ws_w.cell(row=r, column=1, value=lbl)
    lc.font = _f(bold=("WACC" in lbl and "x" in lbl))
    lc.border = _border(); lc.alignment = _align("left")
    if val is not None:
        vc = ws_w.cell(row=r, column=2, value=val)
        vc.number_format = fmt or "General"
        vc.font = _f(bold=("WACC" in lbl and "x" in lbl), color=tc)
        vc.alignment = _align("right"); vc.border = _border()
    sc = ws_w.cell(row=r, column=3, value=wacc_sources.get(lbl, ""))
    sc.font = _f(sz=8, italic=True, color="666666"); sc.alignment = _align("left")


# =============================================================================
# SHEET 3 — HISTORICAL FINANCIALS
# =============================================================================

ws_h = wb.create_sheet("Historical Financials")
ws_h.sheet_view.showGridLines = False
col_w(ws_h, 1, 38)
for c in range(2, 8): col_w(ws_h, c, 12)

sect(ws_h, 1, 1, "NIKE, INC.  —  HISTORICAL FINANCIALS ($mm)  |  FY2020 - FY2024", 6)
hdr_row(ws_h, 2, ["", "FY2020", "FY2021", "FY2022", "FY2023", "FY2024"])

hist_rows = [
    ("INCOME STATEMENT", None, None, True),
    ("Revenue",             list(revenue.values),                        F_USD, True),
    ("  % Growth",          [None] + [((revenue[y]/revenue[y-1])-1)*100
                                       for y in HIST_YEARS[1:]],         F_PCT, False),
    ("Cost of Sales",       list(cogs.values),                           F_USD, False),
    ("Gross Profit",        list(gross_profit.values),                   F_USD, True),
    ("  Gross Margin",      [(gp/r)*100 for gp, r in
                             zip(gross_profit.values, revenue.values)],  F_PCT, False),
    ("SG&A",                list(sga.values),                            F_USD, False),
    ("EBIT",                list(ebit.values),                           F_USD, True),
    ("  EBIT Margin",       [(e/r)*100 for e, r in
                             zip(ebit.values, revenue.values)],          F_PCT, False),
    ("EBITDA",              list(ebitda_hist.values),                    F_USD, True),
    ("  EBITDA Margin",     [(e/r)*100 for e, r in
                             zip(ebitda_hist.values, revenue.values)],   F_PCT, False),
    ("Income Tax Expense",  list(tax_exp.values),                        F_USD, False),
    ("  Effective Tax Rate",[(t/e)*100 for t, e in
                             zip(tax_exp.values, ebit.values)],          F_PCT, False),
    ("Net Income",          list(net_income.values),                     F_USD, True),
    ("  Net Margin",        [(n/r)*100 for n, r in
                             zip(net_income.values, revenue.values)],    F_PCT, False),
    ("CASH FLOW & CAPEX", None, None, True),
    ("D&A",                 list(da.values),                             F_USD, False),
    ("CapEx",               list(capex.values),                          F_USD, False),
    ("  CapEx as % Revenue",[(c/r)*100 for c, r in
                             zip(capex.values, revenue.values)],         F_PCT, False),
    ("Operating Cash Flow", list(cfo.values),                            F_USD, True),
    ("WORKING CAPITAL", None, None, True),
    ("Accounts Receivable", list(ar_hist.values),                        F_USD, False),
    ("Inventories",         list(inv_hist.values),                       F_USD, False),
    ("Accounts Payable",    list(ap_hist.values),                        F_USD, False),
    ("Net Working Capital", list(nwc_hist.values),                       F_USD, True),
    ("  NWC as % Revenue",  [(n/r)*100 for n, r in
                             zip(nwc_hist.values, revenue.values)],      F_PCT, False),
]

row = 3
for item in hist_rows:
    lbl, vals, fmt, bold_flag = item
    if vals is None:
        sect(ws_h, row, 1, lbl, 6); row += 1; continue
    alt   = row % 2 == 0
    lc    = ws_h.cell(row=row, column=1, value=lbl)
    lc.font = _f(bold=bold_flag); lc.fill = _fill(ALT if alt else WHITE)
    lc.alignment = _align("left"); lc.border = _border()
    for j, v in enumerate(vals or []):
        safe_v = None if (v is None or (isinstance(v, float) and np.isnan(v))) else round(float(v), 2)
        dc = ws_h.cell(row=row, column=2+j, value=safe_v)
        dc.number_format = fmt; dc.font = _f(bold=bold_flag)
        dc.fill = _fill(ALT if alt else WHITE)
        dc.alignment = _align("right"); dc.border = _border()
    row += 1


# =============================================================================
# SHEET 4 — PROJECTION SHEETS (one per scenario)
# =============================================================================

def write_proj_sheet(wb, sc_name):
    sc = results[sc_name]
    ws = wb.create_sheet(f"{sc_name} Case")
    ws.sheet_view.showGridLines = False
    col_w(ws, 1, 36)
    for c in range(2, 8): col_w(ws, c, 13)

    sect(ws, 1, 1,
         f"NIKE, INC.  —  {sc_name.upper()} CASE PROJECTIONS ($mm)  |  {scenarios[sc_name]['description']}", 6)
    hdr_row(ws, 2, ["", *[str(y) for y in PROJ_YEARS]])

    def r1(vals): return [round(v, 1) for v in vals]

    proj_rows = [
        ("INCOME STATEMENT", None, None, True, True),
        ("Revenue",               r1(sc["revenues"]),       F_USD, True,  False),
        ("  % Growth",            [g*100 for g in sc["growth_rates"]], F_PCT, False, False),
        ("Gross Profit",          r1(sc["gross_profits"]),  F_USD, False, False),
        ("  Gross Margin",        [m*100 for m in sc["gm_margins"]], F_PCT, False, False),
        ("SG&A",                  r1([rv-gp-eb+rv*da_pct
                                      for rv, gp, eb in zip(sc["revenues"],
                                                             sc["gross_profits"],
                                                             sc["ebitdas"])]),
                                  F_USD, False, False),
        ("EBITDA",                r1(sc["ebitdas"]),         F_USD, True,  False),
        ("  EBITDA Margin",       [m*100 for m in sc["ebitda_margins"]], F_PCT, False, False),
        ("D&A",                   r1(sc["das"]),             F_USD, False, False),
        ("EBIT",                  r1(sc["ebits"]),           F_USD, True,  False),
        ("Taxes",                 r1(sc["taxes"]),           F_USD, False, False),
        ("NOPAT",                 r1(sc["nopats"]),          F_USD, True,  False),
        ("FREE CASH FLOW BRIDGE", None, None, True, True),
        ("NOPAT",                 r1(sc["nopats"]),          F_USD, False, False),
        ("  + D&A",               r1(sc["das"]),             F_USD, False, False),
        ("  - CapEx",             [round(-v, 1) for v in r1(sc["capexs"])], F_USD, False, False),
        ("  - Change in NWC",     [round(-v, 1) for v in r1(sc["delta_nwcs"])], F_USD, False, False),
        ("Unlevered Free Cash Flow", r1(sc["ufcfs"]),        F_USD, True,  False),
        ("  Discount Factor",     [round(1/(1+WACC)**(i+1), 4) for i in range(N)], F_4DP, False, False),
        ("  PV of FCF",           [round(v, 1) for v in sc["disc_fcfs"]], F_USD, False, False),
    ]

    r = 3
    for item in proj_rows:
        lbl, vals, fmt, bold, is_sect = item
        if is_sect and vals is None:
            sect(ws, r, 1, lbl, 6); r += 1; continue
        alt = r % 2 == 0
        lc  = ws.cell(row=r, column=1, value=lbl)
        lc.font = _f(bold=bold); lc.fill = _fill(ALT if alt else WHITE)
        lc.alignment = _align("left"); lc.border = _border()
        for j, v in enumerate(vals or []):
            dc = ws.cell(row=r, column=2+j, value=v)
            dc.number_format = fmt; dc.font = _f(bold=bold)
            dc.fill = _fill(ALT if alt else WHITE)
            dc.alignment = _align("right"); dc.border = _border()
        r += 1

    r0 = r + 1
    sect(ws, r0, 1, "DCF VALUATION SUMMARY  —  EQUITY BRIDGE", 3)
    dcf_items = [
        ("Sum of Discounted FCFs ($mm)",    sc["sum_disc_fcfs"],    F_USD,  False),
        ("Discounted Terminal Value ($mm)", sc["disc_tv"],          F_USD,  False),
        ("Enterprise Value ($mm)",          sc["enterprise_value"], F_USD,  True),
        ("  (-) Net Debt ($mm)",            NET_DEBT_MM,            F_USD,  False),
        ("  (-) Minority Interest ($mm)",   MINORITY_INT_MM,        F_USD,  False),
        ("Equity Value ($mm)",              sc["equity_value"],     F_USD,  True),
        ("  Shares Outstanding (mm)",       SHARES_OUT_MM,          F_USD,  False),
        ("Implied Share Price",             sc["implied_price"],    F_PX,   True),
        ("Current Market Price",            CURRENT_PRICE,          F_PX,   False),
        ("Premium / (Discount) to Market",  sc["premium_discount"]*100, F_PCT, True),
        ("Terminal Value as % of EV",       sc["tv_pct_ev"]*100,   F_PCT,  False),
        ("Terminal Growth Rate",            sc["tv_growth"]*100,   F_PCT,  False),
    ]
    for k, (lbl, val, fmt, bold) in enumerate(dcf_items):
        r = r0 + 1 + k
        bg = ALT if k % 2 else WHITE
        lc = ws.cell(row=r, column=1, value=lbl)
        lc.font = _f(bold=bold); lc.fill = _fill(bg); lc.border = _border()
        vc = ws.cell(row=r, column=2, value=round(val, 2))
        vc.number_format = fmt
        vc.font = _f(bold=bold, color=C_LINK if lbl == "Implied Share Price" else C_CALC)
        vc.fill = _fill(bg); vc.border = _border(); vc.alignment = _align("right")


for sc_name in scenarios:
    write_proj_sheet(wb, sc_name)


# =============================================================================
# SHEET 5 — SENSITIVITY ANALYSIS
# =============================================================================

ws_s = wb.create_sheet("Sensitivity Analysis")
ws_s.sheet_view.showGridLines = False
col_w(ws_s, 1, 16)
for c in range(2, 2+len(tvg_range)): col_w(ws_s, c, 11)

sect(ws_s, 1, 1,
     "SENSITIVITY ANALYSIS  —  Implied Share Price  |  WACC vs. Terminal Growth Rate (Base Case FCFs)",
     len(tvg_range)+2)

ws_s.cell(row=2, column=1,
          value="Gold cell = Base Case assumption. Colour scale: red = low implied price, green = high.").font = _f(sz=8, italic=True, color="555555")
ws_s.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(tvg_range)+1)

hdr = ws_s.cell(row=3, column=1, value="WACC \\ TV Growth")
hdr.font = _f(bold=True, color=WHITE); hdr.fill = _fill(NAVY)
hdr.alignment = _align("center"); hdr.border = _border()

for j, g in enumerate(tvg_range):
    c = ws_s.cell(row=3, column=2+j, value=f"{g*100:.1f}%")
    c.font = _f(bold=True, color=WHITE); c.fill = _fill(NAVY)
    c.alignment = _align("center"); c.border = _border()

for i, w in enumerate(wacc_range):
    row = 4 + i
    rc  = ws_s.cell(row=row, column=1, value=f"{w*100:.0f}%")
    rc.font = _f(bold=True, color=WHITE); rc.fill = _fill(NAVY)
    rc.alignment = _align("center"); rc.border = _border()
    for j, g in enumerate(tvg_range):
        px = sens_matrix.loc[w, g]
        c  = ws_s.cell(row=row, column=2+j)
        c.value         = float(px) if not np.isnan(px) else None
        c.number_format = F_PX
        c.alignment     = _align("center"); c.border = _border()
        if abs(w - round(WACC, 2)) < 0.005 and abs(g - 0.030) < 0.003:
            c.fill = _fill("FFD700"); c.font = _f(bold=True)

last_col = get_column_letter(1 + len(tvg_range))
ws_s.conditional_formatting.add(
    f"B4:{last_col}{3+len(wacc_range)}",
    ColorScaleRule(start_type="min",       start_color="FF4444",
                   mid_type="percentile",  mid_value=50, mid_color="FFFFFF",
                   end_type="max",         end_color="00B050")
)


# =============================================================================
# SHEET 6 — COMPARABLE COMPANIES
# =============================================================================

ws_c = wb.create_sheet("Comparable Companies")
ws_c.sheet_view.showGridLines = False
col_w(ws_c, 1, 22); col_w(ws_c, 2, 10)
for c in range(3, 8): col_w(ws_c, c, 16)

sect(ws_c, 1, 1, "COMPARABLE COMPANY ANALYSIS  —  Footwear & Apparel Peers", 7)
sect(ws_c, 2, 1, "Source: Bloomberg / FactSet Consensus  —  May 2025", 7)

hdr_row(ws_c, 3, ["Company", "Ticker", "EV/EBITDA (NTM)", "P/E (NTM)",
                  "Rev Growth (NTM)", "EBITDA Margin"])

for i, (co, row_data) in enumerate(comps.iterrows()):
    r  = 4 + i
    bg = ALT if i % 2 else WHITE
    for j, (v, fmt) in enumerate([
        (co,                        ""),
        (row_data["Ticker"],        ""),
        (row_data["EV_EBITDA_NTM"], F_MULT),
        (row_data["PE_NTM"],        F_MULT),
        (row_data["Rev_Growth"],    F_PCT),
        (row_data["EBITDA_Margin"], F_PCT),
    ]):
        c = ws_c.cell(row=r, column=1+j, value=v)
        c.number_format = fmt; c.fill = _fill(bg); c.border = _border()
        c.alignment = _align("right" if j > 1 else "left"); c.font = _f()

stat_labels = ["Mean", "Median", "25th Percentile", "75th Percentile"]
stat_funcs  = [lambda s: s.mean(), lambda s: s.median(),
               lambda s: s.quantile(0.25), lambda s: s.quantile(0.75)]
for k, (lbl, fn) in enumerate(zip(stat_labels, stat_funcs)):
    r = 4 + len(comps) + k
    lc = ws_c.cell(row=r, column=1, value=lbl)
    lc.font = _f(bold=True); lc.fill = _fill(MID_BLUE); lc.border = _border()
    for j, col_name in enumerate(["EV_EBITDA_NTM", "PE_NTM", "Rev_Growth", "EBITDA_Margin"]):
        fmt = F_MULT if j < 2 else F_PCT
        c   = ws_c.cell(row=r, column=3+j, value=round(fn(comps[col_name]), 3))
        c.number_format = fmt; c.font = _f(bold=True)
        c.fill = _fill(MID_BLUE); c.border = _border(); c.alignment = _align("right")

r0             = 4 + len(comps) + len(stat_labels) + 2
nke_ntm_ebitda = base["ebitdas"][0]
comp_ev_impl   = comp_ev_ebitda_median * nke_ntm_ebitda
comp_eq_impl   = (comp_ev_impl * 1e6 - NET_DEBT_MM * 1e6) / (SHARES_OUT_MM * 1e6)

sect(ws_c, r0, 1, "NKE IMPLIED VALUATION FROM COMPS", 5)
for k, (lbl, val, fmt) in enumerate([
    ("NKE NTM EBITDA ($mm)",         nke_ntm_ebitda,              F_USD),
    ("Peer Median EV/EBITDA",        comp_ev_ebitda_median,       F_MULT),
    ("Implied EV ($mm)",             comp_ev_impl,                F_USD),
    ("(-) Net Debt ($mm)",           NET_DEBT_MM,                 F_USD),
    ("Implied Equity Value ($mm)",  (comp_ev_impl - NET_DEBT_MM), F_USD),
    ("Implied Share Price",          comp_eq_impl,                F_PX),
]):
    r  = r0 + 1 + k
    bg = ALT if k % 2 else WHITE
    lc = ws_c.cell(row=r, column=1, value=lbl)
    lc.font = _f(bold=("Implied Share" in lbl)); lc.fill = _fill(bg); lc.border = _border()
    vc = ws_c.cell(row=r, column=2, value=round(val, 2))
    vc.number_format = fmt
    vc.font = _f(bold=("Implied Share" in lbl),
                 color=C_LINK if "Implied Share" in lbl else C_CALC)
    vc.fill = _fill(bg); vc.border = _border(); vc.alignment = _align("right")


# =============================================================================
# SHEET 7 — FOOTBALL FIELD
# =============================================================================

ws_ff = wb.create_sheet("Football Field")
ws_ff.sheet_view.showGridLines = False
col_w(ws_ff, 1, 26)
for c in range(2, 7): col_w(ws_ff, c, 14)

sect(ws_ff, 1, 1, "FOOTBALL FIELD  —  Valuation Range Summary  |  NIKE, INC. (NKE)", 6)
hdr_row(ws_ff, 2, ["Methodology", "Low", "High", "Midpoint", "vs. Current Price"])

all_ff = list(football_field.items()) + [
    ("EV/EBITDA Comps (implied)", (base["px_comps_impl"] * 0.90,
                                   base["px_comps_impl"] * 1.10)),
    ("Current Market Price",      (CURRENT_PRICE, CURRENT_PRICE)),
]
for i, (method, (lo, hi)) in enumerate(all_ff):
    r      = 3 + i
    mid    = (lo + hi) / 2
    vs_cur = (mid / CURRENT_PRICE) - 1
    bg     = ALT if i % 2 else WHITE
    up_col = "00B050" if vs_cur > 0 else "FF0000"
    for j, (v, fmt) in enumerate([
        (method, ""), (lo, F_PX), (hi, F_PX), (mid, F_PX), (vs_cur*100, F_PCT)
    ]):
        c = ws_ff.cell(row=r, column=1+j, value=v)
        c.number_format = fmt; c.fill = _fill(bg); c.border = _border()
        c.alignment = _align("right" if j > 0 else "left")
        c.font = _f(bold=(j == 0), color=up_col if j == 4 else C_CALC)

chart = BarChart()
chart.type = "bar"; chart.grouping = "clustered"
chart.title = "Valuation Range  —  Football Field"
chart.y_axis.title = "Implied Share Price ($)"
n_ff = len(all_ff)
chart.add_data(Reference(ws_ff, min_col=2, min_row=2, max_row=2+n_ff-1), titles_from_data=True)
chart.add_data(Reference(ws_ff, min_col=3, min_row=2, max_row=2+n_ff-1), titles_from_data=True)
chart.set_categories(Reference(ws_ff, min_col=1, min_row=3, max_row=2+n_ff))
chart.width = 24; chart.height = 14
ws_ff.add_chart(chart, "A12")


# =============================================================================
# SHEET 8 — MODEL ASSUMPTIONS & DOCUMENTATION
# =============================================================================

ws_a = wb.create_sheet("Model Assumptions")
ws_a.sheet_view.showGridLines = False
col_w(ws_a, 1, 32); col_w(ws_a, 2, 14); col_w(ws_a, 3, 10); col_w(ws_a, 4, 52)

sect(ws_a, 1, 1, "MODEL ASSUMPTIONS & SOURCE DOCUMENTATION", 4)
hdr_row(ws_a, 2, ["Assumption", "Value", "Unit", "Source / Notes"])

all_assumptions = [
    ("MARKET DATA", None, None, None, True),
    ("Current Share Price",     CURRENT_PRICE,    "$",   "Yahoo Finance / Bloomberg, May 2025"),
    ("Shares Outstanding",      SHARES_OUT_MM,    "mm",  "Nike_IS, FY2024 Year-End Shares Outstanding"),
    ("Market Capitalisation",   MARKET_CAP_MM,    "$mm", "Current Price x Shares Outstanding"),
    ("52-Week Low",             70.00,            "$",   "Yahoo Finance, May 2025"),
    ("52-Week High",            115.00,           "$",   "Yahoo Finance, May 2025"),

    ("WACC COMPONENTS", None, None, None, True),
    ("Risk-Free Rate",          RISK_FREE_RATE,   "%",   "US 10Y Treasury, May 2025 (FRED: DGS10)"),
    ("Equity Beta",             BETA,             "x",   "Bloomberg 5Y monthly regression vs. S&P 500, May 2025"),
    ("Equity Risk Premium",     ERP,              "%",   "Damodaran (NYU Stern), Jan 2025 Implied ERP"),
    ("Cost of Equity (CAPM)",   COST_OF_EQUITY,   "%",   "Derived: Rf + Beta x ERP"),
    ("Pre-Tax Cost of Debt",    PRE_TAX_COD,      "%",   "NKE 10-K FY2024, Note 9 — Weighted average coupon rate"),
    ("Effective Tax Rate",      EFF_TAX_RATE,     "%",   "Derived: Nike_IS FY2020-FY2024 average (Tax / EBT)"),
    ("After-Tax Cost of Debt",  COST_OF_DEBT_AT,  "%",   "Derived: Pre-Tax Kd x (1 - Effective Tax Rate)"),
    ("Equity Weight",           EQUITY_WEIGHT,    "%",   "Derived: Market Cap / (Market Cap + Total Debt)"),
    ("WACC",                    WACC,             "%",   "Derived: Ke x E% + Kd(at) x D%"),

    ("BALANCE SHEET — EQUITY BRIDGE", None, None, None, True),
    ("Cash & Equivalents",      CASH_MM,          "$mm", "Nike_BS, FY2024"),
    ("Current Portion of LTD",  LTD_CURRENT_MM,   "$mm", "Nike_BS, FY2024"),
    ("Long-Term Debt",          LTD_LONG_MM,      "$mm", "Nike_BS, FY2023 (FY2024 format change — see Note 9)"),
    ("Total Debt",              TOTAL_DEBT_MM,    "$mm", "Current Portion LTD + Long-Term Debt"),
    ("Net Debt",                NET_DEBT_MM,      "$mm", "Total Debt - Cash (negative = net cash position)"),
    ("Minority Interest",       MINORITY_INT_MM,  "$mm", "Nike_BS, FY2024 — Nike has no minority interest"),

    ("PROJECTION DRIVERS  (FY2020-FY2024 Historical Averages)", None, None, None, True),
    ("COGS as % Revenue",       cogs_pct,         "%",   "Derived from Nike_IS: avg(COGS / Revenue)"),
    ("Gross Margin",            gm_pct,           "%",   "Derived from Nike_IS: avg(Gross Profit / Revenue)"),
    ("SG&A as % Revenue",       sga_pct,          "%",   "Derived from Nike_IS: avg(SGA / Revenue)"),
    ("D&A as % Revenue",        da_pct,           "%",   "Derived from Nike_CF: avg(D&A / Revenue)"),
    ("CapEx as % Revenue",      capex_pct,        "%",   "Derived from Nike_CF: avg(CapEx / Revenue)"),
    ("NWC as % Revenue",        nwc_pct,          "%",   "Derived from Nike_BS: avg((AR + Inventory - AP) / Revenue)"),

    ("SCENARIO ASSUMPTIONS", None, None, None, True),
    ("Bear Rev Growth (Y1-Y5)", str(scenarios["Bear"]["rev_growth"]), "", scenarios["Bear"]["description"]),
    ("Base Rev Growth (Y1-Y5)", str(scenarios["Base"]["rev_growth"]), "", scenarios["Base"]["description"]),
    ("Bull Rev Growth (Y1-Y5)", str(scenarios["Bull"]["rev_growth"]), "", scenarios["Bull"]["description"]),
    ("Bear Terminal Growth",    scenarios["Bear"]["tv_growth"], "%",   "Conservative: below long-run GDP growth"),
    ("Base Terminal Growth",    scenarios["Base"]["tv_growth"], "%",   "In line with nominal GDP and long-run inflation"),
    ("Bull Terminal Growth",    scenarios["Bull"]["tv_growth"], "%",   "DTC-driven above-market steady state"),

    ("COMPARABLE COMPANIES", None, None, None, True),
    ("Peer Median EV/EBITDA",   comp_ev_ebitda_median, "x", "Bloomberg / FactSet Consensus, May 2025"),
    ("Peer Mean EV/EBITDA",     comp_ev_ebitda_mean,   "x", "Bloomberg / FactSet Consensus, May 2025"),
    ("Peer Median P/E",         comp_pe_median,        "x", "Bloomberg / FactSet Consensus, May 2025"),
]

row = 3
for item in all_assumptions:
    is_hdr = item[4] if len(item) > 4 else False
    if is_hdr:
        sect(ws_a, row, 1, item[0], 4); row += 1; continue
    lbl, val, unit, src = item[0], item[1], item[2], item[3]
    bg  = ALT if row % 2 else WHITE
    lc  = ws_a.cell(row=row, column=1, value=lbl)
    lc.font = _f(); lc.fill = _fill(bg); lc.border = _border()
    if val is not None:
        fmt = (F_PCT2 if unit == "%" else
               F_USD  if unit in ["$mm", "$"] else
               "0.00" if unit == "x" else "General")
        vc  = ws_a.cell(row=row, column=2,
                        value=val if isinstance(val, str) else val)
        vc.number_format = fmt
        vc.font = _f(color=C_INPUT if unit in ["%", "$", "$mm", "x"] else C_CALC)
        vc.fill = _fill(bg); vc.border = _border(); vc.alignment = _align("right")
    uc  = ws_a.cell(row=row, column=3, value=unit or "")
    uc.font = _f(sz=8); uc.fill = _fill(bg); uc.border = _border()
    sc_ = ws_a.cell(row=row, column=4, value=src or "")
    sc_.font = _f(sz=8, italic=True, color="555555")
    sc_.fill = _fill(bg); sc_.border = _border()
    sc_.alignment = _align("left", wrap=True)
    ws_a.row_dimensions[row].height = 15
    row += 1


# =============================================================================
# REORDER SHEETS & SAVE
# =============================================================================

sheet_order = [
    "Cover", "WACC Build-Up", "Historical Financials",
    "Bear Case", "Base Case", "Bull Case",
    "Sensitivity Analysis", "Comparable Companies",
    "Football Field", "Model Assumptions",
]
for idx, name in enumerate(sheet_order):
    if name in wb.sheetnames:
        wb.move_sheet(name, offset=wb.sheetnames.index(name) - idx)

wb.save(OUTPUT)

print(f"Model saved: {OUTPUT}")
print(f"Sheets: {', '.join(wb.sheetnames)}")
print(f"WACC: {WACC*100:.2f}%")
print(f"Implied Price  —  Bear: ${bear_px:.2f}  |  Base: ${base_px:.2f}  |  Bull: ${bull_px:.2f}")
print(f"vs. Market (${CURRENT_PRICE:.2f})  —  Base: {(base_px/CURRENT_PRICE-1)*100:.1f}%")
