**Overview**

This project is a Python-based valuation tool for Nike, Inc. It pulls historical data from an Excel database to build a 5-year DCF (Discounted Cash Flow) and a Comparable Company Analysis. The goal was to automate the repetitive parts of financial modeling—like formatting and sensitivity tables—while keeping the valuation logic transparent.

**Financial Logic**

* **WACC Build-up:** Uses the CAPM method. Sourced the Risk-Free Rate from the 10Y Treasury and the ERP from Damodaran’s 2025 data.
* **Projections:** Modeled three scenarios (Bear, Base, Bull). Revenue growth and margins are driven by Nike’s shift toward Direct-to-Consumer (DTC) channels.
* **Unlevered FCF:** Calculated as $NOPAT + D\&A - CapEx - \Delta NWC$.
* **Terminal Value:** Used the Gordon Growth Method with a 3.0% exit growth rate for the base case.

**Model Features**

* **Automated Excel Export:** The script doesn't just calculate numbers; it builds a fully formatted `.xlsx` file from scratch.
* **Formatting Conventions:** Follows standard finance color-coding (Blue for inputs, Black for formulas).
* **Sensitivity Matrix:** Includes a WACC vs. Terminal Growth heatmap to show how share price changes under different macro conditions.
* **Football Field Chart:** Aggregates the DCF, Peer Multiples, and 52-week trading range into a single summary chart.

**How to Run**

1. Make sure `Nike_Financials.xlsx` is in the folder.
2. Run `nike_dcf.py`.
3. The output file `Nike_DCF_Model_vF.xlsx` will be generated with all sheets and charts.

**Key Data Sources**

* **Historical Financials:** Nike 10-K and 10-Q filings (FY2020-2024).
* **Beta/Market Data:** Bloomberg & Yahoo Finance.
* **Peer Multiples:** FactSet consensus for Adidas, Lululemon, and Skechers.
