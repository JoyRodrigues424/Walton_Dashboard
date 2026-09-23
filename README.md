# Walton_Dashboard
A comprehensive dashboard on Walton Hi-Tech Industry companies performance of Past 3 years 
# Walton Hi-Tech Industries PLC — Financial Performance Dashboard (FY22–FY25)

An interactive corporate financial analysis model and executive dashboard built in Microsoft Excel analyzing the multi-year performance of **Walton Hi-Tech Industries PLC** across four fiscal years (2021–22 to 2024–25).

---

## 📌 Executive Summary & Key Insights

* **Top-Line Resilience:** Peak revenue of BDT 81.68B in FY22, followed by a rebound in FY24 (+13.18%) and slight contraction in FY25 (BDT 70.82B, -5.72%). Gross margin remained steady between **35% and 38%**.
* **Capital Structure & Deleveraging:** Debt-to-Equity fell systematically from **0.46x in FY22 down to 0.22x in FY25**, reflecting lower reliance on external debt and sustained interest coverage of 3.43x.
* **Liquidity Buffer:** Current Ratio strengthened from **1.59x to 2.55x**, supported by BDT 5.06B in liquid cash and cash equivalents.
* **Cash Flow Turnaround:** Operating cash flow reversed from negative (-BDT 2.40B in FY22) to robust positive cash generation (**BDT 17.63B in FY25**), yielding BDT 12.43B in Free Cash Flow.

---

## ⚡ Interactive Dashboard Feature: Dynamic Year Selection

The `Dashboard` sheet features dynamic KPI cards that automatically update all solvency, liquidity, and operational efficiency ratios based on user selection.

### How to Use the Dashboard
1. Open the file in **Microsoft Excel Desktop** (Office 2021 or Microsoft 365 recommended).
2. Navigate to the **`Dashboard`** tab.
3. Click on the dropdown cell **`C3`** (labeled **Year:**).
4. Select any fiscal year: **`2023`**, **`2024`**, or **`2025`**.
5. All KPI scorecards under **Financial Health** and **Operating Efficiency** will update instantly.

### Technical Implementation

* **Data Validation:** Applied to cell `C3` using list source `=2023, 2024, 2025`.
* **Dynamic Retrieval via `XLOOKUP`:** The calculation engine maps the selected year directly from `Dashboard!$C$3` to the multi-year metric matrix in the backend sheet:

```excel
=XLOOKUP(Dashboard!$C$3, Calc!$H$8:$J$8, Calc!H9:J9)
