# Tankrich Dashboard: Simple Explanation of Every Value (All Tabs)

This guide explains **where each number comes from**, **how it is calculated**, and **how it is shown** in the UI.

## 1) Data source and loading

- The dashboard reads the `Data Sheet` tab from the uploaded Screener export Excel file.
- It maps fixed row numbers into annual arrays:
  - Sales, costs, profit, dividend
  - Balance-sheet items (equity, reserves, debt, assets, working-capital items)
  - Cash-flow items (CFO, CFI, CFF, net cash flow)
  - Historical price array
- All numeric cells are converted to numbers; blank/invalid become `null`.

## 2) How top header values are shown

- **Company Name**: from metadata in Excel.
- **Price**: `₹` + value with 2 decimals.
- **Market Cap**: shown in Cr (or K Cr formatting for very large values).
- **Latest Year**: last year parsed from annual date row.

## 3) Overview tab

### Growth section

- **Sales CAGR (5Y/10Y)** and **Profit CAGR (5Y/10Y)**:
  - Formula: `((End / Start)^(1/years) - 1) * 100`
  - If missing years or non-positive start/end values, shows `-`.

### Returns & Profitability section

- **ROE (Latest)**
  - `Net Profit / Avg Equity * 100`
  - `Avg Equity = (Current Equity+Reserves + Previous Equity+Reserves) / 2`
- **ROCE (Latest)**
  - `EBIT / Capital Employed * 100`
  - `EBIT = PBT + Interest`
  - `Capital Employed = Avg Equity + Debt`
- **OPM (Latest)**
  - `Operating Profit / Sales * 100`
  - `Operating Profit = Sales - Raw Material - Change in Inventory - Power/Fuel - Other Mfg - Employee Cost - Selling/Admin - Other Expenses`
- **Net Margin (Latest)**
  - `Net Profit / Sales * 100`

### Metric color hints

- Positive highlight thresholds:
  - ROE >= 18
  - ROCE >= 15
  - Sales CAGR 5Y >= 12
  - Profit CAGR 5Y >= 15

### Quality Score card (0 to 100)

Five buckets, each out of 20:

1. **Profitability (20)**
   - Profit-making years points + margin stability points.
2. **Returns (20)**
   - Latest ROCE level points + ROCE trend vs 5 years ago points.
3. **Cash Flow (20)**
   - CFO/Net Profit conversion points + FCF margin points.
4. **Balance Sheet (20)**
   - Debt/Equity points + cash-conversion-cycle points.
5. **Growth (20)**
   - Sales CAGR points + profit CAGR quality vs sales CAGR points.

Rating label:
- `>=90` Exceptional
- `>=75` High Quality
- `>=60` Above Average
- `>=40` Average
- else Low Quality

### Red Flags card

Rules checked:
- Receivables growth much faster than sales (3-year lookback)
- Inventory growth much faster than sales (3-year lookback)
- Other income too large vs operating profit
- CWIP too high vs estimated gross block
  - In current code, **Estimated Gross Block = Net Block × 1.5** (proxy, because gross block row is not directly used in this logic).
  - Red flag triggers when **CWIP > 30% of estimated gross block**.
- Debt surge + declining ROCE

Display:
- Shows count and individual warning cards.
- If none, shows “Clean Balance Sheet”.

## 4) Statements tab

Pure display tables (no extra transformations beyond formatting):

- **P&L table**: Sales, key costs, PBT, Tax, Net Profit, Dividend.
- **Balance Sheet table**: liabilities and assets sections.
- **Cash Flow table**: CFO, CFI, CFF, Net Cash Flow.

Values are shown per FY column with 2-decimal formatting.

## 5) Quarterly tab

- Pulls quarterly date row and key quarterly rows from the same `Data Sheet`.
- Builds quarter labels like `Qx FYyyyy`.
- Table shows quarterly Sales, Operating Profit, Net Profit.
- **Latest QoQ Sales Growth**: `(current quarter sales - previous quarter sales) / previous quarter sales * 100`
- **Avg QoQ Growth (4Q)**: average of latest 4 QoQ growth values.
- **YoY Sales/Profit Growth**: latest quarter vs same quarter one year earlier (index `-5`).
- Line chart plots quarterly sales and quarterly profit.

## 6) Analysis tab

### DuPont (ROIC style)

- `EBIT = PBT + Interest`
- `Tax Rate = Tax / PBT` (fallback 25% if needed)
- `NOPAT = EBIT * (1 - Tax Rate)`
- `Excess Cash = max(0, Cash - 2% of Sales)`
- `Invested Capital = Equity + Debt - Excess Cash`
- `ROIC = NOPAT / Invested Capital`
- Drivers shown:
  - NOPAT Margin = `NOPAT / Sales`
  - Invested Capital Turnover = `Sales / Invested Capital`
- Strategy label based on margin + turnover bands.

### Efficiency Metrics

- Asset Turnover = `Sales / Total Assets`
- Debtor Days = `Receivables / Sales * 365`
- Inventory Days = `Inventory / COGS * 365`
- Cash Conversion Cycle (approx) = `Debtor Days + Inventory Days - 30` (30 assumed payable days)
- Also shows improving/declining trend and text insights.

### Leverage & Solvency

- Debt/Equity = `Debt / (Equity + Reserves)`
- Interest Coverage = `EBIT / Interest`
- Total Debt shown directly
- Qualitative label: conservative/moderate/high leverage.

## 7) Frameworks tab

### Moat Detection (14 points)

Six tests:
1. High & sustained ROIC (3 pts)
2. ROIC stability via std deviation (2 pts)
3. Pricing power via margin stability/trend (3 pts)
4. Scale advantages via asset-turnover improvement (2 pts)
5. Customer stickiness via debtor days behavior + growth (2 pts)
6. Core earnings quality (other income as % of op profit) (2 pts)

Total score mapped to No/Narrow/Wide/Exceptional moat.

### Capital Allocation Scorecard (100 points)

- Reinvestment quality (60 points): based on **average incremental ROIC** over recent periods.
  - Incremental ROIC = `Delta NOPAT / Delta Capital`
  - Capital = Equity+Reserves+Debt
- Deployment logic (40 points): checks whether payout ratio policy matches return quality.
- Final grade A/B/C/D/F from total score.

### Value Migration

Uses 5-year view:
- Sales CAGR
- Margin change (old vs latest net margin)
- ROCE change

Classifies direction as Strong Inward / Inward / Stable / Outward.

### Earning Power Box

Rolling 3-year windows:
- PAT CAGR (3Y)
- Average CFO / Net Profit conversion

Quadrants:
- STAR (high growth, high cash)
- INVESTIGATE (high growth, low cash)
- CASH COW (low growth, high cash)
- RED FLAG (low growth, low cash)

Also shows trajectory pattern commentary.

## 8) Advanced tab

### CAP (Competitive Advantage Period)

- Calculates company WACC:
  - Cost of equity = `Risk-free 7% + beta(1.0)*ERP(6%)`
  - Cost of debt from `Interest / Debt`, adjusted for tax
  - WACC = weighted equity + weighted after-tax debt
- ROIC series by year (using NOPAT proxy `Net Profit * 1.15`)
- **CAP years** = count of years where `ROIC > WACC`
- Shows CAP chart: ROIC line vs WACC line.

### Maintenance vs Growth Capex

For recent years:
- Total capex proxy = `abs(CFI)`
- Maintenance capex proxy = `Depreciation`
- Growth capex = `max(0, Total capex - Maintenance capex)`

Adds insights:
- Capex intensity trend (`capex/sales`)
- Maintenance adequacy
- Growth capex ROI vs EBIT change
- Owner earnings = `CFO - Maintenance capex`
- FCF after maintenance

### Incremental Returns Analysis

- For recent years:
  - Incremental ROIC = `Delta NOPAT / Delta Capital`
  - Historical ROIC proxy = `NOPAT / avg capital`
- Compares both and gives grade from average incremental ROIC.

### Raw Material Sensitivity

- RM intensity = `Raw Material / Sales`
- Gross margin series = `(Sales - RM - Other Mfg) / Sales`
- Volatility = std dev of gross margin
- Labels sensitivity as Low/Medium/High.

### Buffett’s $1 Test

If price history exists:
- Retained earnings (5Y) = sum(Net Profit - Dividend)
- Market cap change = (shares × latest price) - (shares × old price), unit adjusted to Cr
- Ratio = `Market cap change / retained earnings`
- Pass if ratio >= 1

Fallback (if no price series):
- Uses book value change instead of market cap change.

### FLOAT Detection

Signals:
1. Strong negative CCC (float-like funding)
2. High other liabilities as % of sales
3. Other income earned on liability base

If signal(s) exist, marks float detected and shows details.

## 9) Charts tab

### Revenue & Profit Trend
- Line chart of annual Sales and Net Profit.

### Margin Evolution
- Operating margin line and net margin line.

### Returns Analysis
- Bar chart of annual ROE and ROCE.

### Cash Flow Breakdown
- Bar chart of CFO, CFI, CFF.

## 10) Compare tab

- Appears only after adding at least one extra company Excel.
- For each company card it calculates and displays:
  - Market Cap
  - Sales CAGR (5Y)
  - Profit CAGR (5Y)
  - ROE
  - Net Margin
  - Debt/Equity
  - Quality Score

## 11) Presentation/format rules

- Missing/unusable values show as `-` or `N/A`.
- Number format:
  - Standard number: 2 decimals
  - Percent: 1 decimal + `%`
  - Currency: `₹` + 2 decimals
- Tab switching simply toggles active class.


