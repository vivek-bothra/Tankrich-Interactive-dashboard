# Tankrich Fundamental Analysis Dashboard

A comprehensive, client-side fundamental analysis tool for equity research. Upload Excel files from screener.in and get instant, deep financial analysis with advanced investment frameworks.

## 🚀 Live Demo

**[Visit Dashboard](https://vivek-bothra.github.io/Tankrich-Interactive-dashboard/)**

## ✨ Features

### Core Financial Analysis
- ✅ Complete Financial Statements (P&L, Balance Sheet, Cash Flow)
- ✅ Growth Metrics (CAGRs: 3yr, 5yr, 7yr, 10yr)
- ✅ Profitability Ratios (ROE, ROCE, ROIC, Margins)
- ✅ Efficiency Metrics (Asset Turnover, Working Capital, CCC)
- ✅ Leverage Analysis (Debt-to-Equity, Interest Coverage)

### Advanced Frameworks
- 🎯 **Quality Score** - 100-point comprehensive scoring system
- 🚩 **Red Flags Detection** - 7 balance sheet health checks
- 📊 **DuPont Analysis** - ROE decomposition into components
- 🏰 **Moat Analysis** - Competitive advantage indicators
- 💰 **Capital Allocation** - Management quality grading
- 📈 **Value Migration** - Business trajectory analysis

### Visualizations
- Revenue & Profit trends
- Margin evolution
- Returns analysis (ROE/ROCE)
- Cash flow breakdown

## 🎯 How to Use

### 1. Export Data from Screener.in

1. Go to [screener.in](https://www.screener.in/)
2. Search for any company
3. Click on "Export" button
4. Download the Excel file

### 2. Upload to Dashboard

1. Open the [Tankrich Dashboard](https://vivek-bothra.github.io/Tankrich-Interactive-dashboard/)
2. Click "Upload Excel from Screener.in"
3. Select your downloaded file
4. Analysis appears instantly!

## 📊 What Each Section Tells You

### Overview Tab
- **Quality Score**: 100-point rating across 5 dimensions (profitability, returns, cash flow, balance sheet, growth)
- **Red Flags**: Automatic detection of balance sheet issues
- **Growth Metrics**: Historical CAGRs for revenue and profit
- **Returns**: Latest year profitability ratios

### Statements Tab
Complete financial statements with 10 years of historical data

### Analysis Tab
- **DuPont Analysis**: Understand what drives ROE (margins, efficiency, or leverage)
- **Efficiency Metrics**: How well the company uses its assets
- **Leverage**: Debt levels and coverage ratios

### Frameworks Tab
- **Moat Detection**: Does the company have sustainable competitive advantages?
- **Capital Allocation**: How well does management deploy capital?
- **Value Migration**: Is value flowing to or from this business?

### Charts Tab
Interactive visualizations of all key metrics over time

## 🔧 Technical Details

- **100% Client-Side**: No data is sent to any server
- **Privacy First**: All processing happens in your browser
- **No Installation**: Works directly from GitHub Pages
- **Framework**: Vanilla JavaScript with Chart.js
- **Excel Parsing**: SheetJS library

## 📁 Repository Structure

```
├── index.html          # Main HTML structure
├── styles.css          # Professional dashboard styling
├── app.js              # All calculations and logic
└── README.md           # This file
```

## 🎨 Design Philosophy

Professional financial terminal aesthetic with:
- Dark theme optimized for long analysis sessions
- Clear data hierarchy
- High information density without clutter
- Subtle animations for smooth interactions

## 📊 Sample Data

Test the dashboard with these companies:
- Asian Paints (mature, consistent performer)
- RateGain Travel (recently listed, limited history)

## 🚀 Deployment

Automatically deployed to GitHub Pages via GitHub Actions.

Any push to `main` branch triggers a new deployment.

## 🛠️ Local Development

```bash
# Clone the repository
git clone https://github.com/vivek-bothra/Tankrich-Interactive-dashboard.git

# Open index.html in your browser
open index.html
```

No build process required!

## 📖 Understanding the Metrics

### Quality Score (0-100)
- **90-100**: ⭐⭐⭐⭐⭐ Exceptional (Blue-chip compounders)
- **75-89**: ⭐⭐⭐⭐ High Quality (Strong long-term holds)
- **60-74**: ⭐⭐⭐ Above Average (Good businesses)
- **40-59**: ⭐⭐ Average (Cyclical plays)
- **0-39**: ⭐ Low Quality (High risk)

### Red Flags (0-7)
- **0-1 flags**: ✅ Clean balance sheet
- **2-3 flags**: ⚠️ Caution - investigate further
- **4+ flags**: 🔴 High risk - avoid or deep dive

### Moat Score
- **Wide Moat**: Strong sustainable competitive advantages
- **Narrow Moat**: Some competitive advantages
- **No Moat**: Commodity-like business

### Capital Allocation Grade
- **A**: Excellent capital allocator
- **B**: Good capital allocator
- **C**: Average capital allocator
- **D/F**: Poor - avoid management

## 🤝 Contributing

Contributions welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## 📝 License

MIT License - feel free to use and modify

## 🙏 Acknowledgments

Built with insights from:
- Michael Mauboussin (CAP framework)
- Pat Dorsey (Moat analysis)
- Hewitt Heiserman (Earnings quality)
- Warren Buffett (Capital allocation principles)

## 📧 Contact

For questions or suggestions, please open an issue on GitHub.

---

**Made for fundamental investors, by a fundamental investor** 📈
