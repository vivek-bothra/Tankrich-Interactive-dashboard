# Tankrich Interactive Dashboard

This repository hosts a static fundamental analysis dashboard that parses Excel exports from screener.in entirely in the browser. The first phase focuses on file upload, validation, and rendering the company overview plus annual financial statements.

## Features (Phase 1)
- Client-side `.xlsx` parsing with SheetJS.
- Company overview cards with key metrics.
- Annual Profit & Loss, Balance Sheet, and Cash Flow tables.
- Graceful handling of missing data and file validation messaging.
- Basic analysis metrics (growth, profitability, leverage) and charts (revenue, margins, returns).

## Getting Started
1. Open `index.html` in a modern browser.
2. Upload a screener.in Excel export (`.xlsx`).
3. Review the overview and financial statements tabs.

## Tech Stack
- HTML/CSS/JavaScript
- [SheetJS](https://sheetjs.com/) via CDN
- [Chart.js](https://www.chartjs.org/) via CDN

## Next Steps
- Add growth metrics, ratios, and charts.
- Implement advanced frameworks (CAP, moat analysis, quality scoring, etc.).
- Add sample datasets and documentation.
