# Tankrich Interactive Dashboard

This repository hosts a static fundamental analysis dashboard that parses Excel exports from screener.in entirely in the browser. The first phase focuses on file upload, validation, and rendering the company overview plus annual financial statements.

## Live Site
- https://vivek-bothra.github.io/Tankrich-Interactive-dashboard/

## Features (Current)
- Client-side `.xlsx` parsing with SheetJS.
- Company overview cards with key metrics.
- Annual Profit & Loss, Balance Sheet, and Cash Flow tables.
- Graceful handling of missing data and file validation messaging.
- Basic analysis metrics (growth, profitability, leverage) and charts (revenue, margins, returns).
- Advanced frameworks: moat scoring, capex split proxy, incremental ROIC, capital allocation scorecard, value drivers, raw material sensitivity, value migration, and quality score breakdown.

## Getting Started
1. Open `index.html` in a modern browser.
2. Upload a screener.in Excel export (`.xlsx`).
3. Review the overview and financial statements tabs.

## Tech Stack
- HTML/CSS/JavaScript
- [SheetJS](https://sheetjs.com/) via CDN
- [Chart.js](https://www.chartjs.org/) via CDN

## Deployment
- GitHub Actions workflow at `.github/workflows/pages.yml` deploys from `main`.
- In repository settings, set Pages source to `GitHub Actions` if not already set.
