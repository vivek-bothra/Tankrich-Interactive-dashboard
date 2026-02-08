const fileInput = document.getElementById("fileInput");
const uploadStatus = document.getElementById("uploadStatus");

const companyNameEl = document.getElementById("companyName");
const currentPriceEl = document.getElementById("currentPrice");
const marketCapEl = document.getElementById("marketCap");
const latestYearEl = document.getElementById("latestYear");

const overviewNameEl = document.getElementById("overviewName");
const faceValueEl = document.getElementById("faceValue");
const overviewLatestEl = document.getElementById("overviewLatest");
const dataCoverageEl = document.getElementById("dataCoverage");
const highlightsEl = document.getElementById("highlights");
const salesCagrEl = document.getElementById("salesCagr");
const profitCagrEl = document.getElementById("profitCagr");
const roeLatestEl = document.getElementById("roeLatest");
const roceLatestEl = document.getElementById("roceLatest");

const growthMetricsEl = document.getElementById("growthMetrics");
const profitabilityMetricsEl = document.getElementById("profitabilityMetrics");
const efficiencyMetricsEl = document.getElementById("efficiencyMetrics");
const leverageMetricsEl = document.getElementById("leverageMetrics");

const revenueChartEl = document.getElementById("revenueChart");
const marginChartEl = document.getElementById("marginChart");
const returnChartEl = document.getElementById("returnChart");
const capChartEl = document.getElementById("capChart");
const epbChartEl = document.getElementById("epbChart");

const capDurationEl = document.getElementById("capDuration");
const capSignalEl = document.getElementById("capSignal");
const epbCagrEl = document.getElementById("epbCagr");
const epbCashEl = document.getElementById("epbCash");
const epbQuadrantEl = document.getElementById("epbQuadrant");
const redFlagsEl = document.getElementById("redFlags");
const dupontMetricsEl = document.getElementById("dupontMetrics");
const qualityScoreEl = document.getElementById("qualityScore");
const buffettTestEl = document.getElementById("buffettTest");

const plTableEl = document.getElementById("plTable");
const bsTableEl = document.getElementById("bsTable");
const cfTableEl = document.getElementById("cfTable");

const TAB_BUTTONS = document.querySelectorAll(".tab-button");
const TAB_PANELS = document.querySelectorAll(".tab-panel");

const REQUIRED_SHEETS = ["Data Sheet", "Profit & Loss", "Balance Sheet", "Cash Flow"];

const ROWS = {
  companyName: 0,
  faceValue: 6,
  currentPrice: 7,
  marketCap: 8,
  reportDates: 15,
  sales: 16,
  rawMaterial: 17,
  inventoryChange: 18,
  powerFuel: 19,
  otherMfg: 20,
  employeeCost: 21,
  sellingAdmin: 22,
  otherExpenses: 23,
  otherIncome: 24,
  depreciation: 25,
  interest: 26,
  pbt: 27,
  tax: 28,
  netProfit: 29,
  dividend: 30,
  bsDates: 55,
  equity: 56,
  reserves: 57,
  borrowings: 58,
  otherLiabilities: 59,
  totalLiabilities: 60,
  netBlock: 61,
  cwip: 62,
  investments: 63,
  otherAssets: 64,
  totalAssets: 65,
  receivables: 66,
  inventory: 67,
  cash: 68,
  shares: 69,
  cfDates: 80,
  cfo: 81,
  cfi: 82,
  cff: 83,
  netCash: 84,
};

const safeNumber = (value) => {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : null;
};

const formatNumber = (value, suffix = "") => {
  if (value === null || value === undefined) return "-";
  return `${value.toLocaleString(undefined, { maximumFractionDigits: 2 })}${suffix}`;
};

const formatPercent = (value) => {
  if (value === null || value === undefined) return "-";
  return `${value.toFixed(1)}%`;
};

const getRowSlice = (rows, index) => {
  if (!rows[index]) return [];
  return rows[index].slice(4);
};

const calculateMargin = (numerator, denominator) => {
  if (!Number.isFinite(numerator) || !Number.isFinite(denominator) || denominator === 0) return null;
  return (numerator / denominator) * 100;
};

const getLatestValue = (values) => {
  const filtered = values.filter((val) => Number.isFinite(val));
  return filtered.length ? filtered[filtered.length - 1] : null;
};

const getAverage = (values) => {
  const filtered = values.filter((val) => Number.isFinite(val));
  if (!filtered.length) return null;
  return filtered.reduce((sum, val) => sum + val, 0) / filtered.length;
};

const calculateCAGR = (values, years) => {
  if (!values.length || years <= 0) return null;
  const filtered = values.filter((val) => Number.isFinite(val));
  if (filtered.length <= years) return null;
  const end = filtered[filtered.length - 1];
  const start = filtered[filtered.length - 1 - years];
  if (!Number.isFinite(start) || !Number.isFinite(end) || start <= 0 || end <= 0) return null;
  return (Math.pow(end / start, 1 / years) - 1) * 100;
};

const calculateRatio = (numerator, denominator) => {
  if (!Number.isFinite(numerator) || !Number.isFinite(denominator) || denominator === 0) return null;
  return numerator / denominator;
};

const formatRatio = (value) => {
  if (value === null || value === undefined) return "-";
  return `${value.toFixed(2)}x`;
};

const renderMetricList = (container, metrics) => {
  container.innerHTML = "";
  metrics.forEach(({ label, value }) => {
    const row = document.createElement("div");
    row.className = "metric-item";
    const name = document.createElement("span");
    name.textContent = label;
    const metricValue = document.createElement("span");
    metricValue.textContent = value ?? "-";
    row.appendChild(name);
    row.appendChild(metricValue);
    container.appendChild(row);
  });
};

const destroyChart = (chartRef) => {
  if (chartRef?.destroy) {
    chartRef.destroy();
  }
};

let revenueChart;
let marginChart;
let returnChart;
let capChart;
let epbChart;

const renderTable = (container, headers, rows) => {
  if (!headers.length) {
    container.innerHTML = "<p class=\"status-text\">No data available.</p>";
    return;
  }

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");

  headers.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headRow.appendChild(th);
  });

  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    row.forEach((cell, index) => {
      const td = document.createElement("td");
      td.textContent = cell ?? "-";
      if (index > 0 && typeof cell === "string" && cell.endsWith("%")) {
        const numeric = Number(cell.replace("%", ""));
        if (Number.isFinite(numeric)) {
          td.classList.add(numeric >= 0 ? "positive" : "negative");
        }
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
  container.innerHTML = "";
  container.appendChild(table);
};

const parseDataSheet = (sheet) => {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });

  const meta = {
    name: rows[ROWS.companyName]?.[1] ?? "N/A",
    faceValue: safeNumber(rows[ROWS.faceValue]?.[1]),
    currentPrice: safeNumber(rows[ROWS.currentPrice]?.[1]),
    marketCap: safeNumber(rows[ROWS.marketCap]?.[1]),
  };

  const dates = getRowSlice(rows, ROWS.reportDates).map((date) => date ?? "-");

  const metrics = {
    sales: getRowSlice(rows, ROWS.sales).map(safeNumber),
    rawMaterial: getRowSlice(rows, ROWS.rawMaterial).map(safeNumber),
    inventoryChange: getRowSlice(rows, ROWS.inventoryChange).map(safeNumber),
    powerFuel: getRowSlice(rows, ROWS.powerFuel).map(safeNumber),
    otherMfg: getRowSlice(rows, ROWS.otherMfg).map(safeNumber),
    employeeCost: getRowSlice(rows, ROWS.employeeCost).map(safeNumber),
    sellingAdmin: getRowSlice(rows, ROWS.sellingAdmin).map(safeNumber),
    otherExpenses: getRowSlice(rows, ROWS.otherExpenses).map(safeNumber),
    otherIncome: getRowSlice(rows, ROWS.otherIncome).map(safeNumber),
    depreciation: getRowSlice(rows, ROWS.depreciation).map(safeNumber),
    interest: getRowSlice(rows, ROWS.interest).map(safeNumber),
    pbt: getRowSlice(rows, ROWS.pbt).map(safeNumber),
    tax: getRowSlice(rows, ROWS.tax).map(safeNumber),
    netProfit: getRowSlice(rows, ROWS.netProfit).map(safeNumber),
    dividend: getRowSlice(rows, ROWS.dividend).map(safeNumber),
  };

  const balanceSheetDates = getRowSlice(rows, ROWS.bsDates).map((date) => date ?? "-");
  const balanceSheet = {
    equity: getRowSlice(rows, ROWS.equity).map(safeNumber),
    reserves: getRowSlice(rows, ROWS.reserves).map(safeNumber),
    borrowings: getRowSlice(rows, ROWS.borrowings).map(safeNumber),
    otherLiabilities: getRowSlice(rows, ROWS.otherLiabilities).map(safeNumber),
    totalLiabilities: getRowSlice(rows, ROWS.totalLiabilities).map(safeNumber),
    netBlock: getRowSlice(rows, ROWS.netBlock).map(safeNumber),
    cwip: getRowSlice(rows, ROWS.cwip).map(safeNumber),
    investments: getRowSlice(rows, ROWS.investments).map(safeNumber),
    otherAssets: getRowSlice(rows, ROWS.otherAssets).map(safeNumber),
    totalAssets: getRowSlice(rows, ROWS.totalAssets).map(safeNumber),
    receivables: getRowSlice(rows, ROWS.receivables).map(safeNumber),
    inventory: getRowSlice(rows, ROWS.inventory).map(safeNumber),
    cash: getRowSlice(rows, ROWS.cash).map(safeNumber),
    shares: getRowSlice(rows, ROWS.shares).map(safeNumber),
  };

  const cashFlowDates = getRowSlice(rows, ROWS.cfDates).map((date) => date ?? "-");
  const cashFlow = {
    cfo: getRowSlice(rows, ROWS.cfo).map(safeNumber),
    cfi: getRowSlice(rows, ROWS.cfi).map(safeNumber),
    cff: getRowSlice(rows, ROWS.cff).map(safeNumber),
    netCash: getRowSlice(rows, ROWS.netCash).map(safeNumber),
  };

  return { meta, dates, metrics, balanceSheetDates, balanceSheet, cashFlowDates, cashFlow };
};

const updateOverview = (data) => {
  const latestYear = data.dates[data.dates.length - 1] ?? "-";
  const coverage = `${data.dates.length} years`;
  const latestSales = getLatestValue(data.metrics.sales);
  const latestProfit = getLatestValue(data.metrics.netProfit);
  const latestMargin = calculateMargin(latestProfit, latestSales);

  const salesCagr5 = calculateCAGR(data.metrics.sales, 5);
  const profitCagr5 = calculateCAGR(data.metrics.netProfit, 5);

  const equityValues = data.balanceSheet.equity.map((equity, index) => {
    const reserves = data.balanceSheet.reserves[index];
    if (!Number.isFinite(equity) && !Number.isFinite(reserves)) return null;
    return (equity ?? 0) + (reserves ?? 0);
  });

  const latestEquity = getLatestValue(equityValues);
  const latestBorrowings = getLatestValue(data.balanceSheet.borrowings);

  const roeLatest = calculateMargin(latestProfit, latestEquity);
  const capitalEmployed = Number.isFinite(latestEquity) || Number.isFinite(latestBorrowings)
    ? (latestEquity ?? 0) + (latestBorrowings ?? 0)
    : null;
  const roceLatest = calculateMargin(latestProfit, capitalEmployed);

  companyNameEl.textContent = data.meta.name ?? "-";
  currentPriceEl.textContent = formatNumber(data.meta.currentPrice);
  marketCapEl.textContent = formatNumber(data.meta.marketCap);
  latestYearEl.textContent = latestYear;

  overviewNameEl.textContent = data.meta.name ?? "-";
  faceValueEl.textContent = formatNumber(data.meta.faceValue);
  overviewLatestEl.textContent = latestYear;
  dataCoverageEl.textContent = coverage;

  highlightsEl.innerHTML = "";
  const highlights = [
    `Latest Sales: ${formatNumber(latestSales)}`,
    `Latest Net Profit: ${formatNumber(latestProfit)}`,
    `Net Margin: ${formatPercent(latestMargin)}`,
  ];

  highlights.forEach((item) => {
    const p = document.createElement("p");
    p.textContent = item;
    highlightsEl.appendChild(p);
  });

  salesCagrEl.textContent = formatPercent(salesCagr5);
  profitCagrEl.textContent = formatPercent(profitCagr5);
  roeLatestEl.textContent = formatPercent(roeLatest);
  roceLatestEl.textContent = formatPercent(roceLatest);
};

const buildPLTable = (data) => {
  const headers = ["Metric", ...data.dates];
  const rows = [];

  const operatingExpenses = data.metrics.sales.map((_, index) => {
    const values = [
      data.metrics.rawMaterial[index],
      data.metrics.inventoryChange[index],
      data.metrics.powerFuel[index],
      data.metrics.otherMfg[index],
      data.metrics.employeeCost[index],
      data.metrics.sellingAdmin[index],
      data.metrics.otherExpenses[index],
    ];
    if (values.every((val) => val === null)) return null;
    return values.reduce((sum, value) => (Number.isFinite(value) ? sum + value : sum), 0);
  });

  const ebitda = data.metrics.sales.map((sales, index) => {
    if (!Number.isFinite(sales)) return null;
    const expenses = operatingExpenses[index] ?? 0;
    const otherIncome = data.metrics.otherIncome[index] ?? 0;
    return sales - expenses + otherIncome;
  });

  const rowMap = [
    ["Sales", data.metrics.sales],
    ["Raw Material Cost", data.metrics.rawMaterial],
    ["Change in Inventory", data.metrics.inventoryChange],
    ["Power & Fuel", data.metrics.powerFuel],
    ["Other Manufacturing", data.metrics.otherMfg],
    ["Employee Cost", data.metrics.employeeCost],
    ["Selling & Admin", data.metrics.sellingAdmin],
    ["Other Expenses", data.metrics.otherExpenses],
    ["Other Income", data.metrics.otherIncome],
    ["EBITDA", ebitda],
    ["Depreciation", data.metrics.depreciation],
    ["Interest", data.metrics.interest],
    ["Profit Before Tax", data.metrics.pbt],
    ["Tax", data.metrics.tax],
    ["Net Profit", data.metrics.netProfit],
    ["Dividend", data.metrics.dividend],
  ];

  rowMap.forEach(([label, values]) => {
    rows.push([label, ...values.map((value) => (value === null ? "-" : formatNumber(value)))]);
  });

  const marginsRow = data.metrics.sales.map((sales, index) =>
    formatPercent(calculateMargin(data.metrics.netProfit[index], sales))
  );
  rows.push(["Net Margin %", ...marginsRow]);

  renderTable(plTableEl, headers, rows);
};

const buildBSTable = (data) => {
  const headers = ["Metric", ...data.balanceSheetDates];
  const rows = [
    ["Equity Share Capital", data.balanceSheet.equity],
    ["Reserves", data.balanceSheet.reserves],
    ["Borrowings", data.balanceSheet.borrowings],
    ["Other Liabilities", data.balanceSheet.otherLiabilities],
    ["Total Liabilities", data.balanceSheet.totalLiabilities],
    ["Net Block", data.balanceSheet.netBlock],
    ["CWIP", data.balanceSheet.cwip],
    ["Investments", data.balanceSheet.investments],
    ["Other Assets", data.balanceSheet.otherAssets],
    ["Total Assets", data.balanceSheet.totalAssets],
    ["Receivables", data.balanceSheet.receivables],
    ["Inventory", data.balanceSheet.inventory],
    ["Cash & Bank", data.balanceSheet.cash],
    ["Number of Equity Shares", data.balanceSheet.shares],
  ].map(([label, values]) => [label, ...values.map((value) => (value === null ? "-" : formatNumber(value)))]);

  renderTable(bsTableEl, headers, rows);
};

const buildCFTable = (data) => {
  const headers = ["Metric", ...data.cashFlowDates];
  const rows = [
    ["Cash From Operating Activity", data.cashFlow.cfo],
    ["Cash From Investing Activity", data.cashFlow.cfi],
    ["Cash From Financing Activity", data.cashFlow.cff],
    ["Net Cash Flow", data.cashFlow.netCash],
  ].map(([label, values]) => [label, ...values.map((value) => (value === null ? "-" : formatNumber(value)))]);

  renderTable(cfTableEl, headers, rows);
};

const buildAnalysis = (data) => {
  const salesCagr3 = calculateCAGR(data.metrics.sales, 3);
  const salesCagr5 = calculateCAGR(data.metrics.sales, 5);
  const profitCagr3 = calculateCAGR(data.metrics.netProfit, 3);
  const profitCagr5 = calculateCAGR(data.metrics.netProfit, 5);

  const equityValues = data.balanceSheet.equity.map((equity, index) => {
    const reserves = data.balanceSheet.reserves[index];
    if (!Number.isFinite(equity) && !Number.isFinite(reserves)) return null;
    return (equity ?? 0) + (reserves ?? 0);
  });
  const assets = data.balanceSheet.totalAssets;
  const latestProfit = getLatestValue(data.metrics.netProfit);
  const latestEquity = getLatestValue(equityValues);
  const latestAssets = getLatestValue(assets);
  const latestBorrowings = getLatestValue(data.balanceSheet.borrowings);
  const latestSales = getLatestValue(data.metrics.sales);

  const roeLatest = calculateMargin(latestProfit, latestEquity);
  const roceLatest = calculateMargin(latestProfit, (latestEquity ?? 0) + (latestBorrowings ?? 0));
  const roaLatest = calculateMargin(latestProfit, latestAssets);

  const inventory = getLatestValue(data.balanceSheet.inventory);
  const receivables = getLatestValue(data.balanceSheet.receivables);
  const inventoryDays = calculateRatio(inventory, latestSales);
  const debtorDays = calculateRatio(receivables, latestSales);

  const debtToEquity = calculateRatio(latestBorrowings, latestEquity);
  const interest = getLatestValue(data.metrics.interest);
  const ebit = Number.isFinite(latestProfit) && Number.isFinite(interest)
    ? latestProfit + interest
    : null;
  const interestCoverage = calculateRatio(ebit, interest);

  renderMetricList(growthMetricsEl, [
    { label: "Sales CAGR (3Y)", value: formatPercent(salesCagr3) },
    { label: "Sales CAGR (5Y)", value: formatPercent(salesCagr5) },
    { label: "Profit CAGR (3Y)", value: formatPercent(profitCagr3) },
    { label: "Profit CAGR (5Y)", value: formatPercent(profitCagr5) },
  ]);

  renderMetricList(profitabilityMetricsEl, [
    { label: "ROE (Latest)", value: formatPercent(roeLatest) },
    { label: "ROCE (Latest)", value: formatPercent(roceLatest) },
    { label: "ROA (Latest)", value: formatPercent(roaLatest) },
  ]);

  renderMetricList(efficiencyMetricsEl, [
    { label: "Inventory Days (proxy)", value: inventoryDays ? `${(inventoryDays * 365).toFixed(0)} days` : "-" },
    { label: "Debtor Days (proxy)", value: debtorDays ? `${(debtorDays * 365).toFixed(0)} days` : "-" },
  ]);

  renderMetricList(leverageMetricsEl, [
    { label: "Debt-to-Equity", value: formatRatio(debtToEquity) },
    { label: "Interest Coverage", value: formatRatio(interestCoverage) },
  ]);
};

const buildCharts = (data) => {
  const labels = data.dates;
  const sales = data.metrics.sales.map((value) => value ?? null);
  const profit = data.metrics.netProfit.map((value) => value ?? null);

  destroyChart(revenueChart);
  destroyChart(marginChart);
  destroyChart(returnChart);

  revenueChart = new Chart(revenueChartEl, {
    type: "line",
    data: {
      labels,
      datasets: [
        { label: "Sales", data: sales, borderColor: "#2b6cff", backgroundColor: "rgba(43, 108, 255, 0.2)" },
        { label: "Net Profit", data: profit, borderColor: "#1d8f50", backgroundColor: "rgba(29, 143, 80, 0.2)" },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
    },
  });

  const opm = data.metrics.sales.map((salesValue, index) =>
    calculateMargin(
      (salesValue ?? 0) -
        (data.metrics.rawMaterial[index] ?? 0) -
        (data.metrics.inventoryChange[index] ?? 0) -
        (data.metrics.powerFuel[index] ?? 0) -
        (data.metrics.otherMfg[index] ?? 0) -
        (data.metrics.employeeCost[index] ?? 0) -
        (data.metrics.sellingAdmin[index] ?? 0) -
        (data.metrics.otherExpenses[index] ?? 0),
      salesValue
    )
  );
  const npm = data.metrics.sales.map((salesValue, index) =>
    calculateMargin(data.metrics.netProfit[index], salesValue)
  );

  marginChart = new Chart(marginChartEl, {
    type: "line",
    data: {
      labels,
      datasets: [
        { label: "Operating Margin %", data: opm, borderColor: "#f5a524" },
        { label: "Net Margin %", data: npm, borderColor: "#dc2626" },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
    },
  });

  const equityValues = data.balanceSheet.equity.map((equity, index) => {
    const reserves = data.balanceSheet.reserves[index];
    if (!Number.isFinite(equity) && !Number.isFinite(reserves)) return null;
    return (equity ?? 0) + (reserves ?? 0);
  });
  const roeSeries = data.metrics.netProfit.map((profitValue, index) =>
    calculateMargin(profitValue, equityValues[index])
  );
  const roceSeries = data.metrics.netProfit.map((profitValue, index) => {
    const borrowings = data.balanceSheet.borrowings[index];
    const capital = (equityValues[index] ?? 0) + (borrowings ?? 0);
    return calculateMargin(profitValue, capital);
  });

  returnChart = new Chart(returnChartEl, {
    type: "bar",
    data: {
      labels,
      datasets: [
        { label: "ROE %", data: roeSeries, backgroundColor: "rgba(43, 108, 255, 0.6)" },
        { label: "ROCE %", data: roceSeries, backgroundColor: "rgba(29, 143, 80, 0.6)" },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
    },
  });
};

const buildFrameworks = (data) => {
  const labels = data.dates;
  const wacc = 13;
  const equityValues = data.balanceSheet.equity.map((equity, index) => {
    const reserves = data.balanceSheet.reserves[index];
    if (!Number.isFinite(equity) && !Number.isFinite(reserves)) return null;
    return (equity ?? 0) + (reserves ?? 0);
  });

  const taxRateSeries = data.metrics.pbt.map((pbt, index) => {
    const tax = data.metrics.tax[index];
    if (!Number.isFinite(pbt) || pbt === 0 || !Number.isFinite(tax)) return null;
    return tax / pbt;
  });

  const roicSeries = data.metrics.pbt.map((pbt, index) => {
    const interest = data.metrics.interest[index];
    const taxRate = taxRateSeries[index] ?? 0.25;
    const ebit = Number.isFinite(pbt) ? pbt + (interest ?? 0) : null;
    const nopat = Number.isFinite(ebit) ? ebit * (1 - taxRate) : null;
    const capital = (equityValues[index] ?? 0) + (data.balanceSheet.borrowings[index] ?? 0) - (data.balanceSheet.cash[index] ?? 0);
    if (!Number.isFinite(nopat) || !Number.isFinite(capital) || capital === 0) return null;
    return (nopat / capital) * 100;
  });

  const capYears = roicSeries.filter((value) => Number.isFinite(value) && value > wacc).length;
  const capSignal = capYears >= 10 ? "Strong CAP (10+ years)" : capYears >= 5 ? "Moderate CAP (5-10 years)" : "Weak CAP (<5 years)";
  capDurationEl.textContent = capYears ? `${capYears} years` : "-";
  capSignalEl.textContent = capYears ? capSignal : "-";

  destroyChart(capChart);
  capChart = new Chart(capChartEl, {
    type: "line",
    data: {
      labels,
      datasets: [
        { label: "ROIC %", data: roicSeries, borderColor: "#2b6cff" },
        { label: "WACC %", data: labels.map(() => wacc), borderColor: "#dc2626", borderDash: [4, 4] },
      ],
    },
    options: { responsive: true, maintainAspectRatio: false },
  });

  const patCagr3 = calculateCAGR(data.metrics.netProfit, 3);
  const cfo = data.cashFlow.cfo.filter((val) => Number.isFinite(val));
  const pat = data.metrics.netProfit.filter((val) => Number.isFinite(val));
  const cfoPatRatio = cfo.length && pat.length ? getAverage(cfo.slice(-3)) / getAverage(pat.slice(-3)) : null;

  epbCagrEl.textContent = formatPercent(patCagr3);
  epbCashEl.textContent = cfoPatRatio ? formatPercent(cfoPatRatio * 100) : "-";

  let quadrant = "N/A";
  if (Number.isFinite(patCagr3) && Number.isFinite(cfoPatRatio)) {
    quadrant = patCagr3 > 15 && cfoPatRatio > 0.8
      ? "⭐ Star (High Growth + Cash)"
      : patCagr3 > 15 && cfoPatRatio <= 0.8
      ? "Investigate (High Growth, Low Cash)"
      : patCagr3 <= 15 && cfoPatRatio > 0.8
      ? "Cash Cow (Low Growth + Cash)"
      : "Red Flag (Low Growth + Low Cash)";
  }
  epbQuadrantEl.textContent = quadrant;

  destroyChart(epbChart);
  epbChart = new Chart(epbChartEl, {
    type: "scatter",
    data: {
      datasets: [
        {
          label: "Company",
          data: Number.isFinite(patCagr3) && Number.isFinite(cfoPatRatio)
            ? [{ x: patCagr3, y: cfoPatRatio * 100 }]
            : [],
          backgroundColor: "#2b6cff",
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: { title: { display: true, text: "PAT CAGR (3Y) %" } },
        y: { title: { display: true, text: "CFO / PAT %" } },
      },
    },
  });

  const redFlags = [];
  const salesSeries = data.metrics.sales;
  const receivablesSeries = data.balanceSheet.receivables;
  const inventorySeries = data.balanceSheet.inventory;
  const otherAssetsSeries = data.balanceSheet.otherAssets;
  const totalAssetsSeries = data.balanceSheet.totalAssets;
  const otherIncomeSeries = data.metrics.otherIncome;
  const operatingProfitSeries = data.metrics.sales.map((salesValue, index) => {
    if (!Number.isFinite(salesValue)) return null;
    const expenses = [
      data.metrics.rawMaterial[index],
      data.metrics.inventoryChange[index],
      data.metrics.powerFuel[index],
      data.metrics.otherMfg[index],
      data.metrics.employeeCost[index],
      data.metrics.sellingAdmin[index],
      data.metrics.otherExpenses[index],
    ];
    if (expenses.every((value) => value === null)) return null;
    const totalExpenses = expenses.reduce((sum, value) => (Number.isFinite(value) ? sum + value : sum), 0);
    return salesValue - totalExpenses;
  });

  const latestSales = getLatestValue(salesSeries);
  const latestReceivables = getLatestValue(receivablesSeries);
  const latestInventory = getLatestValue(inventorySeries);
  const latestOtherAssets = getLatestValue(otherAssetsSeries);
  const latestTotalAssets = getLatestValue(totalAssetsSeries);
  const latestOtherIncome = getLatestValue(otherIncomeSeries);
  const latestOperatingProfit = getLatestValue(operatingProfitSeries);
  const latestBorrowings = getLatestValue(data.balanceSheet.borrowings);

  if (Number.isFinite(latestReceivables) && Number.isFinite(latestSales)) {
    const receivableIntensity = latestReceivables / latestSales;
    if (receivableIntensity > 0.25) {
      redFlags.push(`Receivables high (${(receivableIntensity * 100).toFixed(0)}% of sales)`);
    }
  }
  if (Number.isFinite(latestInventory) && Number.isFinite(latestSales)) {
    const inventoryIntensity = latestInventory / latestSales;
    if (inventoryIntensity > 0.2) {
      redFlags.push(`Inventory buildup (${(inventoryIntensity * 100).toFixed(0)}% of sales)`);
    }
  }
  if (Number.isFinite(latestOtherAssets) && Number.isFinite(latestTotalAssets)) {
    const otherAssetsShare = latestOtherAssets / latestTotalAssets;
    if (otherAssetsShare > 0.2) {
      redFlags.push(`Other assets elevated (${(otherAssetsShare * 100).toFixed(0)}% of assets)`);
    }
  }
  if (Number.isFinite(latestOtherIncome) && Number.isFinite(latestOperatingProfit) && latestOperatingProfit > 0) {
    const otherIncomeShare = latestOtherIncome / latestOperatingProfit;
    if (otherIncomeShare > 0.5) {
      redFlags.push(`Other income > 50% of operating profit`);
    }
  }

  renderMetricList(redFlagsEl, redFlags.length ? redFlags.map((flag) => ({ label: "Flag", value: flag })) : [
    { label: "Flags", value: "No major flags detected" },
  ]);

  const latestAssets = getLatestValue(totalAssetsSeries);
  const latestEquity = getLatestValue(equityValues);
  const latestProfit = getLatestValue(data.metrics.netProfit);
  const netMargin = calculateMargin(latestProfit, latestSales);
  const assetTurnover = calculateRatio(latestSales, latestAssets);
  const equityMultiplier = calculateRatio(latestAssets, latestEquity);
  const debtToEquity = calculateRatio(latestBorrowings, latestEquity);

  renderMetricList(dupontMetricsEl, [
    { label: "Net Margin", value: formatPercent(netMargin) },
    { label: "Asset Turnover", value: formatRatio(assetTurnover) },
    { label: "Equity Multiplier", value: formatRatio(equityMultiplier) },
  ]);

  let qualityScore = 0;
  if (Number.isFinite(netMargin) && netMargin > 15) qualityScore += 20;
  const latestRoic = roicSeries.filter((val) => Number.isFinite(val)).at(-1);
  if (Number.isFinite(latestRoic) && latestRoic > 15) qualityScore += 20;
  const cfoMargin = Number.isFinite(cfoPatRatio) ? cfoPatRatio : null;
  if (Number.isFinite(cfoMargin) && cfoMargin > 0.8) qualityScore += 20;
  if (Number.isFinite(assetTurnover) && assetTurnover > 1.5) qualityScore += 20;
  if (Number.isFinite(debtToEquity) && debtToEquity < 0.7) qualityScore += 20;

  renderMetricList(qualityScoreEl, [
    { label: "Score (out of 100)", value: Number.isFinite(qualityScore) ? `${qualityScore}` : "-" },
    { label: "Classification", value: qualityScore >= 80 ? "High Quality" : qualityScore >= 60 ? "Above Average" : "Needs Review" },
  ]);

  const retainedEarnings = data.metrics.netProfit.reduce((sum, profit, index) => {
    const dividend = data.metrics.dividend[index];
    if (!Number.isFinite(profit)) return sum;
    return sum + profit - (dividend ?? 0);
  }, 0);
  const firstEquity = equityValues.find((val) => Number.isFinite(val));
  const lastEquity = getLatestValue(equityValues);
  const bookValueChange = Number.isFinite(firstEquity) && Number.isFinite(lastEquity) ? lastEquity - firstEquity : null;
  const buffettRatio = Number.isFinite(bookValueChange) && retainedEarnings !== 0 ? bookValueChange / retainedEarnings : null;

  renderMetricList(buffettTestEl, [
    { label: "Retained Earnings", value: formatNumber(retainedEarnings) },
    { label: "Book Value Change", value: formatNumber(bookValueChange) },
    { label: "Value Creation Ratio", value: buffettRatio ? buffettRatio.toFixed(2) : "-" },
  ]);
};

const validateWorkbook = (workbook) => {
  const sheetNames = workbook.SheetNames;
  const missing = REQUIRED_SHEETS.filter((sheet) => !sheetNames.includes(sheet));
  if (missing.length) {
    return `Missing required sheets: ${missing.join(", ")}`;
  }
  return null;
};

const handleFile = async (file) => {
  uploadStatus.textContent = "Parsing file...";
  try {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array" });
    const validationError = validateWorkbook(workbook);
    if (validationError) {
      uploadStatus.textContent = validationError;
      return;
    }

    const dataSheet = workbook.Sheets["Data Sheet"];
    if (!dataSheet) {
      uploadStatus.textContent = "Could not find 'Data Sheet'.";
      return;
    }

    const parsed = parseDataSheet(dataSheet);
    if (parsed.dates.length < 2) {
      uploadStatus.textContent = "Insufficient data: This file has fewer than 2 years of data.";
    } else {
      uploadStatus.textContent = `Loaded ${file.name}`;
    }

    updateOverview(parsed);
    buildPLTable(parsed);
    buildBSTable(parsed);
    buildCFTable(parsed);
    buildAnalysis(parsed);
    buildCharts(parsed);
    buildFrameworks(parsed);
  } catch (error) {
    uploadStatus.textContent = "Could not parse Excel file. Please ensure it's a screener.in export.";
    console.error(error);
  }
};

fileInput.addEventListener("change", (event) => {
  const file = event.target.files?.[0];
  if (!file) return;
  if (!file.name.endsWith(".xlsx")) {
    uploadStatus.textContent = "Please upload a valid .xlsx file.";
    return;
  }
  handleFile(file);
});

TAB_BUTTONS.forEach((button) => {
  button.addEventListener("click", () => {
    if (button.disabled) return;
    TAB_BUTTONS.forEach((btn) => btn.classList.remove("active"));
    TAB_PANELS.forEach((panel) => panel.classList.remove("active"));
    button.classList.add("active");
    const panel = document.getElementById(button.dataset.tab);
    if (panel) {
      panel.classList.add("active");
    }
  });
});
