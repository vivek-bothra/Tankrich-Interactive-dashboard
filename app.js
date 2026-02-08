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

const moatScoreEl = document.getElementById("moatScore");
const moatBreakdownEl = document.getElementById("moatBreakdown");
const capexSplitEl = document.getElementById("capexSplit");
const incrementalRoicEl = document.getElementById("incrementalRoic");
const capitalAllocationScoreEl = document.getElementById("capitalAllocationScore");
const capitalAllocationBreakdownEl = document.getElementById("capitalAllocationBreakdown");
const valueDriversEl = document.getElementById("valueDrivers");
const rawMaterialSensitivityEl = document.getElementById("rawMaterialSensitivity");
const valueMigrationEl = document.getElementById("valueMigration");
const qualityScoreEl = document.getElementById("qualityScore");
const qualityBreakdownEl = document.getElementById("qualityBreakdown");

const plTableEl = document.getElementById("plTable");
const bsTableEl = document.getElementById("bsTable");
const cfTableEl = document.getElementById("cfTable");

const TAB_BUTTONS = document.querySelectorAll(".tab-button");
const TAB_PANELS = document.querySelectorAll(".tab-panel");

const REQUIRED_SHEETS = ["Data Sheet", "Profit & Loss", "Balance Sheet", "Cash Flow"];

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

const normalizeLabel = (value) => String(value ?? "")
  .toLowerCase()
  .replace(/[^a-z0-9]/g, "");

const findRowIndex = (rows, candidates, start = 0, end = rows.length, occurrence = 1) => {
  const normalizedCandidates = candidates.map(normalizeLabel);
  let matched = 0;
  for (let i = start; i < Math.min(end, rows.length); i += 1) {
    const rowLabel = normalizeLabel(rows[i]?.[0]);
    if (!rowLabel) continue;
    if (normalizedCandidates.some((candidate) => rowLabel.includes(candidate))) {
      matched += 1;
      if (matched === occurrence) return i;
    }
  }
  return -1;
};

const getValueByLabel = (rows, candidates, defaultValue = null, start = 0, end = rows.length, occurrence = 1) => {
  const index = findRowIndex(rows, candidates, start, end, occurrence);
  if (index < 0) return defaultValue;
  return rows[index]?.[1] ?? defaultValue;
};

const getSeriesByLabel = (rows, candidates, start = 0, end = rows.length, occurrence = 1) => {
  const index = findRowIndex(rows, candidates, start, end, occurrence);
  return index >= 0 ? getRowSlice(rows, index) : [];
};

const formatReportDate = (value) => {
  if (value === null || value === undefined || value === "") return "-";

  if (Number.isFinite(value)) {
    const parsed = XLSX?.SSF?.parse_date_code?.(value);
    if (parsed?.y) return String(parsed.y);
    return String(value);
  }

  if (typeof value === "string") {
    const trimmed = value.trim();
    if (/^\d+(\.\d+)?$/.test(trimmed)) {
      const parsed = XLSX?.SSF?.parse_date_code?.(Number(trimmed));
      if (parsed?.y) return String(parsed.y);
    }
    const yearMatch = trimmed.match(/\b(19|20)\d{2}\b/);
    return yearMatch ? yearMatch[0] : trimmed;
  }

  return String(value);
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

const getMedian = (values) => {
  const filtered = values.filter((val) => Number.isFinite(val)).sort((a, b) => a - b);
  if (!filtered.length) return null;
  const mid = Math.floor(filtered.length / 2);
  return filtered.length % 2 ? filtered[mid] : (filtered[mid - 1] + filtered[mid]) / 2;
};

const getStdDev = (values) => {
  const avg = getAverage(values);
  const filtered = values.filter((val) => Number.isFinite(val));
  if (!Number.isFinite(avg) || filtered.length < 2) return null;
  const variance = filtered.reduce((sum, value) => sum + ((value - avg) ** 2), 0) / filtered.length;
  return Math.sqrt(variance);
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

const clamp = (value, min, max) => Math.min(Math.max(value, min), max);

const scoreFromThresholds = (value, thresholds = []) => {
  if (!Number.isFinite(value) || !thresholds.length) return 0;
  let score = 0;
  thresholds.forEach(({ min, points }) => {
    if (value >= min) score = points;
  });
  return score;
};

const scoreFromInverseThresholds = (value, thresholds = []) => {
  if (!Number.isFinite(value) || !thresholds.length) return 0;
  let score = 0;
  thresholds.forEach(({ max, points }) => {
    if (value <= max) score = points;
  });
  return score;
};

const formatRatio = (value) => {
  if (value === null || value === undefined) return "-";
  return `${value.toFixed(2)}x`;
};

const formatScore = (value, maxScore = 100) => {
  if (!Number.isFinite(value)) return "-";
  return `${Math.round(value)}/${maxScore}`;
};

const getSeriesFirstAndLast = (series) => {
  const filtered = series
    .map((value, index) => ({ value, index }))
    .filter(({ value }) => Number.isFinite(value));
  if (filtered.length < 2) return null;
  return {
    first: filtered[0].value,
    last: filtered[filtered.length - 1].value,
    span: filtered[filtered.length - 1].index - filtered[0].index,
  };
};

const calculateIncrementalRatio = (numeratorSeries, denominatorSeries, periodsBack) => {
  const pairs = numeratorSeries
    .map((value, index) => ({ num: value, den: denominatorSeries[index] }))
    .filter(({ num, den }) => Number.isFinite(num) && Number.isFinite(den));
  if (pairs.length <= periodsBack) return null;
  const end = pairs[pairs.length - 1];
  const start = pairs[pairs.length - 1 - periodsBack];
  const deltaDen = end.den - start.den;
  if (!Number.isFinite(deltaDen) || deltaDen === 0) return null;
  return (end.num - start.num) / deltaDen;
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

  const plStart = findRowIndex(rows, ["profitloss"]);
  const quartersStart = findRowIndex(rows, ["quarters"], plStart + 1);
  const bsStart = findRowIndex(rows, ["balancesheet"], (quartersStart >= 0 ? quartersStart + 1 : plStart + 1));
  const cfStart = findRowIndex(rows, ["cashflow"], (bsStart >= 0 ? bsStart + 1 : 0));

  const metaEnd = plStart > 0 ? plStart : rows.length;
  const plEnd = quartersStart > 0 ? quartersStart : (bsStart > 0 ? bsStart : rows.length);
  const bsEnd = cfStart > 0 ? cfStart : rows.length;
  const cfEnd = rows.length;

  const meta = {
    name: getValueByLabel(rows, ["companyname"], "N/A", 0, metaEnd),
    faceValue: safeNumber(getValueByLabel(rows, ["facevalue"], null, 0, metaEnd)),
    currentPrice: safeNumber(getValueByLabel(rows, ["currentprice"], null, 0, metaEnd)),
    marketCap: safeNumber(getValueByLabel(rows, ["marketcapitalization", "marketcap"], null, 0, metaEnd)),
  };

  const dates = getSeriesByLabel(rows, ["reportdate"], plStart, plEnd).map(formatReportDate);

  const metrics = {
    sales: getSeriesByLabel(rows, ["sales"], plStart, plEnd).map(safeNumber),
    rawMaterial: getSeriesByLabel(rows, ["rawmaterialcost"], plStart, plEnd).map(safeNumber),
    inventoryChange: getSeriesByLabel(rows, ["changeininventory"], plStart, plEnd).map(safeNumber),
    powerFuel: getSeriesByLabel(rows, ["powerandfuel", "powerfuel"], plStart, plEnd).map(safeNumber),
    otherMfg: getSeriesByLabel(rows, ["othermfrexp", "othermanufacturing"], plStart, plEnd).map(safeNumber),
    employeeCost: getSeriesByLabel(rows, ["employeecost"], plStart, plEnd).map(safeNumber),
    sellingAdmin: getSeriesByLabel(rows, ["sellingandadmin", "sellingadmin"], plStart, plEnd).map(safeNumber),
    otherExpenses: getSeriesByLabel(rows, ["otherexpenses"], plStart, plEnd).map(safeNumber),
    otherIncome: getSeriesByLabel(rows, ["otherincome"], plStart, plEnd).map(safeNumber),
    depreciation: getSeriesByLabel(rows, ["depreciation"], plStart, plEnd).map(safeNumber),
    interest: getSeriesByLabel(rows, ["interest"], plStart, plEnd).map(safeNumber),
    pbt: getSeriesByLabel(rows, ["profitbeforetax"], plStart, plEnd).map(safeNumber),
    tax: getSeriesByLabel(rows, ["tax"], plStart, plEnd).map(safeNumber),
    netProfit: getSeriesByLabel(rows, ["netprofit"], plStart, plEnd).map(safeNumber),
    dividend: getSeriesByLabel(rows, ["dividendamount", "dividend"], plStart, plEnd).map(safeNumber),
  };

  const balanceSheetDates = getSeriesByLabel(rows, ["reportdate"], bsStart, bsEnd).map(formatReportDate);
  const balanceSheet = {
    equity: getSeriesByLabel(rows, ["equitysharecapital"], bsStart, bsEnd).map(safeNumber),
    reserves: getSeriesByLabel(rows, ["reserves"], bsStart, bsEnd).map(safeNumber),
    borrowings: getSeriesByLabel(rows, ["borrowings"], bsStart, bsEnd).map(safeNumber),
    otherLiabilities: getSeriesByLabel(rows, ["otherliabilities"], bsStart, bsEnd).map(safeNumber),
    totalLiabilities: getSeriesByLabel(rows, ["total"], bsStart, bsEnd, 1).map(safeNumber),
    netBlock: getSeriesByLabel(rows, ["netblock"], bsStart, bsEnd).map(safeNumber),
    cwip: getSeriesByLabel(rows, ["capitalworkinprogress", "cwip"], bsStart, bsEnd).map(safeNumber),
    investments: getSeriesByLabel(rows, ["investments"], bsStart, bsEnd).map(safeNumber),
    otherAssets: getSeriesByLabel(rows, ["otherassets"], bsStart, bsEnd).map(safeNumber),
    totalAssets: getSeriesByLabel(rows, ["total"], bsStart, bsEnd, 2).map(safeNumber),
    receivables: getSeriesByLabel(rows, ["receivables"], bsStart, bsEnd).map(safeNumber),
    inventory: getSeriesByLabel(rows, ["inventory"], bsStart, bsEnd).map(safeNumber),
    cash: getSeriesByLabel(rows, ["cashbank", "cashandbank"], bsStart, bsEnd).map(safeNumber),
    shares: getSeriesByLabel(rows, ["noofequityshares"], bsStart, bsEnd).map(safeNumber),
  };
  if (!balanceSheet.totalAssets.length && balanceSheet.totalLiabilities.length) {
    balanceSheet.totalAssets = [...balanceSheet.totalLiabilities];
  }

  const cashFlowDates = getSeriesByLabel(rows, ["reportdate"], cfStart, cfEnd).map(formatReportDate);
  const cashFlow = {
    cfo: getSeriesByLabel(rows, ["cashfromoperatingactivity"], cfStart, cfEnd).map(safeNumber),
    cfi: getSeriesByLabel(rows, ["cashfrominvestingactivity"], cfStart, cfEnd).map(safeNumber),
    cff: getSeriesByLabel(rows, ["cashfromfinancingactivity"], cfStart, cfEnd).map(safeNumber),
    netCash: getSeriesByLabel(rows, ["netcashflow"], cfStart, cfEnd).map(safeNumber),
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
  const equityValues = data.balanceSheet.equity.map((equity, index) => {
    const reserves = data.balanceSheet.reserves[index];
    if (!Number.isFinite(equity) && !Number.isFinite(reserves)) return null;
    return (equity ?? 0) + (reserves ?? 0);
  });
  const capitalEmployedSeries = equityValues.map((equity, index) => {
    const borrowings = data.balanceSheet.borrowings[index];
    if (!Number.isFinite(equity) && !Number.isFinite(borrowings)) return null;
    return (equity ?? 0) + (borrowings ?? 0);
  });

  const roceSeries = data.metrics.netProfit.map((profitValue, index) =>
    calculateMargin(profitValue, capitalEmployedSeries[index])
  );
  const avgRoce = getAverage(roceSeries);
  const roceVolatility = getStdDev(roceSeries);
  const salesCagr5 = calculateCAGR(data.metrics.sales, 5);
  const marginSeries = data.metrics.sales.map((sales, index) =>
    calculateMargin(data.metrics.netProfit[index], sales)
  );
  const marginVolatility = getStdDev(marginSeries);
  const debtToEquitySeries = data.balanceSheet.borrowings.map((borrowings, index) =>
    calculateRatio(borrowings, equityValues[index])
  );
  const latestDebtToEquity = getLatestValue(debtToEquitySeries);

  const moatRoceScore = scoreFromThresholds(avgRoce, [
    { min: 8, points: 10 },
    { min: 12, points: 20 },
    { min: 18, points: 30 },
  ]);
  const moatGrowthScore = scoreFromThresholds(salesCagr5, [
    { min: 5, points: 8 },
    { min: 10, points: 15 },
    { min: 15, points: 20 },
  ]);
  const moatMarginStabilityScore = scoreFromInverseThresholds(marginVolatility, [
    { max: 8, points: 8 },
    { max: 6, points: 12 },
    { max: 4, points: 18 },
  ]);
  const moatLeverageScore = scoreFromInverseThresholds(latestDebtToEquity, [
    { max: 1.2, points: 8 },
    { max: 0.8, points: 12 },
    { max: 0.4, points: 16 },
  ]);
  const moatScore = moatRoceScore + moatGrowthScore + moatMarginStabilityScore + moatLeverageScore;
  moatScoreEl.textContent = formatScore(moatScore, 84);

  renderMetricList(moatBreakdownEl, [
    { label: "Avg ROCE", value: formatPercent(avgRoce) },
    { label: "ROCE Volatility", value: formatPercent(roceVolatility) },
    { label: "Sales CAGR (5Y)", value: formatPercent(salesCagr5) },
    { label: "Margin Volatility", value: formatPercent(marginVolatility) },
    { label: "Debt-to-Equity (Latest)", value: formatRatio(latestDebtToEquity) },
  ]);

  const grossCapexSeries = data.cashFlow.cfi.map((cfi) => (Number.isFinite(cfi) ? Math.max(-cfi, 0) : null));
  const maintenanceCapexSeries = grossCapexSeries.map((grossCapex, index) => {
    const depreciation = data.metrics.depreciation[index];
    if (!Number.isFinite(grossCapex) || !Number.isFinite(depreciation)) return null;
    return Math.min(grossCapex, Math.max(depreciation, 0));
  });
  const growthCapexSeries = grossCapexSeries.map((grossCapex, index) => {
    const maintenance = maintenanceCapexSeries[index];
    if (!Number.isFinite(grossCapex) || !Number.isFinite(maintenance)) return null;
    return Math.max(grossCapex - maintenance, 0);
  });
  const latestGrossCapex = getLatestValue(grossCapexSeries);
  const latestMaintenanceCapex = getLatestValue(maintenanceCapexSeries);
  const latestGrowthCapex = getLatestValue(growthCapexSeries);
  const maintenanceShare = calculateMargin(latestMaintenanceCapex, latestGrossCapex);
  const growthShare = calculateMargin(latestGrowthCapex, latestGrossCapex);

  renderMetricList(capexSplitEl, [
    { label: "Gross Capex (Latest)", value: formatNumber(latestGrossCapex) },
    { label: "Maintenance Capex (Proxy)", value: formatNumber(latestMaintenanceCapex) },
    { label: "Growth Capex (Proxy)", value: formatNumber(latestGrowthCapex) },
    { label: "Maintenance Share", value: formatPercent(maintenanceShare) },
    { label: "Growth Share", value: formatPercent(growthShare) },
  ]);

  const nopatSeries = data.metrics.pbt.map((pbt, index) => {
    const tax = data.metrics.tax[index];
    if (!Number.isFinite(pbt) && !Number.isFinite(tax)) return null;
    return (pbt ?? 0) - (tax ?? 0);
  });
  const incrementalRoic3y = calculateIncrementalRatio(nopatSeries, capitalEmployedSeries, 3);
  const incrementalRoic5y = calculateIncrementalRatio(nopatSeries, capitalEmployedSeries, 5);
  const reinvestmentRate = calculateRatio(getLatestValue(growthCapexSeries), getLatestValue(data.metrics.netProfit));

  renderMetricList(incrementalRoicEl, [
    {
      label: "Incremental ROIC (3Y)",
      value: formatPercent(Number.isFinite(incrementalRoic3y) ? incrementalRoic3y * 100 : null),
    },
    {
      label: "Incremental ROIC (5Y)",
      value: formatPercent(Number.isFinite(incrementalRoic5y) ? incrementalRoic5y * 100 : null),
    },
    {
      label: "Growth Reinvestment Rate",
      value: formatPercent(Number.isFinite(reinvestmentRate) ? reinvestmentRate * 100 : null),
    },
  ]);

  const cfoSeries = data.cashFlow.cfo;
  const cashConversion = calculateRatio(getLatestValue(cfoSeries), getLatestValue(data.metrics.netProfit));
  const dividendPayout = calculateRatio(getLatestValue(data.metrics.dividend), getLatestValue(data.metrics.netProfit));
  const balanceSheetDiscipline = scoreFromInverseThresholds(latestDebtToEquity, [
    { max: 1.2, points: 10 },
    { max: 0.8, points: 15 },
    { max: 0.4, points: 20 },
  ]);
  const reinvestmentScore = scoreFromThresholds(Number.isFinite(incrementalRoic3y) ? incrementalRoic3y * 100 : null, [
    { min: 8, points: 10 },
    { min: 12, points: 15 },
    { min: 18, points: 20 },
  ]);
  const cashConversionScore = scoreFromThresholds(cashConversion, [
    { min: 0.7, points: 10 },
    { min: 1, points: 15 },
    { min: 1.2, points: 20 },
  ]);
  const payoutBalanceScore = Number.isFinite(dividendPayout)
    ? (dividendPayout >= 0.1 && dividendPayout <= 0.6 ? 20 : 10)
    : 0;
  const capitalAllocationScore = reinvestmentScore + cashConversionScore + balanceSheetDiscipline + payoutBalanceScore;
  capitalAllocationScoreEl.textContent = formatScore(capitalAllocationScore, 80);

  renderMetricList(capitalAllocationBreakdownEl, [
    { label: "Cash Conversion (CFO/Profit)", value: formatRatio(cashConversion) },
    {
      label: "Dividend Payout (Latest)",
      value: formatPercent(Number.isFinite(dividendPayout) ? dividendPayout * 100 : null),
    },
    { label: "Debt-to-Equity (Latest)", value: formatRatio(latestDebtToEquity) },
    {
      label: "Reinvestment Efficiency",
      value: formatPercent(Number.isFinite(incrementalRoic3y) ? incrementalRoic3y * 100 : null),
    },
  ]);

  const medianMargin = getMedian(marginSeries);
  const medianRoce = getMedian(roceSeries);
  const marginExpansion = (() => {
    const points = getSeriesFirstAndLast(marginSeries);
    if (!points) return null;
    return points.last - points.first;
  })();
  const leverageTrend = (() => {
    const points = getSeriesFirstAndLast(debtToEquitySeries);
    if (!points) return null;
    return points.last - points.first;
  })();

  renderMetricList(valueDriversEl, [
    { label: "Growth Driver (Sales CAGR 5Y)", value: formatPercent(salesCagr5) },
    { label: "Margin Structure (Median NPM)", value: formatPercent(medianMargin) },
    { label: "Return Structure (Median ROCE)", value: formatPercent(medianRoce) },
    { label: "Margin Expansion Trend", value: formatPercent(marginExpansion) },
    { label: "Leverage Trend", value: formatRatio(leverageTrend) },
  ]);

  const rmToSalesSeries = data.metrics.rawMaterial.map((rawMaterial, index) =>
    calculateMargin(rawMaterial, data.metrics.sales[index])
  );
  const latestRmRatio = getLatestValue(rmToSalesSeries);
  const avgRmRatio = getAverage(rmToSalesSeries);
  const rmTrend = (() => {
    const points = getSeriesFirstAndLast(rmToSalesSeries);
    if (!points) return null;
    return points.last - points.first;
  })();
  const stressMargin = Number.isFinite(latestRmRatio)
    ? latestRmRatio * 1.1
    : null;

  renderMetricList(rawMaterialSensitivityEl, [
    { label: "Raw Material / Sales (Latest)", value: formatPercent(latestRmRatio) },
    { label: "Raw Material / Sales (Average)", value: formatPercent(avgRmRatio) },
    { label: "Raw Material Trend", value: formatPercent(rmTrend) },
    { label: "+10% RM Shock (proxy ratio)", value: formatPercent(stressMargin) },
  ]);

  const valueMigrationLabel = (() => {
    if (!Number.isFinite(marginExpansion) || !Number.isFinite(leverageTrend) || !Number.isFinite(avgRoce)) return "-";
    if (marginExpansion > 1 && leverageTrend <= 0 && avgRoce >= 15) return "Value Accretive";
    if (marginExpansion >= 0 && leverageTrend <= 0.2 && avgRoce >= 10) return "Neutral";
    return "Value Dilutive";
  })();
  const marketCapPerProfit = calculateRatio(data.meta.marketCap, getLatestValue(data.metrics.netProfit));

  renderMetricList(valueMigrationEl, [
    { label: "Migration Assessment", value: valueMigrationLabel },
    { label: "Margin Expansion", value: formatPercent(marginExpansion) },
    { label: "Leverage Drift", value: formatRatio(leverageTrend) },
    { label: "Market Cap / Profit", value: formatRatio(marketCapPerProfit) },
  ]);

  const growthQualityScore = scoreFromThresholds(salesCagr5, [
    { min: 5, points: 10 },
    { min: 10, points: 20 },
    { min: 15, points: 25 },
  ]);
  const returnQualityScore = scoreFromThresholds(avgRoce, [
    { min: 8, points: 10 },
    { min: 12, points: 20 },
    { min: 18, points: 25 },
  ]);
  const cashQualityScore = scoreFromThresholds(cashConversion, [
    { min: 0.7, points: 10 },
    { min: 1, points: 20 },
    { min: 1.2, points: 25 },
  ]);
  const balanceSheetQualityScore = scoreFromInverseThresholds(latestDebtToEquity, [
    { max: 1.2, points: 10 },
    { max: 0.8, points: 20 },
    { max: 0.4, points: 25 },
  ]);
  const qualityScore = clamp(
    growthQualityScore + returnQualityScore + cashQualityScore + balanceSheetQualityScore,
    0,
    100
  );
  qualityScoreEl.textContent = formatScore(qualityScore, 100);

  renderMetricList(qualityBreakdownEl, [
    { label: "Growth Quality", value: `${growthQualityScore}/25` },
    { label: "Return Quality", value: `${returnQualityScore}/25` },
    { label: "Cash Quality", value: `${cashQualityScore}/25` },
    { label: "Balance Sheet Quality", value: `${balanceSheetQualityScore}/25` },
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
