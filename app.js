// Tankrich Dashboard - Main Application Logic

// Global state
let companyData = null;

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    const fileInputWelcome = document.getElementById('fileInputWelcome');
    
    fileInput.addEventListener('change', handleFileUpload);
    fileInputWelcome.addEventListener('change', handleFileUpload);
    
    // Tab switching
    document.querySelectorAll('.tab').forEach(tab => {
        tab.addEventListener('click', () => switchTab(tab.dataset.tab));
    });
});

// File Upload Handler
async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    document.getElementById('fileName').textContent = file.name;
    document.getElementById('welcomeScreen').classList.add('hidden');
    document.getElementById('loadingState').classList.remove('hidden');
    
    try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data);
        
        // Parse the Excel file
        companyData = parseExcelData(workbook);
        
        // Display data
        displayDashboard();
        
    } catch (error) {
        console.error('Error parsing file:', error);
        alert('Error parsing Excel file. Please ensure it\'s a valid screener.in export.');
        document.getElementById('loadingState').classList.add('hidden');
        document.getElementById('welcomeScreen').classList.remove('hidden');
    }
}

// Parse Excel Data
function parseExcelData(workbook) {
    const dataSheet = workbook.Sheets['Data Sheet'];
    if (!dataSheet) {
        throw new Error('Data Sheet not found');
    }
    
    // Convert sheet to array of arrays
    const raw = XLSX.utils.sheet_to_json(dataSheet, { header: 1, defval: null });
    
    // Extract company meta
    const meta = {
        name: raw[0][1] || 'Unknown Company',
        faceValue: raw[6][1] || null,
        currentPrice: raw[7][1] || null,
        marketCap: raw[8][1] || null
    };
    
    // Extract annual data
    const reportDates = (raw[15] || []).slice(4).filter(d => d);
    const years = reportDates.map(d => new Date(d).getFullYear());
    
    const annual = {
        years: years,
        dates: reportDates,
        sales: extractRow(raw, 16),
        rawMaterial: extractRow(raw, 17),
        changeInventory: extractRow(raw, 18),
        powerFuel: extractRow(raw, 19),
        otherMfg: extractRow(raw, 20),
        employeeCost: extractRow(raw, 21),
        sellingAdmin: extractRow(raw, 22),
        otherExpenses: extractRow(raw, 23),
        otherIncome: extractRow(raw, 24),
        depreciation: extractRow(raw, 25),
        interest: extractRow(raw, 26),
        pbt: extractRow(raw, 27),
        tax: extractRow(raw, 28),
        netProfit: extractRow(raw, 29),
        dividend: extractRow(raw, 30),
        
        // Balance Sheet
        equity: extractRow(raw, 56),
        reserves: extractRow(raw, 57),
        borrowings: extractRow(raw, 58),
        otherLiabilities: extractRow(raw, 59),
        totalLiabilities: extractRow(raw, 60),
        netBlock: extractRow(raw, 61),
        cwip: extractRow(raw, 62),
        investments: extractRow(raw, 63),
        otherAssets: extractRow(raw, 64),
        totalAssets: extractRow(raw, 65),
        receivables: extractRow(raw, 66),
        inventory: extractRow(raw, 67),
        cash: extractRow(raw, 68),
        shares: extractRow(raw, 69),
        
        // Cash Flow
        cfo: extractRow(raw, 81),
        cfi: extractRow(raw, 82),
        cff: extractRow(raw, 83),
        netCashFlow: extractRow(raw, 84),
        
        // Prices
        prices: extractRow(raw, 89)
    };
    
    return { meta, annual };
}

function extractRow(raw, rowIndex) {
    if (!raw[rowIndex]) return [];
    return raw[rowIndex].slice(4).map(v => {
        if (v === null || v === undefined || v === '') return null;
        const num = parseFloat(v);
        return isNaN(num) ? null : num;
    });
}

// Display Dashboard
function displayDashboard() {
    document.getElementById('loadingState').classList.add('hidden');
    document.getElementById('companyHeader').classList.remove('hidden');
    document.getElementById('mainContent').classList.remove('hidden');
    
    // Update company header
    document.getElementById('companyName').textContent = companyData.meta.name;
    document.getElementById('currentPrice').textContent = formatCurrency(companyData.meta.currentPrice);
    document.getElementById('marketCap').textContent = formatLargeNumber(companyData.meta.marketCap) + ' Cr';
    document.getElementById('latestYear').textContent = companyData.annual.years[companyData.annual.years.length - 1] || '-';
    
    // Calculate and display all metrics
    calculateAndDisplayMetrics();
    
    // Display financial statements
    displayFinancialStatements();
    
    // Display charts
    displayCharts();
}

// Calculate and Display Metrics
function calculateAndDisplayMetrics() {
    const { annual } = companyData;
    const n = annual.years.length;
    
    // Growth Metrics
    const salesCAGR5 = calculateCAGR(annual.sales, 5);
    const salesCAGR10 = calculateCAGR(annual.sales, 10);
    const profitCAGR5 = calculateCAGR(annual.netProfit, 5);
    const profitCAGR10 = calculateCAGR(annual.netProfit, 10);
    
    document.getElementById('salesCAGR5').textContent = formatPercent(salesCAGR5);
    document.getElementById('salesCAGR10').textContent = formatPercent(salesCAGR10);
    document.getElementById('profitCAGR5').textContent = formatPercent(profitCAGR5);
    document.getElementById('profitCAGR10').textContent = formatPercent(profitCAGR10);
    
    // Profitability Ratios (Latest Year)
    const latestSales = annual.sales[n-1];
    const latestProfit = annual.netProfit[n-1];
    const latestEquity = (annual.equity[n-1] || 0) + (annual.reserves[n-1] || 0);
    const prevEquity = (annual.equity[n-2] || 0) + (annual.reserves[n-2] || 0);
    const avgEquity = (latestEquity + prevEquity) / 2;
    
    const latestDebt = annual.borrowings[n-1] || 0;
    const capitalEmployed = avgEquity + latestDebt;
    
    const ebit = (annual.pbt[n-1] || 0) + (annual.interest[n-1] || 0);
    
    const roe = avgEquity > 0 ? (latestProfit / avgEquity) * 100 : null;
    const roce = capitalEmployed > 0 ? (ebit / capitalEmployed) * 100 : null;
    
    // Operating Profit = Sales - Operating Expenses
    const operatingProfit = latestSales - 
        (annual.rawMaterial[n-1] || 0) - 
        (annual.changeInventory[n-1] || 0) -
        (annual.powerFuel[n-1] || 0) -
        (annual.otherMfg[n-1] || 0) -
        (annual.employeeCost[n-1] || 0) -
        (annual.sellingAdmin[n-1] || 0) -
        (annual.otherExpenses[n-1] || 0);
    
    const opm = latestSales > 0 ? (operatingProfit / latestSales) * 100 : null;
    const npm = latestSales > 0 ? (latestProfit / latestSales) * 100 : null;
    
    document.getElementById('roeLatest').textContent = formatPercent(roe);
    document.getElementById('roceLatest').textContent = formatPercent(roce);
    document.getElementById('opmLatest').textContent = formatPercent(opm);
    document.getElementById('npmLatest').textContent = formatPercent(npm);
    
    // Apply color coding
    setMetricColor('roeLatest', roe, 18);
    setMetricColor('roceLatest', roce, 15);
    setMetricColor('salesCAGR5', salesCAGR5, 12);
    setMetricColor('profitCAGR5', profitCAGR5, 15);
    
    // Quality Score
    const qualityScore = calculateQualityScore();
    displayQualityScore(qualityScore);
    
    // Red Flags
    const redFlags = detectRedFlags();
    displayRedFlags(redFlags);
    
    // DuPont Analysis
    displayDuPontAnalysis();
    
    // Efficiency Metrics
    displayEfficiencyMetrics();
    
    // Leverage Metrics
    displayLeverageMetrics();
    
    // Moat Analysis
    displayMoatAnalysis();
    
    // Capital Allocation
    displayCapitalAllocation();
    
    // Value Migration
    displayValueMigration();
}

// Calculate CAGR
function calculateCAGR(data, years) {
    if (!data || data.length < years + 1) return null;
    
    const endValue = data[data.length - 1];
    const startValue = data[data.length - 1 - years];
    
    if (!endValue || !startValue || endValue <= 0 || startValue <= 0) return null;
    
    return (Math.pow(endValue / startValue, 1 / years) - 1) * 100;
}

// Quality Score Calculation (100 points)
function calculateQualityScore() {
    const { annual } = companyData;
    const n = annual.years.length;
    let score = 0;
    const breakdown = {};
    
    // 1. Profitability Quality (20 points)
    let profitabilityScore = 0;
    
    // a) Consistent profitability (10 pts)
    const profitableYears = annual.netProfit.filter(p => p && p > 0).length;
    if (profitableYears >= 10) profitabilityScore += 10;
    else if (profitableYears >= 7) profitabilityScore += 7;
    else if (profitableYears >= 5) profitabilityScore += 5;
    
    // b) High & stable margins (10 pts)
    const margins = annual.sales.map((s, i) => {
        if (!s || s === 0) return null;
        const profit = annual.netProfit[i];
        const rm = annual.rawMaterial[i] || 0;
        const other = (annual.powerFuel[i] || 0) + (annual.otherMfg[i] || 0) + 
                      (annual.employeeCost[i] || 0) + (annual.sellingAdmin[i] || 0);
        const opProfit = s - rm - other;
        return (opProfit / s) * 100;
    }).filter(m => m !== null);
    
    const avgMargin = margins.reduce((a, b) => a + b, 0) / margins.length;
    const stdDev = Math.sqrt(margins.reduce((sum, m) => sum + Math.pow(m - avgMargin, 2), 0) / margins.length);
    
    if (avgMargin > 15 && stdDev < 3) profitabilityScore += 10;
    else if (avgMargin > 10 && stdDev < 5) profitabilityScore += 6;
    else if (avgMargin > 5) profitabilityScore += 3;
    
    breakdown.profitability = profitabilityScore;
    score += profitabilityScore;
    
    // 2. Returns Quality (20 points)
    let returnsScore = 0;
    
    // Calculate ROCE for latest year
    const latestEquity = (annual.equity[n-1] || 0) + (annual.reserves[n-1] || 0);
    const prevEquity = (annual.equity[n-2] || 0) + (annual.reserves[n-2] || 0);
    const avgEquity = (latestEquity + prevEquity) / 2;
    const latestDebt = annual.borrowings[n-1] || 0;
    const capitalEmployed = avgEquity + latestDebt;
    const ebit = (annual.pbt[n-1] || 0) + (annual.interest[n-1] || 0);
    const latestROCE = capitalEmployed > 0 ? (ebit / capitalEmployed) * 100 : null;
    
    if (latestROCE && latestROCE > 25) returnsScore += 10;
    else if (latestROCE && latestROCE > 18) returnsScore += 7;
    else if (latestROCE && latestROCE > 12) returnsScore += 4;
    
    // ROCE trend
    if (n >= 5) {
        const oldEquity = (annual.equity[n-6] || 0) + (annual.reserves[n-6] || 0);
        const oldDebt = annual.borrowings[n-6] || 0;
        const oldCapital = oldEquity + oldDebt;
        const oldEbit = (annual.pbt[n-6] || 0) + (annual.interest[n-6] || 0);
        const oldROCE = oldCapital > 0 ? (oldEbit / oldCapital) * 100 : null;
        
        if (latestROCE && oldROCE && latestROCE > oldROCE + 2) returnsScore += 10;
        else if (latestROCE && oldROCE && Math.abs(latestROCE - oldROCE) <= 2) returnsScore += 6;
    }
    
    breakdown.returns = returnsScore;
    score += returnsScore;
    
    // 3. Cash Flow Quality (20 points)
    let cashFlowScore = 0;
    
    // CFO / Net Income
    const latestCFO = annual.cfo[n-1];
    const latestProfit = annual.netProfit[n-1];
    const cfoRatio = latestProfit > 0 ? (latestCFO / latestProfit) * 100 : null;
    
    if (cfoRatio && cfoRatio > 100) cashFlowScore += 10;
    else if (cfoRatio && cfoRatio > 80) cashFlowScore += 7;
    else if (cfoRatio && cfoRatio > 60) cashFlowScore += 4;
    
    // FCF / Sales
    const latestSales = annual.sales[n-1];
    const capex = Math.abs(annual.cfi[n-1] || 0); // Simplified
    const fcf = latestCFO - capex;
    const fcfMargin = latestSales > 0 ? (fcf / latestSales) * 100 : null;
    
    if (fcfMargin && fcfMargin > 10) cashFlowScore += 10;
    else if (fcfMargin && fcfMargin > 5) cashFlowScore += 6;
    else if (fcfMargin && fcfMargin > 0) cashFlowScore += 3;
    
    breakdown.cashFlow = cashFlowScore;
    score += cashFlowScore;
    
    // 4. Balance Sheet Quality (20 points)
    let balanceSheetScore = 0;
    
    // Debt to Equity
    const debtToEquity = latestEquity > 0 ? latestDebt / latestEquity : null;
    
    if (debtToEquity !== null && debtToEquity < 0.3) balanceSheetScore += 10;
    else if (debtToEquity !== null && debtToEquity < 0.7) balanceSheetScore += 6;
    else if (debtToEquity !== null && debtToEquity < 1.5) balanceSheetScore += 3;
    
    // Working Capital Efficiency (simplified CCC)
    const latestReceivables = annual.receivables[n-1] || 0;
    const latestInventory = annual.inventory[n-1] || 0;
    const cogs = (annual.rawMaterial[n-1] || 0) + (annual.otherMfg[n-1] || 0);
    
    const debtorDays = latestSales > 0 ? (latestReceivables / latestSales) * 365 : null;
    const inventoryDays = cogs > 0 ? (latestInventory / cogs) * 365 : null;
    const ccc = debtorDays + inventoryDays; // Simplified, missing creditor days
    
    if (ccc && ccc < 60) balanceSheetScore += 10;
    else if (ccc && ccc < 90) balanceSheetScore += 7;
    else if (ccc && ccc < 120) balanceSheetScore += 4;
    
    breakdown.balanceSheet = balanceSheetScore;
    score += balanceSheetScore;
    
    // 5. Growth Quality (20 points)
    let growthScore = 0;
    
    const salesCAGR5 = calculateCAGR(annual.sales, Math.min(5, n - 1));
    const profitCAGR5 = calculateCAGR(annual.netProfit, Math.min(5, n - 1));
    
    if (salesCAGR5 && salesCAGR5 > 20) growthScore += 10;
    else if (salesCAGR5 && salesCAGR5 > 12) growthScore += 7;
    else if (salesCAGR5 && salesCAGR5 > 7) growthScore += 4;
    
    // Profit leverage
    if (profitCAGR5 && salesCAGR5 && profitCAGR5 > salesCAGR5 + 5) growthScore += 10;
    else if (profitCAGR5 && salesCAGR5 && Math.abs(profitCAGR5 - salesCAGR5) <= 5) growthScore += 6;
    else if (profitCAGR5) growthScore += 2;
    
    breakdown.growth = growthScore;
    score += growthScore;
    
    return { total: score, breakdown };
}

// Display Quality Score
function displayQualityScore(qualityScore) {
    const { total, breakdown } = qualityScore;
    
    document.getElementById('qualityScore').textContent = total;
    
    let rating = 'Low Quality';
    let ratingClass = 'low';
    
    if (total >= 90) { rating = 'Exceptional'; ratingClass = 'excellent'; }
    else if (total >= 75) { rating = 'High Quality'; ratingClass = 'high'; }
    else if (total >= 60) { rating = 'Above Average'; ratingClass = 'above-avg'; }
    else if (total >= 40) { rating = 'Average'; ratingClass = 'average'; }
    
    const ratingEl = document.getElementById('qualityRating');
    ratingEl.textContent = rating;
    ratingEl.className = 'rating ' + ratingClass;
    
    // Breakdown
    const breakdownHTML = `
        <div class="dimension">
            <span class="dimension-label">Profitability</span>
            <span class="dimension-score">${breakdown.profitability}/20</span>
        </div>
        <div class="dimension">
            <span class="dimension-label">Returns</span>
            <span class="dimension-score">${breakdown.returns}/20</span>
        </div>
        <div class="dimension">
            <span class="dimension-label">Cash Flow</span>
            <span class="dimension-score">${breakdown.cashFlow}/20</span>
        </div>
        <div class="dimension">
            <span class="dimension-label">Balance Sheet</span>
            <span class="dimension-score">${breakdown.balanceSheet}/20</span>
        </div>
        <div class="dimension">
            <span class="dimension-label">Growth</span>
            <span class="dimension-score">${breakdown.growth}/20</span>
        </div>
    `;
    
    document.getElementById('qualityBreakdown').innerHTML = breakdownHTML;
}

// Detect Red Flags
function detectRedFlags() {
    const { annual } = companyData;
    const n = annual.years.length;
    const flags = [];
    
    if (n < 3) return flags;
    
    // 1. Receivables growing faster than sales
    const salesGrowth = ((annual.sales[n-1] - annual.sales[n-3]) / annual.sales[n-3]) * 100;
    const receivablesGrowth = ((annual.receivables[n-1] - annual.receivables[n-3]) / annual.receivables[n-3]) * 100;
    
    if (receivablesGrowth > salesGrowth + 10) {
        flags.push({
            title: 'Receivables Growing Faster Than Sales',
            description: `Receivables grew ${receivablesGrowth.toFixed(1)}% vs Sales ${salesGrowth.toFixed(1)}%`,
            severity: 'high'
        });
    }
    
    // 2. Inventory buildup
    const inventoryGrowth = ((annual.inventory[n-1] - annual.inventory[n-3]) / annual.inventory[n-3]) * 100;
    
    if (inventoryGrowth > salesGrowth + 15) {
        flags.push({
            title: 'Inventory Buildup',
            description: `Inventory grew ${inventoryGrowth.toFixed(1)}% vs Sales ${salesGrowth.toFixed(1)}%`,
            severity: 'medium'
        });
    }
    
    // 3. Other Income > 50% of Operating Profit
    const latestSales = annual.sales[n-1];
    const opProfit = latestSales - (annual.rawMaterial[n-1] || 0) - (annual.employeeCost[n-1] || 0) - 
                     (annual.sellingAdmin[n-1] || 0);
    const otherIncome = annual.otherIncome[n-1] || 0;
    
    if (otherIncome > opProfit * 0.5) {
        flags.push({
            title: 'High Other Income',
            description: `Other Income is ${((otherIncome/opProfit)*100).toFixed(0)}% of Operating Profit`,
            severity: 'medium'
        });
    }
    
    // 4. CWIP > 30% of Gross Block
    const cwip = annual.cwip[n-1] || 0;
    const netBlock = annual.netBlock[n-1] || 0;
    const grossBlock = netBlock * 1.5; // Approximation
    
    if (cwip > grossBlock * 0.3) {
        flags.push({
            title: 'High CWIP',
            description: `CWIP is ${((cwip/grossBlock)*100).toFixed(0)}% of Gross Block`,
            severity: 'medium'
        });
    }
    
    // 5. Debt surge with declining ROCE
    if (n >= 3) {
        const debtGrowth = ((annual.borrowings[n-1] - annual.borrowings[n-3]) / (annual.borrowings[n-3] || 1)) * 100;
        const latestEquity = (annual.equity[n-1] || 0) + (annual.reserves[n-1] || 0);
        const oldEquity = (annual.equity[n-3] || 0) + (annual.reserves[n-3] || 0);
        
        const latestROCE = ((annual.pbt[n-1] || 0) + (annual.interest[n-1] || 0)) / (latestEquity + (annual.borrowings[n-1] || 0)) * 100;
        const oldROCE = ((annual.pbt[n-3] || 0) + (annual.interest[n-3] || 0)) / (oldEquity + (annual.borrowings[n-3] || 0)) * 100;
        
        if (debtGrowth > 20 && latestROCE < oldROCE) {
            flags.push({
                title: 'Debt Surge with Declining ROCE',
                description: `Debt up ${debtGrowth.toFixed(0)}% while ROCE declined`,
                severity: 'high'
            });
        }
    }
    
    return flags;
}

// Display Red Flags
function displayRedFlags(flags) {
    const countEl = document.getElementById('redFlagsCount');
    countEl.querySelector('.count').textContent = flags.length;
    
    if (flags.length === 0) {
        countEl.className = 'flags-count clean';
        document.getElementById('redFlagsList').innerHTML = '<div style="color: var(--accent-success); text-align: center; padding: 2rem;">✓ Clean Balance Sheet - No Red Flags Detected</div>';
    } else if (flags.length <= 2) {
        countEl.className = 'flags-count caution';
    } else {
        countEl.className = 'flags-count danger';
    }
    
    const flagsHTML = flags.map(flag => `
        <div class="flag-item">
            <div class="flag-icon">⚠️</div>
            <div class="flag-content">
                <div class="flag-title">${flag.title}</div>
                <div class="flag-description">${flag.description}</div>
            </div>
        </div>
    `).join('');
    
    if (flags.length > 0) {
        document.getElementById('redFlagsList').innerHTML = flagsHTML;
    }
}

// DuPont Analysis
function displayDuPontAnalysis() {
    const { annual } = companyData;
    const n = annual.years.length;
    
    const latestSales = annual.sales[n-1];
    const latestProfit = annual.netProfit[n-1];
    const latestAssets = annual.totalAssets[n-1];
    const latestEquity = (annual.equity[n-1] || 0) + (annual.reserves[n-1] || 0);
    const prevEquity = (annual.equity[n-2] || 0) + (annual.reserves[n-2] || 0);
    const avgEquity = (latestEquity + prevEquity) / 2;
    
    const netMargin = latestSales > 0 ? (latestProfit / latestSales) * 100 : null;
    const assetTurnover = latestAssets > 0 ? latestSales / latestAssets : null;
    const equityMultiplier = avgEquity > 0 ? latestAssets / avgEquity : null;
    const roe = avgEquity > 0 ? (latestProfit / avgEquity) * 100 : null;
    
    const html = `
        <div class="dupont-breakdown">
            <div class="dupont-factor">
                <span class="factor-name">Net Margin</span>
                <span class="factor-value">${formatPercent(netMargin)}</span>
            </div>
            <div class="dupont-factor">
                <span class="factor-name">× Asset Turnover</span>
                <span class="factor-value">${assetTurnover ? assetTurnover.toFixed(2) + 'x' : 'N/A'}</span>
            </div>
            <div class="dupont-factor">
                <span class="factor-name">× Equity Multiplier</span>
                <span class="factor-value">${equityMultiplier ? equityMultiplier.toFixed(2) + 'x' : 'N/A'}</span>
            </div>
            <div class="dupont-factor" style="border-top: 2px solid var(--border-color); margin-top: 1rem; padding-top: 1rem;">
                <span class="factor-name" style="font-weight: 700;">= ROE</span>
                <span class="factor-value" style="color: var(--accent-primary); font-size: 1.5rem;">${formatPercent(roe)}</span>
            </div>
        </div>
        <div style="margin-top: 1.5rem; padding: 1rem; background: var(--bg-secondary); border-radius: 6px; font-size: 0.9rem; color: var(--text-secondary);">
            <strong>Interpretation:</strong> 
            ${netMargin > 15 ? 'Strong profitability indicates pricing power and operational efficiency.' : 
              netMargin > 10 ? 'Good profitability, room for improvement.' : 
              'Low margins suggest competitive pressures or inefficiencies.'}
            ${assetTurnover > 2 ? ' Excellent asset utilization.' : 
              assetTurnover > 1 ? ' Decent asset efficiency.' : 
              ' Capital-intensive business model.'}
        </div>
    `;
    
    document.getElementById('dupontAnalysis').innerHTML = html;
}

// Efficiency Metrics
function displayEfficiencyMetrics() {
    const { annual } = companyData;
    const n = annual.years.length;
    
    const latestSales = annual.sales[n-1];
    const latestAssets = annual.totalAssets[n-1];
    const latestReceivables = annual.receivables[n-1] || 0;
    const latestInventory = annual.inventory[n-1] || 0;
    const cogs = (annual.rawMaterial[n-1] || 0) + (annual.otherMfg[n-1] || 0);
    
    const assetTurnover = latestAssets > 0 ? latestSales / latestAssets : null;
    const debtorDays = latestSales > 0 ? (latestReceivables / latestSales) * 365 : null;
    const inventoryDays = cogs > 0 ? (latestInventory / cogs) * 365 : null;
    const ccc = debtorDays + inventoryDays;
    
    const html = `
        <div style="display: flex; flex-direction: column; gap: 1rem;">
            <div class="dupont-factor">
                <span class="factor-name">Asset Turnover</span>
                <span class="factor-value">${assetTurnover ? assetTurnover.toFixed(2) + 'x' : 'N/A'}</span>
            </div>
            <div class="dupont-factor">
                <span class="factor-name">Debtor Days</span>
                <span class="factor-value">${debtorDays ? Math.round(debtorDays) + ' days' : 'N/A'}</span>
            </div>
            <div class="dupont-factor">
                <span class="factor-name">Inventory Days</span>
                <span class="factor-value">${inventoryDays ? Math.round(inventoryDays) + ' days' : 'N/A'}</span>
            </div>
            <div class="dupont-factor">
                <span class="factor-name">Cash Conversion Cycle</span>
                <span class="factor-value">${ccc ? Math.round(ccc) + ' days' : 'N/A'}</span>
            </div>
        </div>
    `;
    
    document.getElementById('efficiencyMetrics').innerHTML = html;
}

// Leverage Metrics
function displayLeverageMetrics() {
    const { annual } = companyData;
    const n = annual.years.length;
    
    const latestEquity = (annual.equity[n-1] || 0) + (annual.reserves[n-1] || 0);
    const latestDebt = annual.borrowings[n-1] || 0;
    const latestInterest = annual.interest[n-1] || 0;
    const ebit = (annual.pbt[n-1] || 0) + latestInterest;
    
    const debtToEquity = latestEquity > 0 ? latestDebt / latestEquity : null;
    const interestCoverage = latestInterest > 0 ? ebit / latestInterest : null;
    
    const html = `
        <div style="display: flex; flex-direction: column; gap: 1rem;">
            <div class="dupont-factor">
                <span class="factor-name">Debt to Equity</span>
                <span class="factor-value">${debtToEquity ? debtToEquity.toFixed(2) + 'x' : 'N/A'}</span>
            </div>
            <div class="dupont-factor">
                <span class="factor-name">Interest Coverage</span>
                <span class="factor-value">${interestCoverage ? interestCoverage.toFixed(2) + 'x' : 'N/A'}</span>
            </div>
            <div class="dupont-factor">
                <span class="factor-name">Total Debt</span>
                <span class="factor-value">${formatNumber(latestDebt)} Cr</span>
            </div>
            <div style="margin-top: 0.5rem; padding: 0.75rem; background: var(--bg-secondary); border-radius: 6px; font-size: 0.85rem; color: var(--text-secondary);">
                ${debtToEquity < 0.5 ? '✓ Conservative leverage' : 
                  debtToEquity < 1.0 ? '⚠ Moderate leverage' : 
                  '⚠️ High leverage - monitor carefully'}
            </div>
        </div>
    `;
    
    document.getElementById('leverageMetrics').innerHTML = html;
}

// Moat Analysis
function displayMoatAnalysis() {
    const { annual } = companyData;
    const n = annual.years.length;
    
    let moatScore = 0;
    const indicators = [];
    
    // 1. Consistently High ROIC
    let highROICYears = 0;
    for (let i = Math.max(0, n - 7); i < n; i++) {
        const equity = (annual.equity[i] || 0) + (annual.reserves[i] || 0);
        const debt = annual.borrowings[i] || 0;
        const capital = equity + debt;
        const nopat = (annual.netProfit[i] || 0) * 1.15; // Approximation
        const roic = capital > 0 ? (nopat / capital) * 100 : null;
        
        if (roic && roic > 15) highROICYears++;
    }
    
    const roicPass = highROICYears >= 5;
    if (roicPass) moatScore += 2;
    indicators.push({ name: 'High ROIC (7+ years)', pass: roicPass, value: `${highROICYears}/7 years` });
    
    // 2. Pricing Power (Margin stability)
    const margins = annual.sales.slice(-5).map((s, i) => {
        const idx = n - 5 + i;
        if (!s || s === 0) return null;
        const profit = annual.netProfit[idx];
        return (profit / s) * 100;
    }).filter(m => m !== null);
    
    const marginStdDev = Math.sqrt(margins.reduce((sum, m) => {
        const avg = margins.reduce((a, b) => a + b, 0) / margins.length;
        return sum + Math.pow(m - avg, 2);
    }, 0) / margins.length);
    
    const pricingPowerPass = marginStdDev < 3;
    if (pricingPowerPass) moatScore += 2;
    indicators.push({ name: 'Margin Stability', pass: pricingPowerPass, value: `σ = ${marginStdDev.toFixed(1)}%` });
    
    // 3. Capital Efficiency
    const latestSales = annual.sales[n-1];
    const latestAssets = annual.totalAssets[n-1];
    const assetTurnover = latestAssets > 0 ? latestSales / latestAssets : null;
    
    const efficiencyPass = assetTurnover > 1.5;
    if (efficiencyPass) moatScore += 2;
    indicators.push({ name: 'Asset Efficiency', pass: efficiencyPass, value: assetTurnover ? assetTurnover.toFixed(2) + 'x' : 'N/A' });
    
    const moatRating = moatScore >= 5 ? 'Wide Moat' : moatScore >= 3 ? 'Narrow Moat' : 'No Moat';
    
    const html = `
        <div style="margin-bottom: 1.5rem; padding: 1rem; background: var(--bg-secondary); border-radius: 6px;">
            <div style="font-size: 1.25rem; font-weight: 700; margin-bottom: 0.5rem;">
                ${moatRating}
            </div>
            <div style="font-size: 0.9rem; color: var(--text-secondary);">
                Score: ${moatScore}/6
            </div>
        </div>
        <div class="moat-indicators">
            ${indicators.map(ind => `
                <div class="indicator">
                    <span>${ind.name}</span>
                    <span class="indicator-status ${ind.pass ? 'pass' : 'fail'}">
                        ${ind.pass ? '✓' : '✗'} ${ind.value}
                    </span>
                </div>
            `).join('')}
        </div>
    `;
    
    document.getElementById('moatAnalysis').innerHTML = html;
}

// Capital Allocation
function displayCapitalAllocation() {
    const { annual } = companyData;
    const n = annual.years.length;
    
    if (n < 3) {
        document.getElementById('capitalAllocation').innerHTML = '<p style="color: var(--text-muted);">Insufficient data (requires 3+ years)</p>';
        return;
    }
    
    // Dividend consistency
    const dividendYears = annual.dividend.filter(d => d && d > 0).length;
    const dividendScore = dividendYears / n >= 0.7 ? 3 : dividendYears / n >= 0.5 ? 2 : 1;
    
    // Average payout ratio
    const payoutRatios = annual.dividend.map((d, i) => {
        const profit = annual.netProfit[i];
        if (!profit || profit <= 0) return null;
        return (d / profit) * 100;
    }).filter(p => p !== null);
    
    const avgPayout = payoutRatios.reduce((a, b) => a + b, 0) / payoutRatios.length;
    
    // Debt management
    const latestDebt = annual.borrowings[n-1] || 0;
    const oldDebt = annual.borrowings[Math.max(0, n-3)] || 0;
    const debtGrowth = oldDebt > 0 ? ((latestDebt - oldDebt) / oldDebt) * 100 : null;
    
    const debtScore = latestDebt === 0 ? 3 : debtGrowth < 10 ? 2 : 1;
    
    const totalScore = dividendScore + debtScore;
    const grade = totalScore >= 5 ? 'A' : totalScore >= 4 ? 'B' : totalScore >= 3 ? 'C' : 'D';
    
    const html = `
        <div style="margin-bottom: 1.5rem; padding: 1rem; background: var(--bg-secondary); border-radius: 6px;">
            <div style="font-size: 2rem; font-weight: 700; font-family: var(--font-mono);">
                Grade ${grade}
            </div>
            <div style="font-size: 0.9rem; color: var(--text-secondary); margin-top: 0.25rem;">
                Capital Allocation Quality
            </div>
        </div>
        <div style="display: flex; flex-direction: column; gap: 1rem;">
            <div class="dupont-factor">
                <span class="factor-name">Dividend Consistency</span>
                <span class="factor-value">${dividendYears}/${n} years</span>
            </div>
            <div class="dupont-factor">
                <span class="factor-name">Avg Payout Ratio</span>
                <span class="factor-value">${avgPayout.toFixed(1)}%</span>
            </div>
            <div class="dupont-factor">
                <span class="factor-name">Debt Management</span>
                <span class="factor-value">${latestDebt === 0 ? 'Debt Free' : debtGrowth ? debtGrowth.toFixed(1) + '% growth' : 'N/A'}</span>
            </div>
        </div>
    `;
    
    document.getElementById('capitalAllocation').innerHTML = html;
}

// Value Migration
function displayValueMigration() {
    const { annual } = companyData;
    const n = annual.years.length;
    
    if (n < 5) {
        document.getElementById('valueMigration').innerHTML = '<p style="color: var(--text-muted);">Insufficient data (requires 5+ years)</p>';
        return;
    }
    
    const salesCAGR = calculateCAGR(annual.sales, 5);
    const profitCAGR = calculateCAGR(annual.netProfit, 5);
    
    // Margin trend
    const oldMargin = annual.sales[n-6] > 0 ? (annual.netProfit[n-6] / annual.sales[n-6]) * 100 : null;
    const newMargin = annual.sales[n-1] > 0 ? (annual.netProfit[n-1] / annual.sales[n-1]) * 100 : null;
    const marginChange = newMargin - oldMargin;
    
    // ROCE trend
    const oldEquity = (annual.equity[n-6] || 0) + (annual.reserves[n-6] || 0);
    const newEquity = (annual.equity[n-1] || 0) + (annual.reserves[n-1] || 0);
    const oldROCE = oldEquity > 0 ? ((annual.pbt[n-6] || 0) / oldEquity) * 100 : null;
    const newROCE = newEquity > 0 ? ((annual.pbt[n-1] || 0) / newEquity) * 100 : null;
    const roceChange = newROCE - oldROCE;
    
    let direction = 'Stable';
    let strength = 0;
    
    if (salesCAGR > 15 && marginChange > 2 && roceChange > 0) {
        direction = 'Strong Inward';
        strength = 4;
    } else if (salesCAGR > 10 && marginChange > 0) {
        direction = 'Inward';
        strength = 3;
    } else if (salesCAGR > 5) {
        direction = 'Stable';
        strength = 2;
    } else {
        direction = 'Outward';
        strength = 1;
    }
    
    const html = `
        <div style="margin-bottom: 1.5rem; padding: 1rem; background: var(--bg-secondary); border-radius: 6px;">
            <div style="font-size: 1.25rem; font-weight: 700; margin-bottom: 0.5rem;">
                ${direction === 'Strong Inward' ? '🟢' : direction === 'Inward' ? '🟢' : direction === 'Stable' ? '🟡' : '🔴'} ${direction}
            </div>
            <div style="font-size: 0.85rem; color: var(--text-secondary);">
                Value Migration Direction
            </div>
        </div>
        <div style="display: flex; flex-direction: column; gap: 1rem;">
            <div class="dupont-factor">
                <span class="factor-name">Sales CAGR (5Y)</span>
                <span class="factor-value">${formatPercent(salesCAGR)}</span>
            </div>
            <div class="dupont-factor">
                <span class="factor-name">Margin Change</span>
                <span class="factor-value">${marginChange ? (marginChange > 0 ? '+' : '') + marginChange.toFixed(1) + '%' : 'N/A'}</span>
            </div>
            <div class="dupont-factor">
                <span class="factor-name">ROCE Change</span>
                <span class="factor-value">${roceChange ? (roceChange > 0 ? '+' : '') + roceChange.toFixed(1) + '%' : 'N/A'}</span>
            </div>
        </div>
    `;
    
    document.getElementById('valueMigration').innerHTML = html;
}

// Display Financial Statements
function displayFinancialStatements() {
    displayPLStatement();
    displayBalanceSheet();
    displayCashFlow();
}

function displayPLStatement() {
    const { annual } = companyData;
    const years = annual.years;
    
    const table = document.getElementById('plTable');
    
    // Header
    const thead = table.querySelector('thead');
    thead.innerHTML = `
        <tr>
            <th>Particulars</th>
            ${years.map(y => `<th>FY${y}</th>`).join('')}
        </tr>
    `;
    
    // Body
    const tbody = table.querySelector('tbody');
    const rows = [
        { label: 'Sales', data: annual.sales },
        { label: 'Raw Material Cost', data: annual.rawMaterial },
        { label: 'Employee Cost', data: annual.employeeCost },
        { label: 'Selling & Admin', data: annual.sellingAdmin },
        { label: 'Depreciation', data: annual.depreciation },
        { label: 'Other Income', data: annual.otherIncome },
        { label: 'Interest', data: annual.interest },
        { label: 'Profit Before Tax', data: annual.pbt, class: 'total-row' },
        { label: 'Tax', data: annual.tax },
        { label: 'Net Profit', data: annual.netProfit, class: 'total-row' },
        { label: 'Dividend', data: annual.dividend }
    ];
    
    tbody.innerHTML = rows.map(row => `
        <tr${row.class ? ` class="${row.class}"` : ''}>
            <td>${row.label}</td>
            ${row.data.map(v => `<td>${formatNumber(v)}</td>`).join('')}
        </tr>
    `).join('');
}

function displayBalanceSheet() {
    const { annual } = companyData;
    const years = annual.years;
    
    const table = document.getElementById('bsTable');
    
    // Header
    const thead = table.querySelector('thead');
    thead.innerHTML = `
        <tr>
            <th>Particulars</th>
            ${years.map(y => `<th>FY${y}</th>`).join('')}
        </tr>
    `;
    
    // Body
    const tbody = table.querySelector('tbody');
    const rows = [
        { label: 'LIABILITIES', data: [], class: 'category-row' },
        { label: 'Equity Capital', data: annual.equity },
        { label: 'Reserves', data: annual.reserves },
        { label: 'Borrowings', data: annual.borrowings },
        { label: 'Other Liabilities', data: annual.otherLiabilities },
        { label: 'Total Liabilities', data: annual.totalLiabilities, class: 'total-row' },
        { label: '', data: [], class: 'category-row' },
        { label: 'ASSETS', data: [], class: 'category-row' },
        { label: 'Fixed Assets', data: annual.netBlock },
        { label: 'CWIP', data: annual.cwip },
        { label: 'Investments', data: annual.investments },
        { label: 'Receivables', data: annual.receivables },
        { label: 'Inventory', data: annual.inventory },
        { label: 'Cash & Bank', data: annual.cash },
        { label: 'Other Assets', data: annual.otherAssets },
        { label: 'Total Assets', data: annual.totalAssets, class: 'total-row' }
    ];
    
    tbody.innerHTML = rows.map(row => {
        if (row.data.length === 0) {
            return `<tr class="${row.class || ''}"><td colspan="${years.length + 1}">${row.label}</td></tr>`;
        }
        return `
            <tr${row.class ? ` class="${row.class}"` : ''}>
                <td>${row.label}</td>
                ${row.data.map(v => `<td>${formatNumber(v)}</td>`).join('')}
            </tr>
        `;
    }).join('');
}

function displayCashFlow() {
    const { annual } = companyData;
    const years = annual.years;
    
    const table = document.getElementById('cfTable');
    
    // Header
    const thead = table.querySelector('thead');
    thead.innerHTML = `
        <tr>
            <th>Particulars</th>
            ${years.map(y => `<th>FY${y}</th>`).join('')}
        </tr>
    `;
    
    // Body
    const tbody = table.querySelector('tbody');
    const rows = [
        { label: 'Operating Activities', data: annual.cfo },
        { label: 'Investing Activities', data: annual.cfi },
        { label: 'Financing Activities', data: annual.cff },
        { label: 'Net Cash Flow', data: annual.netCashFlow, class: 'total-row' }
    ];
    
    tbody.innerHTML = rows.map(row => `
        <tr${row.class ? ` class="${row.class}"` : ''}>
            <td>${row.label}</td>
            ${row.data.map(v => `<td>${formatNumber(v)}</td>`).join('')}
        </tr>
    `).join('');
}

// Display Charts
function displayCharts() {
    createRevenueChart();
    createMarginChart();
    createReturnsChart();
    createCashflowChart();
}

function createRevenueChart() {
    const { annual } = companyData;
    const ctx = document.getElementById('revenueChart');
    
    new Chart(ctx, {
        type: 'line',
        data: {
            labels: annual.years,
            datasets: [
                {
                    label: 'Revenue',
                    data: annual.sales,
                    borderColor: '#4fc3f7',
                    backgroundColor: 'rgba(79, 195, 247, 0.1)',
                    yAxisID: 'y',
                    tension: 0.3
                },
                {
                    label: 'Net Profit',
                    data: annual.netProfit,
                    borderColor: '#26c281',
                    backgroundColor: 'rgba(38, 194, 129, 0.1)',
                    yAxisID: 'y',
                    tension: 0.3
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    labels: { color: '#e8eaf6' }
                }
            },
            scales: {
                y: {
                    type: 'linear',
                    position: 'left',
                    ticks: { color: '#9fa8c9' },
                    grid: { color: '#2a3764' }
                },
                x: {
                    ticks: { color: '#9fa8c9' },
                    grid: { color: '#2a3764' }
                }
            }
        }
    });
}

function createMarginChart() {
    const { annual } = companyData;
    const ctx = document.getElementById('marginChart');
    
    const margins = annual.sales.map((s, i) => {
        if (!s || s === 0) return null;
        const profit = annual.netProfit[i];
        const rm = annual.rawMaterial[i] || 0;
        const emp = annual.employeeCost[i] || 0;
        const sa = annual.sellingAdmin[i] || 0;
        const opProfit = s - rm - emp - sa;
        return (opProfit / s) * 100;
    });
    
    const netMargins = annual.sales.map((s, i) => {
        if (!s || s === 0) return null;
        return (annual.netProfit[i] / s) * 100;
    });
    
    new Chart(ctx, {
        type: 'line',
        data: {
            labels: annual.years,
            datasets: [
                {
                    label: 'Operating Margin %',
                    data: margins,
                    borderColor: '#f4a742',
                    backgroundColor: 'rgba(244, 167, 66, 0.1)',
                    tension: 0.3
                },
                {
                    label: 'Net Margin %',
                    data: netMargins,
                    borderColor: '#26c281',
                    backgroundColor: 'rgba(38, 194, 129, 0.1)',
                    tension: 0.3
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    labels: { color: '#e8eaf6' }
                }
            },
            scales: {
                y: {
                    ticks: { 
                        color: '#9fa8c9',
                        callback: value => value + '%'
                    },
                    grid: { color: '#2a3764' }
                },
                x: {
                    ticks: { color: '#9fa8c9' },
                    grid: { color: '#2a3764' }
                }
            }
        }
    });
}

function createReturnsChart() {
    const { annual } = companyData;
    const ctx = document.getElementById('returnsChart');
    const n = annual.years.length;
    
    const roes = [];
    const roces = [];
    
    for (let i = 1; i < n; i++) {
        const equity = (annual.equity[i] || 0) + (annual.reserves[i] || 0);
        const prevEquity = (annual.equity[i-1] || 0) + (annual.reserves[i-1] || 0);
        const avgEquity = (equity + prevEquity) / 2;
        
        const debt = annual.borrowings[i] || 0;
        const capitalEmployed = avgEquity + debt;
        
        const profit = annual.netProfit[i] || 0;
        const ebit = (annual.pbt[i] || 0) + (annual.interest[i] || 0);
        
        roes.push(avgEquity > 0 ? (profit / avgEquity) * 100 : null);
        roces.push(capitalEmployed > 0 ? (ebit / capitalEmployed) * 100 : null);
    }
    
    new Chart(ctx, {
        type: 'bar',
        data: {
            labels: annual.years.slice(1),
            datasets: [
                {
                    label: 'ROE %',
                    data: roes,
                    backgroundColor: 'rgba(79, 195, 247, 0.7)',
                    borderColor: '#4fc3f7',
                    borderWidth: 1
                },
                {
                    label: 'ROCE %',
                    data: roces,
                    backgroundColor: 'rgba(38, 194, 129, 0.7)',
                    borderColor: '#26c281',
                    borderWidth: 1
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    labels: { color: '#e8eaf6' }
                }
            },
            scales: {
                y: {
                    ticks: { 
                        color: '#9fa8c9',
                        callback: value => value + '%'
                    },
                    grid: { color: '#2a3764' }
                },
                x: {
                    ticks: { color: '#9fa8c9' },
                    grid: { color: '#2a3764' }
                }
            }
        }
    });
}

function createCashflowChart() {
    const { annual } = companyData;
    const ctx = document.getElementById('cashflowChart');
    
    new Chart(ctx, {
        type: 'bar',
        data: {
            labels: annual.years,
            datasets: [
                {
                    label: 'Operating',
                    data: annual.cfo,
                    backgroundColor: 'rgba(38, 194, 129, 0.7)',
                    borderColor: '#26c281',
                    borderWidth: 1
                },
                {
                    label: 'Investing',
                    data: annual.cfi,
                    backgroundColor: 'rgba(239, 83, 80, 0.7)',
                    borderColor: '#ef5350',
                    borderWidth: 1
                },
                {
                    label: 'Financing',
                    data: annual.cff,
                    backgroundColor: 'rgba(244, 167, 66, 0.7)',
                    borderColor: '#f4a742',
                    borderWidth: 1
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    labels: { color: '#e8eaf6' }
                }
            },
            scales: {
                y: {
                    ticks: { color: '#9fa8c9' },
                    grid: { color: '#2a3764' }
                },
                x: {
                    ticks: { color: '#9fa8c9' },
                    grid: { color: '#2a3764' }
                }
            }
        }
    });
}

// Tab Switching
function switchTab(tabName) {
    // Update tab buttons
    document.querySelectorAll('.tab').forEach(tab => {
        tab.classList.remove('active');
        if (tab.dataset.tab === tabName) {
            tab.classList.add('active');
        }
    });
    
    // Update tab panels
    document.querySelectorAll('.tab-panel').forEach(panel => {
        panel.classList.remove('active');
        if (panel.id === tabName) {
            panel.classList.add('active');
        }
    });
}

// Utility Functions
function formatNumber(num) {
    if (num === null || num === undefined || isNaN(num)) return '-';
    return num.toFixed(2);
}

function formatLargeNumber(num) {
    if (num === null || num === undefined || isNaN(num)) return '-';
    if (num >= 1000) {
        return (num / 1000).toFixed(2) + 'K';
    }
    return num.toFixed(2);
}

function formatCurrency(num) {
    if (num === null || num === undefined || isNaN(num)) return '-';
    return '₹' + num.toFixed(2);
}

function formatPercent(num) {
    if (num === null || num === undefined || isNaN(num)) return '-';
    return num.toFixed(1) + '%';
}

function setMetricColor(elementId, value, threshold) {
    const el = document.getElementById(elementId);
    if (value === null || value === undefined) return;
    
    if (value >= threshold) {
        el.classList.add('positive');
    } else if (value < 0) {
        el.classList.add('negative');
    }
}
