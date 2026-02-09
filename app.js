// Tankrich Dashboard - Main Application Logic
// REVISED VERSION with Advanced Framework Improvements

// Global state
let companyData = null;
let workbookGlobal = null;
let comparisonCompanies = [];

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
    
    // Comparison mode
    document.getElementById('compareFileInput').addEventListener('change', handleComparisonUpload);
    document.getElementById('clearComparison').addEventListener('click', clearComparison);
    
    // Export/Print
    document.getElementById('printBtn').addEventListener('click', () => window.print());
    document.getElementById('exportBtn').addEventListener('click', () => {
        alert('PDF export: Please use Print and "Save as PDF"');
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
        workbookGlobal = workbook;
        
        companyData = parseExcelData(workbook);
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
    if (!dataSheet) throw new Error('Data Sheet not found');
    
    const raw = XLSX.utils.sheet_to_json(dataSheet, { header: 1, defval: null });
    
    // Extract company meta
    const meta = {
        name: raw[0][1] || 'Unknown Company',
        faceValue: raw[6][1] || null,
        currentPrice: raw[7][1] || null,
        marketCap: raw[8][1] || null
    };
    
    // Extract annual data with proper date parsing
    const reportDates = (raw[15] || []).slice(4).filter(d => d);
    
    const years = reportDates.map(d => {
        if (typeof d === 'number') {
            const excelEpoch = new Date(1900, 0, 1);
            const date = new Date(excelEpoch.getTime() + (d - 2) * 86400000);
            return date.getFullYear();
        } else if (d instanceof Date) {
            return d.getFullYear();
        } else if (typeof d === 'string') {
            const parsed = new Date(d);
            return isNaN(parsed.getTime()) ? 1970 : parsed.getFullYear();
        }
        return 1970;
    });
    
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
        
        cfo: extractRow(raw, 81),
        cfi: extractRow(raw, 82),
        cff: extractRow(raw, 83),
        netCashFlow: extractRow(raw, 84),
        
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
    
    document.getElementById('companyName').textContent = companyData.meta.name;
    document.getElementById('currentPrice').textContent = formatCurrency(companyData.meta.currentPrice);
    document.getElementById('marketCap').textContent = formatLargeNumber(companyData.meta.marketCap) + ' Cr';
    document.getElementById('latestYear').textContent = companyData.annual.years[companyData.annual.years.length - 1] || '-';
    
    calculateAndDisplayMetrics();
    displayFinancialStatements();
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
    
    // Profitability Ratios
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
    
    setMetricColor('roeLatest', roe, 18);
    setMetricColor('roceLatest', roce, 15);
    setMetricColor('salesCAGR5', salesCAGR5, 12);
    setMetricColor('profitCAGR5', profitCAGR5, 15);
    
    const qualityScore = calculateQualityScore();
    displayQualityScore(qualityScore);
    
    const redFlags = detectRedFlags();
    displayRedFlags(redFlags);
    
    // Updated frameworks
    displayDuPontAnalysis();
    displayEfficiencyMetrics();
    displayLeverageMetrics();
    displayMoatAnalysis();
    displayCapitalAllocation();
    displayValueMigration();
    displayEarningPowerBox();
    displayCAPAnalysis();
    displayCapexSplit();
    displayIncrementalROIC();
    displayRMSensitivity();
    displayBuffettTest();
    displayFLOATDetection();
    
    try {
        displayQuarterlyAnalysis();
    } catch (error) {
        console.error('Error displaying quarterly analysis:', error);
    }
}

// Calculate CAGR
function calculateCAGR(data, years) {
    if (!data || data.length < years + 1) return null;
    
    const endValue = data[data.length - 1];
    const startValue = data[data.length - 1 - years];
    
    if (!endValue || !startValue || endValue <= 0 || startValue <= 0) return null;
    
    return (Math.pow(endValue / startValue, 1 / years) - 1) * 100;
}

// Quality Score (UNCHANGED - as requested)
function calculateQualityScore(company = null) {
    const data = company || companyData;
    const { annual } = data;
    const n = annual.years.length;
    let score = 0;
    const breakdown = {};
    
    // Profitability Quality (20 points)
    let profitabilityScore = 0;
    const profitableYears = annual.netProfit.filter(p => p && p > 0).length;
    if (profitableYears >= 10) profitabilityScore += 10;
    else if (profitableYears >= 7) profitabilityScore += 7;
    else if (profitableYears >= 5) profitabilityScore += 5;
    
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
    
    // Returns Quality (20 points)
    let returnsScore = 0;
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
    
    // Cash Flow Quality (20 points)
    let cashFlowScore = 0;
    const latestCFO = annual.cfo[n-1];
    const latestProfit = annual.netProfit[n-1];
    const cfoRatio = latestProfit > 0 ? (latestCFO / latestProfit) * 100 : null;
    
    if (cfoRatio && cfoRatio > 100) cashFlowScore += 10;
    else if (cfoRatio && cfoRatio > 80) cashFlowScore += 7;
    else if (cfoRatio && cfoRatio > 60) cashFlowScore += 4;
    
    const latestSales = annual.sales[n-1];
    const capex = Math.abs(annual.cfi[n-1] || 0);
    const fcf = latestCFO - capex;
    const fcfMargin = latestSales > 0 ? (fcf / latestSales) * 100 : null;
    
    if (fcfMargin && fcfMargin > 10) cashFlowScore += 10;
    else if (fcfMargin && fcfMargin > 5) cashFlowScore += 6;
    else if (fcfMargin && fcfMargin > 0) cashFlowScore += 3;
    
    breakdown.cashFlow = cashFlowScore;
    score += cashFlowScore;
    
    // Balance Sheet Quality (20 points)
    let balanceSheetScore = 0;
    const debtToEquity = latestEquity > 0 ? latestDebt / latestEquity : null;
    
    if (debtToEquity !== null && debtToEquity < 0.3) balanceSheetScore += 10;
    else if (debtToEquity !== null && debtToEquity < 0.7) balanceSheetScore += 6;
    else if (debtToEquity !== null && debtToEquity < 1.5) balanceSheetScore += 3;
    
    const latestReceivables = annual.receivables[n-1] || 0;
    const latestInventory = annual.inventory[n-1] || 0;
    const cogs = (annual.rawMaterial[n-1] || 0) + (annual.otherMfg[n-1] || 0);
    
    const debtorDays = latestSales > 0 ? (latestReceivables / latestSales) * 365 : null;
    const inventoryDays = cogs > 0 ? (latestInventory / cogs) * 365 : null;
    const ccc = debtorDays + inventoryDays;
    
    if (ccc && ccc < 60) balanceSheetScore += 10;
    else if (ccc && ccc < 90) balanceSheetScore += 7;
    else if (ccc && ccc < 120) balanceSheetScore += 4;
    
    breakdown.balanceSheet = balanceSheetScore;
    score += balanceSheetScore;
    
    // Growth Quality (20 points)
    let growthScore = 0;
    const salesCAGR5 = calculateCAGR(annual.sales, Math.min(5, n - 1));
    const profitCAGR5 = calculateCAGR(annual.netProfit, Math.min(5, n - 1));
    
    if (salesCAGR5 && salesCAGR5 > 20) growthScore += 10;
    else if (salesCAGR5 && salesCAGR5 > 12) growthScore += 7;
    else if (salesCAGR5 && salesCAGR5 > 7) growthScore += 4;
    
    if (profitCAGR5 && salesCAGR5 && profitCAGR5 > salesCAGR5 + 5) growthScore += 10;
    else if (profitCAGR5 && salesCAGR5 && Math.abs(profitCAGR5 - salesCAGR5) <= 5) growthScore += 6;
    else if (profitCAGR5) growthScore += 2;
    
    breakdown.growth = growthScore;
    score += growthScore;
    
    return { total: score, breakdown };
}

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

// Red Flags Detection (UNCHANGED)
function detectRedFlags() {
    const { annual } = companyData;
    const n = annual.years.length;
    const flags = [];
    
    if (n < 3) return flags;
    
    const salesGrowth = ((annual.sales[n-1] - annual.sales[n-3]) / annual.sales[n-3]) * 100;
    const receivablesGrowth = ((annual.receivables[n-1] - annual.receivables[n-3]) / annual.receivables[n-3]) * 100;
    
    if (receivablesGrowth > salesGrowth + 10) {
        flags.push({
            title: 'Receivables Growing Faster Than Sales',
            description: `Receivables grew ${receivablesGrowth.toFixed(1)}% vs Sales ${salesGrowth.toFixed(1)}%`,
            severity: 'high'
        });
    }
    
    const inventoryGrowth = ((annual.inventory[n-1] - annual.inventory[n-3]) / annual.inventory[n-3]) * 100;
    
    if (inventoryGrowth > salesGrowth + 15) {
        flags.push({
            title: 'Inventory Buildup',
            description: `Inventory grew ${inventoryGrowth.toFixed(1)}% vs Sales ${salesGrowth.toFixed(1)}%`,
            severity: 'medium'
        });
    }
    
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
    
    const cwip = annual.cwip[n-1] || 0;
    const netBlock = annual.netBlock[n-1] || 0;
    const grossBlock = netBlock * 1.5;
    
    if (cwip > grossBlock * 0.3) {
        flags.push({
            title: 'High CWIP',
            description: `CWIP is ${((cwip/grossBlock)*100).toFixed(0)}% of Gross Block`,
            severity: 'medium'
        });
    }
    
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

// ============================================================================
// REVISED: DUPONT ANALYSIS - MAUBOUSSIN'S ROIC APPROACH
// ============================================================================

function displayDuPontAnalysis() {
    const { annual } = companyData;
    const n = annual.years.length;
    
    if (n < 2) {
        document.getElementById('dupontAnalysis').innerHTML = '<p style="color: var(--text-muted);">Insufficient data</p>';
        return;
    }
    
    // Calculate ROIC components
    const latestSales = annual.sales[n-1];
    const latestPBT = annual.pbt[n-1] || 0;
    const latestInterest = annual.interest[n-1] || 0;
    const latestTax = annual.tax[n-1] || 0;
    
    // NOPAT = EBIT × (1 - Tax Rate)
    const ebit = latestPBT + latestInterest;
    const taxRate = latestPBT > 0 ? (latestTax / latestPBT) : 0.25;
    const nopat = ebit * (1 - taxRate);
    
    // Invested Capital = Equity + Debt - Excess Cash
    const latestEquity = (annual.equity[n-1] || 0) + (annual.reserves[n-1] || 0);
    const latestDebt = annual.borrowings[n-1] || 0;
    const latestCash = annual.cash[n-1] || 0;
    const excessCash = Math.max(0, latestCash - (latestSales * 0.02)); // Keep 2% as working cash
    const investedCapital = latestEquity + latestDebt - excessCash;
    
    // ROIC = NOPAT / Invested Capital
    const roic = investedCapital > 0 ? (nopat / investedCapital) * 100 : null;
    
    // ROIC Drivers
    const nopatMargin = latestSales > 0 ? (nopat / latestSales) * 100 : null;
    const icTurnover = investedCapital > 0 ? latestSales / investedCapital : null;
    
    // Strategy Classification (Mauboussin)
    let strategy = '';
    let strategyColor = '';
    
    if (nopatMargin > 15 && icTurnover < 2) {
        strategy = '🎯 Differentiation Strategy';
        strategyColor = 'var(--accent-primary)';
    } else if (nopatMargin < 10 && icTurnover > 3) {
        strategy = '📊 Cost Leadership Strategy';
        strategyColor = 'var(--accent-warning)';
    } else if (nopatMargin > 15 && icTurnover > 2) {
        strategy = '⭐ Exceptional - Both High!';
        strategyColor = 'var(--accent-success)';
    } else {
        strategy = '⚖️ Balanced Model';
        strategyColor = 'var(--text-secondary)';
    }
    
    const html = `
        <div style="display: grid; grid-template-columns: 2fr 1fr; gap: 2rem; margin-bottom: 1.5rem;">
            <div>
                <div class="dupont-breakdown">
                    <div class="dupont-factor">
                        <span class="factor-name">NOPAT Margin</span>
                        <span class="factor-value">${formatPercent(nopatMargin)}</span>
                    </div>
                    <div class="dupont-factor">
                        <span class="factor-name">× Invested Capital Turnover</span>
                        <span class="factor-value">${icTurnover ? icTurnover.toFixed(2) + 'x' : 'N/A'}</span>
                    </div>
                    <div class="dupont-factor" style="border-top: 2px solid var(--border-color); margin-top: 1rem; padding-top: 1rem;">
                        <span class="factor-name" style="font-weight: 700;">= ROIC</span>
                        <span class="factor-value" style="color: var(--accent-primary); font-size: 1.5rem;">${formatPercent(roic)}</span>
                    </div>
                </div>
            </div>
            <div style="padding: 1.5rem; background: var(--bg-secondary); border-radius: 8px; border-left: 3px solid var(--accent-primary);">
                <div style="font-size: 0.85rem; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 0.5rem;">Strategy</div>
                <div style="font-size: 1.1rem; font-weight: 700; color: ${strategyColor}; margin-bottom: 1rem;">
                    ${strategy}
                </div>
                <div style="font-size: 0.85rem; color: var(--text-secondary); line-height: 1.6;">
                    ${nopatMargin > 15 ? 'High margins suggest pricing power, brand strength, or unique products.' : 
                      nopatMargin > 10 ? 'Moderate margins - healthy business.' : 
                      'Low margins - competing on volume/efficiency.'}
                    ${icTurnover > 3 ? ' Very capital efficient.' : 
                      icTurnover > 1.5 ? ' Decent capital efficiency.' : 
                      ' Capital intensive model.'}
                </div>
            </div>
        </div>
        <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 1rem; padding: 1rem; background: var(--bg-secondary); border-radius: 6px;">
            <div>
                <div style="font-size: 0.75rem; color: var(--text-muted); margin-bottom: 0.25rem;">NOPAT</div>
                <div style="font-family: var(--font-mono); font-size: 1.1rem; font-weight: 600;">${formatNumber(nopat)} Cr</div>
            </div>
            <div>
                <div style="font-size: 0.75rem; color: var(--text-muted); margin-bottom: 0.25rem;">Invested Capital</div>
                <div style="font-family: var(--font-mono); font-size: 1.1rem; font-weight: 600;">${formatNumber(investedCapital)} Cr</div>
            </div>
            <div>
                <div style="font-size: 0.75rem; color: var(--text-muted); margin-bottom: 0.25rem;">Tax Rate</div>
                <div style="font-family: var(--font-mono); font-size: 1.1rem; font-weight: 600;">${(taxRate * 100).toFixed(1)}%</div>
            </div>
        </div>
    `;
    
    document.getElementById('dupontAnalysis').innerHTML = html;
}

// ============================================================================
// REVISED: EFFICIENCY METRICS WITH INSIGHTS
// ============================================================================

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
    const ccc = debtorDays + inventoryDays - 30; // Assuming 30 days payable
    
    // Trends
    let assetTurnoverTrend = '';
    let debtorDaysTrend = '';
    let cccTrend = '';
    
    if (n >= 4) {
        const oldAssetTurnover = annual.totalAssets[n-4] > 0 ? annual.sales[n-4] / annual.totalAssets[n-4] : null;
        const oldDebtorDays = annual.sales[n-4] > 0 ? (annual.receivables[n-4] / annual.sales[n-4]) * 365 : null;
        
        if (assetTurnover && oldAssetTurnover) {
            assetTurnoverTrend = assetTurnover > oldAssetTurnover ? ' ↗️ Improving' : ' ↘️ Declining';
        }
        if (debtorDays && oldDebtorDays) {
            debtorDaysTrend = debtorDays < oldDebtorDays ? ' ↗️ Improving' : ' ↘️ Worsening';
        }
    }
    
    // Insights
    let atInsight = '';
    if (assetTurnover > 3) {
        atInsight = '🟢 Highly capital efficient - asset-light model';
    } else if (assetTurnover > 1.5) {
        atInsight = '🟡 Moderate capital efficiency';
    } else {
        atInsight = '🔴 Capital intensive - needs high margins to justify';
    }
    
    let debtorInsight = '';
    if (debtorDays < 30) {
        debtorInsight = '🟢 Excellent collection - strong bargaining power';
    } else if (debtorDays < 60) {
        debtorInsight = '🟡 Normal collection period';
    } else {
        debtorInsight = '⚠️ Slow collections - verify customer quality';
    }
    
    let cccInsight = '';
    if (ccc < 0) {
        cccInsight = '💰 Negative CCC - Earns from float!';
    } else if (ccc < 60) {
        cccInsight = '🟢 Efficient working capital management';
    } else if (ccc < 90) {
        cccInsight = '🟡 Average working capital cycle';
    } else {
        cccInsight = '⚠️ Cash tied up in operations - needs improvement';
    }
    
    const html = `
        <div style="display: flex; flex-direction: column; gap: 1rem;">
            <div class="dupont-factor">
                <div style="display: flex; flex-direction: column;">
                    <span class="factor-name">Asset Turnover${assetTurnoverTrend}</span>
                    <span style="font-size: 0.75rem; color: var(--text-muted); margin-top: 0.25rem;">${atInsight}</span>
                </div>
                <span class="factor-value">${assetTurnover ? assetTurnover.toFixed(2) + 'x' : 'N/A'}</span>
            </div>
            <div class="dupont-factor">
                <div style="display: flex; flex-direction: column;">
                    <span class="factor-name">Debtor Days${debtorDaysTrend}</span>
                    <span style="font-size: 0.75rem; color: var(--text-muted); margin-top: 0.25rem;">${debtorInsight}</span>
                </div>
                <span class="factor-value">${debtorDays ? Math.round(debtorDays) + ' days' : 'N/A'}</span>
            </div>
            <div class="dupont-factor">
                <div style="display: flex; flex-direction: column;">
                    <span class="factor-name">Inventory Days</span>
                    <span style="font-size: 0.75rem; color: var(--text-muted); margin-top: 0.25rem;">
                        ${inventoryDays ? Math.round(inventoryDays) + ' days turnover' : 'N/A'}
                    </span>
                </div>
                <span class="factor-value">${inventoryDays ? Math.round(inventoryDays) + ' days' : 'N/A'}</span>
            </div>
            <div class="dupont-factor">
                <div style="display: flex; flex-direction: column;">
                    <span class="factor-name">Cash Conversion Cycle</span>
                    <span style="font-size: 0.75rem; color: var(--text-muted); margin-top: 0.25rem;">${cccInsight}</span>
                </div>
                <span class="factor-value">${ccc ? Math.round(ccc) + ' days' : 'N/A'}</span>
            </div>
        </div>
    `;
    
    document.getElementById('efficiencyMetrics').innerHTML = html;
}

// Leverage Metrics (UNCHANGED)
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

// ============================================================================
// REVISED: MOAT ANALYSIS - EXPANDED TO 6 TESTS (14 POINTS)
// ============================================================================

function displayMoatAnalysis() {
    const { annual } = companyData;
    const n = annual.years.length;
    
    let moatScore = 0;
    const indicators = [];
    
    // 1. High & Sustained ROIC (3 pts)
    let highROICYears = 0;
    for (let i = Math.max(0, n - 7); i < n; i++) {
        const equity = (annual.equity[i] || 0) + (annual.reserves[i] || 0);
        const debt = annual.borrowings[i] || 0;
        const capital = equity + debt;
        const nopat = (annual.netProfit[i] || 0) * 1.15;
        const roic = capital > 0 ? (nopat / capital) * 100 : null;
        
        if (roic && roic > 15) highROICYears++;
    }
    
    let roicPoints = 0;
    if (highROICYears >= 7) roicPoints = 3;
    else if (highROICYears >= 5) roicPoints = 2;
    else if (highROICYears >= 3) roicPoints = 1;
    
    moatScore += roicPoints;
    indicators.push({ 
        name: 'High & Sustained ROIC', 
        pass: highROICYears >= 5, 
        value: `${highROICYears}/7 years`,
        points: `${roicPoints}/3`
    });
    
    // 2. ROIC Stability (2 pts)
    const roics = [];
    for (let i = Math.max(0, n - 7); i < n; i++) {
        const equity = (annual.equity[i] || 0) + (annual.reserves[i] || 0);
        const debt = annual.borrowings[i] || 0;
        const capital = equity + debt;
        const nopat = (annual.netProfit[i] || 0) * 1.15;
        const roic = capital > 0 ? (nopat / capital) * 100 : null;
        if (roic) roics.push(roic);
    }
    
    let roicStdDev = 0;
    if (roics.length > 0) {
        const avgROIC = roics.reduce((a, b) => a + b, 0) / roics.length;
        roicStdDev = Math.sqrt(roics.reduce((sum, r) => sum + Math.pow(r - avgROIC, 2), 0) / roics.length);
    }
    
    let stabilityPoints = 0;
    if (roicStdDev < 5) stabilityPoints = 2;
    else if (roicStdDev < 10) stabilityPoints = 1;
    
    moatScore += stabilityPoints;
    indicators.push({ 
        name: 'ROIC Stability', 
        pass: roicStdDev < 5, 
        value: `σ = ${roicStdDev.toFixed(1)}%`,
        points: `${stabilityPoints}/2`
    });
    
    // 3. Pricing Power - Margin Stability (3 pts)
    const margins = annual.sales.slice(-5).map((s, i) => {
        const idx = n - 5 + i;
        if (!s || s === 0) return null;
        const profit = annual.netProfit[idx];
        return (profit / s) * 100;
    }).filter(m => m !== null);
    
    let marginStdDev = 0;
    let marginTrend = 0;
    if (margins.length >= 3) {
        const avgMargin = margins.reduce((a, b) => a + b, 0) / margins.length;
        marginStdDev = Math.sqrt(margins.reduce((sum, m) => sum + Math.pow(m - avgMargin, 2), 0) / margins.length);
        marginTrend = margins[margins.length - 1] - margins[0];
    }
    
    let pricingPoints = 0;
    if (marginTrend > 0 && marginStdDev < 3) pricingPoints = 3;
    else if (marginStdDev < 3 || marginTrend > 0) pricingPoints = 2;
    else if (marginStdDev < 5) pricingPoints = 1;
    
    moatScore += pricingPoints;
    indicators.push({ 
        name: 'Pricing Power', 
        pass: marginStdDev < 3, 
        value: `Margin σ = ${marginStdDev.toFixed(1)}%`,
        points: `${pricingPoints}/3`
    });
    
    // 4. Scale Advantages - Improving Efficiency (2 pts)
    let efficiencyPoints = 0;
    if (n >= 4) {
        const latestAssetTurnover = annual.totalAssets[n-1] > 0 ? annual.sales[n-1] / annual.totalAssets[n-1] : 0;
        const oldAssetTurnover = annual.totalAssets[n-4] > 0 ? annual.sales[n-4] / annual.totalAssets[n-4] : 0;
        
        if (latestAssetTurnover > oldAssetTurnover * 1.1) {
            efficiencyPoints = 2;
        } else if (latestAssetTurnover > oldAssetTurnover) {
            efficiencyPoints = 1;
        }
    }
    
    moatScore += efficiencyPoints;
    indicators.push({ 
        name: 'Scale Advantages', 
        pass: efficiencyPoints >= 1, 
        value: efficiencyPoints >= 1 ? 'Improving efficiency' : 'Stable',
        points: `${efficiencyPoints}/2`
    });
    
    // 5. Customer Stickiness - Working Capital (2 pts)
    let stickinessPoints = 0;
    if (n >= 4) {
        const latestDebtorDays = annual.sales[n-1] > 0 ? (annual.receivables[n-1] / annual.sales[n-1]) * 365 : 999;
        const oldDebtorDays = annual.sales[n-4] > 0 ? (annual.receivables[n-4] / annual.sales[n-4]) * 365 : 999;
        
        const salesGrowth = annual.sales[n-1] / annual.sales[n-4];
        
        if (salesGrowth > 1.3 && latestDebtorDays <= oldDebtorDays) {
            stickinessPoints = 2;
        } else if (latestDebtorDays <= oldDebtorDays) {
            stickinessPoints = 1;
        }
    }
    
    moatScore += stickinessPoints;
    indicators.push({ 
        name: 'Customer Stickiness', 
        pass: stickinessPoints >= 1, 
        value: stickinessPoints >= 1 ? 'Strong retention' : 'Variable',
        points: `${stickinessPoints}/2`
    });
    
    // 6. Core Earnings Quality (2 pts)
    const latestSales = annual.sales[n-1];
    const opProfit = latestSales - (annual.rawMaterial[n-1] || 0) - (annual.employeeCost[n-1] || 0) - (annual.sellingAdmin[n-1] || 0);
    const otherIncome = annual.otherIncome[n-1] || 0;
    
    let corePoints = 0;
    if (otherIncome < opProfit * 0.1) {
        corePoints = 2;
    } else if (otherIncome < opProfit * 0.3) {
        corePoints = 1;
    }
    
    moatScore += corePoints;
    indicators.push({ 
        name: 'Core Earnings Quality', 
        pass: otherIncome < opProfit * 0.1, 
        value: `Other Income ${((otherIncome/opProfit)*100).toFixed(0)}% of Op Profit`,
        points: `${corePoints}/2`
    });
    
    // Moat Classification
    let moatRating = 'No Moat';
    let moatColor = 'var(--text-muted)';
    if (moatScore >= 11) {
        moatRating = 'Exceptional Moat';
        moatColor = 'var(--accent-success)';
    } else if (moatScore >= 8) {
        moatRating = 'Wide Moat';
        moatColor = 'var(--accent-primary)';
    } else if (moatScore >= 5) {
        moatRating = 'Narrow Moat';
        moatColor = 'var(--accent-warning)';
    }
    
    const html = `
        <div style="margin-bottom: 1.5rem; padding: 1rem; background: var(--bg-secondary); border-radius: 6px;">
            <div style="font-size: 1.5rem; font-weight: 700; margin-bottom: 0.5rem; color: ${moatColor};">
                ${moatRating}
            </div>
            <div style="font-size: 0.9rem; color: var(--text-secondary);">
                Score: ${moatScore}/14 points
            </div>
        </div>
        <div class="moat-indicators">
            ${indicators.map(ind => `
                <div class="indicator">
                    <div>
                        <div style="font-weight: 600;">${ind.name}</div>
                        <div style="font-size: 0.85rem; color: var(--text-secondary);">${ind.value}</div>
                    </div>
                    <span class="indicator-status ${ind.pass ? 'pass' : 'fail'}">
                        ${ind.pass ? '✓' : '✗'} ${ind.points}
                    </span>
                </div>
            `).join('')}
        </div>
    `;
    
    document.getElementById('moatAnalysis').innerHTML = html;
}

// ============================================================================
// REVISED: CAPITAL ALLOCATION - INCREMENTAL ROIC FOCUSED
// ============================================================================

function displayCapitalAllocation() {
    const { annual } = companyData;
    const n = annual.years.length;
    
    if (n < 4) {
        document.getElementById('capitalAllocation').innerHTML = '<p style="color: var(--text-muted);">Insufficient data (requires 4+ years)</p>';
        return;
    }
    
    // Calculate Incremental ROIC for last 3 years
    const incROICs = [];
    for (let i = 1; i < Math.min(n, 4); i++) {
        const idx = n - i - 1;
        if (idx < 0) continue;
        
        const currEquity = (annual.equity[idx+1] || 0) + (annual.reserves[idx+1] || 0);
        const prevEquity = (annual.equity[idx] || 0) + (annual.reserves[idx] || 0);
        const currDebt = annual.borrowings[idx+1] || 0;
        const prevDebt = annual.borrowings[idx] || 0;
        
        const deltaCapital = (currEquity + currDebt) - (prevEquity + prevDebt);
        const deltaNOPAT = (annual.netProfit[idx+1] || 0) - (annual.netProfit[idx] || 0);
        
        if (deltaCapital > 0) {
            const incROIC = (deltaNOPAT / deltaCapital) * 100;
            incROICs.push(incROIC);
        }
    }
    
    const avgIncROIC = incROICs.length > 0 ? incROICs.reduce((a, b) => a + b, 0) / incROICs.length : null;
    
    // Score based on Incremental ROIC (60% weight = 60 points)
    let reinvestmentScore = 0;
    if (avgIncROIC > 25) {
        reinvestmentScore = 60;
    } else if (avgIncROIC > 18) {
        reinvestmentScore = 45;
    } else if (avgIncROIC > 13) {
        reinvestmentScore = 30;
    } else if (avgIncROIC > 5) {
        reinvestmentScore = 15;
    }
    
    // Deployment Logic Score (40% weight = 40 points)
    let deploymentScore = 0;
    
    const payoutRatios = [];
    for (let i = Math.max(0, n - 3); i < n; i++) {
        if (annual.netProfit[i] > 0) {
            payoutRatios.push((annual.dividend[i] / annual.netProfit[i]) * 100);
        }
    }
    const avgPayout = payoutRatios.length > 0 ? payoutRatios.reduce((a, b) => a + b, 0) / payoutRatios.length : 0;
    
    if (avgIncROIC > 20) {
        // High returns - should retain
        if (avgPayout < 30) {
            deploymentScore = 40; // Correctly reinvesting
        } else if (avgPayout < 50) {
            deploymentScore = 25;
        } else {
            deploymentScore = 10; // Paying out too much
        }
    } else if (avgIncROIC < 13) {
        // Low returns - should return cash
        if (avgPayout > 60) {
            deploymentScore = 40; // Correctly returning cash
        } else if (avgPayout > 40) {
            deploymentScore = 25;
        } else {
            deploymentScore = 10; // Hoarding cash
        }
    } else {
        // Moderate returns - balanced
        if (avgPayout >= 30 && avgPayout <= 60) {
            deploymentScore = 35;
        } else {
            deploymentScore = 20;
        }
    }
    
    const totalScore = reinvestmentScore + deploymentScore;
    
    // Grade
    let grade = 'F';
    let gradeColor = 'var(--accent-danger)';
    if (totalScore >= 85) { grade = 'A'; gradeColor = 'var(--accent-success)'; }
    else if (totalScore >= 70) { grade = 'B'; gradeColor = 'var(--accent-primary)'; }
    else if (totalScore >= 55) { grade = 'C'; gradeColor = 'var(--accent-warning)'; }
    else if (totalScore >= 35) { grade = 'D'; gradeColor = 'var(--text-secondary)'; }
    
    // Recommendation
    let recommendation = '';
    if (avgIncROIC > 20 && avgPayout > 50) {
        recommendation = '⚠️ Company earning high returns but distributing too much. Should retain more for growth.';
    } else if (avgIncROIC < 13 && avgPayout < 40) {
        recommendation = '⚠️ Low incremental returns. Should return more cash to shareholders via dividends or buybacks.';
    } else if (avgIncROIC > 20 && avgPayout < 30) {
        recommendation = '✅ Excellent! High returns on new capital and correctly retaining earnings for reinvestment.';
    } else if (avgIncROIC < 13 && avgPayout > 60) {
        recommendation = '✅ Prudent! Low returns, correctly returning cash to shareholders.';
    } else {
        recommendation = '⚖️ Balanced approach to capital allocation.';
    }
    
    const html = `
        <div style="margin-bottom: 1.5rem; padding: 1.5rem; background: var(--bg-secondary); border-radius: 6px; border-left: 3px solid ${gradeColor};">
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <div>
                    <div style="font-size: 0.75rem; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.05em;">Capital Allocation Grade</div>
                    <div style="font-size: 3rem; font-weight: 700; font-family: var(--font-mono); color: ${gradeColor}; line-height: 1;">
                        ${grade}
                    </div>
                </div>
                <div style="text-align: right;">
                    <div style="font-size: 0.85rem; color: var(--text-secondary);">Total Score</div>
                    <div style="font-size: 1.5rem; font-weight: 700;">${totalScore}/100</div>
                </div>
            </div>
        </div>
        
        <div style="display: flex; flex-direction: column; gap: 1rem;">
            <div class="dupont-factor">
                <span class="factor-name">Avg Incremental ROIC (3Y)</span>
                <span class="factor-value">${formatPercent(avgIncROIC)}</span>
            </div>
            <div class="dupont-factor">
                <span class="factor-name">Reinvestment Quality Score</span>
                <span class="factor-value">${reinvestmentScore}/60 pts</span>
            </div>
            <div class="dupont-factor">
                <span class="factor-name">Deployment Logic Score</span>
                <span class="factor-value">${deploymentScore}/40 pts</span>
            </div>
            <div class="dupont-factor">
                <span class="factor-name">Avg Payout Ratio (3Y)</span>
                <span class="factor-value">${avgPayout.toFixed(1)}%</span>
            </div>
        </div>
        
        <div style="margin-top: 1.5rem; padding: 1rem; background: var(--bg-secondary); border-radius: 6px; font-size: 0.9rem; line-height: 1.6; color: var(--text-secondary);">
            <strong>Analysis:</strong><br/>
            ${recommendation}
        </div>
    `;
    
    document.getElementById('capitalAllocation').innerHTML = html;
}

// Value Migration (UNCHANGED - as requested)
function displayValueMigration() {
    const { annual } = companyData;
    const n = annual.years.length;
    
    if (n < 5) {
        document.getElementById('valueMigration').innerHTML = '<p style="color: var(--text-muted);">Insufficient data (requires 5+ years)</p>';
        return;
    }
    
    const salesCAGR = calculateCAGR(annual.sales, 5);
    const profitCAGR = calculateCAGR(annual.netProfit, 5);
    
    const oldMargin = annual.sales[n-6] > 0 ? (annual.netProfit[n-6] / annual.sales[n-6]) * 100 : null;
    const newMargin = annual.sales[n-1] > 0 ? (annual.netProfit[n-1] / annual.sales[n-1]) * 100 : null;
    const marginChange = newMargin - oldMargin;
    
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

// Continue in next part...

// ============================================================================
// REVISED: EARNING POWER BOX WITH 10-YEAR TRAJECTORY
// ============================================================================

function displayEarningPowerBox() {
    const { annual } = companyData;
    const n = annual.years.length;
    
    if (n < 4) {
        document.getElementById('earningPowerBox').innerHTML = '<p style="color: var(--text-muted);">Requires 3+ years of data</p>';
        return;
    }
    
    // Calculate trajectory for all available years (minimum 3-year windows)
    const trajectory = [];
    
    for (let i = 3; i < n; i++) {
        const startIdx = i - 3;
        const endIdx = i;
        
        // 3-year PAT CAGR
        const patStart = annual.netProfit[startIdx];
        const patEnd = annual.netProfit[endIdx];
        const patCAGR = (patStart > 0 && patEnd > 0) ? 
            (Math.pow(patEnd / patStart, 1/3) - 1) * 100 : null;
        
        // CFO / Net Profit ratio (average of 3 years)
        const cfoRatios = [];
        for (let j = startIdx; j <= endIdx; j++) {
            if (annual.netProfit[j] > 0 && annual.cfo[j]) {
                cfoRatios.push((annual.cfo[j] / annual.netProfit[j]) * 100);
            }
        }
        const avgCFORatio = cfoRatios.length > 0 ? 
            cfoRatios.reduce((a, b) => a + b, 0) / cfoRatios.length : null;
        
        // Determine quadrant
        const highGrowth = patCAGR && patCAGR > 15;
        const highCash = avgCFORatio && avgCFORatio > 80;
        
        let quadrant = '';
        if (highGrowth && highCash) quadrant = 'STAR';
        else if (highGrowth && !highCash) quadrant = 'INVESTIGATE';
        else if (!highGrowth && highCash) quadrant = 'CASH COW';
        else quadrant = 'RED FLAG';
        
        trajectory.push({
            year: annual.years[endIdx],
            growth: patCAGR,
            cash: avgCFORatio,
            quadrant: quadrant
        });
    }
    
    // Current position
    const current = trajectory[trajectory.length - 1];
    
    // Pattern recognition
    let pattern = '';
    if (trajectory.length >= 3) {
        const quadrants = trajectory.map(t => t.quadrant);
        const lastThree = quadrants.slice(-3);
        
        if (lastThree.every(q => q === 'STAR')) {
            pattern = '🌟 Consistent High-Quality Compounder - 3+ years in STAR quadrant';
        } else if (quadrants.includes('RED FLAG') && current.quadrant === 'STAR') {
            pattern = '📈 Successful Turnaround Story - Moved from RED FLAG to STAR';
        } else if (quadrants.includes('STAR') && current.quadrant === 'CASH COW') {
            pattern = '⚠️ Growth Slowing - Matured from STAR to CASH COW';
        } else if (quadrants.slice(-2).every(q => q === current.quadrant)) {
            pattern = `⚖️ Stable Position - Consistent ${current.quadrant}`;
        } else {
            pattern = '🔄 Transitioning - Position changed recently';
        }
    }
    
    const html = `
        <div style="margin-bottom: 1.5rem;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 1rem;">
                <div>
                    <div style="font-size: 0.85rem; color: var(--text-muted);">Latest Position (${current.year})</div>
                    <div style="font-size: 1.25rem; font-weight: 700; color: var(--accent-primary);">
                        ${current.quadrant}
                    </div>
                </div>
                <div style="text-align: right;">
                    <div style="font-size: 0.85rem; color: var(--text-muted);">Growth: ${formatPercent(current.growth)}</div>
                    <div style="font-size: 0.85rem; color: var(--text-muted);">Cash: ${formatPercent(current.cash)}</div>
                </div>
            </div>
            
            <div style="padding: 0.75rem; background: var(--bg-secondary); border-radius: 6px; font-size: 0.85rem; color: var(--text-secondary); margin-bottom: 1rem;">
                <strong>Journey Pattern:</strong> ${pattern}
            </div>
            
            <div style="margin-bottom: 1rem;">
                <div style="font-size: 0.85rem; color: var(--text-muted); margin-bottom: 0.5rem;">Historical Trajectory:</div>
                <div style="display: flex; flex-wrap: wrap; gap: 0.5rem;">
                    ${trajectory.map(t => `
                        <div style="padding: 0.4rem 0.8rem; background: var(--bg-secondary); border-radius: 4px; font-size: 0.8rem;">
                            <strong>${t.year}:</strong> ${t.quadrant}
                        </div>
                    `).join('')}
                </div>
            </div>
        </div>
        
        <div class="epb-grid">
            <div class="epb-quadrant ${current.quadrant === 'INVESTIGATE' ? 'investigate current' : 'investigate'}">
                <div class="epb-quadrant-title">🔍 INVESTIGATE</div>
                <div class="epb-quadrant-desc">High Growth + Low Cash<br/>Capital-intensive growth</div>
            </div>
            <div class="epb-quadrant ${current.quadrant === 'STAR' ? 'star current' : 'star'}">
                <div class="epb-quadrant-title">⭐ STAR</div>
                <div class="epb-quadrant-desc">High Growth + High Cash<br/>Ideal Investment</div>
            </div>
            <div class="epb-quadrant ${current.quadrant === 'RED FLAG' ? 'redflag current' : 'redflag'}">
                <div class="epb-quadrant-title">⚠️ RED FLAG</div>
                <div class="epb-quadrant-desc">Low Growth + Low Cash<br/>Avoid or Deep Value</div>
            </div>
            <div class="epb-quadrant ${current.quadrant === 'CASH COW' ? 'cashcow current' : 'cashcow'}">
                <div class="epb-quadrant-title">💰 CASH COW</div>
                <div class="epb-quadrant-desc">Low Growth + High Cash<br/>Mature Dividend Play</div>
            </div>
        </div>
    `;
    
    document.getElementById('earningPowerBox').innerHTML = html;
}

// ============================================================================
// REVISED: CAP ANALYSIS WITH COMPANY-SPECIFIC WACC
// ============================================================================

function displayCAPAnalysis() {
    const { annual } = companyData;
    const n = annual.years.length;
    
    if (n < 5) {
        document.getElementById('capAnalysis').innerHTML = '<p style="color: var(--text-muted);">Requires 5+ years of data</p>';
        return;
    }
    
    // Calculate company-specific WACC
    const latestEquity = (annual.equity[n-1] || 0) + (annual.reserves[n-1] || 0);
    const latestDebt = annual.borrowings[n-1] || 0;
    const totalCapital = latestEquity + latestDebt;
    
    // Cost of Equity (simplified CAPM)
    const riskFreeRate = 7.0; // 10-year G-Sec India
    const equityRiskPremium = 6.0;
    const beta = 1.0; // Market average
    const costOfEquity = riskFreeRate + (beta * equityRiskPremium);
    
    // Cost of Debt (from actual data)
    const interestExpense = annual.interest[n-1] || 0;
    let costOfDebt = 0;
    let afterTaxCostOfDebt = 0;
    
    if (latestDebt > 0 && interestExpense > 0) {
        costOfDebt = (interestExpense / latestDebt) * 100;
        const taxRate = annual.pbt[n-1] > 0 ? (annual.tax[n-1] / annual.pbt[n-1]) : 0.25;
        afterTaxCostOfDebt = costOfDebt * (1 - taxRate);
    }
    
    // WACC Calculation
    let wacc;
    if (totalCapital > 0) {
        const equityWeight = latestEquity / totalCapital;
        const debtWeight = latestDebt / totalCapital;
        wacc = (equityWeight * costOfEquity) + (debtWeight * afterTaxCostOfDebt);
    } else {
        wacc = costOfEquity; // All equity
    }
    
    // Calculate ROIC for each year
    const roics = [];
    for (let i = 0; i < n; i++) {
        const equity = (annual.equity[i] || 0) + (annual.reserves[i] || 0);
        const debt = annual.borrowings[i] || 0;
        const capital = equity + debt;
        const nopat = (annual.netProfit[i] || 0) * 1.15;
        
        if (capital > 0) {
            roics.push((nopat / capital) * 100);
        } else {
            roics.push(null);
        }
    }
    
    // Count years with ROIC > WACC
    const capYears = roics.filter(r => r && r > wacc).length;
    
    const html = `
        <div style="display: grid; grid-template-columns: 2fr 1fr; gap: 2rem;">
            <div>
                <canvas id="capChart"></canvas>
            </div>
            <div>
                <div style="padding: 1.5rem; background: var(--bg-secondary); border-radius: 6px;">
                    <div style="font-size: 2.5rem; font-weight: 700; margin-bottom: 0.5rem; color: var(--accent-primary);">${capYears} Years</div>
                    <div style="color: var(--text-secondary); margin-bottom: 1.5rem; font-size: 0.9rem;">Competitive Advantage Period</div>
                    <div style="font-size: 0.9rem; line-height: 1.6; margin-bottom: 1.5rem;">
                        ${capYears >= 7 ? '✅ <strong>Strong CAP</strong> - Sustainable competitive advantage' :
                          capYears >= 5 ? '⚠️ <strong>Moderate CAP</strong> - Some competitive position' :
                          '⚠️ <strong>Weak CAP</strong> - Commodity-like business'}
                    </div>
                    <div style="padding-top: 1rem; border-top: 1px solid var(--border-color); font-size: 0.85rem;">
                        <div style="margin-bottom: 0.5rem;">
                            <span style="color: var(--text-muted);">Company WACC:</span>
                            <strong style="float: right;">${wacc.toFixed(1)}%</strong>
                        </div>
                        <div style="margin-bottom: 0.5rem;">
                            <span style="color: var(--text-muted);">Equity Weight:</span>
                            <strong style="float: right;">${totalCapital > 0 ? ((latestEquity/totalCapital)*100).toFixed(0) : 100}%</strong>
                        </div>
                        <div>
                            <span style="color: var(--text-muted);">Debt Weight:</span>
                            <strong style="float: right;">${totalCapital > 0 ? ((latestDebt/totalCapital)*100).toFixed(0) : 0}%</strong>
                        </div>
                    </div>
                    <div style="margin-top: 1.5rem; padding-top: 1rem; border-top: 1px solid var(--border-color);">
                        <div style="font-size: 0.85rem; color: var(--text-secondary);">
                            <strong>Note:</strong> Companies with CAP > 10 years typically have wide moats. Years counted where ROIC > WACC.
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;
    
    document.getElementById('capAnalysis').innerHTML = html;
    
    // Draw CAP chart
    setTimeout(() => {
        const ctx = document.getElementById('capChart');
        if (ctx) {
            new Chart(ctx, {
                type: 'line',
                data: {
                    labels: annual.years,
                    datasets: [
                        {
                            label: 'ROIC %',
                            data: roics,
                            borderColor: '#2ecc71',
                            backgroundColor: 'rgba(46, 204, 113, 0.1)',
                            tension: 0.3,
                            fill: false,
                            borderWidth: 3
                        },
                        {
                            label: `WACC (${wacc.toFixed(1)}%)`,
                            data: Array(n).fill(wacc),
                            borderColor: '#e74c3c',
                            borderDash: [5, 5],
                            fill: false,
                            borderWidth: 2
                        }
                    ]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: true,
                    plugins: {
                        legend: {
                            labels: { color: '#ffffff' }
                        }
                    },
                    scales: {
                        y: {
                            ticks: { 
                                color: '#a0a0a0',
                                callback: value => value + '%'
                            },
                            grid: { color: '#333333' }
                        },
                        x: {
                            ticks: { color: '#a0a0a0' },
                            grid: { color: '#333333' }
                        }
                    }
                }
            });
        }
    }, 100);
}

// ============================================================================
// REVISED: CAPEX SPLIT WITH INSIGHTS
// ============================================================================

function displayCapexSplit() {
    const { annual } = companyData;
    const n = annual.years.length;
    
    if (n < 3) {
        document.getElementById('capexSplit').innerHTML = '<p style="color: var(--text-muted);">Requires 3+ years of data</p>';
        return;
    }
    
    const capexData = [];
    
    for (let i = Math.max(0, n - 5); i < n; i++) {
        const depreciation = annual.depreciation[i] || 0;
        const cfi = Math.abs(annual.cfi[i] || 0);
        const maintenanceCapex = depreciation;
        const growthCapex = Math.max(0, cfi - depreciation);
        const sales = annual.sales[i] || 1;
        const capexIntensity = (cfi / sales) * 100;
        
        capexData.push({
            year: annual.years[i],
            maintenance: maintenanceCapex,
            growth: growthCapex,
            total: cfi,
            intensity: capexIntensity,
            cfo: annual.cfo[i] || 0
        });
    }
    
    // Calculate insights
    const latest = capexData[capexData.length - 1];
    const oldest = capexData[0];
    
    // 1. Capex Intensity Trend
    const intensityTrend = latest.intensity - oldest.intensity;
    const intensityInsight = intensityTrend < -2 ? 
        '🟢 Becoming more capital efficient' : 
        intensityTrend > 2 ? 
        '⚠️ Capital intensity increasing - verify productivity' : 
        '⚖️ Stable capital intensity';
    
    // 2. Maintenance vs Depreciation
    const maintVsDepr = latest.maintenance / (annual.depreciation[n-1] || 1);
    const maintInsight = maintVsDepr < 0.9 ? 
        '⚠️ Under-investing in asset maintenance' : 
        maintVsDepr > 1.3 ? 
        '⚠️ High replacement needs - capital intensive' : 
        '🟢 Appropriate maintenance spending';
    
    // 3. Growth Capex Productivity (3-year lag)
    let growthCapexROI = null;
    let roiInsight = 'N/A';
    if (n >= 4) {
        const investedCapex = capexData.slice(0, -1).reduce((sum, d) => sum + d.growth, 0);
        const ebitChange = (annual.pbt[n-1] + annual.interest[n-1]) - (annual.pbt[n-4] + annual.interest[n-4]);
        if (investedCapex > 0) {
            growthCapexROI = (ebitChange / investedCapex) * 100;
            roiInsight = growthCapexROI > 20 ? 
                '🟢 Growth investments paying off well' : 
                growthCapexROI > 10 ? 
                '🟡 Moderate returns on growth capex' : 
                '⚠️ Growth capex not generating sufficient returns';
        }
    }
    
    // 4. Owner Earnings
    const ownerEarnings = latest.cfo - latest.maintenance;
    const marketCap = companyData.meta.marketCap || 0;
    const ownerEarningsYield = marketCap > 0 ? (ownerEarnings / marketCap) * 100 : null;
    const oeInsight = ownerEarningsYield > 5 ? 
        '🟢 Strong cash generation for shareholders' : 
        ownerEarningsYield > 3 ? 
        '🟡 Decent owner earnings yield' : 
        '⚠️ Low owner earnings yield';
    
    // 5. FCF after Maintenance
    const fcfAfterMaint = latest.cfo - latest.maintenance;
    const fcfInsight = fcfAfterMaint > 0 && fcfAfterMaint > oldest.cfo - oldest.maintenance ? 
        '🟢 Self-sustaining + can fund growth' : 
        fcfAfterMaint > 0 ? 
        '🟡 Generates free cash flow' : 
        '⚠️ Needs external capital to grow';
    
    const html = `
        <div style="margin-bottom: 1.5rem;">
            <div style="display: flex; flex-direction: column; gap: 1rem;">
                ${capexData.map((d, idx) => `
                    <div class="dupont-factor">
                        <span class="factor-name">FY${d.year}</span>
                        <span class="factor-value" style="font-size: 0.85rem;">
                            M: ${formatNumber(d.maintenance)} | G: ${formatNumber(d.growth)} | Total: ${formatNumber(d.total)} Cr
                        </span>
                    </div>
                `).join('')}
            </div>
        </div>
        
        <div style="display: flex; flex-direction: column; gap: 0.75rem; padding: 1rem; background: var(--bg-secondary); border-radius: 6px; font-size: 0.85rem;">
            <div>
                <strong>1. Capex Intensity:</strong> ${latest.intensity.toFixed(1)}% of sales
                <div style="color: var(--text-secondary); margin-top: 0.25rem;">${intensityInsight}</div>
            </div>
            <div>
                <strong>2. Maintenance Spending:</strong> ${(maintVsDepr * 100).toFixed(0)}% of depreciation
                <div style="color: var(--text-secondary); margin-top: 0.25rem;">${maintInsight}</div>
            </div>
            <div>
                <strong>3. Growth Capex ROI:</strong> ${growthCapexROI ? growthCapexROI.toFixed(1) + '%' : 'N/A'}
                <div style="color: var(--text-secondary); margin-top: 0.25rem;">${roiInsight}</div>
            </div>
            <div>
                <strong>4. Owner Earnings:</strong> ${formatNumber(ownerEarnings)} Cr (${ownerEarningsYield ? ownerEarningsYield.toFixed(1) + '% yield' : 'N/A'})
                <div style="color: var(--text-secondary); margin-top: 0.25rem;">${oeInsight}</div>
            </div>
            <div>
                <strong>5. FCF after Maintenance:</strong> ${formatNumber(fcfAfterMaint)} Cr
                <div style="color: var(--text-secondary); margin-top: 0.25rem;">${fcfInsight}</div>
            </div>
        </div>
        
        <div style="margin-top: 1rem; padding: 0.75rem; background: var(--bg-secondary); border-radius: 6px; font-size: 0.85rem; color: var(--text-secondary);">
            <strong>Method:</strong> Maintenance Capex ≈ Depreciation | Owner Earnings = CFO - Maintenance Capex
        </div>
    `;
    
    document.getElementById('capexSplit').innerHTML = html;
}

// Incremental ROIC (UNCHANGED)
function displayIncrementalROIC() {
    const { annual } = companyData;
    const n = annual.years.length;
    
    if (n < 4) {
        document.getElementById('incrementalROIC').innerHTML = '<p style="color: var(--text-muted);">Requires 4+ years of data</p>';
        return;
    }
    
    const rows = [];
    for (let i = 1; i < Math.min(n, 6); i++) {
        const idx = n - i - 1;
        if (idx < 0) continue;
        
        const currEquity = (annual.equity[idx+1] || 0) + (annual.reserves[idx+1] || 0);
        const prevEquity = (annual.equity[idx] || 0) + (annual.reserves[idx] || 0);
        const currDebt = annual.borrowings[idx+1] || 0;
        const prevDebt = annual.borrowings[idx] || 0;
        
        const deltaCapital = (currEquity + currDebt) - (prevEquity + prevDebt);
        const deltaNOPAT = (annual.netProfit[idx+1] || 0) - (annual.netProfit[idx] || 0);
        
        const incROIC = deltaCapital > 0 ? (deltaNOPAT / deltaCapital) * 100 : null;
        
        const avgCapital = ((currEquity + prevEquity) / 2) + ((currDebt + prevDebt) / 2);
        const histROIC = avgCapital > 0 ? ((annual.netProfit[idx+1] || 0) * 1.15 / avgCapital) * 100 : null;
        
        rows.push({
            year: annual.years[idx+1],
            deltaCapital,
            deltaNOPAT,
            incROIC,
            histROIC,
            better: incROIC && histROIC && incROIC > histROIC
        });
    }
    
    const avgIncROIC = rows.filter(r => r.incROIC).reduce((sum, r) => sum + r.incROIC, 0) / rows.filter(r => r.incROIC).length;
    const grade = avgIncROIC > 25 ? 'A' : avgIncROIC > 18 ? 'B' : avgIncROIC > 13 ? 'C' : avgIncROIC > 5 ? 'D' : 'F';
    
    const html = `
        <div style="margin-bottom: 1.5rem; padding: 1rem; background: var(--bg-secondary); border-radius: 6px;">
            <div style="font-size: 2rem; font-weight: 700; font-family: var(--font-mono);">Grade ${grade}</div>
            <div style="font-size: 0.9rem; color: var(--text-secondary); margin-top: 0.25rem;">
                Avg Incremental ROIC: ${formatPercent(avgIncROIC)}
            </div>
        </div>
        <div style="display: flex; flex-direction: column; gap: 0.75rem;">
            ${rows.map(row => `
                <div class="dupont-factor">
                    <span class="factor-name">FY${row.year}</span>
                    <span class="factor-value" style="font-size: 0.85rem; ${row.better ? 'color: var(--accent-success)' : ''}">
                        Inc: ${formatPercent(row.incROIC)} vs Hist: ${formatPercent(row.histROIC)}
                        ${row.better ? ' ✓' : ''}
                    </span>
                </div>
            `).join('')}
        </div>
    `;
    
    document.getElementById('incrementalROIC').innerHTML = html;
}

// RM Sensitivity (UNCHANGED)
function displayRMSensitivity() {
    const { annual } = companyData;
    const n = annual.years.length;
    
    if (n < 3) {
        document.getElementById('rmSensitivity').innerHTML = '<p style="color: var(--text-muted);">Requires 3+ years of data</p>';
        return;
    }
    
    const rmIntensities = annual.sales.map((s, i) => {
        if (!s || s === 0) return null;
        const rm = annual.rawMaterial[i] || 0;
        return (rm / s) * 100;
    });
    
    const latestRMIntensity = rmIntensities[n-1];
    
    const grossMargins = annual.sales.map((s, i) => {
        if (!s || s === 0) return null;
        const rm = annual.rawMaterial[i] || 0;
        const other = (annual.otherMfg[i] || 0);
        return ((s - rm - other) / s) * 100;
    }).filter(m => m !== null);
    
    const avgGM = grossMargins.reduce((a, b) => a + b, 0) / grossMargins.length;
    const gmStdDev = Math.sqrt(grossMargins.reduce((sum, m) => sum + Math.pow(m - avgGM, 2), 0) / grossMargins.length);
    
    let sensitivity = 'Low';
    let sensitivityColor = 'var(--accent-success)';
    if (latestRMIntensity > 60) {
        sensitivity = 'High';
        sensitivityColor = 'var(--accent-danger)';
    } else if (latestRMIntensity > 40) {
        sensitivity = 'Medium';
        sensitivityColor = 'var(--accent-warning)';
    }
    
    const html = `
        <div style="margin-bottom: 1.5rem; padding: 1rem; background: var(--bg-secondary); border-radius: 6px;">
            <div style="font-size: 1.5rem; font-weight: 700; color: ${sensitivityColor};">${sensitivity} Sensitivity</div>
            <div style="font-size: 0.9rem; color: var(--text-secondary); margin-top: 0.25rem;">
                RM Intensity: ${formatPercent(latestRMIntensity)}
            </div>
        </div>
        <div style="display: flex; flex-direction: column; gap: 1rem;">
            <div class="dupont-factor">
                <span class="factor-name">RM % of Sales</span>
                <span class="factor-value">${formatPercent(latestRMIntensity)}</span>
            </div>
            <div class="dupont-factor">
                <span class="factor-name">Gross Margin Volatility</span>
                <span class="factor-value">${gmStdDev.toFixed(1)}% σ</span>
            </div>
            <div class="dupont-factor">
                <span class="factor-name">Avg Gross Margin</span>
                <span class="factor-value">${formatPercent(avgGM)}</span>
            </div>
        </div>
        <div style="margin-top: 1rem; padding: 0.75rem; background: var(--bg-secondary); border-radius: 6px; font-size: 0.85rem; color: var(--text-secondary);">
            ${gmStdDev < 3 ? '✅ Low volatility suggests good pricing power' :
              gmStdDev < 5 ? '⚠️ Moderate volatility - some pricing power' :
              '⚠️ High volatility - weak pricing power, vulnerable to RM shocks'}
        </div>
    `;
    
    document.getElementById('rmSensitivity').innerHTML = html;
}

// ============================================================================
// REVISED: BUFFETT TEST - MARKET CAP METHOD
// ============================================================================

function displayBuffettTest() {
    const { annual } = companyData;
    const n = annual.years.length;
    
    if (n < 6) {
        document.getElementById('buffettTest').innerHTML = '<p style="color: var(--text-muted);">Requires 5+ years of data</p>';
        return;
    }
    
    // Check if price data is available
    const hasPriceData = annual.prices && annual.prices.length >= n && annual.prices[n-1] && annual.prices[n-6];
    
    if (!hasPriceData) {
        // Fallback to book value method
        displayBuffettTestBookValue();
        return;
    }
    
    // MARKET CAP METHOD
    
    // 1. Calculate retained earnings over 5 years
    let retainedEarnings = 0;
    for (let i = n - 6; i < n; i++) {
        const profit = annual.netProfit[i] || 0;
        const dividend = annual.dividend[i] || 0;
        retainedEarnings += (profit - dividend);
    }
    
    // 2. Calculate market cap change
    const shares = annual.shares[n-1] || annual.shares[n-6] || 1; // Shares in lakhs
    const oldPrice = annual.prices[n-6];
    const newPrice = annual.prices[n-1];
    
    // Market cap in Crores
    const oldMarketCap = (shares * oldPrice) / 100; // Convert lakhs to crores
    const newMarketCap = (shares * newPrice) / 100;
    
    const marketCapChange = newMarketCap - oldMarketCap;
    
    // 3. Buffett Ratio
    const buffettRatio = retainedEarnings > 0 ? marketCapChange / retainedEarnings : null;
    
    const passed = buffettRatio && buffettRatio >= 1.0;
    
    let interpretation = '';
    if (buffettRatio > 1.5) {
        interpretation = '⭐ Excellent - Created ₹' + buffettRatio.toFixed(2) + ' of market value per ₹1 retained. Management deploying capital brilliantly.';
    } else if (buffettRatio > 1.0) {
        interpretation = '✅ Good - Created ₹' + buffettRatio.toFixed(2) + ' per ₹1 retained. Management earning their keep.';
    } else if (buffettRatio > 0.7) {
        interpretation = '⚠️ Borderline - Created only ₹' + buffettRatio.toFixed(2) + ' per ₹1 retained. Barely met the test.';
    } else {
        interpretation = '❌ Failed - Created only ₹' + buffettRatio.toFixed(2) + ' per ₹1 retained. Management destroying shareholder value. Should have paid higher dividends.';
    }
    
    const html = `
        <div style="margin-bottom: 1.5rem; padding: 1.5rem; background: ${passed ? 'rgba(46, 204, 113, 0.1)' : 'rgba(239, 83, 80, 0.1)'}; border-radius: 6px; border: 2px solid ${passed ? 'var(--accent-success)' : 'var(--accent-danger)'};">
            <div style="font-size: 1.5rem; font-weight: 700; margin-bottom: 0.5rem;">
                ${passed ? '✅ PASSED' : '❌ FAILED'}
            </div>
            <div style="font-size: 1rem; color: var(--text-secondary);">
                Market Value Test: ${buffettRatio ? buffettRatio.toFixed(2) : 'N/A'}x
            </div>
        </div>
        <div style="display: flex; flex-direction: column; gap: 1rem;">
            <div class="dupont-factor">
                <span class="factor-name">Retained Earnings (5Y)</span>
                <span class="factor-value">${formatNumber(retainedEarnings)} Cr</span>
            </div>
            <div class="dupont-factor">
                <span class="factor-name">Old Market Cap (${annual.years[n-6]})</span>
                <span class="factor-value">${formatNumber(oldMarketCap)} Cr</span>
            </div>
            <div class="dupont-factor">
                <span class="factor-name">New Market Cap (${annual.years[n-1]})</span>
                <span class="factor-value">${formatNumber(newMarketCap)} Cr</span>
            </div>
            <div class="dupont-factor">
                <span class="factor-name">Market Cap Change</span>
                <span class="factor-value">${formatNumber(marketCapChange)} Cr</span>
            </div>
            <div class="dupont-factor">
                <span class="factor-name">Value Created per Re. 1</span>
                <span class="factor-value">₹${buffettRatio ? buffettRatio.toFixed(2) : 'N/A'}</span>
            </div>
        </div>
        <div style="margin-top: 1.5rem; padding: 1rem; background: var(--bg-secondary); border-radius: 6px; font-size: 0.85rem; line-height: 1.6; color: var(--text-secondary);">
            <strong>Market Value Test:</strong> For every Re. 1 retained, has market cap increased by at least Re. 1?<br/>
            ${interpretation}
        </div>
    `;
    
    document.getElementById('buffettTest').innerHTML = html;
}

function displayBuffettTestBookValue() {
    // Fallback book value method
    const { annual } = companyData;
    const n = annual.years.length;
    
    let retainedEarnings = 0;
    for (let i = n - 6; i < n; i++) {
        const profit = annual.netProfit[i] || 0;
        const dividend = annual.dividend[i] || 0;
        retainedEarnings += (profit - dividend);
    }
    
    const oldBookValue = (annual.equity[n-6] || 0) + (annual.reserves[n-6] || 0);
    const newBookValue = (annual.equity[n-1] || 0) + (annual.reserves[n-1] || 0);
    const bookValueChange = newBookValue - oldBookValue;
    
    const ratio = retainedEarnings > 0 ? bookValueChange / retainedEarnings : null;
    const passed = ratio && ratio >= 1.0;
    
    const html = `
        <div style="margin-bottom: 1.5rem; padding: 1.5rem; background: ${passed ? 'rgba(46, 204, 113, 0.1)' : 'rgba(239, 83, 80, 0.1)'}; border-radius: 6px; border: 2px solid ${passed ? 'var(--accent-success)' : 'var(--accent-danger)'};">
            <div style="font-size: 1.5rem; font-weight: 700; margin-bottom: 0.5rem;">
                ${passed ? '✅ PASSED' : '❌ FAILED'}
            </div>
            <div style="font-size: 1rem; color: var(--text-secondary);">
                Book Value Test: ${ratio ? ratio.toFixed(2) : 'N/A'}x
            </div>
        </div>
        <div style="display: flex; flex-direction: column; gap: 1rem;">
            <div class="dupont-factor">
                <span class="factor-name">Retained Earnings (5Y)</span>
                <span class="factor-value">${formatNumber(retainedEarnings)} Cr</span>
            </div>
            <div class="dupont-factor">
                <span class="factor-name">Book Value Change</span>
                <span class="factor-value">${formatNumber(bookValueChange)} Cr</span>
            </div>
            <div class="dupont-factor">
                <span class="factor-name">Value Created per Re. 1</span>
                <span class="factor-value">₹${ratio ? ratio.toFixed(2) : 'N/A'}</span>
            </div>
        </div>
        <div style="margin-top: 1.5rem; padding: 1rem; background: var(--bg-secondary); border-radius: 6px; font-size: 0.85rem; line-height: 1.6; color: var(--text-secondary);">
            <strong>Book Value Method:</strong> Price data not available, using conservative book value test.<br/>
            ${passed ? 
              '✅ Management is deploying retained earnings productively.' :
              '❌ Management destroying value. Should consider higher dividends or buybacks.'}
        </div>
    `;
    
    document.getElementById('buffettTest').innerHTML = html;
}

// ============================================================================
// REVISED: FLOAT DETECTION - EXPANDED ANALYSIS
// ============================================================================

function displayFLOATDetection() {
    const { annual } = companyData;
    const n = annual.years.length;
    
    if (n < 3) {
        document.getElementById('floatDetection').innerHTML = '<p style="color: var(--text-muted);">Requires 3+ years of data</p>';
        return;
    }
    
    const signals = [];
    let floatDetected = false;
    
    // Signal 1: Negative CCC
    const latestSales = annual.sales[n-1];
    const latestReceivables = annual.receivables[n-1] || 0;
    const latestInventory = annual.inventory[n-1] || 0;
    const cogs = (annual.rawMaterial[n-1] || 0) + (annual.otherMfg[n-1] || 0);
    
    const debtorDays = latestSales > 0 ? (latestReceivables / latestSales) * 365 : null;
    const inventoryDays = cogs > 0 ? (latestInventory / cogs) * 365 : null;
    const payableDays = 30; // Approximation
    const ccc = debtorDays + inventoryDays - payableDays;
    
    if (ccc < -30) {
        const floatAmount = (latestSales / 365) * Math.abs(ccc);
        floatDetected = true;
        signals.push({
            name: 'Negative Cash Conversion Cycle',
            value: `${Math.round(ccc)} days`,
            floatAmount: formatNumber(floatAmount) + ' Cr',
            description: `Collects cash ${Math.abs(Math.round(ccc))} days before paying suppliers`
        });
    }
    
    // Signal 2: High Other Liabilities
    const otherLiabilities = annual.otherLiabilities[n-1] || 0;
    const otherLiabPct = latestSales > 0 ? (otherLiabilities / latestSales) * 100 : 0;
    
    if (otherLiabPct > 20) {
        floatDetected = true;
        signals.push({
            name: 'High Other Liabilities',
            value: formatNumber(otherLiabilities) + ' Cr',
            floatAmount: `${otherLiabPct.toFixed(1)}% of sales`,
            description: 'Customer advances/unearned revenue - potential float'
        });
    }
    
    // Signal 3: Float Earnings
    const otherIncome = annual.otherIncome[n-1] || 0;
    const floatEarningsRatio = otherLiabilities > 0 ? (otherIncome / otherLiabilities) * 100 : null;
    
    if (floatEarningsRatio && floatEarningsRatio > 3) {
        const annualFloatBenefit = (otherIncome / (annual.netProfit[n-1] || 1)) * 100;
        signals.push({
            name: 'Float Earnings',
            value: formatNumber(otherIncome) + ' Cr',
            floatAmount: `${floatEarningsRatio.toFixed(1)}% return on liabilities`,
            description: `Earning ${annualFloatBenefit.toFixed(1)}% of net profit from customer funds`
        });
    }
    
    const html = `
        <div style="margin-bottom: 1.5rem; padding: 1.5rem; background: ${floatDetected ? 'rgba(46, 204, 113, 0.1)' : 'rgba(107, 117, 153, 0.1)'}; border-radius: 6px; border: 2px solid ${floatDetected ? 'var(--accent-primary)' : 'var(--border-color)'};">
            <div style="font-size: 1.5rem; font-weight: 700; margin-bottom: 0.5rem;">
                ${floatDetected ? '💰 FLOAT DETECTED' : '❌ No Float Detected'}
            </div>
            <div style="font-size: 0.9rem; color: var(--text-secondary);">
                ${floatDetected ? `${signals.length} signals identified` : 'Standard working capital model'}
            </div>
        </div>
        
        ${signals.length > 0 ? `
            <div style="display: flex; flex-direction: column; gap: 1rem; margin-bottom: 1.5rem;">
                ${signals.map(signal => `
                    <div style="padding: 1rem; background: var(--bg-secondary); border-radius: 6px;">
                        <div style="font-weight: 600; margin-bottom: 0.5rem;">${signal.name}</div>
                        <div style="display: flex; justify-content: space-between; margin-bottom: 0.25rem;">
                            <span style="font-size: 0.85rem; color: var(--text-muted);">Value:</span>
                            <span style="font-family: var(--font-mono);">${signal.value}</span>
                        </div>
                        <div style="display: flex; justify-content: space-between; margin-bottom: 0.5rem;">
                            <span style="font-size: 0.85rem; color: var(--text-muted);">Float Amount:</span>
                            <span style="font-family: var(--font-mono);">${signal.floatAmount}</span>
                        </div>
                        <div style="font-size: 0.85rem; color: var(--text-secondary);">${signal.description}</div>
                    </div>
                `).join('')}
            </div>
        ` : ''}
        
        <div style="padding: 1rem; background: var(--bg-secondary); border-radius: 6px; font-size: 0.85rem; line-height: 1.6; color: var(--text-secondary);">
            <strong>Float Business Models:</strong><br/>
            ${floatDetected ? 
                `This company benefits from float - customer money held before obligations are paid. Common in: Insurance, Asset Management, Retail chains (supplier credit), Exchanges, Subscription businesses.` :
                `No significant float detected. Standard operating model where company pays suppliers before collecting from customers.`}
        </div>
    `;
    
    document.getElementById('floatDetection').innerHTML = html;
}

// Continue with remaining functions (Financial Statements, Charts, etc.)...

// Financial Statements Display (UNCHANGED)
function displayFinancialStatements() {
    displayPLStatement();
    displayBalanceSheet();
    displayCashFlow();
}

function displayPLStatement() {
    const { annual } = companyData;
    const years = annual.years;
    
    const table = document.getElementById('plTable');
    
    const thead = table.querySelector('thead');
    thead.innerHTML = `
        <tr>
            <th>Particulars</th>
            ${years.map(y => `<th>FY${y}</th>`).join('')}
        </tr>
    `;
    
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
    
    const thead = table.querySelector('thead');
    thead.innerHTML = `
        <tr>
            <th>Particulars</th>
            ${years.map(y => `<th>FY${y}</th>`).join('')}
        </tr>
    `;
    
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
    
    const thead = table.querySelector('thead');
    thead.innerHTML = `
        <tr>
            <th>Particulars</th>
            ${years.map(y => `<th>FY${y}</th>`).join('')}
        </tr>
    `;
    
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

// Charts Display (Updated with green colors)
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
                    borderColor: '#2ecc71',
                    backgroundColor: 'rgba(46, 204, 113, 0.1)',
                    yAxisID: 'y',
                    tension: 0.4,
                    borderWidth: 3,
                    pointRadius: 4,
                    pointHoverRadius: 6
                },
                {
                    label: 'Net Profit',
                    data: annual.netProfit,
                    borderColor: '#27ae60',
                    backgroundColor: 'rgba(39, 174, 96, 0.1)',
                    yAxisID: 'y',
                    tension: 0.4,
                    borderWidth: 3,
                    pointRadius: 4,
                    pointHoverRadius: 6
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    labels: { 
                        color: '#ffffff',
                        font: { size: 12, weight: '600' }
                    }
                }
            },
            scales: {
                y: {
                    type: 'linear',
                    position: 'left',
                    ticks: { color: '#a0a0a0' },
                    grid: { color: '#333333' }
                },
                x: {
                    ticks: { color: '#a0a0a0' },
                    grid: { color: '#333333' }
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
                    borderColor: '#2ecc71',
                    backgroundColor: 'rgba(46, 204, 113, 0.1)',
                    tension: 0.4,
                    borderWidth: 3
                },
                {
                    label: 'Net Margin %',
                    data: netMargins,
                    borderColor: '#27ae60',
                    backgroundColor: 'rgba(39, 174, 96, 0.1)',
                    tension: 0.4,
                    borderWidth: 3
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    labels: { color: '#ffffff' }
                }
            },
            scales: {
                y: {
                    ticks: { 
                        color: '#a0a0a0',
                        callback: value => value + '%'
                    },
                    grid: { color: '#333333' }
                },
                x: {
                    ticks: { color: '#a0a0a0' },
                    grid: { color: '#333333' }
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
                    backgroundColor: 'rgba(46, 204, 113, 0.7)',
                    borderColor: '#2ecc71',
                    borderWidth: 1
                },
                {
                    label: 'ROCE %',
                    data: roces,
                    backgroundColor: 'rgba(39, 174, 96, 0.7)',
                    borderColor: '#27ae60',
                    borderWidth: 1
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    labels: { color: '#ffffff' }
                }
            },
            scales: {
                y: {
                    ticks: { 
                        color: '#a0a0a0',
                        callback: value => value + '%'
                    },
                    grid: { color: '#333333' }
                },
                x: {
                    ticks: { color: '#a0a0a0' },
                    grid: { color: '#333333' }
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
                    backgroundColor: 'rgba(46, 204, 113, 0.7)',
                    borderColor: '#2ecc71',
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
                    labels: { color: '#ffffff' }
                }
            },
            scales: {
                y: {
                    ticks: { color: '#a0a0a0' },
                    grid: { color: '#333333' }
                },
                x: {
                    ticks: { color: '#a0a0a0' },
                    grid: { color: '#333333' }
                }
            }
        }
    });
}

// Quarterly Analysis (UNCHANGED)
function displayQuarterlyAnalysis() {
    if (!workbookGlobal) return;
    
    const dataSheet = XLSX.utils.sheet_to_json(workbookGlobal.Sheets['Data Sheet'], { header: 1, defval: null });
    
    const qtrDates = (dataSheet[40] || []).slice(4).filter(d => d);
    if (qtrDates.length === 0) {
        document.getElementById('quarterlyTable').innerHTML = '<p style="padding: 2rem; text-align: center; color: var(--text-muted);">No quarterly data available</p>';
        return;
    }
    
    const quarters = qtrDates.map(d => {
        let date;
        if (typeof d === 'number') {
            const excelEpoch = new Date(1900, 0, 1);
            date = new Date(excelEpoch.getTime() + (d - 2) * 86400000);
        } else if (d instanceof Date) {
            date = d;
        } else {
            date = new Date(d);
        }
        const month = date.getMonth();
        const year = date.getFullYear();
        const qtr = Math.floor(month / 3) + 1;
        return `Q${qtr} FY${year}`;
    });
    
    const qtrSales = extractRow(dataSheet, 41);
    const qtrExpenses = extractRow(dataSheet, 42);
    const qtrProfit = extractRow(dataSheet, 48);
    const qtrOpProfit = extractRow(dataSheet, 49);
    
    const table = document.getElementById('quarterlyTable');
    const thead = table.querySelector('thead');
    thead.innerHTML = `
        <tr>
            <th>Quarter</th>
            ${quarters.map(q => `<th>${q}</th>`).join('')}
        </tr>
    `;
    
    const tbody = table.querySelector('tbody');
    tbody.innerHTML = `
        <tr>
            <td>Sales</td>
            ${qtrSales.slice(0, quarters.length).map(v => `<td>${formatNumber(v)}</td>`).join('')}
        </tr>
        <tr>
            <td>Operating Profit</td>
            ${qtrOpProfit.slice(0, quarters.length).map(v => `<td>${formatNumber(v)}</td>`).join('')}
        </tr>
        <tr class="total-row">
            <td>Net Profit</td>
            ${qtrProfit.slice(0, quarters.length).map(v => `<td>${formatNumber(v)}</td>`).join('')}
        </tr>
    `;
    
    const qoqSalesGrowth = [];
    for (let i = 1; i < qtrSales.length && i < quarters.length; i++) {
        if (qtrSales[i-1] && qtrSales[i]) {
            qoqSalesGrowth.push(((qtrSales[i] - qtrSales[i-1]) / qtrSales[i-1]) * 100);
        } else {
            qoqSalesGrowth.push(null);
        }
    }
    
    document.getElementById('qoqGrowth').innerHTML = `
        <div class="dupont-factor">
            <span class="factor-name">Latest QoQ Sales Growth</span>
            <span class="factor-value">${formatPercent(qoqSalesGrowth[qoqSalesGrowth.length - 1])}</span>
        </div>
        <div class="dupont-factor">
            <span class="factor-name">Avg QoQ Growth (4Q)</span>
            <span class="factor-value">${formatPercent(qoqSalesGrowth.slice(-4).reduce((a,b) => a+b, 0) / 4)}</span>
        </div>
    `;
    
    if (qtrSales.length >= 5) {
        const yoySales = ((qtrSales[qtrSales.length-1] - qtrSales[qtrSales.length-5]) / qtrSales[qtrSales.length-5]) * 100;
        const yoyProfit = ((qtrProfit[qtrProfit.length-1] - qtrProfit[qtrProfit.length-5]) / qtrProfit[qtrProfit.length-5]) * 100;
        
        document.getElementById('yoyGrowth').innerHTML = `
            <div class="dupont-factor">
                <span class="factor-name">YoY Sales Growth</span>
                <span class="factor-value">${formatPercent(yoySales)}</span>
            </div>
            <div class="dupont-factor">
                <span class="factor-name">YoY Profit Growth</span>
                <span class="factor-value">${formatPercent(yoyProfit)}</span>
            </div>
        `;
    }
    
    const ctx = document.getElementById('quarterlyChart');
    new Chart(ctx, {
        type: 'line',
        data: {
            labels: quarters,
            datasets: [
                {
                    label: 'Quarterly Sales',
                    data: qtrSales.slice(0, quarters.length),
                    borderColor: '#2ecc71',
                    backgroundColor: 'rgba(46, 204, 113, 0.1)',
                    tension: 0.3
                },
                {
                    label: 'Quarterly Profit',
                    data: qtrProfit.slice(0, quarters.length),
                    borderColor: '#27ae60',
                    backgroundColor: 'rgba(39, 174, 96, 0.1)',
                    tension: 0.3
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    labels: { color: '#ffffff' }
                }
            },
            scales: {
                y: {
                    ticks: { color: '#a0a0a0' },
                    grid: { color: '#333333' }
                },
                x: {
                    ticks: { color: '#a0a0a0' },
                    grid: { color: '#333333' }
                }
            }
        }
    });
}

// Comparison Mode (UNCHANGED)
async function handleComparisonUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data);
        const parsedData = parseExcelData(workbook);
        
        comparisonCompanies.push(parsedData);
        document.getElementById('compareTab').style.display = 'block';
        displayComparison();
        
        alert(`Added ${parsedData.meta.name} for comparison!`);
    } catch (error) {
        console.error('Error adding company for comparison:', error);
        alert('Error parsing file for comparison');
    }
}

function clearComparison() {
    comparisonCompanies = [];
    document.getElementById('compareTab').style.display = 'none';
    document.getElementById('comparisonContent').innerHTML = '<p style="text-align: center; color: var(--text-muted);">No companies added for comparison yet.</p>';
}

function displayComparison() {
    if (comparisonCompanies.length === 0) {
        document.getElementById('comparisonContent').innerHTML = '<p style="text-align: center; color: var(--text-muted);">No companies added for comparison yet.</p>';
        return;
    }
    
    const companies = [companyData, ...comparisonCompanies];
    
    const html = companies.map(company => {
        const n = company.annual.years.length;
        const salesCAGR5 = calculateCAGR(company.annual.sales, Math.min(5, n-1));
        const profitCAGR5 = calculateCAGR(company.annual.netProfit, Math.min(5, n-1));
        
        const latestEquity = (company.annual.equity[n-1] || 0) + (company.annual.reserves[n-1] || 0);
        const prevEquity = (company.annual.equity[n-2] || 0) + (company.annual.reserves[n-2] || 0);
        const avgEquity = (latestEquity + prevEquity) / 2;
        const roe = avgEquity > 0 ? ((company.annual.netProfit[n-1] || 0) / avgEquity) * 100 : null;
        
        const latestSales = company.annual.sales[n-1];
        const latestProfit = company.annual.netProfit[n-1];
        const npm = latestSales > 0 ? (latestProfit / latestSales) * 100 : null;
        
        const debtToEquity = latestEquity > 0 ? (company.annual.borrowings[n-1] || 0) / latestEquity : null;
        
        return `
            <div class="company-comparison-card">
                <h3>${company.meta.name}</h3>
                <div class="comparison-metrics">
                    <div class="comparison-metric">
                        <span class="comparison-metric-label">Market Cap</span>
                        <span class="comparison-metric-value">${formatLargeNumber(company.meta.marketCap)} Cr</span>
                    </div>
                    <div class="comparison-metric">
                        <span class="comparison-metric-label">Sales CAGR (5Y)</span>
                        <span class="comparison-metric-value">${formatPercent(salesCAGR5)}</span>
                    </div>
                    <div class="comparison-metric">
                        <span class="comparison-metric-label">Profit CAGR (5Y)</span>
                        <span class="comparison-metric-value">${formatPercent(profitCAGR5)}</span>
                    </div>
                    <div class="comparison-metric">
                        <span class="comparison-metric-label">ROE</span>
                        <span class="comparison-metric-value">${formatPercent(roe)}</span>
                    </div>
                    <div class="comparison-metric">
                        <span class="comparison-metric-label">Net Margin</span>
                        <span class="comparison-metric-value">${formatPercent(npm)}</span>
                    </div>
                    <div class="comparison-metric">
                        <span class="comparison-metric-label">Debt/Equity</span>
                        <span class="comparison-metric-value">${debtToEquity ? debtToEquity.toFixed(2) + 'x' : 'N/A'}</span>
                    </div>
                    <div class="comparison-metric">
                        <span class="comparison-metric-label">Quality Score</span>
                        <span class="comparison-metric-value">${calculateQualityScore(company).total}/100</span>
                    </div>
                </div>
            </div>
        `;
    }).join('');
    
    document.getElementById('comparisonContent').innerHTML = html;
}

// Tab Switching
function switchTab(tabName) {
    document.querySelectorAll('.tab').forEach(tab => {
        tab.classList.remove('active');
        if (tab.dataset.tab === tabName) {
            tab.classList.add('active');
        }
    });
    
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
