/**
 * Generates an executive-grade Financial Dashboard featuring:
 * - Top KPI Scorecards
 * - Executive Financial Summary Table
 * - Compact Core Expense Breakdown Table
 * - Visual Analytics (Net Worth Trend, Cash Flow Bar Chart)
 * - Clean 3-Month Expense Share Donut Charts (Reference Data in Off-Screen Columns Z onwards)
 *
 * REPLACES the old "dashboard" file entirely. Needs Settings and Setup.gs.
 *
 * Changes from the old version:
 *  - Expense groups come from the "Dashboard Group" column on the Chart of
 *    Accounts (and the group list in the Settings area), instead of name
 *    patterns hard-coded in this file. Accounts with no group fall back to the
 *    old name rules, so existing workbooks look the same until edited.
 *  - The currency symbol comes from Settings; no more hard-coded "$".
 *  - Expense categories with a net NEGATIVE total (for example after a
 *    reimbursement) are no longer dropped, and accounts that share a name are
 *    added together instead of overwriting each other.
 *  - The "no Financials sheets" message names the right menu item.
 */

const DASHBOARD_MAX_GROUPS_ = 12;   // keeps the breakdown table clear of the reference columns (Z onwards)

function generateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cur = getCurrencySymbol();
  const moneyFmt = dashboardMoneyFormat_(cur);
  const axisFmt = dashboardAxisFormat_(cur);

  // 1. Scan for all "Financials - [Month Year]" sheets
  const sheets = ss.getSheets();
  const financialData = [];

  const monthMap = {
    "january": 0, "february": 1, "march": 2, "april": 3, "may": 4, "june": 5,
    "july": 6, "august": 7, "september": 8, "october": 9, "november": 10, "december": 11
  };

  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name.startsWith("Financials - ")) {
      const parts = name.replace("Financials - ", "").trim().split(" ");
      if (parts.length >= 2) {
        const monthStr = parts[0];
        const yearNum = Number(parts[1]);
        const mIdx = monthMap[monthStr.toLowerCase()];

        if (mIdx !== undefined && !isNaN(yearNum)) {
          const sortDate = new Date(yearNum, mIdx, 1);
          const data = extractFinancialMetrics(sheet);

          financialData.push({
            sheetName: name,
            period: `${monthStr} ${yearNum}`,
            date: sortDate,
            revenue: data.revenue,
            expenses: data.expenses,
            netIncome: data.netIncome,
            assets: data.assets,
            liabilities: data.liabilities,
            equity: data.equity,
            netWorth: data.assets - data.liabilities,
            expenseDetails: data.expenseDetails
          });
        }
      }
    }
  });

  if (financialData.length === 0) {
    SpreadsheetApp.getUi().alert("No 'Financials - [Month Year]' sheets found.\n\nPlease run 'Generate Monthly Reports' on a monthly ledger tab first.");
    return;
  }

  // Sort chronologically
  financialData.sort((a, b) => a.date - b.date);

  // Expense groups: from the Chart of Accounts / Settings, with "Other" always last
  const groupSetup = dashboardGroupSetup_();
  const groups = groupSetup.groups;
  financialData.forEach(d => { d.buckets = bucketExpenses_(d.expenseDetails, groupSetup); });

  // Reset Dashboard Sheet
  let dashboard = ss.getSheetByName("Dashboard") || ss.insertSheet("Dashboard");

  // Unhide any hidden rows before clearing to ensure clean state
  const maxRows = dashboard.getMaxRows();
  if (maxRows > 0) {
    dashboard.showRows(1, maxRows);
  }

  dashboard.clear().clearNotes().clearFormats();
  dashboard.getCharts().forEach(chart => dashboard.removeChart(chart));

  // --- 1. TITLE BANNER ---
  dashboard.getRange("A1:H1").merge()
    .setValue("FINANCIAL DASHBOARD")
    .setFontWeight("bold")
    .setFontSize(13)
    .setHorizontalAlignment("center")
    .setBackground("#1e293b") // Slate dark navy
    .setFontColor("#ffffff");
  dashboard.setRowHeight(1, 38);

  // --- 2. TOP KPI SCORECARDS (Rows 3 to 5) ---
  const latestNetWorth = financialData[financialData.length - 1].netWorth;
  const avgRev = financialData.reduce((s, d) => s + d.revenue, 0) / financialData.length;
  const avgExp = financialData.reduce((s, d) => s + d.expenses, 0) / financialData.length;
  const avgSavingsRate = avgRev > 0 ? (avgRev - avgExp) / avgRev : 0;

  createKpiCard(dashboard, "A3:B5", "CURRENT NET WORTH", latestNetWorth, "money", cur);
  createKpiCard(dashboard, "C3:D5", "AVG MONTHLY REVENUE", avgRev, "money", cur);
  createKpiCard(dashboard, "E3:F5", "AVG MONTHLY EXPENSES", avgExp, "money", cur);
  createKpiCard(dashboard, "G3:H5", "NET SAVINGS RATE", avgSavingsRate, "0.0%", cur);

  // --- 3. TABLE 1: EXECUTIVE SUMMARY TABLE ---
  const startRow1 = 8;
  dashboard.getRange(startRow1 - 1, 1).setValue("1. FINANCIAL SUMMARY BY MONTH")
    .setFontWeight("bold").setFontSize(11).setFontColor("#1e293b");

  const headers = [["Month / Year", "Total Revenue", "Total Expenses", "Net Income", "Total Assets", "Total Liabilities", "Net Worth"]];
  dashboard.getRange(startRow1, 1, 1, 7).setValues(headers)
    .setFontWeight("bold")
    .setBackground("#334155")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  const tableRows = financialData.map(d => [
    d.period,
    d.revenue,
    d.expenses,
    d.netIncome,
    d.assets,
    d.liabilities,
    d.netWorth
  ]);

  const numRows1 = tableRows.length;
  dashboard.getRange(startRow1 + 1, 1, numRows1, 7).setValues(tableRows);

  // Alternating Row Shading
  for (let i = 0; i < numRows1; i++) {
    const rowNum = startRow1 + 1 + i;
    const bg = i % 2 === 0 ? "#ffffff" : "#f8fafc";
    dashboard.getRange(rowNum, 1, 1, 7).setBackground(bg);
  }

  dashboard.getRange(startRow1 + 1, 1, numRows1, 1).setHorizontalAlignment("center");
  dashboard.getRange(startRow1 + 1, 2, numRows1, 6).setNumberFormat(moneyFmt).setHorizontalAlignment("right");
  dashboard.getRange(startRow1 + 1, 7, numRows1, 1).setFontWeight("bold").setFontColor("#1e3a8a");

  // Averages Row
  const endRow1 = startRow1 + 1 + numRows1;
  dashboard.getRange(endRow1, 1).setValue("Monthly Average").setFontWeight("bold").setHorizontalAlignment("center");
  for (let c = 2; c <= 7; c++) {
    const colLetter = getColumnLetter(c);
    dashboard.getRange(endRow1, c).setFormula(`=AVERAGE(${colLetter}${startRow1 + 1}:${colLetter}${endRow1 - 1})`);
  }
  dashboard.getRange(endRow1, 1, 1, 7)
    .setFontWeight("bold")
    .setBackground("#f1f5f9")
    .setBorder(true, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.DOUBLE)
    .setNumberFormat(moneyFmt);

  // --- 4. TABLE 2: EXPENSE BREAKDOWN BY GROUP ---
  const startRow2 = endRow1 + 3;
  dashboard.getRange(startRow2 - 1, 1).setValue("2. EXPENSE BREAKDOWN BY CATEGORY")
    .setFontWeight("bold").setFontSize(11).setFontColor("#1e293b");

  const catHeaders = ["Month / Year", ...groups, "Total Expenses"];
  const numCols2 = catHeaders.length;

  dashboard.getRange(startRow2, 1, 1, numCols2).setValues([catHeaders])
    .setFontWeight("bold")
    .setBackground("#334155")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  const expTableRows = financialData.map(d => [
    d.period,
    ...groups.map(g => d.buckets[g]),
    d.expenses
  ]);

  const numRows2 = expTableRows.length;
  dashboard.getRange(startRow2 + 1, 1, numRows2, numCols2).setValues(expTableRows);

  for (let i = 0; i < numRows2; i++) {
    const rowNum = startRow2 + 1 + i;
    const bg = i % 2 === 0 ? "#ffffff" : "#f8fafc";
    dashboard.getRange(rowNum, 1, 1, numCols2).setBackground(bg);
  }

  dashboard.getRange(startRow2 + 1, 1, numRows2, 1).setHorizontalAlignment("center");
  dashboard.getRange(startRow2 + 1, 2, numRows2, numCols2 - 1).setNumberFormat(moneyFmt).setHorizontalAlignment("right");

  // Category Averages Row
  const endRow2 = startRow2 + 1 + numRows2;
  dashboard.getRange(endRow2, 1).setValue("Monthly Average").setFontWeight("bold").setHorizontalAlignment("center");

  for (let c = 2; c <= numCols2; c++) {
    let colStr = getColumnLetter(c);
    dashboard.getRange(endRow2, c).setFormula(`=AVERAGE(${colStr}${startRow2 + 1}:${colStr}${endRow2 - 1})`);
  }

  dashboard.getRange(endRow2, 1, 1, numCols2)
    .setFontWeight("bold")
    .setBackground("#f1f5f9")
    .setBorder(true, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.DOUBLE)
    .setNumberFormat(moneyFmt);

  // --- 5. CLEAN COLUMN WIDTHS ---
  const widthCols = Math.max(8, numCols2);
  dashboard.setColumnWidth(1, 130); // Period
  for (let c = 2; c <= widthCols; c++) {
    dashboard.setColumnWidth(c, 125);
  }
  dashboard.setColumnWidth(numCols2, 130); // Total Expenses

  // --- 6. CHARTS & VISUAL ANALYTICS SECTION ---
  const chartStartRow = endRow2 + 3;
  dashboard.getRange(chartStartRow - 1, 1).setValue("3. VISUAL ANALYTICS & TRENDS")
    .setFontWeight("bold").setFontSize(11).setFontColor("#1e293b");

  // Chart 1: Net Worth Growth Line Chart
  const netWorthChart = dashboard.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(dashboard.getRange(startRow1, 1, numRows1 + 1, 1)) // X: Period
    .addRange(dashboard.getRange(startRow1, 7, numRows1 + 1, 1)) // Y: Net Worth
    .setPosition(chartStartRow, 1, 0, 0)
    .setOption('title', 'Net Worth Growth Trend')
    .setOption('colors', ['#2563eb'])
    .setOption('pointSize', 6)
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { title: 'Month / Year', textStyle: { fontSize: 10 } })
    .setOption('vAxis', { title: `Net Worth (${cur})`, format: axisFmt, textStyle: { fontSize: 10 } })
    .setOption('width', 490)
    .setOption('height', 290)
    .build();
  dashboard.insertChart(netWorthChart);

  // Chart 2: Cash Flow (Revenue vs Expenses) Bar Chart
  const cashFlowChart = dashboard.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dashboard.getRange(startRow1, 1, numRows1 + 1, 1)) // X: Period
    .addRange(dashboard.getRange(startRow1, 2, numRows1 + 1, 1)) // Revenue
    .addRange(dashboard.getRange(startRow1, 3, numRows1 + 1, 1)) // Expenses
    .setPosition(chartStartRow, 5, 0, 0)
    .setOption('title', 'Monthly Cash Flow (Revenue vs. Expenses)')
    .setOption('colors', ['#059669', '#dc2626'])
    .setOption('legend', { position: 'top' })
    .setOption('hAxis', { title: 'Month / Year', textStyle: { fontSize: 10 } })
    .setOption('vAxis', { title: `Amount (${cur})`, format: axisFmt, textStyle: { fontSize: 10 } })
    .setOption('width', 490)
    .setOption('height', 290)
    .build();
  dashboard.insertChart(cashFlowChart);

  // --- 7. RECENT 3-MONTH EXPENSE SHARE (OFF-SCREEN DATA IN COLUMNS Z+ + 3 DONUT CHARTS) ---
  const recentMonths = financialData.slice(-3); // Extract last 3 months
  const pieChartRow = chartStartRow + 16;

  dashboard.getRange(pieChartRow - 1, 1).setValue("4. RECENT 3-MONTH EXPENSE BREAKDOWN")
    .setFontWeight("bold").setFontSize(11).setFontColor("#1e293b");

  // Write reference data off-screen into Columns Z (26) onwards so no rows are hidden and no charts break
  const refStartCol = 26; // Column Z
  const recentHeaders = ["Category", ...recentMonths.map(m => m.period)];
  const numRecentCols = recentHeaders.length;

  dashboard.getRange(1, refStartCol, 1, numRecentCols).setValues([recentHeaders]);

  // A donut cannot show negative slices, so a net-negative group is shown as zero here
  const recentMatrix = groups.map(g => [g, ...recentMonths.map(m => Math.max(0, m.buckets[g]))]);
  dashboard.getRange(2, refStartCol, groups.length, numRecentCols).setValues(recentMatrix);

  // Generate 1 Donut Chart per recent month directly under Section 4 header
  recentMonths.forEach((m, idx) => {
    const monthColIdx = refStartCol + 1 + idx; // Col AA (27), Col AB (28), Col AC (29)
    const catRange = dashboard.getRange(1, refStartCol, groups.length + 1, 1);    // Categories in Col Z
    const monthRange = dashboard.getRange(1, monthColIdx, groups.length + 1, 1);  // Specific month values

    const startColForChart = 1 + (idx * 3); // Position charts across columns A (1), D (4), G (7)

    const pieChart = dashboard.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(catRange)
      .addRange(monthRange)
      .setPosition(pieChartRow, startColForChart, 0, 0)
      .setOption('title', `${m.period} Expense Share`)
      .setOption('pieHole', 0.35)
      .setOption('pieSliceText', 'percentage')
      .setOption('legend', { position: 'bottom', textStyle: { fontSize: 9 } })
      .setOption('width', 320)
      .setOption('height', 280)
      .build();

    dashboard.insertChart(pieChart);
  });

  // Focus Dashboard
  ss.setActiveSheet(dashboard);
  ss.toast("Executive Dashboard updated cleanly!", "Success", 4);
}

// ---------------------------------------------------------------------------
// EXPENSE GROUPS
// ---------------------------------------------------------------------------

/**
 * Returns { groups: [...], nameToGroup: {accountName: group} }.
 * groups comes from the Settings area, capped, with "Other" always last.
 */
function dashboardGroupSetup_() {
  let chart = {};
  try { chart = readChartOfAccounts(); } catch (e) { chart = {}; }

  const configured = getDashboardGroups().filter(g => g !== 'Other').slice(0, DASHBOARD_MAX_GROUPS_ - 1);
  const groups = configured.concat(['Other']);

  const nameToGroup = {};
  Object.keys(chart).forEach(code => {
    const a = chart[code];
    if (a.type !== 'Expense') return;
    const name = String(a.name).trim();
    // Blank group: fall back to the old name rules (keeps pre-upgrade workbooks unchanged)
    nameToGroup[name] = a.dashboardGroup || legacyDashboardGroup_(name, groups);
  });

  return { groups: groups, nameToGroup: nameToGroup };
}

/** Adds each expense category's amount to its group. Unknown or removed groups go to "Other". */
function bucketExpenses_(details, setup) {
  const buckets = {};
  setup.groups.forEach(g => { buckets[g] = 0; });

  Object.keys(details).forEach(label => {
    const name = String(label).trim();
    let group = setup.nameToGroup[name] || legacyDashboardGroup_(name, setup.groups);
    if (!(group in buckets)) group = 'Other';
    buckets[group] += details[label];
  });
  return buckets;
}

// ---------------------------------------------------------------------------
// DATA EXTRACTION
// ---------------------------------------------------------------------------

/**
 * Extracts high-level financial metrics and individual expense category line items from a Financials sheet.
 */
function extractFinancialMetrics(sheet) {
  const lastRow = sheet.getLastRow();
  let revenue = 0, expenses = 0, netIncome = 0;
  let assets = 0, liabilities = 0, equity = 0;
  let expenseDetails = {};

  if (lastRow < 2) {
    return { revenue, expenses, netIncome, assets, liabilities, equity, expenseDetails };
  }

  // Parse Profit or Loss (Cols K & L, i.e. 11 & 12)
  const plData = sheet.getRange(1, 11, lastRow, 2).getValues();
  let inExpensesSection = false;

  plData.forEach(row => {
    const label = String(row[0]).trim();
    const upperLabel = label.toUpperCase();
    const val = Number(row[1]) || 0;

    if (upperLabel === "TOTAL REVENUE") {
      revenue = val;
    } else if (upperLabel === "EXPENSES") {
      inExpensesSection = true;
    } else if (upperLabel === "TOTAL EXPENSES") {
      expenses = val;
      inExpensesSection = false;
    } else if (upperLabel.startsWith("NET PROFIT")) {
      netIncome = val;
    } else if (inExpensesSection && label !== "" && Math.abs(val) >= 0.005) {
      // Keep negative categories (reimbursements) and add together accounts that share a name
      expenseDetails[label] = (expenseDetails[label] || 0) + val;
    }
  });

  // Parse Balance Sheet (Cols N & O, i.e. 14 & 15)
  const bsData = sheet.getRange(1, 14, lastRow, 2).getValues();
  bsData.forEach(row => {
    const upperLabel = String(row[0]).trim().toUpperCase();
    const val = Number(row[1]) || 0;

    if (upperLabel === "TOTAL ASSETS") assets = val;
    else if (upperLabel === "TOTAL LIABILITIES") liabilities = val;
    else if (upperLabel === "TOTAL EQUITY") equity = val;
  });

  return { revenue, expenses, netIncome, assets, liabilities, equity, expenseDetails };
}

// ---------------------------------------------------------------------------
// FORMATTING HELPERS
// ---------------------------------------------------------------------------

/** Sheets number format with the currency symbol as literal text, e.g. "J$"#,##0.00 */
function dashboardMoneyFormat_(cur) {
  return '"' + String(cur).replace(/"/g, '') + '"#,##0.00';
}

/** Chart axis format (ICU pattern); the symbol is quoted so letters like "J" are literal. */
function dashboardAxisFormat_(cur) {
  return "'" + String(cur).replace(/'/g, "''") + "'#,##0";
}

/** Plain text money with the minus sign in front of the symbol: -$1,234.50 */
function dashboardMoney_(n, cur) {
  const s = Math.abs(n).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (n < 0 ? '-' : '') + cur + s;
}

/**
 * Helper to build top KPI Scorecards.
 * numberFormat: anything containing "%" is shown as a percentage, everything else as money.
 */
function createKpiCard(sheet, cellRange, title, value, numberFormat, cur) {
  const range = sheet.getRange(cellRange);
  range.merge()
    .setBackground("#f8fafc")
    .setBorder(true, true, true, true, null, null, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

  const topCell = range.getCell(1, 1);
  const formattedVal = String(numberFormat).includes("%")
    ? (value * 100).toFixed(1) + "%"
    : dashboardMoney_(value, cur || "$");

  topCell.setValue(`${title}\n${formattedVal}`)
    .setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
}

/**
 * Converts column index to letter string (e.g. 1 -> A, 27 -> AA)
 */
function getColumnLetter(col) {
  let colStr = "";
  let temp = col;
  while (temp > 0) {
    let rem = (temp - 1) % 26;
    colStr = String.fromCharCode(65 + rem) + colStr;
    temp = Math.floor((temp - rem) / 26);
  }
  return colStr;
}
