/**
 * ============================================================================
 * BETTER BUDGET - FINANCIAL REPORTS GENERATION ENGINE
 * ============================================================================
 * Generates five consolidated financial statements on a single sheet:
 * 1. T-Accounts (Cols A–D)
 * 2. Trial Balance (Cols F–I)
 * 3. Profit or Loss Statement (Cols K–L)
 * 4. Balance Sheet (Cols N–O)
 * 5. Indirect Statement of Cash Flows (Cols Q–R)
 */

function generateConsolidatedReports() {
  REPORT_FMT_CACHE_ = null;   // re-read the currency setting on every run
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  // --------------------------------------------------------------------------
  // 1. LEDGER VALIDATION GUARD
  // --------------------------------------------------------------------------
  const validation = validateLedgerSheet(sourceSheet);
  if (!validation.valid) {
    ui.alert(
      "⚠️ Incompatible Sheet Selected",
      `Cannot generate financial reports from the active sheet.\n\nReason: ${validation.reason}\n\nPlease switch to a monthly General Ledger tab and try again.`,
      ui.ButtonSet.OK
    );
    return;
  }

  // --------------------------------------------------------------------------
  // 2. SETUP MONTH / YEAR & SUFFIX
  // --------------------------------------------------------------------------
  let month = String(sourceSheet.getRange("B3").getValue()).trim();
  let year = String(sourceSheet.getRange("B4").getValue()).trim();

  // Fallback if title block B3/B4 is empty or shifted
  if (!month || !year) {
    const sheetParts = sourceSheet.getName().trim().split(" ");
    if (sheetParts.length >= 2) {
      month = sheetParts[0];
      year = sheetParts[1];
    }
  }

  const suffix = `${month} ${year}`;
  const reportSheetName = `Financials - ${suffix}`;

  // --------------------------------------------------------------------------
  // 3. FETCH CHART OF ACCOUNTS (Code, Name, Type)
  // --------------------------------------------------------------------------
  // Reads the Chart of Accounts by column header (see Settings and Setup.gs)
  let accountInfo = {};
  try { accountInfo = readChartOfAccounts(); } catch (e) { accountInfo = {}; }

  // --------------------------------------------------------------------------
  // 4. FETCH TRANSACTIONS (Dynamic Header Detection for Row 8 / Shifted Layouts)
  // --------------------------------------------------------------------------
  const rawSheetData = sourceSheet.getDataRange().getValues();
  const headerRowIndex = findHeaderRow(rawSheetData);
  const headerRow = headerRowIndex !== -1 ? headerRowIndex + 1 : 8;

  const lastRow = sourceSheet.getLastRow();
  let transactions = [];
  if (lastRow > headerRow) {
    const rawData = sourceSheet.getRange(headerRow + 1, 1, (lastRow - headerRow), 5).getValues();
    transactions = rawData.filter(row => row[1] !== "" && row[1] !== null).map(row => {
      row[1] = String(row[1]).trim();
      return row;
    });
  }

  let accountCodes = new Set(transactions.map(row => row[1]));

  // --------------------------------------------------------------------------
  // 5. FETCH PREVIOUS MONTH'S BALANCES & ROLL-FORWARD
  // --------------------------------------------------------------------------
  let prevBalances = {};
  const currentDate = new Date(`${month} 1, ${year}`);
  currentDate.setMonth(currentDate.getMonth() - 1);
  const prevMonthStr = currentDate.toLocaleString('default', { month: 'long' });
  const prevYearStr = currentDate.getFullYear();
  const prevSheetName = `Financials - ${prevMonthStr} ${prevYearStr}`;

  const prevSheet = ss.getSheetByName(prevSheetName);

  if (prevSheet) {
    const lastRowPrev = prevSheet.getLastRow();
    // Trial Balance is in columns F-I (6-9), data starts on row 3
    if (lastRowPrev > 2) {
      const tbData = prevSheet.getRange(3, 6, lastRowPrev - 2, 4).getValues();

      let prevNetIncomeBal = 0;

      tbData.forEach(row => {
        const code = String(row[0]).trim();
        const drAmount = Number(row[2]) || 0;
        const crAmount = Number(row[3]) || 0;

        if (code && accountInfo[code]) {
          const type = accountInfo[code].type;
          let netBal = drAmount - crAmount;

          // 1. Bring forward permanent accounts
          if (["Asset", "Assets", "Liability", "Liabilities", "Equity"].includes(type)) {
            if (Math.abs(netBal) > 0.01) {
              prevBalances[code] = {
                amount: Math.abs(netBal),
                type: netBal > 0 ? "Dr" : "Cr"
              };
              accountCodes.add(code);
            }
          }
          // 2. Catch temporary accounts to calculate Net Profit/Loss
          else if (["Revenue", "Expense", "Expenses"].includes(type)) {
            prevNetIncomeBal += netBal;
          }
        }
      });

      // 3. Roll previous Net Profit/Loss into Retained Earnings
      if (Math.abs(prevNetIncomeBal) > 0.01) {
        const retainedEarningsCode = getRetainedEarningsCode();

        let existingBal = 0;
        if (prevBalances[retainedEarningsCode]) {
          existingBal = prevBalances[retainedEarningsCode].type === "Dr"
            ? prevBalances[retainedEarningsCode].amount
            : -prevBalances[retainedEarningsCode].amount;
        }

        let newTotalBal = existingBal + prevNetIncomeBal;

        if (Math.abs(newTotalBal) > 0.01) {
          prevBalances[retainedEarningsCode] = {
            amount: Math.abs(newTotalBal),
            type: newTotalBal > 0 ? "Dr" : "Cr"
          };
          accountCodes.add(retainedEarningsCode);
        }
      }
    }
  }

  // --------------------------------------------------------------------------
  // 6. PREPARE / CLEAR CONSOLIDATED SHEET
  // --------------------------------------------------------------------------
  let reportSheet = ss.getSheetByName(reportSheetName) || ss.insertSheet(reportSheetName);
  reportSheet.clear().clearNotes().clearFormats();

  // --------------------------------------------------------------------------
  // 7. RENDER T-ACCOUNTS (Columns A-D)
  // --------------------------------------------------------------------------
  let currentRow = 1;
  const sortedCodes = Array.from(accountCodes).sort();
  const finalBalances = {};

  sortedCodes.forEach((code) => {
    const info = accountInfo[code] || { name: "Unknown", type: "Unknown" };
    const accountData = transactions.filter(row => row[1] === code);
    const bForward = prevBalances[code] || { amount: 0, type: "" };

    // Header for individual T-Account
    reportSheet.getRange(currentRow, 1, 1, 4).merge().setValue(`${info.name} - ${code}`)
      .setFontWeight("bold").setHorizontalAlignment("center").setBackground("#eeeeee").setBorder(true,true,true,true,null,null);

    reportSheet.getRange(currentRow + 1, 1, 1, 4).setValues([["Date", "Debit", "Date", "Credit"]])
      .setFontWeight("bold").setBorder(null,null,true,null,null,null,"black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    let debits = accountData.filter(row => row[3] > 0).map(r => ({ d: r[0], n: r[2], a: r[3] }));
    let credits = accountData.filter(row => row[4] > 0).map(r => ({ d: r[0], n: r[2], a: r[4] }));

    // 1. Insert Previous Month's Balance B/F at top
    if (bForward.amount > 0.01) {
      const bf = { d: "Balance B/F", n: "Opening balance", a: bForward.amount };
      bForward.type === "Dr" ? debits.unshift(bf) : credits.unshift(bf);
    }

    // 2. Calculate provisional totals to find closing balance
    let provDrTotal = debits.reduce((sum, item) => sum + Number(item.a), 0);
    let provCrTotal = credits.reduce((sum, item) => sum + Number(item.a), 0);
    let netBal = provDrTotal - provCrTotal;

    // Save actual mathematical balance
    finalBalances[code] = netBal;

    // 3. Inject Balance C/F into lighter side to force account to balance
    if (netBal > 0.01) {
      credits.push({ d: "", n: "Closing balance", a: netBal, isCf: true });
    } else if (netBal < -0.01) {
      debits.push({ d: "", n: "Closing balance", a: Math.abs(netBal), isCf: true });
    }

    // 4. Render rows
    const rowCount = Math.max(debits.length, credits.length, 1);
    let drTotal = 0, crTotal = 0;

    for (let i = 0; i < rowCount; i++) {
      const rIdx = currentRow + 2 + i;
      if (debits[i]) {
        let label = debits[i].isCf ? "Balance C/F" : debits[i].d;
        reportSheet.getRange(rIdx, 1).setValue(label).setNote(debits[i].n);
        reportSheet.getRange(rIdx, 2).setValue(debits[i].a);
        drTotal += Number(debits[i].a);
      }
      if (credits[i]) {
        let label = credits[i].isCf ? "Balance C/F" : credits[i].d;
        reportSheet.getRange(rIdx, 3).setValue(label).setNote(credits[i].n);
        reportSheet.getRange(rIdx, 4).setValue(credits[i].a);
        crTotal += Number(credits[i].a);
      }
    }

    // 5. Draw Borders and Matching Totals
    reportSheet.getRange(currentRow + 1, 2, rowCount + 1, 1).setBorder(null,null,null,true,null,null,"black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    const foot = currentRow + 2 + rowCount;

    reportSheet.getRange(foot, 2).setValue(drTotal).setFontWeight("bold")
      .setBorder(true,null,true,null,null,null,"black", SpreadsheetApp.BorderStyle.DOUBLE).setNumberFormat(reportMoneyFormat_());
    reportSheet.getRange(foot, 4).setValue(crTotal).setFontWeight("bold")
      .setBorder(true,null,true,null,null,null,"black", SpreadsheetApp.BorderStyle.DOUBLE).setNumberFormat(reportMoneyFormat_());

    // 6. Drop down new Balance B/F for next period below totals line
    if (netBal > 0.01) {
      reportSheet.getRange(foot + 1, 1).setValue("Balance B/F").setFontStyle("italic");
      reportSheet.getRange(foot + 1, 2).setValue(netBal).setFontWeight("bold").setNumberFormat(reportMoneyFormat_());
    } else if (netBal < -0.01) {
      reportSheet.getRange(foot + 1, 3).setValue("Balance B/F").setFontStyle("italic");
      reportSheet.getRange(foot + 1, 4).setValue(Math.abs(netBal)).setFontWeight("bold").setNumberFormat(reportMoneyFormat_());
    }

    currentRow += rowCount + 5;
  });

  // --------------------------------------------------------------------------
  // 8. RENDER STATEMENTS 2 - 5 SIDE-BY-SIDE
  // --------------------------------------------------------------------------
  renderTrialBalance(reportSheet, finalBalances, accountInfo, suffix);
  renderProfitLoss(reportSheet, finalBalances, accountInfo, suffix);
  renderBalanceSheet(reportSheet, finalBalances, accountInfo, suffix);
  renderCashFlow(reportSheet, finalBalances, prevBalances, accountInfo, suffix);

  // --------------------------------------------------------------------------
  // 9. FINAL FORMATTING & COLUMN WIDTHS
  // --------------------------------------------------------------------------
  reportSheet.setColumnWidth(1, 90);   // Date (Left)
  reportSheet.setColumnWidth(2, 100);  // Debit Amount
  reportSheet.setColumnWidth(3, 90);   // Date (Right)
  reportSheet.setColumnWidth(4, 100);  // Credit Amount
  reportSheet.setColumnWidth(5, 30);   // Gap
  reportSheet.setColumnWidth(6, 60);   // Code
  reportSheet.setColumnWidth(7, 150);  // Account Name
  reportSheet.setColumnWidth(8, 100);  // Debit
  reportSheet.setColumnWidth(9, 100);  // Credit
  reportSheet.setColumnWidth(10, 30);  // Gap
  reportSheet.setColumnWidth(11, 200); // P&L Description
  reportSheet.setColumnWidth(12, 100); // P&L Amount
  reportSheet.setColumnWidth(13, 30);  // Gap
  reportSheet.setColumnWidth(14, 200); // BS Description
  reportSheet.setColumnWidth(15, 100); // BS Amount
  reportSheet.setColumnWidth(16, 30);  // Gap
  reportSheet.setColumnWidth(17, 220); // Cash Flow Description
  reportSheet.setColumnWidth(18, 100); // Cash Flow Amount

  reportSheet.getRange("A:R").setHorizontalAlignment("center");
  reportSheet.getRange("G:G").setHorizontalAlignment("left");
  reportSheet.getRange("K:K").setHorizontalAlignment("left");
  reportSheet.getRange("N:N").setHorizontalAlignment("left");
  reportSheet.getRange("Q:Q").setHorizontalAlignment("left");

  // Force descriptive code columns to Plain Text
  reportSheet.getRange("F:F").setNumberFormat("@");
  reportSheet.getRange("K:K").setNumberFormat("@");
  reportSheet.getRange("N:N").setNumberFormat("@");

  // Format T-Account transaction columns with Dollar symbol and 2 decimal places
  reportSheet.getRange("B:B").setNumberFormat(reportMoneyFormat_());
  reportSheet.getRange("D:D").setNumberFormat(reportMoneyFormat_());

  SpreadsheetApp.getUi().alert("Consolidated Financial Sheet Generated!");
}

/**
 * Validates whether the active sheet is a valid General Ledger sheet.
 */
function validateLedgerSheet(sheet) {
  const name = sheet.getName();

  const lower = name.trim().toLowerCase();
  if (lower === "dashboard" || lower === "chart of accounts" || lower === "settings" || lower === "start here") {
    return { valid: false, reason: `"${name}" is a system sheet, not a General Ledger.` };
  }

  if (name.startsWith("Financials - ")) {
    return { valid: false, reason: `"${name}" is a report tab. Switch to your monthly General Ledger tab.` };
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { valid: false, reason: `"${name}" does not contain any transaction data.` };
  }

  const data = sheet.getRange(1, 1, Math.min(lastRow, 15), sheet.getLastColumn()).getValues();
  const headerRowIdx = findHeaderRow(data);

  if (headerRowIdx === -1) {
    return { valid: false, reason: `Could not locate headers ("Account", "Debit", "Credit") in top 15 rows of "${name}".` };
  }

  if (lastRow <= headerRowIdx + 1) {
    return { valid: false, reason: `Found headers on Row ${headerRowIdx + 1}, but no transactions are logged below it.` };
  }

  return { valid: true, reason: "" };
}

/**
 * Dynamically scans the top 15 rows to find the header row index (0-based).
 */
function findHeaderRow(data) {
  const maxScan = Math.min(data.length, 15);
  for (let r = 0; r < maxScan; r++) {
    const rowStr = data[r].map(c => String(c).trim().toLowerCase());
    const hasCode = rowStr.some(h => h.includes("account") || h.includes("code"));
    const hasDr = rowStr.some(h => h.includes("debit") || h.includes("dr"));
    const hasCr = rowStr.some(h => h.includes("credit") || h.includes("cr"));

    if (hasCode && hasDr && hasCr) {
      return r;
    }
  }
  return -1;
}

/**
 * Sub-function for Trial Balance (Columns F–I)
 */
function renderTrialBalance(sheet, balances, accountInfo, suffix) {
  let startCol = 6;
  sheet.getRange(1, startCol, 1, 4).merge().setValue(`Trial Balance - ${suffix}`).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#d9d9d9");
  sheet.getRange(2, startCol, 1, 4).setValues([["Code", "Account", "Debit", "Credit"]]).setFontWeight("bold").setBorder(false,false,true,false,false,false);

  let rowIdx = 3;
  let totalDr = 0, totalCr = 0;

  Object.keys(balances).sort().forEach(code => {
    let bal = balances[code];

    if (Math.abs(bal) < 0.01) return;

    let name = accountInfo[code] ? accountInfo[code].name : "Unknown";

    if (bal > 0) {
      sheet.getRange(rowIdx, startCol, 1, 4).setValues([["'" + code, name, bal, ""]]);
      totalDr += bal;
    } else if (bal < 0) {
      sheet.getRange(rowIdx, startCol, 1, 4).setValues([["'" + code, name, "", Math.abs(bal)]]);
      totalCr += Math.abs(bal);
    }
    rowIdx++;
  });

  sheet.getRange(rowIdx, startCol, 1, 4).setValues([["", "TOTAL", totalDr, totalCr]]).setFontWeight("bold").setBorder(true,null,true,null,null,null).setNumberFormat(reportMoneyFormat_());
}

/**
 * Sub-function for Profit or Loss (Columns K–L)
 */
function renderProfitLoss(sheet, balances, accountInfo, suffix) {
  let startCol = 11;
  sheet.getRange(1, startCol, 1, 2).merge().setValue(`Profit or Loss - ${suffix}`).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#d9ead3");

  let rowIdx = 2;
  let revs = [], exps = [], totalRev = 0, totalExp = 0;

  Object.keys(balances).forEach(code => {
    let type = accountInfo[code] ? accountInfo[code].type : "";
    let name = accountInfo[code] ? accountInfo[code].name : "Unknown";
    let bal = balances[code];

    if (Math.abs(bal) < 0.01) return;

    if (type === "Revenue") {
      let val = Math.abs(bal);
      revs.push([name, val]);
      totalRev += val;
    } else if (type === "Expense" || type === "Expenses") {
      let val = bal;
      exps.push([name, val]);
      totalExp += val;
    }
  });

  sheet.getRange(rowIdx++, startCol).setValue("REVENUE").setFontWeight("bold");
  if (revs.length) {
    sheet.getRange(rowIdx, startCol, revs.length, 2).setValues(revs).setNumberFormat(reportMoneyFormat_());
    rowIdx += revs.length;
  }
  sheet.getRange(rowIdx++, startCol, 1, 2).setValues([["Total Revenue", totalRev]]).setFontWeight("bold").setNumberFormat(reportMoneyFormat_());

  rowIdx++;
  sheet.getRange(rowIdx++, startCol).setValue("EXPENSES").setFontWeight("bold");
  if (exps.length) {
    sheet.getRange(rowIdx, startCol, exps.length, 2).setValues(exps).setNumberFormat(reportMoneyFormat_());
    rowIdx += exps.length;
  }
  sheet.getRange(rowIdx++, startCol, 1, 2).setValues([["Total Expenses", totalExp]]).setFontWeight("bold").setNumberFormat(reportMoneyFormat_());

  rowIdx++;
  let net = totalRev - totalExp;
  sheet.getRange(rowIdx, startCol, 1, 2).setValues([["NET PROFIT / (LOSS)", net]]).setFontWeight("bold").setBackground(net >= 0 ? "#d9ead3" : "#f4cccc").setNumberFormat(reportMoneyFormat_());
}

/**
 * Sub-function for Balance Sheet (Columns N–O)
 */
function renderBalanceSheet(sheet, balances, accountInfo, suffix) {
  const startCol = 14;

  sheet.setColumnWidth(13, 30);  // Gap
  sheet.setColumnWidth(14, 200); // Description
  sheet.setColumnWidth(15, 100); // Amount

  sheet.getRange(1, startCol, 1, 2).merge().setValue(`Balance Sheet - ${suffix}`)
    .setFontWeight("bold").setHorizontalAlignment("center").setBackground("#cfe2f3");

  let rowIdx = 2;
  let assets = [], liabilities = [], equity = [];
  let totalAssets = 0, totalLiabilities = 0, totalEquity = 0, netIncome = 0;

  Object.keys(balances).forEach(code => {
    let type = accountInfo[code] ? accountInfo[code].type : "";
    let name = accountInfo[code] ? accountInfo[code].name : "Unknown";
    let bal = balances[code];

    if (type === "Revenue") {
      netIncome += Math.abs(bal);
    } else if (type === "Expense" || type === "Expenses") {
      netIncome -= bal;
    }

    if (Math.abs(bal) < 0.01) return;

    if (type === "Asset" || type === "Assets") {
      assets.push([name, bal]);
      totalAssets += bal;
    } else if (type === "Liability" || type === "Liabilities") {
      let val = Math.abs(bal);
      liabilities.push([name, val]);
      totalLiabilities += val;
    } else if (type === "Equity") {
      let val = bal;
      equity.push([name, Math.abs(val)]);
      totalEquity -= val;
    }
  });

  sheet.getRange(rowIdx++, startCol).setValue("ASSETS").setFontWeight("bold");
  if (assets.length) {
    sheet.getRange(rowIdx, startCol, assets.length, 2).setValues(assets).setNumberFormat(reportMoneyFormat_());
    rowIdx += assets.length;
  }
  sheet.getRange(rowIdx++, startCol, 1, 2).setValues([["Total Assets", totalAssets]])
    .setFontWeight("bold").setNumberFormat(reportMoneyFormat_()).setBorder(true, null, null, null, null, null);

  rowIdx++;

  sheet.getRange(rowIdx++, startCol).setValue("LIABILITIES").setFontWeight("bold");
  if (liabilities.length) {
    sheet.getRange(rowIdx, startCol, liabilities.length, 2).setValues(liabilities).setNumberFormat(reportMoneyFormat_());
    rowIdx += liabilities.length;
  }
  sheet.getRange(rowIdx++, startCol, 1, 2).setValues([["Total Liabilities", totalLiabilities]])
    .setFontWeight("bold").setNumberFormat(reportMoneyFormat_()).setBorder(true, null, null, null, null, null);

  rowIdx++;

  sheet.getRange(rowIdx++, startCol).setValue("EQUITY").setFontWeight("bold");
  if (equity.length) {
    sheet.getRange(rowIdx, startCol, equity.length, 2).setValues(equity).setNumberFormat(reportMoneyFormat_());
    rowIdx += equity.length;
  }

  sheet.getRange(rowIdx++, startCol, 1, 2).setValues([["Net Income / (Loss)", netIncome]]).setNumberFormat(reportMoneyFormat_());
  totalEquity += netIncome;

  sheet.getRange(rowIdx++, startCol, 1, 2).setValues([["Total Equity", totalEquity]])
    .setFontWeight("bold").setNumberFormat(reportMoneyFormat_()).setBorder(true, null, null, null, null, null);

  rowIdx++;

  let totalLE = totalLiabilities + totalEquity;
  let isBalanced = Math.abs(totalAssets - totalLE) < 0.01;

  sheet.getRange(rowIdx, startCol, 1, 2).setValues([["TOTAL LIAB. & EQUITY", totalLE]])
    .setFontWeight("bold").setNumberFormat(reportMoneyFormat_())
    .setBackground(isBalanced ? "#d9ead3" : "#f4cccc");

  sheet.getRange(2, startCol, rowIdx, 1).setHorizontalAlignment("left");
}

// ---------------------------------------------------------------------------
// CURRENCY FORMAT (shared with Cash Flow Statement.gs)
// ---------------------------------------------------------------------------

let REPORT_FMT_CACHE_ = null;

/** Sheets number format using the Currency symbol setting, e.g. "J$"#,##0.00 */
function reportMoneyFormat_() {
  if (REPORT_FMT_CACHE_ === null) {
    let cur = '$';
    try { cur = getCurrencySymbol(); } catch (e) { cur = '$'; }
    REPORT_FMT_CACHE_ = '"' + String(cur).replace(/"/g, '') + '"#,##0.00';
  }
  return REPORT_FMT_CACHE_;
}

// ---------------------------------------------------------------------------
// CASH FLOW STATEMENT (Columns Q–R)
// ---------------------------------------------------------------------------

/**
 * Indirect Statement of Cash Flows (Columns Q-R)
 *
 * REPLACES the old renderCashFlow() in Monthly Reports.gs (delete that one;
 * two functions with the same name will break the project).
 *
 * Changes from the old version:
 *  - Cash accounts and Operating / Investing / Financing grouping come from the
 *    "Cash Flow Class" column on the Chart of Accounts, not hard-coded codes.
 *  - Retained earnings comes from Settings.
 *  - FIX: the cash effect of ANY non-cash account is the negative of the change
 *    in its debit balance. The old code used the opposite sign for Financing,
 *    so loans/equity moved cash the wrong way.
 *  - Accounts with a balance change but no Cash Flow Class are listed in an
 *    "Unclassified" block instead of being silently skipped.
 *  - A check row compares the statement to the actual change in cash balances.
 */
function renderCashFlow(sheet, balances, prevBalances, accountInfo, suffix) {
  const startCol = 17; // Col Q
  const MONEY = reportMoneyFormat_();   // uses the Currency symbol setting

  sheet.getRange(1, startCol, 1, 2).merge().setValue(`Cash Flow Statement - ${suffix}`)
    .setFontWeight("bold").setHorizontalAlignment("center").setBackground("#fce5cd");

  let rowIdx = 2;

  const retainedCode = getRetainedEarningsCode();
  const isCash = code => !!accountInfo[code] && accountInfo[code].cashFlowClass === "Cash";
  const isIncomeStatement = type => type === "Revenue" || type === "Expense" || type === "Expenses";

  // Net income for the month
  let netIncome = 0;
  Object.keys(balances).forEach(code => {
    const type = accountInfo[code] ? accountInfo[code].type : "";
    const bal = balances[code] || 0;
    if (type === "Revenue") netIncome += Math.abs(bal);
    else if (type === "Expense" || type === "Expenses") netIncome -= bal;
  });

  // Beginning & ending cash across all accounts classed as Cash
  let begCash = 0, endCash = 0;
  Object.keys(accountInfo).forEach(code => {
    if (!isCash(code)) return;
    if (prevBalances[code]) {
      begCash += prevBalances[code].type === "Dr" ? prevBalances[code].amount : -prevBalances[code].amount;
    }
    if (balances[code] !== undefined) endCash += balances[code];
  });

  // Group every other balance-sheet movement by its Cash Flow Class
  const items = { Operating: [], Investing: [], Financing: [] };
  const totals = { Operating: netIncome, Investing: 0, Financing: 0 };
  const unclassified = [];

  const allCodes = new Set([...Object.keys(balances), ...Object.keys(prevBalances)]);
  allCodes.forEach(code => {
    if (isCash(code) || code === retainedCode) return;

    const info = accountInfo[code];
    const type = info ? info.type : "";
    if (isIncomeStatement(type)) return;

    const curBal = balances[code] || 0;
    let prevBal = 0;
    if (prevBalances[code]) {
      prevBal = prevBalances[code].type === "Dr" ? prevBalances[code].amount : -prevBalances[code].amount;
    }
    const change = curBal - prevBal;
    if (Math.abs(change) < 0.01) return;

    const label = `Change in ${info ? info.name : "Unknown Account"}`;
    const impact = -change;   // cash effect of any non-cash account = minus its change in debit balance
    const cls = info ? info.cashFlowClass : "";

    if (items[cls]) {
      items[cls].push([label, impact]);
      totals[cls] += impact;
    } else {
      unclassified.push([label, impact]);
    }
  });

  // ---- Render helpers -------------------------------------------------------
  const renderSection = (title, list, total, totalLabel, emptyText, leadRow) => {
    sheet.getRange(rowIdx++, startCol).setValue(title).setFontWeight("bold");
    if (leadRow) {
      sheet.getRange(rowIdx++, startCol, 1, 2).setValues([leadRow]).setNumberFormat(MONEY);
    }
    if (list.length) {
      sheet.getRange(rowIdx, startCol, list.length, 2).setValues(list).setNumberFormat(MONEY);
      rowIdx += list.length;
    } else if (emptyText) {
      sheet.getRange(rowIdx++, startCol, 1, 2).setValues([[emptyText, 0]]).setNumberFormat(MONEY);
    }
    sheet.getRange(rowIdx++, startCol, 1, 2).setValues([[totalLabel, total]])
      .setFontWeight("bold").setNumberFormat(MONEY).setBorder(true, null, null, null, null, null);
    rowIdx++;
  };

  renderSection("CASH FLOW FROM OPERATING ACTIVITIES", items.Operating, totals.Operating,
    "Net Cash from Operating Activities", null, ["Net Income / (Loss)", netIncome]);
  renderSection("CASH FLOW FROM INVESTING ACTIVITIES", items.Investing, totals.Investing,
    "Net Cash from Investing Activities", "No investing activity");
  renderSection("CASH FLOW FROM FINANCING ACTIVITIES", items.Financing, totals.Financing,
    "Net Cash from Financing Activities", "No financing activity");

  // ---- Cash reconciliation --------------------------------------------------
  const netCashChange = totals.Operating + totals.Investing + totals.Financing;
  sheet.getRange(rowIdx++, startCol, 1, 2).setValues([["NET INCREASE / (DECREASE) IN CASH", netCashChange]])
    .setFontWeight("bold").setNumberFormat(MONEY).setBorder(true, null, null, null, null, null);

  sheet.getRange(rowIdx++, startCol, 1, 2).setValues([["Beginning Cash Balance", begCash]]).setNumberFormat(MONEY);
  sheet.getRange(rowIdx++, startCol, 1, 2).setValues([["ENDING CASH BALANCE", endCash]])
    .setFontWeight("bold").setNumberFormat(MONEY).setBackground("#d9ead3");

  // Check: does the statement explain the actual movement in cash?
  const diff = (begCash + netCashChange) - endCash;
  const ok = Math.abs(diff) < 0.01;
  sheet.getRange(rowIdx++, startCol, 1, 2)
    .setValues([[ok ? "Check: statement matches cash balances" : "CHECK FAILED: statement differs from cash by", ok ? 0 : diff]])
    .setNumberFormat(MONEY).setFontStyle("italic")
    .setBackground(ok ? "#d9ead3" : "#f4cccc");

  // Anything that moved but has no Cash Flow Class
  if (unclassified.length) {
    rowIdx++;
    sheet.getRange(rowIdx++, startCol, 1, 2)
      .setValues([["UNCLASSIFIED (set a Cash Flow Class on the Chart of Accounts)", ""]])
      .setFontWeight("bold").setBackground("#f4cccc");
    sheet.getRange(rowIdx, startCol, unclassified.length, 2).setValues(unclassified).setNumberFormat(MONEY);
    rowIdx += unclassified.length;
  }

  sheet.getRange(2, startCol, rowIdx, 1).setHorizontalAlignment("left");
}
