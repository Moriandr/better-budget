/**
 * Rolls the General Ledger forward to a new month.
 *
 * REPLACES the old startNextMonth() file entirely.
 *
 * Changes from the old version:
 *  - The bank account is the "Primary bank account code" from Settings
 *    (no more hard-coded 1020 / "NCB").
 *  - Only runs from a real ledger tab (uses validateLedgerSheet).
 *  - Shows where the opening balance came from, and warns if the Financials
 *    tab and the ledger disagree (a sign the Financials tab is out of date).
 *  - If no opening balance can be found, the new tab's starting balance is
 *    left BLANK instead of silently keeping last month's figure.
 *  - The new tab's "Adjusted Balance" formula and labels are rebuilt for the
 *    primary bank account. The account code is written into the formula, so
 *    changing the setting later never alters past months.
 *
 * Ledger top block (unchanged): B3 month, B4 year, F2/F3 starting balance,
 * F5/F6 adjusted balance. Header row is found automatically (default row 8).
 */
function startNextMonth() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sourceSheet = ss.getActiveSheet();

  // 1. Must be a real ledger tab
  const check = validateLedgerSheet(sourceSheet);
  if (!check.valid) {
    ui.alert("Start New Month",
      `Switch to the monthly General Ledger tab you want to roll forward.\n\nReason: ${check.reason}`,
      ui.ButtonSet.OK);
    return;
  }

  // 2. Current month & year
  const currentMonthStr = String(sourceSheet.getRange("B3").getValue()).trim();
  const currentYear = Number(sourceSheet.getRange("B4").getValue());

  if (!currentMonthStr || isNaN(currentYear)) {
    ui.alert("Error: Active sheet must have a valid Month in B3 and Year in B4.");
    return;
  }

  const monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
  const monthIndex = monthNames.findIndex(m => m.toLowerCase() === currentMonthStr.toLowerCase());
  if (monthIndex === -1) {
    ui.alert(`Error: '${currentMonthStr}' is not a recognized month name.`);
    return;
  }

  // 3. Next month & year; don't overwrite an existing tab
  const nextMonthStr = monthNames[(monthIndex + 1) % 12];
  const nextYear = monthIndex === 11 ? currentYear + 1 : currentYear;
  const newSheetName = `${nextMonthStr} ${nextYear}`;

  if (ss.getSheetByName(newSheetName)) {
    ui.alert(`Sheet '${newSheetName}' already exists.`);
    return;
  }

  // 4. Primary bank account from Settings must exist in the Chart of Accounts
  const bankCode = getPrimaryBankCode();
  let chart = {};
  try { chart = readChartOfAccounts(); } catch (e) { chart = {}; }
  if (!chart[bankCode]) {
    ui.alert("Start New Month",
      `The Primary bank account code in Settings ("${bankCode}") is not in your Chart of Accounts.\n\nFix the setting and try again.`,
      ui.ButtonSet.OK);
    return;
  }
  const bankName = chart[bankCode].name;

  // 5. Work out the closing balance to carry forward
  const ledgerAdjusted = sourceSheet.getRange("F6").getValue();
  const hasLedgerBalance = typeof ledgerAdjusted === "number";

  let closingBalance = null;
  let sourceNote = "";
  let warning = "";

  const finSheet = ss.getSheetByName(`Financials - ${currentMonthStr} ${currentYear}`);
  if (finSheet) {
    // Trial Balance is in columns F-I from row 3. A missing account means a zero balance.
    closingBalance = 0;
    const lastRowFin = finSheet.getLastRow();
    if (lastRowFin >= 3) {
      const tbData = finSheet.getRange(3, 6, lastRowFin - 2, 4).getValues();
      for (const row of tbData) {
        const code = String(row[0]).replace("'", "").trim();
        if (code === bankCode) {
          closingBalance = (Number(row[2]) || 0) - (Number(row[3]) || 0);
          break;
        }
      }
    }
    sourceNote = `taken from the "Financials - ${currentMonthStr} ${currentYear}" trial balance`;

    if (hasLedgerBalance && Math.abs(ledgerAdjusted - closingBalance) > 0.01) {
      warning = `\n\nWARNING: the Financials tab shows ${formatMoney_(closingBalance)} but this ledger's ` +
        `Adjusted Balance (F6) shows ${formatMoney_(ledgerAdjusted)}. If you've added entries since you ` +
        `generated the reports, cancel and run "Generate Monthly Reports" first.`;
    }
  } else if (hasLedgerBalance) {
    closingBalance = ledgerAdjusted;
    sourceNote = `taken from this ledger's Adjusted Balance (F6), because no Financials tab exists for ${currentMonthStr} ${currentYear}`;
  } else {
    warning = `\n\nWARNING: no opening balance could be found (no Financials tab, and F6 is not a number). ` +
      `The new tab's starting balance will be left blank - you'll need to enter it.`;
  }

  // 6. Confirm, showing the balance and where it came from
  const balanceLine = closingBalance === null
    ? "Opening balance: (none found)"
    : `Opening balance for ${bankName}: ${formatMoney_(closingBalance)}\n(${sourceNote})`;

  const response = ui.alert(
    "Start New Month",
    `Create new General Ledger sheet for '${newSheetName}'?\n\n${balanceLine}${warning}`,
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) {
    ss.toast("Action cancelled.", "Cancelled", 3);
    return;
  }

  // 7. Find the header row before copying (default row 8, as before)
  const hIdx = findHeaderRow(sourceSheet.getDataRange().getValues());
  const headerRow = hIdx !== -1 ? hIdx + 1 : 8;

  // 8. Duplicate and configure the new tab
  const newSheet = sourceSheet.copyTo(ss);
  newSheet.setName(newSheetName);
  ss.setActiveSheet(newSheet);

  newSheet.getRange("B3").setValue(nextMonthStr);
  newSheet.getRange("B4").setValue(nextYear);

  newSheet.getRange("F2").setValue(`${bankName} Starting Balance`);
  newSheet.getRange("F5").setValue(`${bankName} Adjusted Balance`);

  if (closingBalance !== null) {
    newSheet.getRange("F3").setValue(closingBalance);
  } else {
    newSheet.getRange("F3").clearContent();
  }

  // Write the account code into the formula (numbers stay numeric, text is quoted)
  const crit = /^\d+$/.test(bankCode) ? bankCode : `"${bankCode}"`;
  newSheet.getRange("F6").setFormula(`=F3-SUMIFS(E:E,B:B,${crit})+SUMIFS(D:D,B:B,${crit})`);

  // 9. Clear last month's transactions (below the header row, columns A-E)
  const lastRowNew = newSheet.getLastRow();
  if (lastRowNew > headerRow) {
    newSheet.getRange(headerRow + 1, 1, lastRowNew - headerRow, 5).clearContent();
  }

  ss.toast(`Successfully created '${newSheetName}'!`, "Success", 4);
}

/** Simple money formatter for dialog text. */
function formatMoney_(n) {
  const neg = n < 0;
  const s = Math.abs(n).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return (neg ? "-" : "") + getCurrencySymbol() + s;
}
