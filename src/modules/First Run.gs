/**
 * ============================================================================
 * BETTER BUDGET - FIRST-RUN SETUP
 * ============================================================================
 * Called at the end of setUpOrUpgradeTool() (Settings and Setup.gs).
 * Needs Styling.gs.
 *
 * For a new user it:
 *   1. creates their first monthly ledger tab (if the workbook has none), using
 *      the layout the rest of the tool expects:
 *        B3 month, B4 year, F2/F3 starting balance, F5/F6 adjusted balance,
 *        table headers on row 8 (Date | Account | Description | Debit | Credit)
 *   2. adds a "Start Here" sheet with plain-English instructions
 *   3. puts the tabs in a sensible order and removes the empty default sheet
 *
 * Existing users are not affected: nothing is created if a ledger tab already
 * exists, and an existing "Start Here" sheet is never overwritten.
 */

const FIRST_RUN_MONTHS_ = ['January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'];

const START_HERE_SHEET_NAME = 'Start Here';

// NOTE: wording is chosen so findHeaderRow() never mistakes this sheet for a ledger.
function startHereRows_() {
  return [
  ['Better Budget - Start Here', ''],
  ['', ''],
  ['1. Add your starting balances',
    'Open Better Budget > Make Transaction Entry and choose "Opening balance". Enter what each bank account, wallet or card held (or owed) on the first day you want to track. Do this once for each one.'],
  ['2. Record your day-to-day activity',
    'Use the same form and pick what happened: spent money, received money, moved money, or got money back. The tool does the bookkeeping for you.'],
  ['3. Set monthly budgets',
    'On the Chart of Accounts tab, type an amount in the Monthly Budget column next to any spending category you want to track.'],
  ['4. See your reports',
    'On a monthly ledger tab, choose Better Budget > Generate Monthly Reports. Then choose Generate Executive Dashboard for the big picture.'],
  ['5. Start each new month',
    'On the latest ledger tab, choose Better Budget > Start New Month. Generate that month\'s reports first so the balances carry forward correctly.'],
  ['6. Get a weekly email (optional)',
    'Choose Better Budget > Weekly Email > Turn on. Every Monday you get a short note showing which budgets are over, or on pace to run over. Use "Send now" to preview it.'],
  ['Good to know',
    'Your settings (currency, main bank account, report email and so on) live on the Chart of Accounts tab, in columns H to J. Rename or add accounts there whenever you like.'],
  ['About',
    'Better Budget v' + BETTER_BUDGET_VERSION + '. A personal budgeting aid, not professional financial, tax or legal advice. Check important figures yourself and keep your own backups.'],
  ];
}

// ---------------------------------------------------------------------------
// ENTRY POINT (called from setUpOrUpgradeTool)
// ---------------------------------------------------------------------------

/** Returns a short summary (possibly empty) to show in the final "Done" dialog. */
function runFirstRunSteps_(ss, ui) {
  const notes = [];
  let ledgerSheet = null;
  let createdStartHere = false;
  const originalActive = ss.getActiveSheet();

  // 1. First ledger tab
  if (!hasLedgerTab_(ss)) {
    const tz = ss.getSpreadsheetTimeZone();
    const nowMonth = Number(Utilities.formatDate(new Date(), tz, 'M')) - 1;
    const nowYear = Number(Utilities.formatDate(new Date(), tz, 'yyyy'));
    const nowLabel = FIRST_RUN_MONTHS_[nowMonth] + ' ' + nowYear;

    const resp = ui.prompt(
      'Create your first ledger',
      'No monthly ledger tab was found, so let\'s create one.\n\n' +
      'Which month do you want to start with? Type it like "October 2026", or leave blank for the current month (' + nowLabel + ').',
      ui.ButtonSet.OK_CANCEL
    );

    if (resp.getSelectedButton() === ui.Button.OK) {
      const text = resp.getResponseText().trim();
      const parsed = text === '' ? { monthIdx: nowMonth, year: nowYear } : parseMonthYear_(text);

      if (!parsed) {
        notes.push('No ledger tab was created because "' + text + '" is not a month and year. ' +
          'Run Set Up / Upgrade Tool again and type something like "October 2026".');
      } else {
        const sheetName = FIRST_RUN_MONTHS_[parsed.monthIdx] + ' ' + parsed.year;
        if (ss.getSheetByName(sheetName)) {
          notes.push('A sheet named "' + sheetName + '" already exists, so no ledger was created.');
        } else {
          ledgerSheet = createLedgerSheet_(ss, parsed.monthIdx, parsed.year);
          notes.push('Created your first ledger tab: ' + sheetName + '.');
        }
      }
    }
  }

  // 2. Start Here sheet
  if (!findSheetCI_(ss, START_HERE_SHEET_NAME)) {
    createStartHereSheet_(ss);
    createdStartHere = true;
    notes.push('Added a "Start Here" sheet with instructions.');
  }

  // 3. New workbook: sensible tab order, and drop the empty default sheet
  if (ledgerSheet && createdStartHere) {
    orderNewWorkbookTabs_(ss, ledgerSheet);
    deleteEmptyDefaultSheet_(ss);
  }

  try { ss.setActiveSheet(ledgerSheet || originalActive); } catch (e) { /* sheet may be gone */ }
  return notes.join('\n');
}

/** Start Here, Chart of Accounts, then the ledger. */
function orderNewWorkbookTabs_(ss, ledger) {
  const order = [findSheetCI_(ss, START_HERE_SHEET_NAME), getChartSheet_(ss), ledger];
  let position = 1;
  order.forEach(function (sh) {
    if (!sh) return;
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(position++);
  });
}

/** Deletes an untouched "Sheet1"-style default sheet. Never deletes a sheet with content. */
function deleteEmptyDefaultSheet_(ss) {
  ss.getSheets().forEach(function (sh) {
    if (/^Sheet\d+$/.test(sh.getName()) && sh.getLastRow() === 0 && ss.getSheets().length > 1) {
      ss.deleteSheet(sh);
    }
  });
}

// ---------------------------------------------------------------------------
// LEDGER TAB
// ---------------------------------------------------------------------------

/** True if any sheet in the workbook is already a monthly ledger. */
function hasLedgerTab_(ss) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (getLedgerTarget_(sheets[i]).ok) return true;
  }
  return false;
}

/** Accepts "October 2026", "Oct 2026", "oct. 2026". Returns {monthIdx, year} or null. */
function parseMonthYear_(text) {
  const m = String(text || '').trim().match(/^([A-Za-z]{3,9})\.?,?\s+(\d{4})$/);
  if (!m) return null;
  const word = m[1].toLowerCase();
  const idx = FIRST_RUN_MONTHS_.findIndex(function (n) { return n.toLowerCase().indexOf(word) === 0; });
  const year = Number(m[2]);
  if (idx === -1 || year < 1990 || year > 2100) return null;
  return { monthIdx: idx, year: year };
}

/**
 * The top of a ledger tab (rows 1-8, columns A-F) as a 2D array.
 * Kept separate so the layout can be tested without a spreadsheet.
 * Cell positions are fixed because the other files read them:
 *   B3 month, B4 year, F2 / F3 starting balance, F5 / F6 adjusted balance.
 */
function ledgerTopBlock_(monthName, year, bankName, bankCode) {
  const crit = /^\d+$/.test(bankCode) ? bankCode : '"' + bankCode + '"';
  return [
    ['General Ledger', '', '', '', '', ''],
    ['Period', '', '', '', '', bankName + ' Starting Balance'],
    ['Month', monthName, '', '', '', ''],
    ['Year', year, '', '', '', ''],
    ['', '', '', '', '', bankName + ' Adjusted Balance'],
    ['', '', '', '', '', '=F3-SUMIFS(E:E,B:B,' + crit + ')+SUMIFS(D:D,B:B,' + crit + ')'],
    ['Tip: record entries from the Better Budget menu (Make Transaction Entry).', '', '', '', '', ''],
    ['Date', 'Account', 'Description', 'Debit', 'Credit', ''],
  ];
}

/** Builds a new, empty, styled monthly ledger tab and returns it. */
function createLedgerSheet_(ss, monthIdx, year) {
  const monthName = FIRST_RUN_MONTHS_[monthIdx];

  // Bank label comes from the user's own chart and settings
  const bankCode = getPrimaryBankCode();
  let bankName = 'Bank';
  try {
    const chart = readChartOfAccounts();
    if (chart[bankCode]) bankName = String(chart[bankCode].name);
  } catch (e) { /* keep the generic label */ }

  const sh = ss.insertSheet(monthName + ' ' + year);   // activates the new sheet

  // Trim the endless empty columns; keep A-H
  if (sh.getMaxColumns() > 8) sh.deleteColumns(9, sh.getMaxColumns() - 8);

  const block = ledgerTopBlock_(monthName, year, bankName, bankCode);
  sh.getRange(1, 1, block.length, 6).setValues(block);
  sh.setHiddenGridlines(true);
  sh.setTabColor('#2563eb');

  // Title bar across the whole top block
  sh.getRange('A1:F1').merge()
    .setFontWeight('bold').setFontSize(14).setFontColor('#ffffff').setBackground(BB_NAVY_)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 38);

  // Period block (B3 month, B4 year)
  sh.getRange('A2').setFontSize(9).setFontWeight('bold').setFontColor(BB_MUTED_);
  sh.getRange('A3:A4').setFontWeight('bold').setFontColor(BB_MUTED_);
  sh.getRange('B3:B4').setFontWeight('bold').setBackground(BB_LIGHT_).setHorizontalAlignment('left');
  sh.getRange('B4').setNumberFormat('0');
  styleTableLines_(sh.getRange('B3:B4'));

  // Bank block: small labels over larger value boxes (F3 is typed in, F6 is calculated)
  sh.getRange('F2').setFontSize(9).setFontWeight('bold').setFontColor(BB_MUTED_);
  sh.getRange('F5').setFontSize(9).setFontWeight('bold').setFontColor(BB_MUTED_);
  sh.getRange('F3').setNumberFormat('#,##0.00').setFontSize(12).setFontWeight('bold')
    .setBackground(BB_INPUT_).setHorizontalAlignment('right');
  sh.getRange('F6').setNumberFormat('#,##0.00').setFontSize(12).setFontWeight('bold')
    .setBackground(BB_LIGHT_).setHorizontalAlignment('right');
  styleTableLines_(sh.getRange('F3'));
  styleTableLines_(sh.getRange('F6'));

  // Tip
  sh.getRange('A7').setFontSize(9).setFontStyle('italic').setFontColor(BB_MUTED_);

  // Entry table: header row 8, then ruled rows
  styleHeader_(sh.getRange('A8:E8'));
  sh.setRowHeight(8, 28);

  const rowsBelow = Math.max(sh.getMaxRows() - 8, 1);
  const entries = sh.getRange(9, 1, rowsBelow, 5);
  entries.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  entries.setVerticalAlignment('middle');
  sh.getRange(9, 1, rowsBelow, 1).setNumberFormat('yyyy-mm-dd').setHorizontalAlignment('center');
  sh.getRange(9, 2, rowsBelow, 1).setHorizontalAlignment('center');
  sh.getRange(9, 3, rowsBelow, 1).setHorizontalAlignment('left');
  sh.getRange(9, 4, rowsBelow, 2).setNumberFormat('#,##0.00').setHorizontalAlignment('right');

  sh.setColumnWidth(1, 110);
  sh.setColumnWidth(2, 90);
  sh.setColumnWidth(3, 280);
  sh.setColumnWidth(4, 110);
  sh.setColumnWidth(5, 110);
  sh.setColumnWidth(6, 270);
  sh.setFrozenRows(8);

  return sh;
}

// ---------------------------------------------------------------------------
// START HERE SHEET
// ---------------------------------------------------------------------------

function createStartHereSheet_(ss) {
  const sh = ss.insertSheet(START_HERE_SHEET_NAME, 0);   // first tab; activates it
  if (sh.getMaxColumns() > 2) sh.deleteColumns(3, sh.getMaxColumns() - 2);

  const rows = startHereRows_().length;
  sh.getRange(1, 1, rows, 2).setValues(startHereRows_());
  sh.setHiddenGridlines(true);
  sh.setTabColor('#059669');

  // Title bar
  sh.getRange(1, 1, 1, 2).merge()
    .setFontWeight('bold').setFontSize(15).setFontColor('#ffffff').setBackground(BB_NAVY_)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 44);
  sh.setRowHeight(2, 14);

  // Steps: bold heading on the left, wrapped explanation on the right
  const count = rows - 2;
  const body = sh.getRange(3, 1, count, 2);
  body.setVerticalAlignment('middle').setWrap(true).setFontSize(11);
  styleTableLines_(body);
  sh.getRange(3, 1, count, 1).setFontWeight('bold').setFontColor(BB_NAVY_).setBackground(BB_LIGHT_);
  sh.getRange(3, 2, count, 1).setFontColor(BB_SLATE_);
  sh.setRowHeights(3, count, 66);

  // The last two rows are side notes, not steps
  sh.getRange(rows - 1, 1, 2, 2).setFontStyle('italic').setFontColor(BB_MUTED_).setFontSize(10);
  sh.setRowHeights(rows - 1, 2, 50);

  sh.setColumnWidth(1, 250);
  sh.setColumnWidth(2, 580);
  return sh;
}
