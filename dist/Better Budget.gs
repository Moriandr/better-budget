// ============================================================================
// BETTER BUDGET - COMPLETE SCRIPT
// ============================================================================
// This single file contains the whole tool. Sections, in order:
//   1. Custom Menu          2. Settings and Setup    3. Styling
//   4. First Run            5. Record sidebar        6. Monthly Reports
//   7. New Month            8. Dashboard             9. Weekly Report
//  10. About
// The entry form lives in the separate SidebarHTML file.
// ============================================================================


// ============================================================================
// SECTION 1 OF 10: CUSTOM MENU
// ============================================================================

/**
 * Adds the "Better Budget" menu whenever the spreadsheet is opened.
 * (After pasting this in, reload the sheet once so the menu rebuilds.)
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Better Budget')
    .addItem('Make Transaction Entry', 'showSidebar')
    .addItem('Start New Month', 'startNextMonth')
    .addItem('Generate Monthly Reports', 'generateConsolidatedReports')
    .addItem('Generate Executive Dashboard', 'generateDashboard')
    .addSeparator()
    .addSubMenu(ui.createMenu('Weekly Email')
      .addItem('Turn on', 'turnOnWeeklyEmail')
      .addItem('Turn off', 'turnOffWeeklyEmail')
      .addItem('Status', 'weeklyEmailStatus')
      .addItem('Send now', 'sendWeeklyReportNow'))
    .addSeparator()
    .addItem('Set Up / Upgrade Tool', 'setUpOrUpgradeTool')
    .addItem('About', 'showAbout')
    .addToUi();
}


// ============================================================================
// SECTION 2 OF 10: SETTINGS AND SETUP
// ============================================================================

/**
 * ============================================================================
 * BETTER BUDGET - CHART OF ACCOUNTS, SETTINGS & SETUP
 * ============================================================================
 * REPLACES the previous "Settings and Setup.gs" (and its Settings sheet).
 *
 * Everything lives on the "Chart of Accounts" sheet:
 *
 *   A-F   Chart of Accounts table
 *         Account Code | Account Name | Category | Monthly Budget | Cash Flow Class | Dashboard Group
 *   H-J   Settings table   (Setting | Value | What it does)
 *   L     Dashboard groups list (the choices for the Dashboard Group column)
 *
 * Run setUpOrUpgradeTool() from the menu. It is safe to run more than once and
 * never overwrites values you have already filled in. On a sheet with no
 * accounts it offers to load a generic starter chart.
 *
 * Public helpers used by the other files:
 *   readChartOfAccounts(), getSetting(), getCurrencySymbol(),
 *   getPrimaryBankCode(), getRetainedEarningsCode(), getOpeningBalanceCode(),
 *   getDashboardGroups(), findSheetCI_(), getChartSheet_()
 */

const CHART_SHEET_NAME = 'Chart of Accounts';

// Layout of the Chart of Accounts sheet
const CHART_COLS = 6;        // table occupies columns A:F
const SETTINGS_COL = 8;      // settings table starts in column H
const GROUPS_COL = 12;       // dashboard groups list is column L
const SETTINGS_ROWS = 30;    // rows scanned when reading settings / groups

const ACCOUNT_CATEGORIES = ['Asset', 'Liability', 'Equity', 'Revenue', 'Expense'];
const CASH_FLOW_CLASSES = ['Cash', 'Operating', 'Investing', 'Financing'];
const DEFAULT_DASHBOARD_GROUPS = ['Housing', 'Groceries', 'Dining Out', 'Utilities', 'Transport', 'Other'];

const CHART_HEADERS = ['Account Code', 'Account Name', 'Category', 'Monthly Budget', 'Cash Flow Class', 'Dashboard Group'];
const CHART_BUDGET_HEADER = 'Monthly Budget';
const CHART_CASHFLOW_HEADER = 'Cash Flow Class';
const CHART_GROUP_HEADER = 'Dashboard Group';

const SETTING_DEFAULTS = [
  ['Currency symbol', '$', 'Shown in reports, the dashboard and emails.'],
  ['Primary bank account code', '1020', 'The account whose balance is carried into each new month.'],
  ['Retained earnings account code', '3020', 'Where each month\'s profit or loss is rolled into.'],
  ['Opening balance account code', '3010', 'Used to record the starting balances of your accounts.'],
  ['Report email', '', 'Where the weekly report is sent. Blank = your own Google account.'],
];

/**
 * Generic starter chart: [code, name, category, monthly budget, cash flow class, dashboard group]
 */
const STARTER_CHART = [
  // Assets
  ['1010', 'Cash on Hand', 'Asset', '', 'Cash', ''],
  ['1020', 'Primary Bank Account', 'Asset', '', 'Cash', ''],
  ['1030', 'Savings Account', 'Asset', '', 'Cash', ''],
  ['1040', 'Mobile Wallet / Prepaid Card', 'Asset', '', 'Cash', ''],
  ['1100', 'Investments', 'Asset', '', 'Investing', ''],
  ['1200', 'Property & Equipment', 'Asset', '', 'Investing', ''],
  ['1300', 'Money Owed to Me', 'Asset', '', 'Operating', ''],
  // Liabilities
  ['2010', 'Credit Card', 'Liability', '', 'Operating', ''],
  ['2020', 'Bills & Amounts Owed', 'Liability', '', 'Operating', ''],
  ['2100', 'Loans', 'Liability', '', 'Financing', ''],
  ['2200', 'Mortgage', 'Liability', '', 'Financing', ''],
  // Equity
  ['3010', 'Opening Balances', 'Equity', '', 'Financing', ''],
  ['3020', 'Retained Earnings', 'Equity', '', '', ''],
  // Revenue
  ['4010', 'Salary & Wages', 'Revenue', '', '', ''],
  ['4020', 'Business & Side Income', 'Revenue', '', '', ''],
  ['4030', 'Interest & Dividends', 'Revenue', '', '', ''],
  ['4090', 'Other Income', 'Revenue', '', '', ''],
  // Expenses
  ['5010', 'Housing (Rent / Mortgage)', 'Expense', '', '', 'Housing'],
  ['5020', 'Groceries', 'Expense', '', '', 'Groceries'],
  ['5030', 'Dining Out', 'Expense', '', '', 'Dining Out'],
  ['5040', 'Utilities', 'Expense', '', '', 'Utilities'],
  ['5050', 'Transportation', 'Expense', '', '', 'Transport'],
  ['5060', 'Insurance', 'Expense', '', '', 'Other'],
  ['5070', 'Healthcare', 'Expense', '', '', 'Other'],
  ['5080', 'Subscriptions & Phone', 'Expense', '', '', 'Other'],
  ['5090', 'Personal Care', 'Expense', '', '', 'Other'],
  ['5100', 'Entertainment', 'Expense', '', '', 'Other'],
  ['5110', 'Education', 'Expense', '', '', 'Other'],
  ['5120', 'Gifts & Donations', 'Expense', '', '', 'Other'],
  ['5130', 'Bank Fees & Interest', 'Expense', '', '', 'Other'],
  ['5900', 'Miscellaneous', 'Expense', '', '', 'Other'],
];

// ---------------------------------------------------------------------------
// MENU ENTRY POINT
// ---------------------------------------------------------------------------

/**
 * Sets up (or upgrades) the Chart of Accounts sheet:
 *  - loads the starter chart if there are no accounts yet
 *  - creates the Settings table and dashboard groups list
 *  - adds missing columns and dropdowns, and fills BLANK cells only
 */
function setUpOrUpgradeTool() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const previousSheet = ss.getActiveSheet();

  let chart = getChartSheet_(ss);
  let loadedStarter = false;

  if (!chart || countAccounts_(chart) === 0) {
    const load = ui.alert(
      'Set Up Tool',
      'No accounts were found.\n\nLoad the starter Chart of Accounts? It is a ready-made, generic list of ' +
      'account types (bank, credit card, salary, groceries, rent and so on) that you can edit afterwards.',
      ui.ButtonSet.YES_NO
    );
    if (load !== ui.Button.YES) {
      ss.toast('Cancelled. Nothing was changed.', 'Cancelled', 3);
      return;
    }
    chart = loadStarterChart_(ss);
    loadedStarter = true;
  } else {
    const answer = ui.alert(
      'Upgrade Tool',
      'This will:\n' +
      '- add the Settings table (columns H-J) and Dashboard groups list (column L) if missing\n' +
      '- add any missing Chart of Accounts columns, with dropdown lists\n' +
      '- fill BLANK cells only, using your current rules\n\n' +
      'Existing values are never overwritten. Tip: File > Make a copy first if you want a backup.\n\nContinue?',
      ui.ButtonSet.YES_NO
    );
    if (answer !== ui.Button.YES) {
      ss.toast('Cancelled. Nothing was changed.', 'Cancelled', 3);
      return;
    }
  }

  const migrated = ensureSettingsTable_(ss, chart);
  const result = upgradeChartOfAccounts_(ss, chart);
  restoreActiveSheet_(ss, previousSheet);

  // New workbook? Create the first ledger tab and a "Start Here" sheet (see First Run.gs)
  const firstRun = runFirstRunSteps_(ss, ui);

  // Sanity-check the account codes stored in Settings
  const chartMap = readChartOfAccounts();
  const warnings = [];
  [['Primary bank account code', 'Primary bank'], ['Retained earnings account code', 'Retained earnings'],
   ['Opening balance account code', 'Opening balance']].forEach(function (pair) {
    const code = getSetting(pair[0]);
    if (!chartMap[code]) warnings.push(pair[1] + ' code "' + code + '" is not in your Chart of Accounts.');
  });

  ui.alert(
    'Done',
    (loadedStarter ? 'Starter Chart of Accounts loaded.\n' : '') +
    'Columns added: ' + (result.addedColumns.length ? result.addedColumns.join(', ') : 'none (already present)') + '\n' +
    'Cash Flow Class cells filled: ' + result.cashFlowFilled + '\n' +
    'Dashboard Group cells filled: ' + result.groupFilled +
    (migrated ? '\n\nYour values from the old "Settings" sheet were copied to the Chart of Accounts sheet. ' +
                'You can now delete the old "Settings" sheet.' : '') +
    (firstRun ? '\n\n' + firstRun : '') +
    (warnings.length ? '\n\nCheck these:\n- ' + warnings.join('\n- ') : ''),
    ui.ButtonSet.OK
  );
}

// ---------------------------------------------------------------------------
// SETTINGS (stored on the Chart of Accounts sheet, columns H-J)
// ---------------------------------------------------------------------------

/** Returns a setting's value as text, or its default if the row is missing or blank. */
function getSetting(key) {
  const def = SETTING_DEFAULTS.filter(function (d) { return d[0] === key; })[0];
  const fallback = def ? def[1] : '';
  const chart = getChartSheet_();
  if (!chart || chart.getMaxColumns() < SETTINGS_COL + 1) return fallback;

  const rows = chart.getRange(1, SETTINGS_COL, SETTINGS_ROWS, 2).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === key.toLowerCase()) {
      const v = String(rows[i][1]).trim();
      return v !== '' ? v : fallback;
    }
  }
  return fallback;
}

function getCurrencySymbol() { return getSetting('Currency symbol'); }
function getPrimaryBankCode() { return getSetting('Primary bank account code'); }
function getRetainedEarningsCode() { return getSetting('Retained earnings account code'); }
function getOpeningBalanceCode() { return getSetting('Opening balance account code'); }

/** The dashboard groups the user has defined (column L of the Chart of Accounts sheet). */
function getDashboardGroups() {
  const chart = getChartSheet_();
  if (!chart || chart.getMaxColumns() < GROUPS_COL) return DEFAULT_DASHBOARD_GROUPS.slice();
  const vals = chart.getRange(2, GROUPS_COL, SETTINGS_ROWS, 1).getValues();
  const groups = vals.map(function (r) { return String(r[0]).trim(); })
    .filter(function (g) { return g !== ''; });
  return groups.length ? groups : DEFAULT_DASHBOARD_GROUPS.slice();
}

/**
 * Creates the Settings table and groups list on the chart sheet if missing, and
 * adds any missing setting rows. If an older separate "Settings" sheet exists,
 * its values are copied over. Returns true if values were migrated from it.
 */
function ensureSettingsTable_(ss, chart) {
  if (chart.getMaxColumns() < GROUPS_COL) {
    chart.insertColumnsAfter(chart.getMaxColumns(), GROUPS_COL - chart.getMaxColumns());
  }

  const legacy = readLegacySettingsSheet_(ss);
  let migrated = false;

  // Headers
  const styleHeader = function (range) {
    range.setFontWeight('bold').setBackground('#334155').setFontColor('#ffffff');
  };
  if (String(chart.getRange(1, SETTINGS_COL).getValue()).trim() === '') {
    styleHeader(chart.getRange(1, SETTINGS_COL, 1, 3).setValues([['Setting', 'Value', 'What it does']]));
  }
  if (String(chart.getRange(1, GROUPS_COL).getValue()).trim() === '') {
    styleHeader(chart.getRange(1, GROUPS_COL).setValue('Dashboard groups'));
  }

  // Setting rows: add any that are missing
  const existing = chart.getRange(2, SETTINGS_COL, SETTINGS_ROWS - 1, 1).getValues()
    .map(function (r) { return String(r[0]).trim().toLowerCase(); });
  let nextRow = 2;
  while (nextRow - 2 < existing.length && existing[nextRow - 2] !== '') nextRow++;

  SETTING_DEFAULTS.forEach(function (d) {
    if (existing.indexOf(d[0].toLowerCase()) !== -1) return;
    const legacyValue = legacy && legacy.values[d[0].toLowerCase()];
    const value = (legacyValue !== undefined && legacyValue !== '') ? legacyValue : d[1];
    if (legacyValue !== undefined && legacyValue !== '') migrated = true;
    chart.getRange(nextRow, SETTINGS_COL + 1).setNumberFormat('@');   // keep codes as text
    chart.getRange(nextRow, SETTINGS_COL, 1, 3).setValues([[d[0], value, d[2]]]);
    nextRow++;
  });

  // Dashboard groups list
  const groupCells = chart.getRange(2, GROUPS_COL, SETTINGS_ROWS - 1, 1).getValues();
  const hasGroups = groupCells.some(function (r) { return String(r[0]).trim() !== ''; });
  if (!hasGroups) {
    const list = (legacy && legacy.groups.length) ? legacy.groups : DEFAULT_DASHBOARD_GROUPS;
    if (legacy && legacy.groups.length) migrated = true;
    chart.getRange(2, GROUPS_COL, list.length, 1).setValues(list.map(function (g) { return [g]; }));
  }

  // Widths: A:F chart, G gap, H:J settings, K gap, L groups
  const widths = [100, 260, 100, 120, 130, 140, 30, 230, 110, 430, 30, 160];
  widths.forEach(function (w, i) { chart.setColumnWidth(i + 1, w); });
  chart.setFrozenRows(1);
  styleSettingsTable_(chart);   // see Styling.gs

  return migrated;
}

/** Reads the old separate "Settings" sheet (key/value in A:B, groups in E), if present. */
function readLegacySettingsSheet_(ss) {
  const sh = findSheetCI_(ss, 'Settings');
  if (!sh || sh.getLastRow() < 2) return null;
  const out = { values: {}, groups: [] };
  sh.getRange(1, 1, Math.min(sh.getLastRow(), SETTINGS_ROWS), 2).getValues().forEach(function (r) {
    const k = String(r[0]).trim().toLowerCase();
    if (k && k !== 'setting') out.values[k] = String(r[1]).trim();
  });
  if (sh.getMaxColumns() >= 5) {
    out.groups = sh.getRange(2, 5, Math.min(sh.getLastRow(), SETTINGS_ROWS), 1).getValues()
      .map(function (r) { return String(r[0]).trim(); })
      .filter(function (g) { return g !== ''; });
  }
  return out;
}

// ---------------------------------------------------------------------------
// CHART OF ACCOUNTS
// ---------------------------------------------------------------------------

function getChartSheet_(ss) {
  return findSheetCI_(ss || SpreadsheetApp.getActiveSpreadsheet(), CHART_SHEET_NAME);
}

/** Number of accounts in column A below the header. Ignores the Settings table and groups list. */
function countAccounts_(chart) {
  const rows = chart.getMaxRows() - 1;
  if (rows < 1) return 0;
  return chart.getRange(2, 1, rows, 1).getValues()
    .filter(function (r) { return String(r[0]).trim() !== ''; }).length;
}

/** Writes the generic starter chart. Only called when no accounts exist. */
function loadStarterChart_(ss) {
  let chart = getChartSheet_(ss);
  if (!chart) chart = ss.insertSheet(CHART_SHEET_NAME);   // this activates the new sheet

  chart.getRange(1, 1, 1, CHART_COLS).setValues([CHART_HEADERS])
    .setFontWeight('bold').setBackground('#334155').setFontColor('#ffffff');
  chart.getRange(2, 1, STARTER_CHART.length, 1).setNumberFormat('@');   // codes stay text
  chart.getRange(2, 1, STARTER_CHART.length, CHART_COLS).setValues(STARTER_CHART);
  chart.getRange(2, 4, STARTER_CHART.length, 1).setNumberFormat('#,##0.00');
  styleStarterChart_(chart, STARTER_CHART.length);   // see Styling.gs
  return chart;
}

/**
 * Reads the Chart of Accounts table (columns A:F) into
 * { code: {name, type, budget, cashFlowClass, dashboardGroup} }.
 * Columns are found by header text. The type column may be headed "Category" or "Type".
 * Plural values ("Assets") are normalised to singular.
 */
function readChartOfAccounts() {
  const sh = getChartSheet_();
  if (!sh) throw new Error('Sheet "' + CHART_SHEET_NAME + '" not found. Run Better Budget > Set Up / Upgrade Tool.');

  const lastRow = sh.getLastRow();
  if (lastRow < 1) return {};
  const width = Math.min(CHART_COLS, sh.getMaxColumns());
  const data = sh.getRange(1, 1, lastRow, width).getValues();

  const hdr = data[0].map(function (h) { return String(h).trim().toLowerCase(); });
  const find = function (names, fallback) {
    for (let i = 0; i < names.length; i++) {
      const idx = hdr.indexOf(names[i]);
      if (idx !== -1) return idx;
    }
    return fallback;
  };
  const typeCol = find(['category', 'type'], 2);
  const budgetCol = find([CHART_BUDGET_HEADER.toLowerCase()], -1);
  const cfCol = find([CHART_CASHFLOW_HEADER.toLowerCase()], -1);
  const groupCol = find([CHART_GROUP_HEADER.toLowerCase()], -1);

  const out = {};
  data.forEach(function (row, i) {
    if (i === 0 || row[0] === '' || row[0] === null) return;
    const budgetRaw = budgetCol === -1 ? '' : row[budgetCol];
    const budgetNum = Number(budgetRaw);
    out[String(row[0]).trim()] = {
      name: row[1],
      type: normalizeCategory_(row[typeCol]),
      budget: (budgetRaw !== '' && !isNaN(budgetNum) && budgetNum > 0) ? budgetNum : null,
      cashFlowClass: cfCol === -1 ? '' : String(row[cfCol]).trim(),
      dashboardGroup: groupCol === -1 ? '' : String(row[groupCol]).trim(),
    };
  });
  return out;
}

function normalizeCategory_(t) {
  const s = String(t === null || t === undefined ? '' : t).trim();
  const l = s.toLowerCase();
  if (l.indexOf('asset') === 0) return 'Asset';
  if (l.indexOf('liabilit') === 0) return 'Liability';
  if (l.indexOf('equity') === 0) return 'Equity';
  if (l === 'revenue' || l === 'income') return 'Revenue';
  if (l.indexOf('expense') === 0) return 'Expense';
  return s;
}

/** Adds missing columns (within A:F), dropdowns, and fills BLANK cells with sensible defaults. */
function upgradeChartOfAccounts_(ss, chart) {
  const addedColumns = [];

  const header = chart.getRange(1, 1, 1, CHART_COLS).getValues()[0].map(function (h) { return String(h).trim(); });

  const ensureCol = function (title) {
    let idx = -1;
    for (let i = 0; i < header.length; i++) {
      if (header[i].toLowerCase() === title.toLowerCase()) { idx = i; break; }
    }
    if (idx !== -1) return idx + 1;

    for (let i = 0; i < header.length; i++) {
      if (header[i] === '') { idx = i; break; }
    }
    if (idx === -1) {
      throw new Error('The Chart of Accounts table (columns A-F) has no free column for "' + title +
        '". Move or delete a column inside A-F and try again.');
    }
    header[idx] = title;
    chart.getRange(1, idx + 1).setValue(title);
    chart.getRange(1, 1).copyTo(chart.getRange(1, idx + 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    addedColumns.push(title);
    return idx + 1;
  };

  ensureCol(CHART_BUDGET_HEADER);
  const cfCol = ensureCol(CHART_CASHFLOW_HEADER);
  const groupCol = ensureCol(CHART_GROUP_HEADER);

  // Category column: "Category" or "Type", else the 3rd column
  let typeCol = 3;
  for (let i = 0; i < header.length; i++) {
    const l = header[i].toLowerCase();
    if (l === 'category' || l === 'type') { typeCol = i + 1; break; }
  }

  const rowsBelow = Math.max(chart.getMaxRows() - 1, 1);

  // Dropdowns (warning-only, so existing unusual values are not rejected)
  const listRule = function (values) {
    return SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(true).build();
  };
  chart.getRange(2, typeCol, rowsBelow, 1).setDataValidation(listRule(ACCOUNT_CATEGORIES));
  chart.getRange(2, cfCol, rowsBelow, 1).setDataValidation(listRule(CASH_FLOW_CLASSES));

  const groupRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(chart.getRange(2, GROUPS_COL, SETTINGS_ROWS - 1, 1), true)
    .setAllowInvalid(true).build();
  chart.getRange(2, groupCol, rowsBelow, 1).setDataValidation(groupRule);

  // Pre-fill BLANK cells
  const lastRow = chart.getLastRow();
  let cashFlowFilled = 0, groupFilled = 0;
  if (lastRow >= 2) {
    const n = lastRow - 1;
    const data = chart.getRange(2, 1, n, CHART_COLS).getValues();

    const cfOut = [];
    const groupOut = [];
    data.forEach(function (row) {
      const code = String(row[0]).trim();
      const type = normalizeCategory_(row[typeCol - 1]);
      let cf = String(row[cfCol - 1]).trim();
      let grp = String(row[groupCol - 1]).trim();

      if (code !== '' && cf === '') {
        const guess = legacyCashFlowClass_(code, type);
        if (guess) { cf = guess; cashFlowFilled++; }
      }
      if (code !== '' && grp === '' && type === 'Expense') {
        grp = legacyDashboardGroup_(String(row[1]));
        groupFilled++;
      }
      cfOut.push([cf]);
      groupOut.push([grp]);
    });
    chart.getRange(2, cfCol, n, 1).setValues(cfOut);
    chart.getRange(2, groupCol, n, 1).setValues(groupOut);
  }

  return { addedColumns: addedColumns, cashFlowFilled: cashFlowFilled, groupFilled: groupFilled };
}

// ---------------------------------------------------------------------------
// LEGACY RULES (the logic that used to be hard-coded in renderCashFlow and
// generateDashboard). Used ONLY to pre-fill blank cells during an upgrade.
// ---------------------------------------------------------------------------

function legacyCashFlowClass_(code, type) {
  if (type !== 'Asset' && type !== 'Liability' && type !== 'Equity') return '';
  if (code === getRetainedEarningsCode()) return '';

  const cashCodes = ['1010', '1020', '1021', '1022', '1030', '1300'];
  if (cashCodes.indexOf(code) !== -1) return 'Cash';
  if (code === '1500' || code === '2500' || code.indexOf('20') === 0) return 'Operating';
  if (code.indexOf('11') === 0 || code.indexOf('12') === 0) return 'Investing';
  if (code.indexOf('21') === 0 || code.indexOf('22') === 0 || code.indexOf('30') === 0) return 'Financing';
  return '';
}

function legacyDashboardGroup_(name, groupsOpt) {
  const groups = groupsOpt || getDashboardGroups();   // pass the list in to avoid re-reading the sheet
  const pick = function (preferred) { return groups.indexOf(preferred) !== -1 ? preferred : 'Other'; };
  if (/rent|mortgage|housing/i.test(name)) return groups.indexOf('Rent') !== -1 ? 'Rent' : pick('Housing');
  if (/grocer/i.test(name)) return pick('Groceries');
  if (/dining/i.test(name)) return pick('Dining Out');
  if (/utilit/i.test(name)) return pick('Utilities');
  if (/transport|gas/i.test(name)) return pick('Transport');
  return 'Other';
}

// ---------------------------------------------------------------------------
// SMALL HELPERS
// ---------------------------------------------------------------------------

/** Case-insensitive sheet lookup ("Chart Of Accounts" == "Chart of Accounts"). */
function findSheetCI_(ss, name) {
  const target = String(name).trim().toLowerCase();
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().trim().toLowerCase() === target) return sheets[i];
  }
  return null;
}

/** insertSheet() activates the new sheet; put the user back where they were. */
function restoreActiveSheet_(ss, sheet) {
  try { ss.setActiveSheet(sheet); } catch (e) { /* sheet may have been removed */ }
}


// ============================================================================
// SECTION 3 OF 10: STYLING
// ============================================================================

/**
 * ============================================================================
 * BETTER BUDGET - SHARED STYLING
 * ============================================================================
 * NEW file. One look for everything the tool generates for a new user, using
 * the same slate palette as the Dashboard.
 *
 * Used by:
 *   Settings and Setup.gs  -> styleStarterChart_(), styleSettingsTable_()
 *   First Run.gs           -> styleHeader_() and the BB_* colours
 *
 * Existing users' own Chart of Accounts keeps its formatting: only the starter
 * chart (new workbooks) and the Settings table are styled.
 */

const BB_NAVY_ = '#1e293b';
const BB_SLATE_ = '#334155';
const BB_LIGHT_ = '#f1f5f9';
const BB_BORDER_ = '#cbd5e1';
const BB_INPUT_ = '#eff6ff';
const BB_MUTED_ = '#64748b';

const BB_CATEGORY_COLORS_ = {
  Asset:     { bg: '#dbeafe', fg: '#1e3a8a' },
  Liability: { bg: '#ffedd5', fg: '#9a3412' },
  Equity:    { bg: '#ede9fe', fg: '#5b21b6' },
  Revenue:   { bg: '#dcfce7', fg: '#166534' },
  Expense:   { bg: '#fee2e2', fg: '#991b1b' },
};

/** Dark table header: white bold text, centred. */
function styleHeader_(range) {
  range.setFontWeight('bold').setFontColor('#ffffff').setBackground(BB_SLATE_)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
}

/** Thin grey grid lines around and between the rows of a table. */
function styleTableLines_(range) {
  range.setBorder(true, true, true, true, false, true, BB_BORDER_, SpreadsheetApp.BorderStyle.SOLID);
}

// ---------------------------------------------------------------------------
// CHART OF ACCOUNTS (starter chart only)
// ---------------------------------------------------------------------------

function styleStarterChart_(chart, n) {
  styleHeader_(chart.getRange(1, 1, 1, CHART_COLS));
  chart.setRowHeight(1, 34);
  chart.setFrozenRows(1);
  chart.setTabColor('#475569');
  chart.getRange(1, 4).setNote('Type a monthly amount next to any expense you want to track. ' +
    'The reports and the weekly email compare your spending with it.');

  const body = chart.getRange(2, 1, n, CHART_COLS);
  body.setVerticalAlignment('middle');
  styleTableLines_(body);

  chart.getRange(2, 1, n, 1).setHorizontalAlignment('center').setFontColor(BB_MUTED_);
  chart.getRange(2, 3, n, 1).setHorizontalAlignment('center');
  chart.getRange(2, 4, n, 1).setHorizontalAlignment('right').setNumberFormat('#,##0.00');
  chart.getRange(2, 5, n, 1).setHorizontalAlignment('center');

  // Budgets only apply to expenses: grey out the other rows' budget cells
  const types = chart.getRange(2, 3, n, 1).getValues();
  chart.getRange(2, 4, n, 1).setBackgrounds(types.map(function (r) {
    return [normalizeCategory_(r[0]) === 'Expense' ? '#ffffff' : BB_LIGHT_];
  }));

  // A heavier line where the type of account changes (assets | liabilities | ...)
  for (let i = 1; i < types.length; i++) {
    if (types[i][0] !== types[i - 1][0]) {
      chart.getRange(2 + i, 1, 1, CHART_COLS)
        .setBorder(true, null, null, null, null, null, BB_MUTED_, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  }

  applyCategoryColors_(chart, 3);
}

/**
 * Colours the Category cells (Asset = blue, Expense = red and so on). Uses
 * conditional formatting, so it keeps working when accounts are added or changed.
 * Safe to run repeatedly: it replaces its own earlier rules.
 */
function applyCategoryColors_(chart, typeCol) {
  const rows = Math.max(chart.getMaxRows() - 1, 1);
  const range = chart.getRange(2, typeCol, rows, 1);

  const kept = chart.getConditionalFormatRules().filter(function (rule) {
    const cond = rule.getBooleanCondition();
    if (!cond || cond.getCriteriaType() !== SpreadsheetApp.BooleanCriteria.TEXT_EQUAL_TO) return true;
    const values = cond.getCriteriaValues();
    const ranges = rule.getRanges();
    const ours = values.length && BB_CATEGORY_COLORS_[values[0]] && ranges.length &&
      ranges[0].getColumn() === typeCol;
    return !ours;
  });

  Object.keys(BB_CATEGORY_COLORS_).forEach(function (cat) {
    const c = BB_CATEGORY_COLORS_[cat];
    kept.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(cat)
      .setBackground(c.bg)
      .setFontColor(c.fg)
      .setBold(true)
      .setRanges([range])
      .build());
  });
  chart.setConditionalFormatRules(kept);
}

// ---------------------------------------------------------------------------
// SETTINGS TABLE (columns H-J) AND DASHBOARD GROUPS (column L)
// ---------------------------------------------------------------------------

function styleSettingsTable_(chart) {
  const n = SETTING_DEFAULTS.length;
  styleHeader_(chart.getRange(1, SETTINGS_COL, 1, 3));
  styleHeader_(chart.getRange(1, GROUPS_COL));

  const body = chart.getRange(2, SETTINGS_COL, n, 3);
  body.setVerticalAlignment('middle');
  styleTableLines_(body);

  chart.getRange(2, SETTINGS_COL, n, 1).setFontWeight('bold');
  // Values are the cells people edit: tint them like input boxes
  chart.getRange(2, SETTINGS_COL + 1, n, 1)
    .setBackground(BB_INPUT_).setFontWeight('bold').setHorizontalAlignment('center');
  chart.getRange(2, SETTINGS_COL + 2, n, 1)
    .setFontColor(BB_MUTED_).setFontSize(9).setWrap(true);

  // Dashboard groups: the filled cells plus a few spare ones to invite additions
  const spare = 4;
  const groupRows = Math.min(getDashboardGroups().length + spare, SETTINGS_ROWS - 1);
  const groups = chart.getRange(2, GROUPS_COL, groupRows, 1);
  groups.setBackground(BB_INPUT_).setVerticalAlignment('middle');
  styleTableLines_(groups);
}


// ============================================================================
// SECTION 4 OF 10: FIRST RUN
// ============================================================================

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


// ============================================================================
// SECTION 5 OF 10: RECORD SIDEBAR
// ============================================================================

/**
 * ============================================================================
 * BETTER BUDGET - TRANSACTION ENTRY SIDEBAR
 * ============================================================================
 * REPLACES "Record sidebar.gs" entirely. Use together with the new
 * SidebarHTML.html.
 *
 * Users describe what happened in plain English ("I spent money", "I received
 * money", "I moved money", "I got money back", "opening balance") and this file
 * turns it into the correct double-entry lines. The ledger underneath is
 * unchanged: every simple entry writes two rows (one debit, one credit).
 * An "Advanced" mode keeps the old one-line debit/credit entry.
 *
 * Depends on:
 *   Settings and Setup.gs  -> readChartOfAccounts, getCurrencySymbol,
 *                             getPrimaryBankCode, getOpeningBalanceCode
 *   Monthly Reports.gs     -> findHeaderRow
 *
 * Removed from the old version: getChartOfAccounts(), processForm()
 */

const ENTRY_MONTHS_ = ['January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'];

const ENTRY_CATEGORIES_ = ['Asset', 'Liability', 'Equity', 'Revenue', 'Expense'];

// Which account categories each entry type accepts. a = first dropdown, b = second.
const ENTRY_RULES_ = {
  spent:    { a: ['Expense'],            b: ['Asset', 'Liability'] },
  received: { a: ['Revenue'],            b: ['Asset'] },
  moved:    { a: ['Asset', 'Liability'], b: ['Asset', 'Liability'] },
  refund:   { a: ['Expense'],            b: ['Asset', 'Liability'] },
  opening:  { a: ['Asset', 'Liability'], b: null },
  advanced: { a: null,                   b: null },
};

// ---------------------------------------------------------------------------
// SIDEBAR
// ---------------------------------------------------------------------------

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('SidebarHTML')
    .setTitle('Record a Transaction')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Called by the sidebar when it opens and whenever it regains focus (the user
 * may have switched tabs). Describes the active tab and the available accounts.
 */
function getSidebarContext() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const target = getLedgerTarget_(sheet);
  if (!target.ok) {
    return { ok: false, reason: target.reason + ' Switch to a monthly ledger tab to record entries.' };
  }

  let chart;
  try {
    chart = readChartOfAccounts();
  } catch (e) {
    return { ok: false, reason: e.message };
  }

  const accounts = [];
  Object.keys(chart).forEach(function (code) {
    const a = chart[code];
    if (ENTRY_CATEGORIES_.indexOf(a.type) === -1) return;
    accounts.push({ code: code, name: String(a.name), type: a.type });
  });
  if (!accounts.length) {
    return { ok: false, reason: 'Your Chart of Accounts has no accounts yet. Run Better Budget > Set Up / Upgrade Tool.' };
  }

  const bank = getPrimaryBankCode();
  const hasMonth = target.monthIdx !== -1 && target.year > 0;

  return {
    ok: true,
    tabName: sheet.getName(),
    monthKey: hasMonth ? target.year + '-' + ('0' + (target.monthIdx + 1)).slice(-2) : '',
    currency: getCurrencySymbol(),
    primaryBank: chart[bank] ? bank : '',
    accounts: accounts,
    outOfBalance: entryOutOfBalance_(sheet, target.headerRow),
  };
}

/**
 * Saves one entry from the sidebar. Always returns { ok, message | error, outOfBalance }.
 * Expected problems (bad input, wrong tab) come back as { ok: false, error }.
 */
function saveEntry(payload) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) {
    return { ok: false, error: 'The sheet is busy. Please try again in a moment.' };
  }
  try {
    return saveEntryLocked_(payload || {});
  } finally {
    lock.releaseLock();
  }
}

function saveEntryLocked_(p) {
  const fail = function (msg) { return { ok: false, error: msg }; };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // 1. Must be a real ledger tab (checked again here: the user may have switched tabs)
  const target = getLedgerTarget_(sheet);
  if (!target.ok) return fail(target.reason + ' Switch to a monthly ledger tab and try again.');

  let chart;
  try { chart = readChartOfAccounts(); } catch (e) { return fail(e.message); }

  // 2. Date, and it must belong to this tab's month
  const parts = String(p.date || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!parts) return fail('Please choose a valid date.');
  const y = Number(parts[1]), m = Number(parts[2]) - 1, d = Number(parts[3]);
  const date = new Date(y, m, d);
  if (date.getFullYear() !== y || date.getMonth() !== m || date.getDate() !== d) {
    return fail('Please choose a valid date.');
  }
  if (target.monthIdx !== -1 && target.year > 0 && (m !== target.monthIdx || y !== target.year)) {
    return fail('This tab is for ' + ENTRY_MONTHS_[target.monthIdx] + ' ' + target.year +
      ', but the date is in ' + ENTRY_MONTHS_[m] + ' ' + y +
      '. Change the date, or switch to that month\'s tab.');
  }

  // 3. Amount(s) and entry type
  const type = String(p.type || '');
  const amount = type === 'advanced' ? 0 : entryParseAmount_(p.amount);
  if (type !== 'advanced' && amount === null) return fail('Please enter an amount greater than zero.');

  const built = buildEntry_(type, p, chart, amount,
    String(p.description || '').trim(), getCurrencySymbol(), getOpeningBalanceCode());
  if (built.error) return fail(built.error);

  // 4. Rows to write: Date | Account | Description | Debit | Credit
  const rows = built.lines.map(function (l) {
    const code = /^[1-9]\d{0,14}$/.test(l.code) ? Number(l.code) : l.code;
    return [date, code, built.desc, l.debit > 0 ? l.debit : '', l.credit > 0 ? l.credit : ''];
  });

  // 5. Write below the last used row of the table (columns A-E)
  const first = target.headerRow + 1;
  let lastUsed = target.headerRow;
  const maxRows = sheet.getMaxRows();
  if (maxRows >= first) {
    const vals = sheet.getRange(first, 1, maxRows - first + 1, 5).getValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      if (vals[i].some(function (c) { return c !== '' && c !== null; })) { lastUsed = first + i; break; }
    }
  }
  const startRow = lastUsed + 1;
  const lastNeeded = startRow + rows.length - 1;
  if (lastNeeded > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), lastNeeded - sheet.getMaxRows());
  }
  sheet.getRange(startRow, 1, rows.length, 5).setValues(rows);
  SpreadsheetApp.flush();

  return {
    ok: true,
    message: 'Saved to ' + sheet.getName() + ': ' + built.summary,
    outOfBalance: entryOutOfBalance_(sheet, target.headerRow),
  };
}

// ---------------------------------------------------------------------------
// ENTRY LOGIC (turns plain-English entries into debit/credit lines)
// ---------------------------------------------------------------------------

/**
 * Returns { lines: [{code, debit, credit}], desc, summary } or { error }.
 * Pure function: no sheet access.
 */
function buildEntry_(type, p, chart, amount, desc, cur, openingCode) {
  const rule = ENTRY_RULES_[type];
  if (!rule) return { error: 'Unknown entry type.' };

  const codeA = String(p.accountA || '').trim();
  const codeB = String(p.accountB || '').trim();
  const a = chart[codeA];
  if (!a) return { error: 'Please choose an account.' };
  if (rule.a && rule.a.indexOf(a.type) === -1) {
    return { error: '"' + a.name + '" is not the right kind of account for this entry.' };
  }

  // Advanced: a single debit or credit line
  if (type === 'advanced') {
    const dr = p.debit !== '' && p.debit != null ? entryParseAmount_(p.debit) : 0;
    const cr = p.credit !== '' && p.credit != null ? entryParseAmount_(p.credit) : 0;
    if (dr === null || cr === null) return { error: 'Amounts must be greater than zero.' };
    if ((dr > 0) === (cr > 0)) return { error: 'Enter either a Debit or a Credit amount, not both or neither.' };
    const amt = dr > 0 ? dr : cr;
    return {
      lines: [{ code: codeA, debit: dr, credit: cr }],
      desc: desc || a.name,
      summary: entryMoney_(amt, cur) + ' ' + (dr > 0 ? 'debit' : 'credit') + ' to ' + a.name +
        '. Remember to record the matching line for the other side.',
    };
  }

  const money = entryMoney_(amount, cur);

  // Opening balance: one account against the Opening Balance equity account
  if (type === 'opening') {
    const eq = chart[openingCode];
    if (!eq) {
      return { error: 'The opening balance account code in Settings ("' + openingCode +
        '") is not in your Chart of Accounts.' };
    }
    const isAsset = a.type === 'Asset';
    return {
      lines: isAsset
        ? [{ code: codeA, debit: amount, credit: 0 }, { code: openingCode, debit: 0, credit: amount }]
        : [{ code: openingCode, debit: amount, credit: 0 }, { code: codeA, debit: 0, credit: amount }],
      desc: desc || 'Opening balance',
      summary: 'opening balance of ' + money + ' for ' + a.name + '.',
    };
  }

  // Everything else involves two accounts
  const b = chart[codeB];
  if (!b) return { error: 'Please choose both accounts.' };
  if (rule.b && rule.b.indexOf(b.type) === -1) {
    return { error: '"' + b.name + '" is not the right kind of account for this entry.' };
  }
  if (codeA === codeB) return { error: 'Please choose two different accounts.' };

  if (type === 'spent') {            // debit what it was for, credit where the money came from
    return {
      lines: [{ code: codeA, debit: amount, credit: 0 }, { code: codeB, debit: 0, credit: amount }],
      desc: desc || a.name,
      summary: money + ' for ' + a.name + ', paid from ' + b.name + '.',
    };
  }
  if (type === 'received') {         // debit where it went, credit the income source
    return {
      lines: [{ code: codeB, debit: amount, credit: 0 }, { code: codeA, debit: 0, credit: amount }],
      desc: desc || a.name,
      summary: money + ' received as ' + a.name + ', deposited into ' + b.name + '.',
    };
  }
  if (type === 'moved') {            // debit the destination, credit the source
    return {
      lines: [{ code: codeB, debit: amount, credit: 0 }, { code: codeA, debit: 0, credit: amount }],
      desc: desc || ('Transfer from ' + a.name + ' to ' + b.name),
      summary: 'moved ' + money + ' from ' + a.name + ' to ' + b.name + '.',
    };
  }
  if (type === 'refund') {           // credit the original expense, debit where the money landed
    return {
      lines: [{ code: codeB, debit: amount, credit: 0 }, { code: codeA, debit: 0, credit: amount }],
      desc: desc || ('Refund - ' + a.name),
      summary: money + ' back for ' + a.name + ', received into ' + b.name + '.',
    };
  }
  return { error: 'Unknown entry type.' };
}

/** Positive number rounded to cents, or null. */
function entryParseAmount_(v) {
  const n = Number(v);
  if (!isFinite(n) || n <= 0 || n > 1e12) return null;
  const r = Math.round((n + Number.EPSILON) * 100) / 100;
  return r > 0 ? r : null;
}

function entryMoney_(n, cur) {
  return cur + Math.abs(n).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

// ---------------------------------------------------------------------------
// LEDGER HELPERS
// ---------------------------------------------------------------------------

/**
 * Decides whether a sheet is a monthly ledger and where its table starts.
 * Unlike validateLedgerSheet(), an EMPTY ledger (a brand-new month) is valid.
 */
function getLedgerTarget_(sheet) {
  const name = sheet.getName();
  const lower = name.trim().toLowerCase();

  if (name === 'Dashboard' || lower === 'chart of accounts' || lower === 'settings' || lower === 'start here') {
    return { ok: false, reason: '"' + name + '" is not a ledger tab.' };
  }
  if (name.indexOf('Financials - ') === 0) {
    return { ok: false, reason: '"' + name + '" is a report tab, not a ledger.' };
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { ok: false, reason: '"' + name + '" is empty.' };

  const top = sheet.getRange(1, 1, Math.min(lastRow, 15), lastCol).getValues();
  const idx = findHeaderRow(top);
  if (idx === -1) {
    return { ok: false, reason: 'Could not find the ledger headers (Date, Account, Debit, Credit) on "' + name + '".' };
  }

  const monthName = String(sheet.getRange('B3').getValue()).trim().toLowerCase();
  const year = Number(sheet.getRange('B4').getValue());
  const monthIdx = ENTRY_MONTHS_.findIndex(function (mm) { return mm.toLowerCase() === monthName; });

  return {
    ok: true,
    name: name,
    headerRow: idx + 1,
    monthIdx: monthIdx,
    year: isNaN(year) ? 0 : year,
  };
}

/** Total debits minus total credits on the tab (should be 0). */
function entryOutOfBalance_(sheet, headerRow) {
  const last = sheet.getLastRow();
  if (last <= headerRow) return 0;
  const vals = sheet.getRange(headerRow + 1, 4, last - headerRow, 2).getValues();
  let dr = 0, cr = 0;
  vals.forEach(function (r) {
    dr += Number(r[0]) || 0;
    cr += Number(r[1]) || 0;
  });
  return Math.round((dr - cr) * 100) / 100;
}


// ============================================================================
// SECTION 6 OF 10: MONTHLY REPORTS
// ============================================================================

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


// ============================================================================
// SECTION 7 OF 10: NEW MONTH
// ============================================================================

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


// ============================================================================
// SECTION 8 OF 10: DASHBOARD
// ============================================================================

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


// ============================================================================
// SECTION 9 OF 10: WEEKLY REPORT
// ============================================================================

/**
 * ============================================================================
 * BETTER BUDGET - WEEKLY BUDGET REPORT
 * ============================================================================
 * Emails a weekly summary of expense spending vs. monthly budgets.
 *
 * Uses the shared pieces of the tool:
 *   - readChartOfAccounts(), getCurrencySymbol(), getSetting()  (Settings and Setup.gs)
 *   - validateLedgerSheet() and findHeaderRow()                 (Monthly Reports.gs)
 *
 * Setup for the user: type a Monthly Budget next to the expense accounts they
 * want tracked (Chart of Accounts, column D), then use
 * Better Budget > Weekly Email > Turn on. Optional: set "Report email" in
 * the Settings table to send the report somewhere other than their own account.
 *
 * All names here start with "weekly"/"Weekly"/"WEEKLY" so they can't clash
 * with names in other files.
 */

const WEEKLY_CONFIG = {
  PACE_TOLERANCE: 10,                 // % points ahead of month progress before "Ahead of pace"
  PROJECTION_MIN_DAY: 7,              // don't judge pace before this day of the month
  TRIGGER_DAY: ScriptApp.WeekDay.MONDAY,
  TRIGGER_HOUR: 8,
};

const WEEKLY_MONTHS = ['January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'];

// ---------------------------------------------------------------------------
// MAIN
// ---------------------------------------------------------------------------

function sendWeeklyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();

  // "Today" as a plain local date, built from the spreadsheet's time zone
  const p = Utilities.formatDate(new Date(), tz, 'yyyy-M-d').split('-').map(Number);
  const today = new Date(p[0], p[1] - 1, p[2]);
  const endOfToday = new Date(p[0], p[1] - 1, p[2], 23, 59, 59);
  const weekStart = new Date(p[0], p[1] - 1, p[2] - 6);   // last 7 days incl. today
  const monthStart = new Date(p[0], p[1] - 1, 1);
  const daysInMonth = new Date(p[0], p[1], 0).getDate();
  const dayOfMonth = p[2];
  const monthPct = (dayOfMonth / daysInMonth) * 100;

  const accounts = weeklyReadAccounts_();

  // This month's ledger, plus last month's if the 7-day window crosses into it
  const monthRows = weeklyReadLedger_(weeklyFindLedgerSheet_(ss, today));
  let weekRows = monthRows;
  if (weekStart < monthStart) {
    weekRows = weekRows.concat(weeklyReadLedger_(weeklyFindLedgerSheet_(ss, weekStart)));
  }

  // Aggregate expense spend per account (expenses are net debits)
  const spend = {};
  const getSpend = function (code) {
    return spend[code] || (spend[code] = { month: 0, week: 0 });
  };

  monthRows.forEach(function (r) {
    if (!weeklyIsExpense_(accounts[r.code])) return;
    getSpend(r.code).month += r.debit - r.credit;
  });

  const bigTx = [];
  weekRows.forEach(function (r) {
    if (!weeklyIsExpense_(accounts[r.code])) return;
    if (!r.date || r.date < weekStart || r.date > endOfToday) return;
    const net = r.debit - r.credit;
    getSpend(r.code).week += net;
    if (net > 0) bigTx.push({ date: r.date, desc: r.desc, name: accounts[r.code].name, amount: net });
  });

  // Split into budgeted / unbudgeted
  const budgeted = [];
  const unbudgeted = [];
  Object.keys(spend).forEach(function (code) {
    const a = accounts[code];
    const s = spend[code];
    if (a.budget === null) {
      if (Math.abs(s.month) > 0.005 || Math.abs(s.week) > 0.005) {
        unbudgeted.push({ name: a.name, month: s.month, week: s.week });
      }
      return;
    }
    budgeted.push(weeklyAssess_(code, a, s, dayOfMonth, daysInMonth, monthPct));
  });

  // Budgeted accounts with no spending yet still belong in the report
  Object.keys(accounts).forEach(function (code) {
    const a = accounts[code];
    if (weeklyIsExpense_(a) && a.budget !== null && !spend[code]) {
      budgeted.push(weeklyAssess_(code, a, { month: 0, week: 0 }, dayOfMonth, daysInMonth, monthPct));
    }
  });

  const rank = { 'OVER BUDGET': 0, 'On pace to exceed': 1, 'Ahead of pace': 2, 'On track': 3 };
  budgeted.sort(function (a, b) { return rank[a.status] - rank[b.status] || b.pctUsed - a.pctUsed; });
  unbudgeted.sort(function (a, b) { return b.month - a.month; });
  bigTx.sort(function (a, b) { return b.amount - a.amount; });

  const html = weeklyBuildHtml_(budgeted, unbudgeted, bigTx.slice(0, 5), {
    ssName: ss.getName(), noBudgets: budgeted.length === 0, ssUrl: ss.getUrl(), weekStart: weekStart, today: today,
    dayOfMonth: dayOfMonth, daysInMonth: daysInMonth, monthPct: monthPct,
  });

  const flagged = budgeted.filter(function (r) { return r.status !== 'On track'; }).length;
  const subject = 'Weekly budget report - ' + WEEKLY_MONTHS[today.getMonth()].slice(0, 3) + ' ' +
    today.getDate() + ', ' + today.getFullYear() +
    (flagged ? ' (' + flagged + ' need attention)' : ' (all on track)');

  MailApp.sendEmail({
    to: weeklyRecipient_(),
    subject: subject,
    htmlBody: html,
  });
}

/** Menu: sends the report immediately. */
function sendWeeklyReportNow() {
  const ui = SpreadsheetApp.getUi();
  try {
    sendWeeklyReport();
  } catch (e) {
    ui.alert('Could not send the weekly report', String(e.message || e), ui.ButtonSet.OK);
    return;
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('Weekly report emailed to ' + weeklyRecipient_() + '.', 'Done', 4);
}

/** Menu: Weekly Email > Turn on. Safe to repeat: replaces any earlier schedule. */
function turnOnWeeklyEmail() {
  const ui = SpreadsheetApp.getUi();
  createWeeklyReportTrigger();
  ui.alert('Weekly email is on',
    'A report will be emailed to ' + weeklyRecipient_() + ' every ' + weeklyDayName_() +
    ' around ' + weeklyHourLabel_() + '.\n\n' +
    'Tip: use Better Budget > Weekly Email > Send now to see what it looks like.\n' +
    'To send it to a different address, type it next to "Report email" in the Settings table on the Chart of Accounts.',
    ui.ButtonSet.OK);
}

/** Menu: Weekly Email > Turn off. */
function turnOffWeeklyEmail() {
  const ui = SpreadsheetApp.getUi();
  const removed = weeklyDeleteTriggers_();
  ui.alert(removed ? 'Weekly email is off' : 'Weekly email was already off',
    removed ? 'No more weekly reports will be sent. You can turn it back on at any time.'
            : 'There was no weekly schedule to remove.', ui.ButtonSet.OK);
}

/** Menu: Weekly Email > Status. */
function weeklyEmailStatus() {
  const ui = SpreadsheetApp.getUi();
  const on = weeklyHasTrigger_();
  ui.alert('Weekly email: ' + (on ? 'ON' : 'OFF'),
    on ? 'Sent to ' + weeklyRecipient_() + ' every ' + weeklyDayName_() + ' around ' + weeklyHourLabel_() + '.'
       : 'Use Better Budget > Weekly Email > Turn on to start receiving it.',
    ui.ButtonSet.OK);
}

/** Creates the weekly schedule (no dialogs, so it also works from the editor). */
function createWeeklyReportTrigger() {
  weeklyDeleteTriggers_();
  ScriptApp.newTrigger('sendWeeklyReport')
    .timeBased()
    .onWeekDay(WEEKLY_CONFIG.TRIGGER_DAY)
    .atHour(WEEKLY_CONFIG.TRIGGER_HOUR)
    .create();
}

function weeklyDeleteTriggers_() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendWeeklyReport') { ScriptApp.deleteTrigger(t); n++; }
  });
  return n;
}

function weeklyHasTrigger_() {
  return ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'sendWeeklyReport';
  });
}

function weeklyDayName_() {
  const d = String(WEEKLY_CONFIG.TRIGGER_DAY);   // "MONDAY"
  return d.charAt(0) + d.slice(1).toLowerCase();
}

function weeklyHourLabel_() {
  const h = WEEKLY_CONFIG.TRIGGER_HOUR;
  return (h % 12 === 0 ? 12 : h % 12) + (h < 12 ? ' am' : ' pm');
}

/** Report email setting if it looks like an address, otherwise the user's own account. */
function weeklyRecipient_() {
  let v = '';
  try { v = String(getSetting('Report email') || '').trim(); } catch (e) { v = ''; }
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v)) return v;
  return Session.getEffectiveUser().getEmail();
}

// ---------------------------------------------------------------------------
// ASSESSMENT
// ---------------------------------------------------------------------------

function weeklyAssess_(code, a, s, dayOfMonth, daysInMonth, monthPct) {
  const pctUsed = a.budget > 0 ? (s.month / a.budget) * 100 : 0;
  const projected = (s.month / dayOfMonth) * daysInMonth;
  let status = 'On track';
  if (s.month > a.budget) {
    status = 'OVER BUDGET';
  } else if (dayOfMonth >= WEEKLY_CONFIG.PROJECTION_MIN_DAY) {
    if (projected > a.budget) status = 'On pace to exceed';
    else if (pctUsed > monthPct + WEEKLY_CONFIG.PACE_TOLERANCE) status = 'Ahead of pace';
  }
  return {
    code: code, name: a.name, budget: a.budget, spent: s.month, week: s.week,
    pctUsed: pctUsed, projected: projected, remaining: a.budget - s.month, status: status,
  };
}

// ---------------------------------------------------------------------------
// DATA ACCESS
// ---------------------------------------------------------------------------

function weeklyIsExpense_(acc) {
  return !!acc && (acc.type === 'Expense' || acc.type === 'Expenses');
}

/** Chart of Accounts via the shared reader: { code: {name, type, budget|null} } */
function weeklyReadAccounts_() {
  const chart = readChartOfAccounts();
  const out = {};
  Object.keys(chart).forEach(function (code) {
    out[code] = { name: chart[code].name, type: chart[code].type, budget: chart[code].budget };
  });
  if (!Object.keys(out).length) {
    throw new Error('The Chart of Accounts is empty. Run Better Budget > Set Up / Upgrade Tool first.');
  }
  return out;
}

/**
 * Finds the General Ledger tab for the month containing `date`.
 * 1) tries a tab named "October 2026" style, 2) falls back to scanning every
 * tab's title block (B3 = month, B4 = year), the same cells your report engine uses.
 */
function weeklyFindLedgerSheet_(ss, date) {
  const month = WEEKLY_MONTHS[date.getMonth()];
  const year = String(date.getFullYear());

  const byName = ss.getSheetByName(month + ' ' + year);
  if (byName && validateLedgerSheet(byName).valid) return byName;

  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const sh = sheets[i];
    if (!validateLedgerSheet(sh).valid) continue;
    const m = String(sh.getRange('B3').getValue()).trim().toLowerCase();
    const y = String(sh.getRange('B4').getValue()).trim();
    if (m === month.toLowerCase() && y === year) return sh;
  }
  return null;
}

/** Reads ledger rows: Date | Account | Description | Debit | Credit */
function weeklyReadLedger_(sheet) {
  if (!sheet) return [];
  const raw = sheet.getDataRange().getValues();
  const hIdx = findHeaderRow(raw);
  const start = (hIdx !== -1 ? hIdx : 7) + 1;   // your engine defaults to header on row 8

  const out = [];
  for (let i = start; i < raw.length; i++) {
    const r = raw[i];
    if (r[1] === '' || r[1] === null) continue;
    out.push({
      date: weeklyParseDate_(r[0]),
      code: String(r[1]).trim(),
      desc: String(r[2] || ''),
      debit: Number(r[3]) || 0,
      credit: Number(r[4]) || 0,
    });
  }
  return out;
}

/** Handles real dates and "2026-10-11" text (as sent by your sidebar). */
function weeklyParseDate_(v) {
  if (v instanceof Date) return v;
  if (typeof v === 'string') {
    const m = v.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

// ---------------------------------------------------------------------------
// EMAIL
// ---------------------------------------------------------------------------

let WEEKLY_CUR_CACHE_ = null;
function weeklyCurrency_() {
  if (WEEKLY_CUR_CACHE_ === null) {
    try { WEEKLY_CUR_CACHE_ = getCurrencySymbol(); } catch (e) { WEEKLY_CUR_CACHE_ = '$'; }
  }
  return WEEKLY_CUR_CACHE_;
}

function weeklyMoney_(n) {
  const neg = n < 0;
  const s = Math.abs(n).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (neg ? '-' : '') + weeklyCurrency_() + s;
}

function weeklyEsc_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function weeklyShortDate_(d) {
  return WEEKLY_MONTHS[d.getMonth()].slice(0, 3) + ' ' + d.getDate();
}

function weeklyBuildHtml_(budgeted, unbudgeted, bigTx, meta) {
  const colors = {
    'OVER BUDGET': '#c62828', 'On pace to exceed': '#e65100',
    'Ahead of pace': '#b8860b', 'On track': '#2e7d32',
  };
  const sum = function (arr, k) { return arr.reduce(function (a, r) { return a + r[k]; }, 0); };

  const totalBudget = sum(budgeted, 'budget');
  const totalSpent = sum(budgeted, 'spent');
  const weekAll = sum(budgeted, 'week') + sum(unbudgeted, 'week');

  let h = '<div style="font-family:Arial,sans-serif;max-width:760px">';
  h += '<h2 style="margin-bottom:4px">Weekly budget report</h2>';
  h += '<p style="color:#555;margin-top:0">' + weeklyShortDate_(meta.weekStart) + ' - ' +
       weeklyShortDate_(meta.today) + ' &nbsp;|&nbsp; Day ' + meta.dayOfMonth + ' of ' +
       meta.daysInMonth + ' (' + meta.monthPct.toFixed(0) + '% of the month)</p>';

  h += '<p><b>Spent this week (all expenses):</b> ' + weeklyMoney_(weekAll) +
       '<br><b>Budgeted categories, month to date:</b> ' + weeklyMoney_(totalSpent) +
       ' of ' + weeklyMoney_(totalBudget) + ' (' +
       (totalBudget ? (totalSpent / totalBudget * 100).toFixed(0) : 0) + '%)</p>';

  if (meta.noBudgets) {
    h += '<p style="background:#fff8e1;padding:8px;border-radius:4px">No budgets are set yet. Type a monthly amount next to ' +
         'your expense accounts (Monthly Budget column on the Chart of Accounts) to get over-budget and pace alerts.</p>';
  }

  h += '<table style="border-collapse:collapse;width:100%;font-size:13px">';
  h += '<tr style="background:#eee;text-align:left"><th style="padding:6px">Category</th><th>Budget</th>' +
       '<th>Spent</th><th>% used</th><th>Left</th><th>This week</th><th>Projected</th><th>Status</th></tr>';
  budgeted.forEach(function (r) {
    h += '<tr style="border-bottom:1px solid #ddd">' +
      '<td style="padding:6px">' + weeklyEsc_(r.name) + '</td>' +
      '<td>' + weeklyMoney_(r.budget) + '</td>' +
      '<td>' + weeklyMoney_(r.spent) + '</td>' +
      '<td>' + r.pctUsed.toFixed(0) + '%</td>' +
      '<td>' + weeklyMoney_(r.remaining) + '</td>' +
      '<td>' + weeklyMoney_(r.week) + '</td>' +
      '<td>' + weeklyMoney_(r.projected) + '</td>' +
      '<td style="color:' + colors[r.status] + ';font-weight:bold">' + r.status + '</td></tr>';
  });
  h += '</table>';

  if (unbudgeted.length) {
    h += '<h3>Spending with no budget set</h3><ul>';
    unbudgeted.forEach(function (u) {
      h += '<li>' + weeklyEsc_(u.name) + ': ' + weeklyMoney_(u.month) + ' this month (' +
           weeklyMoney_(u.week) + ' this week)</li>';
    });
    h += '</ul>';
  }

  if (bigTx.length) {
    h += '<h3>Largest expenses this week</h3><ul>';
    bigTx.forEach(function (t) {
      h += '<li>' + weeklyShortDate_(t.date) + ' - ' + weeklyEsc_(t.desc || t.name) +
           ' (' + weeklyEsc_(t.name) + '): ' + weeklyMoney_(t.amount) + '</li>';
    });
    h += '</ul>';
  }

  h += '<p style="color:#777;font-size:11px">Projected = spend so far / days elapsed x days in month. ' +
       'Pace flags start on day ' + WEEKLY_CONFIG.PROJECTION_MIN_DAY + '. ' +
       '<a href="' + meta.ssUrl + '">Open ' + weeklyEsc_(meta.ssName || 'your sheet') + '</a></p></div>';
  return h;
}


// ============================================================================
// SECTION 10 OF 10: ABOUT
// ============================================================================

/**
 * ============================================================================
 * BETTER BUDGET - VERSION AND ABOUT
 * ============================================================================
 * NEW file. Bump BETTER_BUDGET_VERSION whenever you publish a new copy of the
 * template, so users (and you, when they send feedback) know what they have.
 */

const BETTER_BUDGET_VERSION = '1.0.0';

const BETTER_BUDGET_DISCLAIMER =
  'Better Budget is a personal budgeting aid. It is not professional financial, tax or legal advice, ' +
  'and it comes with no warranty. Check important figures yourself and keep your own backups ' +
  '(File > Make a copy).';

/** Menu: Better Budget > About. */
function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Better Budget v' + BETTER_BUDGET_VERSION,
    BETTER_BUDGET_DISCLAIMER + '\n\n' +
    'Your data stays in this spreadsheet. The only thing the tool ever sends out is the optional weekly ' +
    'email, and it goes to the address in your Settings (or your own Google account).',
    ui.ButtonSet.OK);
}
