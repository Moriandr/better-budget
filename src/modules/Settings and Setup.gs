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
