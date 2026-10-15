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
