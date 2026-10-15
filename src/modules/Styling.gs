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
