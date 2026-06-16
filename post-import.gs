/**
 * WH40k Point Efficiency Calculator — post-import processing for Google Sheets.
 *
 * Run this AFTER each manual "File → Import → Replace data" of the .xlsx.
 * It re-applies the things the xlsx→Sheets import drops or breaks:
 *   1. Native checkboxes on the boolean input cells.
 *   2. Frozen rows on the main tabs.
 *   3. A hidden record of where the checkboxes are, so the boxes survive being
 *      cleared with the Delete key (see "Why the hidden map" below).
 *
 * Install (once):
 *   Open the master Sheet → Extensions → Apps Script → paste this file → Save.
 *   Reload the Sheet; a "WH40k" menu appears in the menu bar.
 * Use (each release):
 *   Import as usual, then click  WH40k → Apply post-processing.
 *   (First run only: approve the one-time authorization prompt.)
 *
 * Why the hidden map:
 *   Google Sheets' Delete key removes a checkbox cell's *data-validation rule*,
 *   not just its value. So once a box is cleared there is nothing on the cell to
 *   tell us it was ever a checkbox — and on Input/Army Rules/Targets, where the
 *   checkbox cells are detected automatically, a cleared box is indistinguishable
 *   from any blank cell. To restore it on the next edit, onEdit needs an
 *   independent record of which cells are checkboxes. We store that record in a
 *   hidden sheet (CHECKBOX_MAP_SHEET) so it travels with the file when a user
 *   does File → Make a copy, and so simple triggers (which run as the editing
 *   user) can read it without any extra authorization.
 */

// ---------------------------------------------------------------------------
// CONFIG  — the only part you may ever need to touch
// ---------------------------------------------------------------------------

// Where to look for boolean cells to turn into checkboxes.
//   ranges:      a list of A1 ranges to scan, or `null` to scan the whole used
//                area. Only cells holding a LITERAL true/false (no formula
//                behind them) are converted, so computed booleans on these tabs
//                are left untouched — and these are the cells made delete-proof.
//   forceRanges: a list of A1 ranges whose boolean cells become checkboxes even
//                when a formula drives them (e.g. an indicator column that
//                mirrors a setting). These get the checkbox VALIDATION only —
//                their value/formula is never touched — and they are NOT made
//                delete-proof, because "restoring" one would overwrite its
//                formula with a literal false. Auto-detection can't tell a
//                display formula-checkbox (Targets!Q) from an ordinary computed
//                boolean (Army Rules!A, PPP!D5), so these must be listed by hand.
var CHECKBOX_TARGETS = [
  { sheet: 'PPP',        ranges: ['AE9:AE14', 'AE16:AE20', 'AG9:AG20'] },
  { sheet: 'Input',      ranges: null },   // ranges shift per release → scan all
  { sheet: 'Army Rules', ranges: null },
  { sheet: 'Targets',    ranges: null, forceRanges: ['Q21:Q45'] },
];

// Tabs to freeze, and how many rows to freeze on each.
var FREEZE = [
  { sheet: 'Input',  rows: 20 },
  { sheet: 'Damage', rows: 20 },
  { sheet: 'Kills',  rows: 20 },
  { sheet: 'PPW',    rows: 20 },
  { sheet: 'PPP',    rows: 20 },
];

// Hidden sheet that records, per tab, the A1 ranges holding checkboxes.
// Written by applyPostProcessing; read by onEdit. Safe to leave hidden.
var CHECKBOX_MAP_SHEET = '_wh40k_checkbox_map';
// ---------------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('WH40k')
    .addItem('Apply post-processing (checkboxes + freeze)', 'applyPostProcessing')
    .addToUi();
}

/**
 * Keeps checkboxes "delete-proof". Pressing Delete on a checkbox cell clears
 * both its value AND its checkbox rule, so the box vanishes. When that happens
 * on a configured tab, this rewrites the checkbox rule and a `false` value so
 * the box reappears, simply unchecked, instead of disappearing.
 *
 * It only touches cells that the hidden map says are checkboxes, so ordinary
 * data cells (and the formula cells beside them) are never rewritten.
 * Programmatic writes don't fire onEdit, so there is no loop.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  var range = e.range;
  var sheet = range.getSheet();
  var sheetName = sheet.getName();

  // Fast reject: only the configured tabs can ever hold these checkboxes.
  var tabs = CHECKBOX_TARGETS.map(function (t) { return t.sheet; });
  if (tabs.indexOf(sheetName) === -1) return;

  var map = readCheckboxMap(e.source);
  var checkboxRanges = map[sheetName];
  if (!checkboxRanges) return;   // map not built yet (run Apply post-processing)

  // One read of the edited block, one batched write at the end — no per-column
  // round-trips, so the cost barely grows with the size of the deletion.
  var eTop = range.getRow();
  var eLeft = range.getColumn();
  var values = range.getValues();
  var boxes = checkboxRanges.map(parseA1Bounds);

  // Collect just-cleared cells that the map says are checkboxes.
  var cleared = [];
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      var v = values[r][c];
      if (v !== '' && v !== null) continue;          // only blank (deleted) cells
      var row = eTop + r, col = eLeft + c;
      if (isInAnyBox(row, col, boxes)) cleared.push([row, col]);
    }
  }
  if (cleared.length === 0) return;

  // Re-insert checkboxes (rule + unchecked) on exactly those cells, in one call.
  // consolidateToA1Ranges keeps the list small (one entry per column run); since
  // only cleared cells are listed, boxes the user left checked are never reset.
  sheet.getRangeList(consolidateToA1Ranges(cleared).split(',')).insertCheckboxes();
}

function applyPostProcessing() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var report = [];
  var totalBoxes = 0;
  var mapRows = [];   // [sheetName, "A1,A1,…"] rows for the hidden map

  // 1) Checkboxes — and remember where they are.
  CHECKBOX_TARGETS.forEach(function (target) {
    var sheet = ss.getSheetByName(target.sheet);
    if (!sheet) {
      report.push('⚠ Checkbox tab not found: "' + target.sheet + '"');
      return;
    }
    var blocks = target.ranges
      ? target.ranges.map(function (a1) { return sheet.getRange(a1); })
      : [sheet.getDataRange()];   // null → whole used area
    var cells = [];
    blocks.forEach(function (block) {
      cells = cells.concat(applyCheckboxesToBlock(block));
    });
    totalBoxes += cells.length;
    if (cells.length > 0) {
      mapRows.push([target.sheet, consolidateToA1Ranges(cells)]);
    }
    report.push('• ' + target.sheet + ': ' + cells.length + ' checkbox' + (cells.length === 1 ? '' : 'es'));

    // Formula-driven display checkboxes: validation only, kept out of the map.
    if (target.forceRanges) {
      var forced = 0;
      target.forceRanges.forEach(function (a1) {
        forced += applyCheckboxValidationOnly(sheet.getRange(a1));
      });
      if (forced > 0) {
        totalBoxes += forced;
        report.push('  ↳ +' + forced + ' formula-driven (display-only, not delete-proofed)');
      }
    }
  });

  // 2) Persist the checkbox locations so onEdit can restore cleared boxes.
  writeCheckboxMap(ss, mapRows);

  // 3) Frozen rows (idempotent — safe to re-run).
  FREEZE.forEach(function (f) {
    var sheet = ss.getSheetByName(f.sheet);
    if (!sheet) {
      report.push('⚠ Freeze tab not found: "' + f.sheet + '"');
      return;
    }
    sheet.setFrozenRows(f.rows);
    report.push('• ' + f.sheet + ': frozen at row ' + f.rows);
  });

  SpreadsheetApp.getUi().alert(
    'Post-processing complete',
    'Applied ' + totalBoxes + ' checkbox(es) total.\n\n' + report.join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Converts every literal-boolean cell in `block` to a checkbox, leaving every
 * other cell — and its existing data validation — untouched.
 * Returns the absolute [row, col] (1-based) of each cell it converted.
 */
function applyCheckboxesToBlock(block) {
  var values = block.getValues();
  var formulas = block.getFormulas();
  var rules = block.getDataValidations();   // read existing rules so we preserve them
  var checkbox = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  var baseRow = block.getRow();
  var baseCol = block.getColumn();

  var cells = [];
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      // Literal boolean = a true/false VALUE with no formula behind it.
      if (formulas[r][c] === '' && typeof values[r][c] === 'boolean') {
        rules[r][c] = checkbox;
        cells.push([baseRow + r, baseCol + c]);
      }
    }
  }
  if (cells.length > 0) block.setDataValidations(rules);
  return cells;
}

/**
 * Adds checkbox VALIDATION to every boolean-valued cell in `block` — including
 * formula-driven ones — without touching any value or formula. Use for display
 * checkboxes that mirror a formula (e.g. Targets!Q21:Q45). Returns the count.
 */
function applyCheckboxValidationOnly(block) {
  var values = block.getValues();   // formula results; a formula bool reads as boolean here
  var rules = block.getDataValidations();
  var checkbox = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  var count = 0;
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      if (typeof values[r][c] === 'boolean') {   // skip blanks/text/numbers in the range
        rules[r][c] = checkbox;
        count++;
      }
    }
  }
  if (count > 0) block.setDataValidations(rules);
  return count;
}

// ---------------------------------------------------------------------------
// Checkbox-map helpers
// ---------------------------------------------------------------------------

/** 1 → "A", 27 → "AA", etc. */
function colToLetter(col) {
  var s = '';
  while (col > 0) {
    var m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = (col - m - 1) / 26;
  }
  return s;
}

/** "A" → 1, "AA" → 27, etc. */
function letterToCol(letters) {
  var col = 0;
  for (var i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return col;
}

/** "G21" → {row: 21, col: 7}. */
function parseA1Cell(s) {
  var m = s.match(/^([A-Z]+)(\d+)$/);
  return { row: Number(m[2]), col: letterToCol(m[1]) };
}

/** "G21:G200" (or single "G21") → numeric {top, bottom, left, right} bounds. */
function parseA1Bounds(a1) {
  var parts = a1.split(':');
  var a = parseA1Cell(parts[0]);
  var b = parts[1] ? parseA1Cell(parts[1]) : a;
  return {
    top: Math.min(a.row, b.row), bottom: Math.max(a.row, b.row),
    left: Math.min(a.col, b.col), right: Math.max(a.col, b.col)
  };
}

/** True if (row, col) falls inside any of the given numeric bounds. */
function isInAnyBox(row, col, boxes) {
  for (var i = 0; i < boxes.length; i++) {
    var b = boxes[i];
    if (row >= b.top && row <= b.bottom && col >= b.left && col <= b.right) return true;
  }
  return false;
}

/**
 * Compacts a list of [row, col] cells into A1 ranges, joining each column's
 * consecutive rows into one range, e.g. [[21,7],[22,7],…,[36,15]] → "G21:G…,O21:O36".
 */
function consolidateToA1Ranges(cells) {
  var byCol = {};
  cells.forEach(function (rc) {
    var col = rc[1];
    (byCol[col] = byCol[col] || []).push(rc[0]);
  });
  var parts = [];
  Object.keys(byCol).forEach(function (colStr) {
    var letter = colToLetter(Number(colStr));
    var rows = byCol[colStr].sort(function (a, b) { return a - b; });
    var start = rows[0], prev = rows[0];
    for (var i = 1; i <= rows.length; i++) {
      if (i < rows.length && rows[i] === prev + 1) { prev = rows[i]; continue; }
      parts.push(start === prev ? letter + start : letter + start + ':' + letter + prev);
      if (i < rows.length) { start = prev = rows[i]; }
    }
  });
  return parts.join(',');
}

/** Writes the {sheet → "A1,A1,…"} map to the hidden map sheet. */
function writeCheckboxMap(ss, mapRows) {
  var sheet = ss.getSheetByName(CHECKBOX_MAP_SHEET);
  if (!sheet) sheet = ss.insertSheet(CHECKBOX_MAP_SHEET);
  sheet.clear();
  sheet.getRange(1, 1, 1, 2).setValues([['Sheet', 'CheckboxRanges']]);
  if (mapRows.length > 0) {
    sheet.getRange(2, 1, mapRows.length, 2).setValues(mapRows);
  }
  sheet.hideSheet();
}

/** Reads the hidden map sheet into {sheetName: ["G21:G200", "O21:O36", …]}. */
function readCheckboxMap(ss) {
  var sheet = ss.getSheetByName(CHECKBOX_MAP_SHEET);
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {   // row 0 is the header
    var name = data[i][0];
    var ranges = data[i][1];
    if (name && ranges) map[name] = String(ranges).split(',');
  }
  return map;
}
