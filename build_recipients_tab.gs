/**
 * Builds (or refreshes) a "Recipients" tab in the main Bandwidth Bot sheet.
 * Sources:
 *   - Employee sheet : 1oOSF5iCl_lFvZBhR5oO8BoYahN7i03s_SL5PV7AyGt8
 *   - Brand mapping  : 1s65SeaAlKyMDirDBM-txFMOCEzRw9ZjufhL7fgZrnho
 *
 * Run buildRecipientsTab() once to create/refresh the tab.
 */

// ── Config ────────────────────────────────────────────────────────────────────
const MAIN_SHEET_ID     = '19LCHZ2uc3MbCkZWQBqVoMzVNB8OuqxdJNHlb4DYLFJY';
const EMP_SHEET_ID      = '1oOSF5iCl_lFvZBhR5oO8BoYahN7i03s_SL5PV7AyGt8';
const BRAND_SHEET_ID    = '1s65SeaAlKyMDirDBM-txFMOCEzRw9ZjufhL7fgZrnho';
const EMP_GID           = '1165838487';
const BRAND_GID         = '1996161867';
const OUTPUT_TAB        = 'Recipients';

// Designations to include (case-insensitive partial match)
const ALLOWED_DESIGNATIONS = [
  'product manager',
  'associate product manager',
  'director of product',
  'director, product',
  'associate director',
  'vp of product',
  'vp, product',
  'head of product',
  'platform',
  'product management',
];

// ── Main function ─────────────────────────────────────────────────────────────
function buildRecipientsTab() {
  Logger.log('=== buildRecipientsTab START ===');

  // 1. Read employee data
  const empRows = _readSheet(EMP_SHEET_ID, EMP_GID);
  Logger.log('Employee sheet rows: ' + empRows.length);
  Logger.log('Employee headers: ' + empRows[0].join(' | '));

  // 2. Read brand mapping data
  const brandRows = _readSheet(BRAND_SHEET_ID, BRAND_GID);
  Logger.log('Brand sheet rows: ' + brandRows.length);
  Logger.log('Brand headers: ' + brandRows[0].join(' | '));

  // 3. Parse employees
  const empHeader = empRows[0].map(h => h.toString().trim().toLowerCase());
  const employees = _parseEmployees(empRows, empHeader);
  Logger.log('Filtered employees: ' + employees.length);

  // 4. Parse brand mapping (name → brands[])
  const brandHeader = brandRows[0].map(h => h.toString().trim().toLowerCase());
  const brandMap    = _parseBrandMap(brandRows, brandHeader);
  Logger.log('Brand map entries: ' + Object.keys(brandMap).length);

  // 5. Merge: attach brand to each employee
  const merged = employees.map(e => {
    const key    = e.name.toLowerCase().trim();
    const brands = brandMap[key] || _fuzzyBrandLookup(key, brandMap) || [];
    return { ...e, brands };
  });

  // 6. Write output tab
  _writeOutputTab(merged);
  Logger.log('=== buildRecipientsTab DONE ===');
}

// ── Parse employees ───────────────────────────────────────────────────────────
function _parseEmployees(rows, header) {
  // Try to detect columns
  const iName   = _colIdx(header, ['name', 'employee name', 'full name', 'emp name', 'member name']);
  const iEmail  = _colIdx(header, ['email', 'work email', 'email address', 'corporate email', 'innovaccer email']);
  const iDesig  = _colIdx(header, ['designation', 'title', 'job title', 'role', 'position', 'level']);
  const iStatus = _colIdx(header, ['status', 'active', 'employment status', 'emp status']);

  Logger.log(`Cols → name:${iName} email:${iEmail} designation:${iDesig} status:${iStatus}`);

  const result = [];
  for (let i = 1; i < rows.length; i++) {
    const row   = rows[i];
    const name  = iName  >= 0 ? _str(row, iName)  : '';
    const email = iEmail >= 0 ? _str(row, iEmail) : '';
    const desig = iDesig >= 0 ? _str(row, iDesig) : '';
    const status = iStatus >= 0 ? _str(row, iStatus).toLowerCase() : 'active';

    if (!name) continue;

    // Skip inactive employees
    if (iStatus >= 0 && status && !status.includes('active') && status !== '') continue;

    // Filter by designation
    if (!_matchesDesig(desig)) continue;

    result.push({ name, email, designation: desig });
  }
  return result;
}

function _matchesDesig(desig) {
  const d = desig.toLowerCase().trim();
  if (!d) return false;
  return ALLOWED_DESIGNATIONS.some(allowed => d.includes(allowed));
}

// ── Parse brand map ───────────────────────────────────────────────────────────
function _parseBrandMap(rows, header) {
  const iName  = _colIdx(header, ['name', 'employee name', 'full name', 'member name', 'pm name']);
  const iBrand = _colIdx(header, ['brand', 'solution', 'product', 'brand name', 'solution name', 'account']);
  const iBrand2 = _colIdx(header, ['brand 2', 'solution 2', 'brand2', 'secondary brand']);
  const iBrand3 = _colIdx(header, ['brand 3', 'solution 3', 'brand3']);

  Logger.log(`Brand cols → name:${iName} brand:${iBrand} brand2:${iBrand2} brand3:${iBrand3}`);

  const map = {};
  for (let i = 1; i < rows.length; i++) {
    const row  = rows[i];
    const name = iName >= 0 ? _str(row, iName).toLowerCase().trim() : '';
    if (!name) continue;

    const brands = [];
    if (iBrand  >= 0 && _str(row, iBrand))  brands.push(_str(row, iBrand));
    if (iBrand2 >= 0 && _str(row, iBrand2)) brands.push(_str(row, iBrand2));
    if (iBrand3 >= 0 && _str(row, iBrand3)) brands.push(_str(row, iBrand3));

    map[name] = brands;
  }
  return map;
}

function _fuzzyBrandLookup(nameKey, brandMap) {
  // Try matching on first + last name only (ignore middle names / suffixes)
  const parts = nameKey.split(/\s+/).filter(Boolean);
  if (parts.length < 2) return null;
  const first = parts[0], last = parts[parts.length - 1];
  for (const k of Object.keys(brandMap)) {
    if (k.includes(first) && k.includes(last)) return brandMap[k];
  }
  return null;
}

// ── Write output tab ──────────────────────────────────────────────────────────
function _writeOutputTab(people) {
  const ss  = SpreadsheetApp.openById(MAIN_SHEET_ID);
  let tab   = ss.getSheetByName(OUTPUT_TAB);

  if (tab) {
    tab.clearContents();
  } else {
    tab = ss.insertSheet(OUTPUT_TAB);
  }

  // Header row
  const headers = ['Name', 'Email', 'Designation', 'Brand 1', 'Brand 2', 'Brand 3'];
  tab.appendRow(headers);

  // Style header
  const hRange = tab.getRange(1, 1, 1, headers.length);
  hRange.setBackground('#0066CC').setFontColor('#ffffff').setFontWeight('bold');

  // Data rows
  for (const p of people) {
    tab.appendRow([
      p.name,
      p.email,
      p.designation,
      p.brands[0] || '',
      p.brands[1] || '',
      p.brands[2] || '',
    ]);
  }

  // Auto-resize columns
  for (let c = 1; c <= headers.length; c++) {
    tab.autoResizeColumn(c);
  }

  // Freeze header row
  tab.setFrozenRows(1);

  Logger.log(`Written ${people.length} rows to "${OUTPUT_TAB}" tab.`);
  SpreadsheetApp.getUi().alert(`Done! ${people.length} recipients written to "${OUTPUT_TAB}" tab.\n\nCheck the Logs (View → Logs) for details on column detection.`);
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function _readSheet(sheetId, gid) {
  const url  = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv&gid=${gid}`;
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const csv  = resp.getContentText();
  return Utilities.parseCsv(csv);
}

function _colIdx(header, candidates) {
  for (const c of candidates) {
    const i = header.indexOf(c);
    if (i >= 0) return i;
  }
  // Partial match fallback
  for (const c of candidates) {
    const i = header.findIndex(h => h.includes(c) || c.includes(h));
    if (i >= 0) return i;
  }
  return -1;
}

function _str(row, idx) {
  return (idx < row.length ? row[idx] : '').toString().trim();
}

// ── Run this first to inspect headers before building ─────────────────────────
function inspectHeaders() {
  const empRows   = _readSheet(EMP_SHEET_ID, EMP_GID);
  const brandRows = _readSheet(BRAND_SHEET_ID, BRAND_GID);
  Logger.log('=== EMPLOYEE HEADERS ===');
  Logger.log(empRows[0].join(' | '));
  Logger.log('=== EMPLOYEE SAMPLE ROW ===');
  Logger.log(empRows[1] ? empRows[1].join(' | ') : '(empty)');
  Logger.log('=== BRAND HEADERS ===');
  Logger.log(brandRows[0].join(' | '));
  Logger.log('=== BRAND SAMPLE ROW ===');
  Logger.log(brandRows[1] ? brandRows[1].join(' | ') : '(empty)');
}
