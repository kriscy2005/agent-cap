// Bandwidth Bot — Google Apps Script Web App
// Deploy as: Execute as ME, Who has access: ANYONE
// Then paste the Web App URL into bandwidth-bot-py/.env as APPS_SCRIPT_URL

const SHEET_ID = '19LCHZ2uc3MbCkZWQBqVoMzVNB8OuqxdJNHlb4DYLFJY';
const LOG_TAB  = 'Bandwidth_Log';
const HEADERS  = ['Ticket Key','Summary','Brand','Person Name','Designation','Role','Bandwidth %','Submitted At'];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // Route by action
    if (data.action === 'createRecipientsTab') {
      return _createRecipientsTab(data.rows || []);
    }

    // Default: log bandwidth entries
    const entries = data.entries || [];
    if (!entries.length) return _json({ error: 'No entries' });

    const ss    = SpreadsheetApp.openById(SHEET_ID);
    let   sheet = ss.getSheetByName(LOG_TAB);

    if (!sheet) {
      sheet = ss.insertSheet(LOG_TAB);
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length)
           .setFontWeight('bold')
           .setBackground('#0066CC')
           .setFontColor('#ffffff');
    }

    const now  = new Date().toISOString();
    const rows = entries.map(r => [
      r.key, r.summary || '', r.brand || '',
      r.personName, r.designation || '', r.role || 'Other',
      r.bandwidth, now
    ]);

    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length)
         .setValues(rows);

    return _json({ success: true, rows: rows.length });
  } catch (err) {
    return _json({ error: err.message });
  }
}

function _createRecipientsTab(rows) {
  const RCPT_TAB = 'Recipients';
  const RCPT_HEADERS = ['Name', 'Email', 'Designation', 'Brand 1', 'Brand 2', 'Brand 3'];

  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(RCPT_TAB);
  if (sheet) {
    sheet.clearContents();
  } else {
    sheet = ss.insertSheet(RCPT_TAB);
  }

  // Header row
  sheet.appendRow(RCPT_HEADERS);
  sheet.getRange(1, 1, 1, RCPT_HEADERS.length)
       .setFontWeight('bold')
       .setBackground('#0066CC')
       .setFontColor('#ffffff');

  // Data rows
  const data = rows.map(r => [
    r.name, r.email, r.designation,
    r.brand1 || '', r.brand2 || '', r.brand3 || ''
  ]);
  if (data.length) {
    sheet.getRange(2, 1, data.length, RCPT_HEADERS.length).setValues(data);
  }

  // Freeze + autosize
  sheet.setFrozenRows(1);
  for (let c = 1; c <= RCPT_HEADERS.length; c++) sheet.autoResizeColumn(c);

  return _json({ success: true, tab: RCPT_TAB, rows: data.length });
}

// Health check — visit the web app URL in browser to confirm it's live
function doGet() {
  return _json({ status: 'ok', tab: LOG_TAB });
}

function _json(obj, code) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
