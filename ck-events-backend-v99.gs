/**
 * ╔══════════════════════════════════════════════════════════════════╗
 * ║       CK Events — Google Apps Script Backend v99 FULL          ║
 * ║       Rule 1: Dynamic Payment Rows with Date & Time in PDF.    ║
 * ║       Rule 2: Hard Single-Use Link Lock (Layer 1 dead page).   ║
 * ║       PIN = 272727. Payment_History_JSON includes timeFormatted.║
 * ╚══════════════════════════════════════════════════════════════════╝
 */

// ═══════════════════════════════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════════════════════════════
var SPREADSHEET_ID        = '1EBaoa2yDzk7aMDxQjrh7r782PiSH7u_5d5UeHR8seIk';
var DRIVE_ROOT_FOLDER_ID  = '1c37rw0F_1cRgvu-ISazdz2PS65g7wwhN';
var ADMIN_PHONE           = '60126597702';

// ── PART E: Global Approval PIN ────────────────────────────────────
var APPROVAL_PIN = '272727';

var INVOICE_MASTER_SHEET_NAME = 'INVOICE_MASTER';
var PAYMENT_LOG_SHEET_NAME    = 'PAYMENT_LOG';
var FILE_REGISTRY_SHEET_NAME  = 'FILE_REGISTRY';

var DB_HEADERS = [
  'Invoice_No', 'Document_Type', 'Date', 'Status', 'Payment_Status',
  'Bill_To_Company', 'Bill_To_Attn', 'Client_Phone', 'Bill_To_Address',
  'Venue', 'Event_Date_from', 'Event_Date_To', 'Setup_Date', 'Setup_time',
  'Dismantle_Date', 'Dismantle_Time', 'Invoice_Items_JSON',
  'Subtotal', 'Discount', 'Grand_Total', 'Payment_made', 'Balance_Due',
  'Latest_Payment_Date', 'Latest_Payment_Ref', 'Latest_Drive_Link',
  'Client_Folder_Link', 'Invoice_Folder_Link', 'Latest_Invoice_Version',
  'Special_Notes', 'Remarks', 'Created_By', 'Created_Timestamp',
  'Last_Updated_By', 'Last_Updated_Timestamp',
  'RECORD_MODE', 'LAST_UPDATED_AT_MY', 'LAST_UPDATED_AT_ISO',
  'UPDATED_FIELDS', 'CHANGELOG_JSON',
  'Payment_History_JSON'   // ← NEW v98: stores approved payment history
];

// ═══════════════════════════════════════════════════════════════════
// MALAYSIA DATE/TIME HELPERS
// ═══════════════════════════════════════════════════════════════════
var MY_TZ = 'Asia/Kuala_Lumpur';
var MY_MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

function _nowMY_() {
  return new Date(Utilities.formatDate(new Date(), MY_TZ, "yyyy-MM-dd'T'HH:mm:ss"));
}
function _formatDateMY_(d) {
  if (!d) return '';
  if (typeof d === 'string') {
    var m = d.match(/^(\d{1,2})[\/\-]([A-Za-z]{3})[\/\-](\d{4})$/);
    if (m) return _pad2_(+m[1]) + ' ' + m[2].charAt(0).toUpperCase()+m[2].slice(1,3).toLowerCase() + ' ' + m[3];
    m = d.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) { var dt = new Date(+m[1], +m[2]-1, +m[3]); return _pad2_(dt.getDate()) + ' ' + MY_MONTHS[dt.getMonth()] + ' ' + dt.getFullYear(); }
    if (/^\d{2} [A-Z][a-z]{2} \d{4}$/.test(d)) return d;
    return d;
  }
  try { return _pad2_(d.getDate()) + ' ' + MY_MONTHS[d.getMonth()] + ' ' + d.getFullYear(); } catch(e) { return ''; }
}
function _formatTimeMY_(t) {
  if (!t) return '';
  var h24, min;
  var m = String(t).match(/^(\d{1,2}):(\d{2})/);
  if (m) { h24 = +m[1]; min = +m[2]; }
  else {
    var m12 = String(t).match(/^(\d{1,2})(?::(\d{2}))?\s*(AM|PM|am|pm)$/i);
    if (m12) {
      var hr = +m12[1]; min = +(m12[2]||0);
      var ap = m12[3].toUpperCase();
      h24 = (ap === 'PM' && hr !== 12) ? hr+12 : (ap === 'AM' && hr === 12 ? 0 : hr);
    } else return t;
  }
  var apStr = h24 >= 12 ? 'pm' : 'am';
  var h12 = h24 > 12 ? h24-12 : (h24 === 0 ? 12 : h24);
  var timeStr = h12 + ':' + _pad2_(min) + apStr;
  var label = (h24 >= 0 && h24 < 12) ? 'Morning' : (h24 >= 12 && h24 < 18) ? 'Afternoon' : (h24 >= 18 && h24 < 22) ? 'Evening' : 'Night';
  return label + ' ' + timeStr;
}
function _formatDateTimeMY_() {
  var now = new Date();
  var myStr = Utilities.formatDate(now, MY_TZ, 'dd MMM yyyy, hh:mmaa');
  return myStr.replace(/AM$/, 'am').replace(/PM$/, 'pm');
}
function _pad2_(n) { return String(n).length < 2 ? '0'+n : String(n); }

// PAYMENT_LOG column positions (1-based)
var PL_COL = {
  TOKEN:           1,
  INVOICE_NO:      2,
  CLIENT_NAME:     3,
  AMOUNT:          4,
  DATE:            5,
  TIME:            6,
  METHOD:          7,
  REF:             8,
  NOTE:            9,
  PROOF_LINK:      10,
  APPROVAL_STATUS: 11,
  SOURCE:          12,
  SUBMITTED_AT:    13,
  APPROVED_AT:     14,
  APPROVED_BY:     15,
  RECEIPT_LINK:    16,
  SENT_AT:         17,
  SENT_BY:         18,
  APPROVE_URL:     19,
  REJECT_REASON:   20
};

// ═══════════════════════════════════════════════════════════════════
// HELPER: Build Payment_History_JSON from PAYMENT_LOG for an invoice
// ═══════════════════════════════════════════════════════════════════
function _buildPaymentHistoryJSON_(logSheet, invoiceNo) {
  try {
    var logData   = logSheet.getDataRange().getValues();
    var lHeaders  = logData[0];
    var lInvCol   = lHeaders.indexOf('INVOICE_NO');
    var lStatCol  = lHeaders.indexOf('APPROVAL_STATUS');
    var lAmtCol   = lHeaders.indexOf('AMOUNT');
    var lDateCol  = lHeaders.indexOf('DATE');
    var lTimeCol  = lHeaders.indexOf('TIME');
    var lMethCol  = lHeaders.indexOf('METHOD');
    var lProofCol = lHeaders.indexOf('PROOF_LINK');

    var entries = [];
    for (var i = 1; i < logData.length; i++) {
      if (String(logData[i][lInvCol] || '').trim() !== invoiceNo) continue;
      var st = String(logData[i][lStatCol] || '').toUpperCase();
      if (st !== 'APPROVED') continue;
      var rawTime = String(logData[i][lTimeCol] || '');
      entries.push({
        amount:       String(logData[i][lAmtCol]   || ''),
        date:         String(logData[i][lDateCol]  || ''),
        time:         rawTime,
        timeFormatted: rawTime ? _formatTimeMY_(rawTime) : '',
        method:       String(logData[i][lMethCol]  || ''),
        proofLink:    String(logData[i][lProofCol] || '')
      });
    }

    // Sort oldest-first by date string (lexicographic is fine for DD-Mon-YYYY)
    entries.sort(function(a, b) {
      var da = a.date || '', db = b.date || '';
      return da < db ? -1 : da > db ? 1 : 0;
    });

    return JSON.stringify(entries);
  } catch(e) {
    Logger.log('_buildPaymentHistoryJSON_ error: ' + e.message);
    return '[]';
  }
}

// ═══════════════════════════════════════════════════════════════════
// doGet
// ═══════════════════════════════════════════════════════════════════
function doGet(e) {
  var action = String((e.parameter && e.parameter.action) || '').toLowerCase().trim();
  if (action === 'approveform' || action === 'approve') return _serveAccountantApprovalPage_(e);
  if (action === 'clientpage')        return _serveClientPage_(e);
  if (action === 'forwardtoaccountant') return _serveForwardToAccountant_(e);
  return HtmlService.createHtmlOutput(
    '<html><body style="font-family:sans-serif;padding:40px;text-align:center">'
    + '<h2>CK Event Management System</h2>'
    + '<p style="color:#64748b">Web App is running correctly ✅</p>'
    + '</body></html>'
  ).setTitle('CK Events');
}

// ═══════════════════════════════════════════════════════════════════
// doPost
// ═══════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    var data   = JSON.parse(e.postData.contents);
    var action = String(data.action || '').toLowerCase().trim();
    if (action === 'upsert')                return handleUpsert_(data);
    if (action === 'search')                return handleSearch_(data);
    if (action === 'getinvoice')            return handleGetInvoice_(data);
    if (action === 'clientsubmitpayment')   return handleClientSubmitPayment_(data);
    if (action === 'pendpayment')           return handlePendPayment_(data);
    if (action === 'marksenttoaccountant')  return handleMarkSentToAccountant_(data);
    if (action === 'getpending')            return handleGetPending_(data);
    if (action === 'confirmpayment')        return handleConfirmPayment_(data);
    if (action === 'rejectpayment')         return handleRejectPayment_(data);
    if (action === 'admincreatedinvoice')   return handleAdminCreatedInvoice_(data);
    if (action === 'uploadapprovedpdf')     return handleUploadApprovedPdf_(data);
    return _jsonErr_('Unknown action: ' + action);
  } catch(err) {
    return _jsonErr_('doPost error: ' + err.message);
  }
}

// ═══════════════════════════════════════════════════════════════════
// handleUpsert_
// ═══════════════════════════════════════════════════════════════════
function handleUpsert_(data) {
  try {
    var ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
    var masterSheet = ss.getSheetByName(INVOICE_MASTER_SHEET_NAME);
    var row         = data.row        || {};
    var pdfBase64   = data.pdfBase64  || '';
    var fileName    = data.fileName   || ((row.Invoice_No || 'invoice') + '.pdf');
    var invoiceNo   = String(row.Invoice_No || '').trim();
    if (!invoiceNo) throw new Error('Invoice_No is required');

    var sheetData = masterSheet.getDataRange().getValues();
    var headers   = sheetData[0] || [];
    if (!headers.length || String(headers[0]).trim() === '') {
      masterSheet.clearContents();
      masterSheet.appendRow(DB_HEADERS);
      sheetData = masterSheet.getDataRange().getValues();
      headers   = sheetData[0];
    }

    var auditCols = ['RECORD_MODE','LAST_UPDATED_AT_MY','LAST_UPDATED_AT_ISO','UPDATED_FIELDS','CHANGELOG_JSON','Payment_History_JSON'];
    var headersChanged = false;
    auditCols.forEach(function(col) {
      if (headers.indexOf(col) < 0) { headers.push(col); headersChanged = true; }
    });
    if (headersChanged) masterSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    var fileUrl = '', clientFolderLink = '', invoiceFolderLink = '', version = 1;
    if (pdfBase64) {
      var rootFolder = DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);
      var clientName    = String(row.Bill_To_Company || 'Unknown').replace(/[\/\\:*?"<>|]/g, '').trim() || 'Unknown';
      var clientFolders = rootFolder.getFoldersByName(clientName);
      var clientFolder  = clientFolders.hasNext() ? clientFolders.next() : rootFolder.createFolder(clientName);
      clientFolderLink  = clientFolder.getUrl();
      var invFolders    = clientFolder.getFoldersByName(invoiceNo);
      var invFolder     = invFolders.hasNext() ? invFolders.next() : clientFolder.createFolder(invoiceNo);
      invoiceFolderLink = invFolder.getUrl();
      var existingFiles = invFolder.getFilesByType('application/pdf');
      var versionCount  = 0;
      while (existingFiles.hasNext()) { existingFiles.next(); versionCount++; }
      version = versionCount + 1;
      var bytes = Utilities.base64Decode(pdfBase64);
      var blob  = Utilities.newBlob(bytes, 'application/pdf', fileName);
      var file  = invFolder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
      try {
        var regSheet = ss.getSheetByName(FILE_REGISTRY_SHEET_NAME);
        if (regSheet) regSheet.appendRow([invoiceNo, fileName, fileUrl, new Date().toISOString(), version, row.Document_Type || '', row.Bill_To_Company || '']);
      } catch(regErr) {}
    }

    if (fileUrl)           row['Latest_Drive_Link']      = fileUrl;
    if (version)           row['Latest_Invoice_Version'] = version;
    if (clientFolderLink)  row['Client_Folder_Link']      = clientFolderLink;
    if (invoiceFolderLink) row['Invoice_Folder_Link']     = invoiceFolderLink;

    sheetData = masterSheet.getDataRange().getValues();
    headers   = sheetData[0];
    var invColIdx      = headers.indexOf('Invoice_No');
    var existingRowIdx = -1;
    var oldRecord      = {};
    if (invColIdx >= 0) {
      for (var i = 1; i < sheetData.length; i++) {
        if (String(sheetData[i][invColIdx]).trim() === invoiceNo) {
          existingRowIdx = i + 1;
          for (var j = 0; j < headers.length; j++) oldRecord[headers[j]] = sheetData[i][j];
          break;
        }
      }
    }

    var mode = existingRowIdx > 0 ? 'UPDATED' : 'CREATED';
    var nowISO = new Date().toISOString();
    var nowMY  = _formatDateTimeMY_();
    var updatedFields = [], changes = [];
    var SKIP_DIFF = {
      'Last_Updated_Timestamp':1,'Last_Updated_By':1,'Created_Timestamp':1,'Created_By':1,
      'Latest_Invoice_Version':1,'Latest_Drive_Link':1,'Client_Folder_Link':1,'Invoice_Folder_Link':1,
      'RECORD_MODE':1,'LAST_UPDATED_AT_MY':1,'LAST_UPDATED_AT_ISO':1,'UPDATED_FIELDS':1,'CHANGELOG_JSON':1,'Payment_History_JSON':1
    };

    if (mode === 'UPDATED') {
      var checkFields = Object.keys(row).concat(Object.keys(oldRecord));
      var seen = {};
      checkFields.forEach(function(field) {
        if (SKIP_DIFF[field] || seen[field]) return;
        seen[field] = true;
        var oldVal = String(oldRecord[field] || '').trim();
        var newVal = String(row[field] || '').trim();
        if (!isNaN(parseFloat(oldVal)) && !isNaN(parseFloat(newVal))) {
          if (parseFloat(oldVal).toFixed(2) !== parseFloat(newVal).toFixed(2)) {
            updatedFields.push(field); changes.push({ field: field, oldValue: oldVal, newValue: newVal });
          }
        } else if (oldVal !== newVal && newVal !== '') {
          updatedFields.push(field); changes.push({ field: field, oldValue: oldVal, newValue: newVal });
        }
      });
      if (updatedFields.length === 0) {
        return _jsonOk_({
          invoiceNo: invoiceNo, mode: 'NO_CHANGES', rowIndex: existingRowIdx,
          updatedAtMY: nowMY, updatedFields: [], changes: [],
          driveLink: String(oldRecord['Latest_Drive_Link'] || ''), fileUrl: fileUrl, url: fileUrl,
          message: 'No changes detected — nothing updated'
        });
      }
    }

    if (mode === 'CREATED') {
      row['Created_Timestamp'] = row['Created_Timestamp'] || nowISO;
      row['Created_By']        = row['Created_By']        || 'Admin';
    }
    row['Last_Updated_Timestamp'] = nowISO;
    row['Last_Updated_By']        = 'Admin';
    row['RECORD_MODE']            = mode;
    row['LAST_UPDATED_AT_MY']     = nowMY;
    row['LAST_UPDATED_AT_ISO']    = nowISO;
    row['UPDATED_FIELDS']         = updatedFields.join(', ');
    row['CHANGELOG_JSON']         = JSON.stringify(changes);

    if (existingRowIdx > 0) {
      for (var h = 0; h < headers.length; h++) {
        var colName = headers[h];
        if (!colName) continue;
        var val = row[colName];
        if (val !== undefined && val !== null && val !== '') masterSheet.getRange(existingRowIdx, h + 1).setValue(val);
      }
      var forceWrite = ['Latest_Drive_Link','Latest_Invoice_Version','Last_Updated_Timestamp',
        'RECORD_MODE','LAST_UPDATED_AT_MY','LAST_UPDATED_AT_ISO','UPDATED_FIELDS','CHANGELOG_JSON'];
      forceWrite.forEach(function(col) {
        var idx = headers.indexOf(col);
        if (idx >= 0 && row[col] !== undefined) masterSheet.getRange(existingRowIdx, idx+1).setValue(row[col]);
      });
    } else {
      var newRow = headers.map(function(h) { var v = row[h]; return (v !== undefined && v !== null) ? v : ''; });
      masterSheet.appendRow(newRow);
    }

    return _jsonOk_({
      invoiceNo: invoiceNo, mode: mode, rowIndex: existingRowIdx > 0 ? existingRowIdx : -1,
      updatedAtMY: nowMY, updatedFields: updatedFields, changes: changes,
      fileUrl: fileUrl, url: fileUrl,
      Latest_Drive_Link: fileUrl || String(oldRecord['Latest_Drive_Link'] || ''),
      driveLink: fileUrl || String(oldRecord['Latest_Drive_Link'] || ''),
      clientFolderLink: clientFolderLink, invoiceFolderLink: invoiceFolderLink, version: version,
      message: mode === 'UPDATED' ? ('Updated ' + updatedFields.length + ' field(s)') : 'New record created'
    });
  } catch(err) {
    return _jsonErr_('handleUpsert_ error: ' + err.message);
  }
}

// ═══════════════════════════════════════════════════════════════════
// handleSearch_
// ═══════════════════════════════════════════════════════════════════
function handleSearch_(data) {
  try {
    var ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
    var masterSheet = ss.getSheetByName(INVOICE_MASTER_SHEET_NAME);
    var sheetData   = masterSheet.getDataRange().getValues();
    var headers     = sheetData[0];
    var query       = String(data.query || '').trim().toLowerCase();
    var field       = String(data.field || 'all').trim().toLowerCase();
    var limit       = parseInt(data.limit) || 200;
    var results     = [];
    for (var i = 1; i < sheetData.length; i++) {
      var rowObj = {};
      for (var j = 0; j < headers.length; j++) rowObj[headers[j]] = sheetData[i][j];
      if (!rowObj['Invoice_No']) continue;
      if (query) {
        var searchIn = field === 'all' ? Object.values(rowObj).join(' ').toLowerCase() : String(rowObj[field] || '').toLowerCase();
        if (searchIn.indexOf(query) === -1) continue;
      }
      results.push(rowObj);
      if (results.length >= limit) break;
    }
    results.reverse();
    return _jsonOk_({ results: results, count: results.length });
  } catch(err) {
    return _jsonErr_('handleSearch_ error: ' + err.message);
  }
}

// ═══════════════════════════════════════════════════════════════════
// handleGetInvoice_ — v98: attaches Payment_History_JSON to response
// ═══════════════════════════════════════════════════════════════════
function handleGetInvoice_(data) {
  try {
    var key = String(data.invoiceNo || data.Invoice_No || data.docId || data.pendingId || data.query || '').trim();
    if (!key) throw new Error('invoiceNo is required');

    var ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
    var masterSheet = ss.getSheetByName(INVOICE_MASTER_SHEET_NAME);
    var values      = masterSheet.getDataRange().getValues();
    if (!values || values.length < 2) throw new Error('INVOICE_MASTER is empty');

    var headers = values[0].map(function(h){ return String(h).trim(); });
    var invCol  = headers.indexOf('Invoice_No');
    if (invCol < 0) throw new Error('Invoice_No column not found');

    function normalize(v) {
      return String(v == null ? '' : v).replace(/\u00A0/g,' ').replace(/[–—]/g,'-').replace(/\.pdf$/i,'').trim();
    }
    var target = normalize(key);

    for (var r = 1; r < values.length; r++) {
      var cell = normalize(values[r][invCol]);
      if (!cell) continue;
      if (cell === target || cell.indexOf(target) >= 0 || target.indexOf(cell) >= 0) {
        var rowObj = {};
        for (var c = 0; c < headers.length; c++) rowObj[headers[c]] = values[r][c];

        // ── v98: Attach fresh Payment_History_JSON from PAYMENT_LOG ──
        try {
          var logSheet = ss.getSheetByName(PAYMENT_LOG_SHEET_NAME);
          var freshHistJson = _buildPaymentHistoryJSON_(logSheet, String(rowObj['Invoice_No'] || '').trim());
          rowObj['Payment_History_JSON'] = freshHistJson;
          // Also write it back to the sheet for the PDF builder
          var phColIdx = headers.indexOf('Payment_History_JSON');
          if (phColIdx >= 0) {
            masterSheet.getRange(r + 1, phColIdx + 1).setValue(freshHistJson);
          }
        } catch(histErr) {
          Logger.log('Payment history attach error: ' + histErr.message);
          rowObj['Payment_History_JSON'] = rowObj['Payment_History_JSON'] || '[]';
        }

        return _jsonOk_({ row: rowObj });
      }
    }
    throw new Error('Invoice not found: ' + target);
  } catch (err) {
    return _jsonErr_('handleGetInvoice_ error: ' + err.message);
  }
}

// ═══════════════════════════════════════════════════════════════════
// handleClientSubmitPayment_
// ═══════════════════════════════════════════════════════════════════
function handleClientSubmitPayment_(data) {
  try {
    var invoiceNo   = String(data.invoiceNo   || '').trim();
    var paymentData = data.paymentData        || {};
    var slipBase64  = data.slipBase64         || '';
    var slipMime    = data.slipMime           || 'image/jpeg';
    if (!invoiceNo) throw new Error('invoiceNo is required');

    var ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    var logSheet  = ss.getSheetByName(PAYMENT_LOG_SHEET_NAME);
    var webAppUrl = ScriptApp.getService().getUrl();
    var token     = _genToken_();
    var now       = new Date();
    var proofLink = '';
    if (slipBase64) proofLink = _uploadProofSlip_(slipBase64, slipMime, invoiceNo, token);

    var approveUrl = webAppUrl
      + '?action=approveForm'
      + '&invoice='   + encodeURIComponent(invoiceNo)
      + '&token='     + encodeURIComponent(token)
      + '&pendingId=' + encodeURIComponent(token)
      + '&amt='       + encodeURIComponent(paymentData.amount    || '')
      + '&client='    + encodeURIComponent(paymentData.clientName || '');

    var newRow = new Array(20).fill('');
    newRow[PL_COL.TOKEN - 1]           = token;
    newRow[PL_COL.INVOICE_NO - 1]      = invoiceNo;
    newRow[PL_COL.CLIENT_NAME - 1]     = paymentData.clientName || '';
    newRow[PL_COL.AMOUNT - 1]          = paymentData.amount     || '';
    newRow[PL_COL.DATE - 1]            = paymentData.date       || '';
    newRow[PL_COL.TIME - 1]            = paymentData.time       || now.toLocaleTimeString('en-MY');
    newRow[PL_COL.METHOD - 1]          = paymentData.method     || '';
    newRow[PL_COL.REF - 1]             = paymentData.ref        || '';
    newRow[PL_COL.NOTE - 1]            = paymentData.note       || '';
    newRow[PL_COL.PROOF_LINK - 1]      = proofLink;
    newRow[PL_COL.APPROVAL_STATUS - 1] = 'PENDING';
    newRow[PL_COL.SOURCE - 1]          = 'client';
    newRow[PL_COL.SUBMITTED_AT - 1]    = now.toISOString();
    newRow[PL_COL.APPROVE_URL - 1]     = approveUrl;

    _ensurePaymentLogHeaders_(logSheet);
    logSheet.appendRow(newRow);

    return _jsonOk_({ token: token, proofLink: proofLink, approveUrl: approveUrl, adminPhone: ADMIN_PHONE });
  } catch(err) {
    return _jsonErr_('handleClientSubmitPayment_ error: ' + err.message);
  }
}

// ═══════════════════════════════════════════════════════════════════
// handlePendPayment_
// ═══════════════════════════════════════════════════════════════════
function handlePendPayment_(data) {
  try {
    var invoiceNo   = String(data.invoiceNo   || '').trim();
    var paymentData = data.paymentData        || {};
    var token       = String(data.token       || _genToken_()).trim();
    var slipBase64  = data.slipBase64         || '';
    var slipMime    = data.slipMime           || 'image/jpeg';
    if (!invoiceNo) throw new Error('invoiceNo is required');

    var ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
    var logSheet = ss.getSheetByName(PAYMENT_LOG_SHEET_NAME);
    var proofLink = '';
    if (slipBase64) proofLink = _uploadProofSlip_(slipBase64, slipMime, invoiceNo, token);

    _ensurePaymentLogHeaders_(logSheet);
    var existing = _findPaymentLogRow_(logSheet, token);
    if (existing) {
      _updatePaymentLogStatus_(logSheet, token, 'PENDING', proofLink || '', '');
    } else {
      var now    = new Date();
      var newRow = new Array(20).fill('');
      newRow[PL_COL.TOKEN - 1]           = token;
      newRow[PL_COL.INVOICE_NO - 1]      = invoiceNo;
      newRow[PL_COL.CLIENT_NAME - 1]     = paymentData.clientName || '';
      newRow[PL_COL.AMOUNT - 1]          = paymentData.amount     || '';
      newRow[PL_COL.DATE - 1]            = paymentData.date       || '';
      newRow[PL_COL.TIME - 1]            = paymentData.time       || '';
      newRow[PL_COL.METHOD - 1]          = paymentData.method     || '';
      newRow[PL_COL.REF - 1]             = paymentData.ref        || '';
      newRow[PL_COL.NOTE - 1]            = paymentData.note       || '';
      newRow[PL_COL.PROOF_LINK - 1]      = proofLink || paymentData.proofLink || '';
      newRow[PL_COL.APPROVAL_STATUS - 1] = 'PENDING';
      newRow[PL_COL.SOURCE - 1]          = paymentData.source     || 'admin';
      newRow[PL_COL.SUBMITTED_AT - 1]    = now.toISOString();
      logSheet.appendRow(newRow);
    }
    return _jsonOk_({ token: token, proofLink: proofLink });
  } catch(err) {
    return _jsonErr_('handlePendPayment_ error: ' + err.message);
  }
}

// ═══════════════════════════════════════════════════════════════════
// handleMarkSentToAccountant_
// ═══════════════════════════════════════════════════════════════════
function handleMarkSentToAccountant_(data) {
  try {
    var pendingId   = String(data.pendingId || data.token || '').trim();
    var invoiceNo   = String(data.invoiceNo || '').trim();
    var sentBy      = String(data.sentBy    || 'Admin').trim();
    var paymentData = data.paymentData      || {};
    if (!pendingId || !invoiceNo) throw new Error('Missing pendingId or invoiceNo');

    var ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
    var logSheet    = ss.getSheetByName(PAYMENT_LOG_SHEET_NAME);
    var masterSheet = ss.getSheetByName(INVOICE_MASTER_SHEET_NAME);
    var now         = new Date();
    var webAppUrl   = ScriptApp.getService().getUrl();

    var approveUrl = webAppUrl
      + '?action=approveForm'
      + '&invoice='   + encodeURIComponent(invoiceNo)
      + '&pendingId=' + encodeURIComponent(pendingId)
      + '&token='     + encodeURIComponent(pendingId)
      + '&amt='       + encodeURIComponent(paymentData.amount    || '')
      + '&client='    + encodeURIComponent(paymentData.clientName || '');

    _ensurePaymentLogHeaders_(logSheet);
    var existingRow = _findPaymentLogRow_(logSheet, pendingId);

    if (existingRow) {
      var data2   = logSheet.getDataRange().getValues();
      var headers = data2[0];
      var tokCol  = headers.indexOf('TOKEN'); if (tokCol < 0) tokCol = 0;
      var stCol   = headers.indexOf('APPROVAL_STATUS');
      var satCol  = headers.indexOf('SENT_AT');
      var sbyCol  = headers.indexOf('SENT_BY');
      var auCol   = headers.indexOf('APPROVE_URL');
      for (var i = 1; i < data2.length; i++) {
        if (String(data2[i][tokCol]).trim() !== pendingId) continue;
        var rn = i + 1;
        if (stCol  >= 0) logSheet.getRange(rn, stCol  + 1).setValue('SENT_TO_ACCOUNTANT');
        if (satCol >= 0) logSheet.getRange(rn, satCol + 1).setValue(now.toISOString());
        if (sbyCol >= 0) logSheet.getRange(rn, sbyCol + 1).setValue(sentBy);
        if (auCol  >= 0) logSheet.getRange(rn, auCol  + 1).setValue(approveUrl);
        break;
      }
    } else {
      var invoiceRow = _findInvoiceRow_(masterSheet, invoiceNo);
      var clientName = (invoiceRow && invoiceRow['Bill_To_Company']) || paymentData.clientName || '';
      var newRow = new Array(20).fill('');
      newRow[PL_COL.TOKEN - 1]           = pendingId;
      newRow[PL_COL.INVOICE_NO - 1]      = invoiceNo;
      newRow[PL_COL.CLIENT_NAME - 1]     = clientName;
      newRow[PL_COL.AMOUNT - 1]          = paymentData.amount  || '';
      newRow[PL_COL.DATE - 1]            = paymentData.date    || '';
      newRow[PL_COL.TIME - 1]            = paymentData.time    || '';
      newRow[PL_COL.METHOD - 1]          = paymentData.method  || '';
      newRow[PL_COL.REF - 1]             = paymentData.ref     || '';
      newRow[PL_COL.NOTE - 1]            = paymentData.note    || '';
      newRow[PL_COL.PROOF_LINK - 1]      = paymentData.proofLink || '';
      newRow[PL_COL.APPROVAL_STATUS - 1] = 'SENT_TO_ACCOUNTANT';
      newRow[PL_COL.SOURCE - 1]          = paymentData.source  || 'admin';
      newRow[PL_COL.SUBMITTED_AT - 1]    = now.toISOString();
      newRow[PL_COL.SENT_AT - 1]         = now.toISOString();
      newRow[PL_COL.SENT_BY - 1]         = sentBy;
      newRow[PL_COL.APPROVE_URL - 1]     = approveUrl;
      logSheet.appendRow(newRow);
    }
    return _jsonOk_({ pendingId: pendingId, invoiceNo: invoiceNo, status: 'SENT_TO_ACCOUNTANT', approveUrl: approveUrl });
  } catch(err) {
    return _jsonErr_('handleMarkSentToAccountant_ error: ' + err.message);
  }
}

// ═══════════════════════════════════════════════════════════════════
// handleGetPending_
// ═══════════════════════════════════════════════════════════════════
function handleGetPending_(data) {
  try {
    var ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
    var logSheet = ss.getSheetByName(PAYMENT_LOG_SHEET_NAME);
    var rows     = logSheet.getDataRange().getValues();
    var headers  = rows[0];
    var pending  = [];
    for (var i = 1; i < rows.length; i++) {
      var obj = {};
      for (var j = 0; j < headers.length; j++) obj[headers[j]] = rows[i][j];
      var status = String(obj['APPROVAL_STATUS'] || '').trim().toUpperCase();
      if (status !== 'PENDING' && status !== 'SENT_TO_ACCOUNTANT') continue;
      pending.push({
        token: obj['TOKEN'] || '', invoiceNo: obj['INVOICE_NO'] || '',
        clientName: obj['CLIENT_NAME'] || '', amount: obj['AMOUNT'] || '',
        date: obj['DATE'] || '', time: obj['TIME'] || '', method: obj['METHOD'] || '',
        ref: obj['REF'] || '', note: obj['NOTE'] || '', proofLink: obj['PROOF_LINK'] || '',
        source: obj['SOURCE'] || 'admin', savedAt: obj['SUBMITTED_AT'] || '',
        status: status, approveUrl: obj['APPROVE_URL'] || '', sentAt: obj['SENT_AT'] || ''
      });
    }
    return _jsonOk_({ pending: pending });
  } catch(err) {
    return _jsonErr_('handleGetPending_ error: ' + err.message);
  }
}

// ═══════════════════════════════════════════════════════════════════
// handleConfirmPayment_ — PART E: PIN + Layer-2 lock + PART F (Special_Notes)
//                         v98: also writes Payment_History_JSON to INVOICE_MASTER
// ═══════════════════════════════════════════════════════════════════
function handleConfirmPayment_(data) {
  try {
    // ── PART E: PIN check ─────────────────────────────────────────
    if (!data.pin || String(data.pin).trim() !== APPROVAL_PIN) {
      return _jsonErr_('Unauthorized: Invalid PIN');
    }

    var token       = String(data.token     || '').trim();
    var invoiceNo   = String(data.invoiceNo || '').trim();
    var paymentData = data.paymentData      || {};
    if (!token || !invoiceNo) throw new Error('token and invoiceNo are required');

    var ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
    var logSheet    = ss.getSheetByName(PAYMENT_LOG_SHEET_NAME);
    var masterSheet = ss.getSheetByName(INVOICE_MASTER_SHEET_NAME);

    // ── PART E Layer 2: Re-check status before updating ───────────
    var logRowCheck = _findPaymentLogRow_(logSheet, token);
    if (logRowCheck) {
      var checkStatus = String(logRowCheck['APPROVAL_STATUS'] || '').toUpperCase();
      if (checkStatus === 'APPROVED' || checkStatus === 'REJECTED') {
        return _jsonErr_('Too late! This payment was already processed.');
      }
    }

    var invoiceRow = _findInvoiceRow_(masterSheet, invoiceNo);
    if (!invoiceRow) throw new Error('Invoice not found: ' + invoiceNo);

    var amount     = parseFloat(String(paymentData.amount || '0').replace(/[^0-9.]/g,'')) || 0;
    var grandTotal = parseFloat(String(invoiceRow['Grand_Total'] || invoiceRow['Subtotal'] || '0').replace(/[^0-9.]/g,'')) || 0;
    var prevPaid   = parseFloat(String(invoiceRow['Payment_made'] || '0').replace(/[^0-9.]/g,'')) || 0;
    var newPaid    = prevPaid + amount;
    var newBalance = Math.max(0, grandTotal - newPaid);
    var payStatus  = newBalance <= 0.01 ? 'Fully Paid' : 'Partially Paid';
    var nowMY      = _formatDateTimeMY_();
    var nowISO     = new Date().toISOString();

    // ── STEP 1: Update PAYMENT_LOG status to APPROVED ─────────────
    _updatePaymentLogStatus_(logSheet, token, 'APPROVED', '', '');

    // ── STEP 2: Build Payment_History_JSON from ALL approved entries ─
    var freshHistJson = _buildPaymentHistoryJSON_(logSheet, invoiceNo);

    // ── STEP 3: Update INVOICE_MASTER ─────────────────────────────
    var masterData = masterSheet.getDataRange().getValues();
    var mHeaders   = masterData[0];
    var invCol     = mHeaders.indexOf('Invoice_No');
    var masterRowNum = -1;

    for (var i = 1; i < masterData.length; i++) {
      if (String(masterData[i][invCol]).trim() !== invoiceNo) continue;
      masterRowNum = i + 1;
      _setSheetCell_(masterSheet, masterRowNum, mHeaders, 'Payment_made',           newPaid);
      _setSheetCell_(masterSheet, masterRowNum, mHeaders, 'Balance_Due',            newBalance);
      _setSheetCell_(masterSheet, masterRowNum, mHeaders, 'Payment_Status',         payStatus);
      _setSheetCell_(masterSheet, masterRowNum, mHeaders, 'Latest_Payment_Date',    paymentData.date || nowMY);
      _setSheetCell_(masterSheet, masterRowNum, mHeaders, 'Latest_Payment_Ref',     paymentData.ref  || '');
      _setSheetCell_(masterSheet, masterRowNum, mHeaders, 'Last_Updated_Timestamp', nowISO);
      _setSheetCell_(masterSheet, masterRowNum, mHeaders, 'LAST_UPDATED_AT_MY',     nowMY);
      _setSheetCell_(masterSheet, masterRowNum, mHeaders, 'LAST_UPDATED_AT_ISO',    nowISO);
      _setSheetCell_(masterSheet, masterRowNum, mHeaders, 'RECORD_MODE',            'PAYMENT_APPROVED');
      _setSheetCell_(masterSheet, masterRowNum, mHeaders, 'Payment_History_JSON',   freshHistJson);
      if (payStatus === 'Fully Paid') _setSheetCell_(masterSheet, masterRowNum, mHeaders, 'Status', 'Paid');

      // ── PART F: Clear old payment history from Special_Notes ─────
      try {
        var existingNotes = String(invoiceRow['Special_Notes'] || '');
        var cleanedNotes  = existingNotes.split('--- PAYMENT HISTORY ---')[0].trim();
        // Do NOT append old-style history to Special_Notes anymore (handled by Payment_History_JSON)
        if (cleanedNotes !== existingNotes) {
          _setSheetCell_(masterSheet, masterRowNum, mHeaders, 'Special_Notes', cleanedNotes);
        }
      } catch(notesErr) {
        Logger.log('Special_Notes cleanup error: ' + notesErr.message);
      }

      break;
    }

    var existingDriveLink = invoiceRow['Latest_Drive_Link'] || '';
    return _jsonOk_({
      invoiceNo: invoiceNo, amount: amount, newPaid: newPaid,
      balance: newBalance.toFixed(2), paymentStatus: payStatus,
      driveLink: existingDriveLink, needsBeautifulPdf: true,
      paymentHistoryJSON: freshHistJson,
      updatedAtMY: nowMY,
      updatedFields: ['Payment_made','Balance_Due','Payment_Status','Payment_History_JSON']
    });
  } catch(err) {
    return _jsonErr_('handleConfirmPayment_ error: ' + err.message);
  }
}

// ═══════════════════════════════════════════════════════════════════
// handleUploadApprovedPdf_
// ═══════════════════════════════════════════════════════════════════
function handleUploadApprovedPdf_(data) {
  try {
    var invoiceNo = String(data.invoiceNo  || '').trim();
    var pdfBase64 = data.pdfBase64         || '';
    var fileName  = data.fileName          || (invoiceNo + '_approved.pdf');
    if (!invoiceNo) throw new Error('invoiceNo is required');
    if (!pdfBase64) throw new Error('pdfBase64 is required');

    var ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
    var masterSheet = ss.getSheetByName(INVOICE_MASTER_SHEET_NAME);
    var invoiceRow  = _findInvoiceRow_(masterSheet, invoiceNo);
    if (!invoiceRow) throw new Error('Invoice not found: ' + invoiceNo);

    var clientName   = String(invoiceRow['Bill_To_Company'] || 'Unknown').replace(/[\/\\:*?"<>|]/g,'').trim() || 'Unknown';
    var rootFolder   = DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);
    var cFolders     = rootFolder.getFoldersByName(clientName);
    var clientFolder = cFolders.hasNext() ? cFolders.next() : rootFolder.createFolder(clientName);
    var iFolders     = clientFolder.getFoldersByName(invoiceNo);
    var invFolder    = iFolders.hasNext() ? iFolders.next() : clientFolder.createFolder(invoiceNo);

    var bytes   = Utilities.base64Decode(pdfBase64);
    var blob    = Utilities.newBlob(bytes, 'application/pdf', fileName);
    var pdfFile = invFolder.createFile(blob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileUrl = pdfFile.getUrl();

    var masterData = masterSheet.getDataRange().getValues();
    var mHeaders   = masterData[0];
    var invCol     = mHeaders.indexOf('Invoice_No');
    for (var i = 1; i < masterData.length; i++) {
      if (String(masterData[i][invCol]).trim() !== invoiceNo) continue;
      _setSheetCell_(masterSheet, i+1, mHeaders, 'Latest_Drive_Link', fileUrl);
      _setSheetCell_(masterSheet, i+1, mHeaders, 'LAST_UPDATED_AT_MY', _formatDateTimeMY_());
      break;
    }

    try {
      var logSheet2   = ss.getSheetByName(PAYMENT_LOG_SHEET_NAME);
      var logData2    = logSheet2.getDataRange().getValues();
      var lHeaders2   = logData2[0];
      var lInvCol     = lHeaders2.indexOf('INVOICE_NO');
      var lReceiptCol = lHeaders2.indexOf('RECEIPT_LINK');
      var lApprAtCol  = lHeaders2.indexOf('APPROVED_AT');
      var lStatusCol2 = lHeaders2.indexOf('APPROVAL_STATUS');
      var nowMY2      = _formatDateTimeMY_();
      for (var li = logData2.length - 1; li >= 1; li--) {
        if (String(logData2[li][lInvCol]).trim() !== invoiceNo) continue;
        if (String(logData2[li][lStatusCol2] || '').toUpperCase() !== 'APPROVED') continue;
        if (lReceiptCol >= 0) logSheet2.getRange(li+1, lReceiptCol+1).setValue(fileUrl);
        if (lApprAtCol  >= 0 && !logData2[li][lApprAtCol]) logSheet2.getRange(li+1, lApprAtCol+1).setValue(nowMY2);
        break;
      }
    } catch(logErr) { Logger.log('PAYMENT_LOG receipt update error: ' + logErr.message); }

    return _jsonOk_({ invoiceNo: invoiceNo, driveLink: fileUrl, message: 'PDF uploaded and Latest_Drive_Link updated' });
  } catch(err) {
    return _jsonErr_('handleUploadApprovedPdf_ error: ' + err.message);
  }
}

// ═══════════════════════════════════════════════════════════════════
// handleRejectPayment_ — PART E: PIN + reason + single-use lock
// ═══════════════════════════════════════════════════════════════════
function handleRejectPayment_(data) {
  try {
    if (!data.pin || String(data.pin).trim() !== APPROVAL_PIN) {
      return _jsonErr_('Unauthorized: Invalid PIN');
    }
    var token     = String(data.token     || '').trim();
    var invoiceNo = String(data.invoiceNo || '').trim();
    var reason    = String(data.reason    || 'No reason given').trim();
    if (!token) throw new Error('token is required');

    var ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
    var logSheet = ss.getSheetByName(PAYMENT_LOG_SHEET_NAME);

    // ── Layer 2 lock ──────────────────────────────────────────────
    var logRowCheck = _findPaymentLogRow_(logSheet, token);
    if (logRowCheck) {
      var checkStatus = String(logRowCheck['APPROVAL_STATUS'] || '').toUpperCase();
      if (checkStatus === 'APPROVED' || checkStatus === 'REJECTED') {
        return _jsonErr_('Too late! This payment was already processed.');
      }
    }

    _updatePaymentLogStatus_(logSheet, token, 'REJECTED', '', reason);
    return _jsonOk_({ token: token, invoiceNo: invoiceNo, reason: reason });
  } catch(err) {
    return _jsonErr_('handleRejectPayment_ error: ' + err.message);
  }
}

// ═══════════════════════════════════════════════════════════════════
// handleAdminCreatedInvoice_
// ═══════════════════════════════════════════════════════════════════
function handleAdminCreatedInvoice_(data) {
  data.pdfBase64 = data.pdfBase64 || '';
  return handleUpsert_(data);
}

// ═══════════════════════════════════════════════════════════════════
// _serveClientPage_ — PART A (one link) + PART B (PDF/Pay buttons) + PART C (two WA buttons)
// ═══════════════════════════════════════════════════════════════════
function _serveClientPage_(e) {
  var invoiceNo = String(e.parameter.invoice || '').trim();
  if (!invoiceNo) {
    return HtmlService.createHtmlOutput('<html><body style="font-family:sans-serif;padding:40px;text-align:center"><h2>Invoice not found</h2></body></html>');
  }

  var ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
  var masterSheet = ss.getSheetByName(INVOICE_MASTER_SHEET_NAME);
  var invoiceRow  = _findInvoiceRow_(masterSheet, invoiceNo);

  if (!invoiceRow) {
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family:sans-serif;padding:40px;text-align:center">'
      + '<h2 style="color:#dc2626">Invoice Not Found</h2>'
      + '<p>Invoice <strong>' + _esc_(invoiceNo) + '</strong> could not be found.</p>'
      + '</body></html>'
    );
  }

  var grandTotal  = parseFloat(String(invoiceRow['Grand_Total']  || '0').replace(/[^0-9.]/g,'')) || 0;
  var prevPaid    = parseFloat(String(invoiceRow['Payment_made'] || '0').replace(/[^0-9.]/g,'')) || 0;
  var balance     = Math.max(0, grandTotal - prevPaid);
  var driveLink   = invoiceRow['Latest_Drive_Link'] || '';
  var clientName  = String(invoiceRow['Bill_To_Company'] || '');
  var eventDate   = String(invoiceRow['Event_Date_from'] || '');
  var venue       = String(invoiceRow['Venue'] || '');
  var webAppUrl   = ScriptApp.getService().getUrl();

  // ── PART A: Client page URL — the ONLY link sent in WA ────────
  var clientPageUrl = webAppUrl + '?action=clientpage&invoice=' + encodeURIComponent(invoiceNo);

  // ── PART B: Quick-link buttons ─────────────────────────────────
  var topLinksHtml = '';
  if (driveLink) {
    topLinksHtml += '<a href="' + _esc_(driveLink) + '" target="_blank" class="quick-link">📄 View / Download PDF →</a>';
  }
  if (balance > 0.01) {
    topLinksHtml += '<a href="#payCard" class="quick-link quick-link-green" onclick="document.getElementById(\'payCard\').scrollIntoView({behavior:\'smooth\'});return false;">💳 Make Payment →</a>';
  }

  var payFormHtml = '';
  if (balance > 0.01) {
    payFormHtml = ''
      + '<div class="card" id="payCard">'
      + '<div style="font-weight:900;font-size:17px;margin-bottom:4px">💳 Submit Payment</div>'
      + '<div style="font-size:12px;color:#94a3b8;margin-bottom:16px">Fill all fields, upload slip, then submit</div>'
      + '<div class="field-group"><label class="field-lbl">Invoice No</label>'
      + '<input class="field-inp" id="f_invoice" value="' + _esc_(invoiceNo) + '" readonly style="background:#f8fafc;color:#64748b"></div>'
      + '<div class="field-group"><label class="field-lbl">Client Name</label>'
      + '<input class="field-inp" id="f_client" value="' + _esc_(clientName) + '" readonly style="background:#f8fafc;color:#64748b"></div>'
      + '<div class="field-group"><label class="field-lbl">Amount Paying (RM) <span style="color:#dc2626">*</span></label>'
      + '<input class="field-inp" id="f_amt" type="number" step="0.01" min="0.01" placeholder="e.g. 1110.00"></div>'
      + '<div class="field-group"><label class="field-lbl">Payment Method <span style="color:#dc2626">*</span></label>'
      + '<select class="field-inp" id="f_method"><option value="">-- Select method --</option>'
      + '<option>Bank Transfer</option><option>DuitNow QR</option><option>DuitNow Transfer</option>'
      + '<option>Cash</option><option>Cheque</option><option>Online Banking (FPX)</option>'
      + '<option>Credit/Debit Card</option><option>Other</option></select></div>'
      + '<div class="field-group"><label class="field-lbl">Reference / Transaction No</label>'
      + '<input class="field-inp" id="f_ref" placeholder="e.g. TT123456789 (optional)"></div>'
      + '<div class="field-group"><label class="field-lbl">Date &amp; Time</label>'
      + '<input class="field-inp" id="f_datetime" type="datetime-local"></div>'
      + '<div class="field-group"><label class="field-lbl">Payment Slip <span style="color:#94a3b8;font-weight:500">(JPG / PNG / PDF)</span></label>'
      + '<label for="f_slip" id="slipLabel" style="display:flex;align-items:center;gap:10px;padding:12px 14px;border:2px dashed #cbd5e1;border-radius:10px;cursor:pointer;background:#f8fafc;color:#64748b;font-size:13px;font-weight:600;transition:border-color .2s">'
      + '<span style="font-size:22px">📎</span><span id="slipLabelTxt">Tap to upload slip</span></label>'
      + '<input id="f_slip" type="file" accept="image/jpeg,image/jpg,image/png,application/pdf" style="display:none"></div>'
      + '<div id="slipPreview" style="display:none;margin:-4px 0 12px">'
      + '<img id="slipImg" style="max-width:100%;border-radius:8px;border:1px solid #e2e8f0;display:none">'
      + '<div id="slipPdfTag" style="display:none;padding:9px 12px;background:#eff6ff;border-radius:8px;border:1px solid #bfdbfe;font-size:12px;color:#1d4ed8;font-weight:700"></div>'
      + '</div>'
      + '<button class="btn btn-blue" id="btnSubmit" onclick="doSubmit()">📤 Submit to System</button>'
      + '<div id="submitMsg" style="text-align:center;margin-top:10px;font-size:13px;font-weight:700;min-height:18px"></div>'
      + '</div>'

      // ── PART C: Success card with TWO WhatsApp buttons ─────────
      + '<div class="card" id="waCard" style="display:none;border:2px solid #22c55e">'
      + '<div style="font-weight:900;font-size:16px;margin-bottom:6px;color:#15803d">✅ Submitted Successfully!</div>'
      + '<div style="font-size:13px;color:#64748b;margin-bottom:14px">Tap the buttons below to notify admin and accountant via WhatsApp.</div>'
      + '<div id="waMsgPreview" style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;padding:12px;font-size:12px;color:#166534;white-space:pre-wrap;word-break:break-word;margin-bottom:12px;line-height:1.6"></div>'
      + '<button class="btn btn-blue" id="btnWaAdmin" onclick="openWAAdmin()">💬 WhatsApp Admin</button>'
      + '<div id="btnWaAcctWrapper" style="margin-top:8px">'
      + '<button class="btn btn-green" id="btnWaAcct" onclick="openWAAcct()">💬 WhatsApp Accountant</button>'
      + '<div id="noAcctNote" style="display:none;font-size:11px;color:#94a3b8;text-align:center;margin-top:6px">Accountant number not set — please WhatsApp admin only.</div>'
      + '</div>'
      + '</div>';
  }

  var html = '<!DOCTYPE html><html><head><meta charset="utf-8">'
    + '<meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<title>Invoice ' + _esc_(invoiceNo) + '</title>'
    + '<style>'
    + '*{box-sizing:border-box;margin:0;padding:0}'
    + 'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;background:#f1f5f9;padding:16px 12px 32px}'
    + '.card{background:#fff;border-radius:16px;padding:20px;max-width:520px;margin:0 auto 14px;box-shadow:0 2px 16px rgba(0,0,0,.07)}'
    + '.row{display:flex;justify-content:space-between;align-items:center;padding:8px 0;border-bottom:1px solid #f1f5f9;font-size:14px}'
    + '.row:last-child{border-bottom:none}'
    + '.lbl{color:#94a3b8;font-size:13px}.val{font-weight:700;text-align:right;max-width:65%}'
    + '.field-group{margin-bottom:10px}'
    + '.field-lbl{display:block;font-size:12px;font-weight:700;color:#64748b;margin-bottom:4px;text-transform:uppercase;letter-spacing:.04em}'
    + '.field-inp{width:100%;padding:11px 13px;border:1.5px solid #e2e8f0;border-radius:10px;font-size:14px;color:#1e293b;outline:none;transition:border-color .2s;background:#fff}'
    + '.field-inp:focus{border-color:#2563eb}'
    + '.btn{display:block;width:100%;padding:15px;border:none;border-radius:12px;font-size:15px;font-weight:800;cursor:pointer;margin-top:6px;letter-spacing:.02em;transition:opacity .2s}'
    + '.btn:disabled{opacity:.45;cursor:not-allowed}'
    + '.btn-blue{background:linear-gradient(135deg,#2563eb,#1d4ed8);color:#fff}'
    + '.btn-green{background:linear-gradient(135deg,#16a34a,#15803d);color:#fff}'
    + '.quick-link{display:block;text-align:center;margin:0 auto 10px;max-width:520px;padding:12px;background:#eff6ff;border-radius:10px;color:#2563eb;font-weight:700;font-size:13px;text-decoration:none}'
    + '.quick-link-green{background:#f0fdf4;color:#16a34a}'
    + '</style>'
    + '</head><body>'
    + topLinksHtml
    + '<div class="card">'
    + '<div style="background:linear-gradient(135deg,#1e3a8a,#3b82f6);border-radius:12px;padding:18px 20px;color:#fff;margin-bottom:16px">'
    + '<div style="font-size:11px;opacity:.75;font-weight:700;text-transform:uppercase;letter-spacing:.12em">CK Event Management</div>'
    + '<div style="font-size:22px;font-weight:900;margin-top:4px">' + _esc_(invoiceNo) + '</div>'
    + '</div>'
    + '<div class="row"><span class="lbl">Client</span><span class="val">' + _esc_(clientName) + '</span></div>'
    + '<div class="row"><span class="lbl">Event Date</span><span class="val">' + _esc_(eventDate) + '</span></div>'
    + '<div class="row"><span class="lbl">Venue</span><span class="val">' + _esc_(venue) + '</span></div>'
    + '<div class="row"><span class="lbl">Grand Total</span><span class="val">RM ' + grandTotal.toFixed(2) + '</span></div>'
    + '<div class="row"><span class="lbl">Amount Paid</span><span class="val" style="color:#059669">RM ' + prevPaid.toFixed(2) + '</span></div>'
    + '<div class="row"><span class="lbl">Balance Due</span><span class="val" style="color:' + (balance <= 0.01 ? '#059669' : '#dc2626') + ';font-size:18px">RM ' + balance.toFixed(2) + '</span></div>'
    + (balance <= 0.01 ? '<div style="margin-top:14px;padding:12px;background:#f0fdf4;border-radius:10px;text-align:center;color:#15803d;font-weight:800;font-size:15px">✅ Fully Paid — Thank you!</div>' : '')
    + '</div>'
    + payFormHtml
    + '<script>'
    + 'var INVOICE_NO=' + JSON.stringify(invoiceNo) + ';'
    + 'var CLIENT_NAME=' + JSON.stringify(clientName) + ';'
    + 'var EVENT_DATE=' + JSON.stringify(eventDate) + ';'
    + 'var VENUE=' + JSON.stringify(venue) + ';'
    + 'var ADMIN_PHONE=' + JSON.stringify(ADMIN_PHONE) + ';'
    + 'var ACCOUNTANT_PHONE="";'
    + 'var WEB_APP_URL=' + JSON.stringify(webAppUrl) + ';'

    + '(function(){var dt=document.getElementById("f_datetime");if(!dt)return;var n=new Date(),pad=function(x){return String(x).padStart(2,"0");};dt.value=n.getFullYear()+"-"+pad(n.getMonth()+1)+"-"+pad(n.getDate())+"T"+pad(n.getHours())+":"+pad(n.getMinutes());})();'

    + '(function(){var fi=document.getElementById("f_slip");if(!fi)return;fi.addEventListener("change",function(){var f=fi.files[0];if(!f)return;document.getElementById("slipLabelTxt").textContent="✅ "+f.name;document.getElementById("slipLabel").style.borderColor="#22c55e";document.getElementById("slipLabel").style.color="#16a34a";var prev=document.getElementById("slipPreview");var img=document.getElementById("slipImg");var pdf=document.getElementById("slipPdfTag");prev.style.display="block";if(f.type.indexOf("pdf")>=0){img.style.display="none";pdf.style.display="block";pdf.textContent="📄 "+f.name;}else{pdf.style.display="none";img.style.display="block";var rdr=new FileReader();rdr.onload=function(ev){img.src=ev.target.result;};rdr.readAsDataURL(f);}});})();'

    + 'var _waMsgAdmin="";var _waMsgAcct="";var _submitted=false;'

    + 'function doSubmit(){'
    + 'if(_submitted){alert("Already submitted.");return;}'
    + 'var amt=document.getElementById("f_amt").value.trim();'
    + 'var method=document.getElementById("f_method").value.trim();'
    + 'var ref=document.getElementById("f_ref").value.trim();'
    + 'var dtVal=document.getElementById("f_datetime").value;'
    + 'var slipFi=document.getElementById("f_slip");'
    + 'var msgEl=document.getElementById("submitMsg");'
    + 'if(!amt||isNaN(parseFloat(amt))||parseFloat(amt)<=0){msgEl.style.color="#dc2626";msgEl.textContent="⚠️ Please enter a valid amount.";return;}'
    + 'if(!method){msgEl.style.color="#dc2626";msgEl.textContent="⚠️ Please select a payment method.";return;}'
    + 'var btn=document.getElementById("btnSubmit");'
    + 'btn.disabled=true;btn.textContent="⏳ Uploading...";'
    + 'msgEl.style.color="#64748b";msgEl.textContent="Please wait...";'
    + 'var dtDisplay=dtVal?dtVal.replace("T"," "):(new Date()).toLocaleString("en-MY");'
    + 'function send(b64,mime){'
    + 'var payload={action:"clientSubmitPayment",invoiceNo:INVOICE_NO,slipBase64:b64,slipMime:mime,paymentData:{clientName:CLIENT_NAME,amount:amt,method:method,ref:ref,date:dtDisplay,time:"",source:"client"}};'
    + 'fetch(WEB_APP_URL,{method:"POST",redirect:"follow",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify(payload)})'
    + '.then(function(r){return r.json();})'
    + '.then(function(j){'
    + 'if(j.status==="success"){'
    + '_submitted=true;btn.textContent="✅ Submitted";'
    + 'msgEl.style.color="#059669";msgEl.textContent="Saved! Now notify via WhatsApp 👇";'
    // PART C: Build message body
    + 'var slipLine=j.proofLink?"\\n🔗 *Slip:* "+j.proofLink:"\\n📎 No slip";'
    + 'var approvalLine=j.approveUrl?"\\n\\n🔐 *Approval Link:*\\n"+j.approveUrl:"";'
    + 'var forwardUrl=WEB_APP_URL+"?action=forwardToAccountant&token="+encodeURIComponent(j.token)+"&invoice="+encodeURIComponent(INVOICE_NO)+"&acct="+encodeURIComponent(ACCOUNTANT_PHONE);'
    + 'var baseMsg="*🔔 PAYMENT SUBMITTED — CK Events*"+"\\n━━━━━━━━━━━━━━━━━━━━"+"\\n📋 *Invoice:* "+INVOICE_NO+"\\n👤 *Client:* "+CLIENT_NAME+"\\n📅 *Event:* "+EVENT_DATE+"\\n📍 *Venue:* "+VENUE+"\\n━━━━━━━━━━━━━━━━━━━━"+"\\n💵 *Amount:* RM "+amt+"\\n🏦 *Method:* "+method+(ref?"\\n🔢 *Ref No:* "+ref:"")+"\\n🕐 *Submitted:* "+dtDisplay+slipLine+approvalLine+"\\n━━━━━━━━━━━━━━━━━━━━"+"\\n_Please approve or reject via the link above._";'
    + '_waMsgAdmin=baseMsg+(ACCOUNTANT_PHONE?"\\n\\n➡️ *Forward to Accountant:*\\n"+forwardUrl:"");'
    + '_waMsgAcct=baseMsg;'
    + 'document.getElementById("waMsgPreview").textContent=_waMsgAdmin;'
    + 'document.getElementById("waCard").style.display="block";'
    + 'document.getElementById("waCard").scrollIntoView({behavior:"smooth",block:"center"});'
    + 'if(!ACCOUNTANT_PHONE){document.getElementById("btnWaAcct").style.display="none";document.getElementById("noAcctNote").style.display="block";}'
    + '}else{btn.disabled=false;btn.textContent="📤 Submit to System";msgEl.style.color="#dc2626";msgEl.textContent="❌ "+j.message;}'
    + '}).catch(function(){btn.disabled=false;btn.textContent="📤 Submit to System";msgEl.style.color="#dc2626";msgEl.textContent="❌ Network error.";});}'
    + 'var slipFile=slipFi&&slipFi.files.length>0?slipFi.files[0]:null;'
    + 'if(slipFile){btn.textContent="⏳ Reading file...";var rdr=new FileReader();rdr.onload=function(ev){var b64=ev.target.result.split(",")[1];send(b64,slipFile.type);};rdr.readAsDataURL(slipFile);}'
    + 'else{send("","");}'
    + '}'

    + 'function openWAAdmin(){if(!_waMsgAdmin){alert("Please submit first.");return;}window.open("https://wa.me/"+ADMIN_PHONE+"?text="+encodeURIComponent(_waMsgAdmin),"_blank");}'
    + 'function openWAAcct(){if(!ACCOUNTANT_PHONE){alert("Accountant number not set.");return;}if(!_waMsgAcct){alert("Please submit first.");return;}window.open("https://wa.me/"+ACCOUNTANT_PHONE+"?text="+encodeURIComponent(_waMsgAcct),"_blank");}'

    + '<\/script></body></html>';

  return HtmlService.createHtmlOutput(html).setTitle('Invoice ' + _esc_(invoiceNo));
}

// ═══════════════════════════════════════════════════════════════════
// PART D: _serveForwardToAccountant_
// ═══════════════════════════════════════════════════════════════════
function _serveForwardToAccountant_(e) {
  var token     = String(e.parameter.token   || '').trim();
  var invoiceNo = String(e.parameter.invoice || '').trim();
  var acct      = String(e.parameter.acct    || '').trim();

  function errPage(msg) {
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family:sans-serif;padding:40px;text-align:center">'
      + '<h2 style="color:#dc2626">⛔ Error</h2><p>' + msg + '</p>'
      + '</body></html>'
    ).setTitle('Error');
  }

  if (!token) return errPage('Invalid token.');
  if (!acct)  return errPage('Accountant number not set.');

  var ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
  var logSheet = ss.getSheetByName(PAYMENT_LOG_SHEET_NAME);
  var logRow   = _findPaymentLogRow_(logSheet, token);
  if (!logRow) return errPage('Invalid or expired link.');

  var webAppUrl  = ScriptApp.getService().getUrl();
  var approveUrl = String(logRow['APPROVE_URL'] || '');
  if (!approveUrl && invoiceNo) {
    approveUrl = webAppUrl
      + '?action=approveForm&invoice=' + encodeURIComponent(invoiceNo)
      + '&token=' + encodeURIComponent(token)
      + '&pendingId=' + encodeURIComponent(token)
      + '&amt=' + encodeURIComponent(String(logRow['AMOUNT'] || ''))
      + '&client=' + encodeURIComponent(String(logRow['CLIENT_NAME'] || ''));
  }

  var msg = '*🔔 PAYMENT FOR APPROVAL — CK Events*'
    + '\n━━━━━━━━━━━━━━━━━━━━'
    + '\n📋 *Invoice:* ' + String(logRow['INVOICE_NO'] || invoiceNo)
    + '\n👤 *Client:* '  + String(logRow['CLIENT_NAME'] || '')
    + '\n💵 *Amount:* RM ' + String(logRow['AMOUNT'] || '')
    + '\n🏦 *Method:* '  + String(logRow['METHOD'] || '')
    + (logRow['REF']  ? '\n🔢 *Ref No:* ' + String(logRow['REF']) : '')
    + (logRow['DATE'] ? '\n📅 *Date:* '   + String(logRow['DATE']) : '')
    + (logRow['TIME'] ? '\n🕐 *Time:* '   + String(logRow['TIME']) : '')
    + (logRow['NOTE'] ? '\n📝 *Note:* '   + String(logRow['NOTE']) : '')
    + (logRow['PROOF_LINK'] ? '\n🔗 *Slip:* ' + String(logRow['PROOF_LINK']) : '\n📎 No slip')
    + (approveUrl ? '\n\n🔐 *Approval Link:*\n' + approveUrl : '')
    + '\n━━━━━━━━━━━━━━━━━━━━'
    + '\n_Please approve or reject via the link above._';

  var waUrl = 'https://wa.me/' + encodeURIComponent(acct) + '?text=' + encodeURIComponent(msg);

  var html = '<!DOCTYPE html><html><head><meta charset="utf-8">'
    + '<meta http-equiv="refresh" content="0;url=' + _esc_(waUrl) + '">'
    + '<title>Forwarding to Accountant…</title>'
    + '<style>body{font-family:sans-serif;padding:40px;text-align:center;background:#f1f5f9}'
    + '.card{background:#fff;border-radius:16px;padding:30px;max-width:400px;margin:0 auto;box-shadow:0 2px 16px rgba(0,0,0,.07)}'
    + 'a{color:#2563eb;font-weight:700}</style>'
    + '</head><body>'
    + '<div class="card">'
    + '<div style="font-size:40px;margin-bottom:12px">💬</div>'
    + '<div style="font-weight:900;font-size:18px;margin-bottom:8px">Forwarding to Accountant…</div>'
    + '<div style="color:#64748b;font-size:13px;margin-bottom:20px">Opening WhatsApp with the payment details for approval.</div>'
    + '<a href="' + _esc_(waUrl) + '" style="display:inline-block;padding:14px 28px;background:#16a34a;color:#fff;border-radius:10px;text-decoration:none">💬 Open WhatsApp</a>'
    + '</div>'
    + '<script>setTimeout(function(){window.location.href=' + JSON.stringify(waUrl) + ';},500);<\/script>'
    + '</body></html>';

  return HtmlService.createHtmlOutput(html).setTitle('Forwarding to Accountant');
}

// ═══════════════════════════════════════════════════════════════════
// PART E: _serveAccountantApprovalPage_
// ═══════════════════════════════════════════════════════════════════
function _serveAccountantApprovalPage_(e) {
  var invoiceNo = String(e.parameter.invoice   || e.parameter.invoice_no || e.parameter.invoiceNo || '').trim();
  var token     = String(e.parameter.token     || e.parameter.pendingId  || '').trim();
  var amount    = String(e.parameter.amt       || e.parameter.amount     || '').trim();
  var client    = String(e.parameter.client    || '').trim();

  if (!token) {
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family:sans-serif;padding:40px;text-align:center">'
      + '<h2 style="color:#dc2626">⛔ Unauthorized</h2><p>Invalid or missing token.</p>'
      + '</body></html>'
    ).setTitle('Unauthorized');
  }

  var ss         = SpreadsheetApp.openById(SPREADSHEET_ID);
  var logSheet   = ss.getSheetByName(PAYMENT_LOG_SHEET_NAME);
  var logRow     = _findPaymentLogRow_(logSheet, token);
  var curStatus  = logRow ? String(logRow['APPROVAL_STATUS'] || '').toUpperCase() : '';

  // ── LAYER 1: Hard Page Lock ─────────────────────────────────────────
  // If already APPROVED or REJECTED, return a dead "Link Expired" page immediately.
  if (curStatus === 'APPROVED' || curStatus === 'REJECTED') {
    var expiredHtml = '<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
      + '<title>Link Expired</title>'
      + '<style>body{font-family:Arial,system-ui,sans-serif;background:#f1f5f9;margin:0;display:flex;align-items:center;justify-content:center;min-height:100vh;padding:16px}'
      + '.card{background:#fff;border-radius:16px;box-shadow:0 8px 30px rgba(0,0,0,.09);padding:40px 32px;max-width:440px;width:100%;text-align:center}'
      + '.icon{font-size:56px;margin-bottom:16px}'
      + 'h2{font-size:22px;font-weight:900;color:#dc2626;margin:0 0 12px}'
      + 'p{color:#64748b;font-size:14px;line-height:1.6;margin:0}'
      + '.badge{display:inline-block;margin-top:18px;padding:6px 18px;border-radius:99px;font-size:12px;font-weight:700;letter-spacing:.05em;'
      + (curStatus === 'APPROVED' ? 'background:#d1fae5;color:#065f46' : 'background:#fee2e2;color:#991b1b')
      + '}'
      + '</style></head><body>'
      + '<div class="card">'
      + '<div class="icon">⚠️</div>'
      + '<h2>Link Expired</h2>'
      + '<p>This payment has already been processed by the Admin or Accountant.</p>'
      + '<div class="badge">' + curStatus + '</div>'
      + '</div>'
      + '</body></html>';
    return HtmlService.createHtmlOutput(expiredHtml).setTitle('Link Expired');
  }

  var alreadyDone = false;
  var clientPhone = '';

  if (logRow) clientPhone = String(logRow['CLIENT_PHONE'] || '');
  if (!clientPhone && invoiceNo) {
    try {
      var masterSheet2 = ss.getSheetByName(INVOICE_MASTER_SHEET_NAME);
      var invRow2      = _findInvoiceRow_(masterSheet2, invoiceNo);
      if (invRow2) clientPhone = String(invRow2['Client_Phone'] || '');
    } catch(ex) {}
  }

  var WEB_APP_URL = ScriptApp.getService().getUrl();
  var CK_APP_URL  = 'https://canopyrentalsmalaysia.com/invoice-asli';

  var html =
    '<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<title>Approve Payment — ' + _esc_(invoiceNo) + '</title>'
    + '<style>'
    + 'body{font-family:Arial,system-ui,-apple-system,sans-serif;background:#f1f5f9;margin:0;padding:16px}'
    + '.card{max-width:560px;margin:0 auto;background:#fff;border-radius:16px;box-shadow:0 8px 30px rgba(0,0,0,.08);padding:18px}'
    + '.h{font-size:18px;font-weight:800;color:#0f172a;margin-bottom:12px}'
    + '.row{display:flex;justify-content:space-between;gap:10px;padding:8px 0;border-bottom:1px solid #eef2f7;font-size:14px}'
    + '.lbl{color:#64748b}.val{font-weight:800;color:#0f172a;text-align:right}'
    + '.btn{width:100%;border:none;border-radius:12px;padding:14px 16px;font-size:15px;font-weight:900;cursor:pointer;margin-top:10px}'
    + '.approve{background:#059669;color:#fff}.reject{background:#dc2626;color:#fff}.wa-btn{background:#16a34a;color:#fff}'
    + '.pin-field{width:100%;padding:11px 13px;border:1.5px solid #e2e8f0;border-radius:10px;font-size:16px;outline:none;margin:10px 0;text-align:center;letter-spacing:4px}'
    + '.pin-field:focus{border-color:#2563eb}'
    + 'textarea{width:100%;padding:10px;border:1.5px solid #e2e8f0;border-radius:10px;font-size:13px;resize:vertical;outline:none;margin-top:8px}'
    + '</style></head><body>'
    + '<div class="card">'
    + '<div class="h">💳 Approve / Reject Payment</div>'
    + '<div class="row"><div class="lbl">Invoice No</div><div class="val">' + _esc_(invoiceNo) + '</div></div>'
    + '<div class="row"><div class="lbl">Client</div><div class="val">' + _esc_(client) + '</div></div>'
    + '<div class="row"><div class="lbl">Amount</div><div class="val">RM ' + _esc_(amount) + '</div></div>'
    + '<div class="row"><div class="lbl">Status</div><div class="val">' + _esc_(curStatus || 'PENDING') + '</div></div>'

    + '<div style="margin-top:14px"><label style="font-size:12px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:.04em">🔐 Approval PIN</label>'
    + '<input type="password" class="pin-field" id="pinInput" placeholder="••••••" maxlength="10">'
    + '</div>'

    + (alreadyDone
        ? '<div style="margin-top:12px;padding:12px;background:#fef3c7;border-radius:10px;font-weight:700;color:#92400e;text-align:center">⚠️ Already processed: ' + _esc_(curStatus) + '</div>'
        : '<button id="btnApprove" class="btn approve">✅ Approve Payment</button>'
          + '<button id="btnShowReject" class="btn reject" style="margin-top:8px">❌ Reject Payment</button>'
          + '<div id="rejectSection" style="display:none;margin-top:10px">'
          + '<label style="font-size:12px;font-weight:700;color:#64748b;text-transform:uppercase">Reason for Rejection <span style="color:#dc2626">*</span></label>'
          + '<textarea id="rejectReason" rows="3" placeholder="e.g. Slip unclear, amount mismatch, wrong reference..."></textarea>'
          + '<button id="btnConfirmReject" class="btn reject" style="margin-top:8px">Confirm Reject</button>'
          + '</div>'
      )

    + '<div id="status" style="margin-top:10px;padding:10px;border-radius:12px;background:#f8fafc;border:1px solid #e2e8f0;font-size:13px;color:#0f172a">Ready.</div>'
    + '<div id="waNotifyWrapper" style="display:none;margin-top:10px">'
    + '<button id="btnNotifyClient" class="btn wa-btn" onclick="notifyClient()">💬 Notify Client via WhatsApp</button>'
    + '<div id="waNotifyNote" style="font-size:11px;color:#94a3b8;text-align:center;margin-top:4px"></div>'
    + '</div>'
    + '<div style="margin-top:10px;font-size:12px;color:#64748b">After approval, you will be redirected to generate the updated PDF.</div>'
    + '</div>'
    + '<script>'
    + 'var INVOICE_NO=' + JSON.stringify(invoiceNo) + ';'
    + 'var TOKEN=' + JSON.stringify(token) + ';'
    + 'var AMOUNT=' + JSON.stringify(amount) + ';'
    + 'var CLIENT=' + JSON.stringify(client) + ';'
    + 'var CLIENT_PHONE=' + JSON.stringify(clientPhone) + ';'
    + 'var WEB_APP=' + JSON.stringify(WEB_APP_URL) + ';'
    + 'var CK_APP=' + JSON.stringify(CK_APP_URL) + ';'
    + 'var ALREADY_DONE=' + JSON.stringify(alreadyDone) + ';'
    + 'var _rejectReason="";'
    + 'function $s(s){return document.querySelector(s);}'
    + 'function setStatus(html){$s("#status").innerHTML=html;}'
    + 'async function post(action,payload){'
    + '  var res=await fetch(WEB_APP,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify(Object.assign({action:action},payload||{}))});'
    + '  return res.json();'
    + '}'
    + 'function getPin(){return ($s("#pinInput")||{}).value||"";}'
    + 'function topRedirect(url){try{if(window.top){window.top.location.replace(url);return;}}catch(e){}window.location.replace(url);}'

    + 'async function approve(){'
    + '  var pin=getPin();'
    + '  if(!pin){setStatus("<span style=\'color:#dc2626;font-weight:900\'>⚠️ Please enter your PIN.</span>");return;}'
    + '  if(ALREADY_DONE){setStatus("<span style=\'color:#dc2626\'>Already processed.</span>");return;}'
    + '  $s("#btnApprove").disabled=true;'
    + '  setStatus("⏳ Approving…");'
    + '  var j;'
    + '  try{j=await post("confirmPayment",{token:TOKEN,invoiceNo:INVOICE_NO,pin:pin,paymentData:{amount:AMOUNT}});}'
    + '  catch(e){setStatus("<span style=\'color:#dc2626\'>❌ Network error.</span>");$s("#btnApprove").disabled=false;return;}'
    + '  if(!j||j.status!=="success"){setStatus("<span style=\'color:#dc2626\'>❌ "+(j&&j.message?j.message:"Unknown error")+"</span>");$s("#btnApprove").disabled=false;return;}'
    + '  setStatus("<span style=\'color:#059669;font-weight:900\'>✅ Approved!</span> Redirecting to generate updated PDF…");'
    + '  var target=CK_APP+"?mode=ACCOUNTANT_PDF&invoiceNo="+encodeURIComponent(INVOICE_NO)+"&token="+encodeURIComponent(TOKEN)+"&gasUrl="+encodeURIComponent(WEB_APP);'
    + '  topRedirect(target);'
    + '}'

    + 'function showReject(){$s("#rejectSection").style.display="block";$s("#rejectSection").scrollIntoView({behavior:"smooth"});}'

    + 'async function confirmReject(){'
    + '  var pin=getPin();'
    + '  if(!pin){setStatus("<span style=\'color:#dc2626\'>⚠️ Please enter your PIN.</span>");return;}'
    + '  var reason=($s("#rejectReason")||{}).value.trim();'
    + '  if(!reason){setStatus("<span style=\'color:#dc2626\'>⚠️ Please enter a rejection reason.</span>");return;}'
    + '  if(ALREADY_DONE){setStatus("<span style=\'color:#dc2626\'>Already processed.</span>");return;}'
    + '  $s("#btnConfirmReject").disabled=true;'
    + '  setStatus("⏳ Rejecting…");'
    + '  var j;'
    + '  try{j=await post("rejectPayment",{token:TOKEN,invoiceNo:INVOICE_NO,pin:pin,reason:reason});}'
    + '  catch(e){setStatus("<span style=\'color:#dc2626\'>❌ Network error.</span>");$s("#btnConfirmReject").disabled=false;return;}'
    + '  if(!j||j.status!=="success"){setStatus("<span style=\'color:#dc2626\'>❌ "+(j&&j.message?j.message:"Unknown error")+"</span>");$s("#btnConfirmReject").disabled=false;return;}'
    + '  _rejectReason=reason;'
    + '  setStatus("<span style=\'color:#dc2626;font-weight:900\'>❌ Payment rejected.</span> Reason: "+reason);'
    + '  $s("#waNotifyWrapper").style.display="block";'
    + '  if(!CLIENT_PHONE)$s("#waNotifyNote").textContent="Client phone not on file — notify manually.";'
    + '}'

    + 'function notifyClient(){'
    + '  if(!CLIENT_PHONE){alert("Client phone not available.");return;}'
    + '  var msg="Payment Rejected — CK Events\\nInvoice: "+INVOICE_NO+"\\nReason: "+_rejectReason+"\\nYour payment is not verified/not reflected. Please contact admin or resubmit slip using the invoice link.";'
    + '  window.open("https://wa.me/"+CLIENT_PHONE+"?text="+encodeURIComponent(msg),"_blank");'
    + '}'

    + (alreadyDone ? '' :
        'if($s("#btnApprove"))$s("#btnApprove").addEventListener("click",approve);'
        + 'if($s("#btnShowReject"))$s("#btnShowReject").addEventListener("click",showReject);'
        + 'if($s("#btnConfirmReject"))$s("#btnConfirmReject").addEventListener("click",confirmReject);'
      )
    + '<\/script>'
    + '</body></html>';

  return HtmlService.createHtmlOutput(html).setTitle('Approve Payment');
}

// ═══════════════════════════════════════════════════════════════════
// _buildInvoiceHtmlForPdf_ — v98: dynamic payment rows in totals box
// ═══════════════════════════════════════════════════════════════════
function _buildInvoiceHtmlForPdf_(row) {
  var items = [];
  try { items = JSON.parse(row['Invoice_Items_JSON'] || '[]'); } catch(e) {}

  var docTitle = (row['Document_Type'] || '').toUpperCase() === 'INVOICE'
    ? 'INVOICE / OFFICIAL RECEIPT'
    : 'OFFICIAL QUOTATION / INVOICE';

  var subtotal   = parseFloat(String(row['Subtotal']    || '0').replace(/[^0-9.]/g,'')) || 0;
  var grandTotal = parseFloat(String(row['Grand_Total'] || row['Subtotal'] || '0').replace(/[^0-9.]/g,'')) || 0;
  var paid       = parseFloat(String(row['Payment_made']|| '0').replace(/[^0-9.]/g,'')) || 0;
  var balance    = (function(){ var bd = parseFloat(String(row['Balance_Due'] || '').replace(/[^0-9.]/g,'')); return (!isNaN(bd) && bd >= 0) ? bd : Math.max(0, grandTotal - paid); })();
  var isInvoice  = (row['Document_Type'] || '').toUpperCase() === 'INVOICE';

  // ── v98: Parse Payment_History_JSON ───────────────────────────
  var paymentHistory = [];
  try {
    var histRaw = row['Payment_History_JSON'] || '[]';
    paymentHistory = JSON.parse(histRaw);
    if (!Array.isArray(paymentHistory)) paymentHistory = [];
  } catch(e) { paymentHistory = []; }

  function fmt(n) { return parseFloat(n || 0).toFixed(2); }

  var itemsHtml = '';
  for (var i = 0; i < items.length; i++) {
    var it = items[i], qty = parseFloat(it.qty||0), rt = parseFloat(it.rate||0), am = qty*rt;
    itemsHtml += '<tr style="border-bottom:1px solid #e2e8f0">'
      + '<td style="width:50%;padding:6px 8px 6px 12px;font-size:11px;font-weight:700;vertical-align:top">' + _esc_(it.item||'') + '</td>'
      + '<td style="width:15%;padding:6px 8px;font-size:11px;text-align:right;font-family:monospace;vertical-align:top">' + fmt(rt) + '</td>'
      + '<td style="width:10%;padding:6px 8px;font-size:11px;text-align:center;font-family:monospace;vertical-align:top">' + qty + '</td>'
      + '<td style="width:25%;padding:6px 12px 6px 8px;font-size:11px;text-align:right;font-weight:700;font-family:monospace;vertical-align:top">' + fmt(am) + '</td>'
      + '</tr>';
  }

  // ── v98: Build totals box with dynamic payment rows ───────────
  // Layout (invoice mode):
  //   Subtotal: RM xxx
  //   Grand Total: RM xxx
  //   [if payments exist]
  //     1st Payment (DD Mon): - RM xxx
  //     2nd Payment (DD Mon): - RM xxx
  //     ...
  //   Total Paid: - RM xxx
  //   Balance Due: RM xxx  (or "Fully Paid" in green)

  var totalsHtml = '<div style="display:flex;justify-content:space-between;padding:5px 0;border-bottom:1px solid #000;font-size:11px"><span style="font-weight:700">Subtotal</span><span style="font-family:monospace">RM ' + fmt(subtotal) + '</span></div>';

  if (isInvoice) {
    // Grand Total row
    totalsHtml += '<div style="display:flex;justify-content:space-between;padding:5px 0;border-bottom:1px solid #e2e8f0;font-size:11px"><span>Grand Total</span><span style="font-family:monospace">RM ' + fmt(grandTotal) + '</span></div>';

    // Individual payment rows (only if there are approved payments)
    if (paymentHistory.length > 0) {
      var ordinals = ['1st','2nd','3rd','4th','5th','6th','7th','8th','9th','10th'];
      for (var pi = 0; pi < paymentHistory.length; pi++) {
        var entry  = paymentHistory[pi];
        var pAmt   = parseFloat(String(entry.amount || '0').replace(/[^0-9.]/g,'')) || 0;
        var pDate  = entry.date ? _formatDateMY_(entry.date) : '';
        var pTime  = entry.timeFormatted || (entry.time ? _formatTimeMY_(entry.time) : '');
        var pMeta  = '';
        if (pDate && pTime) pMeta = ' (on ' + pDate + ' @ ' + pTime + ')';
        else if (pDate)     pMeta = ' (on ' + pDate + ')';
        var pLabel = (ordinals[pi] || (pi+1)+'th') + ' Payment' + pMeta;
        totalsHtml += '<div style="display:flex;justify-content:space-between;padding:4px 0 4px 10px;border-bottom:1px solid #f1f5f9;font-size:10px;color:#64748b;font-style:italic">'
          + '<span>' + _esc_(pLabel) + '</span>'
          + '<span style="font-family:monospace;color:#059669">- RM ' + fmt(pAmt) + '</span>'
          + '</div>';
      }
    }

    // Total Paid row
    totalsHtml += '<div style="display:flex;justify-content:space-between;padding:5px 0;border-bottom:1px solid #e2e8f0;font-size:11px;font-weight:700"><span>Total Paid</span><span style="font-family:monospace;color:#059669">- RM ' + fmt(paid) + '</span></div>';

    // Balance Due / Fully Paid row
    if (balance <= 0.01) {
      totalsHtml += '<div style="display:flex;justify-content:space-between;padding:5px 0;font-size:11px;font-weight:700;color:#059669"><span>Balance Due</span><span style="font-family:monospace">✅ Fully Paid</span></div>';
    } else {
      totalsHtml += '<div style="display:flex;justify-content:space-between;padding:5px 0;font-size:11px;font-weight:700"><span>Balance Due</span><span style="font-family:monospace">RM ' + fmt(balance) + '</span></div>';
    }

    // Big Total box (shows Balance Due or 0)
    var displayTotal = balance <= 0.01 ? 0 : balance;
    var totalBoxContent = balance <= 0.01
      ? '<span style="font-size:11px;text-transform:uppercase">Total Due</span><span style="font-family:monospace;color:#059669">✅ Fully Paid</span>'
      : '<span style="font-size:11px;text-transform:uppercase">Total Due</span><span style="font-family:monospace">RM ' + fmt(displayTotal) + '</span>';
    totalsHtml = totalsHtml + '<div style="display:flex;justify-content:space-between;font-size:14px;font-weight:700;background:#d4af37;padding:8px 10px;margin-top:4px;border-left:4px solid #000">' + totalBoxContent + '</div>';
  } else {
    // Quotation: just subtotal + total
    totalsHtml += '<div style="display:flex;justify-content:space-between;font-size:14px;font-weight:700;background:#d4af37;padding:8px 10px;margin-top:4px;border-left:4px solid #000"><span style="font-size:11px;text-transform:uppercase">Total</span><span style="font-family:monospace">RM ' + fmt(grandTotal) + '</span></div>';
  }

  var evDateFrom  = row['Event_Date_from'] || '';
  var evDateTo    = row['Event_Date_To']   || '';
  var evDateStr   = (evDateFrom && evDateTo && evDateFrom !== evDateTo) ? (evDateFrom + ' - ' + evDateTo) : (evDateFrom || '-');
  var setupStr    = (row['Setup_Date'] && row['Setup_time'])         ? (row['Setup_Date']     + ' @ ' + row['Setup_time'])     : (row['Setup_Date']     || '-');
  var dismantleStr= (row['Dismantle_Date'] && row['Dismantle_Time']) ? (row['Dismantle_Date'] + ' @ ' + row['Dismantle_Time']) : (row['Dismantle_Date'] || '-');
  var address     = _esc_(String(row['Bill_To_Address'] || '').replace(/, /g, '\n'));

  // Notes — only Special_Notes (no old payment history appended)
  var notesHtml = '';
  var cleanSpecialNotes = String(row['Special_Notes'] || '').split('--- PAYMENT HISTORY ---')[0].trim();
  if (cleanSpecialNotes || row['Remarks']) {
    notesHtml = '<div style="margin-bottom:20px">';
    if (cleanSpecialNotes) notesHtml += '<div style="font-size:9px;font-weight:700;background:#d4af37;color:#000;text-transform:uppercase;letter-spacing:.1em;padding:2px 8px;display:inline-block;margin-bottom:4px">Special Notes</div><p style="font-size:11px;color:#334155;white-space:pre-line;background:#fefce8;padding:8px;border-left:2px solid #d4af37">' + _esc_(cleanSpecialNotes) + '</p>';
    if (row['Remarks']) notesHtml += '<div style="font-size:9px;font-weight:700;background:#d4af37;color:#000;text-transform:uppercase;letter-spacing:.1em;padding:2px 8px;display:inline-block;margin-bottom:4px;margin-top:8px">Remarks</div><p style="font-size:11px;color:#334155;font-style:italic;white-space:pre-line">' + _esc_(row['Remarks']) + '</p>';
    notesHtml += '</div>';
  }

  var termsData = [
    'Availability & Booking Confirmation:All quotations are subject to availability until confirmed with payment.',
    'Booking & Deposit:A 50% non-refundable deposit is required to secure the booking.',
    'Balance Payment:Remaining 50% balance must be paid in full at least 7 days before the event date.',
    'Payment Confirmation:Please provide proof of payment once made.',
    'Cancellation Policy:All payments made are strictly non-refundable in the event of cancellation.',
    'Delivery & Collection Timing (Shared Transport):Times are approximate, subject to traffic and site conditions.',
    'Premium Timing Option:Dedicated transport & team available at additional cost — must be requested in advance.',
    'Venue Access & Readiness:Client is responsible for ensuring clear access and venue readiness at agreed time.',
    'Additional Charges:Any last-minute additions, overtime, or after-hours services not stated will be billed separately.',
    'Damage, Loss & Liability:Client is fully responsible for any loss or damage to rented items during rental period.',
    'Force Majeure:We shall not be held liable for delays due to circumstances beyond our control.',
    'Agreement:Confirmation of booking and/or payment signifies acceptance of all the above terms.'
  ];
  var termsHtml = '';
  for (var t = 0; t < termsData.length; t++) {
    var tp = termsData[t].split(':'), th = tp.shift(), tt = tp.join(':');
    termsHtml += '<div style="display:flex;align-items:flex-start;margin-bottom:3px;font-size:9px;line-height:1.3"><div style="min-width:18px;font-weight:700">' + (t+1) + '.</div><div style="flex:1;text-align:justify"><strong>' + _esc_(th) + ':</strong>' + _esc_(tt) + '</div></div>';
  }

  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#1e293b;background:#fff}</style></head><body>'
    + '<div style="width:794px;min-height:1123px;padding:40px;margin:0 auto;position:relative;background:#fff">'
    + '<div style="position:absolute;top:0;left:0;right:0;height:12px;background:#000;border-bottom:2px solid #d4af37"></div>'
    + '<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-top:20px;margin-bottom:20px">'
    + '<div style="width:60%"><h1 style="font-size:22px;font-weight:900;letter-spacing:.1em;color:#c59d2e;text-transform:uppercase">' + _esc_(docTitle) + '</h1>'
    + '<div style="margin-top:14px"><div style="display:flex;margin-bottom:6px"><span style="font-weight:700;width:90px;text-transform:uppercase;font-size:10px">Date</span><span>' + _esc_(row['Date']||'') + '</span></div>'
    + '<div style="display:flex"><span style="font-weight:700;width:90px;text-transform:uppercase;font-size:10px">Invoice No.</span><span style="font-weight:700">' + _esc_(row['Invoice_No']||'') + '</span></div></div></div>'
    + '<div style="width:40%;text-align:right"><div style="font-weight:700;font-size:15px;text-transform:uppercase">CK Ballroom<br>Rawang Resources</div>'
    + '<div style="font-size:9px;font-weight:700;color:#c59d2e;text-transform:uppercase;letter-spacing:.1em;margin:4px 0">Operating as CK Event Management</div>'
    + '<div style="font-size:10px;color:#475569;line-height:1.5">No. 5, Jalan STR 1, Saujana Teknologi Park, 48000 Rawang, Selangor<br><strong>Reg No:</strong> 201903353628<br><strong>Tel:</strong> 012-659 7702</div></div></div>'
    + '<div style="height:1px;background:#000;margin-bottom:20px;position:relative"><div style="position:absolute;right:0;top:0;height:100%;width:33%;background:#d4af37"></div></div>'
    + '<div style="display:flex;gap:24px;margin-bottom:20px">'
    + '<div style="width:50%"><div style="font-size:9px;font-weight:700;background:#d4af37;color:#000;text-transform:uppercase;padding:2px 8px;display:inline-block;margin-bottom:6px">Bill To</div>'
    + '<div style="font-weight:700;font-size:14px;border-left:4px solid #d4af37;padding-left:10px;margin-bottom:4px">' + _esc_(row['Bill_To_Company']||'') + '</div>'
    + (row['Bill_To_Attn'] ? '<div style="font-size:10px;color:#475569;padding-left:14px;margin-bottom:4px">' + _esc_(row['Bill_To_Attn']) + '</div>' : '')
    + '<div style="font-size:10px;color:#475569;white-space:pre-line;padding-left:14px">' + address + '</div></div>'
    + '<div style="width:50%;background:#f8fafc;padding:12px;border:1px solid #e2e8f0">'
    + '<div style="margin-bottom:8px;padding-bottom:6px;border-bottom:1px solid #e2e8f0"><div style="font-size:9px;font-weight:700;background:#d4af37;color:#000;text-transform:uppercase;padding:1px 8px;display:inline-block;margin-bottom:4px">Event Date</div><div style="font-size:11px;font-weight:600">' + _esc_(evDateStr) + '</div></div>'
    + '<div style="margin-bottom:8px;padding-bottom:6px;border-bottom:1px solid #e2e8f0"><div style="font-size:9px;font-weight:700;background:#d4af37;color:#000;text-transform:uppercase;padding:1px 8px;display:inline-block;margin-bottom:4px">Venue</div><div style="font-size:10px">' + _esc_(row['Venue']||'-') + '</div></div>'
    + '<div style="margin-bottom:8px;padding-bottom:6px;border-bottom:1px solid #e2e8f0"><div style="font-size:9px;font-weight:700;background:#d4af37;color:#000;text-transform:uppercase;padding:1px 8px;display:inline-block;margin-bottom:4px">Setup</div><div style="font-size:10px">' + _esc_(setupStr) + '</div></div>'
    + '<div><div style="font-size:9px;font-weight:700;background:#d4af37;color:#000;text-transform:uppercase;padding:1px 8px;display:inline-block;margin-bottom:4px">Dismantle</div><div style="font-size:10px">' + _esc_(dismantleStr) + '</div></div>'
    + '</div></div>'
    + '<table style="width:100%;border-collapse:collapse;margin-bottom:20px"><thead><tr style="background:#d4af37;color:#000;border-bottom:2px solid #000">'
    + '<th style="width:50%;padding:6px 8px 6px 12px;text-align:left;font-weight:700;text-transform:uppercase;font-size:10px">Description</th>'
    + '<th style="width:15%;padding:6px 8px;text-align:right;font-weight:700;text-transform:uppercase;font-size:10px">Rate</th>'
    + '<th style="width:10%;padding:6px 8px;text-align:center;font-weight:700;text-transform:uppercase;font-size:10px">Qty</th>'
    + '<th style="width:25%;padding:6px 12px 6px 8px;text-align:right;font-weight:700;text-transform:uppercase;font-size:10px">Amount</th>'
    + '</tr></thead><tbody>' + itemsHtml + '</tbody></table>'
    + '<div style="display:flex;justify-content:flex-end;margin-bottom:20px"><div style="width:280px">' + totalsHtml + '</div></div>'
    + notesHtml
    + '<div style="border-top:2px solid #e2e8f0;padding-top:14px;margin-bottom:14px"><div style="display:flex;gap:24px">'
    + '<div style="width:50%"><div style="font-size:9px;font-weight:700;background:#d4af37;color:#000;text-transform:uppercase;padding:2px 8px;display:inline-block;margin-bottom:8px">Payment Details</div>'
    + '<div style="font-size:10px;line-height:1.8"><div style="display:flex"><span style="width:80px;color:#64748b">Bank:</span><span style="font-weight:700">CIMB Bank</span></div>'
    + '<div style="display:flex"><span style="width:80px;color:#64748b">Account:</span><span style="font-weight:700">8010399709</span></div>'
    + '<div style="display:flex"><span style="width:80px;color:#64748b">Name:</span><span>CK Ballroom Rawang Resources</span></div>'
    + '<div style="display:flex"><span style="width:80px;color:#64748b">SWIFT:</span><span>CIBBMYKL</span></div></div></div>'
    + '<div style="width:50%;text-align:right;font-size:10px;color:#64748b;display:flex;flex-direction:column;justify-content:flex-end">'
    + '<p style="font-weight:700;font-size:13px;color:#c59d2e;font-style:italic;margin-bottom:4px">Thank you for your business!</p>'
    + '<p>Please send proof of payment to the number above.</p></div></div></div>'
    + '<div style="margin-bottom:14px"><div style="font-size:9px;font-weight:700;background:#d4af37;color:#000;text-transform:uppercase;padding:3px 12px;display:inline-block;margin-bottom:6px">TERMS &amp; CONDITIONS</div>'
    + '<div>' + termsHtml + '</div></div>'
    + '<div style="font-size:9px;color:#94a3b8;text-align:center;padding-bottom:20px">Computer generated document. No signature required.</div>'
    + '<div style="position:absolute;bottom:0;left:0;right:0;height:12px;background:#000;border-top:2px solid #d4af37"></div>'
    + '</div></body></html>';
}

// ═══════════════════════════════════════════════════════════════════
// _generateAndSaveInvoicePdf_
// ═══════════════════════════════════════════════════════════════════
function _generateAndSaveInvoicePdf_(row) {
  try {
    var invoiceNo  = String(row['Invoice_No'] || 'invoice').trim();
    var clientName = String(row['Bill_To_Company'] || 'Unknown').replace(/[\/\\:*?"<>|]/g,'').trim() || 'Unknown';
    var stamp      = Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'yyyyMMdd_HHmm');
    var fileName   = invoiceNo + '_updated_' + stamp;

    var html = _buildInvoiceHtmlForPdf_(row);
    var pdfBlob;
    var tempHtmlFileId = null;

    try {
      var boundary  = '-------314159265358979323846';
      var metadata  = JSON.stringify({ name: fileName, mimeType: 'application/vnd.google-apps.document' });
      var payload   = '--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + metadata + '\r\n'
        + '--' + boundary + '\r\nContent-Type: text/html\r\n\r\n' + html + '\r\n'
        + '--' + boundary + '--';
      var uploadResp = UrlFetchApp.fetch(
        'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart',
        { method: 'POST', headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), 'Content-Type': 'multipart/related; boundary="' + boundary + '"' }, payload: payload, muteHttpExceptions: true }
      );
      var uploadJson = JSON.parse(uploadResp.getContentText());
      tempHtmlFileId = uploadJson.id;
      if (!tempHtmlFileId) throw new Error('Upload failed');

      var exportUrl = 'https://www.googleapis.com/drive/v3/files/' + tempHtmlFileId + '/export?mimeType=application/pdf';
      var response  = UrlFetchApp.fetch(exportUrl, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
      try { DriveApp.getFileById(tempHtmlFileId).setTrashed(true); tempHtmlFileId = null; } catch(x) {}

      if (response.getResponseCode() === 200) {
        pdfBlob = response.getBlob();
        pdfBlob.setName(fileName + '.pdf');
        pdfBlob.setContentType('application/pdf');
      } else {
        throw new Error('Drive export HTTP ' + response.getResponseCode());
      }
    } catch(exportErr) {
      Logger.log('HTML→PDF export failed: ' + exportErr.message);
      if (tempHtmlFileId) { try { DriveApp.getFileById(tempHtmlFileId).setTrashed(true); } catch(x) {} }
      pdfBlob = _docAppFallbackPdf_(row, fileName);
      if (!pdfBlob) throw new Error('Both PDF methods failed');
    }

    var rootFolder   = DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);
    var cFolders     = rootFolder.getFoldersByName(clientName);
    var clientFolder = cFolders.hasNext() ? cFolders.next() : rootFolder.createFolder(clientName);
    var iFolders     = clientFolder.getFoldersByName(invoiceNo);
    var invFolder    = iFolders.hasNext() ? iFolders.next() : clientFolder.createFolder(invoiceNo);

    var pdfFile = invFolder.createFile(pdfBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return pdfFile.getUrl();
  } catch(e) {
    Logger.log('_generateAndSaveInvoicePdf_ error: ' + e.message);
    return '';
  }
}

// ═══════════════════════════════════════════════════════════════════
// _docAppFallbackPdf_
// ═══════════════════════════════════════════════════════════════════
function _docAppFallbackPdf_(row, fileName) {
  var tempDocId = null;
  try {
    function fmt(n)  { return parseFloat(n || 0).toFixed(2); }
    function safe(s) { return String(s || ''); }
    var doc   = DocumentApp.create('TEMP_' + fileName);
    tempDocId = doc.getId();
    var body  = doc.getBody();
    var items = [];
    try { items = JSON.parse(row['Invoice_Items_JSON'] || '[]'); } catch(e) {}
    var subtotal   = parseFloat(String(row['Subtotal']     || '0').replace(/[^0-9.]/g,'')) || 0;
    var grandTotal = parseFloat(String(row['Grand_Total']  || '0').replace(/[^0-9.]/g,'')) || 0;
    var paid       = parseFloat(String(row['Payment_made'] || '0').replace(/[^0-9.]/g,'')) || 0;
    var balance    = parseFloat(String(row['Balance_Due']  || '0').replace(/[^0-9.]/g,'')) || Math.max(0, grandTotal - paid);
    var isInvoice  = safe(row['Document_Type']).toUpperCase() === 'INVOICE';
    body.setPageWidth(595.28); body.setMarginTop(36); body.setMarginBottom(36); body.setMarginLeft(50); body.setMarginRight(50);
    var tp = body.appendParagraph(isInvoice ? 'INVOICE / OFFICIAL RECEIPT' : 'OFFICIAL QUOTATION / INVOICE');
    tp.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph('CK Ballroom Rawang Resources — CK Event Management').setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendTable([['Invoice No.',safe(row['Invoice_No'])],['Date',safe(row['Date'])],['Client',safe(row['Bill_To_Company'])]]);
    body.appendParagraph('');
    var iRows = [['Description','Rate (RM)','Qty','Amount (RM)']];
    for (var i = 0; i < items.length; i++) { var it=items[i],q=parseFloat(it.qty||0),r=parseFloat(it.rate||0); iRows.push([safe(it.item),fmt(r),String(q),fmt(q*r)]); }
    body.appendTable(iRows);
    body.appendParagraph('');
    var tRows = [['Subtotal','RM '+fmt(subtotal)]];
    if (isInvoice) {
      tRows.push(['Grand Total','RM '+fmt(grandTotal)]);
      // Add individual payment rows in fallback too
      var phFallback = [];
      try { phFallback = JSON.parse(row['Payment_History_JSON'] || '[]'); } catch(e) {}
      var ordFallback = ['1st','2nd','3rd','4th','5th','6th','7th','8th','9th','10th'];
      for (var pi = 0; pi < phFallback.length; pi++) {
        var pAmt = parseFloat(String(phFallback[pi].amount||'0').replace(/[^0-9.]/g,''))||0;
        var pDate = phFallback[pi].date ? _formatDateMY_(phFallback[pi].date) : '';
        tRows.push([(ordFallback[pi]||'#'+pi)+' Payment'+(pDate?' ('+pDate+')':''),'- RM '+fmt(pAmt)]);
      }
      tRows.push(['Total Paid','- RM '+fmt(paid)]);
      tRows.push(['Balance Due', balance<=0.01 ? '✅ Fully Paid' : 'RM '+fmt(balance)]);
    } else {
      tRows.push(['Total','RM '+fmt(grandTotal)]);
    }
    body.appendTable(tRows);
    if (row['Special_Notes']) body.appendParagraph('Special Notes: ' + safe(row['Special_Notes']).split('--- PAYMENT HISTORY ---')[0].trim());
    if (row['Remarks']) body.appendParagraph('Remarks: ' + safe(row['Remarks']));
    doc.saveAndClose();
    var docFile = DriveApp.getFileById(tempDocId);
    var blob    = docFile.getAs(MimeType.PDF);
    blob.setName(fileName + '.pdf');
    try { docFile.setTrashed(true); } catch(x) {}
    return blob;
  } catch(e) {
    Logger.log('_docAppFallbackPdf_ error: ' + e.message);
    if (tempDocId) { try { DriveApp.getFileById(tempDocId).setTrashed(true); } catch(x) {} }
    return null;
  }
}

// ═══════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════
function _genToken_() { return Math.random().toString(36).slice(2, 10).toUpperCase(); }

function _jsonOk_(obj) {
  obj.status = 'success';
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function _jsonErr_(msg) {
  return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: msg })).setMimeType(ContentService.MimeType.JSON);
}
function _esc_(str) {
  return String(str || '').replace(/\\/g,'\\\\').replace(/'/g,"\\'").replace(/"/g,'\\"').replace(/\r?\n/g,' ').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
function _findInvoiceRow_(sheet, invoiceNo) {
  if (!sheet || !invoiceNo) return null;
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var invCol  = headers.indexOf('Invoice_No');
  if (invCol < 0) return null;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][invCol]).trim() === String(invoiceNo).trim()) {
      var obj = {};
      for (var j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j];
      return obj;
    }
  }
  return null;
}
function _findPaymentLogRow_(sheet, token) {
  if (!sheet || !token) return null;
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var tokCol  = headers.indexOf('TOKEN'); if (tokCol < 0) tokCol = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][tokCol]).trim() === String(token).trim()) {
      var obj = {};
      for (var j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j];
      return obj;
    }
  }
  return null;
}
function _updatePaymentLogStatus_(sheet, token, newStatus, receiptLink, rejectReason) {
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var tokCol  = headers.indexOf('TOKEN');           if (tokCol < 0) tokCol = 0;
  var stCol   = headers.indexOf('APPROVAL_STATUS');
  var recCol  = headers.indexOf('RECEIPT_LINK');
  var atCol   = headers.indexOf('APPROVED_AT');
  var rrCol   = headers.indexOf('REJECT_REASON');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][tokCol]).trim() !== String(token).trim()) continue;
    if (stCol  >= 0) sheet.getRange(i + 1, stCol  + 1).setValue(newStatus);
    if (atCol  >= 0) sheet.getRange(i + 1, atCol  + 1).setValue(new Date().toISOString());
    if (receiptLink  && recCol >= 0) sheet.getRange(i + 1, recCol + 1).setValue(receiptLink);
    if (rejectReason && rrCol  >= 0) sheet.getRange(i + 1, rrCol  + 1).setValue(rejectReason);
    break;
  }
}
function _setSheetCell_(sheet, rowNum, headers, colName, value) {
  var idx = headers.indexOf(colName);
  if (idx >= 0) sheet.getRange(rowNum, idx + 1).setValue(value);
}
function _ensurePaymentLogHeaders_(logSheet) {
  var data    = logSheet.getDataRange().getValues();
  var headers = data[0] || [];
  if (!headers.length || String(headers[0]).trim() === '') {
    var plHeaders = [
      'TOKEN','INVOICE_NO','CLIENT_NAME','AMOUNT','DATE','TIME',
      'METHOD','REF','NOTE','PROOF_LINK','APPROVAL_STATUS','SOURCE',
      'SUBMITTED_AT','APPROVED_AT','APPROVED_BY','RECEIPT_LINK',
      'SENT_AT','SENT_BY','APPROVE_URL','REJECT_REASON'
    ];
    logSheet.clearContents();
    logSheet.appendRow(plHeaders);
  }
}
function _uploadProofSlip_(base64Data, mime, invoiceNo, token) {
  try {
    var rootFolder   = DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);
    var targetFolder = rootFolder;
    try {
      var allClientFolders = rootFolder.getFolders();
      while (allClientFolders.hasNext()) {
        var cf = allClientFolders.next();
        var invFolders = cf.getFoldersByName(invoiceNo);
        if (invFolders.hasNext()) { targetFolder = invFolders.next(); break; }
      }
    } catch(fe) {}
    var ext      = mime.indexOf('pdf') >= 0 ? 'pdf' : (mime.indexOf('png') >= 0 ? 'png' : 'jpg');
    var filename = 'PaymentSlip_' + invoiceNo + '_' + token + '.' + ext;
    var bytes    = Utilities.base64Decode(base64Data);
    var blob     = Utilities.newBlob(bytes, mime, filename);
    var file     = targetFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch(e) {
    Logger.log('Proof upload error: ' + e.message);
    return '';
  }
}

// ═══════════════════════════════════════════════════════════════════
// TEST FUNCTIONS
// ═══════════════════════════════════════════════════════════════════
function testSetup() {
  Logger.log('=== CK Events GAS v98 Test ===');
  try { var ss = SpreadsheetApp.openById(SPREADSHEET_ID); Logger.log('✅ Spreadsheet: ' + ss.getName()); } catch(e) { Logger.log('❌ Spreadsheet error: ' + e.message); }
  try { var folder = DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID); Logger.log('✅ Drive folder: ' + folder.getName()); } catch(e) { Logger.log('❌ Drive folder error: ' + e.message); }
  Logger.log('APPROVAL_PIN is set to: ' + APPROVAL_PIN);
  Logger.log('Payment_History_JSON column: enabled in DB_HEADERS ✅');
  Logger.log('=== Test complete ===');
}

function testPdfGeneration() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var masterSheet = ss.getSheetByName(INVOICE_MASTER_SHEET_NAME);
  var rows = masterSheet.getDataRange().getValues();
  var headers = rows[0];
  if (rows.length < 2) { Logger.log('ERROR: No invoice rows found'); return; }
  var row = {};
  for (var h = 0; h < headers.length; h++) row[headers[h]] = rows[1][h];
  Logger.log('Testing PDF for invoice: ' + row['Invoice_No']);
  // Attach fresh payment history
  var logSheet = ss.getSheetByName(PAYMENT_LOG_SHEET_NAME);
  row['Payment_History_JSON'] = _buildPaymentHistoryJSON_(logSheet, String(row['Invoice_No']||'').trim());
  try {
    var url = _generateAndSaveInvoicePdf_(row);
    Logger.log(url ? '✅ PDF created at: ' + url : '❌ FAILED: empty string');
  } catch(e) { Logger.log('❌ EXCEPTION: ' + e.message); }
}
