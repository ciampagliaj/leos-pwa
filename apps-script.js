// ── LEO'S RECEIPTS — GOOGLE APPS SCRIPT ──
// Paste this entire file into your Apps Script editor, then
// Deploy → Manage deployments → edit existing → New version → Update
//
// Handles three actions (passed via ?action= URL param):
//   drive   — download image from Cloudinary and save to Drive folder
//   expense — write a row to the Expenses sheet
//   income  — write a row to the Income sheet

var SPREADSHEET_ID = '1ExM-7GB74gDchC242iKJdGAVPEBNVcYB';
var DRIVE_FOLDER   = "Leo's Receipts";

function doGet(e) {
  try {
    var action = (e.parameter.action || 'drive').toLowerCase();

    if (action === 'expense') {
      writeExpenseRow(e.parameter);
      if (e.parameter.url) saveToDrive(e.parameter.url, e.parameter.category || 'Other', e.parameter.vendor || '', e.parameter.date || '');
      return ok();
    }

    if (action === 'income') {
      writeIncomeRow(e.parameter);
      return ok();
    }

    // action === 'drive' — legacy Drive-only call
    if (e.parameter.url) {
      saveToDrive(e.parameter.url, e.parameter.category || 'Other', e.parameter.vendor || '', e.parameter.date || '');
    }
    return ok();

  } catch (err) {
    return fail(err.message);
  }
}

// ── EXPENSES SHEET ──
// Columns: Date | Vendor/Supplier | Category | Description | Amount pre-HST | HST Paid | Total Paid | Payment Method | Receipt # | Notes
function writeExpenseRow(p) {
  var ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet  = ss.getSheetByName('Expenses');
  if (!sheet) throw new Error('Sheet "Expenses" not found');

  var amount = parseFloat(p.amount) || 0;
  var hst    = Math.round(amount * 0.13 * 100) / 100;
  var total  = Math.round((amount + hst)  * 100) / 100;

  sheet.appendRow([
    p.date          || '',   // Date
    p.vendor        || '',   // Vendor/Supplier
    p.category      || '',   // Category
    p.notes         || '',   // Description
    amount.toFixed(2),       // Amount pre-HST ($)
    hst.toFixed(2),          // HST Paid ($)
    total.toFixed(2),        // Total Paid ($)
    p.paymentMethod || '',   // Payment Method
    p.receiptNum    || '',   // Receipt #
    ''                       // Notes (extra — description above covers it)
  ]);
}

// ── INCOME SHEET ──
// Columns: INV # | Date | Client Name | Job Type | Description | Labour | Materials | Subtotal | HST 13% | Total Invoiced | Deposit | Balance | Status | Est. Ref.
function writeIncomeRow(p) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Income');
  if (!sheet) throw new Error('Sheet "Income" not found');

  var labour    = parseFloat(p.labour)    || 0;
  var materials = parseFloat(p.materials) || 0;
  var deposit   = parseFloat(p.deposit)   || 0;
  var subtotal  = labour + materials;
  var hst       = Math.round(subtotal * 0.13  * 100) / 100;
  var total     = Math.round((subtotal + hst)  * 100) / 100;
  var balance   = Math.round((total - deposit) * 100) / 100;

  sheet.appendRow([
    p.invoiceNum  || '',   // INV #
    p.date        || '',   // Date
    p.clientName  || '',   // Client Name
    p.jobType     || '',   // Job Type
    p.notes       || '',   // Description
    labour.toFixed(2),     // Labour ($)
    materials.toFixed(2),  // Materials ($)
    subtotal.toFixed(2),   // Subtotal ($)
    hst.toFixed(2),        // HST Charged (13%)
    total.toFixed(2),      // Total Invoiced ($)
    deposit.toFixed(2),    // Deposit Received ($)
    balance.toFixed(2),    // Balance Owing ($)
    'Sent',                // Status
    ''                     // Est. Ref.
  ]);
}

// ── DRIVE UPLOAD ──
function saveToDrive(url, category, vendor, date) {
  var folders = DriveApp.getFoldersByName(DRIVE_FOLDER);
  var parent  = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER);

  var subs = parent.getFoldersByName(category);
  var sub  = subs.hasNext() ? subs.next() : parent.createFolder(category);

  var blob     = UrlFetchApp.fetch(url).getBlob();
  var safeName = (vendor || 'receipt').replace(/[^a-zA-Z0-9 \-]/g, '').trim();
  blob.setName(safeName + '_' + (date || '') + '.jpg');
  sub.createFile(blob);
}

// ── HELPERS ──
function ok()       { return json({ ok: true }); }
function fail(msg)  { return json({ ok: false, error: msg }); }
function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
