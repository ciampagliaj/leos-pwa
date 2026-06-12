// ── LEO'S RECEIPTS — GOOGLE APPS SCRIPT ──
// Deploy as Web App → Execute as: Me → Who has access: Anyone

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
    if (e.parameter.url) saveToDrive(e.parameter.url, e.parameter.category || 'Other', e.parameter.vendor || '', e.parameter.date || '');
    return ok();
  } catch (err) {
    return fail(err.message);
  }
}

function getSheet(p, sheetName) {
  var id = p.spreadsheetId || '1ExM-7GB74gDchC242iKJdGAVPEBNVcYB';
  var ss = SpreadsheetApp.openById(id);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Tab "' + sheetName + '" not found in spreadsheet');
  return sheet;
}

function writeExpenseRow(p) {
  var sheet  = getSheet(p, 'Expenses');
  var amount = parseFloat(p.amount) || 0;
  var hst    = Math.round(amount * 0.13 * 100) / 100;
  var total  = Math.round((amount + hst) * 100) / 100;
  sheet.appendRow([
    p.date          || '',
    p.vendor        || '',
    p.category      || '',
    p.notes         || '',
    amount.toFixed(2),
    hst.toFixed(2),
    total.toFixed(2),
    p.paymentMethod || '',
    p.receiptNum    || '',
    ''
  ]);
}

function writeIncomeRow(p) {
  var sheet     = getSheet(p, 'Income');
  var labour    = parseFloat(p.labour)    || 0;
  var materials = parseFloat(p.materials) || 0;
  var deposit   = parseFloat(p.deposit)   || 0;
  var subtotal  = labour + materials;
  var hst       = Math.round(subtotal * 0.13 * 100) / 100;
  var total     = Math.round((subtotal + hst) * 100) / 100;
  var balance   = Math.round((total - deposit) * 100) / 100;
  sheet.appendRow([
    p.invoiceNum  || '',
    p.date        || '',
    p.clientName  || '',
    p.jobType     || '',
    p.notes       || '',
    labour.toFixed(2),
    materials.toFixed(2),
    subtotal.toFixed(2),
    hst.toFixed(2),
    total.toFixed(2),
    deposit.toFixed(2),
    balance.toFixed(2),
    'Sent',
    ''
  ]);
}

function saveToDrive(url, category, vendor, date) {
  var folders = DriveApp.getFoldersByName("Leo's Receipts");
  var parent  = folders.hasNext() ? folders.next() : DriveApp.createFolder("Leo's Receipts");
  var subs    = parent.getFoldersByName(category);
  var sub     = subs.hasNext() ? subs.next() : parent.createFolder(category);
  var blob    = UrlFetchApp.fetch(url).getBlob();
  var name    = (vendor || 'receipt').replace(/[^a-zA-Z0-9 \-]/g, '').trim();
  blob.setName(name + '_' + (date || '') + '.jpg');
  sub.createFile(blob);
}

function ok()      { return json({ ok: true }); }
function fail(msg) { return json({ ok: false, error: msg }); }
function json(d)   {
  return ContentService.createTextOutput(JSON.stringify(d))
    .setMimeType(ContentService.MimeType.JSON);
}

function testAccess() {
  var ss = SpreadsheetApp.openById('1ExM-7GB74gDchC242iKJdGAVPEBNVcYB');
  Logger.log('Found: ' + ss.getName());
  var sheet = ss.getSheetByName('Expenses');
  Logger.log('Expenses tab: ' + (sheet ? 'YES' : 'NOT FOUND'));
  var sheet2 = ss.getSheetByName('Income');
  Logger.log('Income tab: ' + (sheet2 ? 'YES' : 'NOT FOUND'));
}
