// ═══════════════════════════════════════════════════════
//  iSPICE ORDERS — Google Apps Script
//  Paste this entire file into Google Apps Script,
//  then Deploy → New deployment → Web app → Anyone
// ═══════════════════════════════════════════════════════

function doGet(e) {
  const action = e.parameter.action || 'list';

  try {
    if (action === 'add') {
      const order = JSON.parse(decodeURIComponent(e.parameter.order || '{}'));
      addOrder(order);
      return respond({ ok: true });
    }

    if (action === 'list') {
      return respond(listOrders());
    }

    if (action === 'clear') {
      clearOrders();
      return respond({ ok: true, cleared: true });
    }

    return respond({ ok: true, message: 'iSpice Orders API ready' });

  } catch (err) {
    return respond({ ok: false, error: err.toString() });
  }
}

// ── Helpers ──────────────────────────────────────────

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Orders');
  if (!sheet) {
    sheet = ss.insertSheet('Orders');
    sheet.appendRow(['id','name','item','qty','comment','catKey','catLabel','catEmoji','catCls','time']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function addOrder(order) {
  const sheet = getSheet();
  sheet.appendRow([
    order.id       || Date.now(),
    order.name     || '',
    order.item     || '',
    order.qty      || 1,
    order.comment  || '',
    order.catKey   || '',
    order.catLabel || '',
    order.catEmoji || '',
    order.catCls   || '',
    order.time     || new Date().toLocaleTimeString()
  ]);
}

function listOrders() {
  const sheet = getSheet();
  if (sheet.getLastRow() <= 1) return [];
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  return data.slice(1).map(row => {
    const o = {};
    headers.forEach((h, i) => { o[h] = row[i]; });
    o.qty = Number(o.qty) || 1;
    return o;
  });
}

function clearOrders() {
  const sheet = getSheet();
  if (sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
  }
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
