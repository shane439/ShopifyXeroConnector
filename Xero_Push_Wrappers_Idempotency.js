/* ============================================================================
 * Optional wrappers (validation detect, strip ItemCode, idempotency, etc.)
 * ==========================================================================*/

function _xp_isValidation400_(resp) {
  try {
    if (resp.getResponseCode() !== 400) return false;
    var j = JSON.parse(resp.getContentText());
    return String(j && j.Type).toLowerCase().indexOf('validation') >= 0;
  } catch (e) { return false; }
}
function _xp_stripItemCodeOnly_(lines) {
  var clone = [];
  for (var i = 0; i < lines.length; i++) {
    var l = JSON.parse(JSON.stringify(lines[i] || {}));
    if ('ItemCode' in l) delete l.ItemCode;
    clone.push(l);
  }
  return clone;
}
function _xp_findExistingDraftByReference_(accessToken, tenantId, reference) {
  try {
    var where = 'Reference=="' + reference.replace(/"/g,'\\"') + '" AND Status=="DRAFT"';
    var url = 'https://api.xero.com/api.xro/2.0/Invoices?where=' + encodeURIComponent(where) + '&page=1';
    var resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + accessToken, 'xero-tenant-id': tenantId, Accept: 'application/json' },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() >= 300) return false;
    var j = JSON.parse(resp.getContentText());
    var arr = (j && j.Invoices) || [];
    return arr.length > 0;
  } catch (e) { return false; }
}
function _xp_resolveMarkerCols_(headers) {
  return {
    pushedAtIdx:  headers.indexOf(XP_LOCAL_MARKER_COLS.pushedAt)  + 1,
    invoiceIdIdx: headers.indexOf(XP_LOCAL_MARKER_COLS.invoiceId) + 1
  };
}
function _xp_rowAlreadyPushed_(row, mk) {
  try {
    if (!mk.pushedAtIdx || !mk.invoiceIdIdx) return false;
    var pushedAt  = row[mk.pushedAtIdx  - 1];
    var invoiceId = row[mk.invoiceIdIdx - 1];
    return Boolean(pushedAt && invoiceId);
  } catch (e) { return false; }
}
function _xp_markRowPushed_(sheet, absRowIndex, mk, invoiceId) {
  try {
    if (!mk.pushedAtIdx || !mk.invoiceIdIdx) return;
    var now = new Date();
    sheet.getRange(absRowIndex, mk.pushedAtIdx).setValue(now);
    sheet.getRange(absRowIndex, mk.invoiceIdIdx).setValue(invoiceId);
  } catch (e) { /* best effort */ }
}
