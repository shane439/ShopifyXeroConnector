/** ======================================================================
 * Diag Trace (prepend-at-top logger)
 * Sheet: Diag_Trace (auto-created)
 * Every call inserts at row 2 so newest logs are always on top.
 * ====================================================================== */

function DX_log_(phase, orderRef, step, status, note, extraObj) {
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName('Diag_Trace');
    if (!sh) {
      sh = ss.insertSheet('Diag_Trace');
      sh.appendRow(['Timestamp','Phase','OrderRef','Step','Status','Note','ExtraJSON']);
    }
    // Insert at top (row 2)
    sh.insertRowsAfter(1, 1);
    var row = [
      new Date(),
      String(phase || ''),
      String(orderRef || ''),
      String(step || ''),
      String(status || ''),
      String(note || ''),
      extraObj ? JSON.stringify(extraObj) : ''
    ];
    sh.getRange(2, 1, 1, row.length).setValues([row]);
  } catch(e) {
    Logger.log('DX_log_ error: ' + e.message);
  }
}

/** Compact a payload for logging (safe, size-limited). */
function DX_compactInvoice_(payload) {
  try {
    var inv = payload && payload.Invoices && payload.Invoices[0];
    if (!inv) return {};
    return {
      Reference: inv.Reference || '',
      Date: inv.Date || '',
      LineAmountTypes: inv.LineAmountTypes || '',
      Lines: (inv.LineItems || []).map(function(li, i){
        return {
          i: i,
          Desc: (li.Description || '').slice(0,80),
          Qty:  li.Quantity,
          Price: li.UnitAmount,
          AC:   li.AccountCode || '',
          IT:   li.ItemCode || '',
          TX:   li.TaxType || ''
        };
      })
    };
  } catch(e){ return {}; }
}

/** Extract a human ref from payload for logs */
function DX_refFromPayload_(payload) {
  try {
    var inv = payload && payload.Invoices && payload.Invoices[0];
    return (inv && inv.Reference) || '';
  } catch(e) { return ''; }
}
