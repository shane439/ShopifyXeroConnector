/** ======================================================================
 * Diag Trace (prepend-at-top logger) + timing/context + raw-body vault
 * Sheet: Diag_Trace (auto-created); newest logs always at the top (row 2)
 * Optional raw store: Diag_Raw (top-inserted, capped size)
 * ====================================================================== */

var __DX_CTX__ = null;
if (typeof DX_ENABLE_RAW_VAULT === 'undefined') var DX_ENABLE_RAW_VAULT = true;   // set false to disable raw store
if (typeof DX_RAW_VAULT_MAX === 'undefined')    var DX_RAW_VAULT_MAX = 400;       // keep at most this many rows

function DX_setCtx_(obj) { __DX_CTX__ = obj || {}; }
function DX_ctx_() { return __DX_CTX__ || {}; }

function DX_uuid_() {
  var s = Utilities.getUuid();
  return s && s.replace(/-/g,'').slice(0,16);
}

function DX_startTimer_(label) {
  var ctx = DX_ctx_();
  ctx.timers = ctx.timers || {};
  ctx.timers[label] = new Date().getTime();
}

function DX_endTimer_(label) {
  var ctx = DX_ctx_();
  var t0 = ctx.timers && ctx.timers[label];
  if (!t0) return null;
  var ms = new Date().getTime() - t0;
  ctx.timers[label] = null;
  return ms;
}

function DX_log_(phase, orderRef, step, status, note, extraObj) {
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName('Diag_Trace');
    if (!sh) {
      sh = ss.insertSheet('Diag_Trace');
      sh.appendRow([
        'Timestamp','TraceRunId','OrderIdx','Phase','OrderRef','Step',
        'Status','Note','DurationMS','ExtraJSON'
      ]);
    }
    // upgrade-safe headers
    var header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    if (header.length < 10) {
      sh.clear();
      sh.appendRow(['Timestamp','TraceRunId','OrderIdx','Phase','OrderRef','Step','Status','Note','DurationMS','ExtraJSON']);
    }
    var ctx = DX_ctx_();
    var ms = (extraObj && typeof extraObj.__stopTimerLabel === 'string')
      ? DX_endTimer_(extraObj.__stopTimerLabel) : null;

    sh.insertRowsAfter(1, 1); // prepend at top
    sh.getRange(2, 1, 1, 10).setValues([[
      new Date(),
      String(ctx.TraceRunId || ''),
      String(ctx.OrderIdx != null ? ctx.OrderIdx : ''),
      String(phase || ''),
      String(orderRef || ''),
      String(step || ''),
      String(status || ''),
      String(note || ''),
      ms != null ? ms : '',
      extraObj ? JSON.stringify(extraObj) : ''
    ]]);
  } catch(e) {
    Logger.log('DX_log_ error: ' + e.message);
  }
}

/** Compact payload for logging (safe, size-limited). */
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

/** Human ref from payload for logs */
function DX_refFromPayload_(payload) {
  try {
    var inv = payload && payload.Invoices && payload.Invoices[0];
    return (inv && inv.Reference) || '';
  } catch(e) { return ''; }
}

/** Optional raw response vault (prepend, capped) */
function DX_raw_(label, text) {
  try {
    if (!DX_ENABLE_RAW_VAULT) return;
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName('Diag_Raw');
    if (!sh) {
      sh = ss.insertSheet('Diag_Raw');
      sh.appendRow(['Timestamp','TraceRunId','OrderRef','Label','Raw']);
    }
    var ctx = DX_ctx_();
    sh.insertRowsAfter(1, 1);
    sh.getRange(2,1,1,5).setValues([[new Date(), String(ctx.TraceRunId||''), String(ctx.orderRef||ctx.OrderIdx||''), String(label||''), String(text||'')]]);
    // cap rows
    var last = sh.getLastRow();
    if (last > DX_RAW_VAULT_MAX+1) sh.deleteRows(DX_RAW_VAULT_MAX+2, last-(DX_RAW_VAULT_MAX+1));
  } catch(e){ Logger.log('DX_raw_ error: '+e.message); }
}
