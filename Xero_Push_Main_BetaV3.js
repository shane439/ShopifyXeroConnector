/** ======================================================================
 * Xero Push Main — BetaV3 with Diag_Trace instrumentation
 * Uses your Xeroitemslist resolver (already in project).
 * Entry: XL_pushSelectedDraft_InclusiveOnly_Standalone_BetaV3
 * ====================================================================== */

function XL_pushSelectedDraft_InclusiveOnly_Standalone_BetaV3() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('OrdersInbox');
  if (!sh) { ui.alert('OrdersInbox not found'); return; }

  var sel = sh.getActiveRange();
  if (!sel) { ui.alert('Select one or more rows in OrdersInbox.'); return; }

  var col = _oi_columns_(sh);
  if (col.order_id == null || col.order_number == null || col.raw_json == null) {
    ui.alert('OrdersInbox missing required headers (order_id, order_number, raw_json).'); return;
  }

  if (typeof _xi_load_ === 'function') _xi_load_(); // warm resolver

  var total=0, ok=0, fail=0, downgraded=0;
  var results = [];

  var rows = sel.getValues();
  for (var i=0;i<rows.length;i++){
    var row1 = sel.getRow() + i;
    try{
      var r         = rows[i];
      var orderId   = _get(r, col.order_id);
      var orderNo   = _get(r, col.order_number);
      var createdAt = _get(r, col.created_at);
      var raw       = _get(r, col.raw_json);
      if (!orderId) continue;

      total++;
      var ref = 'Shopify #' + orderNo;
      DX_log_('MAIN', ref, 'row:start', 'info', 'Start row build');

      var contact = _xp_buildContactFromRow_(r, col);
      var dateStr = _asYMD_(createdAt);

      var shop  = _safeParseJSON_(raw, {});
      var items = Array.isArray(shop.line_items) ? shop.line_items : [];

      DX_log_('MAIN', ref, 'row:lines', 'info', 'items='+items.length);

      // Build lineItems using your Xeroitemslist resolver
      var lineItems = items.map(function(li){
        var qty  = Number(li && li.quantity || 0);
        var unit = (li && li.price != null) ? Number(li.price) : (li && li.total && qty ? Number(li.total)/qty : 0);
        var m    = (typeof _xi_resolveBySKU_ === 'function') ? _xi_resolveBySKU_(li) : {};
        var ln = { Description: String(li && li.title || ''), Quantity: qty, UnitAmount: unit };
        if (m.ItemCode) ln.ItemCode = m.ItemCode; else if (li && li.sku != null) ln.ItemCode = String(li.sku);
        if (m.AccountCode) ln.AccountCode = m.AccountCode;
        if (m.TaxType)     ln.TaxType     = m.TaxType;
        return ln;
      });

      var payload = {
        Invoices: [{
          Type: 'ACCREC',
          Status: 'DRAFT',
          LineAmountTypes: 'Inclusive',
          Reference: ref, Contact: contact, Date: dateStr,
          LineItems: lineItems
        }]
      };

      DX_log_('MAIN', ref, 'post:call', 'info', 'POST wrapper', {snap: DX_compactInvoice_(payload)});

      var res = XP_postInvoiceDraftInclusive_(payload);

      if (res.code===200 || res.code===201){
        ok++;
        DX_log_('MAIN', ref, 'post:ok', '200/201', 'Draft created');
        _markPushed_(sh, row1, col, res);

        var dIdx = (res.meta && res.meta.downgradedIdxs) ? res.meta.downgradedIdxs : [];
        if (dIdx.length){
          downgraded += dIdx.length;
          dIdx.forEach(function(ix){
            var li = items[ix] || {};
            results.push([new Date(), String(orderNo), ix, String(li.sku||''), String(li.title||''), 'DOWNGRADED', 'Targeted/progressive']);
          });
        } else {
          results.push([new Date(), String(orderNo), '', '', '', 'OK', '']);
        }
      } else {
        fail++;
        DX_log_('MAIN', ref, 'post:fail', String(res.code), (res.body||'').slice(0,400));
        results.push([new Date(), String(orderNo), '', '', '', 'FAILED', 'HTTP '+res.code]);
      }

    } catch(e){
      fail++;
      DX_log_('MAIN', 'row#'+row1, 'exception', 'error', e.message);
      results.push([new Date(), '', '', '', '', 'FAILED', e.message]);
      ui.alert('Row '+row1+' → '+e.message);
    }
  }

  _writePushResults_(results);

  var msg = [
    'Orders processed: ' + total,
    'Successful: ' + ok,
    'Failed: ' + fail,
    'Downgraded lines (targeted/progressive): ' + downgraded
  ].join('\n');
  ss.toast('Push summary — see Push_Results & Diag_Trace', 'Xero Push', 8);
  ui.alert('Xero Push Summary', msg, ui.ButtonSet.OK);
}

/* === helpers (unchanged) === */
function _oi_columns_(sh){
  var headers=(sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]||[]).map(String);
  function idxOf(name){ var n=name.toLowerCase(); for(var i=0;i<headers.length;i++) if(headers[i].toLowerCase().trim()===n) return i; return null; }
  return { order_id: idxOf('order_id'), order_number: idxOf('order_number'),
           created_at: idxOf('created_at'), customer_name: idxOf('customer_name'),
           customer_email: idxOf('customer_email'), raw_json: idxOf('raw_json'),
           pushed_at: idxOf('PushedAt'), pushed_invoice: idxOf('PushedInvoiceID') };
}
function _get(row, ix){ return (ix==null? null : row[ix]); }
function _asYMD_(d){ try{ if(d&&d.toISOString) return d.toISOString().slice(0,10); return Utilities.formatDate(new Date(d),'UTC','yyyy-MM-dd'); }catch(e){ return Utilities.formatDate(new Date(),'UTC','yyyy-MM-dd'); } }
function _safeParseJSON_(s, fb){ try{ return s? JSON.parse(s): fb; }catch(e){ return fb; } }
function _xp_buildContactFromRow_(row, col){ var out={}; var name=_get(row,col.customer_name)||''; var email=_get(row,col.customer_email)||''; if(name) out.Name=name; if(email) out.EmailAddress=email; return out; }
function _markPushed_(sh, row1, col, res){ try{ if (col.pushed_at!=null) sh.getRange(row1, col.pushed_at+1).setValue(new Date()); if (col.pushed_invoice!=null){ var id=(res.json&&res.json.Invoices&&res.json.Invoices[0]&&res.json.Invoices[0].InvoiceID)||''; sh.getRange(row1, col.pushed_invoice+1).setValue(id);} }catch(e){ Logger.log('Stamping error: '+e.message);} }
function _writePushResults_(rows){ if(!rows||!rows.length) return; var ss=SpreadsheetApp.getActive(); var sh=ss.getSheetByName('Push_Results'); if(!sh){ sh=ss.insertSheet('Push_Results'); sh.appendRow(['Timestamp','OrderNumber','LineIdx','SKU','Title','Action','Note']); } sh.getRange(sh.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows); }
