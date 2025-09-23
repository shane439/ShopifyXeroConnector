/** ======================================================================
 * Xero Push Main — BetaV3
 * Deltas (this build):
 *  - Skip cancelled Shopify orders (cancelled_at / cancel_reason / financial_status)
 *  - Summary shows skipped count + ⏭ entries.
 *  - For downgraded lines, include truncated Description and a short reason code.
 * ====================================================================== */

if (typeof DX_SUPPRESS_ROW_ALERTS === 'undefined') var DX_SUPPRESS_ROW_ALERTS = true;

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

  var trace = DX_uuid_();
  DX_setCtx_({ TraceRunId: trace, OrderIdx: '' });
  DX_log_('MAIN', '', 'batch:start', 'info', 'Begin batch', { rows: sel.getNumRows(), strategy: (typeof XP_STRATEGY_ORDER!=='undefined'?XP_STRATEGY_ORDER:'A') });

  if (typeof _xi_load_ === 'function') _xi_load_(); // warm resolver

  var total=0, ok=0, fail=0, downgraded=0, skipped=0;
  var results = [];
  var summary = [];
  var postdatedRefs = [];

  var rows = sel.getValues();
  for (var i=0;i<rows.length;i++){
    var row1 = sel.getRow() + i;
    DX_setCtx_({ TraceRunId: trace, OrderIdx: row1 });

    try{
      var r         = rows[i];
      var orderId   = _get(r, col.order_id);
      var orderNo   = _get(r, col.order_number);
      var createdAt = _get(r, col.created_at);
      var fstatus   = _get(r, col.financial_status);
      var raw       = _get(r, col.raw_json);
      if (!orderId) continue;

      total++;
      var ref = 'Shopify #' + orderNo;

      var shop  = _safeParseJSON_(raw, {});
      if (_isCancelled_(shop, fstatus)) {
        skipped++;
        DX_log_('MAIN', ref, 'skip:cancelled', 'skip', 'Shopify order cancelled', { cancelled_at: shop.cancelled_at || '', cancel_reason: shop.cancel_reason || '', financial_status: fstatus || '' });
        results.push([new Date(), String(orderNo), '', '', '', 'SKIPPED', 'cancelled']);
        summary.push('⏭ '+orderNo+' (cancelled)');
        continue;
      }

      DX_log_('MAIN', ref, 'row:start', 'info', 'Start order', { row: row1 });
      DX_startTimer_('order-post');

      var contact = _xp_buildContactFromRow_(r, col);
      var orderDate = _safeDate_(createdAt);

      var items = Array.isArray(shop.line_items) ? shop.line_items : [];
      var orderNote = (shop && shop.note) ? String(shop.note) : '';

      var dueDate = _determineDueDate_(orderDate, orderNote);
      if (dueDate.isPostdated) postdatedRefs.push(String(orderNo));

      var lineItems = items.map(function(li){
        var qty = Number(li && li.quantity || 0);
        var unitPrice = (li && li.price != null) ? Number(li.price) : 0;

        var totalDisc = 0;
        if (li && li.total_discount != null) totalDisc = Number(li.total_discount) || 0;
        else if (Array.isArray(li && li.discount_allocations)) {
          li.discount_allocations.forEach(function(da){ var a = Number(da && da.amount || 0); if (!isNaN(a)) totalDisc += a; });
        }

        var unitAfterDisc = (qty > 0) ? Math.max(0, unitPrice - (totalDisc/qty)) : unitPrice;

        var m = (typeof _xi_resolveBySKU_ === 'function') ? _xi_resolveBySKU_(li) : {};
        var ln = { Description: String(li && li.title || ''), Quantity: qty, UnitAmount: unitAfterDisc };
        if (m.ItemCode){ ln.ItemCode = m.ItemCode; } else if (li && li.sku != null) { ln.ItemCode = String(li.sku); }
        if (m.AccountCode){ ln.AccountCode = m.AccountCode; }
        if (m.TaxType){ ln.TaxType = m.TaxType; }

        if (totalDisc > 0) ln.Description += "\n[Discount Applied: -" + totalDisc.toFixed(2) + "]";
        return ln;
      });

      if (orderNote) {
        lineItems.push({ Description: "[Note: " + orderNote + "]", Quantity: 0, UnitAmount: 0 });
      }

      var payload = {
        Invoices: [{
          Type: 'ACCREC',
          Status: 'DRAFT',
          LineAmountTypes: 'Inclusive',
          Reference: ref,
          Contact: contact,
          Date: _asYMD_(orderDate),
          DueDate: _asYMD_(dueDate.date),
          LineItems: lineItems
        }]
      };

      var res = XP_postInvoiceDraftInclusive_(payload);

      if (res.code===200 || res.code===201){
        ok++;
        var inv = (res.json && res.json.Invoices && res.json.Invoices[0]) || {};
        var invId = inv.InvoiceID || '';
        var invNo = inv.InvoiceNumber || '';
        DX_log_('MAIN', ref, 'post:ok', '200/201', 'Draft created', { invoiceId: invId, invoiceNo: invNo, headers: res.headers||{}, meta: res.meta||{}, __stopTimerLabel:'order-post' });

        _markPushed_(sh, row1, col, res);

        var dIdx = (res.meta && res.meta.downgradedIdxs) ? res.meta.downgradedIdxs : [];
        if (dIdx.length){
          downgraded += dIdx.length;
          dIdx.forEach(function(ix){
            var li = items[ix] || {};
            var sku = String(li.sku||'');
            var desc = _trunc_(String((lineItems[ix] && lineItems[ix].Description) || li.title || ''), 26);
            var reason = (res.meta && res.meta.downgradeReason) || '';
            var code = _abbrReason_(reason);
            results.push([new Date(), String(orderNo), ix, sku, (li.title||''), 'DOWNGRADED', reason]);
            summary.push('• ↓ '+orderNo+'  ln '+ix+'  SKU='+sku+'  "'+desc+'"  (R:'+code+')');
          });
        }
        results.push([new Date(), String(orderNo), '', '', '', 'OK', 'InvoiceID='+invId+' / No='+invNo]);
        summary.push('✓ '+orderNo+'  →  '+(invNo||invId||'DRAFT'));

      } else {
        fail++;

        var parsed = (res.code===400 && res.meta && res.meta.parsed400)
          ? res.meta.parsed400
          : (res.code===400 ? _xp_parse400_(res.body || '') : {indexes:[],messages:[],summary:''});

        var reasons = _xp_classifyReasons_(parsed.messages || []);

        var idxsToReport = parsed.indexes.length ? parsed.indexes : (items.map(function(_,i){return i;}));
        idxsToReport.forEach(function(ix){
          var li = items[ix] || {};
          var sku = String(li.sku || '');
          var title = String(li.title || '');
          var desc = _trunc_(title, 26);
          var reason = reasons[ix] || reasons._generic || (parsed.messages.slice(0,2).join(' | ') || ('HTTP '+res.code));
          results.push([new Date(), String(orderNo), ix, sku, title, 'FAILED', reason]);
          summary.push('✗ '+orderNo+'  ln '+ix+'  SKU='+sku+'  "'+desc+'"  ('+reason+')');
        });

        summary.push('✗ '+orderNo+'  →  '+(parsed.summary || ('HTTP '+res.code))+(parsed.indexes.length?('  idx='+parsed.indexes.join(',')):''));
      }

    } catch(e){
      fail++;
      DX_log_('MAIN', 'row#'+row1, 'exception', 'error', e.message, { stack: (e && e.stack) ? String(e.stack).slice(0,600) : '' , __stopTimerLabel:'order-post' });
      if (!DX_SUPPRESS_ROW_ALERTS) SpreadsheetApp.getUi().alert('Row '+row1+' → '+e.message);
      results.push([new Date(), '', '', '', '', 'FAILED', e.message]);
      summary.push('✗ row#'+row1+' → '+e.message);
    }
  }

  _writePushResults_(results);

  var msg = [
    'Orders processed: ' + total,
    'Successful: ' + ok,
    'Failed: ' + fail,
    'Skipped: ' + skipped + ' (cancelled)',
    'Downgraded lines (targeted/progressive): ' + downgraded
  ].join('\n');

  var extras = '';
  if (postdatedRefs.length) extras = '\n\nPostdated/Credit terms detected for orders: ' + postdatedRefs.join(', ');

  DX_log_('MAIN', '', 'batch:end', 'info', 'End batch', { total: total, ok: ok, fail: fail, skipped: skipped, downgraded: downgraded, pd: postdatedRefs });

  var pretty = [
    'Xero Push Summary ',
    '------------------',
    msg, '', 'Details:',
    summary.length ? summary.join('\n') : '(no additional details)'
  ].join('\n') + extras;
  _showCopyDialog_(pretty);
}

/* === helpers ========================================================== */
function _oi_columns_(sh){
  var headers=(sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]||[]).map(String);
  function idxOf(name){ var n=name.toLowerCase(); for(var i=0;i<headers.length;i++) if(headers[i].toLowerCase().trim()===n) return i; return null; }
  return { order_id: idxOf('order_id'), order_number: idxOf('order_number'),
           created_at: idxOf('created_at'), financial_status: idxOf('financial_status'),
           customer_name: idxOf('customer_name'), customer_email: idxOf('customer_email'),
           raw_json: idxOf('raw_json'), pushed_at: idxOf('PushedAt'), pushed_invoice: idxOf('PushedInvoiceID') };
}
function _get(row, ix){ return (ix==null? null : row[ix]); }
function _safeDate_(d){ try{ if(d && d.toISOString) return d; return new Date(d); }catch(e){ return new Date(); } }
function _asYMD_(d){ try{ if(d&&d.toISOString) return d.toISOString().slice(0,10); return Utilities.formatDate(new Date(d),'UTC','yyyy-MM-dd'); }catch(e){ return Utilities.formatDate(new Date(),'UTC','yyyy-MM-dd'); } }
function _safeParseJSON_(s, fb){ try{ return s? JSON.parse(s): fb; }catch(e){ return fb; } }
function _xp_buildContactFromRow_(row, col){ var out={}; var name=_get(row,col.customer_name)||''; var email=_get(row,col.customer_email)||''; if(name) out.Name=name; if(email) out.EmailAddress=email; return out; }
function _markPushed_(sh, row1, col, res){ try{ if (col.pushed_at!=null) sh.getRange(row1, col.pushed_at+1).setValue(new Date()); if (col.pushed_invoice!=null){ var id=(res.json&&res.json.Invoices&&res.json.Invoices[0]&&res.json.Invoices[0].InvoiceID)||''; sh.getRange(row1, col.pushed_invoice+1).setValue(id);} }catch(e){ Logger.log('Stamping error: '+e.message);} }
function _writePushResults_(rows){ if(!rows||!rows.length) return; var ss=SpreadsheetApp.getActive(); var sh=ss.getSheetByName('Push_Results'); if(!sh){ sh=ss.insertSheet('Push_Results'); sh.appendRow(['Timestamp','OrderNumber','LineIdx','SKU','Title','Action','Note']); } sh.getRange(sh.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows); }
function _showCopyDialog_(text) {
  function esc(s){ return String(s).replace(/[&<>]/g, function(c){ return ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]); }); }
  var html = HtmlService.createHtmlOutput(
    '<div style="font:13px/1.4 -apple-system,Segoe UI,Arial,Roboto,system-ui;padding:12px 12px 4px 12px;">'+
      '<div style="margin-bottom:8px;color:#111;font-weight:600">Copy & share</div>'+
      '<textarea id="t" style="width:100%;height:360px;white-space:pre; font-family:ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;">'+esc(text)+'</textarea>'+
      '<div style="margin-top:8px;text-align:right">'+
        '<button onclick="var ta=document.getElementById(\'t\'); ta.focus(); ta.select(); try{navigator.clipboard.writeText(ta.value);}catch(e){}">Copy</button>'+
        '<button onclick="google.script.host.close()" style="margin-left:8px">Close</button>'+
      '</div>'+
    '</div>'
  ).setWidth(620).setHeight(480);
  SpreadsheetApp.getUi().showModalDialog(html, 'Xero Push — Details');
}
function _xp_classifyReasons_(messages){
  var out = { _generic: '' };
  if (!messages || !messages.length) return out;
  var lower = messages.join(' || ').toLowerCase();
  function has(s){ return lower.indexOf(s) >= 0; }

  if (has('account code'))   out._generic = 'Account code invalid/missing';
  if (has('tax type'))       out._generic = (out._generic? out._generic+'; ' : '') + 'TaxType invalid/missing';
  if (has('item code') || has('item not found')) out._generic = (out._generic? out._generic+'; ' : '') + 'ItemCode unknown';

  (messages || []).forEach(function(msg){
    var m = /LineItems\[(\d+)\].*?(Account|Tax|Item)/i.exec(String(msg||''));
    if (m) {
      var idx = parseInt(m[1], 10);
      var kind = (m[2] || '').toLowerCase();
      var hint = '';
      if (kind.indexOf('account')>=0) hint = 'Account code invalid/missing';
      else if (kind.indexOf('tax')>=0) hint = 'TaxType invalid/missing';
      else if (kind.indexOf('item')>=0) hint = 'ItemCode unknown';
      out[idx] = hint || 'Line validation';
    }
  });
  if (!out._generic) out._generic = messages.slice(0,2).join(' | ');
  return out;
}
function _determineDueDate_(orderDate, note) {
  var isPD = /\b(cr|credit|pd|post\s*dated|postdated)\b/i.test(String(note||''));
  var d = new Date(orderDate); if (isPD) d.setDate(d.getDate()+14);
  return { date: d, isPostdated: isPD };
}
function _isCancelled_(shop, finStatus) {
  if (!shop && !finStatus) return false;
  var fs = String(finStatus||'').toLowerCase();
  if (fs.indexOf('cancel') >= 0) return true;
  var cAt = shop && shop.cancelled_at;
  var cReason = shop && shop.cancel_reason;
  return !!(cAt || cReason);
}
function _trunc_(s, n){ s=String(s||'').replace(/\s+/g,' ').trim(); return (s.length>n)? (s.slice(0,n-1)+'…') : s; }
function _abbrReason_(r){
  r = String(r||'').toLowerCase();
  if (r.indexOf('probe.dropitemcode')>=0 || r.indexOf('passa.dropitemcode')>=0) return 'SKU';
  if (r.indexOf('droptax')>=0) return 'TAX';
  if (r.indexOf('forcefallbackac')>=0) return 'AC';
  if (r.indexOf('incremental')>=0) return 'MIN';
  if (r.indexOf('probe')>=0) return 'PRB';
  return 'MIN';
}
