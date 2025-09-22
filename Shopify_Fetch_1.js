/* ============================================================================
 * Shopify Fetch → OrdersInbox
 * ==========================================================================*/

function SF_promptFetchByDate_(){
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Fetch Orders for Day', 'Enter date (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const day = r.getResponseText().trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(day)) { ui.alert('Invalid date. Use YYYY-MM-DD.'); return; }

  // Port of Spain is UTC-4 (no DST).
  const fromISO = day + 'T00:00:00-04:00';
  const toISO   = day + 'T23:59:59-04:00';
  const all = SF_fetchOrders_({fromISO: fromISO, toISO: toISO, status:'any'});
  const n = SF_writeOrders_(all);
  ui.alert(n ? ('Fetched ' + n + ' order(s) into OrdersInbox.') : 'No orders for that date.');
}

function SF_fetchOrders_(opts) {
  const base = S_shopifyBase_().base;
  const token = S_shopifyBase_().token;
  const status = String((opts && opts.status) || 'any').toLowerCase()==='any' ? 'any' : undefined;
  const paramsBase = { limit:250, status: status, created_at_min: opts && opts.fromISO || undefined, created_at_max: opts && opts.toISO || undefined, order:'created_at asc' };
  var all=[]; var sinceId=null;
  while(true){
    var p = {}; Object.keys(paramsBase).forEach(function(k){ if(paramsBase[k]!==undefined) p[k]=paramsBase[k]; });
    if (sinceId) p.since_id = sinceId;
    const url = base + '/orders.json?' + Object.keys(p).map(function(k){ return encodeURIComponent(k)+'='+encodeURIComponent(String(p[k])); }).join('&');
    const res = UrlFetchApp.fetch(url, {method:'get', muteHttpExceptions:true, headers:{'X-Shopify-Access-Token': token}});
    if (res.getResponseCode() !== 200) throw new Error('Shopify GET '+res.getResponseCode()+' :: '+res.getContentText().slice(0,400));
    const batch = (JSON.parse(res.getContentText()).orders)||[];
    all = all.concat(batch);
    if (!batch.length || batch.length<250) break;
    sinceId = batch[batch.length-1].id;
  }
  return all;
}

function SF_getOrdersSheet_(){
  const ss=SpreadsheetApp.getActive();
  var sh=ss.getSheetByName('OrdersInbox');
  if(!sh){
    sh=ss.insertSheet('OrdersInbox');
    sh.getRange(1,1,1,12).setValues([[
      'order_id','order_number','created_at','financial_status','currency',
      'customer_name','customer_email','line_count','subtotal_price',
      'total_tax','total_price','raw_json'
    ]]);
    sh.setFrozenRows(1);
  }
  if (SF_AUTO_ADD_MARKER_COLUMNS) _ensureMarkerColumns_(sh);
  return sh;
}

// === Compact a Shopify order so it fits under the 50k cell limit ===
function SF_compactOrder_(o){
  var c = {
    id: o.id,
    order_number: o.order_number,
    created_at: o.created_at,
    currency: o.currency,
    email: o.email || (o.customer && o.customer.email) || ''
  };

  if (o.customer) {
    c.customer = {
      first_name: o.customer.first_name,
      last_name:  o.customer.last_name,
      email:      o.customer.email
    };
  }

  if (o.billing_address && o.billing_address.name) {
    c.billing_address = { name: o.billing_address.name };
  }

  var note = (o.note || '').toString();
  if (note.length > 4000) note = note.substring(0, 4000);
  c.note = note;

  if (Array.isArray(o.note_attributes) && o.note_attributes.length) {
    var na = [];
    for (var i = 0; i < o.note_attributes.length && na.length < 50; i++) {
      var a = o.note_attributes[i] || {};
      var k = (a.name || a.key || '').toString();
      var v = (a.value || '').toString();
      if (k.length > 100) k = k.substring(0, 100);
      if (v.length > 500) v = v.substring(0, 500);
      if (k || v) na.push({ name: k, value: v });
    }
    c.note_attributes = na;
  }

  c.line_items = (o.line_items || []).map(function(li){
    return {
      id:        li.id,
      title:     li.title || li.name,
      name:      li.name,
      sku:       li.sku,
      quantity:  li.quantity,
      price:     li.price,
      price_set: (li.price_set && li.price_set.shop_money)
                  ? { shop_money: { amount: li.price_set.shop_money.amount } }
                  : undefined,
      discount_allocations: Array.isArray(li.discount_allocations)
        ? li.discount_allocations.map(function(d){
            return { amount: d.amount || (d.amount_set && d.amount_set.shop_money && d.amount_set.shop_money.amount) || 0 };
          })
        : []
    };
  });

  c.shipping_lines = Array.isArray(o.shipping_lines)
    ? o.shipping_lines.map(function(sl){ return { price: sl.price }; })
    : [];

  c.total_discounts = o.total_discounts;
  return c;
}

function SF_writeOrders_(orders){
  if(!orders || !orders.length) return 0;
  const sh = SF_getOrdersSheet_();

  const rows = orders.map(function(o){
    const name = (o.customer ? [o.customer.first_name,o.customer.last_name].filter(Boolean).join(' ') : '');
    const email = o.email || (o.customer && o.customer.email) || '';

    // Choose raw_json payload based on flag
    var raw = JSON.stringify(o);
    if (SF_USE_COMPACT_JSON) {
      var compact = SF_compactOrder_(o);
      raw = JSON.stringify(compact);
      // Progressive trims if still large (very rare)
      if (raw.length > 49000) {
        compact.note_attributes = [];
        raw = JSON.stringify(compact);
      }
      if (raw.length > 49000) {
        compact.note = (compact.note || '').substring(0, 2000);
        raw = JSON.stringify(compact);
      }
      if (raw.length > 49000) {
        for (var i=0; i<(compact.line_items||[]).length; i++) {
          if (compact.line_items[i] && compact.line_items[i].price_set) delete compact.line_items[i].price_set;
        }
        raw = JSON.stringify(compact);
      }
    }

    return [
      o.id, o.order_number, o.created_at, o.financial_status||'', o.currency||'',
      name, email, Array.isArray(o.line_items)?o.line_items.length:0,
      Number(o.subtotal_price||0), Number(o.total_tax||0), Number(o.total_price||0),
      raw
    ];
  });

  const startRow = Math.max(2, sh.getLastRow() + 1);
  sh.getRange(startRow,1,rows.length,rows[0].length).setValues(rows);
  return rows.length;
}

function SF_testShopify_(){
  const cfg = S_shopifyBase_();
  const res = UrlFetchApp.fetch(cfg.base + '/shop.json', {method:'get', muteHttpExceptions:true, headers:{'X-Shopify-Access-Token': cfg.token}});
  SpreadsheetApp.getUi().alert('Shopify /shop.json → HTTP ' + res.getResponseCode());
}

// Auto-add marker columns if enabled
function _ensureMarkerColumns_(sh) {
  try {
    var headers = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0].map(String);
    var need = [];
    if (headers.indexOf(XP_LOCAL_MARKER_COLS.pushedAt)  < 0) need.push(XP_LOCAL_MARKER_COLS.pushedAt);
    if (headers.indexOf(XP_LOCAL_MARKER_COLS.invoiceId) < 0) need.push(XP_LOCAL_MARKER_COLS.invoiceId);
    if (!need.length) return;
    var startCol = sh.getLastColumn() + 1;
    sh.insertColumnsAfter(sh.getLastColumn(), need.length);
    sh.getRange(1, startCol, 1, need.length).setValues([need]);
  } catch(e){}
}
