// /** =========================================================================
//  *  Xero Push Main — Reset V1 (Standard) + Targeted Per-Line Fallback
//  *  - Entry: XL_pushSelectedDraft_InclusiveOnly_Standalone  (same as original)
//  *  - Standard behavior: DRAFT, Inclusive, keep healthy lines intact
//  *  - Fallback on Xero 400: only the failing item lines are downgraded
//  *    (strip ItemCode, set safe AccountCode, tag description), NOT the whole order
//  *  - Last-resort guard: if failing-line mapping is unknown, perform global
//  *    downgrade for that order (to avoid getting stuck)
//  *  - Uses your existing XL_getAccessToken_() and XL_ensureTenantId_()
//  * ========================================================================== */

// /** ---- Tunables for the targeted fallback (kept local to this file) ---- */
// var XP_DOWNGRADE_TAG                 = '[DOWNGRADED]'; // prefixed on downgraded lines
// var XP_FALLBACK_REVENUE_ACCT        = '200';          // safe AccountCode ('' to omit)
// var XP_KEEP_TAXCODE_ON_DOWNGRADE    = true;           // keep TaxType on downgraded lines

// /** MAIN ENTRY (same name as your original) */
// function XL_pushSelectedDraft_InclusiveOnly_Standalone() {
//   var ss  = SpreadsheetApp.getActive();
//   var sh  = ss.getSheetByName('OrdersInbox');
//   var sel = sh.getActiveRange();
//   if (!sel) { SpreadsheetApp.getUi().alert('Select one or more rows in OrdersInbox.'); return; }

//   // Resolve column positions by header names (more robust than hard-coded indexes)
//   var col = _oi_columns_(sh);

//   var rows = sel.getValues();
//   for (var i = 0; i < rows.length; i++) {
//     try {
//       var r   = rows[i];
//       var oid = _get(r, col.order_id);
//       if (!oid) continue;

//       var orderNo   = _get(r, col.order_number);
//       var createdAt = _get(r, col.created_at);
//       var raw       = _get(r, col.raw_json);

//       var ref      = 'Shopify #' + orderNo;                     // Reference
//       var contact  = _xp_buildContactFromRow_(r, col);          // minimal Contact
//       var dateStr  = _asYMD_(createdAt);                        // 'YYYY-MM-DD'

//       // Build line array (wrapped with meta so we can target by line)
//       var shop  = _safeParseJSON_(raw, {});
//       var items = _asArray_(shop.line_items || []);
//       var wrapped = items.map(function(li, idx){
//         var rawLine   = _xp_buildRawLine_FromShopify_(li);      // map Shopify → Xero line
//         var canDown   = true;                                   // allow downgrade on this line if needed
//         var reasons   = [];
//         if (!rawLine.ItemCode)   reasons.push('NO_ITEMCODE');
//         if (!rawLine.AccountCode) reasons.push('NO_ACCOUNT');   // informational
//         return _xp_wrapLine_(rawLine, { lineId: String(idx+1), canDowngrade: canDown, reasons: reasons });
//       });

//       // 1) Try full payload (healthy path)
//       var payload = _xp_buildInvoicePayload_(ref, contact, _xp_plainLines_(wrapped), dateStr);
//       var res     = _xp_xeroPostInvoices_(payload);
//       if (res.code === 200 || res.code === 201) {
//         _markPushed_(sh, sel.getRow() + i, col, res);
//         continue;
//       }

//       // 2) Targeted fallback ONLY on validation 400
//       if (res.code !== 400) throw new Error('Xero error ' + res.code + ': ' + (res.body || ''));
//       var failingIdxs = _xp_extractFailingLineIndexes_(res.body || '');  // e.g. [0,2]
//       if (!failingIdxs.length) {
//         // Unknown mapping → last-resort global downgrade for this order only
//         var globalDowngraded = wrapped.map(function(w){
//           return _xp_downgradeSpecificLines_([w], [w.__meta.lineId])[0];
//         });
//         payload = _xp_buildInvoicePayload_(ref, contact, _xp_plainLines_(globalDowngraded), dateStr);
//         res     = _xp_xeroPostInvoices_(payload);
//         if (res.code !== 200 && res.code !== 201) throw new Error('Xero push failed: ' + res.code + ' ' + (res.body || ''));
//         _markPushed_(sh, sel.getRow() + i, col, res);
//         continue;
//       }

//       // 3) Map failing indexes → our wrapped lineIds, downgrade ONLY those
//       var failingIds = failingIdxs
//         .filter(function(j){ return wrapped[j] && wrapped[j].__meta; })
//         .map(function(j){ return wrapped[j].__meta.lineId; });

//       var repaired = _xp_downgradeSpecificLines_(wrapped, failingIds, '[NO-ITEMCODE]');
//       payload = _xp_buildInvoicePayload_(ref, contact, _xp_plainLines_(repaired), dateStr);
//       res     = _xp_xeroPostInvoices_(payload);
//       if (res.code !== 200 && res.code !== 201) throw new Error('Xero push failed: ' + res.code + ' ' + (res.body || ''));
//       _markPushed_(sh, sel.getRow() + i, col, res);

//     } catch (e) {
//       Logger.log('Push error row ' + (sel.getRow()+i) + ': ' + e.message);
//       SpreadsheetApp.getUi().alert('Row ' + (sel.getRow()+i) + ' → ' + e.message);
//     }
//   }
// }

// /** ------------------------------ Helpers ------------------------------ */

// // Resolve OrdersInbox column positions by header label (case-insensitive)
// function _oi_columns_(sh) {
//   var headers = (sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0] || []).map(String);
//   function idx(h){ var k = headers.findIndex(function(x){ return x.toLowerCase().trim() === h; }); return k >= 0 ? k : null; }
//   return {
//     order_id:        idx('order_id'),
//     order_number:    idx('order_number'),
//     created_at:      idx('created_at'),
//     financial_status:idx('financial_status'),
//     currency:        idx('currency'),
//     customer_name:   idx('customer_name'),
//     customer_email:  idx('customer_email'),
//     line_count:      idx('line_count'),
//     subtotal_price:  idx('subtotal_price'),
//     total_tax:       idx('total_tax'),
//     total_price:     idx('total_price'),
//     raw_json:        idx('raw_json'),
//     pushed_at:       idx('PushedAt'),         // may be null if column not added yet
//     pushed_invoice:  idx('PushedInvoiceID')   // may be null
//   };
// }
// function _get(row, ix){ return (ix==null ? null : row[ix]); }
// function _asYMD_(d){
//   try { return d && d.toISOString ? d.toISOString().slice(0,10) : Utilities.formatDate(new Date(d), 'UTC', 'yyyy-MM-dd'); }
//   catch(e){ return Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd'); }
// }
// function _safeParseJSON_(s, fallback){ try{ return s ? JSON.parse(s) : fallback; }catch(e){ return fallback; } }

// // Build a minimal Contact payload from row
// function _xp_buildContactFromRow_(row, col) {
//   var name  = _get(row, col.customer_name)  || '';
//   var email = _get(row, col.customer_email) || '';
//   var out = {};
//   if (name)  out.Name = name;
//   if (email) out.EmailAddress = email;
//   return out;
// }

// // Map Shopify line_item → Xero LineItem fields you use
// function _xp_buildRawLine_FromShopify_(li) {
//   // Customize here if you have stronger mappings/GL rules elsewhere
//   var qty = Number(li.quantity || 0);
//   var unit = (li.price != null) ? Number(li.price) : (li.total && qty ? Number(li.total)/qty : 0);
//   return {
//     Description:  String(li.title || ''),
//     Quantity:     qty,
//     UnitAmount:   unit,
//     AccountCode:  li.account_code || '',                    // your GL mapping (if any) can fill this
//     ItemCode:     (li.sku || '').toString(),
//     TaxType:      (li.tax_code || (li.tax_lines && li.tax_lines[0] && li.tax_lines[0].title) || ''),
//     DiscountRate: null // keep null unless you explicitly compute %
//   };
// }

// // Wrap a line with metadata we can target on fallback
// function _xp_wrapLine_(raw, meta) {
//   return {
//     __meta: {
//       lineId: (meta && meta.lineId) || Utilities.getUuid(),
//       canDowngrade: !!(meta && meta.canDowngrade),
//       reasons: (meta && meta.reasons) || []
//     },
//     Description: raw.Description,
//     Quantity:    raw.Quantity,
//     UnitAmount:  raw.UnitAmount,
//     AccountCode: raw.AccountCode,
//     ItemCode:    raw.ItemCode,
//     TaxType:     raw.TaxType,
//     DiscountRate: raw.DiscountRate
//   };
// }

// // Strip meta so payload is legal
// function _xp_plainLines_(wrapped) {
//   return wrapped.map(function(w){
//     var out = { Description: w.Description, Quantity: w.Quantity, UnitAmount: w.UnitAmount };
//     if (w.AccountCode)          out.AccountCode  = w.AccountCode;
//     if (w.ItemCode)             out.ItemCode     = w.ItemCode;
//     if (w.TaxType)              out.TaxType      = w.TaxType;
//     if (w.DiscountRate != null) out.DiscountRate = w.DiscountRate;
//     return out;
//   });
// }

// // Downgrade ONLY specific lineIds (strip ItemCode, set fallback Account, tag description)
// function _xp_downgradeSpecificLines_(wrappedLines, failingLineIds, reasonTag) {
//   reasonTag = reasonTag || XP_DOWNGRADE_TAG || '[DOWNGRADED]';
//   var fallbackAcct = (typeof XP_FALLBACK_REVENUE_ACCT === 'string' ? XP_FALLBACK_REVENUE_ACCT : '');
//   var keepTax      = (typeof XP_KEEP_TAXCODE_ON_DOWNGRADE === 'boolean' ? XP_KEEP_TAXCODE_ON_DOWNGRADE : true);

//   var failing = {};
//   (failingLineIds || []).forEach(function(id){ failing[id] = true; });

//   return wrappedLines.map(function(w){
//     if (!failing[w.__meta.lineId] || !w.__meta.canDowngrade) return w;

//     var desc = w.Description || '';
//     if (reasonTag && desc.indexOf(reasonTag) !== 0) desc = reasonTag + ' ' + desc;

//     var d = {
//       __meta: w.__meta,
//       Description: desc,
//       Quantity:    w.Quantity,
//       UnitAmount:  w.UnitAmount
//     };
//     if (fallbackAcct) d.AccountCode = fallbackAcct;     // optional safe GL
//     if (keepTax && w.TaxType) d.TaxType = w.TaxType;   // preserve tax if present
//     if (w.DiscountRate != null) d.DiscountRate = w.DiscountRate;
//     return d; // Note: no ItemCode on downgraded lines
//   });
// }

// // Pull failing line indexes like "LineItems[2].ItemCode is invalid"
// function _xp_extractFailingLineIndexes_(bodyText) {
//   try {
//     var body = JSON.parse(bodyText || '{}');
//     var set = {};
//     (JSON.stringify(body).toLowerCase().match(/lineitems\[(\d+)\]/g) || [])
//       .forEach(function(tok){
//         var m = /lineitems\[(\d+)\]/i.exec(tok);
//         if (m) set[parseInt(m[1],10)] = 1;
//       });
//     return Object.keys(set).map(function(k){ return parseInt(k,10); });
//   } catch(e) { return []; }
// }

// // Build the Inclusive DRAFT payload
// function _xp_buildInvoicePayload_(ref, contact, lineItems, dateStr) {
//   return {
//     Invoices: [{
//       Type: 'ACCREC',
//       Status: 'DRAFT',
//       LineAmountTypes: 'Inclusive',
//       Reference: ref,
//       Contact: contact,
//       Date: dateStr,
//       LineItems: lineItems
//     }]
//   };
// }

// // POST to Xero (uses your token helpers)
// function _xp_xeroPostInvoices_(payload) {
//   var url = 'https://api.xero.com/api.xro/2.0/Invoices';
//   var access = XL_getAccessToken_();          // from your Tokens script
//   var tenant = XL_ensureTenantId_(access);    // from your Tokens script
//   var res = UrlFetchApp.fetch(url, {
//     method: 'post',
//     muteHttpExceptions: true,
//     headers: {
//       'Authorization': 'Bearer ' + access,
//       'Xero-tenant-id': tenant,
//       'Accept': 'application/json',
//       'Content-Type': 'application/json'
//     },
//     payload: JSON.stringify(payload)
//   });
//   return {
//     code: res.getResponseCode(),
//     body: res.getContentText(),
//     json: _safeParseJSON_(res.getContentText(), {})
//   };
// }

// // Stamp PushedAt & PushedInvoiceID (if those columns exist)
// function _markPushed_(sh, rowIndex1, col, res) {
//   if (col.pushed_at != null)      sh.getRange(rowIndex1, col.pushed_at + 1).setValue(new Date());
//   if (col.pushed_invoice != null) {
//     var id = (res.json && res.json.Invoices && res.json.Invoices[0] && res.json.Invoices[0].InvoiceID) || '';
//     sh.getRange(rowIndex1, col.pushed_invoice + 1).setValue(id);
//   }
// }
