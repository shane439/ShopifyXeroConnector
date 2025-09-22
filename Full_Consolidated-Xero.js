// /** Xero_Push_Main_CONSOLIDATED.gs
//  * Consolidation only. No improvising.
//  * - Discounts as VALUE (your current behavior)
//  * - Uses exact helper implementations from your tabs so nothing is undefined
//  * - Feature flags present (idempotency / 400 fallback / 429 backoff / attempt notes)
//  */

// // ========================= Menu (simple) =========================
// function onOpen() { try { SpreadsheetApp.getUi() .createMenu('Xero Push') .addItem('Push Selected (Inclusive Only)', 'XL_pushSelectedDraft_InclusiveOnly_Standalone') .addToUi(); } catch (e) {} }


// /* ============================================================================
//  * Feature Flags (baseline: ON as per your last spec)
//  * ==========================================================================*/
// var XP_IDEMPOTENCY_LOCAL_MARKER   = true;   // write InvoiceID to row; skip already pushed rows
// var XP_IDEMPOTENCY_XERO_CHECK     = true;   // query Xero for existing DRAFT by Reference; skip if found
// var XP_RATE_LIMIT_BACKOFF         = true;   // add per-invoice delay + one 60s retry on 429
// var XP_FALLBACK_ON_VALIDATION400  = true;   // on 400 only: retry once with ItemCode removed (GL kept)
// var XP_LOG_ATTEMPT_LAYER          = true;   // include (A/B) path info in popup/errors

// // Tunables (used only when the related flag is true)
// var XP_DELAY_MS                   = 1000;   // delay between invoices when backoff is enabled
// var XP_RETRY_429_SLEEP_MS         = 60000;  // sleep before single retry on 429 (60s)
// var XP_LOCAL_MARKER_COLS = {               // add these columns to OrdersInbox (or adjust)
//   pushedAt:   'PushedAt',
//   invoiceId:  'PushedInvoiceID'
// };

// /* ============================================================================
//  * 0) Compat + Auth (EXACT pull from your tabs)
//  *    - S_get_, S_bool_, S_int_, S_xeroAuth_  (SX_Compat_Settings)
//  *    - XL_getAccessToken_, XL_ensureTenantId_ (SX_Xero_Lite)
//  * ==========================================================================*/

// // ---- from SX_Compat_Settings ----
// const HARDCODE_OVERRIDE = false;
// const HARD = { /* kept as-is; values managed via Script Properties/CONFIG in your sheet */ };

// function S_get_(key, defVal) {
//   if (HARDCODE_OVERRIDE && Object.prototype.hasOwnProperty.call(HARD, key)) {
//     const v = HARD[key];
//     if (v !== null && v !== undefined && String(v) !== '') return String(v);
//   }
//   try {
//     const v = PropertiesService.getScriptProperties().getProperty(String(key));
//     if (v !== null && v !== undefined && String(v) !== '') return String(v);
//   } catch (e) {}
//   try {
//     const ss = SpreadsheetApp.getActive();
//     const tabs = ['CONFIG','config','Settings'];
//     for (var t=0; t<tabs.length; t++) {
//       const sh = ss.getSheetByName(tabs[t]);
//       if (sh && sh.getLastRow() >= 2) {
//         const rows = sh.getRange(2,1, sh.getLastRow()-1, 2).getValues();
//         const hit = rows.find(r => String(r[0]).trim() === String(key));
//         if (hit && String(hit[1]).trim() !== '') return String(hit[1]).trim();
//       }
//     }
//   } catch (e) {}
//   return defVal;
// }
// function S_bool_(key, def=false) { return ['true','1','yes','y'].includes(String(S_get_(key, def)).trim().toLowerCase()); }
// function S_int_(key, def=0) { const n = parseInt(S_get_(key, def), 10); return isNaN(n) ? def : n; }

// function S_xeroAuth_() {
//   const clientId     = S_get_('XERO_CLIENT_ID','');
//   const clientSecret = S_get_('XERO_CLIENT_SECRET','');
//   const tenantId     = S_get_('XERO_TENANT_ID','');
//   const refresh      = S_get_('XERO_REFRESH_TOKEN','');
//   if (!clientId || !clientSecret) throw new Error('Xero client credentials missing.');
//   if (!refresh) throw new Error('Xero refresh token missing. Reconnect OAuth.');
//   return { clientId, clientSecret, tenantId, refresh };
// }

// // ---- from SX_Xero_Lite ----
// function XL_getAccessToken_(){
//   if (typeof S_xeroAuth_ !== 'function') throw new Error('S_xeroAuth_ not found. Check Compat_Settings.');
//   const a = S_xeroAuth_();
//   const tok = UrlFetchApp.fetch('https://identity.xero.com/connect/token', {
//     method:'post', contentType:'application/x-www-form-urlencoded',
//     payload: { grant_type: 'refresh_token', refresh_token: a.refresh },
//     headers: { Authorization: 'Basic ' + Utilities.base64Encode(a.clientId + ':' + a.clientSecret) },
//     muteHttpExceptions:true
//   });
//   const code = tok.getResponseCode();
//   const txt  = tok.getContentText();
//   if (code!==200) throw new Error('Xero token failed '+code+' :: '+txt.slice(0,400));
//   const body = JSON.parse(txt);
//   if (body.refresh_token) {
//     try { PropertiesService.getScriptProperties().setProperty('XERO_REFRESH_TOKEN', body.refresh_token); } catch(e){}
//   }
//   return body.access_token;
// }

// function XL_ensureTenantId_(accessToken) {
//   let tenantId = (typeof S_get_ === 'function') ? S_get_('XERO_TENANT_ID','') : '';
//   if (tenantId) return tenantId;
//   const con = UrlFetchApp.fetch('https://api.xero.com/connections', {
//     headers:{Authorization:'Bearer '+accessToken, Accept:'application/json'},
//     muteHttpExceptions:true
//   });
//   const code = con.getResponseCode();
//   const txt  = con.getContentText();
//   if (code===200) {
//     const arr = JSON.parse(txt) || [];
//     if (arr[0] && (arr[0].tenantId || arr[0].tenant_id)) {
//       tenantId = arr[0].tenantId || arr[0].tenant_id;
//       try { PropertiesService.getScriptProperties().setProperty('XERO_TENANT_ID', tenantId); } catch(e){}
//       return tenantId;
//     }
//   }
//   throw new Error('Xero tenant not set. Reconnect OAuth.');
// }

// /* ============================================================================
//  * 1) ENTRY POINT — your working A-path with optional wrappers behind flags
//  * ==========================================================================*/
// function XL_pushSelectedDraft_InclusiveOnly_Standalone() {
//   var ss = SpreadsheetApp.getActive();
//   var sh = ss.getSheetByName('OrdersInbox');
//   if (!sh) { SpreadsheetApp.getUi().alert('OrdersInbox not found.'); return; }

//   var headers = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0].map(String);
//   var col = function(n){ return headers.indexOf(n) + 1; };

//   var ranges = (sh.getActiveRangeList() ? sh.getActiveRangeList().getRanges() : [sh.getActiveRange()]).filter(Boolean);
//   if (!ranges.length) { SpreadsheetApp.getUi().alert('Select one or more data rows.'); return; }

//   var ok = 0, fail = 0, errs = [];
//   var accessToken = XL_getAccessToken_();
//   var tenantId    = XL_ensureTenantId_(accessToken);

//   var mk = _xp_resolveMarkerCols_(headers);

//   ranges.forEach(function(rng){
//     var rows = sh.getRange(rng.getRow(), 1, rng.getNumRows(), sh.getLastColumn()).getValues();
//     rows.forEach(function(r, idx){
//       try {
//         if (String(r[0]).toLowerCase() === 'order_id') return;

//         if (XP_IDEMPOTENCY_LOCAL_MARKER && _xp_rowAlreadyPushed_(r, mk)) {
//           if (XP_LOG_ATTEMPT_LAYER) errs.push('SKIP (already pushed) row='+(rng.getRow()+idx));
//           ok++; return;
//         }

//         var raw = String(r[col('raw_json') - 1] || '').trim();
//         if (!raw) throw new Error('raw_json empty for selected row.');
//         var order = JSON.parse(raw);

//         var reference = 'Shopify #' + String(order && order.order_number || '');

//         if (XP_IDEMPOTENCY_XERO_CHECK && _xp_findExistingDraftByReference_(accessToken, tenantId, reference)) {
//           if (XP_LOG_ATTEMPT_LAYER) errs.push('SKIP (existing draft on Xero) '+reference);
//           ok++; return;
//         }

//         if (XP_RATE_LIMIT_BACKOFF) Utilities.sleep(XP_DELAY_MS);

//         // 1) Build lines — discounts as VALUE (your current behavior)
//         var lineBuild = _buildMonetaryLinesWithDiscounts_(order);

//         // 2) Notes line
//         var notes = _extractNotes_(order);
//         if (notes) {
//           var noteLine = _makeDescriptionOnlyLine_(notes);
//           if (noteLine) lineBuild.lines.push(noteLine);
//         }

//         // 3) Payload (Inclusive + order date = due date)
//         var payload = _buildInclusiveInvoicePayload_UsingOrderDate_(order, lineBuild.lines);

//         // A-path POST
//         var res = UrlFetchApp.fetch('https://api.xero.com/api.xro/2.0/Invoices', {
//           method: 'post',
//           headers: {
//             Authorization: 'Bearer ' + accessToken,
//             'xero-tenant-id': tenantId,
//             Accept: 'application/json',
//             'Content-Type': 'application/json'
//           },
//           payload: JSON.stringify(payload),
//           muteHttpExceptions: true
//         });

//         // Optional 429 backoff + single retry
//         if (XP_RATE_LIMIT_BACKOFF && res.getResponseCode() === 429) {
//           Utilities.sleep(XP_RETRY_429_SLEEP_MS);
//           res = UrlFetchApp.fetch('https://api.xero.com/api.xro/2.0/Invoices', {
//             method: 'post',
//             headers: {
//               Authorization: 'Bearer ' + accessToken,
//               'xero-tenant-id': tenantId,
//               Accept: 'application/json',
//               'Content-Type': 'application/json'
//             },
//             payload: JSON.stringify(payload),
//             muteHttpExceptions: true
//           });
//         }

//         // Optional 400 fallback: strip ItemCode only
//         if (XP_FALLBACK_ON_VALIDATION400 && _xp_isValidation400_(res)) {
//           if (XP_LOG_ATTEMPT_LAYER) errs.push('Fallback B for '+reference);
//           var fallbackLines = _xp_stripItemCodeOnly_(lineBuild.lines);
//           var retryPayload  = _buildInclusiveInvoicePayload_UsingOrderDate_(order, fallbackLines);
//           if (XP_RATE_LIMIT_BACKOFF) Utilities.sleep(XP_DELAY_MS);
//           res = UrlFetchApp.fetch('https://api.xero.com/api.xro/2.0/Invoices', {
//             method: 'post',
//             headers: {
//               Authorization: 'Bearer ' + accessToken,
//               'xero-tenant-id': tenantId,
//               Accept: 'application/json',
//               'Content-Type': 'application/json'
//             },
//             payload: JSON.stringify(retryPayload),
//             muteHttpExceptions: true
//           });
//         }

//         if (res.getResponseCode() >= 300) {
//           throw new Error('Create failed ' + res.getResponseCode() + ' :: ' + res.getContentText().slice(0, 1200));
//         }

//         // Enforce Inclusive via header-only PUT (no LineItems in PUT)
//         try {
//           var created = JSON.parse(res.getContentText());
//           var inv = created && created.Invoices && created.Invoices[0];
//           if (inv && inv.InvoiceID) {
//             _forceInclusiveHeaderOnly_(accessToken, tenantId, inv);
//             if (XP_IDEMPOTENCY_LOCAL_MARKER) {
//               _xp_markRowPushed_(sh, (rng.getRow() + idx), mk, inv.InvoiceID);
//             }
//           }
//         } catch (e) { /* ignore */ }

//         ok++;
//       } catch (e) {
//         fail++;
//         errs.push(String(e && e.message ? e.message : e));
//       }
//     });
//   });

//   SpreadsheetApp.getUi().alert('Inclusive push — ok: ' + ok + '  fail: ' + fail + (errs.length ? ('\n' + errs.join('\n')) : ''));
// }

// /* ============================================================================
//  * 2) Builders & helpers — EXACT behavior you’re running now
//  * ==========================================================================*/

// /** Discounts as VALUE (net UnitAmount) + pro-rate order.total_discounts if needed */
// function _buildMonetaryLinesWithDiscounts_(order) {
//   var out = [];
//   var items = (order && Array.isArray(order.line_items)) ? order.line_items : [];

//   var productMeta = [];
//   for (var i = 0; i < items.length; i++) {
//     var li = items[i] || {};
//     var qty = Number(li.quantity || 0);
//     if (!qty) continue;

//     var unit = _toNumberSafe_(li.price, li.price_set && li.price_set.shop_money && li.price_set.shop_money.amount);
//     var desc = (li.title || '').toString().substring(0, 4000);

//     var line = { Description: desc, Quantity: qty, UnitAmount: unit, TaxType: 'OUTPUT' };
//     if (li.sku) line.ItemCode = String(li.sku).trim();

//     var liDisc = _sumDiscountAllocations_(li);
//     if (liDisc > 0) {
//       var ext = qty * unit;
//       if (ext > 0) {
//         var netExt = ext - liDisc;
//         var netUnit = netExt / qty;
//         line.UnitAmount = Math.round(netUnit * 100) / 100;
//       }
//     }

//     out.push(line);
//     productMeta.push({ index: out.length - 1, extended: qty * unit });
//   }

//   var shippingTotal = Array.isArray(order && order.shipping_lines)
//     ? order.shipping_lines.reduce(function(s,x){ return s + _toNumberSafe_(x.price); }, 0)
//     : 0;
//   if (shippingTotal > 0) {
//     out.push({ Description: 'Shipping', Quantity: 1, UnitAmount: Number(shippingTotal), TaxType: 'OUTPUT' });
//   }

//   var anyAdjusted = out.some(function(l, idx){ return l.UnitAmount !== undefined && l.UnitAmount !== productMeta[idx]?.extended; });
//   var orderTotalDiscount = _toNumberSafe_(order && order.total_discounts);
//   if (!anyAdjusted && orderTotalDiscount > 0 && productMeta.length > 0) {
//     var base = productMeta.reduce(function(s, m){ return s + (m.extended > 0 ? m.extended : 0); }, 0);
//     if (base > 0) {
//       for (var j = 0; j < productMeta.length; j++) {
//         var m = productMeta[j];
//         if (m.extended <= 0) continue;
//         var share = orderTotalDiscount * (m.extended / base);
//         var netExt = m.extended - share;
//         var netUnit = netExt / (out[m.index].Quantity || 1);
//         out[m.index].UnitAmount = Math.round(netUnit * 100) / 100;
//       }
//     }
//   }

//   return { lines: out };
// }

// function _extractNotes_(order) {
//   try {
//     var parts = [];
//     if (order && order.note && String(order.note).trim()) parts.push(String(order.note).trim());
//     if (order && Array.isArray(order.note_attributes)) {
//       order.note_attributes.forEach(function(attr){
//         if (!attr) return;
//         var k = (attr.name || attr.key || '').toString().trim();
//         var v = (attr.value || '').toString().trim();
//         if (k || v) parts.push(k ? (k + ': ' + v) : v);
//       });
//     }
//     var all = parts.join('\n').trim();
//     return all ? all : '';
//   } catch (e) { return ''; }
// }

// function _makeDescriptionOnlyLine_(text) {
//   var t = (text || '').toString();
//   return t ? { Description: t.substring(0, 4000), Quantity: 0, UnitAmount: 0, TaxType: 'NONE', AccountCode: null } : null;
// }

// function _buildInclusiveInvoicePayload_UsingOrderDate_(order, lineItems) {
//   var name =
//     (order && order.customer
//       ? [order.customer.first_name, order.customer.last_name].filter(Boolean).join(' ')
//       : (order && order.billing_address && order.billing_address.name)) ||
//     'Shopify Customer';
//   var email = (order && (order.email || (order.customer && order.customer.email))) || undefined;

//   var created = new Date(order && order.created_at ? order.created_at : new Date());
//   var xDate = created.toISOString().slice(0,10);

//   var invoice = {
//     Type: 'ACCREC',
//     Status: 'DRAFT',
//     Date: xDate,
//     DueDate: xDate,
//     Reference: 'Shopify #' + String(order && order.order_number || ''),
//     Contact: (email ? { Name: name, EmailAddress: email } : { Name: name }),
//     LineItems: lineItems,
//     LineAmountTypes: 'Inclusive'
//   };

//   if (order && order.currency) invoice.CurrencyCode = String(order.currency);
//   return { Invoices: [ invoice ] };
// }

// function _sumDiscountAllocations_(li) {
//   try {
//     var arr = Array.isArray(li && li.discount_allocations) ? li.discount_allocations : [];
//     var sum = 0;
//     for (var i = 0; i < arr.length; i++) {
//       var a = arr[i] || {};
//       sum += _toNumberSafe_(a.amount, a.amount_set && a.amount_set.shop_money && a.amount_set.shop_money.amount);
//     }
//     return Number(sum) || 0;
//   } catch (e) { return 0; }
// }
// function _rateFromAmount_(discountAmt, extendedAmt) { // retained for parity if you ever revert to % mode
//   if (!extendedAmt || extendedAmt <= 0) return 0;
//   var pct = (Number(discountAmt) / Number(extendedAmt)) * 100;
//   if (!isFinite(pct) || pct <= 0) return 0;
//   pct = Math.max(0, Math.min(100, pct));
//   return Math.round(pct * 10000) / 10000;
// }
// function _toNumberSafe_() {
//   for (var i = 0; i < arguments.length; i++) {
//     var v = arguments[i];
//     if (v === null || v === undefined) continue;
//     var n = Number(v);
//     if (!isNaN(n)) return n;
//     var s = String(v).trim();
//     if (s) {
//       var n2 = Number(s);
//       if (!isNaN(n2)) return n2;
//     }
//   }
//   return 0;
// }

// function _forceInclusiveHeaderOnly_(accessToken, tenantId, inv) {
//   var contact = {};
//   if (inv.Contact && inv.Contact.ContactID) contact = { ContactID: inv.Contact.ContactID };
//   else if (inv.Contact && inv.Contact.Name) contact = { Name: inv.Contact.Name };
//   else contact = { Name: 'Shopify Customer' };

//   var fixPayload = { Invoices: [{ InvoiceID: inv.InvoiceID, Type: 'ACCREC', Status: 'DRAFT', Contact: contact, LineAmountTypes: 'Inclusive' }] };
//   var resPut = UrlFetchApp.fetch('https://api.xero.com/api.xro/2.0/Invoices', {
//     method: 'put',
//     headers: { Authorization: 'Bearer ' + accessToken, 'xero-tenant-id': tenantId, Accept: 'application/json', 'Content-Type': 'application/json' },
//     payload: JSON.stringify(fixPayload),
//     muteHttpExceptions: true
//   });
//   if (resPut.getResponseCode() >= 300) {
//     throw new Error('Inclusive header-only fix-up failed ' + resPut.getResponseCode() + ' :: ' + resPut.getContentText().slice(0, 1200));
//   }
// }

// /* ============================================================================
//  * 3) Optional wrappers (only if flags true)
//  * ==========================================================================*/
// function _xp_isValidation400_(resp) {
//   try {
//     if (resp.getResponseCode() !== 400) return false;
//     var j = JSON.parse(resp.getContentText());
//     return String(j && j.Type).toLowerCase().indexOf('validation') >= 0;
//   } catch (e) { return false; }
// }
// function _xp_stripItemCodeOnly_(lines) {
//   var clone = [];
//   for (var i = 0; i < lines.length; i++) {
//     var l = JSON.parse(JSON.stringify(lines[i] || {}));
//     if ('ItemCode' in l) delete l.ItemCode;
//     clone.push(l);
//   }
//   return clone;
// }
// function _xp_findExistingDraftByReference_(accessToken, tenantId, reference) {
//   try {
//     var where = 'Reference=="' + reference.replace(/"/g,'\\"') + '" AND Status=="DRAFT"';
//     var url = 'https://api.xero.com/api.xro/2.0/Invoices?where=' + encodeURIComponent(where) + '&page=1';
//     var resp = UrlFetchApp.fetch(url, {
//       method: 'get',
//       headers: { Authorization: 'Bearer ' + accessToken, 'xero-tenant-id': tenantId, Accept: 'application/json' },
//       muteHttpExceptions: true
//     });
//     if (resp.getResponseCode() >= 300) return false;
//     var j = JSON.parse(resp.getContentText());
//     var arr = (j && j.Invoices) || [];
//     return arr.length > 0;
//   } catch (e) { return false; }
// }
// function _xp_resolveMarkerCols_(headers) {
//   return {
//     pushedAtIdx:  headers.indexOf(XP_LOCAL_MARKER_COLS.pushedAt)  + 1,
//     invoiceIdIdx: headers.indexOf(XP_LOCAL_MARKER_COLS.invoiceId) + 1
//   };
// }
// function _xp_rowAlreadyPushed_(row, mk) {
//   try {
//     if (!mk.pushedAtIdx || !mk.invoiceIdIdx) return false;
//     var pushedAt  = row[mk.pushedAtIdx  - 1];
//     var invoiceId = row[mk.invoiceIdIdx - 1];
//     return Boolean(pushedAt && invoiceId);
//   } catch (e) { return false; }
// }
// function _xp_markRowPushed_(sheet, absRowIndex, mk, invoiceId) {
//   try {
//     if (!mk.pushedAtIdx || !mk.invoiceIdIdx) return;
//     var now = new Date();
//     sheet.getRange(absRowIndex, mk.pushedAtIdx).setValue(now);
//     sheet.getRange(absRowIndex, mk.invoiceIdIdx).setValue(invoiceId);
//   } catch (e) { /* best effort */ }
// }
