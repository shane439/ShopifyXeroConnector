// /** Xero_Push_Main_CONSOLIDATED.gs — Full consolidation (no improvising)
//  *
//  * Includes:
//  *  - Menus: Fetch Orders (Shopify), Push Selected (Inclusive Only), Test Shopify/Xero, Xero OAuth (prompt-paste)
//  *  - Settings helpers (CONFIG / Script Properties) + Shopify base builder
//  *  - Shopify fetch → OrdersInbox
//  *  - Xero OAuth helpers (prompt/paste) + optional WebApp doGet handler
//  *  - Xero push (Draft, VAT Inclusive), value-discounts, notes line, currency carry-through
//  *  - Idempotency: local row markers + Xero DRAFT pre-check by Reference
//  *  - Backoff: 1s pacing + single 60s retry on 429
//  *
//  * Behavior locked to existing tabs from your Sheet. Sources: Shopify_Lite, Xero_Lite, Compat_Settings,
//  * Xero_OAuth_Lite, Xero_OAuth_WebApp. Secrets are read from Script Properties/CONFIG.
//  */

// // ========================= 0) Menu =========================
// function onOpen() {
//   const ui = SpreadsheetApp.getUi();
//   try {
//     ui.createMenu('Shopify ↔ Xero')
//       .addItem('Fetch Orders by Date…', 'SF_promptFetchByDate_')
//       .addSeparator()
//       .addItem('Push Selected (Inclusive Only)', 'XL_pushSelectedDraft_InclusiveOnly_Standalone')
//       .addSeparator()
//       .addItem('Test Shopify', 'SF_testShopify_')
//       .addItem('Test Xero', 'XL_testXero_')
//       .addSeparator()
//       .addItem('Xero → Reconnect (OAuth)…', 'XO_startAuth')
//       .addToUi();
//   } catch (e) {}
// }

// /* ============================================================================
//  * 1) Feature Flags (baseline: ON per your spec)
//  * ==========================================================================*/
// var XP_IDEMPOTENCY_LOCAL_MARKER   = true;   // write InvoiceID to row; skip already pushed rows
// var XP_IDEMPOTENCY_XERO_CHECK     = true;   // query Xero for existing DRAFT by Reference; skip if found
// var XP_RATE_LIMIT_BACKOFF         = true;   // per-invoice delay + one 60s retry on 429
// var XP_FALLBACK_ON_VALIDATION400  = true;   // on HTTP 400 “Validation” only: retry without ItemCode
// var XP_LOG_ATTEMPT_LAYER          = true;   // include (A/B) path info in alert

// // New, localized hardeners (OFF by default — enable only if/when you approve)
// var SF_USE_COMPACT_JSON           = true;  // when true, compacts raw_json to avoid 50k cell limit
// var SF_AUTO_ADD_MARKER_COLUMNS    = false;  // when true, auto-adds PushedAt/PushedInvoiceID headers if missing

// // Tunables (used only when the related flag is true)
// var XP_DELAY_MS                   = 1000;   // 1s pacing
// var XP_RETRY_429_SLEEP_MS         = 60000;  // 60s before retry
// var XP_LOCAL_MARKER_COLS = {
//   pushedAt:   'PushedAt',
//   invoiceId:  'PushedInvoiceID'
// };

// /* ============================================================================
//  * 2) Settings & Base Builders (from Compat_Settings + Shopify_Lite)
//  *    - We default to Script Properties / CONFIG. No hardcoded secrets.
//  * ==========================================================================*/

// // Hardcoded override OFF (use Script Properties / CONFIG tab)
// const HARDCODE_OVERRIDE = false;
// const HARD = {}; // no inline secrets

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
//     const tabs = ['CONFIG', 'config', 'Settings'];
//     for (var t = 0; t < tabs.length; t++) {
//       const sh = ss.getSheetByName(tabs[t]);
//       if (sh && sh.getLastRow() >= 2) {
//         const rows = sh.getRange(2,1, sh.getLastRow()-1, 2).getValues();
//         const hit = rows.find(function(r){ return String(r[0]).trim() === String(key); });
//         if (hit && String(hit[1]).trim() !== '') return String(hit[1]).trim();
//       }
//     }
//   } catch (e) {}
//   return defVal;
// }
// function S_bool_(key, def) { return ['true','1','yes','y'].includes(String(S_get_(key, def)).trim().toLowerCase()); }
// function S_int_(key, def)  { const n = parseInt(S_get_(key, def), 10); return isNaN(n) ? def : n; }

// // Shopify base (verbatim behavior)
// function S_shopifyBase_() {
//   var dom = String(S_get_('SHOPIFY_STORE_DOMAIN','')).trim().replace(/^https?:\/\//,'').replace(/\/+$/,'');
//   if (!/\.myshopify\.com$/i.test(dom)) throw new Error('SHOPIFY_STORE_DOMAIN must end with .myshopify.com');

//   var raw = S_get_('SHOPIFY_API_VERSION', '2024-07');
//   var ver = '';
//   if (Object.prototype.toString.call(raw) === '[object Date]') {
//     var d = raw;
//     ver = String(d.getFullYear()) + '-' + ('0' + (d.getMonth()+1)).slice(-2);
//   } else {
//     var m = String(raw).match(/(\d{4}-\d{2})/);
//     ver = m ? m[1] : '';
//   }
//   if (!/^\d{4}-\d{2}$/.test(ver)) throw new Error('SHOPIFY_API_VERSION must be YYYY-MM (e.g., 2024-07)');

//   var token = S_get_('SHOPIFY_ADMIN_TOKEN', '');
//   if (!token) throw new Error('SHOPIFY_ADMIN_TOKEN missing.');

//   return { dom: dom, ver: ver, token: token, base: 'https://' + dom + '/admin/api/' + ver };
// }

// // Xero auth settings
// function S_xeroAuth_() {
//   const clientId     = S_get_('XERO_CLIENT_ID','');
//   const clientSecret = S_get_('XERO_CLIENT_SECRET','');
//   const tenantId     = S_get_('XERO_TENANT_ID','');
//   const refresh      = S_get_('XERO_REFRESH_TOKEN','');
//   if (!clientId || !clientSecret) throw new Error('Xero client credentials missing.');
//   if (!refresh) throw new Error('Xero refresh token missing. Reconnect OAuth.');
//   return { clientId: clientId, clientSecret: clientSecret, tenantId: tenantId, refresh: refresh };
// }

// /* ============================================================================
//  * 3) Shopify Fetch (from Shopify_Lite)
//  * ==========================================================================*/

// function SF_promptFetchByDate_(){
//   const ui = SpreadsheetApp.getUi();
//   const r = ui.prompt('Fetch Orders for Day', 'Enter date (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
//   if (r.getSelectedButton() !== ui.Button.OK) return;
//   const day = r.getResponseText().trim();
//   if (!/^\d{4}-\d{2}-\d{2}$/.test(day)) { ui.alert('Invalid date. Use YYYY-MM-DD.'); return; }

//   // Port of Spain is UTC-4 (no DST).
//   const fromISO = day + 'T00:00:00-04:00';
//   const toISO   = day + 'T23:59:59-04:00';
//   const all = SF_fetchOrders_({fromISO: fromISO, toISO: toISO, status:'any'});
//   const n = SF_writeOrders_(all);
//   ui.alert(n ? ('Fetched ' + n + ' order(s) into OrdersInbox.') : 'No orders for that date.');
// }

// function SF_fetchOrders_(opts) {
//   const base = S_shopifyBase_().base;
//   const token = S_shopifyBase_().token;
//   const status = String((opts && opts.status) || 'any').toLowerCase()==='any' ? 'any' : undefined;
//   const paramsBase = { limit:250, status: status, created_at_min: opts && opts.fromISO || undefined, created_at_max: opts && opts.toISO || undefined, order:'created_at asc' };
//   var all=[]; var sinceId=null;
//   while(true){
//     var p = {}; Object.keys(paramsBase).forEach(function(k){ if(paramsBase[k]!==undefined) p[k]=paramsBase[k]; });
//     if (sinceId) p.since_id = sinceId;
//     const url = base + '/orders.json?' + Object.keys(p).map(function(k){ return encodeURIComponent(k)+'='+encodeURIComponent(String(p[k])); }).join('&');
//     const res = UrlFetchApp.fetch(url, {method:'get', muteHttpExceptions:true, headers:{'X-Shopify-Access-Token': token}});
//     if (res.getResponseCode() !== 200) throw new Error('Shopify GET '+res.getResponseCode()+' :: '+res.getContentText().slice(0,400));
//     const batch = (JSON.parse(res.getContentText()).orders)||[];
//     all = all.concat(batch);
//     if (!batch.length || batch.length<250) break;
//     sinceId = batch[batch.length-1].id;
//   }
//   return all;
// }

// function SF_getOrdersSheet_(){
//   const ss=SpreadsheetApp.getActive();
//   var sh=ss.getSheetByName('OrdersInbox');
//   if(!sh){
//     sh=ss.insertSheet('OrdersInbox');
//     sh.getRange(1,1,1,12).setValues([[
//       'order_id','order_number','created_at','financial_status','currency',
//       'customer_name','customer_email','line_count','subtotal_price',
//       'total_tax','total_price','raw_json'
//     ]]);
//     sh.setFrozenRows(1);
//   }
//   if (SF_AUTO_ADD_MARKER_COLUMNS) _ensureMarkerColumns_(sh);
//   return sh;
// }

// // === Compact a Shopify order so it fits under the 50k cell limit (OFF by default) ===
// function SF_compactOrder_(o){
//   var c = {
//     id: o.id,
//     order_number: o.order_number,
//     created_at: o.created_at,
//     currency: o.currency,
//     email: o.email || (o.customer && o.customer.email) || ''
//   };

//   if (o.customer) {
//     c.customer = {
//       first_name: o.customer.first_name,
//       last_name:  o.customer.last_name,
//       email:      o.customer.email
//     };
//   }

//   if (o.billing_address && o.billing_address.name) {
//     c.billing_address = { name: o.billing_address.name };
//   }

//   var note = (o.note || '').toString();
//   if (note.length > 4000) note = note.substring(0, 4000);
//   c.note = note;

//   if (Array.isArray(o.note_attributes) && o.note_attributes.length) {
//     var na = [];
//     for (var i = 0; i < o.note_attributes.length && na.length < 50; i++) {
//       var a = o.note_attributes[i] || {};
//       var k = (a.name || a.key || '').toString();
//       var v = (a.value || '').toString();
//       if (k.length > 100) k = k.substring(0, 100);
//       if (v.length > 500) v = v.substring(0, 500);
//       if (k || v) na.push({ name: k, value: v });
//     }
//     c.note_attributes = na;
//   }

//   c.line_items = (o.line_items || []).map(function(li){
//     return {
//       id:        li.id,
//       title:     li.title || li.name,
//       name:      li.name,
//       sku:       li.sku,
//       quantity:  li.quantity,
//       price:     li.price,
//       price_set: (li.price_set && li.price_set.shop_money)
//                   ? { shop_money: { amount: li.price_set.shop_money.amount } }
//                   : undefined,
//       discount_allocations: Array.isArray(li.discount_allocations)
//         ? li.discount_allocations.map(function(d){
//             return { amount: d.amount || (d.amount_set && d.amount_set.shop_money && d.amount_set.shop_money.amount) || 0 };
//           })
//         : []
//     };
//   });

//   c.shipping_lines = Array.isArray(o.shipping_lines)
//     ? o.shipping_lines.map(function(sl){ return { price: sl.price }; })
//     : [];

//   c.total_discounts = o.total_discounts;
//   return c;
// }

// function SF_writeOrders_(orders){
//   if(!orders || !orders.length) return 0;
//   const sh = SF_getOrdersSheet_();

//   const rows = orders.map(function(o){
//     const name = (o.customer ? [o.customer.first_name,o.customer.last_name].filter(Boolean).join(' ') : '');
//     const email = o.email || (o.customer && o.customer.email) || '';

//     // Choose raw_json payload based on flag
//     var raw = JSON.stringify(o);
//     if (SF_USE_COMPACT_JSON) {
//       var compact = SF_compactOrder_(o);
//       raw = JSON.stringify(compact);
//       // Progressive trims if still large (very rare)
//       if (raw.length > 49000) {
//         compact.note_attributes = [];
//         raw = JSON.stringify(compact);
//       }
//       if (raw.length > 49000) {
//         compact.note = (compact.note || '').substring(0, 2000);
//         raw = JSON.stringify(compact);
//       }
//       if (raw.length > 49000) {
//         for (var i=0; i<(compact.line_items||[]).length; i++) {
//           if (compact.line_items[i] && compact.line_items[i].price_set) delete compact.line_items[i].price_set;
//         }
//         raw = JSON.stringify(compact);
//       }
//     }

//     return [
//       o.id, o.order_number, o.created_at, o.financial_status||'', o.currency||'',
//       name, email, Array.isArray(o.line_items)?o.line_items.length:0,
//       Number(o.subtotal_price||0), Number(o.total_tax||0), Number(o.total_price||0),
//       raw
//     ];
//   });

//   const startRow = Math.max(2, sh.getLastRow() + 1);
//   sh.getRange(startRow,1,rows.length,rows[0].length).setValues(rows);
//   return rows.length;
// }

// function SF_testShopify_(){
//   const cfg = S_shopifyBase_();
//   const res = UrlFetchApp.fetch(cfg.base + '/shop.json', {method:'get', muteHttpExceptions:true, headers:{'X-Shopify-Access-Token': cfg.token}});
//   SpreadsheetApp.getUi().alert('Shopify /shop.json → HTTP ' + res.getResponseCode());
// }

// /* ============================================================================
//  * 4) Xero OAuth (from Xero_OAuth_Lite) + Access Token Refresh (from Xero_Lite)
//  * ==========================================================================*/

// function XO_startAuth(){
//   const clientId = S_get_('XERO_CLIENT_ID','');
//   const redirect = S_get_('XERO_REDIRECT_URI','');
//   if (!clientId || !redirect) { SpreadsheetApp.getUi().alert('Missing XERO_CLIENT_ID or XERO_REDIRECT_URI in CONFIG/Properties.'); return; }

//   const scopes = [
//     'openid','profile','email','offline_access',
//     'accounting.settings','accounting.transactions','accounting.contacts','accounting.journals.read'
//   ].join(' ');
//   const url = 'https://login.xero.com/identity/connect/authorize'
//     + '?response_type=code'
//     + '&client_id=' + encodeURIComponent(clientId)
//     + '&redirect_uri=' + encodeURIComponent(redirect)
//     + '&scope=' + encodeURIComponent(scopes)
//     + '&prompt=consent';

//   const ui = SpreadsheetApp.getUi();
//   ui.alert(
//     'Xero Reconnect',
//     '1) Open this URL in your browser and approve:\n\n' + url + '\n\n2) You will be redirected to the Redirect URI with a long "code" in the URL.\n3) Copy ONLY that code (not the whole URL).\n4) Then run: Xero → Reconnect (OAuth)… again to paste the code.',
//     ui.ButtonSet.OK
//   );

//   const r = ui.prompt('Paste the OAuth CODE here:', ui.ButtonSet.OK_CANCEL);
//   if (r.getSelectedButton() !== ui.Button.OK) return;
//   const code = r.getResponseText().trim();
//   if (!code) { ui.alert('No code pasted.'); return; }

//   XO_exchangeCode_(code);
// }

// function XO_exchangeCode_(code){
//   const clientId = S_get_('XERO_CLIENT_ID','');
//   const clientSecret = S_get_('XERO_CLIENT_SECRET','');
//   const redirect = S_get_('XERO_REDIRECT_URI','');

//   const tok = UrlFetchApp.fetch('https://identity.xero.com/connect/token', {
//     method:'post', contentType:'application/x-www-form-urlencoded',
//     payload: { grant_type: 'authorization_code', code: code, redirect_uri: redirect },
//     headers: { Authorization: 'Basic ' + Utilities.base64Encode(clientId + ':' + clientSecret) },
//     muteHttpExceptions:true
//   });
//   const codeHttp = tok.getResponseCode();
//   if (codeHttp!==200) { SpreadsheetApp.getUi().alert('Token exchange failed '+codeHttp+' :: '+tok.getContentText().slice(0,400)); return; }

//   const body = JSON.parse(tok.getContentText());
//   if (body.refresh_token) PropertiesService.getScriptProperties().setProperty('XERO_REFRESH_TOKEN', body.refresh_token);

//   // Discover tenant ID
//   const con = UrlFetchApp.fetch('https://api.xero.com/connections', {headers:{Authorization:'Bearer '+body.access_token, Accept:'application/json'}});
//   if (con.getResponseCode()===200) {
//     const arr = JSON.parse(con.getContentText()) || [];
//     if (arr[0] && arr[0].tenantId) {
//       PropertiesService.getScriptProperties().setProperty('XERO_TENANT_ID', arr[0].tenantId);
//     }
//   }
//   SpreadsheetApp.getUi().alert('Xero connected. Refresh token stored and tenant discovered.');
// }

// // Access token refresh (verbatim behavior)
// function XL_getAccessToken_(){
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
//   var tenantId = S_get_('XERO_TENANT_ID','');
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
//  * 5) Xero Push (Inclusive Only) — your working A-path with wrappers by flags
//  * ==========================================================================*/

// function XL_testXero_() {
//   try {
//     XL_getAccessToken_();
//     SpreadsheetApp.getUi().alert('Xero token ok. Tenant: ' + (S_get_('XERO_TENANT_ID','(missing)')));
//   } catch(e){ SpreadsheetApp.getUi().alert('Xero error: ' + e.message); }
// }

// function XL_pushSelectedDraft_InclusiveOnly_Standalone() {
//   var ss = SpreadsheetApp.getActive();
//   var sh = ss.getSheetByName('OrdersInbox');
//   if (!sh) { SpreadsheetApp.getUi().alert('OrdersInbox not found.'); return; }

//   var headers = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0].map(String);
//   var col = function(n){ return headers.indexOf(n) + 1; };

//   var ranges = (sh.getActiveRangeList() ? sh.getActiveRangeList().getRanges() : [sh.getActiveRange()]).filter(function(x){return !!x;});
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

//         // 1) Build lines — discounts as VALUE (current behavior)
//         var lineBuild = _buildMonetaryLinesWithDiscounts_(order);

//         // 2) Notes line
//         var notes = _extractNotes_(order);
//         if (notes) {
//           var noteLine = _makeDescriptionOnlyLine_(notes);
//           if (noteLine) lineBuild.lines.push(noteLine);
//         }

//         // 3) Payload (Inclusive + order date = due date)
//         var payload = _buildInclusiveInvoicePayload_UsingOrderDate_(order, lineBuild.lines);

//         // POST
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

//         // 429 backoff + single retry
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

//         // 400-validation fallback: strip ItemCode only
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
//  * 6) Builders & helpers — exact behavior (value-discounts, notes, payload)
//  * ==========================================================================*/

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

//   // Order-level discount pro-rata path preserved as in your current code
//   var anyAdjusted = false;
//   for (var aidx = 0; aidx < out.length; aidx++) {
//     var pm = productMeta[aidx];
//     var pmExt = (pm && pm.extended !== undefined) ? pm.extended : undefined;
//     var lineUA = out[aidx] && out[aidx].UnitAmount;
//     if (pmExt !== undefined && lineUA !== undefined && lineUA !== pmExt) { anyAdjusted = true; break; }
//   }

//   var orderTotalDiscount = _toNumberSafe_(order && order.total_discounts);
//   if (!anyAdjusted && orderTotalDiscount > 0 && productMeta.length > 0) {
//     var base = 0;
//     for (var b = 0; b < productMeta.length; b++) {
//       var m = productMeta[b];
//       if (m && m.extended > 0) base += m.extended;
//     }
//     if (base > 0) {
//       for (var j = 0; j < productMeta.length; j++) {
//         var m2 = productMeta[j];
//         if (!m2 || m2.extended <= 0) continue;
//         var share = orderTotalDiscount * (m2.extended / base);
//         var netExt = m2.extended - share;
//         var qty2 = out[m2.index].Quantity || 1;
//         var netUnit = netExt / qty2;
//         out[m2.index].UnitAmount = Math.round(netUnit * 100) / 100;
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
//       for (var i=0; i<order.note_attributes.length; i++) {
//         var attr = order.note_attributes[i];
//         if (!attr) continue;
//         var k = (attr.name || attr.key || '').toString().trim();
//         var v = (attr.value || '').toString().trim();
//         if (k || v) parts.push(k ? (k + ': ' + v) : v);
//       }
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
//       ? [order.customer.first_name, order.customer.last_name].filter(function(v){return !!v;}).join(' ')
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
//  * 7) Optional wrappers (validation detect, strip ItemCode, idempotency, etc.)
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

// // Auto-add marker columns if enabled
// function _ensureMarkerColumns_(sh) {
//   try {
//     var headers = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0].map(String);
//     var need = [];
//     if (headers.indexOf(XP_LOCAL_MARKER_COLS.pushedAt)  < 0) need.push(XP_LOCAL_MARKER_COLS.pushedAt);
//     if (headers.indexOf(XP_LOCAL_MARKER_COLS.invoiceId) < 0) need.push(XP_LOCAL_MARKER_COLS.invoiceId);
//     if (!need.length) return;
//     var startCol = sh.getLastColumn() + 1;
//     sh.insertColumnsAfter(sh.getLastColumn(), need.length);
//     sh.getRange(1, startCol, 1, need.length).setValues([need]);
//   } catch(e){}
// }

// /* ============================================================================
//  * 8) (Optional) WebApp OAuth handler from Xero_OAuth_WebApp
//  *     - If you deploy this project as a Web App (Execute as: Me; Access: Anyone),
//  *       set XERO_REDIRECT_URI to that /exec URL. Then you can use XERO_StartOAuth()
//  *       to open the authorize link; upon redirect, doGet(e) will display success.
//  * ==========================================================================*/

// const XERO_TOKEN_URL = 'https://identity.xero.com/connect/token';
// const XERO_CONNS_URL = 'https://api.xero.com/connections';

// function XERO_StartOAuth() {
//   const state = Utilities.getUuid();
//   PropertiesService.getScriptProperties().setProperty('XERO_OAUTH_STATE', state);
//   const url = xeroBuildAuthUrl_(state);
//   SpreadsheetApp.getUi().showModalDialog(
//     HtmlService.createHtmlOutput(
//       '<p>Click to authorize Xero in a new tab:</p>'
//       + '<p><a href="' + url + '" target="_blank" rel="noopener">Authorize Xero</a></p>'
//       + '<p>After approval, you will be redirected to the Web App and see a success page.</p>'),
//     'Xero Authorization'
//   );
// }
// function xeroBuildAuthUrl_(state) {
//   const clientId = S_get_('XERO_CLIENT_ID','');
//   const redirectUri = S_get_('XERO_REDIRECT_URI','');
//   if (!clientId || !redirectUri) throw new Error('XERO_CLIENT_ID or XERO_REDIRECT_URI missing.');
//   const params = {
//     response_type: 'code',
//     client_id: clientId,
//     redirect_uri: redirectUri,
//     scope: [
//       'offline_access',
//       'accounting.transactions',
//       'accounting.contacts',
//       'accounting.settings',
//       'accounting.attachments'
//     ].join(' '),
//     state: state
//   };
//   const q = Object.keys(params).map(function(k){ return encodeURIComponent(k)+'='+encodeURIComponent(params[k]); }).join('&');
//   return 'https://login.xero.com/identity/connect/authorize?' + q;
// }
// function doGet(e) {
//   try {
//     const params = e && e.parameter ? e.parameter : {};
//     const code = params.code || '';
//     const state = params.state || '';
//     if (!code) return htmlPage_('Xero OAuth', 'Missing "code" in redirect. Did you approve the app?');

//     const expected = PropertiesService.getScriptProperties().getProperty('XERO_OAUTH_STATE') || '';
//     if (expected && state !== expected) {
//       return htmlPage_('Xero OAuth', 'Invalid state. Please restart authorization from the Sheet menu.');
//     }

//     const redirectUri = S_get_('XERO_REDIRECT_URI','');
//     const token = xeroExchangeCodeForTokens_(code, redirectUri);

//     if (token.refresh_token) {
//       PropertiesService.getScriptProperties().setProperty('XERO_REFRESH_TOKEN', token.refresh_token);
//     }

//     var tenantId = S_get_('XERO_TENANT_ID','');
//     if (!tenantId) {
//       const conns = xeroGetConnections_(token.access_token);
//       if (conns && conns.length) {
//         tenantId = conns[0].tenantId || conns[0].tenant_id || '';
//         if (tenantId) PropertiesService.getScriptProperties().setProperty('XERO_TENANT_ID', tenantId);
//       }
//     }

//     return htmlPage_(
//       'Xero OAuth',
//       '<h2>✅ Xero Connected</h2>'
//       + '<p>Refresh token saved. Tenant: <b>' + (tenantId || '(not set)') + '</b></p>'
//       + '<p>You can close this tab and return to the Sheet.</p>'
//     );
//   } catch (err) {
//     return htmlPage_('Xero OAuth', '❌ Error: ' + String(err && err.message || err));
//   }
// }
// function xeroExchangeCodeForTokens_(code, redirectUri) {
//   const a = S_xeroAuth_();
//   const body = { grant_type: 'authorization_code', code: code, redirect_uri: redirectUri };
//   const tokenResp = UrlFetchApp.fetch(XERO_TOKEN_URL, {
//     method: 'post',
//     contentType: 'application/x-www-form-urlencoded',
//     payload: Object.keys(body).map(function(k){ return encodeURIComponent(k)+'='+encodeURIComponent(body[k]); }).join('&'),
//     headers: { Authorization: 'Basic ' + Utilities.base64Encode(a.clientId + ':' + a.clientSecret) },
//     muteHttpExceptions: true
//   });
//   const codeResp = tokenResp.getResponseCode();
//   const text = tokenResp.getContentText();
//   if (codeResp < 200 || codeResp >= 300) throw new Error('Token exchange failed ' + codeResp + ': ' + text);
//   return JSON.parse(text);
// }
// function xeroGetConnections_(accessToken) {
//   const resp = UrlFetchApp.fetch(XERO_CONNS_URL, {
//     method: 'get',
//     headers: { Authorization: 'Bearer ' + accessToken },
//     muteHttpExceptions: true
//   });
//   const code = resp.getResponseCode();
//   const txt = resp.getContentText();
//   if (code < 200 || code >= 300) throw new Error('Connections failed ' + code + ': ' + txt);
//   return JSON.parse(txt);
// }
// function htmlPage_(title, bodyHtml) {
//   return HtmlService.createHtmlOutput(
//     '<html><head><meta charset="utf-8"><title>'+title+'</title></head>'
//     + '<body style="font-family:system-ui;padding:24px;line-height:1.5">'
//     + bodyHtml + '</body></html>'
//   ).setTitle(title);
// }
