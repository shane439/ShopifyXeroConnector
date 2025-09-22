// /** === Xero Push Main — Beta V1 with targeted per-line downgrade === */

// // Entry point (wire this in your menu alongside the original)
// function XL_pushSelectedDraft_InclusiveOnly_Standalone() {
//   var ss = SpreadsheetApp.getActive();
//   var sh = ss.getSheetByName('OrdersInbox');
//   var sel = sh.getActiveRange();
//   if (!sel) { SpreadsheetApp.getUi().alert('Select one or more rows in OrdersInbox first.'); return; }

//   var values = sel.getValues();
//   for (var i = 0; i < values.length; i++) {
//     try {
//       var row = values[i];
//       // Adjust indexes if your sheet schema differs
//       var orderId   = row[0];
//       var orderNo   = row[1];
//       var createdAt = row[2]; // Date
//       var rawJson   = row[11]; // raw_json (stringified JSON)

//       if (!orderId) continue;

//       // Reference & contact (adjust to your existing builder if you have one)
//       var ref = "Shopify #" + orderNo;
//       var contact = _xp_buildContactFromRow_(row); // tiny helper below

//       var dateStr = (createdAt && createdAt.toISOString) ? createdAt.toISOString().slice(0,10)
//                     : Utilities.formatDate(new Date(), "UTC", "yyyy-MM-dd");

//       // Build wrapped lines from the raw_json (adjust to your actual schema if needed)
//       var shop = {};
//       try { shop = JSON.parse(rawJson || '{}'); } catch(e) {}
//       var lineItems = _asArray_(shop.line_items || []);

//       var wrappedLines = lineItems.map(function(li, idx){
//         var raw = _xp_buildRawLine_FromShopify_(li); // helper below; keep your own mapping if you already have one
//         var canDown = true; // eligible for downgrade if needed
//         var reasons = [];
//         if (!raw.ItemCode) reasons.push('NO_ITEMCODE');
//         if (!raw.AccountCode) reasons.push('NO_ACCOUNT'); // informational
//         return _xp_wrapLine_(raw, { lineId: String(idx+1), canDowngrade: canDown, reasons: reasons });
//       });

//       // Branch by flag
//       var res = XP_BETA1_ENABLE_PARTIAL_DOWNGRADE
//         ? _xp_createOrRepairDraftInclusive_(ref, contact, wrappedLines, dateStr)
//         : _xp_createDraftInclusive_Old_(ref, contact, wrappedLines, dateStr);

//       if (res.code !== 200 && res.code !== 201) {
//         throw new Error('Xero push failed: ' + res.code + ' ' + (res.body || ''));
//       }

//       // Stamp idempotency markers if you use them (PushedAt / PushedInvoiceID)
//       sh.getRange(sel.getRow() + i, 13).setValue(new Date()); // PushedAt
//       try {
//         var invId = (res.json && res.json.Invoices && res.json.Invoices[0] && res.json.Invoices[0].InvoiceID) || '';
//         sh.getRange(sel.getRow() + i, 14).setValue(invId);     // PushedInvoiceID
//       } catch(e){}

//     } catch (e) {
//       Logger.log('Row ' + (sel.getRow()+i) + ' push error: ' + e.message);
//       SpreadsheetApp.getUi().alert('Push error at row ' + (sel.getRow()+i) + ' → ' + e.message);
//     }
//   }
// }

// /** Old behavior: global downgrade (for parity when flag is OFF) */
// function _xp_createDraftInclusive_Old_(ref, contact, wrappedLines, dateStr) {
//   var lines = _xp_plainLines_(wrappedLines).map(function(l){
//     // global strip of ItemCode (your current “downgrade everything” behavior)
//     delete l.ItemCode;
//     return l;
//   });
//   var payload = _xp_buildInvoicePayload_(ref, contact, lines, dateStr);
//   return _xp_xeroPostInvoices_(payload);
// }

// /** New behavior: try full payload → on 400, downgrade only failing lines */
// function _xp_createOrRepairDraftInclusive_(ref, contact, wrappedLines, dateStr) {
//   var payload = _xp_buildInvoicePayload_(ref, contact, _xp_plainLines_(wrappedLines), dateStr);
//   var res = _xp_xeroPostInvoices_(payload);
//   if (res.code === 200 || res.code === 201) return res;

//   if (res.code !== 400) throw new Error('Xero error ' + res.code + ': ' + (res.body || ''));

//   // Identify failing line indexes from Xero error body, then map to our lineIds
//   var failingIdxs = _xp_extractFailingLineIndexes_(res.body || '');
//   if (!failingIdxs.length) {
//     // fallback to old behavior for this invoice only
//     var fallback = wrappedLines.map(function(w){ return _xp_downgradeSpecificLines_([w], [w.__meta.lineId])[0]; });
//     payload = _xp_buildInvoicePayload_(ref, contact, _xp_plainLines_(fallback), dateStr);
//     return _xp_xeroPostInvoices_(payload);
//   }

//   var failingIds = failingIdxs
//     .filter(function(i){ return wrappedLines[i] && wrappedLines[i].__meta; })
//     .map(function(i){ return wrappedLines[i].__meta.lineId; });

//   var repaired = _xp_downgradeSpecificLines_(wrappedLines, failingIds, '[NO-ITEMCODE]');
//   payload = _xp_buildInvoicePayload_(ref, contact, _xp_plainLines_(repaired), dateStr);
//   return _xp_xeroPostInvoices_(payload);
// }

// /** Build Xero invoice payload (Inclusive + Draft) */
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

// /** Thin POST wrapper (uses your existing token helpers) */
// function _xp_xeroPostInvoices_(payload) {
//   var url = 'https://api.xero.com/api.xro/2.0/Invoices';
//   var access = XL_getAccessToken_();     // keep your existing helper name
//   var tenant = XL_ensureTenantId_(access);
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
//   return { code: res.getResponseCode(), body: res.getContentText(), json: JSON.parse(res.getContentText() || '{}') };
// }

// /** Minimal contact builder from the OrdersInbox row (customize as needed) */
// function _xp_buildContactFromRow_(row) {
//   var name = row[5] || ''; // customer_name
//   var email = row[6] || ''; // customer_email
//   var out = {};
//   if (name) out.Name = name;
//   if (email) out.EmailAddress = email;
//   return out;
// }

// /** Minimal raw-line builder from Shopify line_item object (customize to your mapping) */
// function _xp_buildRawLine_FromShopify_(li) {
//   return {
//     Description: (li.title || ''),
//     Quantity: Number(li.quantity || 0),
//     UnitAmount: Number(li.price || 0),
//     AccountCode: li.account_code || '', // if you map GLs externally, resolve here
//     ItemCode: li.sku || '',
//     TaxType: li.tax_code || '',
//     DiscountRate: (li.total_discount && li.price) ? Math.round((100 * Number(li.total_discount||0))/Math.max(1,Number(li.price))) : null
//   };
// }
// // Alias: expose a Beta-named entry that calls your real Beta function
// function XL_pushSelectedDraft_InclusiveOnly_Standalone_BetaV1(){
//   return XL_pushSelectedDraft_InclusiveOnly_Standalone();
// }

