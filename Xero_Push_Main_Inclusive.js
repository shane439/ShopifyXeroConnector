// /* ============================================================================
//  * Xero Push (Inclusive Only) — main entry & quick test
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
