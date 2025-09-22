// /** === Builders + helpers for Beta per-line downgrade === */

// // Tiny utility: does a global function exist?
// function _hasFn_(name){ try { return typeof this[name] === 'function'; } catch(e){ return false; } }
// // Normalize to array
// function _asArray_(x){ return Array.isArray(x) ? x : (x ? [x] : []); }

// // Wrap a line with metadata so we can target only failing lines later
// function _xp_wrapLine_(raw, meta) {
//   return {
//     __meta: {
//       lineId: (meta && meta.lineId) || Utilities.getUuid(),
//       canDowngrade: !!(meta && meta.canDowngrade),
//       reasons: (meta && meta.reasons) || []
//     },
//     // Xero-accepted fields
//     Description: raw.Description,
//     Quantity: raw.Quantity,
//     UnitAmount: raw.UnitAmount,
//     AccountCode: raw.AccountCode,
//     ItemCode: raw.ItemCode,
//     TaxType: raw.TaxType,
//     DiscountRate: raw.DiscountRate
//   };
// }

// // Strip meta → plain Xero LineItems
// function _xp_plainLines_(wrapped) {
//   return wrapped.map(function(w){
//     var out = { Description: w.Description, Quantity: w.Quantity, UnitAmount: w.UnitAmount };
//     if (w.AccountCode)        out.AccountCode  = w.AccountCode;
//     if (w.ItemCode)           out.ItemCode     = w.ItemCode;
//     if (w.TaxType)            out.TaxType      = w.TaxType;
//     if (w.DiscountRate != null) out.DiscountRate = w.DiscountRate;
//     return out;
//   });
// }

// // Downgrade only specific lineIds: strip ItemCode, set fallback AccountCode, keep tax if configured
// function _xp_downgradeSpecificLines_(wrappedLines, failingLineIds, reasonTag) {
//   reasonTag = reasonTag || (typeof XP_DOWNGRADE_TAG === 'string' ? XP_DOWNGRADE_TAG : '[DOWNGRADED]');
//   var fallbackAcct = (typeof XP_FALLBACK_REVENUE_ACCT === 'string' ? XP_FALLBACK_REVENUE_ACCT : '');
//   var keepTax = (typeof XP_KEEP_TAXCODE_ON_DOWNGRADE === 'boolean' ? XP_KEEP_TAXCODE_ON_DOWNGRADE : true);

//   var failing = {};
//   (failingLineIds || []).forEach(function(id){ failing[id] = true; });

//   return wrappedLines.map(function(w){
//     if (!failing[w.__meta.lineId] || !w.__meta.canDowngrade) return w;
//     var desc = w.Description || '';
//     if (reasonTag && desc.indexOf(reasonTag) !== 0) desc = reasonTag + ' ' + desc;

//     var d = {
//       __meta: w.__meta,
//       Description: desc,
//       Quantity: w.Quantity,
//       UnitAmount: w.UnitAmount
//     };
//     if (fallbackAcct) d.AccountCode = fallbackAcct;
//     if (keepTax && w.TaxType) d.TaxType = w.TaxType;
//     if (w.DiscountRate != null) d.DiscountRate = w.DiscountRate;
//     return d; // no ItemCode
//   });
// }

// // Parse Xero 400 response looking for "LineItems[<n>]" indices that failed
// function _xp_extractFailingLineIndexes_(bodyText) {
//   try {
//     var body = JSON.parse(bodyText || '{}');
//     var set = {};
//     (JSON.stringify(body).toLowerCase().match(/lineitems\[(\d+)\]/g) || []).forEach(function(tok){
//       var m = /lineitems\[(\d+)\]/i.exec(tok); if (m) set[parseInt(m[1],10)] = 1;
//     });
//     return Object.keys(set).map(function(k){ return parseInt(k,10); });
//   } catch(e) { return []; }
// }
