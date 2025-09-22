// var __XP_GL_CACHE = null;

// function _xp_glMap_load_(){
//   if (__XP_GL_CACHE) return __XP_GL_CACHE;
//   var ss = SpreadsheetApp.getActive();
//   var sh = ss.getSheetByName('GL_Map');
//   var map = { byKey: {}, defaultAC: '', defaultTax: '' };
//   if (sh) {
//     var last = sh.getLastRow(), lastC = sh.getLastColumn();
//     if (last >= 2 && lastC >= 2) {
//       var data = sh.getRange(1,1,last,lastC).getValues();
//       var hdr = data.shift().map(function(s){return String(s||'').toLowerCase().trim();});
//       var keyIx  = hdr.indexOf('key');
//       var acIx   = hdr.indexOf('account_code');
//       var taxIx  = hdr.indexOf('tax_type');
//       data.forEach(function(r){
//         var k = keyIx>=0? String(r[keyIx]||'').trim() : '';
//         var ac = acIx>=0? String(r[acIx]||'').trim() : '';
//         var tx = taxIx>=0? String(r[taxIx]||'').trim() : '';
//         if (!k) return;
//         if (k.toUpperCase() === 'DEFAULT') { map.defaultAC = ac; map.defaultTax = tx; return; }
//         map.byKey[k] = { AccountCode: ac, TaxType: tx };
//       });
//     }
//   }
//   __XP_GL_CACHE = map;
//   return map;
// }

// function _xp_resolveGL_(li){
//   var map = _xp_glMap_load_();
//   var sku   = (li && li.sku != null) ? String(li.sku).trim() : '';
//   var ptype = (li && li.product_type != null) ? String(li.product_type).trim() : '';
//   if (sku && map.byKey[sku])   return map.byKey[sku];
//   if (ptype && map.byKey[ptype]) return map.byKey[ptype];
//   var out = {};
//   if (map.defaultAC)  out.AccountCode = map.defaultAC;
//   if (map.defaultTax) out.TaxType     = map.defaultTax;
//   return out;
// }
