/** ========================================================================
 * Resolver for Xeroitemslist (EXACT HEADERS SUPPORTED):
 *   *ItemCode, ItemName, Quantity, PurchasesDescription, PurchasesUnitPrice,
 *   PurchasesAccount, PurchasesTaxRate, SalesDescription, SalesUnitPrice,
 *   SalesAccount, SalesTaxRate, InventoryAssetAccount, CostOfGoodsSoldAccount,
 *   Status, InventoryType
 *
 * Used by BetaV3 push to fill: ItemCode, AccountCode (SalesAccount), TaxType (SalesTaxRate)
 * Fallback: a row with *ItemCode = 'DEFAULT' provides default SalesAccount/TaxRate.
 * Matching key: Shopify line_item.sku is matched to *ItemCode.
 * ======================================================================== */

var __XI_CACHE = null;

function _xi_load_() {
  if (__XI_CACHE) return __XI_CACHE;

  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Xeroitemslist');
  var out = { byItemCode: {}, def: {} };

  if (!sh) { __XI_CACHE = out; return out; }

  var lastR = sh.getLastRow(), lastC = sh.getLastColumn();
  if (lastR < 2 || lastC < 2) { __XI_CACHE = out; return out; }

  var data = sh.getRange(1, 1, lastR, lastC).getValues();
  var hdr  = data.shift().map(function(s){ return String(s||'').trim(); });

  // Find indices by exact header names (case-insensitive, tolerate leading *)
  function idxOf(label){
    var needle = String(label||'').toLowerCase();
    for (var i=0;i<hdr.length;i++){
      if (hdr[i] && hdr[i].toString().toLowerCase().replace(/^\*/, '') === needle.replace(/^\*/, '')) return i;
    }
    return -1;
  }

  var ix = {
    itemCode:    idxOf('*ItemCode'),
    itemName:    idxOf('ItemName'),
    salesAcct:   idxOf('SalesAccount'),
    salesTax:    idxOf('SalesTaxRate')
  };

  data.forEach(function(r){
    var key = ix.itemCode>=0 ? String(r[ix.itemCode]||'').trim() : '';
    if (!key) return;

    var rec = {
      ItemCode:    key,
      ItemName:    ix.itemName>=0  ? String(r[ix.itemName]||'').trim()  : '',
      AccountCode: ix.salesAcct>=0 ? String(r[ix.salesAcct]||'').trim() : '',
      TaxType:     ix.salesTax>=0  ? String(r[ix.salesTax]||'').trim()  : ''
    };

    if (key.toUpperCase() === 'DEFAULT') {
      out.def = rec;                            // store default (AccountCode, TaxType)
    } else {
      out.byItemCode[key] = rec;                // direct lookup by ItemCode
    }
  });

  __XI_CACHE = out;
  return out;
}

/** Resolve mapping for a Shopify line_item using *ItemCode == sku */
function _xi_resolveBySKU_(lineItem) {
  var cache = _xi_load_();
  var sku = (lineItem && lineItem.sku != null) ? String(lineItem.sku).trim() : '';
  if (sku && cache.byItemCode[sku]) return cache.byItemCode[sku];
  return cache.def || {}; // fallback to DEFAULT row if present
}
