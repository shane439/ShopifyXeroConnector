/* ============================================================================
 * Builders & Helpers (value-discounts, notes, payload)
 * ==========================================================================*/

function _buildMonetaryLinesWithDiscounts_(order) {
  var out = [];
  var items = (order && Array.isArray(order.line_items)) ? order.line_items : [];

  var productMeta = [];
  for (var i = 0; i < items.length; i++) {
    var li = items[i] || {};
    var qty = Number(li.quantity || 0);
    if (!qty) continue;

    var unit = _toNumberSafe_(li.price, li.price_set && li.price_set.shop_money && li.price_set.shop_money.amount);
    var desc = (li.title || '').toString().substring(0, 4000);

    var line = { Description: desc, Quantity: qty, UnitAmount: unit, TaxType: 'OUTPUT' };
    if (li.sku) line.ItemCode = String(li.sku).trim();

    var liDisc = _sumDiscountAllocations_(li);
    if (liDisc > 0) {
      var ext = qty * unit;
      if (ext > 0) {
        var netExt = ext - liDisc;
        var netUnit = netExt / qty;
        line.UnitAmount = Math.round(netUnit * 100) / 100;
      }
    }

    out.push(line);
    productMeta.push({ index: out.length - 1, extended: qty * unit });
  }

  var shippingTotal = Array.isArray(order && order.shipping_lines)
    ? order.shipping_lines.reduce(function(s,x){ return s + _toNumberSafe_(x.price); }, 0)
    : 0;
  if (shippingTotal > 0) {
    out.push({ Description: 'Shipping', Quantity: 1, UnitAmount: Number(shippingTotal), TaxType: 'OUTPUT' });
  }

  // Order-level discount pro-rata path preserved as in your current code
  var anyAdjusted = false;
  for (var aidx = 0; aidx < out.length; aidx++) {
    var pm = productMeta[aidx];
    var pmExt = (pm && pm.extended !== undefined) ? pm.extended : undefined;
    var lineUA = out[aidx] && out[aidx].UnitAmount;
    if (pmExt !== undefined && lineUA !== undefined && lineUA !== pmExt) { anyAdjusted = true; break; }
  }

  var orderTotalDiscount = _toNumberSafe_(order && order.total_discounts);
  if (!anyAdjusted && orderTotalDiscount > 0 && productMeta.length > 0) {
    var base = 0;
    for (var b = 0; b < productMeta.length; b++) {
      var m = productMeta[b];
      if (m && m.extended > 0) base += m.extended;
    }
    if (base > 0) {
      for (var j = 0; j < productMeta.length; j++) {
        var m2 = productMeta[j];
        if (!m2 || m2.extended <= 0) continue;
        var share = orderTotalDiscount * (m2.extended / base);
        var netExt = m2.extended - share;
        var qty2 = out[m2.index].Quantity || 1;
        var netUnit = netExt / qty2;
        out[m2.index].UnitAmount = Math.round(netUnit * 100) / 100;
      }
    }
  }

  return { lines: out };
}

function _extractNotes_(order) {
  try {
    var parts = [];
    if (order && order.note && String(order.note).trim()) parts.push(String(order.note).trim());
    if (order && Array.isArray(order.note_attributes)) {
      for (var i=0; i<order.note_attributes.length; i++) {
        var attr = order.note_attributes[i];
        if (!attr) continue;
        var k = (attr.name || attr.key || '').toString().trim();
        var v = (attr.value || '').toString().trim();
        if (k || v) parts.push(k ? (k + ': ' + v) : v);
      }
    }
    var all = parts.join('\n').trim();
    return all ? all : '';
  } catch (e) { return ''; }
}

function _makeDescriptionOnlyLine_(text) {
  var t = (text || '').toString();
  return t ? { Description: t.substring(0, 4000), Quantity: 0, UnitAmount: 0, TaxType: 'NONE', AccountCode: null } : null;
}

function _buildInclusiveInvoicePayload_UsingOrderDate_(order, lineItems) {
  var name =
    (order && order.customer
      ? [order.customer.first_name, order.customer.last_name].filter(function(v){return !!v;}).join(' ')
      : (order && order.billing_address && order.billing_address.name)) ||
    'Shopify Customer';
  var email = (order && (order.email || (order.customer && order.customer.email))) || undefined;

  var created = new Date(order && order.created_at ? order.created_at : new Date());
  var xDate = created.toISOString().slice(0,10);

  var invoice = {
    Type: 'ACCREC',
    Status: 'DRAFT',
    Date: xDate,
    DueDate: xDate,
    Reference: 'Shopify #' + String(order && order.order_number || ''),
    Contact: (email ? { Name: name, EmailAddress: email } : { Name: name }),
    LineItems: lineItems,
    LineAmountTypes: 'Inclusive'
  };

  if (order && order.currency) invoice.CurrencyCode = String(order.currency);
  return { Invoices: [ invoice ] };
}

function _sumDiscountAllocations_(li) {
  try {
    var arr = Array.isArray(li && li.discount_allocations) ? li.discount_allocations : [];
    var sum = 0;
    for (var i = 0; i < arr.length; i++) {
      var a = arr[i] || {};
      sum += _toNumberSafe_(a.amount, a.amount_set && a.amount_set.shop_money && a.amount_set.shop_money.amount);
    }
    return Number(sum) || 0;
  } catch (e) { return 0; }
}
function _toNumberSafe_() {
  for (var i = 0; i < arguments.length; i++) {
    var v = arguments[i];
    if (v === null || v === undefined) continue;
    var n = Number(v);
    if (!isNaN(n)) return n;
    var s = String(v).trim();
    if (s) {
      var n2 = Number(s);
      if (!isNaN(n2)) return n2;
    }
  }
  return 0;
}

function _forceInclusiveHeaderOnly_(accessToken, tenantId, inv) {
  var contact = {};
  if (inv.Contact && inv.Contact.ContactID) contact = { ContactID: inv.Contact.ContactID };
  else if (inv.Contact && inv.Contact.Name) contact = { Name: inv.Contact.Name };
  else contact = { Name: 'Shopify Customer' };

  var fixPayload = { Invoices: [{ InvoiceID: inv.InvoiceID, Type: 'ACCREC', Status: 'DRAFT', Contact: contact, LineAmountTypes: 'Inclusive' }] };
  var resPut = UrlFetchApp.fetch('https://api.xero.com/api.xro/2.0/Invoices', {
    method: 'put',
    headers: { Authorization: 'Bearer ' + accessToken, 'xero-tenant-id': tenantId, Accept: 'application/json', 'Content-Type': 'application/json' },
    payload: JSON.stringify(fixPayload),
    muteHttpExceptions: true
  });
  if (resPut.getResponseCode() >= 300) {
    throw new Error('Inclusive header-only fix-up failed ' + resPut.getResponseCode() + ' :: ' + resPut.getContentText().slice(0, 1200));
  }
}
