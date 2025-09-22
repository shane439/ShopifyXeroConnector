/** ======================================================================
 * Xero Wrappers + Idempotency (FULL REPLACEMENT)
 * Public entry: XP_postInvoiceDraftInclusive_(payload)
 * Behaviors:
 *  - Never posts if LineItems empty
 *  - 429 backoff
 *  - 400 validation → layered line repair:
 *      a) Targeted by parsed indexes (ItemCode → TaxType → AccountCode)
 *      b) If no indexes parsed, progressive one-line probe
 *      c) Last-resort global downgrade
 *  - Rich Diag_Trace logs at every point
 * ====================================================================== */

/* Flags (safe defaults if not defined elsewhere) */
if (typeof XP_ENABLE_TARGETED_LINE_DOWNGRADE === 'undefined') var XP_ENABLE_TARGETED_LINE_DOWNGRADE = true; // recommend ON
if (typeof XP_DOWNGRADE_TAG === 'undefined')                 var XP_DOWNGRADE_TAG = '[DOWNGRADED]';
if (typeof XP_FALLBACK_REVENUE_ACCT === 'undefined')         var XP_FALLBACK_REVENUE_ACCT = '';   // keep '' to preserve mapped AC unless forced
if (typeof XP_KEEP_TAXCODE_ON_DOWNGRADE === 'undefined')     var XP_KEEP_TAXCODE_ON_DOWNGRADE = true;

/** Public entry */
function XP_postInvoiceDraftInclusive_(payload) {
  var access = XL_getAccessToken_();
  var tenant = XL_ensureTenantId_(access);
  var ref = DX_refFromPayload_(payload);

  DX_log_('WRAPPER', ref, 'POST:begin', 'info', 'Attempting first POST', {snap: DX_compactInvoice_(payload)});

  if (!_xp_hasAnyLineItems_(payload)) {
    DX_log_('WRAPPER', ref, 'guard:no_lines', 'skip', 'Payload has no LineItems');
    return { code: 0, body: 'Skipped POST: no LineItems', json: {} };
  }

  var url = 'https://api.xero.com/api.xro/2.0/Invoices';
  var baseOpts = {
    method: 'post',
    muteHttpExceptions: true,
    headers: {
      'Authorization': 'Bearer ' + access,
      'Xero-tenant-id': tenant,
      'Accept': 'application/json',
      'Content-Type': 'application/json'
    }
  };

  // 1) First attempt
  var opts1 = Object.assign({}, baseOpts, { payload: JSON.stringify(payload) });
  var r1 = _xp_fetchWithBackoff_(url, opts1, 3);
  DX_log_('WRAPPER', ref, 'POST:first', String(r1.code), 'response', _dx_respMeta_(r1));

  if (r1.code === 200 || r1.code === 201) return r1;
  if (r1.code !== 400) return r1; // non-validation: bail with full log

  // Parse 400 details
  var vinfo = _xp_parse400_(r1.body);
  DX_log_('WRAPPER', ref, '400:parsed', 'info', vinfo.summary, vinfo);

  // 2) Layered repair if flag ON
  var inv = _getInv_(payload), lines = (inv.LineItems || []);

  if (XP_ENABLE_TARGETED_LINE_DOWNGRADE) {
    var targetIdxs = vinfo.indexes.length ? vinfo.indexes : []; // when element indexes available
    var pass = 0;

    // Pass A: itemcode-only (remove ItemCode on offenders)
    pass++;
    var resA = _xp_tryRepair_(url, baseOpts, inv, lines, targetIdxs, { dropItemCode: true });
    if (_ok_(resA)) return _withMeta_(resA, targetIdxs.length? targetIdxs : resA.meta && resA.meta.downgradedIdxs || []);
    DX_log_('WRAPPER', ref, '400:passA:itemcode', String(resA.code), 'after itemcode removal', _dx_respMeta_(resA));

    // Pass B: if message hints tax problem, or still 400 → also drop/normalize TaxType on offenders
    pass++;
    var resB = _xp_tryRepair_(url, baseOpts, inv, lines, targetIdxs, { dropItemCode: true, dropTaxType: true });
    if (_ok_(resB)) return _withMeta_(resB, targetIdxs.length? targetIdxs : resB.meta && resB.meta.downgradedIdxs || []);
    DX_log_('WRAPPER', ref, '400:passB:tax', String(resB.code), 'after drop tax', _dx_respMeta_(resB));

    // Pass C: if AccountCode invalid (or still failing), set fallback AccountCode on offenders
    pass++;
    var resC = _xp_tryRepair_(url, baseOpts, inv, lines, targetIdxs, { dropItemCode: true, dropTaxType: true, forceFallbackAC: true });
    if (_ok_(resC)) return _withMeta_(resC, targetIdxs.length? targetIdxs : resC.meta && resC.meta.downgradedIdxs || []);
    DX_log_('WRAPPER', ref, '400:passC:acct', String(resC.code), 'after force fallback AC', _dx_respMeta_(resC));

    // If we had no explicit indexes from Xero, try progressive single-line probe
    if (!targetIdxs.length) {
      DX_log_('WRAPPER', ref, '400:probe', 'info', 'Try progressive single-line repair');
      var resP = _xp_progressiveProbe_(url, baseOpts, inv, lines);
      if (_ok_(resP)) return resP;
      DX_log_('WRAPPER', ref, '400:probe:done', String(resP.code), 'progressive failed', _dx_respMeta_(resP));
    }
  } else {
    DX_log_('WRAPPER', ref, '400:targeted', 'off', 'Flag disabled');
  }

  // 3) Global downgrade (last resort)
  DX_log_('WRAPPER', ref, '400:global', 'info', 'Attempt global downgrade of all lines');
  var allDown = lines.map(function(li){ return _downgrade(li, {dropItemCode:true, dropTaxType:false, forceFallbackAC:false}); });
  var rG = _xp_post_(url, baseOpts, Object.assign({}, inv, { LineItems: allDown }));
  DX_log_('WRAPPER', ref, '400:global:post', String(rG.code), 'after global', _dx_respMeta_(rG));
  return _withMeta_(rG, []); // global: leave meta empty

  // ---- inner helpers ----
  function _ok_(res){ return res && (res.code === 200 || res.code === 201); }
  function _withMeta_(res, idxs){ res = res || {}; res.meta = res.meta || {}; res.meta.downgradedIdxs = idxs || []; return res; }
}

/** 429 backoff */
function _xp_fetchWithBackoff_(url, opts, maxRetries) {
  var attempt=0, wait=1000;
  while (true) {
    var r = UrlFetchApp.fetch(url, opts);
    var out = { code: r.getResponseCode(), body: r.getContentText(), json: _safeParseJSON_(r.getContentText(), {}), headers: _safeHeaders_(r) };
    if (out.code !== 429) return out;
    if (attempt >= (maxRetries||0)) return out;
    Utilities.sleep(wait); wait = Math.min(wait*2, 8000); attempt++;
  }
}

/** Low-level POST with same options; returns {code, body, json, headers} */
function _xp_post_(url, baseOpts, invoiceObj) {
  var payload = { Invoices: [ invoiceObj ] };
  if (!_hasLines_({Invoices:[invoiceObj]})) return { code: 0, body: 'Skipped POST: zero lines', json: {} };
  var r = UrlFetchApp.fetch(url, Object.assign({}, baseOpts, { payload: JSON.stringify(payload) }));
  return { code: r.getResponseCode(), body: r.getContentText(), json: _safeParseJSON_(r.getContentText(), {}), headers: _safeHeaders_(r) };
}

/** Try repair with specific toggles on offenders */
function _xp_tryRepair_(url, baseOpts, inv, lines, idxs, cfg) {
  var target = (idxs && idxs.length) ? idxs : lines.map(function(_,i){return i;}); // all lines if no explicit indexes
  var repaired = lines.map(function(li, i){ return (target.indexOf(i)>=0) ? _downgrade(li, cfg) : li; });
  var res = _xp_post_(url, baseOpts, Object.assign({}, inv, { LineItems: repaired }));
  if (res && (res.code === 200 || res.code === 201)) {
    res.meta = { downgradedIdxs: target.slice() };
  }
  return res;
}

/** Progressive one-by-one: itemcode→tax→acct per index until success */
function _xp_progressiveProbe_(url, baseOpts, inv, lines) {
  for (var i=0;i<lines.length;i++) {
    // Pass 1: drop itemcode only on i
    var l1 = lines.slice(); l1[i] = _downgrade(l1[i], {dropItemCode:true});
    var r1 = _xp_post_(url, baseOpts, Object.assign({}, inv, { LineItems: l1 }));
    if (r1.code===200 || r1.code===201) { r1.meta = { downgradedIdxs:[i] }; return r1; }

    // Pass 2: drop itemcode & tax on i
    var l2 = lines.slice(); l2[i] = _downgrade(l2[i], {dropItemCode:true, dropTaxType:true});
    var r2 = _xp_post_(url, baseOpts, Object.assign({}, inv, { LineItems: l2 }));
    if (r2.code===200 || r2.code===201) { r2.meta = { downgradedIdxs:[i] }; return r2; }

    // Pass 3: drop itemcode & tax + fallback AC on i
    var l3 = lines.slice(); l3[i] = _downgrade(l3[i], {dropItemCode:true, dropTaxType:true, forceFallbackAC:true});
    var r3 = _xp_post_(url, baseOpts, Object.assign({}, inv, { LineItems: l3 }));
    if (r3.code===200 || r3.code===201) { r3.meta = { downgradedIdxs:[i] }; return r3; }
  }
  // If none worked, return last attempt
  return r3 || r2 || r1 || { code: 400, body: 'Progressive probe failed', json: {} };
}

/** Parse Xero 400: extract messages + element indexes */
function _xp_parse400_(bodyText) {
  var out = { indexes: [], messages: [], summary: '' };
  try {
    var json = _safeParseJSON_(bodyText, {});
    var elements = json.Elements || [];
    elements.forEach(function(el, eidx){
      (el.ValidationErrors || []).forEach(function(ve){
        var msg = String(ve.Message || '');
        out.messages.push(msg);
        var m = msg.match(/LineItems\[(\d+)\]/i);
        if (m) out.indexes.push(parseInt(m[1],10));
      });
    });
    // fallback: regex scan whole body
    if (!out.indexes.length) {
      (String(bodyText).toLowerCase().match(/lineitems\[(\d+)\]/g) || []).forEach(function(tok){
        var m = /lineitems\[(\d+)\]/i.exec(tok);
        if (m) out.indexes.push(parseInt(m[1],10));
      });
    }
    out.indexes = Array.from(new Set(out.indexes));
    out.summary = (out.messages.slice(0,3).join(' | ') || '400 validation');
  } catch(e) {}
  return out;
}

/** Downgrade one line with options */
function _downgrade(li, cfg) {
  cfg = cfg || {};
  var tag = (typeof XP_DOWNGRADE_TAG === 'string' ? XP_DOWNGRADE_TAG : '[DOWNGRADED]');
  var keepTax = (typeof XP_KEEP_TAXCODE_ON_DOWNGRADE === 'boolean' ? XP_KEEP_TAXCODE_ON_DOWNGRADE : true);
  var d = JSON.parse(JSON.stringify(li || {}));
  // ensure description tag
  if (tag) {
    var desc = d.Description || '';
    if (desc.indexOf(tag) !== 0) d.Description = tag + ' ' + desc;
  }
  if (cfg.dropItemCode) delete d.ItemCode;
  if (cfg.dropTaxType || !keepTax) delete d.TaxType;
  if (cfg.forceFallbackAC && XP_FALLBACK_REVENUE_ACCT) d.AccountCode = XP_FALLBACK_REVENUE_ACCT;
  return d;
}

/** Helpers */
function _getInv_(payload){ return (payload && payload.Invoices && payload.Invoices[0]) || {}; }
function _hasLines_(payload){ try{ var a=payload.Invoices[0].LineItems; return Array.isArray(a)&&a.length>0; }catch(e){return false;} }
function _xp_hasAnyLineItems_(payload){ return _hasLines_(payload); }
function _safeParseJSON_(s, fb){ try{ return s? JSON.parse(s): fb; }catch(e){ return fb; } }
function _safeHeaders_(resp){ try{ return resp.getAllHeaders ? resp.getAllHeaders() : {}; }catch(e){ return {}; } }
function _dx_respMeta_(res){ return { code: res.code, hdr: (res.headers||{}), body: (res.body||'').slice(0,400) }; }
