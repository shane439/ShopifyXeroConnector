/** ======================================================================
 * Xero Wrappers + Idempotency (BetaV3)
 * Deltas:
 *  - Greedy incremental LAST-RESORT with heuristic ordering when Xero
 *    doesn't return LineItems[n] indexes (minimize fewest possible lines).
 *  - If Xero DOES return indexes, we still minimize only those.
 *  - No "[MINIMAL]" prefixes.
 * ====================================================================== */

if (typeof XP_STRATEGY_ORDER === 'undefined') var XP_STRATEGY_ORDER = 'A';
if (typeof XP_ENABLE_TARGETED_LINE_DOWNGRADE === 'undefined') var XP_ENABLE_TARGETED_LINE_DOWNGRADE = true;
if (typeof XP_ENABLE_PROGRESSIVE_PROBE === 'undefined') var XP_ENABLE_PROGRESSIVE_PROBE = true;
if (typeof XP_ENABLE_GLOBAL_DOWNGRADE === 'undefined') var XP_ENABLE_GLOBAL_DOWNGRADE = true;

if (typeof XP_ANNOTATE_DOWNGRADE_IN_DESCRIPTION === 'undefined') var XP_ANNOTATE_DOWNGRADE_IN_DESCRIPTION = false;
if (typeof XP_DOWNGRADE_TAG === 'undefined') var XP_DOWNGRADE_TAG = '[DOWNGRADED]';

if (typeof XP_FALLBACK_REVENUE_ACCT === 'undefined') var XP_FALLBACK_REVENUE_ACCT = '';
if (typeof XP_KEEP_TAXCODE_ON_DOWNGRADE === 'undefined') var XP_KEEP_TAXCODE_ON_DOWNGRADE = true;

function XP_postInvoiceDraftInclusive_(payload) {
  var ref = DX_refFromPayload_(payload);
  DX_log_('WRAPPER', ref, 'POST:begin', 'info', 'First POST', { snap: DX_compactInvoice_(payload), strategy: XP_STRATEGY_ORDER });
  DX_startTimer_('post-first');

  if (!_hasLines_({ Invoices: [ (payload && payload.Invoices && payload.Invoices[0]) || {} ] })) {
    DX_log_('WRAPPER', ref, 'guard:no_lines', 'skip', 'No LineItems');
    return _ensureMeta_({ code: 0, body: 'Skipped POST: no LineItems', json: {}, headers: {} }, { strategy: XP_STRATEGY_ORDER });
  }

  var access = XL_getAccessToken_();
  var tenant = XL_ensureTenantId_(access);
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

  var r1 = _xp_fetchWithBackoff_(url, Object.assign({}, baseOpts, { payload: JSON.stringify(payload) }), 3);
  DX_log_('WRAPPER', ref, 'POST:first', String(r1.code), 'response', { __stopTimerLabel:'post-first', meta:_dx_respMeta_(r1) });
  if (_ok_(r1)) return _ensureMeta_(r1, { strategy: XP_STRATEGY_ORDER });
  if (r1.code !== 400) return _ensureMeta_(r1, { strategy: XP_STRATEGY_ORDER });

  var inv   = _getInv_(payload), lines = (inv.LineItems || []);
  var parsed = _xp_parse400_(r1.body);
  DX_log_('WRAPPER', ref, '400:parsed', 'info', parsed.summary, { idxs: parsed.indexes, msgs: parsed.messages.slice(0,6) });

  var tried = { A:[], B:[], C:[], PROBE:[] };

  function runTargeted() {
    if (!XP_ENABLE_TARGETED_LINE_DOWNGRADE) return null;
    var targetIdxs = parsed.indexes.length ? parsed.indexes : lines.map(function(_,i){return i;});
    tried.A = targetIdxs.slice();
    var resA = _xp_tryRepair_(url, baseOpts, inv, lines, targetIdxs, { dropItemCode: true }, 'passA.dropItemCode');
    if (_ok_(resA)) return _withMetaTried_(resA, _actualIdxs_(parsed.indexes, resA), tried, 'A');

    tried.B = targetIdxs.slice();
    var resB = _xp_tryRepair_(url, baseOpts, inv, lines, targetIdxs, { dropItemCode: true, dropTaxType: true }, 'passB.dropTaxType');
    if (_ok_(resB)) return _withMetaTried_(resB, _actualIdxs_(parsed.indexes, resB), tried, 'B');

    tried.C = targetIdxs.slice();
    var resC = _xp_tryRepair_(url, baseOpts, inv, lines, targetIdxs, { dropItemCode: true, dropTaxType: true, forceFallbackAC: true }, 'passC.forceFallbackAC');
    if (_ok_(resC)) return _withMetaTried_(resC, _actualIdxs_(parsed.indexes, resC), tried, 'C');

    return null;
  }

  function runProgressive() {
    if (!XP_ENABLE_PROGRESSIVE_PROBE) return null;
    var resP = _xp_progressiveProbe_(url, baseOpts, inv, lines);
    tried.PROBE = (resP.meta && resP.meta.downgradedIdxs) ? resP.meta.downgradedIdxs.slice() : [];
    if (_ok_(resP)) return _withMetaTried_(resP, resP.meta.downgradedIdxs || [], tried, 'PROBE');
    return resP;
  }

  var res = (XP_STRATEGY_ORDER === 'B') ? (runProgressive() || runTargeted())
                                        : (runTargeted() || runProgressive());
  if (_ok_(res)) return _ensureMeta_(res, { parsed400: parsed, tried: tried, strategy: XP_STRATEGY_ORDER });

  // === LAST RESORT ===
  if (parsed.indexes.length) {
    DX_log_('WRAPPER', ref, '400:lastResort:indexed', 'warn', 'Minimize parsed indexes', { idxs: parsed.indexes });
    var minimized = lines.map(function(li, i){
      return (parsed.indexes.indexOf(i) === -1) ? li : _minimalizeLine_(li);
    });
    var rIndexed = _xp_post_(url, baseOpts, Object.assign({}, inv, { LineItems: minimized }));
    if (_ok_(rIndexed)) { rIndexed.meta.downgradedIdxs = parsed.indexes.slice(); rIndexed.meta.downgradeReason = 'lastResort'; }
    return _ensureMeta_(rIndexed, { parsed400: parsed, tried: tried, strategy: XP_STRATEGY_ORDER });
  }

  // No indexes → greedy incremental with heuristics
  DX_log_('WRAPPER', ref, '400:lastResort:incremental', 'warn', 'No indexes; running greedy with heuristic ordering');
  var inc = _xp_incrementalMinimizeHeuristic_(url, baseOpts, inv, lines);
  DX_log_('WRAPPER', ref, '400:lastResort:incremental:result', String(inc.code),
    { minimizedIdxs: inc.meta && inc.meta.downgradedIdxs || [], body: (inc.body||'').slice(0,240) });
  return _ensureMeta_(inc, { parsed400: parsed, tried: tried, strategy: XP_STRATEGY_ORDER });
}

/** Heuristic score: higher = more suspicious */
function _suspectScore_(li) {
  var s = 0;
  if (!li) return 0;
  if (!li.ItemCode) s += 3;
  if (!li.AccountCode) s += 2;
  if (!li.TaxType) s += 2;
  var t = (li.TaxType || '').toLowerCase();
  if (t && (t === 'sales tax' || /sales\s*tax(?!.*\d)/.test(t))) s += 1; // vague tax type
  return s;
}

/** Greedy incremental minimizer with heuristic ordering */
/**
 * Incremental minimization:
 *   Pass 1: try ONE line at a time (TAX → AC → RAW). Stop on first success.
 *   Pass 2: if still failing, grow a set (heuristic order) and for the set try TAX → AC → RAW.
 * Notes are never candidates. AC stage: set fallback account if configured;
 * otherwise drop AccountCode (keep ItemCode) + drop TaxType.
 */
function _xp_incrementalMinimizeHeuristic_(url, baseOpts, inv, lines) {
  var parsed = (DX_getCtx_ && DX_getCtx_().Parsed400) || { accountHints:[], taxHints:[] }; // optional context
  var ordered = _candidateOrder_(lines, parsed);
  if (!ordered.length) {
    return _ensureMeta_({ code: 400, body: 'no candidates to minimize', json: {}, headers: {} });
  }

  function postWithSet(setIdxs, mode, metaTag) {
    var version = lines.map(function(li, idx){
      if (setIdxs.indexOf(idx) === -1) return li;
      if (mode === 'TAX') {
        var d1 = JSON.parse(JSON.stringify(li || {}));
        delete d1.TaxType;
        return d1;
      }
      if (mode === 'AC') {
        var d2 = JSON.parse(JSON.stringify(li || {}));
        delete d2.TaxType;
        if (XP_FALLBACK_REVENUE_ACCT && String(XP_FALLBACK_REVENUE_ACCT).trim()) {
          d2.AccountCode = XP_FALLBACK_REVENUE_ACCT;
        } else {
          delete d2.AccountCode; // let Item / Xero default resolve
        }
        return d2;
      }
      return _minimalizeLine_(li);
    });

    var r = _xp_post_(url, baseOpts, Object.assign({}, inv, { LineItems: version }));
    if (_ok_(r)) {
      r.meta.downgradedIdxs = setIdxs.slice();
      r.meta.downgradeReason = 'lastResort.incremental:' + metaTag;
    }
    return r;
  }

  // ===== PASS 1: single-line attempts (cheapest) =====
  for (var a = 0; a < ordered.length; a++) {
    var idx = ordered[a];

    var r1 = postWithSet([idx], 'TAX', 'TAX');
    if (_ok_(r1)) return r1;

    var r2 = postWithSet([idx], 'AC', 'AC');
    if (_ok_(r2)) return r2;

    var r3 = postWithSet([idx], 'RAW', 'RAW');
    if (_ok_(r3)) return r3;
  }

  // ===== PASS 2: grow a set, retry TAX → AC → RAW for the set =====
  var minimized = [];
  var last = null;

  for (var k = 0; k < ordered.length; k++) {
    minimized.push(ordered[k]);

    last = postWithSet(minimized, 'TAX', 'TAX');
    if (_ok_(last)) return last;

    last = postWithSet(minimized, 'AC', 'AC');
    if (_ok_(last)) return last;

    last = postWithSet(minimized, 'RAW', 'RAW');
    if (_ok_(last)) return last;
  }

  return last || _ensureMeta_({ code: 400, body: 'incrementalMinimize failed', json: {}, headers: {} });
}

/** Keep description/qty/unit; drop ItemCode/AccountCode/TaxType */
function _minimalizeLine_(li) {
  li = li || {};
  return {
    Description: (li.Description || ''),
    Quantity: (li.Quantity != null ? li.Quantity : 0),
    UnitAmount: (li.UnitAmount != null ? li.UnitAmount : 0)
  };
}

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
function _xp_post_(url, baseOpts, invoiceObj) {
  DX_startTimer_('post-one');
  var payload = { Invoices: [ invoiceObj ] };
  if (!_hasLines_(payload)) return _ensureMeta_({ code: 0, body: 'Skipped POST: zero lines', json: {}, headers: {} });
  var r = UrlFetchApp.fetch(url, Object.assign({}, baseOpts, { payload: JSON.stringify(payload) }));
  var res = { code: r.getResponseCode(), body: r.getContentText(), json: _safeParseJSON_(r.getContentText(), {}), headers: _safeHeaders_(r) };
  DX_log_('WRAPPER', DX_refFromPayload_({Invoices:[invoiceObj]}), 'post:one', String(res.code), 'single post', { __stopTimerLabel:'post-one', meta:_dx_respMeta_(res) });
  return _ensureMeta_(res);
}

function _xp_tryRepair_(url, baseOpts, inv, lines, idxs, cfg, reasonTag) {
  var target = (idxs && idxs.length) ? idxs.slice() : lines.map(function(_,i){return i;});
  var repaired = lines.map(function(li, i){ return (target.indexOf(i)>=0) ? _maybeAnnotate_(_downgrade(li, cfg), reasonTag) : li; });
  var res = _xp_post_(url, baseOpts, Object.assign({}, inv, { LineItems: repaired }));
  if (_ok_(res)) { res.meta.downgradedIdxs = target.slice(); res.meta.downgradeReason = reasonTag || ''; }
  return res;
}

function _xp_progressiveProbe_(url, baseOpts, inv, lines) {
  var last = null;
  for (var i=0;i<lines.length;i++) {
    var l1 = lines.slice(); l1[i] = _maybeAnnotate_(_downgrade(l1[i], {dropItemCode:true}), 'probe.dropItemCode');
    var r1 = _xp_post_(url, baseOpts, Object.assign({}, inv, { LineItems: l1 })); last = r1; if (_ok_(r1)){ r1.meta.downgradedIdxs=[i]; r1.meta.downgradeReason='probe.dropItemCode'; return r1; }

    var l2 = lines.slice(); l2[i] = _maybeAnnotate_(_downgrade(l2[i], {dropItemCode:true, dropTaxType:true}), 'probe.dropTaxType');
    var r2 = _xp_post_(url, baseOpts, Object.assign({}, inv, { LineItems: l2 })); last = r2; if (_ok_(r2)){ r2.meta.downgradedIdxs=[i]; r2.meta.downgradeReason='probe.dropTaxType'; return r2; }

    var l3 = lines.slice(); l3[i] = _maybeAnnotate_(_downgrade(l3[i], {dropItemCode:true, dropTaxType:true, forceFallbackAC:true}), 'probe.forceFallbackAC');
    var r3 = _xp_post_(url, baseOpts, Object.assign({}, inv, { LineItems: l3 })); last = r3; if (_ok_(r3)){ r3.meta.downgradedIdxs=[i]; r3.meta.downgradeReason='probe.forceFallbackAC'; return r3; }
  }
  return last || _ensureMeta_({ code: 400, body: 'Progressive probe failed', json: {}, headers: {} });
}

function _xp_parse400_(bodyText) {
  var out = { indexes: [], messages: [], summary: '', accountHints: [], taxHints: [] };
  try {
    var body = String(bodyText || '');
    var json = _safeParseJSON_(body, null);

    if (json && json.Elements) {
      (json.Elements || []).forEach(function(el){
        (el.ValidationErrors || []).forEach(function(ve){
          var msg = String(ve.Message || '');
          if (msg) out.messages.push(msg);
        });
      });
    }
    if (json && json.Message && out.messages.length === 0) out.messages.push(String(json.Message));
    if (json && json.ErrorNumber && out.messages.length === 0) out.messages.push('ErrorNumber ' + json.ErrorNumber);

    var hay = (out.messages.join(' || ') + ' ' + body);

    // 1) Try capture explicit LineItems[n]
    var m, re = /LineItems\[(\d+)\]/gi;
    while ((m = re.exec(hay)) !== null) out.indexes.push(parseInt(m[1],10));
    out.indexes = Array.from(new Set(out.indexes));

    // 2) Capture account code hints → accountHints: ["4210", "4230"]
    var acc, reAcc = /account\s*code\s*'(\d+)'/gi;
    while ((acc = reAcc.exec(hay)) !== null) out.accountHints.push(String(acc[1]));
    out.accountHints = Array.from(new Set(out.accountHints));

    // 3) Capture tax type names → taxHints: ["Sales Tax", "VAT", ...]
    var tx, reTx = /taxtype\s*code\s*'([^']+)'/gi;
    while ((tx = reTx.exec(hay)) !== null) out.taxHints.push(String(tx[1]));
    out.taxHints = Array.from(new Set(out.taxHints));

    out.summary = (out.messages.slice(0,3).join(' | ') || '400 validation');
  } catch(e) {
    out.summary = '400 validation (parse error)';
  }
  return out;
}


function _downgrade(li, cfg) {
  cfg = cfg || {};
  var d = JSON.parse(JSON.stringify(li || {}));
  var keepTax = (typeof XP_KEEP_TAXCODE_ON_DOWNGRADE === 'boolean' ? XP_KEEP_TAXCODE_ON_DOWNGRADE : true);
  if (cfg.dropItemCode) delete d.ItemCode;
  if (cfg.dropTaxType || !keepTax) delete d.TaxType;
  if (cfg.forceFallbackAC && XP_FALLBACK_REVENUE_ACCT) d.AccountCode = XP_FALLBACK_REVENUE_ACCT;
  return d;
}
function _maybeAnnotate_(d, reasonTag) {
  if (!XP_ANNOTATE_DOWNGRADE_IN_DESCRIPTION) return d;
  var tag = (typeof XP_DOWNGRADE_TAG === 'string' ? XP_DOWNGRADE_TAG : '[DOWNGRADED]');
  var desc = d.Description || '';
  var reason = reasonTag ? ' {'+reasonTag+'}' : '';
  if (desc.indexOf(tag) !== 0) d.Description = tag + reason + ' ' + desc;
  return d;
}

function _ok_(res){ return res && (res.code===200 || res.code===201); }
function _ensureMeta_(res, extra){ res = res || {}; res.meta = res.meta || { downgradedIdxs: [] }; if (extra) for (var k in extra) res.meta[k]=extra[k]; return res; }
function _withMetaTried_(res, idxs, tried, branch){ res=_ensureMeta_(res); res.meta.downgradedIdxs=idxs||[]; res.meta.tried=tried||{}; res.meta.branch=branch||''; return res; }
function _actualIdxs_(idxsFromParse, res){ return (idxsFromParse && idxsFromParse.length) ? idxsFromParse : (res.meta && res.meta.downgradedIdxs || []); }
function _getInv_(payload){ return (payload && payload.Invoices && payload.Invoices[0]) || {}; }
function _hasLines_(payload){ try{ var a=payload.Invoices[0].LineItems; return Array.isArray(a)&&a.length>0; }catch(e){return false;} }
function _safeParseJSON_(s, fb){ try{ return s? JSON.parse(s): fb; }catch(e){ return fb; } }
function _safeHeaders_(resp){ try{ return resp.getAllHeaders ? resp.getAllHeaders() : {}; }catch(e){ return {}; } }
function _dx_respMeta_(res){ var h=res.headers||{}; return { code:res.code, reqid:h['xero-correlation-id']||h['x-correlation-id']||'', rate:{rem:h['x-rate-limit-remaining']||'',limit:h['x-rate-limit-problem']||''}, body:(res.body||'').slice(0,280)}; }

function _lineIsNote_(li) {
  var desc = String((li && li.Description) || '');
  var qty  = (li && li.Quantity != null) ? Number(li.Quantity) : 0;
  var unit = (li && li.UnitAmount != null) ? Number(li.UnitAmount) : 0;
  return (qty === 0 && unit === 0 && /^\s*\[note:\s*/i.test(desc));
}

function _candidateOrder_(lines, parsed400) {
  // Make a candidate list excluding notes; prefer known account code offenders
  var accSet = new Set(parsed400 && parsed400.accountHints || []);
  var cand = [];

  for (var i=0;i<lines.length;i++) {
    if (_lineIsNote_(lines[i])) continue;
    var li = lines[i] || {};
    var acct = String(li.AccountCode || '');
    var score = _suspectScore_(li);

    // If error mentioned specific accounts, heavily prefer those
    if (accSet.size && accSet.has(acct)) score += 10;

    cand.push({ i:i, s:score });
  }

  cand.sort(function(a,b){ return b.s - a.s; });
  return cand.map(function(x){ return x.i; });
}
function _suspectScore_(li) {
  var s = 0;
  if (!li) return 0;
  if (!li.ItemCode)    s += 3;
  if (!li.AccountCode) s += 2;
  if (!li.TaxType)     s += 1;

  var t = (li.TaxType || '').toLowerCase();
  if (t === 'sales tax' || /sales\s*tax(?!.*\d)/.test(t)) s += 2;

  if (/\[discount applied:/i.test(String(li.Description||''))) s += 0.5;
  return s;
}
