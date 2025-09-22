function _hasFn_(name){ try{return typeof this[name]==='function';}catch(e){return false;} }
function _asArray_(x){ return Array.isArray(x)?x:(x?[x]:[]); }

function _shopifyCfg_(){
  if(!_hasFn_('S_shopifyBase_')) throw new Error('S_shopifyBase_() not found (check SX_Compat_Settings).');
  var cfg=S_shopifyBase_(); if(!cfg||!cfg.base||!cfg.token) throw new Error('Invalid Shopify base config.');
  return cfg;
}
function _nextLink_(headers){
  if(!headers) return null; var link=headers.Link||headers.link||headers['LINK']||'';
  var m=/<([^>]+)>\s*;\s*rel="next"/i.exec(link); return m?m[1]:null;
}

// --- V2 BY DATE (define if missing) ---
if(!_hasFn_('SFv2_fetchOrdersByDate_')){
  function SFv2_fetchOrdersByDate_(fromISO,toISO){
    var cfg=_shopifyCfg_();
    var url = cfg.base + '/orders.json?status=any&limit=250&order=created_at%20asc'
            + '&created_at_min=' + encodeURIComponent(fromISO)
            + '&created_at_max=' + encodeURIComponent(toISO);
    var H={'X-Shopify-Access-Token':cfg.token}, out=[];
    while(url){
      var res=UrlFetchApp.fetch(url,{method:'get',headers:H,muteHttpExceptions:true});
      if(res.getResponseCode()>=400) throw new Error('Shopify fetch error '+res.getResponseCode());
      var body=JSON.parse(res.getContentText()||'{}'); out=out.concat(_asArray_(body.orders||[]));
      url=_nextLink_(res.getHeaders());
    }
    return out;
  }
}

// --- V2 BY NUMBERS (define if missing) ---
if(!_hasFn_('SFv2_fetchOrdersByNumbers_')){
  function SFv2_fetchOrdersByNumbers_(numbers){
    var cfg=_shopifyCfg_(), H={'X-Shopify-Access-Token':cfg.token}, out=[];
    numbers=_asArray_(numbers);
    for(var i=0;i<numbers.length;i++){
      var n=String(numbers[i]).trim(); if(!n) continue;
      var url=cfg.base+'/orders.json?status=any&limit=250&name='+encodeURIComponent('#'+n);
      var res=UrlFetchApp.fetch(url,{method:'get',headers:H,muteHttpExceptions:true});
      if(res.getResponseCode()>=400) throw new Error('Shopify fetch error '+res.getResponseCode()+' for #'+n);
      var body=JSON.parse(res.getContentText()||'{}'); out=out.concat(_asArray_(body.orders||[]));
    }
    return out;
  }
}

function _ensureOrdersInbox_(){
  var ss=SpreadsheetApp.getActive(), sh=ss.getSheetByName('OrdersInbox');
  var hdr=['order_id','order_number','created_at','financial_status','currency',
           'customer_name','customer_email','line_count','subtotal_price','total_tax','total_price','raw_json'];
  if(!sh){ sh=ss.insertSheet('OrdersInbox'); sh.getRange(1,1,1,hdr.length).setValues([hdr]); }
  else { var h=sh.getRange(1,1,1,hdr.length).getValues()[0]; if(!h[0]) sh.getRange(1,1,1,hdr.length).setValues([hdr]); }
  return sh;
}
function _compactJSON_(o,max){ var s=JSON.stringify(o); return (max&&s.length>max)?(s.slice(0,max-3)+'...'):s; }
function _SF_writeOrders_V2_(orders){
  var sh=_ensureOrdersInbox_(), rows=[];
  _asArray_(orders).forEach(function(o){
    if(!o) return; var c=o.customer||{};
    rows.push([
      o.id||'', String(o.name||'').replace(/^#/,''),
      o.created_at||'', o.financial_status||'', o.currency||'',
      (c.first_name?c.first_name+' ':'')+(c.last_name||''), c.email||'',
      (o.line_items?o.line_items.length:0), Number(o.subtotal_price||0),
      Number(o.total_tax||0), Number(o.total_price||0), _compactJSON_(o,49000)
    ]);
  });
  if(!rows.length) return 0;
  sh.getRange(sh.getLastRow()+1,1,rows.length,12).setValues(rows);
  return rows.length;
}
function _SF_writeOrders_orLegacy_(orders){
  return _hasFn_('SF_writeOrders_') ? SF_writeOrders_(_asArray_(orders)) : _SF_writeOrders_V2_(orders);
}
function SFv2_dispatchFetch(payload){
  if(!payload||!payload.mode) throw new Error('No mode specified.');
  var mode=String(payload.mode), orders=[];
  var MAX_RANGE=200, MAX_LIST=200;

  if(mode==='date'){
    var d=String(payload.date||'');
    if(!/^\d{4}-\d{2}-\d{2}$/.test(d)) throw new Error('Date must be YYYY-MM-DD.');
    var fromISO=d+'T00:00:00-04:00', toISO=d+'T23:59:59-04:00';
    orders=SFv2_fetchOrdersByDate_(fromISO,toISO);

  }else if(mode==='range'){
    var s=Number(payload.rangeStart||0), e=Number(payload.rangeEnd||0);
    if(!(s>0&&e>0&&e>=s)) throw new Error('Provide a valid From/To range.');
    if(e-s+1>MAX_RANGE) throw new Error('Range too large (max '+MAX_RANGE+').');
    var nums=[]; for(var n=s;n<=e;n++) nums.push(n);
    orders=SFv2_fetchOrdersByNumbers_(nums);

  }else if(mode==='list'){
    var list=String(payload.listNumbers||'').split(',').map(function(x){return x.trim();}).filter(String);
    if(!list.length) throw new Error('Provide at least one order number.');
    if(list.length>MAX_LIST) throw new Error('Too many numbers (max '+MAX_LIST+').');
    orders=SFv2_fetchOrdersByNumbers_(list);

  }else{
    throw new Error('Unsupported mode: '+mode);
  }

  var written=_SF_writeOrders_orLegacy_(orders)||0;
  return 'Fetched '+written+' order(s) into OrdersInbox.';
}
function SFv2_openDialog(){
  var html=HtmlService.createHtmlOutput(
    '<div style="font:14px system-ui;padding:16px;max-width:460px;">'
    + '<h2 style="margin:0 0 12px;">Fetch Orders (V2)</h2>'
    + '<p style="margin:8px 0 12px;line-height:1.5">Choose one method. V2 uses cursor pagination (Shopify-recommended) and always refetches.</p>'
    + '<label style="display:block;margin-top:8px;"><input type="radio" name="mode" value="date" checked> By Date</label>'
    + '<div id="byDate" style="margin:8px 0 12px 24px;"><label>Select date (UTC-4):<br><input type="date" id="date" style="width:100%"></label></div>'
    + '<label style="display:block;margin-top:8px;"><input type="radio" name="mode" value="range"> By Order # Range</label>'
    + '<div id="byRange" style="margin:8px 0 12px 24px;display:none"><div style="display:flex;gap:8px;">'
    + '<label style="flex:1">From #<input type="number" id="rangeStart" style="width:100%"></label>'
    + '<label style="flex:1">To #<input type="number" id="rangeEnd" style="width:100%"></label></div>'
    + '<div style="color:#666;font-size:12px;margin-top:4px;">Range cap: 200 orders.</div></div>'
    + '<label style="display:block;margin-top:8px;"><input type="radio" name="mode" value="list"> By Specific Order #s</label>'
    + '<div id="byList" style="margin:8px 0 12px 24px;display:none"><label>Comma-separated (e.g., 1001,1007,1013)<br>'
    + '<input type="text" id="listNumbers" placeholder="1001,1007,1013" style="width:100%"></label>'
    + '<div style="color:#666;font-size:12px;margin-top:4px;">Max 200 numbers.</div></div>'
    + '<div style="display:flex;gap:8px;justify-content:flex-end;margin-top:16px;">'
    + '<button type="button" onclick="google.script.host.close()">Cancel</button>'
    + '<button id="goBtn" type="button" style="font-weight:600">Fetch</button></div>'
    + '<script>'
    + 'const qs=(n)=>document.querySelector(n);'
    + 'function setYesterdayUTC4(){try{var now=new Date();var utc=new Date(now.getTime()+now.getTimezoneOffset()*60000);var tz4=new Date(utc.getTime()-4*3600*1000);tz4.setDate(tz4.getDate()-1);var y=tz4.toISOString().slice(0,10);var el=document.getElementById("date");if(el&&!el.value)el.value=y;}catch(e){}}'
    + 'function wireModes(){Array.from(document.getElementsByName("mode")).forEach(function(r){r.addEventListener("change",function(){qs("#byDate").style.display=r.value==="date"?"block":"none";qs("#byRange").style.display=r.value==="range"?"block":"none";qs("#byList").style.display=r.value==="list"?"block":"none";});});}'
    + 'function submitForm(){var mode=Array.from(document.getElementsByName("mode")).find(x=>x.checked).value;var p={mode:mode};'
    + 'if(mode==="date"){p.date=(qs("#date").value||"").trim();if(!/^\\d{4}-\\d{2}-\\d{2}$/.test(p.date)){alert("Pick a date.");return;}}'
    + 'if(mode==="range"){p.rangeStart=(qs("#rangeStart").value||"").trim();p.rangeEnd=(qs("#rangeEnd").value||"").trim();}'
    + 'if(mode==="list"){p.listNumbers=(qs("#listNumbers").value||"").trim();}'
    + 'var b=qs("#goBtn");b.disabled=true;b.textContent="Fetching…";'
    + 'google.script.run.withSuccessHandler(function(msg){alert(msg||"Done");google.script.host.close();})'
    + '.withFailureHandler(function(err){alert(String(err));b.disabled=false;b.textContent="Fetch";})'
    + '.SFv2_dispatchFetch(p);}'
    + 'document.addEventListener("DOMContentLoaded",function(){setYesterdayUTC4();wireModes();qs("#goBtn").addEventListener("click",submitForm);});'
    + '</script></div>'
  ).setWidth(500).setHeight(460);
  SpreadsheetApp.getUi().showModalDialog(html,'Fetch Orders (V2)');
}

function SFv2_openDialog() {
  var html = HtmlService.createHtmlOutput(
    '<div style="font:14px system-ui;padding:16px;max-width:460px;">'
    + '<h2 style="margin:0 0 12px;">Fetch Orders (V2)</h2>'
    + '<p style="margin:8px 0 12px;line-height:1.5">'
    + 'Choose one method. V2 uses cursor pagination (Shopify-recommended) and always refetches.'
    + '</p>'

    + '<label style="display:block;margin-top:8px;"><input type="radio" name="mode" value="date" checked> By Date</label>'
    + '<div id="byDate" style="margin:8px 0 12px 24px;">'
    + '  <label>Select date (UTC-4):<br><input type="date" id="date" style="width:100%"></label>'
    + '</div>'

    + '<label style="display:block;margin-top:8px;"><input type="radio" name="mode" value="range"> By Order # Range</label>'
    + '<div id="byRange" style="margin:8px 0 12px 24px;display:none">'
    + '  <div style="display:flex;gap:8px;">'
    + '    <label style="flex:1">From #<input type="number" id="rangeStart" style="width:100%"></label>'
    + '    <label style="flex:1">To #<input type="number" id="rangeEnd" style="width:100%"></label>'
    + '  </div>'
    + '  <div style="color:#666;font-size:12px;margin-top:4px;">Range cap: 200 orders for safety.</div>'
    + '</div>'

    + '<label style="display:block;margin-top:8px;"><input type="radio" name="mode" value="list"> By Specific Order #s</label>'
    + '<div id="byList" style="margin:8px 0 12px 24px;display:none">'
    + '  <label>Comma-separated (e.g., 1001,1007,1013)<br>'
    + '  <input type="text" id="listNumbers" placeholder="1001,1007,1013" style="width:100%"></label>'
    + '  <div style="color:#666;font-size:12px;margin-top:4px;">Max 200 numbers per run.</div>'
    + '</div>'

    + '<div style="display:flex;gap:8px;justify-content:flex-end;margin-top:16px;">'
    + '  <button type="button" onclick="google.script.host.close()">Cancel</button>'
    + '  <button id="goBtn" type="button" style="font-weight:600">Fetch</button>'
    + '</div>'

    + '<script>'
    + 'const qs = (n)=>document.querySelector(n);'
    + 'function setYesterdayUTC4(){'
    + '  try{'
    + '    var now=new Date();'
    + '    var utc=new Date(now.getTime()+now.getTimezoneOffset()*60000);'
    + '    var tz4=new Date(utc.getTime()-4*3600*1000);'
    + '    tz4.setDate(tz4.getDate()-1);'
    + '    var y=tz4.toISOString().slice(0,10);'
    + '    var el=document.getElementById("date"); if(el && !el.value) el.value=y;'
    + '  }catch(e){}'
    + '}'
    + 'function wireModeToggles(){'
    + '  Array.from(document.getElementsByName("mode")).forEach(function(r){'
    + '    r.addEventListener("change", function(){'
    + '      qs("#byDate").style.display  = r.value==="date"  ? "block":"none";'
    + '      qs("#byRange").style.display = r.value==="range" ? "block":"none";'
    + '      qs("#byList").style.display  = r.value==="list"  ? "block":"none";'
    + '    });'
    + '  });'
    + '}'
    + 'function submitForm(){'
    + '  var mode = Array.from(document.getElementsByName("mode")).find(x=>x.checked).value;'
    + '  var payload={mode:mode};'
    + '  if(mode==="date"){ payload.date = (qs("#date").value||"").trim(); if(!payload.date){alert("Pick a date."); return;} }'
    + '  if(mode==="range"){ payload.rangeStart = (qs("#rangeStart").value||"").trim(); payload.rangeEnd = (qs("#rangeEnd").value||"").trim(); }'
    + '  if(mode==="list"){ payload.listNumbers = (qs("#listNumbers").value||"").trim(); }'
    + '  var btn=qs("#goBtn"); btn.disabled=true; btn.textContent="Fetching…";'
    + '  google.script.run.withSuccessHandler(function(msg){'
    + '      alert(msg||"Done"); google.script.host.close();'
    + '    }).withFailureHandler(function(err){'
    + '      alert(String(err)); btn.disabled=false; btn.textContent="Fetch";'
    + '    }).SFv2_dispatchFetch(payload);'
    + '}'
    + 'document.addEventListener("DOMContentLoaded", function(){'
    + '  setYesterdayUTC4(); wireModeToggles(); qs("#goBtn").addEventListener("click", submitForm);'
    + '});'
    + '</script>'
    + '</div>'
  ).setWidth(500).setHeight(460);
  SpreadsheetApp.getUi().showModalDialog(html, 'Fetch Orders (V2)');
}
function SFv2_dispatchFetch(payload) {
  if (!payload || !payload.mode) throw new Error('No mode specified.');
  var mode = String(payload.mode);

  // Guardrail settings (keep aligned with your Flags_Config if present)
  var MAX_RANGE = 200;
  var MAX_LIST  = 200;

  var fetched = 0, orders = [];

  if (mode === 'date') {
    var d = String(payload.date || '');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) throw new Error('Date must be YYYY-MM-DD.');
    var fromISO = d + 'T00:00:00-04:00';
    var toISO   = d + 'T23:59:59-04:00';
    orders = SFv2_fetchOrdersByDate_(fromISO, toISO);

  } else if (mode === 'range') {
    var s = Number(payload.rangeStart || 0), e = Number(payload.rangeEnd || 0);
    if (!(s>0 && e>0 && e>=s)) throw new Error('Provide a valid From/To range.');
    if (e - s + 1 > MAX_RANGE) throw new Error('Range too large (max ' + MAX_RANGE + ').');
    orders = SFv2_fetchOrdersByNumbers_(Array.apply(null, Array(e-s+1)).map(function(_,i){return s+i;}));

  } else if (mode === 'list') {
    var list = String(payload.listNumbers || '').split(',').map(function(x){return x.trim();}).filter(String);
    if (!list.length) throw new Error('Provide at least one order number.');
    if (list.length > MAX_LIST) throw new Error('Too many numbers (max ' + MAX_LIST + ').');
    orders = SFv2_fetchOrdersByNumbers_(list);

  } else {
    throw new Error('Unsupported mode: ' + mode);
  }

  fetched = SF_writeOrders_(orders) || 0;
  return 'Fetched ' + fetched + ' order(s) into OrdersInbox.';
}
