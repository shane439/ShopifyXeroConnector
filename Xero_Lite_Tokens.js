/* ============================================================================
 * Xero Access Token Refresh & Tenant (from Xero_Lite)
 * ==========================================================================*/

function XL_getAccessToken_(){
  const a = S_xeroAuth_();
  const tok = UrlFetchApp.fetch('https://identity.xero.com/connect/token', {
    method:'post', contentType:'application/x-www-form-urlencoded',
    payload: { grant_type: 'refresh_token', refresh_token: a.refresh },
    headers: { Authorization: 'Basic ' + Utilities.base64Encode(a.clientId + ':' + a.clientSecret) },
    muteHttpExceptions:true
  });
  const code = tok.getResponseCode();
  const txt  = tok.getContentText();
  if (code!==200) throw new Error('Xero token failed '+code+' :: '+txt.slice(0,400));
  const body = JSON.parse(txt);
  if (body.refresh_token) {
    try { PropertiesService.getScriptProperties().setProperty('XERO_REFRESH_TOKEN', body.refresh_token); } catch(e){}
  }
  return body.access_token;
}

function XL_ensureTenantId_(accessToken) {
  var tenantId = S_get_('XERO_TENANT_ID','');
  if (tenantId) return tenantId;
  const con = UrlFetchApp.fetch('https://api.xero.com/connections', {
    headers:{Authorization:'Bearer '+accessToken, Accept:'application/json'},
    muteHttpExceptions:true
  });
  const code = con.getResponseCode();
  const txt  = con.getContentText();
  if (code===200) {
    const arr = JSON.parse(txt) || [];
    if (arr[0] && (arr[0].tenantId || arr[0].tenant_id)) {
      tenantId = arr[0].tenantId || arr[0].tenant_id;
      try { PropertiesService.getScriptProperties().setProperty('XERO_TENANT_ID', tenantId); } catch(e){}
      return tenantId;
    }
  }
  throw new Error('Xero tenant not set. Reconnect OAuth.');
}
