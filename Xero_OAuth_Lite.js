/* ============================================================================
 * Xero OAuth (prompt/paste) — from Xero_OAuth_Lite
 * ==========================================================================*/

function XO_startAuth(){
  const clientId = S_get_('XERO_CLIENT_ID','');
  const redirect = S_get_('XERO_REDIRECT_URI','');
  if (!clientId || !redirect) { SpreadsheetApp.getUi().alert('Missing XERO_CLIENT_ID or XERO_REDIRECT_URI in CONFIG/Properties.'); return; }

  const scopes = [
    'openid','profile','email','offline_access',
    'accounting.settings','accounting.transactions','accounting.contacts','accounting.journals.read'
  ].join(' ');
  const url = 'https://login.xero.com/identity/connect/authorize'
    + '?response_type=code'
    + '&client_id=' + encodeURIComponent(clientId)
    + '&redirect_uri=' + encodeURIComponent(redirect)
    + '&scope=' + encodeURIComponent(scopes)
    + '&prompt=consent';

  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Xero Reconnect',
    '1) Open this URL in your browser and approve:\n\n' + url + '\n\n2) You will be redirected to the Redirect URI with a long "code" in the URL.\n3) Copy ONLY that code (not the whole URL).\n4) Then run: Xero → Reconnect (OAuth)… again to paste the code.',
    ui.ButtonSet.OK
  );

  const r = ui.prompt('Paste the OAuth CODE here:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const code = r.getResponseText().trim();
  if (!code) { ui.alert('No code pasted.'); return; }

  XO_exchangeCode_(code);
}

function XO_exchangeCode_(code){
  const clientId = S_get_('XERO_CLIENT_ID','');
  const clientSecret = S_get_('XERO_CLIENT_SECRET','');
  const redirect = S_get_('XERO_REDIRECT_URI','');

  const tok = UrlFetchApp.fetch('https://identity.xero.com/connect/token', {
    method:'post', contentType:'application/x-www-form-urlencoded',
    payload: { grant_type: 'authorization_code', code: code, redirect_uri: redirect },
    headers: { Authorization: 'Basic ' + Utilities.base64Encode(clientId + ':' + clientSecret) },
    muteHttpExceptions:true
  });
  const codeHttp = tok.getResponseCode();
  if (codeHttp!==200) { SpreadsheetApp.getUi().alert('Token exchange failed '+codeHttp+' :: '+tok.getContentText().slice(0,400)); return; }

  const body = JSON.parse(tok.getContentText());
  if (body.refresh_token) PropertiesService.getScriptProperties().setProperty('XERO_REFRESH_TOKEN', body.refresh_token);

  // Discover tenant ID
  const con = UrlFetchApp.fetch('https://api.xero.com/connections', {headers:{Authorization:'Bearer '+body.access_token, Accept:'application/json'}});
  if (con.getResponseCode()===200) {
    const arr = JSON.parse(con.getContentText()) || [];
    if (arr[0] && arr[0].tenantId) {
      PropertiesService.getScriptProperties().setProperty('XERO_TENANT_ID', arr[0].tenantId);
    }
  }
  SpreadsheetApp.getUi().alert('Xero connected. Refresh token stored and tenant discovered.');
}
