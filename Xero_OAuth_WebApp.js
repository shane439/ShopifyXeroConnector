/* ============================================================================
 * (Optional) WebApp OAuth handler from Xero_OAuth_WebApp
 *  - Deploy as Web App (Execute as: Me; Access: Anyone) and set XERO_REDIRECT_URI
 * ==========================================================================*/

const XERO_TOKEN_URL = 'https://identity.xero.com/connect/token';
const XERO_CONNS_URL = 'https://api.xero.com/connections';

function XERO_StartOAuth() {
  const state = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty('XERO_OAUTH_STATE', state);
  const url = xeroBuildAuthUrl_(state);
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(
      '<p>Click to authorize Xero in a new tab:</p>'
      + '<p><a href="' + url + '" target="_blank" rel="noopener">Authorize Xero</a></p>'
      + '<p>After approval, you will be redirected to the Web App and see a success page.</p>'),
    'Xero Authorization'
  );
}
function xeroBuildAuthUrl_(state) {
  const clientId = S_get_('XERO_CLIENT_ID','');
  const redirectUri = S_get_('XERO_REDIRECT_URI','');
  if (!clientId || !redirectUri) throw new Error('XERO_CLIENT_ID or XERO_REDIRECT_URI missing.');
  const params = {
    response_type: 'code',
    client_id: clientId,
    redirect_uri: redirectUri,
    scope: [
      'offline_access',
      'accounting.transactions',
      'accounting.contacts',
      'accounting.settings',
      'accounting.attachments'
    ].join(' '),
    state: state
  };
  const q = Object.keys(params).map(function(k){ return encodeURIComponent(k)+'='+encodeURIComponent(params[k]); }).join('&');
  return 'https://login.xero.com/identity/connect/authorize?' + q;
}
function doGet(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};
    const code = params.code || '';
    const state = params.state || '';
    if (!code) return htmlPage_('Xero OAuth', 'Missing "code" in redirect. Did you approve the app?');

    const expected = PropertiesService.getScriptProperties().getProperty('XERO_OAUTH_STATE') || '';
    if (expected && state !== expected) {
      return htmlPage_('Xero OAuth', 'Invalid state. Please restart authorization from the Sheet menu.');
    }

    const redirectUri = S_get_('XERO_REDIRECT_URI','');
    const token = xeroExchangeCodeForTokens_(code, redirectUri);

    if (token.refresh_token) {
      PropertiesService.getScriptProperties().setProperty('XERO_REFRESH_TOKEN', token.refresh_token);
    }

    var tenantId = S_get_('XERO_TENANT_ID','');
    if (!tenantId) {
      const conns = xeroGetConnections_(token.access_token);
      if (conns && conns.length) {
        tenantId = conns[0].tenantId || conns[0].tenant_id || '';
        if (tenantId) PropertiesService.getScriptProperties().setProperty('XERO_TENANT_ID', tenantId);
      }
    }

    return htmlPage_(
      'Xero OAuth',
      '<h2>✅ Xero Connected</h2>'
      + '<p>Refresh token saved. Tenant: <b>' + (tenantId || '(not set)') + '</b></p>'
      + '<p>You can close this tab and return to the Sheet.</p>'
    );
  } catch (err) {
    return htmlPage_('Xero OAuth', '❌ Error: ' + String(err && err.message || err));
  }
}
function xeroExchangeCodeForTokens_(code, redirectUri) {
  const a = S_xeroAuth_();
  const body = { grant_type: 'authorization_code', code: code, redirect_uri: redirectUri };
  const tokenResp = UrlFetchApp.fetch(XERO_TOKEN_URL, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: Object.keys(body).map(function(k){ return encodeURIComponent(k)+'='+encodeURIComponent(body[k]); }).join('&'),
    headers: { Authorization: 'Basic ' + Utilities.base64Encode(a.clientId + ':' + a.clientSecret) },
    muteHttpExceptions: true
  });
  const codeResp = tokenResp.getResponseCode();
  const text = tokenResp.getContentText();
  if (codeResp < 200 || codeResp >= 300) throw new Error('Token exchange failed ' + codeResp + ': ' + text);
  return JSON.parse(text);
}
function xeroGetConnections_(accessToken) {
  const resp = UrlFetchApp.fetch(XERO_CONNS_URL, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + accessToken },
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  const txt = resp.getContentText();
  if (code < 200 || code >= 300) throw new Error('Connections failed ' + code + ': ' + txt);
  return JSON.parse(txt);
}
function htmlPage_(title, bodyHtml) {
  return HtmlService.createHtmlOutput(
    '<html><head><meta charset="utf-8"><title>'+title+'</title></head>'
    + '<body style="font-family:system-ui;padding:24px;line-height:1.5">'
    + bodyHtml + '</body></html>'
  ).setTitle(title);
}
