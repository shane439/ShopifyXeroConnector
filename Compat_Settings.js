/* ============================================================================
 * Settings & Base Builders (from Compat_Settings + Shopify_Lite)
 * - Uses Script Properties / CONFIG tab. No hardcoded secrets.
 * ==========================================================================*/

// Hardcoded override OFF (use Script Properties / CONFIG tab)
const HARDCODE_OVERRIDE = false;
const HARD = {}; // no inline secrets

function S_get_(key, defVal) {
  if (HARDCODE_OVERRIDE && Object.prototype.hasOwnProperty.call(HARD, key)) {
    const v = HARD[key];
    if (v !== null && v !== undefined && String(v) !== '') return String(v);
  }
  try {
    const v = PropertiesService.getScriptProperties().getProperty(String(key));
    if (v !== null && v !== undefined && String(v) !== '') return String(v);
  } catch (e) {}
  try {
    const ss = SpreadsheetApp.getActive();
    const tabs = ['CONFIG', 'config', 'Settings'];
    for (var t = 0; t < tabs.length; t++) {
      const sh = ss.getSheetByName(tabs[t]);
      if (sh && sh.getLastRow() >= 2) {
        const rows = sh.getRange(2,1, sh.getLastRow()-1, 2).getValues();
        const hit = rows.find(function(r){ return String(r[0]).trim() === String(key); });
        if (hit && String(hit[1]).trim() !== '') return String(hit[1]).trim();
      }
    }
  } catch (e) {}
  return defVal;
}
function S_bool_(key, def) { return ['true','1','yes','y'].includes(String(S_get_(key, def)).trim().toLowerCase()); }
function S_int_(key, def)  { const n = parseInt(S_get_(key, def), 10); return isNaN(n) ? def : n; }

// Shopify base (verbatim behavior)
function S_shopifyBase_() {
  var dom = String(S_get_('SHOPIFY_STORE_DOMAIN','')).trim().replace(/^https?:\/\//,'').replace(/\/+$/,'');
  if (!/\.myshopify\.com$/i.test(dom)) throw new Error('SHOPIFY_STORE_DOMAIN must end with .myshopify.com');

  var raw = S_get_('SHOPIFY_API_VERSION', '2024-07');
  var ver = '';
  if (Object.prototype.toString.call(raw) === '[object Date]') {
    var d = raw;
    ver = String(d.getFullYear()) + '-' + ('0' + (d.getMonth()+1)).slice(-2);
  } else {
    var m = String(raw).match(/(\d{4}-\d{2})/);
    ver = m ? m[1] : '';
  }
  if (!/^\d{4}-\d{2}$/.test(ver)) throw new Error('SHOPIFY_API_VERSION must be YYYY-MM (e.g., 2024-07)');

  var token = S_get_('SHOPIFY_ADMIN_TOKEN', '');
  if (!token) throw new Error('SHOPIFY_ADMIN_TOKEN missing.');

  return { dom: dom, ver: ver, token: token, base: 'https://' + dom + '/admin/api/' + ver };
}

// Xero auth settings
function S_xeroAuth_() {
  const clientId     = S_get_('XERO_CLIENT_ID','');
  const clientSecret = S_get_('XERO_CLIENT_SECRET','');
  const tenantId     = S_get_('XERO_TENANT_ID','');
  const refresh      = S_get_('XERO_REFRESH_TOKEN','');
  if (!clientId || !clientSecret) throw new Error('Xero client credentials missing.');
  if (!refresh) throw new Error('Xero refresh token missing. Reconnect OAuth.');
  return { clientId: clientId, clientSecret: clientSecret, tenantId: tenantId, refresh: refresh };
}
