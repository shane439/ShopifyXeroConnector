// /* ============================================================================
//  * Feature Flags & Tunables (single source of truth)
//  * ==========================================================================*/
// // === New Beta Flags ===
// var XP_BETA1_ENABLE_PARTIAL_DOWNGRADE = false;   // <-- toggle this ON to activate new behavior
// var XP_DOWNGRADE_TAG          = '[DOWNGRADED]';  // prefix in Description for downgraded lines
// var XP_FALLBACK_REVENUE_ACCT  = '200';           // safe AccountCode, or '' to omit
// var XP_KEEP_TAXCODE_ON_DOWNGRADE = true;         // keep original TaxType on downgrade

// var XP_IDEMPOTENCY_LOCAL_MARKER   = true;   // write InvoiceID to row; skip already pushed rows
// var XP_IDEMPOTENCY_XERO_CHECK     = true;   // query Xero for existing DRAFT by Reference; skip if found
// var XP_RATE_LIMIT_BACKOFF         = true;   // per-invoice delay + one 60s retry on 429
// var XP_FALLBACK_ON_VALIDATION400  = true;   // on HTTP 400 “Validation” only: retry without ItemCode
// var XP_LOG_ATTEMPT_LAYER          = true;   // include (A/B) path info in alert

// // Localized hardeners (your current states preserved)
// var SF_USE_COMPACT_JSON           = true;   // compact raw_json to avoid 50k cell limit
// var SF_AUTO_ADD_MARKER_COLUMNS    = false;  // auto-add PushedAt/PushedInvoiceID if missing

// // Tunables (used only when the related flag is true)
// var XP_DELAY_MS                   = 1000;   // 1s pacing
// var XP_RETRY_429_SLEEP_MS         = 60000;  // 60s before retry
// var XP_LOCAL_MARKER_COLS = {
//   pushedAt:   'PushedAt',
//   invoiceId:  'PushedInvoiceID'
// };
