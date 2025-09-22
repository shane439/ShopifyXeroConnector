// --- Safer Xero push flags (default OFF) ---
var XP_ENABLE_TARGETED_LINE_DOWNGRADE = false;  // if true: only failing lines are downgraded on 400
var XP_DOWNGRADE_TAG                 = '[DOWNGRADED]'; // prefix on the Description of downgraded lines
var XP_FALLBACK_REVENUE_ACCT         = '200';          // safe AccountCode for downgraded lines; '' to omit
var XP_KEEP_TAXCODE_ON_DOWNGRADE     = true;           // keep the original TaxType on downgraded lines
