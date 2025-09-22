function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Shopify ↔ Xero')
    // Fetch (V2)
    .addItem('Fetch Orders (Standard – V2)…', 'SFv2_openDialog')

    .addSeparator()

    // Push: CURRENT vs BETA V3
    .addItem('Push Selected (Inclusive Only) — CURRENT', 'XL_pushSelectedDraft_InclusiveOnly_Standalone')
    .addItem('Push Selected (Inclusive Only) — Beta V3', 'XL_pushSelectedDraft_InclusiveOnly_Standalone_BetaV3')

    .addSeparator()

    // Utilities (keep these if you have them)
    .addItem('Test Shopify', 'SF_testShopify_')
    .addItem('Test Xero', 'XL_testXero_')

    .addToUi();
}
