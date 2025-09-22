// function onOpen(){
//   var ui = SpreadsheetApp.getUi();
//   ui.createMenu('Shopify ↔ Xero')
//     .addItem('Fetch Orders (Standard – V2)…', 'SFv2_openDialog')
//     .addSeparator()
//     .addItem('Push Selected (Inclusive Only)', 'XL_pushSelected_Dispatch_')
//     .addSeparator()
//     .addItem('Test Shopify', 'SF_testShopify_')
//     .addItem('Test Xero', 'XL_testXero_')
//     .addToUi();
// }

// /** Routes to CURRENT or BETA by flag + presence; avoids “function not found”. */
// function XL_pushSelected_Dispatch_(){
//   try {
//     var wantBeta = (typeof XP_BETA1_ENABLE_PARTIAL_DOWNGRADE !== 'undefined') && XP_BETA1_ENABLE_PARTIAL_DOWNGRADE;
//     var hasBeta  = (typeof this['XL_pushSelectedDraft_InclusiveOnly_Standalone_BetaV1'] === 'function');
//     if (wantBeta && hasBeta) return XL_pushSelectedDraft_InclusiveOnly_Standalone_BetaV1();
//     return XL_pushSelectedDraft_InclusiveOnly_Standalone(); // CURRENT
//   } catch (e) {
//     SpreadsheetApp.getUi().alert('Dispatch error: ' + e.message);
//     throw e;
//   }
// }
