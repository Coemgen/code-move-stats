/**
 * Function to be called by the monthly Script Trigger to set up the totals
 * spreadsheet for each month and a yearly stats spreadsheet for each year.
 * @function monthlyRunMain
 */
// eslint-disable-next-line no-unused-vars
function zMonthlyRunMain() {
  "use strict";
  zMonthlyRun.main();
}

/**
 * Run this periodically to add or remove staff from the monthly totals
 * template spreadsheet.
 */
// eslint-disable-next-line no-unused-vars
function zInitCodeMoveTemplateMain() {
  "use strict";
  zInitCodeMoveTemplate.main();
}

/**
 * Use this function to relink a deleted yearly stats spreadsheet to its
 * associated monthly totals spreadsheets.  This should only need to be run
 * if there is a problem with the link between yearly stats and monthly totals
 * sheets.
 */
// eslint-disable-next-line no-unused-vars
function zRestoreImportedDataMain() {
  "use strict";
  zRestoreImportedData.main(2020);
}

/**
 * Run this periodically to add or remove staff from the monthly totals
 * template spreadsheet.
 */
// eslint-disable-next-line no-unused-vars
function zInitStatsTemplateMain() {
  "use strict";
  zInitStatsTemplate.main();
}
