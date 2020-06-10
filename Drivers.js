/*jslint browser:true, white:true*/
/*global 
DriveApp, InitCodeMoveTemplate, InitStatsTemplate, MonthlyRun, 
PropertiesService, RestoreImportedData, SendEmail
*/

/**
 * @file Driver functions used as wrappers for running public methods.
 */

// Project documentation at:  https://coemgen.github.io/code-move-stats/2.0.0/index.html

/**
 * Function to be called by the monthly Script Trigger to set up the totals
 * spreadsheet for each month and a yearly stats spreadsheet for each year.
 * @function monthlyRunMain
 */
// eslint-disable-next-line no-unused-vars
function monthlyRunMain() {
  "use strict";
  MonthlyRun.main();
}

/**
 * Run this periodically to add or remove staff from the montly totals
 * template spreadsheet.
 */
// eslint-disable-next-line no-unused-vars
function initCodeMoveTemplateMain() {
  "use strict";
  InitCodeMoveTemplate.main();
}

/**
 * Links the yearly stats spreadsheet template Weekend Days sheet to its
 * Imported Data sheet.  This should only need to be run for a brand new yearly
 * stats template spreadsheet.
 */
// eslint-disable-next-line no-unused-vars
function initStatsTemplateMain() {
  "use strict";
  InitStatsTemplate.main();
}

/**
 * Use this function to relink a deleted yearly stats spreadsheet to its
 * associated monthly totals spreadsheets.  This should only need to be run
 * if there is a problem with the link between yearly stats and monthly totals
 * sheets.
 */
// eslint-disable-next-line no-unused-vars
function restoreImportedDataMain() {
  "use strict";
  RestoreImportedData.main(2020);
}

/**
 * Use this funtion for testing the project.
 * Months are numbered 0..11
 */
// eslint-disable-next-line no-unused-vars
function monthlyRunTest() {
  "use strict";
  const numMonths = 5;
  const startYear = 2020;
  const startMonth = 0;
  const monthArr = Array.from({
    "length": numMonths
  });
  monthArr.forEach(function (ignore, index) {
    MonthlyRun.main(startYear, startMonth + index);
  });

  return undefined;
}

/**
 * Use this funtion for testing sending email
 */
// eslint-disable-next-line no-unused-vars
function emailTest() {
  "use strict";
  
  const dataFolder = DriveApp.getFolderById(
    PropertiesService.getScriptProperties().getProperty("dataFolderId")
  );
  const yearStr = "2020";
  const folderIterator = dataFolder.getFoldersByName(yearStr);
  const yearFolder = folderIterator.next();
  const fileIterator = yearFolder.getFilesByName(yearStr + "-stats");
  const yearlyStatsFile = fileIterator.next();

  SendEmail.main(yearlyStatsFile, "Jun", "testing");
  
  return undefined;
}