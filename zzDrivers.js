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
function zzmonthlyRunMain() {
  "use strict";
  zzMonthlyRun.main();
}

/**
 * Run this periodically to add or remove staff from the monthly totals
 * template spreadsheet.
 */
// eslint-disable-next-line no-unused-vars
function zzinitCodeMoveTemplateMain() {
  "use strict";
  zzInitCodeMoveTemplate.main();
}

/**
 * Links the yearly stats spreadsheet template Weekend Days sheet to its
 * Imported Data sheet.  This should only need to be run for a brand new yearly
 * stats template spreadsheet.
 */
// eslint-disable-next-line no-unused-vars
function zzinitStatsTemplateMain() {
  "use strict";
  zzInitStatsTemplate.main();
}

/**
 * Use this function to relink a deleted yearly stats spreadsheet to its
 * associated monthly totals spreadsheets.  This should only need to be run
 * if there is a problem with the link between yearly stats and monthly totals
 * sheets.
 */
// eslint-disable-next-line no-unused-vars
function zzrestoreImportedDataMain() {
  "use strict";
  zzRestoreImportedData.main(2020);
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
  var fileIterator = yearFolder.getFilesByName(yearStr + "-stats");
  const yearlyStatsFile = fileIterator.next();

  fileIterator = yearFolder.getFiles();
  var fileObj = fileIterator.next();
      fileObj = fileIterator.next();
  var fileName = fileObj.getName();
  Logger.log("fileName: " + fileName);
  if (fileName.match(/^\d{4}-\d{2}$/) !== null)
  {
    var monthlyStatsFile = fileObj;
    //SendEmail.main(monthlyStatsFile,"Jun","testing","Weekend Code Move Count");
  }

  SendEmail.main(yearlyStatsFile, "Jun", "testing","OHS Stat"); // (yearlyStatsFile, monthStr, flagTesting, monthlyStatsFile, typeEmail)

  return undefined;
}


/**
 * Use this function to replace =IMPORTRANGE formula with static cell values
 * to remove "allow access" button.
 */
// eslint-disable-next-line no-unused-vars
function removeImportRangeMain() {
  "use strict";
  RemoveImportRange.main(2020);

  // Actual use may be:
  // 1.) restoreImportedDataMain() - to connect/update data with current and previous monthly sheet data
  // 2.) removeImportRangeMain() - remove ImportRange/"allow access" dependency
}