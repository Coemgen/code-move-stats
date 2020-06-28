/*jslint browser:true, white:true*/
/*global
DriveApp, InitCodeMoveTemplate, InitStatsTemplate, MonthlyRun,
PropertiesService, RestoreImportedData, SendEmail
*/

// Project documentation at:  https://coemgen.github.io/code-move-stats/2.0.0/index.html

/**
 * @file Defines driver functions for running public methods.
 *
 * <p>Google Script Project values must be set for the following
 * {@linkcode https://developers.google.com/apps-script/guides/properties
 * script properties}:
 * <ul>
 *  <li><b>codeMoveTemplateId</b>&nbsp;&ndash;&nbsp;the spreadsheet id for the
 *  montly totals template</li>
 *  <li><b>dataFolderId</b>&nbsp;&ndash;&nbsp;the parent folder for yearly data
 *  folders</li>
 *  <li><b>distributionType</b>&nbsp;&ndash;&nbsp;
 *  Email distribution type (values are: <b>live</b> or <b>test</b>)</li>
 *  <li><b>groupEmail</b>&nbsp;&ndash;&nbsp;the Google Group email associated
 *  with this project</li>
 *  <li><b>yearlyStatsTemplateId</b>&nbsp;&ndash;&nbsp;the spreadsheet id for
 *  the yearly stats template</li>
 * </ul>
 * <p>Run the script using the {@linkcode
 * https://developers.google.com/apps-script/guides/v8-runtime V8 Runtime}.
 */

/**
 * @namespace Drivers
 */

/**
 * Function to be called by a monthly {@linkcode
 * https://developers.google.com/apps-script/guides/triggers/installable
 * Trigger} to set up the totals spreadsheet for each month and a yearly
 * stats spreadsheet for each year.
 * @function monthlyRunMain
 * @memberof Drivers
 * @public
 */
// eslint-disable-next-line no-unused-vars
function monthlyRunMain() {
  "use strict";
  MonthlyRun.main();
}

/**
 * Updates staff on the Code Move Count template spreadsheet.
 * <br>Run this function:
 * <ol>
 * <li>Monthly via a {@linkcode
 * https://developers.google.com/apps-script/guides/triggers/installable
 * Trigger} before running the <b><code>monthlyRunMain()</code></b> function</li>
 * <li>Periodically when changes have been made to staffing in the associated
 * Code Move Google Group</li>
 * </ol>
 * @function initCodeMoveTemplateMain
 * @memberof Drivers
 * @public
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
 * @function initStatsTemplateMain
 * @memberof Drivers
 * @private
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
 * @function restoreImportedDataMain
 * @memberof Drivers
 * @private
 */
// eslint-disable-next-line no-unused-vars
function restoreImportedDataMain() {
  "use strict";
  RestoreImportedData.main(2020);
}

/**
 * Use this funtion for testing the project.
 * Months are numbered 0..11
 * @function monthlyRunTest
 * @memberof Drivers
 * @private
 */
// eslint-disable-next-line no-unused-vars
function monthlyRunTest() {
  "use strict";
  const numMonths = 6;
  const startYear = 2019;
  const startMonth = 11;
  const monthArr = Array.from({
    "length": numMonths
  });
  monthArr.forEach(function (ignore, index) {
    MonthlyRun.main(startYear, startMonth + index);
  });

  return undefined;
}

/**
 * Function to be called by a weekly {@linkcode
 * https://developers.google.com/apps-script/guides/triggers/installable
 * Trigger} to send a reminder for the Code Move Group to update the
 * spreadsheet.
 * @function sendWeeklyReminder
 * @memberof Drivers
 * @public
 */
// eslint-disable-next-line no-unused-vars
function sendWeeklyReminder() {
  "use strict";
  const d = new Date();
  const yearStr = d.getFullYear().toString();
  const month = d.getMonth() + 1;
  const monthStr = month.toString().padStart(2, 0);
  const dataFolder = DriveApp.getFolderById(
    PropertiesService.getScriptProperties().getProperty("dataFolderId")
  );
  const folderIterator = dataFolder.getFoldersByName(yearStr);
  const yearFolder = folderIterator.next();
  const fileIterator = yearFolder.getFilesByName(yearStr + "-" + monthStr);
  const codeMoveFileId = fileIterator.next().getId();
  const reminder = true;
  SendEmail.main(codeMoveFileId, yearStr, monthStr, reminder);
}