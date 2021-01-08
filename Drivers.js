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
 * Email distribution type (values are: <b>live</b> or <b>test</b>)</li>
 *  <li><b>groupEmail</b>&nbsp;&ndash;&nbsp;the Google Group email associated
 *  with this project</li>
 *  <li><b>yearlyStatsTemplateId</b>&nbsp;&ndash;&nbsp;the spreadsheet id for
 *  the yearly stats template</li>
 *  <li><b>googleSiteUrl</b>&nbsp;&ndash;&nbsp;the url for the project's
 *  associated Google Site</li>
 * </ul>
 * <p>Run the script using the {@linkcode
 * https://developers.google.com/apps-script/guides/v8-runtime V8 Runtime}.
 */

/**
 * @namespace Drivers
 */

/**
 * Function to be called to set permissions when new Google Script script class
 * methods are being called.
 */
function getScriptPermissions() {
  return;
}

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
 * Use this funtion for testing the project.
 * Months are numbered 0..11
 * @function monthlyRunTest
 * @memberof Drivers
 * @private
 */
// eslint-disable-next-line no-unused-vars
function monthlyRunTest() {
  "use strict";
  const numMonths = 1;
  const startYear = 2017;
  const startMonth = 10;
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
  const dayInt = d.getDate();
  const yearStr = d.getFullYear().toString();
  const month = d.getMonth() + 1;
  const monthStr = month.toString().padStart(2, 0);
  const dataFolder = DriveApp.getFolderById(
    PropertiesService.getScriptProperties().getProperty("dataFolderId")
  );
  const folderIterator = dataFolder.getFoldersByName(yearStr);
  const yearFolder = folderIterator.next();
  const fileIterator = yearFolder.getFilesByName(
    "Weekend Code Move Count"
    + " "
    + yearStr
    + "-"
    + monthStr);
  const codeMoveFileId = fileIterator.next().getId();
  const reminder = true;

  if (dayInt > 0 && dayInt < 5) {
    // Avoid sending out a reminder email if the day is part of the same weekend that includes the 1st.
    return;
  }

  SendEmail.main(codeMoveFileId, yearStr, monthStr, reminder);
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
function sendMonthlyOhsStatsReminder() {
  "use strict";

  // Pseudo Code
  // 1) Confirm it's the first of the month, before sending out an email
  // 2) Get file URL for email message body
  // 3) Compose email contents/options
  // 4) Send email

  // Comment line below for testing
  //if ( (new Date().getMonth() !== 1) ) return; // Not the 1st.

  var strYear = new Date().getFullYear().toString().trim();
  var folderData = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty("dataFolderId"));
  var folderIterator = folderData.getFoldersByName(strYear);
  var folderYear = (folderIterator.hasNext()) ? folderIterator.next() : false;

  if (!folderYear) {
    Logger.log("YYYY Folder was not found: " + strYear);
    return; // No matching YYYY folder found
  }

  var fileName = "Weekend Days OHS Stat tracking information " + strYear;
  var fileIterator = folderYear.getFilesByName(fileName);
  var fileOhsStats = (fileIterator.hasNext()) ? fileIterator.next() : false;

  if (!fileOhsStats) {
    Logger.log("OHS Stats spreadsheet was not found: " + fileName);
    return; // No matching OHS Stat sheet found
  }

  var strFileUrl = fileOhsStats.getUrl();

  var recipients = "jeburns@meditech.com,"
    + "kgriffin@meditech.com,"
    + "rhomsey@meditech.com,"
    + "directs.rhomsey@meditech.com,"
    //    + "mjcarnino@meditech.com,"
    //    + "bporter@meditech.com,"
    //    + "kmahoney@meditech.com,"
    //    + "kkoppy@meditech.com,"
    //    + "agrachuk@meditech.com,"
    //    + "kallfrey@meditech.com,"
    + "kellis@meditech.com,"
    + "sgetchell@meditech.com,"
    + "eyip@meditech.com";
  var subject = "MONTHLY: " + fileName;

  var body = '<p>Click the following link to access the sheet: <a href="{file.getUrl}">{file.getName}</a></p>';
  body += '<p>If you see #REF! in the cell, click the cell, and then click Allow Access to connect the data.</p>';
  body += '<p>Please enter all non-automated data values.</p>';
  body += '<p>If you have any questions/comments, please contact James E Burns or Kevin Griffin.</p>';

  body = body.replace(/\{file.getName\}/g, fileOhsStats.getName())
    .replace(/\{file.getUrl\}/g, fileOhsStats.getUrl());


  var options = {
    htmlBody: body
  };

  // ---- Use Gamil Service to send email(s) ----
  GmailApp.sendEmail(recipients, subject, body, options);
}