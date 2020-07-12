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
function zsendMonthlyOhsStatsReminder() {
  "use strict";
  
  // Pseudo Code
  // 1) Confirm it's the first of the month, before sending out an email
  // 2) Get file URL for email message body
  // 3) Compose email contents/options
  // 4) Send email
  
  // Comment line below for testing
  //if ( (new Date().getMonth() !== 1) ) return; // Not the 1st.
  
  var strYear = new Date().getFullYear().toString().trim();
  var folderData = DriveApp.getFolderById( PropertiesService.getScriptProperties().getProperty("dataFolderId") );
  var folderIterator = folderData.getFoldersByName( strYear );
  var folderYear = (folderIterator.hasNext()) ? folderIterator.next() : false;
  
  if ( !folderYear )
  {
    Logger.log("YYYY Folder was not found: " + strYear);
    return; // No matching YYYY folder found
  }
  
  var fileName = "Weekend Days OHS Stat tracking information " + strYear;
  var fileIterator = folderYear.getFilesByName( fileName );
  var fileOhsStats = (fileIterator.hasNext()) ? fileIterator.next() : false;
  
  if ( !fileOhsStats ) {
    Logger.log("OHS Stats spreadsheet was not found: " + fileName);
    return; // No matching OHS Stat sheet found
  }
  
  var strFileUrl = fileOhsStats.getUrl();
  
  var recipients = "jeburns@meditech.com";
  var subject = "MONTHLY: " + fileName;
  
  var body  = '<p>Click the following link to access the sheet: <a href="{file.getUrl}">{file.getName}</a></p>';
      body += '<p>If you see !#REF in the cell, click the cell, and then click Allow Access to connect the data.</p>';
      body += '<p>Please enter all non-automated data values.</p>';
      body += '<p>If you have any questions/comments, please contact James E Burns or Kevin Griffin.</p>';
  
      body = body.replace(/\{file.getName\}/g, fileOhsStats.getName())
                 .replace(/\{file.getUrl\}/g, fileOhsStats.getUrl());
  
  
  var options = {htmlBody: body};

  // ---- Use Gamil Service to send email(s) ----
  GmailApp.sendEmail(recipients, subject, body, options);
}