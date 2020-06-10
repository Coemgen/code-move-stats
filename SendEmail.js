/*jslint browser:true, long:true, white:true*/
/*global
DriveApp, GmailApp, PropertiesService, StaffUtilities
*/
// eslint-disable-next-line no-unused-vars
const SendEmail = (
  function (DriveApp, GmailApp, PropertiesService) {

    function main(yearlyStatsFile, monthStr, flagTesting) {

      // --- Gather Email Addresses --

      var userEmailAddress;

      if (flagTesting) {
        userEmailAddress = ["jeburns@meditech.com", "kgriffin@meditech.com"];

      } else {

        const groupEmail = PropertiesService.getScriptProperties()
          .getProperty("groupEmail");
        userEmailAddress = StaffUtilities.getObjArr(groupEmail)
          .map((userObj) => userObj.email);
      }

      // --- Build Email Message Components ---
      var recipients, subject, body, options;

      if (userEmailAddress.length > 0) {

        recipients = (userEmailAddress.length == 1)
          ? userEmailAddress[0]
          : recipients = userEmailAddress.join(",");

        subject = 'MONTHLY: {file.getName} is now available for editing in Google Drive!';
        subject = subject.replace(/\{file.getName\}/g, yearlyStatsFile.getName());

        body = '<p>Click the following link to access the new sheet: <a href="{file.getUrl}">{file.getName}</a></p>\
              <div><br></div>\
<div>Hi everyone,<br><br>This is your monthly reminder message of the Weekend Code Move Count Spreadsheet! The new spreadsheet tab for {getMonthInfo} has been created. Please remember to update the spreadsheet each and every weekend. Thanks</div>';
        body = body.replace(/\{file.getName\}/g, yearlyStatsFile.getName())
          .replace(/\{getMonthInfo\}/g, monthStr)
          .replace(/\{file.getUrl\}/g, yearlyStatsFile.getUrl());

        options = {
          htmlBody: body
        };

        // ---- Use Gamil Service to send email(s) ----
        GmailApp.sendEmail(recipients, subject, body, options);
      }

    }
    return Object.freeze({
      main
    });

  }(DriveApp, GmailApp, PropertiesService));