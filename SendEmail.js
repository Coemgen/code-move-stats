/*jslint browser:true, long:true, white:true*/
/*global
DriveApp, MailApp, PropertiesService, StaffUtilities
*/
// eslint-disable-next-line no-unused-vars
const SendEmail = (
  function (DriveApp, MailApp, PropertiesService) {

    // eslint-disable-next-line max-statements
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

        subject = "MONTHLY: {file.getName} is now available \
for editing in Google Drive!";
        subject = subject.replace(
          /\{file.getName\}/g, yearlyStatsFile.getName());

        body = "<p>Click the following link to access the new sheet:";
        body += " <a href=\"{file.getUrl}\">{file.getName}</a>";
        body += "</p><div><br></div>";
        body += "<div>Hi everyone,<br><br>This is your monthly reminder";
        body += " message of the Weekend Code Move Count Spreadsheet! The new";
        body += " spreadsheet tab for {getMonthInfo} has been created. Please";
        body += " remember to update the spreadsheet each and every weekend.";
        body += " Thanks</div>";
        body = body.replace(/\{file.getName\}/g, yearlyStatsFile.getName())
          .replace(/\{getMonthInfo\}/g, monthStr)
          .replace(/\{file.getUrl\}/g, yearlyStatsFile.getUrl());

        options = {
          htmlBody: body
        };

        // ---- Use Gamil Service to send email(s) ----
        MailApp.sendEmail(recipients, subject, body, options);
      }

    }
    return Object.freeze({
      main
    });

  }(DriveApp, MailApp, PropertiesService));