/*jslint browser:true, long:true, white:true*/
/*global
MailApp, PropertiesService, StaffUtilities
*/

/**
 * @file Code for sending emails.  Create a script trigger to periodically run
 * SendEmail.main() from Drivers.gs.
 *
 * <p>Before using the script, values must be set for these {@linkcode
 * https://developers.google.com/apps-script/guides/properties
 * script properties}:
 * <ul>
 *  <li><b>distributionType</b>&nbsp;&ndash;&nbsp;
 * Email distribution type (values are: <b>live</b> or <b>test</b>)</li>
 *  <li><b>groupEmail</b>&nbsp;&ndash;&nbsp;the Google Group email associated with this project</li>
 * </ul>
 *
 * <p>Run the script using the {@linkcode
 * https://developers.google.com/apps-script/guides/v8-runtime V8 Runtime}.
 *
 * @author James Burns
 * @author Kevin Griffin <kevin.griffin@gmail.com>
 */

/**
 * @namespace SendEmail
 */

// eslint-disable-next-line no-unused-vars
const SendEmail = (

  function (MailApp, PropertiesService) {
    "use strict";

    /**
     * Takes a month number 01-12 and returns its corresponding month name
     * in Jan-Dec format.
     * @param {string} monthStr in 01-12 format
     * @returns {string} month in Jan-Dec format
     */
    function formatMonthStr(monthStr) {
      const d = new Date();
      d.setMonth(Number(monthStr) - 1);
      return d.toLocaleDateString(
        "en-US", {
          dateStyle: "long"
        }
      ).split(" ")[0];
    }

    /**
     * Send email to weekend code move group members
     * @function main
     * @memberof! SendEmail
     * @public
     * @param {string} codeMoveFileId
     * @param {string} yearStr
     * @param {string} monthStr
     * @param {boolean} reminder is true for weekly reminders
     */
    function main(codeMoveFileId, yearStr, monthStr, reminder) {

      const curMonth = formatMonthStr(monthStr);
      const distType = PropertiesService.getScriptProperties()
        .getProperty("distributionType");
      const notifType = (
        (reminder === true) ? "REMINDER" : "ATTENTION");
      const notifPeriod = (
        (reminder === true) ? "weekly" : "monthly");
      const subject = `${notifType}: Weekend Code Move Count for ${curMonth}`
        + ` ${yearStr} is available for editing in Google Drive!`;
      const body = `Click the following link to access the current sheet: `
        + `${"https://docs.google.com/spreadsheets/d/" + codeMoveFileId}`
        + ` [Weekend Code Move Count ${curMonth} ${yearStr}]`
        + "\n\nHi everyone,\n\n"
        + `This is your ${notifPeriod} reminder message for the Weekend`
        + ` Code Move Count Spreadsheet! Please remember to update the`
        + ` spreadsheet each and every weekend. Thanks`;
      const htmlBody = `<p>Click the following link to access the current`
        + ` sheet: <a href="`
        + `${"https: //docs.google.com/spreadsheets/d/" + codeMoveFileId}">`
        + `Weekend Code Move Count ${curMonth} ${yearStr}</a></p><div><br>`
        + `</div><div>Hi everyone, <br><br> This is your ${notifPeriod} `
        + `reminder message for the Weekend Code Move Count Spreadsheet!`
        + ` Please remember to update the spreadsheet each and every weekend.`
        + ` Thanks</div>`;
      const options = {
        "htmlBody": htmlBody
      };
      let recipients = "";

      if (distType === "live") {
        recipients = StaffUtilities.getObjArr(
            PropertiesService.getScriptProperties()
            .getProperty("groupEmail")
          ).map((userObj) => userObj.email)
          .toString();
      } else if (distType === "test") {
        recipients = "eyip@meditech.com,jeburns@meditech.com,"
          + "kgriffin@meditech.com";
      } else {
        recipients = "kevin.griffin@gmail.com";
      }

      MailApp.sendEmail(recipients, subject, body, options);
    }

    return Object.freeze({
      main
    });

  }(MailApp, PropertiesService));