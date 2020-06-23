/*jslint browser:true, long:true, white:true*/
/*global
MailApp, PropertiesService, StaffUtilities
*/

/**
 * @file Code for sending emails.  Create a script trigger to periodically run 
 * SendEmail.main() from Drivers.gs.
 *
 * <p>Before using the script, values must be set for these {@linkcode
 * https://developers.google.com/apps-script/guides/properties|
 * script properties}:
 * <ul>
 *  <li><b>groupEmail</b>&nbsp;&ndash;&nbsp;the Google Group email associated with this project</li>
 * </ul>
 *
 * <p>Run the script using the {@linkcode
 * https://developers.google.com/apps-script/guides/v8-runtime|V8 Runtime}.
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
     * Takes a month number 01-12 and returns its corresponding month short
     * name in Jan-Dec format.
     * @param {string} monthStr in 1-12 format
     * @returns {string} mont in Jan format
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
     * @param {string} monthStr 
     * @param {string} distType live, test, or undefined
     */
    function main(codeMoveFileId, monthStr, distType) {

      const curMonth = formatMonthStr(monthStr);
      const subject = `MONTHLY: ${curMonth} is now available
      for editing in Google Drive!`;
      const body = `Hi everyone,<br><br>This is your monthly reminder message 
      for the Weekend Code Move Count Spreadsheet! A new spreadsheet has been 
      created for ${curMonth} at url 
      ${"https://docs.google.com/spreadsheets/d/" + codeMoveFileId}. 
      Please remember to update the spreadsheet each and every weekend. 
      Thanks`;
      const htmlBody = `<p>Click the following link to access the new sheet: 
        <a href="${"https://docs.google.com/spreadsheets/d/" + codeMoveFileId}">
        ${curMonth}</a></p><div><br>
        </div><div>Hi everyone,<br><br>This is your monthly reminder message for
        the Weekend Code Move Count Spreadsheet! A new spreadsheet has been 
        created for ${curMonth}. Please remember to update
        the spreadsheet each and every weekend. Thanks</div>`;
      const options = {
        htmlBody: htmlBody
      };
      let recipients = "";

      if (distType === "live") {
        recipients = StaffUtilities.getObjArr(
            PropertiesService.getScriptProperties()
            .getProperty("groupEmail")
          ).map((userObj) => userObj.email)
          .toString();
      } else if (distType === "test") {
        recipients = "jeburns@meditech.com,kgriffin@meditech.com";
      } else {
        recipients = "kevin.griffin@gmail.com";
      }

      MailApp.sendEmail(recipients, subject, body, options);
    }

    return Object.freeze({
      main
    });

  }(MailApp, PropertiesService));