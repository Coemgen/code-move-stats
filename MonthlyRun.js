/*jslint browser:true, long:true, white:true*/
/*global
DriveApp, FIRST_STAFF_ROW, PropertiesService, SendEmail, SitesApp,
SpreadsheetApp
*/

/**
 * @file Defines the <code><b>MonthlyRun</b></code> module.  The module builds
 * the file structure used for compiling and viewing stats.  Create a script
 * trigger to run the <code><b>MonthlyRun.main()</b></code> method via a driver
 * function one time each month.
 * <p>Before using this module, values must be set for these {@linkcode
 * https://developers.google.com/apps-script/guides/properties
 * script properties}:
 * <ul>
 *  <li><b>yearlyStatsTemplateId</b>&nbsp;&ndash;&nbsp;the spreadsheet id for
 *  the yearly stats template</li>
 *  <li><b>dataFolderId</b>&nbsp;&ndash;&nbsp;the parent folder for yearly data
 *  folders</li>
 *  <li><b>codeMoveTemplateId</b>&nbsp;&ndash;&nbsp;the spreadsheet id for the
 *  monthly totals template</li>
 *  <li><b>googleSiteUrl</b>&nbsp;&ndash;&nbsp;the url for the project's
 *  associated Google Site</li>
 * </ul>
 * @author Kevin Griffin <kevin.griffin@gmail.com>
 */

/******************************************************************************/

/**
 * @namespace MonthlyRun
 */

// eslint-disable-next-line no-unused-vars
const MonthlyRun = (

  function (DriveApp, PropertiesService, SitesApp, SpreadsheetApp) {
    "use strict";

    /**
     * Creates and returns a new yearly stats file object.  Also adds the stats
     * file's url to the associated Google Site.
     * @function addYearlyStatsFile
     * @memberof MonthlyRun
     * @private
     * @param {Object} yearlyStatsTemplate
     * @param {string} yearStr
     * @param {Object} yearFolder
     * @returns {Object}
     */
    function addYearlyStatsFile(yearlyStatsTemplate, yearStr, yearFolder) {
      const yearlyStatsFile = DriveApp.getFileById(yearlyStatsTemplate.getId())
        .makeCopy(
          ("Weekend Days OHS Stat tracking information "
            + yearStr),
          yearFolder);
      const site = SitesApp.getSiteByUrl(
        PropertiesService.getScriptProperties().getProperty("googleSiteUrl")
      );
      const ohsStatsListPage = site.getChildByName("ohs-stats");
      const values = [
        "<a href=\"" + yearlyStatsFile.getUrl() + "\">"
        + yearStr + " OHS Stats</a>"
      ];
      ohsStatsListPage.addListItem(values);

      return yearlyStatsFile;
    }

    /**
     * Returns a reference to the folder object, for the current year, and its
     * yearly stats spreadsheet.  If the folder does not already exist, a new
     * one will be created and populated with a yearly stats spreadsheet.
     * @function getYearFolder
     * @memberof MonthlyRun
     * @private
     * @param {string} yearStr
     * @returns {object[]} - [yearfolder, yearlyStatsFile] object references
     */
    function getYearFolder(yearStr) {
      const yearlyStatsTemplate = SpreadsheetApp.openById(
        PropertiesService.getScriptProperties()
        .getProperty("yearlyStatsTemplateId")
      );
      // find root folder
      const dataFolder = DriveApp.getFolderById(
        PropertiesService.getScriptProperties().getProperty("dataFolderId")
      );
      const folderIterator = dataFolder.getFoldersByName(yearStr);
      const yearFolder = (
        (folderIterator.hasNext() === true)
        ? folderIterator.next()
        : dataFolder.createFolder(yearStr));
      const fileIterator = yearFolder.getFilesByName(
        "Weekend Days OHS Stat tracking information "
        + yearStr);
      const yearlyStatsFile = (
        (fileIterator.hasNext() === true)
        ? fileIterator.next()
        : addYearlyStatsFile(yearlyStatsTemplate, yearStr, yearFolder));

      return [yearFolder, yearlyStatsFile];
    }

    /**
     * Creates and returns a new monthly code moves file object.  Also adds the code
     * moves file's url to the associated Google Site.
     * @param {Object} codeMoveTemplate
     * @param {string} codeMoveSheetName
     * @param {Object} yearFolder
     * @param {Object} dateObj
     * @returns {Object}
     */
    function addCodeMoveFile(
      codeMoveTemplate, codeMoveSheetName, yearFolder, dateObj) {
      const codeMoveFile = codeMoveTemplate.makeCopy(
        codeMoveSheetName, yearFolder
      );
      const site = SitesApp.getSiteByUrl(
        PropertiesService.getScriptProperties().getProperty("googleSiteUrl")
      );
      const codeMovePage = site.getChildByName("code-move-counts");
      const year = dateObj.getFullYear();
      const month = dateObj.toLocaleDateString("en-US", {
          "month": "numeric"
        })
        .padStart(2, "0");
      const urlStr = "<a href=\"" + codeMoveFile.getUrl() + "\">"
        + year
        + "-"
        + month
        + "</a>";
      const values = [
        urlStr
      ];
      codeMovePage.addListItem(values);

      return codeMoveFile;
    }

    /**
     * Returns a the file ID for the monthly data file object, for the current
     * month.  If the file does not already exist, a new one will be created.
     * @function getCodeMoveFile
     * @memberof MonthlyRun
     * @private
     * @param {Object} yearFolder - the data folder
     * @param {string} yearMonthStr - used for naming the spreadsheet and sheets
     * @param {Object} dateObj - JavaScript date object for current month
     * @returns {Object} monthly data spreadsheet file object
     */
    function getCodeMoveFile(yearFolder, yearMonthStr, dateObj) {
      // get code move template
      const codeMoveTemplate = DriveApp.getFileById(
        PropertiesService.getScriptProperties()
        .getProperty("codeMoveTemplateId")
      );
      // get or create YYYY-MM spreadsheet
      const codeMoveSheetName = yearMonthStr;
      const fileIterator = yearFolder.getFilesByName(codeMoveSheetName);
      const yearMonthFileFound = fileIterator.hasNext();
      const codeMoveFile = (
        (yearMonthFileFound)
        ? fileIterator.next()
        : addCodeMoveFile(
          codeMoveTemplate, codeMoveSheetName, yearFolder, dateObj));
      let spreadsheet = {};

      // edit new month spreadsheets
      if (yearMonthFileFound === false) {
        spreadsheet = SpreadsheetApp.openById(
          codeMoveFile.getId()
        );
        spreadsheet.getSheetByName("Totals")
          .getRange("A1:A3")
          .setValue(dateObj);

        // for each staff member update link to and from Totals sheet
        spreadsheet.getSheetByName("Totals").getRange("A4:A23")
          .getValues().map((nameArr) => nameArr[0])
          .filter((name) => name).forEach(
            function (name, index) {
              // eslint-disable-next-line no-undef
              const row = FIRST_STAFF_ROW + index; /* jshint ignore:line */
              const sheet = spreadsheet.getSheetByName(name);
              const email = sheet.getRange("B1").getValue();

              spreadsheet.getSheetByName("Totals").getRange("A" + row).setValue(
                "=HYPERLINK(\""
                + "https://docs.google.com/spreadsheets/d/"
                + spreadsheet.getId()
                + "/edit#gid="
                + spreadsheet.getSheetByName(name).getSheetId()
                + "\", \""
                + name
                + "\")"
              );
              spreadsheet.getSheetByName(name).getRange("A1").setValue(
                "=HYPERLINK(\""
                + "https://docs.google.com/spreadsheets/d/"
                + spreadsheet.getId()
                + "/edit#gid="
                + spreadsheet.getSheetByName("Totals").getSheetId()
                + "\", \""
                + "Totals"
                + "\")"
              );

              // set protections
              try {
                sheet.protect().setDomainEdit(false)
                  .addEditor(email);
                sheet.getRange("A1:Z2").protect()
                  .setDomainEdit(false)
                  .removeEditor(email);
              } catch (e) {
                console.log(e);
              }

              return undefined;
            }
          );
      } // end if statement

      return codeMoveFile;
    }

    /**
     * Populate yearly stats Spreadsheet's Weekend Days sheet with references to
     * cells in the year's monthly data sheets.
     * @function updateYearlyStatsFile
     * @memberof MonthlyRun
     * @private
     * @param {Object} yearlyStatsFile
     * @param {Object} codeMoveFile
     * @param {number} month - 0 to 11
     * @param {string} yearMonthStr - "Weekend Code Move Count YYYY-MM" format
     * @returns {undefined}
     */
    function updateYearlyStatsFile(
      yearlyStatsFile, codeMoveFile, month, yearMonthStr) {
      const spreadsheet = SpreadsheetApp.openById(yearlyStatsFile.getId());
      const importedDataSheet = spreadsheet.getSheetByName("Imported Data");
      const row = 2;
      const column = month + 2;
      const colChar = String.fromCharCode(66 + month);
      const formula = "=IMPORTRANGE(" + colChar + "2,\"Totals!AD1:AD\")";

      spreadsheet.getSheetByName("Weekend Days")
        .getRange("A1")
        .setValue("Weekend Days OHS Stats " + yearMonthStr.slice(24, 28));

      importedDataSheet.getRange(row, column).setValue(codeMoveFile.getUrl());
      importedDataSheet.getRange(row + 1, column).setFormula(formula);

      return undefined;
    }

    /**
     * Gets the current year's data folder.  If the folder does not exist then it
     * will be created and populated with a yearly stats spreadsheet. A monthly
     * totals data entry spreadsheet will be created for the current month and
     * linked to the yearly stats spreadsheet.
     * @function main
     * @memberof! MonthlyRun
     * @public
     * @param {number} [testYear=undefined] - YYYY format
     * @param {number} [testMonth=undefined] - 0...11
     * @returns {undefined}
     */
    // eslint-disable-next-line no-unused-vars
    function main(testYear = undefined, testMonth = undefined) {

      let yearFolder = {};
      let codeMoveFile = {};
      let yearlyStatsFile = {};

      const dateObj = (
        ((testYear !== undefined) && (testMonth !== undefined))
        ? new Date(testYear, testMonth)
        : new Date());
      const yearStr = dateObj.getFullYear().toString();
      const month = dateObj.getMonth();
      const monthStr = String(month + 1).toString().padStart(2, "0");
      const yearMonthStr = "Weekend Code Move Count"
        + " "
        + yearStr
        + "-"
        + monthStr;

      /* jshint ignore:start */
      // See: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Operators/Destructuring_assignment#Array_destructuring
      /* jshint ignore:end */
      [yearFolder, yearlyStatsFile] = getYearFolder(yearStr);

      codeMoveFile = getCodeMoveFile(yearFolder, yearMonthStr, dateObj);

      updateYearlyStatsFile(
        yearlyStatsFile, codeMoveFile, month, yearMonthStr);

      SendEmail.main(codeMoveFile.getId(), yearStr, monthStr);

      return undefined;
    }

    return Object.freeze({
      main,
      updateYearlyStatsFile
    });

  }(DriveApp, PropertiesService, SitesApp, SpreadsheetApp));

/******************************************************************************/