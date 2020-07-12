/*jslint browser:true, long:true, white:true*/
/*global
DriveApp, FIRST_STAFF_ROW, PropertiesService, SendEmail, SpreadsheetApp
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
 *  montly totals template</li>
 * </ul>
 * @author Kevin Griffin <kevin.griffin@gmail.com>
 */

/******************************************************************************/

/**
 * @namespace MonthlyRun
 */

// eslint-disable-next-line no-unused-vars
const MonthlyRun = (

  function (DriveApp, PropertiesService, SpreadsheetApp) {
    "use strict";

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
        : DriveApp.getFileById(yearlyStatsTemplate.getId())
        .makeCopy(
          ("Weekend Days OHS Stat tracking information "
            + yearStr),
          yearFolder));

      return [yearFolder, yearlyStatsFile];
    }

    /**
     * Returns a the file ID for the monthly data file object, for the current
     * month.  If the file does not already exist, a new one will be created.
     * @function getCodeMoveFolder
     * @memberof MonthlyRun
     * @private
     * @param {Object} yearFolder - the data folder
     * @param {string} yearMonthStr - used for naming the spreadsheet and sheets
     * @param {Object} dateObj - JavaScript date object for current month
     * @returns {string} monthly data spreadsheet file ID
     */
    function getCodeMoveFileId(yearFolder, yearMonthStr, dateObj) {
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
        : codeMoveTemplate.makeCopy(codeMoveSheetName, yearFolder));
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

      return codeMoveFile.getId();
    }

    /**
     *
     * @param {string} url
     * @param {string} label
     */
    function getHyperlinkFormula(url, label) {
      return "=HYPERLINK(\"" + url + "\",\"" + label + "\")";
    }

    /**
     * Associate yearly stats imported data sheet with monthly code move sheet.
     * @function setMonthlyToStatsSheetLink
     * @memberof! MonthlyRun
     * @private
     * @param {Object} yearlyStatsSheet
     * @param {string} codeMoveFileId
     * @returns {setMonthlyToStatsSheetLink~monthlyCellToStatsCellLink}
     */
    function setMonthlyToStatsSheetLink(yearlyStatsSheet, codeMoveFileId) {
      /**
       * Link monthly totals cell to yearly stats sheet cell.
       * @function monthlyCellToStatsCellLink
       * @memberof! MonthlyRun
       * @private
       * @param {string} monthlyRange - monthly totals cell ref. in A1 notation
       * @param {string} yearlyRange - yearly stats cell ref. in A1 notation
       */
      const monthlyCellToStatsCellLink = function (monthlyRange, yearlyRange) {
        yearlyStatsSheet.getRange(yearlyRange).setFormula("=IMPORTRANGE("
          + "\"https://docs.google.com/spreadsheets/d/"
          + codeMoveFileId
          + "\",\"Totals!" + monthlyRange + "\")"
        );
      };
      return monthlyCellToStatsCellLink;
    }

    /**
     * Link Monthly Code Moves total's sheet footer values to
     * the yearly stats sheet.
     * @param {Object} yearlyStatsSheet
     * @param {string} codeMoveFileId
     * @param {number} row Yearly Stats Imported Data sheet row
     */
    function linkCodeMovesFooterToStats(yearlyStatsSheet, codeMoveFileId, row) {

      const monthlyCellToStatsCellLink = setMonthlyToStatsSheetLink(
        yearlyStatsSheet, codeMoveFileId);

      // Grand Totals (range B24:AH24)
      monthlyCellToStatsCellLink("B24:AH24", "B" + row);

      // H26 PE/MD Code Move Total (calculated on Weekend Stats sheet)

      // H27 Application Code Move Total (calculated on Weekend Stats sheet)

      // H29 Magic Update Total
      monthlyCellToStatsCellLink("H29", "AI" + row);

      // H30 C/S Update Total
      monthlyCellToStatsCellLink("H30", "AJ" + row);

      // H31 Expanse Update Total
      monthlyCellToStatsCellLink("H31", "AK" + row);

      // H34 Ring Deletion Total
      monthlyCellToStatsCellLink("H34", "AL" + row);

      // P34 TEST Setup Total
      monthlyCellToStatsCellLink("P34", "AM" + row);

      // H33 HCIS Deletion Total
      monthlyCellToStatsCellLink("H33", "AN" + row);

      // P36 Additions to Ship Source
      monthlyCellToStatsCellLink("P36", "AO" + row);

      return undefined;
    }

    /**
     * Populate yearly stats Spreadsheet's Weekend Days sheet with references to
     * cells in the year's monthly data sheets.
     * @function updateYearlyStatsFile
     * @memberof MonthlyRun
     * @private
     * @param {Object} yearlyStatsFile
     * @param {string} codeMoveFileId
     * @param {number} month - 0 to 11
     * @param {string} yearMonthStr - "Weekend Code Move Count YYYY-MM" format
     * @returns {undefined}
     */
    function updateYearlyStatsFile(
      yearlyStatsFile, codeMoveFileId, month, yearMonthStr) {
      const spreadsheet = SpreadsheetApp.openById(yearlyStatsFile.getId());
      const weekendDaysSheet = spreadsheet.getSheetByName("Weekend Days");
      const yearlyStatsSheet = spreadsheet.getSheetByName("Imported Data");
      // Yearly Stats Imported Data sheet row number.
      const row = month + 1;
      let codeMoveSheetUrl = "https://docs.google.com/spreadsheets/d/"
        + codeMoveFileId;
      let codeMoveSheetLabel = yearMonthStr;
      let codeMoveSheetHyperlinkFormula = getHyperlinkFormula(
        codeMoveSheetUrl, codeMoveSheetLabel);
      // Month columns range from B (ascii 66) to M (ascii 77)
      let colLetter = String.fromCharCode(66 + month);
      let colNumbers = [2, 13, 24, 27, 29, 34, 37];

      weekendDaysSheet.getRange("A1")
        .setValue("Weekend Days OHS Stats " + yearMonthStr.slice(24, 28));

      yearlyStatsSheet.getRange("A" + row)
        .setValue(yearMonthStr);

      yearlyStatsSheet.getRange("A" + row)
        .setFormula(codeMoveSheetHyperlinkFormula);

      // link month's column headers to its source spreadsheet
      colNumbers.forEach(
        function (colNumber) {
          const cellLocation = colLetter + colNumber.toString();
          const cellValue = weekendDaysSheet.getRange(cellLocation).getValue();
          weekendDaysSheet.getRange(cellLocation)
            .setFormula(getHyperlinkFormula(codeMoveSheetUrl, cellValue));
        }
      );

      linkCodeMovesFooterToStats(yearlyStatsSheet, codeMoveFileId, row);

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
      let codeMoveFileId = "";
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

      codeMoveFileId = getCodeMoveFileId(yearFolder, yearMonthStr, dateObj);

      updateYearlyStatsFile(
        yearlyStatsFile, codeMoveFileId, month, yearMonthStr);

      SendEmail.main(codeMoveFileId, yearStr, monthStr);

      return undefined;
    }

    return Object.freeze({
      main,
      updateYearlyStatsFile
    });

  }(DriveApp, PropertiesService, SpreadsheetApp));

/******************************************************************************/