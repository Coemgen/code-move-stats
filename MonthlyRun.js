/*jslint long:true, white:true*/

"use strict";

/**
 * @file This script creates the file structure used for compiling and viewing
 * stats.  Create a trigger to run its driver, monthlyRunMain, each month.
 *
 * <p>Before running the script, values must be set for these {@linkcode
 * https://developers.google.com/apps-script/guides/properties|
 * script properties}:
 * <ul>
 *  <li>yearlyStatsTemplateId</li>
 *  <li>dataFolderId</li>
 *  <li>codeMoveTemplateId</li>
 * </ul>
 *
 * <p>Run the script using the {@linkcode
 * https://developers.google.com/apps-script/guides/v8-runtime|V8 Runtime}.
 *
 * @author Kevin Griffin <kevin.griffin@gmail.com>
 */

/* jshint ignore:start */
/**
 * Declare global variables to satisfy linter expectations for strict mode.
 * - note that re-declaring a var type global that already has a value does not
 * - affect its value.
 * - See: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Statements/var#Description
 */
/* jshint ignore:end */
/** @global */
var DriveApp;
/** @global */
var PropertiesService;
/** @global */
var SpreadsheetApp;

/******************************************************************************/

/**
 * Returns a reference to the folder object, for the current year, and its
 * yearly stats spreadsheet.  If the folder does not already exist, a new one
 * will be created and populated with a yearly stats spreadsheet.
 * @param {string} yearStr
 * @returns {object[]} - References to year folder and yearly stats file objects
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
  const fileIterator = yearFolder.getFilesByName(yearStr + "-stats");
  const yearlyStatsFile = (
    (fileIterator.hasNext() === true)
    ? fileIterator.next()
    : DriveApp.getFileById(yearlyStatsTemplate.getId())
    .makeCopy((yearStr + "-stats"), yearFolder));

  return [yearFolder, yearlyStatsFile];
}

/**
 * Returns a the file ID for the monthly data file object, for the current
 * month.  If the file does not already exist, a new one will be created.
 * @param {Object} yearFolder - the data folder
 * @param {string} yearMonthStr - used for naming the spreadsheet and sheets
 * @returns {string} monthly data spreadsheet file ID
 */
function getCodeMoveFileId(yearFolder, yearMonthStr) {
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
  var spreadsheet = {};

  // edit new month spreadsheets
  if (yearMonthFileFound === false) {
    spreadsheet = SpreadsheetApp.openById(
      codeMoveFile.getId()
    );
    spreadsheet.getSheetByName("Totals")
      .getRange("A1:A3")
      .setValue(yearMonthStr);

    // for each staff member update link to and from Totals sheet
    spreadsheet.getSheetByName("Totals").getRange("A4:A23")
      .getValues().map((nameArr) => nameArr[0])
      .filter((name) => name).forEach(
        function (name, index) {
          // eslint-disable-next-line no-undef
          const row = FIRST_STAFF_ROW + index; /* jshint ignore:line */
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
        
        return undefined;
        }
      );
  } // end if statement

  return codeMoveFile.getId();
}

/**
 * Populate yearly stats Spreadsheet's Weekend Days sheet with references to
 * cells in the year's monthly data sheets.
 * @param {Object} yearlyStatsFile
 * @param {string} codeMoveFileId
 * @param {number} month - 0 to 11
 * @param {string} yearMonthStr - YYYY-MM format
 * @returns {undefined}
 */
function updateYearlyStatsFile(
  yearlyStatsFile, codeMoveFileId, month, yearMonthStr) {
  const yearlyStatsSheet = SpreadsheetApp.openById(yearlyStatsFile.getId())
    .getSheetByName("Imported Data");
  const row = month + 1;

  yearlyStatsSheet.getRange("A" + row).setValue(yearMonthStr);

  yearlyStatsSheet.getRange("B" + row).setFormula("=IMPORTRANGE("
    + "\"https://docs.google.com/spreadsheets/d/"
    + codeMoveFileId
    + "\",\"Totals!B24:AH24\")"
  );

  yearlyStatsSheet.getRange("AI" + row).setFormula("=IMPORTRANGE("
    + "\"https://docs.google.com/spreadsheets/d/"
    + codeMoveFileId
    + "\",\"Totals!H29\")"
  );

  yearlyStatsSheet.getRange("AJ" + row).setFormula("=IMPORTRANGE("
    + "\"https://docs.google.com/spreadsheets/d/"
    + codeMoveFileId
    + "\",\"Totals!H30\")"
  );

  yearlyStatsSheet.getRange("AK" + row).setFormula("=IMPORTRANGE("
    + "\"https://docs.google.com/spreadsheets/d/"
    + codeMoveFileId
    + "\",\"Totals!H31\")"
  );

  return undefined;
}

/**
 * Gets the current year's data folder.  If the folder does not exits then it
 * will be created.  The current year's folder will then be populated with
 * a yearly stats view spreadsheet and a monthly data entry spreadsheet.
 * @param {number} [testYear=undefined] - YYYY format
 * @param {number} [testMonth=undefined] - 0...11
 * @returns {undefined}
 */
// eslint-disable-next-line no-unused-vars
function monthlyRunMain(testYear = undefined, testMonth = undefined) {

  var yearFolder = {};
  var codeMoveFileId = "";
  var yearlyStatsFile = {};

  const dateObj = (
    ((testYear !== undefined) && (testMonth !== undefined))
    ? new Date(testYear, testMonth)
    : new Date());
  const yearStr = dateObj.getFullYear().toString();
  const month = dateObj.getMonth();
  const monthStr = String(month + 1).toString().padStart(2, "0");
  const yearMonthStr = yearStr + "-" + monthStr;

  /* jshint ignore:start */
  // See: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Operators/Destructuring_assignment#Array_destructuring
  /* jshint ignore:end */
  [yearFolder, yearlyStatsFile] = getYearFolder(yearStr);

  codeMoveFileId = getCodeMoveFileId(yearFolder, yearMonthStr);

  updateYearlyStatsFile(yearlyStatsFile, codeMoveFileId, month, yearMonthStr);

  return undefined;
}

/******************************************************************************/
