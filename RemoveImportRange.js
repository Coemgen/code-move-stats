/*jslint browser:true, long:true, white:true*/
/*global DriveApp, MonthlyRun, PropertiesService, SpreadsheetApp*/

/*
 * Issue:
 *  The OHS Stats Sheet (Management) requires the user to allow access to the the =IMPORTRANGE data used on the "Imported Data" sheet.
 *
 * Solution:
 *  By changing the cell formula (=IMPORTRANGE) to static values, this should remove the dependency and "allow access" requirement
 *
 * Drawbacks:
 *  To make sure the static values reflect the current values from the Monthly sheet, the following may need to be done:
 *    A. Change the formula to static values, once the monthly sheet has been changed to readonly for all users (less work)
 *    B. Routinely, through use of a Google Script Trigger, compare the values in the Imported Data sheet with the Monthly sheet and update accordingly. (more work)
 *        - Call RestoreImportedData to fix/update values and then RemoveImportRange to make them static values again
 */

// eslint-disable-next-line no-unused-vars
const RemoveImportRange = (

  function (DriveApp, MonthlyRun, PropertiesService) {
    "use strict";

    /**
     * Replaces =IMPORTRANGE with static values on "Imported Data" sheet
     */

    function isImportRangeConnected(sheetSource, rangeSource) {
      var cellCellCheck = sheetSource.getRange(rangeSource);
      var formulaCellCheck = cellCellCheck.getFormula();
      var valueCellCheck = cellCellCheck.getValue();

      if (valueCellCheck == "#REF!"
        && formulaCellCheck.indexOf("=IMPORTRANGE") == 0)
        return false;
      else
        return true;
    }

    // ; ----------------------------------------------------------------------------

    function getCellInformation(sheetSource, rangeSource) {

      var cellLocal = sheetSource.getRange(rangeSource);
      var formulaCell = cellLocal.getFormula();
      var valueCell = cellLocal.getValue();

      return [cellLocal, formulaCell, valueCell];
    }

    // ; ----------------------------------------------------------------------------

    function getImportRangeArgs(vFormula) {
      // Gather requirements
      var argImportRange = vFormula.match(/\((.+)\)/g)[0]
        .split(",").map(x => x.replace("(", "").replace(")", "")
          .replace("\"", "").replace("\"", ""));
      var argUrl = argImportRange[0];
      var argRange = argImportRange[1]; // includes the Sheet!Range
      //Logger.log("URL: " + argUrl);
      //Logger.log("Range: " + argRange);

      return [argUrl, argRange];
    }

    // ; ----------------------------------------------------------------------------

    function getTargetRangeValuesByImportRangeFormula(sourceFormula) {

      var argUrl, argRange;
      [argUrl, argRange] = getImportRangeArgs(sourceFormula);

      var targetSpreadsheet = SpreadsheetApp.openByUrl((argUrl + "/edit")); // SpreadsheetApp.openByUrl(url) doesn't appear to work without the /edit suffix
      var targetValues = targetSpreadsheet.getRange(argRange).getValues();
      //Logger.log("Values: " + JSON.stringify(targetValues));

      return targetValues;
    }

    // ; ----------------------------------------------------------------------------

    function setLocalRangeValues(
      colLetter, rowNumber, sheetSource,
      rowValues, isSourceConnected) {

      var rangeReplace = "";

      switch (colLetter) {
      case "B":
        rangeReplace = "B" + rowNumber + ":AH" + rowNumber;
        break;

      default:
        rangeReplace = colLetter + rowNumber + ":" + colLetter + rowNumber;
      }

      if (!isSourceConnected)
        // Removes IMPORTRANGE formula using values grabbed from target sheet via Server script; Allow access = N
        sheetSource.getRange(rangeReplace).setValues(rowValues);
      else
        // Removes IMPORTRANGE formula using values already imported to sheet; Allow access = Y
        sheetSource.getRange(rangeReplace)
        .setValues(sheetSource.getRange(rangeReplace)
          .getValues());


      return undefined;
    }

    // ; ----------------------------------------------------------------------------

    function getAndSetValuesForColumnLetter(sheetSource, colLetter, rowNumber) {

      // Checking B - =IMPORTRANGE("https://docs.google.com/spreadsheets/d/1tDQLgXEH17h4rKO-4OSLCTVDmEan2iSFDmLAx8aH7n4","Totals!B24:AH24")
      var rangeCell = colLetter + rowNumber;

      var formulaCell;
        [formulaCell] = getCellInformation(sheetSource, rangeCell);

      if (!isImportRangeConnected(sheetSource, rangeCell)) {
        // Get Values - gathers requirements from formula arguments
        // Gather values as the sheet isn't connected; Allow access:N
        var targetValues = getTargetRangeValuesByImportRangeFormula(
          formulaCell);

        // Set Values
        setLocalRangeValues(
          colLetter, rowNumber, sheetSource, targetValues, false);

      } else { // Is connected so turn =IMPORTRANGE into static data
        setLocalRangeValues(colLetter, rowNumber, sheetSource, "", true); // Otherwise just set the values as the sheet is already connected.
      }
    }

    // ; ----------------------------------------------------------------------------

    function main(yearToProcess) {

      const dataFolder = DriveApp.getFolderById(
        PropertiesService.getScriptProperties()
        .getProperty("dataFolderId"));
      const yearFolder = dataFolder.getFoldersByName(yearToProcess).next();

      var yearlyStatsFile = yearFolder.getFilesByName(
        yearToProcess + "-stats").next();
      var spreadsheetYearlyStats = SpreadsheetApp.open(yearlyStatsFile);

      if ((spreadsheetYearlyStats == null)
        || (spreadsheetYearlyStats == undefined))
        throw "Stats sheet was not found.";

      var sheetImportedData = spreadsheetYearlyStats.getSheetByName(
        "Imported Data");

      if ((sheetImportedData == null) || (sheetImportedData == undefined))
        throw "Imported Data sheet was not found.";

      // -------------------
      // Remove all cell formulas, replacing with data values
      //var dataImportedData = sheetImportedData.getDataRange().getValues();
      // replace cell formula with cell data
      //sheetImportedData.getDataRange().setValues(dataImportedData); // confirm that formula for the cells are gone. Restore version to reset.
      // -------------------

      // --------------------
      // * ONLY WORKS IF ACCESS HAD BEEN GIVEN ALREADY *
      // Only update B1:#AN, since A# is a =HYPERLINK cell we want to keep
      //var rowLast = sheetImportedData.getLastRow();
      //
      //if ( (rowLast < 1) || (isNaN(rowLast)) )
      //  return; // No data yet populated to the sheet.
      //
      //var rangeData = "B1:AN" + rowLast; // colLetter_colNumber:colLetter_colNumber
      //var dataImportedData = sheetImportedData.getRange(rangeData).getValues();
      //
      // Replace cell formulas with values
      //sheetImportedData.getRange(rangeData).setValues(dataImportedData);
      // --------------------

      // Go through data rows
      // if these column cells--B, AI-AN (AI, AJ, AK, AL, AM, AN)-- have this data, we need to open those sheets and get the values to put in here:
      //   formula: =IMPORTRANGE("https://docs.google.com/spreadsheets/d/1tDQLgXEH17h4rKO-4OSLCTVDmEan2iSFDmLAx8aH7n4","Totals!B24:AH24")
      //   value: #REF!
      //
      //   2-34  35-40 (35  36  37  38  39  40)
      //   B,    AI-AN (AI, AJ, AK, AL, AM, AN)
      //
      var rowLast = sheetImportedData.getLastRow();

      for (var i = 1; i <= rowLast; i++) {
        // Checking B - Example: =IMPORTRANGE("https://docs.google.com/spreadsheets/d/1tDQLgXEH17h4rKO-4OSLCTVDmEan2iSFDmLAx8aH7n4","Totals!B24:AH24")
        getAndSetValuesForColumnLetter(sheetImportedData, "B", i);
        getAndSetValuesForColumnLetter(sheetImportedData, "AI", i);
        getAndSetValuesForColumnLetter(sheetImportedData, "AJ", i);
        getAndSetValuesForColumnLetter(sheetImportedData, "AK", i);
        getAndSetValuesForColumnLetter(sheetImportedData, "AL", i);
        getAndSetValuesForColumnLetter(sheetImportedData, "AM", i);
        getAndSetValuesForColumnLetter(sheetImportedData, "AN", i);
      }

      return undefined;
    }

    return Object.freeze({
      main
    });
  }(DriveApp, MonthlyRun, PropertiesService)

);