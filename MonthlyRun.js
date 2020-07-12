/*jslint browser:true, long:true, white:true*/
/*global DriveApp, FIRST_STAFF_ROW, Logger, PropertiesService, SpreadsheetApp*/

/**
 * @file Code for building the file structure used for compiling and viewing
 * stats.  Create a script trigger to run MonthlyRun.main() from its driver
 * function once a month.
 *
 * <p>Before using the script, values must be set for these {@linkcode
 * https://developers.google.com/apps-script/guides/properties|
 * script properties}:
 * <ul>
 *  <li><b>yearlyStatsTemplateId</b>&nbsp;&ndash;&nbsp;the spreadsheet id for the yearly stats template</li>
 *  <li><b>dataFolderId</b>&nbsp;&ndash;&nbsp;the parent folder for yearly data folders</li>
 *  <li><b>codeMoveTemplateId</b>&nbsp;&ndash;&nbsp;the spreadsheet id for the montly totals template</li>
 * </ul>
 *
 * <p>Run the script using the {@linkcode
 * https://developers.google.com/apps-script/guides/v8-runtime|V8 Runtime}.
 *
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
     * yearly stats spreadsheet.  If the folder does not already exist, a new one
     * will be created and populated with a yearly stats spreadsheet.
     * @function getYearFolder
     * @memberof MonthlyRun
     * @private
     * @param {string} yearStr
     * @returns {object[]} - [yearfolder, yearlyStatsFile] object references
     */
    // 2020.06.20
    function getYearFolder(yearStr) {
//      /*
//      const yearlyStatsTemplate = SpreadsheetApp.openById(
//        PropertiesService.getScriptProperties()
//        .getProperty("yearlyStatsTemplateId")
//      );
//      */
      // find root folder
      const dataFolder = DriveApp.getFolderById(
        PropertiesService.getScriptProperties().getProperty("dataFolderId")
      );
      const folderIterator = dataFolder.getFoldersByName(yearStr);
      const yearFolderFound = folderIterator.hasNext();
//      /*
//      const yearFolder = (
//        (folderIterator.hasNext() === true)
//        ? folderIterator.next()
//        : dataFolder.createFolder(yearStr));
//        */
      const yearFolder = (
        (yearFolderFound)
        ? folderIterator.next()
        : dataFolder.createFolder(yearStr));
      
      // ------------
      // 2020.06.20 - separate get folder and get file
//      /*
//      const fileIterator = yearFolder.getFilesByName(yearStr + "-stats");
//      const yearlyStatsFile = (
//        (fileIterator.hasNext() === true)
//        ? fileIterator.next()
//        : DriveApp.getFileById(yearlyStatsTemplate.getId())
//        .makeCopy((yearStr + "-stats"), yearFolder));
//
//      return [yearFolder, yearlyStatsFile];
//      */
      return yearFolder;
      // ------------
    }
    
    function getYearlyStatsFile(yearStr, yearFolder){
      
      const yearlyStatsTemplate = SpreadsheetApp.openById(
        PropertiesService.getScriptProperties()
        .getProperty("yearlyStatsTemplateId")
      );
      
      const fileIterator = yearFolder.getFilesByName(yearStr + "-stats");
      const yearFileFound = fileIterator.hasNext();
      const yearlyStatsFile = (
        (yearFileFound)
        ? fileIterator.next()
        : DriveApp.getFileById(yearlyStatsTemplate.getId())
        .makeCopy((yearStr + "-stats"), yearFolder));

      return [yearlyStatsFile, yearFileFound];
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
    function getCodeMoveFile(yearFolder, yearMonthStr, dateObj) {
    //function getCodeMoveFileId(yearFolder, yearMonthStr, dateObj) {
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
      
//      /*
//      var spreadsheet = {};
//
//      // edit new month spreadsheets
//      if (yearMonthFileFound === false) {
//        spreadsheet = SpreadsheetApp.openById(
//          codeMoveFile.getId()
//        );
//        spreadsheet.getSheetByName("Totals")
//          .getRange("A1:A3")
//          .setValue(dateObj);
//
//        // for each staff member update link to and from Totals sheet
//        spreadsheet.getSheetByName("Totals").getRange("A4:A23")
//          .getValues().map((nameArr) => nameArr[0])
//          .filter((name) => name).forEach(
//            function (name, index) {
//              // eslint-disable-next-line no-undef
//              const row = FIRST_STAFF_ROW + index; /* jshint ignore:line *//*
//              const sheet = spreadsheet.getSheetByName(name);
//              const email = sheet.getRange("B1").getValue();
//
//              spreadsheet.getSheetByName("Totals").getRange("A" + row).setValue(
//                "=HYPERLINK(\""
//                + "https://docs.google.com/spreadsheets/d/"
//                + spreadsheet.getId()
//                + "/edit#gid="
//                + spreadsheet.getSheetByName(name).getSheetId()
//                + "\", \""
//                + name
//                + "\")"
//              );
//              spreadsheet.getSheetByName(name).getRange("A1").setValue(
//                "=HYPERLINK(\""
//                + "https://docs.google.com/spreadsheets/d/"
//                + spreadsheet.getId()
//                + "/edit#gid="
//                + spreadsheet.getSheetByName("Totals").getSheetId()
//                + "\", \""
//                + "Totals"
//                + "\")"
//              );
//
//              // set protections
//              try {
//                sheet.protect().setDomainEdit(false)
//                  .addEditor(email);
//                sheet.getRange("A1:Z2").protect()
//                  .setDomainEdit(false)
//                  .removeEditor(email);
//              } catch (e) {
//                console.log(e);
//              }
//
//              return undefined;
//            }
//          );
//      } // end if statement
//      */

      //return codeMoveFile.getId();
      //return codeMoveFile;
      return [codeMoveFile, yearMonthFileFound]
    }
    
    
    function updateMonthlyStatsFile(codeMoveFile, dateObj){
      
      var spreadsheet = SpreadsheetApp.openById(codeMoveFile.getId());
      
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
    }

    // --------------------
    // 2020.06.14
    function getHyperlinkFormula(url, label) {
      return "=HYPERLINK(\"" + url + "\",\"" + label + "\")";
    }
    // --------------------

    /**
     * Populate yearly stats Spreadsheet's Weekend Days sheet with references to
     * cells in the year's monthly data sheets.
     * @function updateYearlyStatsFile
     * @memberof MonthlyRun
     * @private
     * @param {Object} yearlyStatsFile
     * @param {string} codeMoveFileId
     * @param {number} month - 0 to 11
     * @param {string} yearMonthStr - YYYY-MM format
     * @returns {undefined}
     */
    function updateYearlyStatsFile(
      yearlyStatsFile, codeMoveFileId, month, yearMonthStr) {
      const spreadsheet = SpreadsheetApp.openById(yearlyStatsFile.getId());
      const weekendDaysSheet = spreadsheet.getSheetByName("Weekend Days");
      const yearlyStatsSheet = spreadsheet.getSheetByName("Imported Data");
      const row = month + 1;

      weekendDaysSheet.getRange("A1")
        .setValue("Weekend Days OHS Stats " + yearMonthStr.slice(0, 4));

      yearlyStatsSheet.getRange("A" + row)
        .setValue(yearMonthStr);

      // --------------------
      // 2020.06.14 - Set month date to link formula, that will also be used for month name in *OHS Stats sheet*
      var codeMoveSheetUrl = "https://docs.google.com/spreadsheets/d/" + codeMoveFileId;
      var codeMoveSheetLabel = yearMonthStr;
      var codeMoveSheetHyperlinkFormula = getHyperlinkFormula(
        codeMoveSheetUrl, codeMoveSheetLabel);

      yearlyStatsSheet.getRange("A" + row)
        .setFormula(codeMoveSheetHyperlinkFormula);

      // 2020.06.14 - Set Month Column Header per section (minus Miscellaneous) to Month Hyperlink on the *Weekend Days sheet*
      var colLetter = String.fromCharCode(66 + (row - 1));
      Logger.log("Column Letter: " + colLetter);
      var colNumbers = [2, 13, 24, 27, 32, 35];

      for (var i = 0; i < colNumbers.length; i++) {
        var cellLocation = colLetter + colNumbers[i];
        var cellValue = weekendDaysSheet.getRange(cellLocation).getValue();

        weekendDaysSheet.getRange(cellLocation)
          .setFormula(getHyperlinkFormula(codeMoveSheetUrl, cellValue));
      }
      // --------------------

      // Grand Totals
      yearlyStatsSheet.getRange("B" + row).setFormula("=IMPORTRANGE("
        + "\"https://docs.google.com/spreadsheets/d/"
        + codeMoveFileId
        + "\",\"Totals!B24:AH24\")"
      );

      // H26 PE/MD Code Move Total (calculated on Stats Weekend Stats sheet)

      // H27 Application Code Move Total (calculated on Stats Weekend Stats sheet)

      // H29 Magic Update Total
      yearlyStatsSheet.getRange("AI" + row).setFormula("=IMPORTRANGE("
        + "\"https://docs.google.com/spreadsheets/d/"
        + codeMoveFileId
        + "\",\"Totals!H29\")"
      );

      // H30 C/S Update Total
      yearlyStatsSheet.getRange("AJ" + row).setFormula("=IMPORTRANGE("
        + "\"https://docs.google.com/spreadsheets/d/"
        + codeMoveFileId
        + "\",\"Totals!H30\")"
      );

      // H31 Expanse Update Total
      yearlyStatsSheet.getRange("AK" + row).setFormula("=IMPORTRANGE("
        + "\"https://docs.google.com/spreadsheets/d/"
        + codeMoveFileId
        + "\",\"Totals!H31\")"
      );

      // H34 Ring Deletion Total
      yearlyStatsSheet.getRange("AL" + row).setFormula("=IMPORTRANGE("
        + "\"https://docs.google.com/spreadsheets/d/"
        + codeMoveFileId
        + "\",\"Totals!H34\")"
      );

      // P34 TEST Setup Total
      yearlyStatsSheet.getRange("AM" + row).setFormula("=IMPORTRANGE("
        + "\"https://docs.google.com/spreadsheets/d/"
        + codeMoveFileId
        + "\",\"Totals!P34\")"
      );

      // H33 HCIS Deletion Total
      yearlyStatsSheet.getRange("AN" + row).setFormula("=IMPORTRANGE("
        + "\"https://docs.google.com/spreadsheets/d/"
        + codeMoveFileId
        + "\",\"Totals!H33\")"
      );

      return undefined;
    }

    
    function maybeSendReminder(){
      var now = new Date();
      var intDay = now.getDate();
  
      if ( intDay > 6 )
        return true; // After the 6th or 6 days, we should only need to remind
      else
        return false; // Before the 6th day, there's a good chance we need to create, which will send an email, so don't also send a reminder email
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

      // --- SETUP ---
      const dateObj = (
        ((testYear !== undefined) && (testMonth !== undefined))
        ? new Date(testYear, testMonth)
        : new Date());
      const yearStr = dateObj.getFullYear().toString();
      const month = dateObj.getMonth();
      const monthStr = String(month + 1).toString().padStart(2, "0");
      const yearMonthStr = yearStr + "-" + monthStr;
      const monthName = dateObj.toLocaleString('default', { month: 'short' });
      
      // --- GET or CREATE Files ---
      var yearFolder = getYearFolder(yearStr);
      
      var yearlyStatsFile, yearFileFound;
      //var yearlyStatsFile = getYearlyStatsFile(yearStr, yearFolder);
      [yearlyStatsFile, yearFileFound] = getYearlyStatsFile(yearStr, yearFolder);
      
      //var codeMoveFile = getCodeMoveFile(yearFolder, yearMonthStr, dateObj);
      var codeMoveFile, yearMonthFileFound;
      [codeMoveFile, yearMonthFileFound] = getCodeMoveFile(yearFolder, yearMonthStr, dateObj);
      
      var codeMoveFileId = codeMoveFile.getId();
      
      // --- Update built files, maybe send email ---
      if ( yearMonthFileFound === false) { // We should really only be updating the file if it never existed before; once a month
        updateMonthlyStatsFile(codeMoveFile, dateObj);
        SendEmail.main(codeMoveFile,monthName,"testing","Weekend Code Move Count");
      }
     
      if ( !maybeSendReminder() || yearFileFound === false ) { // Sheet updates can happen every month, which would update the same file
        updateYearlyStatsFile(yearlyStatsFile, codeMoveFileId, month, yearMonthStr); // Updates YearlyStatsFile: set formulas (HYPERLINK, IMPORTRANGE)
        SendEmail.main(yearlyStatsFile, monthName, "testing","OHS Stat"); // (yearlyStatsFile, monthStr, flagTesting, monthlyStatsFile, typeEmail)  
      }
      /* jshint ignore:start */
      // See: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Operators/Destructuring_assignment#Array_destructuring
      /* jshint ignore:end */
      //[yearFolder, yearlyStatsFile] = getYearFolder(yearStr);
      //codeMoveFileId = getCodeMoveFile(yearFolder, yearMonthStr, dateObj); // Includes SendEmail for WeekendCodeMoveCount
      
      if ( maybeSendReminder() ) { // If past the 6th of the month, we assume that the file was created and only reminders need to be sent out
        SendEmail.main(codeMoveFile,monthName,"testing","Weekend Code Move Count");
        // Maybe send Rob reminder emails.
      }
      
      // --- Exit code ---
      return undefined;
    }

    return Object.freeze({
      main,
      updateYearlyStatsFile
    });

  }(DriveApp, PropertiesService, SpreadsheetApp));

/******************************************************************************/