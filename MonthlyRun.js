/*jslint browser:true, long:true, white:true*/
/*global
DriveApp, FIRST_STAFF_ROW, PropertiesService, SendEmail, SitesApp,
SpreadsheetApp, SupTechStats
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

  function (
    DriveApp, PropertiesService, SitesApp, SpreadsheetApp) {
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
      const urlLink = yearlyStatsFile.getUrl();
      const urlName = yearStr + " OHS Stats";
      const values = [
        "<a href=\"" + urlLink + "\">" + urlName + "</a>"
      ];
      // remove any existing links for the current year
      ohsStatsListPage.getListItems()
        .filter(
          (row) => row.getValueByName("Spreadsheet Links")
          .match(/>([^<]+)/)[1] === urlName
        )
        .forEach((link) => link.deleteListItem());
      // add current year to list
      ohsStatsListPage.addListItem(values);

      return yearlyStatsFile;
    }

    /**
     * Creates and returns a new supervisor/tech stats file object.  Also adds 
     * the file's url to the associated Google Site.
     * @function addYearlySupTechFile
     * @memberof MonthlyRun
     * @private
     * @param {*} yearlySupTechTemplate 
     * @param {*} yearStr 
     * @param {*} yearFolder 
     * @returns {Object}
     */
    function addYearlySupTechFile(yearlySupTechTemplate, yearStr, yearFolder) {
      const yearlySupTechFile = DriveApp.getFileById(
          yearlySupTechTemplate.getId())
        .makeCopy(
          ("Weekend Supervisor/Tech Stats "
            + yearStr),
          yearFolder);
      // update Google Site
      const site = SitesApp.getSiteByUrl(
        PropertiesService.getScriptProperties().getProperty("googleSiteUrl")
      );
      const supTechLogsListPage = site.getChildByName("supervisor-tech-logs");
      const urlLink = yearlySupTechFile.getUrl();
      const urlName = "Weekend Supervisor/Tech Stats " + yearStr;
      const values = [
        "<a href=\"" + urlLink + "\">" + urlName + "</a>"
      ];
      // remove any existing links for the current year
      supTechLogsListPage.getListItems()
        .filter(
          (row) => row.getValueByName("Spreadsheet Links")
          .match(/>([^<]+)/)[1] === urlName
        )
        .forEach((link) => link.deleteListItem());
      // add current year to list
      supTechLogsListPage.addListItem(values);
      // initialize Yearly Sup/Tech formula with the current year.
      SupTechStats.yearlyInit(yearlySupTechFile, yearStr);

      return yearlySupTechFile;
    }

    /**
     * Returns a reference to the folder object, for the current year, and its
     * yearly OHS and Sup/Tech stats spreadsheets.  If the folder does not 
     * already exist, a new one will be created and populated with a yearly OHS
     * and Suppervisor/Tech stats spreadsheets.
     * @function getYearlyFolderAndFiles
     * @memberof MonthlyRun
     * @private
     * @param {string} yearStr
     * @returns {object[]}
     */
    function getYearlyFolderAndFiles(yearStr) {
      // get Yearly OHS Stats template ID
      const yearlyStatsTemplate = SpreadsheetApp.openById(
        PropertiesService.getScriptProperties()
        .getProperty("yearlyStatsTemplateId")
      );
      // get Yearly Supervisor/Tech Stats template ID
      const yearlySupTechTemplate = SpreadsheetApp.openById(
        PropertiesService.getScriptProperties()
        .getProperty("yearlySupTechTemplateId")
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
      // get yearly OHS Stats file
      let fileIterator = yearFolder.getFilesByName(
        "Weekend Days OHS Stat tracking information "
        + yearStr);
      const yearlyStatsFile = (
        (fileIterator.hasNext() === true)
        ? fileIterator.next()
        : addYearlyStatsFile(yearlyStatsTemplate, yearStr, yearFolder));
      // get yearly Supervisor/Tech Logs file
      fileIterator = yearFolder.getFilesByName(
        "Weekend Supervisor/Tech Stats "
        + yearStr);
      const yearlySupTechFile = (
        (fileIterator.hasNext() === true)
        ? fileIterator.next()
        : addYearlySupTechFile(yearlySupTechTemplate, yearStr, yearFolder));

      return [yearFolder, yearlyStatsFile, yearlySupTechFile];
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
      const urlLink = codeMoveFile.getUrl();
      const urlName = year + "-" + month;
      const urlStr = "<a href=\"" + urlLink + "\">" + urlName + "</a>";
      const values = [
        urlStr
      ];
      // remove any existing links for the current month
      codeMovePage.getListItems()
        .filter(
          (row) => row.getValueByName("Spreadsheet Links")
          .match(/>([^<]+)/)[1] === urlName
        )
        .forEach((link) => link.deleteListItem());
      // add current month to list
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

        // for each staff member update link from Totals sheet
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

    function getTechFormula(yearStr, month, rowNum, colLetter) {
      const year = Number(yearStr);
      return `=IF(TODAY()>=Date(${year},${month},1),`
        + `IFERROR('Imported Data'!${colLetter}${rowNum},0),"")`;
    }

    function addSupTechFormulas(yearlyStatsSs, yearStr) {
      const weekendDaysSheet = yearlyStatsSs.getSheetByName("Weekend Days");

      Array.from({
        length: 12
      }).forEach((ignored, index) => {
        const colLetter = String.fromCharCode(66 + index);
        const month = index + 1;
        // 30 'Imported Data'!B53 Magic Updates supported
        console.log(`${colLetter}:30`);
        Logger.log(`${colLetter}:30`);
        weekendDaysSheet.getRange(`${colLetter}:30`)
          .setFormula(getTechFormula(yearStr, month, 53, colLetter));
        // 31 'Imported Data'!B54 CS Updates supported
        weekendDaysSheet.getRange(`${colLetter}:31`)
          .setFormula(getTechFormula(yearStr, month, 54, colLetter));
        // 32 'Imported Data'!B55 Expanse Updates supported
        weekendDaysSheet.getRange(`${colLetter}:32`)
          .setFormula(getTechFormula(yearStr, month, 55, colLetter));
        // 35 'Imported Data'!B56 UWI code moves
        weekendDaysSheet.getRange(`${colLetter}:35`)
          .setFormula(getTechFormula(yearStr, month, 56, colLetter));
        // 36 'Imported Data'!B52 Tech code moves
        weekendDaysSheet.getRange(`${colLetter}:36`)
          .setFormula(getTechFormula(yearStr, month, 52, colLetter));
        // 42 'Imported Data'!B$49 Maintenance/Downtime projects
        weekendDaysSheet.getRange(`${colLetter}:42`)
          .setFormula(getTechFormula(yearStr, month, 49, colLetter));
        // 43 'Imported Data'!B$40 CSCT Messages
        weekendDaysSheet.getRange(`${colLetter}:43`)
          .setFormula(getTechFormula(yearStr, month, 40, colLetter));
        // 44 'Imported Data'!B$42 Development Projects
        weekendDaysSheet.getRange(`${colLetter}:44`)
          .setFormula(getTechFormula(yearStr, month, 42, colLetter));
        // 45 'Imported Data'!B$46 Large Scale Projects 
        weekendDaysSheet.getRange(`${colLetter}:45`)
          .setFormula(getTechFormula(yearStr, month, 46, colLetter));
        // 46 'Imported Data'!B$41 Data Recoveries
        weekendDaysSheet.getRange(`${colLetter}:46`)
          .setFormula(getTechFormula(yearStr, month, 41, colLetter));
        // 47 'Imported Data'!B$43 Health Check - Post Downtime 
        weekendDaysSheet.getRange(`${colLetter}:47`)
          .setFormula(getTechFormula(yearStr, month, 43, colLetter));
        // 48 'Imported Data'!B$44 Health Check Resolution
        weekendDaysSheet.getRange(`${colLetter}:48`)
          .setFormula(getTechFormula(yearStr, month, 44, colLetter));
        // 49 'Imported Data'!B$48 MaaS
        weekendDaysSheet.getRange(`${colLetter}:49`)
          .setFormula(getTechFormula(yearStr, month, 48, colLetter));
        // 50 'Imported Data'!B$47 LIVE tasks support
        weekendDaysSheet.getRange(`${colLetter}:50`)
          .setFormula(getTechFormula(yearStr, month, 47, colLetter));
        // 51 'Imported Data'!B$51 Stipend/Non Stipend
        weekendDaysSheet.getRange(`${colLetter}:51`)
          .setFormula(getTechFormula(yearStr, month, 51, colLetter));
        // 52 'Imported Data'!B$39 6.x pathway code deliveries
        weekendDaysSheet.getRange(`${colLetter}:52`)
          .setFormula(getTechFormula(yearStr, month, 39, colLetter));
        // 53 'Imported Data'!B$45 Infrastructure projects
        weekendDaysSheet.getRange(`${colLetter}:53`)
          .setFormula(getTechFormula(yearStr, month, 45, colLetter));
        // 54 'Imported Data'!B$50 Scheduled projects
        weekendDaysSheet.getRange(`${colLetter}:54`)
          .setFormula(getTechFormula(yearStr, month, 50, colLetter));
      });

      return undefined;
    }

    /**
     * Populate yearly stats Spreadsheet's Weekend Days sheet with references to
     * cells in the year's monthly data sheets.
     * @function updateYearlyStatsFile
     * @memberof MonthlyRun
     * @private
     * @param {Object} yearlyStatsFile
     * @param {Object} yearlySupTechFile
     * @param {Object} codeMoveFile
     * @param {number} month - 0 to 11
     * @param {string} yearMonthStr - "Weekend Code Move Count YYYY-MM" format
     * @param {string} yearStr
     * @returns {undefined}
     */
    function updateYearlyStatsFile(
      yearlyStatsFile, yearlySupTechFile, codeMoveFile,
      month, yearMonthStr, yearStr) {
      const spreadsheet = SpreadsheetApp.openById(yearlyStatsFile.getId());
      const importedDataSheet = spreadsheet.getSheetByName("Imported Data");
      const row = 2;
      const column = month + 2;
      const colChar = String.fromCharCode(66 + month);
      const formula = "=IMPORTRANGE(" + colChar + "2,\"Totals!AD1:AD36\")";

      spreadsheet.getSheetByName("Weekend Days")
        .getRange("A1")
        .setValue("Weekend Days OHS Stats " + yearMonthStr.slice(24, 28));

      importedDataSheet.getRange(row, column).setValue(codeMoveFile.getUrl());
      importedDataSheet.getRange(row + 1, column).setFormula(formula);
      // yearlySupTechFile to yearlyStats sheet
      importedDataSheet.getRange("B39").setFormula(
        `=IMPORTRANGE("${yearlySupTechFile.getUrl()}}","Index!B2:M19")`
      );

      // TODO: add formulas for OHS stats sheet tech cells
      // add formulas linking Weekend Days sheet sup/tech #'s to Imported Data
      addSupTechFormulas(spreadsheet, yearStr);

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
      let yearlySupTechFile = {};

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
      [
        yearFolder,
        yearlyStatsFile,
        yearlySupTechFile
      ] = getYearlyFolderAndFiles(yearStr);

      codeMoveFile = getCodeMoveFile(yearFolder, yearMonthStr, dateObj);

      updateYearlyStatsFile(
        yearlyStatsFile, yearlySupTechFile, codeMoveFile,
        month, yearMonthStr, yearStr
      );

      // TODO: remove this?
      SendEmail.main(codeMoveFile.getId(), yearStr, monthStr);

      return undefined;
    }

    return Object.freeze({
      main,
      updateYearlyStatsFile
    });

  }(DriveApp, PropertiesService, SitesApp, SpreadsheetApp));

/******************************************************************************/
