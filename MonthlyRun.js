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

      return yearlyStatsFile; // Returns File
    }

    /**
     * Update the newly created OHS Stat Tracking Information spreadsheet
     * to show monthly Supervisor/Tech Stats when the month has occurred.
     * Set formulas on 'Weekend Days sheet.'
     * @function editNewYearlyStatsFile
     * @param {File} yearlyStatsFile 
     * @returns {undefined}
     */
    function editNewYearlyStatsFile(yearlyStatsFile){
      var ss = SpreadsheetApp.open(yearlyStatsFile);
      var sheet = ss.getSheetByName("Weekend Days");
      
      var row = 42;
      var rows = 12; // 13
      var year = new Date().getFullYear();
      
      var tRow, tRowValues;
      var colChar, formula;
      var tRange, tMonth;
      
      /*
      // Set monthly header HYPERLINK - Can be added in Supervisor/Tech Monthly Run
      for (var i = 0; i < 12; i++)
      {
        colChar = String.fromCharCode(66 + i);
        tMonth = sheet.getRange( (colChar + "41") ).getValue();
        formula = '=HYPERLINK("'+yearlyStatsFile.getUrl()+'","'+tMonth+'")'; // =HYPERLINK('Imported Data'!B$2,"Jan")
        sheet.getRange( (colChar + "41") ).setFormula(formula);
      }
      */
      
      for (var i = 0; i < rows; i ++)
      {
        tRow = row + i;
        tRowValues = [];
        tRowValues[0] = []; // One row of data
        
        for (var x = 0; x < 12; x++)
        {
          formula = "";
          colChar = String.fromCharCode(66 + x);
          switch (tRow.toString())
          {
            case "42": // Maintenance/Downtime projects    
              formula = '=IF(TODAY()>=Date('+year+','+(x+1)+',1),\'Imported Data\'!'+colChar+'$49,"")';
              break;
              
            case "43": // CSCT Messages
              formula = '=IF(TODAY()>=Date('+year+','+(x+1)+',1),\'Imported Data\'!'+colChar+'$40,"")';
              break;
            
            case "44": // Development Projects (CSTS)
              formula = '=IF(TODAY()>=Date('+year+','+(x+1)+',1),\'Imported Data\'!'+colChar+'$42,"")';
              break;
            
            case "45": // Large Scale Projects
              formula = '=IF(TODAY()>=Date('+year+','+(x+1)+',1),\'Imported Data\'!'+colChar+'$46,"")';
              break;
              
            case "46": // Data Recoveries (once per shift)
              formula = '=IF(TODAY()>=Date('+year+','+(x+1)+',1),\'Imported Data\'!'+colChar+'$41,"")';
              break;
              
            case "47": // Health Check - Post Downtime
              formula = '=IF(TODAY()>=Date('+year+','+(x+1)+',1),\'Imported Data\'!'+colChar+'$43,"")';
              break;
              
            case "48": // Health Check Resolution (Number of issues fixed)
              formula = '=IF(TODAY()>=Date('+year+','+(x+1)+',1),\'Imported Data\'!'+colChar+'$44,"")';
              break;
              
            case "49": // MaaS (Number of tasks open/updated/fixed result of monitoring)
              formula = '=IF(TODAY()>=Date('+year+','+(x+1)+',1),\'Imported Data\'!'+colChar+'$48,"")';
              break;
              
            case "50": // LIVE tasks support (Volume hit 7 day window)
              formula = '=IF(TODAY()>=Date('+year+','+(x+1)+',1),\'Imported Data\'!'+colChar+'$47,"")';
              break;
              
            case "51": // Stipend/Non Stipend staff involved in working an issue
              formula = '=IF(TODAY()>=Date('+year+','+(x+1)+',1),\'Imported Data\'!'+colChar+'$51,"")';
              break;
              
            case "52": // 6.x pathway code deliveries (unique to weekend days)
              formula = '=IF(TODAY()>=Date('+year+','+(x+1)+',1),\'Imported Data\'!'+colChar+'$39,"")';
              break;
              
            case "53": // Infrastructure projects (footnote A)
              formula = '=IF(TODAY()>=Date('+year+','+(x+1)+',1),\'Imported Data\'!'+colChar+'$45,"")';
              break;
              
            case "54": // Scheduled projects (Footnote B)
              formula = '=IF(TODAY()>=Date('+year+','+(x+1)+',1),\'Imported Data\'!'+colChar+'$50,"")';
              break;             
          }
          tRowValues[0].push(formula);
        }
        
        // Set formula values
        tRange = "B" + tRow + ":M" + tRow; // Ex: B42:M42
        sheet.getRange(tRange).setFormulas(tRowValues);
      }
      
      return undefined;
    }

    /**
     * Adds spreadsheet link to Google Site page
     * @function addYearlyStatsFileToGoogleSite
     * @memberof MonthlyRun
     * @param {File} yearlyStatsFile
     * @returns {undefined}
     */
    function addYearlyStatsFileToGoogleSite(yearlyStatsFile, yearStr) {
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

      return undefined;
    }

    /**
     * Send Google Mail to Recipients, informing them that the new yearly stats
     * sheet is available.
     * @function sendYearlyStatsEmail
     * @memberof MonthlyRun
     * @param {File} yearlyStatsFile
     * @returns {undefined}
     */
    function sendYearlyStatsEmail(yearlyStatsFile, yearStr){
      // Send email when a new OHS Stats sheet is created
      const urlLink = yearlyStatsFile.getUrl();
      const urlName = yearStr + " OHS Stats";
      //const recipients = "eyip@meditech.com,jeburns@meditech.com,kgriffin@meditech.com,rhomsey@meditech.com";
      const recipients = "jeburns@meditech.com,kgriffin@meditech.com";
      const subject = urlName + " has been created!";
      const body = `Click the following link to access the current sheet:`
        + `${urlLink}`
        + `${urlName}`
        + "\n\nAttention:\n\n"
        + `This is your yearly message for the Weekend OHS Stats Spreadsheet! Thanks`;
      const htmlBody = `<p>`
        + `<div>This is your yearly message for the Weekend OHS Stats Spreadsheet!</div>`
        + `<div>Click the following link to access the current sheet: `
        + `<a href="${urlLink}">${urlName}</a>`
        + `</div>`
        + `<div>Thanks!</div>`
        + `</p>`;        
      const options = {
        "htmlBody": htmlBody
      };
      
      MailApp.sendEmail(recipients, subject, body, options);

      return undefined;
    }

    /**
     * Returns a reference to the folder object, for the current year.
     * Returns false if it does not exist.
     * @function getYearFolder
     * @memberof MonthlyRun
     * @private
     * @param {Folder} dataFolder (Parent folder)
     * @param {string} yearStr (Format: YYYY)
     * @returns {Folder} yearfolder or {Boolean} false
     */
    function getYearFolder(dataFolder, yearStr) {

      const folderIterator = dataFolder.getFoldersByName(
        yearStr
      ); // FolderIterator — A collection of all folders that are children of the current folder and have the given name.

      const yearFolder = (
        (folderIterator.hasNext() === true)
        ? folderIterator.next()
        : false
      );
      
      return yearFolder; // Returns Folder or False
    }

    /**
     * Returns a reference to the file object, yearly stats file for the current year,
     * Returns false if the file does not already exist.
     * @function getYearlyStatsFile
     * @memberof MonthlyRun
     * @private
     * @param {string} yearStr
     * @param {Folder} yearFolder 
     * @returns {File} yearlyStatsFile or False
     */
    function getYearlyStatsFile(yearStr, yearFolder) {
      const fileIterator = yearFolder.getFilesByName(
        "Weekend Days OHS Stat tracking information "
        + yearStr); // FileIterator — A collection of all files that are children of the current folder and have the given name.

      const yearlyStatsFile = (
        (fileIterator.hasNext() === true)
        ? fileIterator.next()
        : false
      ); // Returns File or False

      return yearlyStatsFile; // Returns File or False
    }

    /**
     * Creates and returns a new monthly code moves file object.  Also adds the code
     * moves file's url to the associated Google Site.
     * @function addCodeMoveFile
     * @memberof MonthlyRun
     * @param {Object} codeMoveTemplate
     * @param {string} codeMoveSheetName
     * @param {Object} yearFolder
     * @param {Object} dateObj
     * @returns {Object}
     */
    function addCodeMoveFile(codeMoveTemplate, codeMoveSheetName, yearFolder, dateObj) {
      const codeMoveFile = codeMoveTemplate.makeCopy(
        codeMoveSheetName, yearFolder
      );

      return codeMoveFile;
    }

    /**
     * Updates Google Site with new Code Move File info
     * @function addCodeMoveFileToGoogle
     * @memberof MonthlyRun
     * @param {File} codeMoveFile 
     * @param {Object} dateObj 
     * @returns {undefined}
     */
    function addCodeMoveFileToGoogle(codeMoveFile, dateObj){
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

      return undefined;
    }

    /**
     * Returns the File for the monthly data file object, for the current month.
     * Returns false if the file does not exist.
     * @function getCodeMoveFile
     * @memberof MonthlyRun
     * @private
     * @param {Object} yearFolder - the data folder
     * @param {string} codeMoveSheetName - used for naming the spreadsheet and sheets
     * @param {Object} dateObj - JavaScript date object for current month
     * @returns File or False
     */
    function getCodeMoveFile(yearFolder, codeMoveSheetName, dateObj) {
      const fileIterator = yearFolder.getFilesByName(
        codeMoveSheetName
      ); // Returns FileIterator — A collection of all files that are children of the current folder and have the given name.

      const codeMoveFile = (
        (fileIterator.hasNext())
        ? fileIterator.next()
        : false
      ); // Returns File or False

      return codeMoveFile; // Returns File or False
    }

    /**
     * Updates new Code Move File with values, formulas, and protection ranges
     * @function editNewCodeMoveFile
     * @memberof MonthlyRun
     * @param {File} codeMoveFile 
     * @param {Object} dateObj
     * @returns {undefined}
     */
    function editNewCodeMoveFile(codeMoveFile, dateObj){

      // edit new month spreadsheets
      let spreadsheet = SpreadsheetApp.openById(
        codeMoveFile.getId()
      ); // Returns Spreadsheet

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

        return undefined;
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
    function updateYearlyStatsFile(yearlyStatsFile, codeMoveFile, month, yearMonthStr) {
      const spreadsheet = SpreadsheetApp.openById(yearlyStatsFile.getId());
      const importedDataSheet = spreadsheet.getSheetByName("Imported Data");
      const row = 2;
      const column = month + 2;
      const colChar = String.fromCharCode(66 + month);
      const formula = "=IMPORTRANGE(" + colChar + "2,\"Totals!AD1:AD36\")"; // Needs to be updated if more data is captured in AD column

      spreadsheet.getSheetByName("Weekend Days")
        .getRange("A1")
        .setValue("Weekend Days OHS Stats " + yearMonthStr.slice(24, 28));

      importedDataSheet.getRange(row, column).setValue(codeMoveFile.getUrl());
      importedDataSheet.getRange(row + 1, column).setFormula(formula);

      return undefined;
    }

    /**
     * A monthly totals data entry spreadsheet will be created for the current month and linked to the yearly stats spreadsheet.
     * 1.) Gets the current year's data folder.
     *    If the folder does not exist, then it will be created
     * 2.) Gets the yearly OHS stats spreadsheet.
     *    If the spreadsheet does not exist:
     *      - it will be created
     *      - it will be linked to Google Site
     *      - an email will be sent notifying recipients of a new file
     * 3.) Gets the monthly weekend code move count spreadsheet
     *    If the spreadsheet does not exist:
     *      - it will be created
     *      - it will be linked to Google Site
     *      - it will be updated (values, formulas, protection ranges)
     * 4.) Fills in "Imported Data" hidden site on Yearly Stats spreadsheet to link data
     *     from Weekend Code Move Count to OHS Stats Tracking Information
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

      // Data Folder (Root)
      // \_Year Folder (YYYY)
      //   \_Files

      const dataFolder = DriveApp.getFolderById(
        PropertiesService.getScriptProperties().getProperty("dataFolderId")
      ); // Returns Folder or Throws a scripting exception if the folder does not exist or the user does not have permission to access it.

      yearFolder = getYearFolder(dataFolder, yearStr); // Returns Folder or False
      if ( yearFolder === false ) {
        yearFolder = dataFolder.createFolder(yearStr); // Returns Folder
      }

      const yearlyStatsTemplate = SpreadsheetApp.openById(
        PropertiesService.getScriptProperties().getProperty("yearlyStatsTemplateId")
      ); // Returns Spreadsheet

      yearlyStatsFile = getYearlyStatsFile(yearStr, yearFolder); // Returns File or False
      if ( yearlyStatsFile === false ) {
        yearlyStatsFile = addYearlyStatsFile(yearlyStatsTemplate, yearStr, yearFolder); // Returns File
        editNewYearlyStatsFile(yearlyStatsFile); // Later move to when yearly is a new file
        addYearlyStatsFileToGoogleSite(yearlyStatsFile, yearStr); // Returns undefined
        sendYearlyStatsEmail(yearlyStatsFile, yearStr); // Returns undefiend
      }

      const codeMoveTemplate = DriveApp.getFileById(
        PropertiesService.getScriptProperties()
        .getProperty("codeMoveTemplateId")
      );
      const codeMoveSheetName = yearMonthStr;

      codeMoveFile = getCodeMoveFile(yearFolder, codeMoveSheetName, dateObj); // Returns File or False
      if ( codeMoveFile === false ) {
        codeMoveFile = addCodeMoveFile(codeMoveTemplate, codeMoveSheetName, yearFolder, dateObj); // Returns File
        addCodeMoveFileToGoogle(codeMoveFile, dateObj) // Returns undefined
        editNewCodeMoveFile(codeMoveFile, dateObj); // Returns undefined
      }

      updateYearlyStatsFile(yearlyStatsFile, codeMoveFile, month, yearMonthStr);
      
      // Replaced line below as Emails are sent by other functions, called by triggers: function sendWeeklyReminder
      //SendEmail.main(codeMoveFile.getId(), yearStr, monthStr);

      return undefined;
    }

    return Object.freeze({
      main,
      updateYearlyStatsFile
    });

  }(DriveApp, PropertiesService, SitesApp, SpreadsheetApp));

/******************************************************************************/