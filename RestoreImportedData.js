/*jslint browser:true, long:true, white:true*/
/*global DriveApp, MonthlyRun, PropertiesService*/

/**
 * @file Code for relinking monthly totals to a new yearly stats spreadsheet.
 * Relinking should only be needed to repair data corruption.
 */

/**
 * @namespace RestoreImportedData
 */

// eslint-disable-next-line no-unused-vars
const RestoreImportedData = (

  function (DriveApp, MonthlyRun, PropertiesService) {
    "use strict";

    /**
     * Relinks a deleted yearly stats spreadsheet to its associated monthly
     * totals spreadsheets.  This should only need to be run if there is a
     * problem with the link between yearly stats and monthly totals sheets.
     * @function main
     * @memberof! RestoreImportedData
     * @public
     * @param {(number|string)} yearToRestore - YYYY format
     */
    // eslint-disable-next-line no-unused-vars
    function main(yearToRestore) {
      const dataFolder = DriveApp.getFolderById(
        PropertiesService.getScriptProperties()
        .getProperty("dataFolderId")
      );
      const yearFolder = dataFolder.getFoldersByName(yearToRestore).next();
      const fileIterator = yearFolder.getFiles();
      var fileNameIdArr = [];
      var fileObj = {};
      var fileName = "";
      var yearlyStatsFile = yearFolder.getFilesByName(
        yearToRestore + "-stats"
      ).next();

      while (fileIterator.hasNext() === true) {
        fileObj = fileIterator.next();
        fileName = fileObj.getName();
        if (fileName.match(/^\d{4}-\d{2}$/) !== null) {
          fileNameIdArr.push([fileName, fileObj.getId()]);
        }
      }
      fileNameIdArr.sort().forEach(
        function (fileNameId, index) {
          var fileId = fileNameId[1];
          // TODO: extract month offset from month name
          var month = index;

          fileName = fileNameId[0];
          MonthlyRun.updateYearlyStatsFile(
            yearlyStatsFile, fileId, month, fileName
          );

          return undefined;
        }
      );

      return undefined;
    }

    return Object.freeze({
      main
    });
  }(DriveApp, MonthlyRun, PropertiesService));
