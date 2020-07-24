/*jslint browser:true, long:true, white:true*/
/*global DriveApp, InitStatsTemplate, MonthlyRun, PropertiesService*/

/**
 * @file Code for relinking monthly totals to a new yearly stats spreadsheet.
 * Relinking should only be needed to repair data corruption.
 */

/**
 * @namespace RestoreImportedData
 */

// eslint-disable-next-line no-unused-vars
const zzRestoreImportedData = (

  function (DriveApp, zzMonthlyRun, PropertiesService) {
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
        function (fileNameId) {
          var fileId = fileNameId[1];
          // filename months are numbered from 1..12
          var monthOffset = Number(fileNameId[0].slice(5)) - 1;

          fileName = fileNameId[0];
          zzMonthlyRun.updateYearlyStatsFile(
            yearlyStatsFile, fileId, monthOffset, fileName
          );

          // ---------------
          // 2020.06.14
          // Initialize the Yearly Stats for each month - fixes broken formula references because a spreadsheet was deleted =SUM(!REF:!REF)
          InitStatsTemplate.main(yearlyStatsFile);
          // ---------------

          return undefined;
        }
      );

      return undefined;
    }

    return Object.freeze({
      main
    });
  }(DriveApp, MonthlyRun, PropertiesService));