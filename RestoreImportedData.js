/*jslint long:true, white:true*/

"use strict";

/**
 * @file Driver for the updateYearlyStatsFile function.
 */
 
var DriveApp;
var PropertiesService;
var updateYearlyStatsFile;
var YEAR_TO_RESTORE = 2020;

// eslint-disable-next-line no-unused-vars
function restoreImportedDataMain() {
  const dataFolder = DriveApp.getFolderById(
    PropertiesService.getScriptProperties()
    .getProperty("dataFolderId")
  );
  const yearFolder = dataFolder.getFoldersByName(
    YEAR_TO_RESTORE
  ).next();
  const fileIterator = yearFolder.getFiles();
  var fileNameIdArr = [];
  var fileObj = {};
  var fileName = "";
  var yearlyStatsFile = yearFolder.getFilesByName(
    YEAR_TO_RESTORE + "-stats"
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

      fileName = fileNameId[0];
      updateYearlyStatsFile(yearlyStatsFile, fileId, index, fileName);

      return undefined;
    }
  );

  return undefined;
}
