/*global DriveApp, PropertiesService, SitesApp, SpreadsheetApp
 */

/**
 * @file Defines the <code><b>SupTechStats</b></code> module.  This module has
 * functions for setting formulas for the Supervisor/Tech Stats spreadsheet.
 */

/**
 * @namespace SupTechStats
 */

// eslint-disable-next-line no-unused-vars
const SupTechStats = (

  function (DriveApp, PropertiesService, SitesApp, SpreadsheetApp) {

    function yearlyInit(yearlySupTechFile, yearStr) {
      const spreadsheet = SpreadsheetApp.openById(yearlySupTechFile.getId());
      const indexSheet = spreadsheet.getSheetByName("Index");
      const deliveriesSs = SpreadsheetApp.openById(
        PropertiesService.getScriptProperties()
        .getProperty("deliveriesSpreadsheetId")
      );
      // get list of spreadsheet tab names sorted alphabetically
      const tabArr = spreadsheet.getSheets()
        .filter(
          (sheet) => {
            const name = sheet.getName();
            name !== "Index" && name !== "References" && name !== "Template";
          }).sort();
      const nextYearStr = `${parseInt(yearStr) + 1}`;

      // special case for 6.x Pathway Code Deliveries
      spreadsheet.getSheetByName("6.x Pathway Code Deliveries")
        .getRange("A3").setFormula(`=QUERY(IMPORTRANGE("\
      ${deliveriesSs.getUrl()}","Sheet1!A10:G"),\
      "Select Col1, Col2, Col3, Col5, Col6 Where\
      (Col5 >= date '${yearStr}-01-01' AND Col5 < date '${nextYearStr}-01-01')\
      OR\
      (Col6 >= date '${yearStr}-01-01' AND Col6 < date '${nextYearStr}-01-01')\
      AND\
      (dayOfWeek(Col5)=1 OR dayOfWeek(Col5)>=2 OR dayOfWeek(Col5)=6 OR\
      dayOfWeek(Col5)>=7 OR dayOfWeek(Col6)=1 OR dayOfWeek(Col6)>=2 OR\
      dayOfWeek(Col6)=6 OR dayOfWeek(Col6)>=7)")`);

      // construct formulas for current year

      return undefined;
    }

    return Object.freeze({
      //   main,
      yearlyInit
    });

  }(DriveApp, PropertiesService, SitesApp, SpreadsheetApp));
