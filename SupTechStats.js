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

    function pathwaysGetDataFormula(deliveriesSs, yearStr, nextYearStr) {
      return `=QUERY(IMPORTRANGE("\
      ${deliveriesSs.getUrl()}","Sheet1!A10:G"),\
      "Select Col1, Col2, Col3, Col5, Col6 Where\
      (Col5 >= date '${yearStr}-01-01' AND Col5 < date '${nextYearStr}-01-01')\
      OR\
      (Col6 >= date '${yearStr}-01-01' AND Col6 < date '${nextYearStr}-01-01')\
      AND\
      (dayOfWeek(Col5)=1 OR dayOfWeek(Col5)>=2 OR dayOfWeek(Col5)=6 OR\
      dayOfWeek(Col5)>=7 OR dayOfWeek(Col6)=1 OR dayOfWeek(Col6)>=2 OR\
      dayOfWeek(Col6)=6 OR dayOfWeek(Col6)>=7)")`;
    }

    function pathwaysDisplayDataFormula(tabArr, yearNum, monthNum, endOfMonth) {
      return `=IFERROR(ISNA(QUERY('${tabArr[0]}'!A3:F,"Select * WHERE\
        year(D)=${yearNum}\
        AND (D >= date '${yearNum}-${monthNum}-1'\
        AND D <= date '${yearNum}-${monthNum}-${endOfMonth}')\
        AND (dayOfWeek(D)=1 OR dayOfWeek(D)=2 OR dayOfWeek(D)=6\
        OR dayOfWeek(D)=7)")),0,ROWS(QUERY('${tabArr[0]}'!A3:F,"Select *\
        WHERE year(D)=${yearNum}/
        AND (D >= date '${yearNum}-${monthNum}-01'\
        AND D <= date '${yearNum}-${monthNum}-${endOfMonth}')\
        AND (dayOfWeek(D)=1 OR dayOfWeek(D)=2\
        OR dayOfWeek(D)=6\
        OR dayOfWeek(D)=7)")))+IF(ISERR(QUERY('${tabArr[0]}'\
        !A3:F,"Select * WHERE year(E)=${yearNum}\
        AND (E >= date '${yearNum}-${monthNum}-01'\
        AND E <= date '${yearNum}-${monthNum}-${endOfMonth}')\
        AND (dayOfWeek(E)=1 OR dayOfWeek(E)=2 OR dayOfWeek(E)=6\
        OR dayOfWeek(E)=7)")),0,ROWS(QUERY('${tabArr[0]}'!A3:F,"Select\
        * WHERE year(E)=${yearNum} AND (E >= date '${yearNum}-${monthNum}-01'\
        AND E <= date '${yearNum}-${monthNum}-${endOfMonth}')\
        AND (dayOfWeek(E)=1 OR dayOfWeek(E)=2\
        OR dayOfWeek(E)=6 OR dayOfWeek(E)=7)"))),"")`;
    }

    function genericDisplayDataFormula(
      sheetName, yearNum, monthNum, endOfMonth) {
      return `=COUNTIFS('${sheetName}'!$A:$A,\
      ">="&Date(${yearNum},${monthNum},1),'${sheetName}'!$A:$A,\
      "<"&Date(${yearNum},${monthNum},${endOfMonth}))`;
    }

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
        .getRange("A3").setFormula(
          pathwaysGetDataFormula(deliveriesSs, yearStr, nextYearStr)
        );

      // construct formulas for current year
      Array.from({
        length: 12
        // eslint-disable-next-line max-statements
      }).forEach((ignored, index) => {
        const yearNum = Number(yearStr);
        const monthNum = index + 1;
        const endOfMonth = new Date(yearStr, monthNum, 0).getDate();
        const colLetter = String.fromCharCode(66 + index);
        // 6.x Pathway Code Deliveries
        indexSheet.getRange(`"${colLetter}:2"`).setFormula(
          pathwaysDisplayDataFormula(tabArr, yearNum, monthNum, endOfMonth)
        );
        // CSCT Messages
        `=COUNT(UNIQUE(FILTER('${tabArr[1]}'!$B:$B,'${tabArr[1]}'!$E:$E\
        >= Date(${yearNum},${monthNum},1),'${tabArr[1]}'!$E:$E\
        <= Date(${yearNum},${monthNum},${endOfMonth}))))`;
        // Data Recoveries
        genericDisplayDataFormula(tabArr[2], yearNum, monthNum, endOfMonth);
        // Development Projects (CSTS)
        genericDisplayDataFormula(tabArr[3], yearNum, monthNum, endOfMonth);
        // Health Check
        genericDisplayDataFormula(tabArr[4], yearNum, monthNum, endOfMonth);
        // Health Check - Resolution
        `=SUMIFS('${tabArr[5]}'!E:E,'${tabArr[5]}'!A:A,\
        ">="&Date(${yearNum},${monthNum},1),'${tabArr[5]}'!A:A,\
        "<"&Date(${yearNum},${monthNum},${endOfMonth}))`;
        // Infrastructure Projects
        genericDisplayDataFormula(tabArr[6], yearNum, monthNum, endOfMonth);
        // Large Scale Projects
        genericDisplayDataFormula(tabArr[7], yearNum, monthNum, endOfMonth);
        // LIVE Tasks Support
        genericDisplayDataFormula(tabArr[8], yearNum, monthNum, endOfMonth);
        // MaaS
        genericDisplayDataFormula(tabArr[9], yearNum, monthNum, endOfMonth);
        // Maintenance/Downtime Projects
        genericDisplayDataFormula(tabArr[10], yearNum, monthNum, endOfMonth);
        // Scheduled Projects
        genericDisplayDataFormula(tabArr[11], yearNum, monthNum, endOfMonth);
        // Stipend/Non Stipend
        genericDisplayDataFormula(tabArr[12], yearNum, monthNum, endOfMonth);
        // Tech Code Moves
        genericDisplayDataFormula(tabArr[13], yearNum, monthNum, endOfMonth);
        // Updates Supported (MG)
        `=COUNTIFS('${tabArr[14]}'!$D:$D,\
        ">="&Date(${yearNum},${monthNum},1),'${tabArr[14]}'!$D:$D,\
        "<"&Date(${yearNum},${monthNum},${endOfMonth}),'${tabArr[14]}'!$G:$G,\
        "Magic",'${tabArr[14]}'!$K:$K,"Yes")`;
        // Updates Supported (CS)
        `=COUNTIFS('${tabArr[15]}'!$D:$D,\
        ">="&Date(${yearNum},${monthNum},1),'${tabArr[15]}'!$D:$D,\
        "<"&Date(${yearNum},${monthNum},${endOfMonth}),'${tabArr[15]}'!$G:$G,\
        "CS",'${tabArr[15]}'!$K:$K,"Yes")`;
        // Updates Supported (Exp)
        `=COUNTIFS('${tabArr[16]}'!$D:$D,\
        ">="&Date(${yearNum},${monthNum},1),'${tabArr[16]}'!$D:$D,\
        "<"&Date(${yearNum},${monthNum},${endOfMonth}),'${tabArr[16]}'!$G:$G,\
        "6.08",'${tabArr[16]}'!$K:$K,"Yes")+COUNTIFS('${tabArr[16]}'!$D:$D,\
        ">="&Date(${yearNum},${monthNum},1),'${tabArr[16]}'!$D:$D,\
        "<"&Date(${yearNum},${monthNum},${endOfMonth}),'${tabArr[16]}'!$G:$G,\
        "6.15",'${tabArr[16]}'!$K:$K,"Yes")+COUNTIFS('${tabArr[16]}'!$D:$D,\
        ">="&Date(${yearNum},${monthNum},1),'${tabArr[16]}'!$D:$D,\
        "<"&Date(${yearNum},${monthNum},${endOfMonth}),'${tabArr[16]}'!$G:$G,\
        "Expanse",'${tabArr[16]}'!$K:$K,"Yes")`;
        // UWI Code Moves
        genericDisplayDataFormula(tabArr[17], yearNum, monthNum, endOfMonth);
      });

      return undefined;
    }

    return Object.freeze({
      //   main,
      yearlyInit
    });

  }(DriveApp, PropertiesService, SitesApp, SpreadsheetApp));
