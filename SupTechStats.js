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
      Array.from({
        length: 12
        // eslint-disable-next-line max-statements
      }).forEach((ignored, index) => {
        const yearNum = Number(yearStr);
        const monthNum = index + 1;
        const endOfMonth = new Date(yearStr, monthNum, 0).getDate();
        // TODO: add this to a range on ss
        `=IFERROR(ISNA(QUERY('${tabArr[0]}'!A3:F,"Select * WHERE\
        year(D)=${yearNum}\
        AND (D >= date '${yearNum}-${monthNum}-1'\
        AND D <= date '${yearNum}-${monthNum}-${endOfMonth}')\
        AND (dayOfWeek(D)=1 OR dayOfWeek(D)=2 OR dayOfWeek(D)=6\
        OR dayOfWeek(D)=7)")),0,ROWS(QUERY('${tabArr[0]}'!A3:F,"Select *\
        WHERE year(D)=${yearNum} AND (D >= date '${yearNum}-${monthNum}-01'\
        AND D <= date '${yearNum}-${monthNum}-${endOfMonth}')\
        AND (dayOfWeek(D)=1 OR dayOfWeek(D)=2\
        OR dayOfWeek(D)=6 OR dayOfWeek(D)=7)")))+IF(ISERR(QUERY('${tabArr[0]}'\
        !A3:F,"Select * WHERE year(E)=${yearNum}\
        AND (E >= date '${yearNum}-${monthNum}-01'\
        AND E <= date '${yearNum}-${monthNum}-${endOfMonth}')\
        AND (dayOfWeek(E)=1 OR dayOfWeek(E)=2 OR dayOfWeek(E)=6\
        OR dayOfWeek(E)=7)")),0,ROWS(QUERY('${tabArr[0]}'!A3:F,"Select * WHERE\
        year(E)=${yearNum} AND (E >= date '${yearNum}-${monthNum}-01'\
        AND E <= date '${yearNum}-${monthNum}-${endOfMonth}')\
        AND (dayOfWeek(E)=1 OR dayOfWeek(E)=2\
        OR dayOfWeek(E)=6 OR dayOfWeek(E)=7)"))),"")`;
        //
        `=COUNT(UNIQUE(FILTER('CSCT Messages'!$B:$B,'CSCT Messages'!$E:$E >= Date(2021,1,1),'CSCT Messages'!$E:$E <= Date(2021,2,1))))`;
        `=COUNTIFS('Data Recoveries'!$A:$A, ">="&Date(2021,1,1),'Data Recoveries'!$A:$A,"<"&Date(2021,2,1))`;
        `=COUNTIFS('Development Projects (CSTS)'!$A:$A, ">="&Date(2021,1,1),'Development Projects (CSTS)'!$A:$A,"<"&Date(2021,2,1))`;
        `=COUNTIFS('Health Check'!$A:$A, ">="&Date(2021,1,1),'Health Check'!$A:$A,"<"&Date(2021,2,1))`;
        `=SUMIFS('Health Check'!E:E,'Health Check'!A:A,">="&Date(2021,1,1),'Health Check'!A:A,"<"&Date(2021,2,1))`;
        `=COUNTIFS('Infrastructure Projects'!$A:$A, ">="&Date(2021,1,1),'Infrastructure Projects'!$A:$A,"<"&Date(2021,2,1))`;
        `=COUNTIFS('Large Scale Projects'!$A:$A, ">="&Date(2021,1,1),'Large Scale Projects'!$A:$A,"<"&Date(2021,2,1))`;
        `=COUNTIFS('LIVE Tasks Support'!$A:$A, ">="&Date(2021,1,1),'LIVE Tasks Support'!$A:$A,"<"&Date(2021,2,1))`;
        `=COUNTIFS(MaaS!$A:$A, ">="&Date(2021,1,1),MaaS!$A:$A,"<"&Date(2021,2,1))`;
        `=COUNTIFS('Maintenance/Downtime Projects'!$A:$A, ">="&Date(2021,1,1),'Maintenance/Downtime Projects'!$A:$A,"<"&Date(2021,2,1))`;
        `=COUNTIFS('Scheduled Projects'!$A:$A, ">="&Date(2021,1,1),'Scheduled Projects'!$A:$A,"<"&Date(2021,2,1))`;
        `=COUNTIFS('Stipend/Non Stipend'!$A:$A, ">="&Date(2021,1,1),'Stipend/Non Stipend'!$A:$A,"<"&Date(2021,2,1))`;
        `=COUNTIFS('Tech Code Moves'!$A:$A, ">="&Date(2021,1,1),'Tech Code Moves'!$A:$A,"<"&Date(2021,2,1))`;
        `=COUNTIFS('Updates Supported'!$D:$D, ">="&Date(2021,1,1),'Updates Supported'!$D:$D,"<"&Date(2021,2,1),'Updates Supported'!$G:$G,"Magic",'Updates Supported'!$K:$K,"Yes")`;
        `=COUNTIFS('Updates Supported'!$D:$D, ">="&Date(2021,1,1),'Updates Supported'!$D:$D,"<"&Date(2021,2,1),'Updates Supported'!$G:$G,"CS",'Updates Supported'!$K:$K,"Yes")`;
        `=COUNTIFS('Updates Supported'!$D:$D, ">="&Date(2021,1,1),'Updates Supported'!$D:$D,"<"&Date(2021,2,1),'Updates Supported'!$G:$G,"6.08",'Updates Supported'!$K:$K,"Yes")+COUNTIFS('Updates Supported'!$D:$D, ">="&Date(2021,1,1),'Updates Supported'!$D:$D,"<"&Date(2021,2,1),'Updates Supported'!$G:$G,"6.15",'Updates Supported'!$K:$K,"Yes")+COUNTIFS('Updates Supported'!$D:$D, ">="&Date(2021,1,1),'Updates Supported'!$D:$D,"<"&Date(2021,2,1),'Updates Supported'!$G:$G,"Expanse",'Updates Supported'!$K:$K,"Yes")`;
        `=COUNTIFS('UWI Code Moves'!$A:$A, ">="&Date(2021,1,1),'UWI Code Moves'!$A:$A,"<"&Date(2021,2,1))`;
      });

      return undefined;
    }

    return Object.freeze({
      //   main,
      yearlyInit
    });

  }(DriveApp, PropertiesService, SitesApp, SpreadsheetApp));
