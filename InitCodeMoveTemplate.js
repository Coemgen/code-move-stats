/*jslint browser:true, long:true, white:true*/
/*global PropertiesService, SpreadsheetApp, StaffUtilities*/

/**
 * @file Defines the <code><b>InitCodeMoveTemplate</b></code> module.  This
 * module initializes the Code Move Counts Template Spreadsheet.
 * <p>Before using this module, values must be set for the following
 * {@linkcode https://developers.google.com/apps-script/guides/properties
 * script properties}:
 * <ul>
 *  <li><b>codeMoveTemplateId</b>&nbsp;&ndash;&nbsp;the spreadsheet id for the
 *  montly totals template</li>
 *  <li><b>groupEmail</b>&nbsp;&ndash;&nbsp;the Google Group email associated
 *  with this project</li>
 * </ul>
 * @author Kevin Griffin <kevin.griffin@gmail.com>
 */

/******************************************************************************/

/**
 * Monthly totals spreadsheet first staff member row.  Rows 1, 2, and 3 are
 * the totals spreadsheet header.
 * @constant
 * @type {number}
 * @default
 */
const FIRST_STAFF_ROW = 4;

/**
 * Monthly totals spreadsheet last staff member row.  Rows 24 and after are the
 * totals spreadsheet footer
 * @constant
 * @type {number}
 * @default
 */
const LAST_STAFF_ROW = 23;

/**
 * @namespace InitCodeMoveTemplate
 */

// eslint-disable-next-line no-unused-vars
const InitCodeMoveTemplate = (

  function (PropertiesService, SpreadsheetApp) {
    "use strict";

    /**
     * Gets the values from a spreadsheet sheet column
     * @function getColumnArray
     * @memberof InitCodeMoveTemplate
     * @private
     * @param {Object} sheetObj - The spreadsheet sheet object
     * @param {string} columnStr - The column to get in A1 or R1C1 notation
     * @returns {object[]} The array of column values
     */
    function getColumnArray(sheetObj, columnStr) {
      return sheetObj.getRange(columnStr)
        .getValues()
        .map((curVal) => curVal[0]).filter((curVal) => curVal);
    }

    /**
     * Initialize sheets for staff members
     * @function addStaffSheets
     * @memberof InitCodeMoveTemplate
     * @private
     * @param {object[]} staffObjArr - Array of {name,email} objects
     * @param {Object} spreadsheet
     * @returns {undefined}
     */
    function addStaffSheets(staffObjArr, spreadsheet) {
      const nameEmailMatrix = staffObjArr.map(
        (staffObj) => [staffObj.name, staffObj.email]
      ).sort();

      // delete existing user sheets
      spreadsheet.getSheets().forEach(
        function (sheet) {
          const sheetName = sheet.getName();
          if (sheetName !== "Totals"
            && sheetName !== "References"
            && sheetName !== "Staff") {
            spreadsheet.deleteSheet(sheet);
          }
        }
      );

      nameEmailMatrix.forEach(function (nameEmailArr) {
        const name = nameEmailArr[0];
        const email = nameEmailArr[1];
        var sheet = spreadsheet.getSheetByName("Staff")
          .copyTo(spreadsheet)
          .setName(name);

        sheet.getRange("B1:C1").setValue(email);

      });

      return undefined;
    }

    /**
     * Combines the header options from the References sheet into a matrix (an
     * array of an array of header strings) that will be used as an iterator for
     * building cell formulas.
     * @function buildHeaderMatrix
     * @memberof InitCodeMoveTemplate
     * @private
     * @param {string[]} platformArr - Array of platform type strings
     * @param {string[]} dirRingArr - Array of ring type strings
     * @param {string[]} peMdNonArr - Array of action type strings
     * @param {string[]} bundleTypesArr - Array of bundle action strings
     * @returns {string[][]} - [[dirRing,platform,peMedNon|bundleType],...]
     */
    function buildHeaderMatrix(
      dirRingArr, platformArr, peMdNonArr, bundleTypesArr) {
      return dirRingArr.reduce(
        (acc, dirRing) => [
          ...acc,
          ...platformArr.reduce(
            (acc, platform) => [...acc, ...peMdNonArr.map(
              (activity) => [dirRing, platform, activity])], []),
          ...bundleTypesArr.map(
            (bundleType) => [dirRing, "Bundle", bundleType])
        ], []);
    }

    /**
     * Insert staff member name into cell A of the member's row.
     * Build spreadsheet formula strings with staff member row numbers then
     * insert the formulas into the staff member's data cells.
     * @function populateStaffRows
     * @memberof InitCodeMoveTemplate
     * @private
     * @param {string[]} staffNameArr - Array of staff name strings
     * @param {Object} totalsSheet - Template sheet for staff member stats
     * @param {string[][]} headerMatrix - Arrays of [dirRing,platform,peMdNon]
     * @param {Object} spreadsheet - Template of Code Move Count spreadsheet
     * @returns {undefined}
     */
    function populateStaffRows(
      staffNameArr, totalsSheet, headerMatrix) {
      const noOfRows = LAST_STAFF_ROW - FIRST_STAFF_ROW + 1;
      const column = 1;
      const numColumns = 34;

      if (staffNameArr.length > noOfRows) {
        // throw an error
        throw "Staff list is too long for current configuration";
      }

      totalsSheet.getRange(FIRST_STAFF_ROW, column, noOfRows, numColumns)
      .clearContent();

      staffNameArr.forEach(function (name, index) {
        const row = FIRST_STAFF_ROW + index;

        totalsSheet.getRange("A" + row).setValue(name);
        totalsSheet.getRange("B" + row + ":AH" + row)
          .setFormulas(
            [
              headerMatrix.map(function (headerArr) {
                var formula = "";

                if (headerArr[2] === "10+ Changes") {
                  formula = "=COUNTIFS('"
                    + name + "'!C2:C,\"=" + headerArr[0] + "\","
                    + "'" + name + "'!G2:G,\"=Yes\","
                    + "'" + name + "'!B2:B,\"=Change Move\")";
                } else if (headerArr[2] === "Addl. Staff?") {
                  formula = "=COUNTIFS('"
                    + name + "'!C2:C,\"=" + headerArr[0] + "\","
                    + "'" + name + "'!H2:H,\"=Yes\","
                    + "'" + name + "'!B2:B,\"=Change Move\")";
                } else {
                  formula = "=COUNTIFS('"
                    + name + "'!C2:C,\"=" + headerArr[0] + "\","
                    + "'" + name + "'!D2:D,\"=" + headerArr[1] + "\","
                    + "'" + name + "'!E2:E,\"=" + headerArr[2] + "\","
                    + "'" + name + "'!B2:B,\"=Change Move\")";
                }
                return formula;
              })
            ]
          );
      });

      return undefined;
    }

    /**
     * Construct and return spreadsheet formula
     * @function getFooterFormula
     * @memberof InitCodeMoveTemplate
     * @private
     * @param {string[]} staffNameArr
     * @param {object} matchObj1 - {"key":"search term","cell":"A1:A2"}
     * @param {object} matchObj2 - {"key":"search term","cell":"A"}
     * @returns {string} - A spreadsheet formula
     */
    function getFooterFormula(staffNameArr, matchObj1, matchObj2) {
      var formulaStr = "=SUM(";

      // add users to formula string
      formulaStr += staffNameArr.reduce(function (acc, name, index) {
        var value = "";

        if (index > 0) {
          value += ",";
        }
        value += "COUNTIFS("
          + "'" + name + "'!" + matchObj1.cell + ",\"" + matchObj1.key + "\"";
        if (matchObj2 !== undefined) {
          value += ",'" + name + "'!" + matchObj2.cell
            + ",\"" + matchObj2.key + "\"";
        }
        value += ")";

        return acc + value;
      }, "");
      // terminate formula string
      formulaStr += ")";

      return formulaStr;
    }

    /**
     * Wrapper for calling getFooterFormula to get a formula string for adding to
     * a footer totals cell.
     * @function setUpdatesTotal
     * @memberof InitCodeMoveTemplate
     * @private
     * @param {Object} totalsSheet
     * @param {string[]} staffNameArr
     * @param {string} platform - Magic, Expanse, Client/Server
     * @param {string} action - What the programmer did
     * @param {string} cell - The cell where totals should display
     * @returns {undefined}
     */
    function setUpdatesTotal(
      totalsSheet, staffNameArr, platform, action, cell) {
      const matchObj1 = {
        "key": platform,
        "cell": "D2:D"
      };
      const matchObj2 = {
        "key": action,
        "cell": "B2:B"
      };
      const formulaStr = getFooterFormula(staffNameArr, matchObj1, matchObj2);

      totalsSheet.getRange(cell).setValue(formulaStr);

      return undefined;
    }

    /**
     * Wrapper for calling getFooterFormula to get a formula string for adding to
     * a footer totals cell.
     * @function setHcisDeletionsTotal
     * @memberof InitCodeMoveTemplate
     * @private
     * @param {Object} totalsSheet
     * @param {string[]} staffNameArr
     * @returns {undefined}
     */
    function setHcisDeletionsTotal(totalsSheet, staffNameArr) {
      const action = "HCIS Deletion";
      const cell = "H33";

      const matchObj1 = {
        "key": action,
        "cell": "B2:B"
      };
      const formulaStr = getFooterFormula(staffNameArr, matchObj1);

      totalsSheet.getRange(cell).setValue(formulaStr);

      return undefined;
    }

    /**
     * Wrapper for calling getFooterFormula to get a formula string for adding to
     * a footer totals cell.
     * @function setRingDeletionsTotal
     * @memberof InitCodeMoveTemplate
     * @private
     * @param {Object} totalsSheet
     * @param {string[]} staffNameArr
     * @returns {undefined}
     */
    function setRingDeletionsTotal(totalsSheet, staffNameArr) {
      const action = "Ring Deletion";
      const cell = "H34";

      const matchObj1 = {
        "key": action,
        "cell": "B2:B"
      };
      const formulaStr = getFooterFormula(staffNameArr, matchObj1);

      totalsSheet.getRange(cell).setValue(formulaStr);

      return undefined;
    }

    /**
     * Wrapper for calling getFooterFormula to get a formula string for adding to
     * a footer totals cell.
     * @function setTestSetupsTotal
     * @memberof InitCodeMoveTemplate
     * @private
     * @param {Object} totalsSheet
     * @param {string[]} staffNameArr
     * @returns {undefined}
     */
    function setTestSetupsTotal(totalsSheet, staffNameArr) {
      const action = "Dir./Ring Setup";
      const ring = "Test";
      const cell = "P34";
      const matchObj1 = {
        "key": action,
        "cell": "B2:B"
      };
      const matchObj2 = {
        "key": ring,
        "cell": "C2:C"
      };
      const formulaStr = getFooterFormula(staffNameArr, matchObj1, matchObj2);

      totalsSheet.getRange(cell).setValue(formulaStr);

      return undefined;
    }

    /**
     * Wrapper for calling getFooterFormula to get a formula string for adding to
     * a footer totals cell.
     * @function additionsToShipSource
     * @memberof InitCodeMoveTemplate
     * @private
     * @param {Object} totalsSheet
     * @param {string[]} staffNameArr
     * @returns {undefined}
     */
    function additionsToShipSource(totalsSheet, staffNameArr) {
      const action = "Add to Ship Source";
      const cell = "P36";
      const matchObj1 = {
        "key": action,
        "cell": "B2:B"
      };
      const formulaStr = getFooterFormula(staffNameArr, matchObj1);

      totalsSheet.getRange(cell).setValue(formulaStr);

      return undefined;
    }
    
    /**
     * Initialize the Monthly Totals Template Sheet with staff names and
     * spreadsheet formulas.  Run <code><b>InitCodeMoveTemplate.main()</b>
     * </code> when there are staffing changes.
     * @function main
     * @memberof! InitCodeMoveTemplate
     * @public
     * @returns {undefined}
     */
    // eslint-disable-next-line no-unused-vars
    function main() { // jshint ignore:line
      const spreadsheet = SpreadsheetApp.openById(
        PropertiesService.getScriptProperties()
        .getProperty("codeMoveTemplateId")
      );
      const referencesSheet = spreadsheet.getSheetByName("References");
      const email = PropertiesService.getScriptProperties()
        .getProperty("groupEmail");
      const staffObjArr = StaffUtilities.getObjArr(email);
      const staffNameArr = StaffUtilities.getNameArr(staffObjArr);
      const dirRingArr = getColumnArray(referencesSheet, "A:A");
      const platformArr = getColumnArray(referencesSheet, "B:B");
      const peMdNonArr = getColumnArray(referencesSheet, "C:C").slice(1);
      const bundleTypesArr = getColumnArray(referencesSheet, "E:E").slice(1);
      const totalsSheet = spreadsheet.getSheetByName("Totals");
      const headerMatrix = buildHeaderMatrix(
        platformArr, dirRingArr, peMdNonArr, bundleTypesArr);

      // add staff sheets
      addStaffSheets(staffObjArr, spreadsheet);

      // populate staff rows
      populateStaffRows(staffNameArr, totalsSheet, headerMatrix, spreadsheet);

      // populate footer cells
      setUpdatesTotal(totalsSheet, staffNameArr,
        "Magic", "Dir./Ring Update", "H29");
      setUpdatesTotal(totalsSheet, staffNameArr,
        "Client/Server", "Dir./Ring Update", "H30");
      setUpdatesTotal(totalsSheet, staffNameArr,
        "Expanse", "Dir./Ring Update", "H31");
      setHcisDeletionsTotal(totalsSheet, staffNameArr);
      setRingDeletionsTotal(totalsSheet, staffNameArr);
      setTestSetupsTotal(totalsSheet, staffNameArr);
      additionsToShipSource(totalsSheet, staffNameArr);

      return undefined;
    }

    return Object.freeze({
      main
    });
  }(PropertiesService, SpreadsheetApp));

/******************************************************************************/