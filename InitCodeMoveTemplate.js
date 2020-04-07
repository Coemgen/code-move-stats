/*jslint long:true, white:true*/

"use strict";

/**
 * @file This script initializes the Code Move Counts Template Spreadsheet.
 * <p>Before running the script, values must be set for the following
 * {@linkcode https://developers.google.com/apps-script/guides/properties|
 * script properties}:
 * <ul>
 *  <li>codeMoveTemplateId</li>
 *  <li>groupEmail</li>
 * </ul>
 * <p>Run the script using the {@linkcode
 * https://developers.google.com/apps-script/guides/v8-runtime|V8 Runtime}.
 * @author Kevin Griffin <kevin.griffin@gmail.com>
 */

/* jshint ignore:start */
/**
 * Declare global variables to satisfy linter expectations for strict mode.
 * - note that re-declaring a var type global that already has a value does not
 * - affect its value.
 * - See: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Statements/var#Description
 */
/* jshint ignore:end */
/** @global */
var PropertiesService;
/** @global */
var SpreadsheetApp;
/** @global */
var getStaffNameArr;
/** @global */
var getStaffObjArr;

/******************************************************************************/

/** @constant
@type {string}
@default
*/
const FIRST_STAFF_ROW = 4;

/** @constant
@type {string}
@default
*/
const LAST_STAFF_ROW = 23;

/**
 * Gets the values from a spreadsheet sheet column
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
 * @param {object[]} staffObjArr - Array of {name,email} objects
 * @param {Object} spreadsheet
 * @returns {undefined}
 */
function addStaffSheets(staffObjArr, spreadsheet) {
  const nameEmailMatrix = staffObjArr.map(
    (staffObj) => [staffObj.name, staffObj.email]
  ).sort();

  nameEmailMatrix.forEach(function (nameEmailArr) {
    const name = nameEmailArr[0];
    const email = nameEmailArr[1];
    var sheet = spreadsheet.getSheetByName(name);

    // delete existing sheet
    if (sheet) {
      spreadsheet.deleteSheet(sheet);
    }

    sheet = spreadsheet.getSheetByName("Staff")
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
 * @param {string[]} platformArr - Array of platform type strings
 * @param {string[]} dirRingArr - Array of ring type strings
 * @param {string[]} peMdNonArr - Array of action type strings
 * @param {string[]} bundleTypesArr - Array of bundle action strings
 * @returns {string[][]} - [[dirRing,platform,peMedNon|bundleType],...]
 */
function buildHeaderMatrix(
  dirRingArr, platformArr, peMdNonArr, bundleTypesArr) {
  return dirRingArr.reduce(
    (acc, dirRing) => acc.concat(platformArr.reduce(
      (acc, platform) => acc.concat(
        // returns array format [[dirRing, ...], ...] for each platform
        peMdNonArr.map(
          (curVal) => [dirRing, platform].concat(curVal))
      ), []).concat(
      // returns array format [[dirRing, ...],...] for each bundleType
      bundleTypesArr.map(
        (bundleType) => [dirRing, "Bundle", bundleType])
    )), []);
}

/**
 * Insert staff member name into cell A of the member's row.
 * Build spreadsheet formula strings with staff member row numbers then
 * insert the formulas into the staff member's data cells.
 * @param {string[]} staffNameArr - Array of staff name strings
 * @param {Object} totalsSheet - Template sheet for staff member stats
 * @param {string[][]} headerMatrix - Arrays of [dirRing,platform,peMdNon]
 * @param {Object} spreadsheet - Template of Code Move Count spreadsheet
 * @returns {undefined}
 */
function populateStaffRows(
  staffNameArr, totalsSheet, headerMatrix) {
  const noOfRows = LAST_STAFF_ROW - FIRST_STAFF_ROW + 1;

  if (staffNameArr.length > noOfRows) {
    // throw an error
    throw "Staff list is too long for current configuration";
  }

  staffNameArr.forEach(function (name, index) {
    const row = FIRST_STAFF_ROW + index;

    totalsSheet.getRange("A" + row).setValue(name);
    totalsSheet.getRange("B" + row + ":AH" + row)
      .clearContent()
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
 * @param {Object} totalsSheet
 * @param {string[]} staffNameArr
 * @param {string} platform - Magic, Expanse, Client/Server
 * @param {string} action - What the programmer did
 * @param {string} cell - The cell where totals should display
 * @returns {undefined}
 */
function setUpdatesTotal(totalsSheet, staffNameArr, platform, action, cell) {
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
 * @param {Object} totalsSheet
 * @param {string[]} staffNameArr
 * @returns {undefined}
 */
function setRingDeletionsTotal(totalsSheet, staffNameArr) {
  const action = "Dir./Ring Deletion";
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
 * @param {Object} totalsSheet
 * @param {string[]} staffNameArr
 * @returns {undefined}
 */
function setTestSetupsTotal(totalsSheet, staffNameArr) {
  const action = "Dir./Ring Setup";
  const ring = "Test";
  const cell = "H34";
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
 * Initialize the Totals Template Sheet.
 * @returns {undefined}
 */
// eslint-disable-next-line no-unused-vars
function initializeTemplateMain() { // jshint ignore:line
  const spreadsheet = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties()
    .getProperty("codeMoveTemplateId")
  );
  const referencesSheet = spreadsheet.getSheetByName("References");
  const email = PropertiesService.getScriptProperties()
    .getProperty("groupEmail");
  const staffObjArr = getStaffObjArr(email);
  const staffNameArr = getStaffNameArr(staffObjArr);
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
  setRingDeletionsTotal(totalsSheet, staffNameArr);
  setTestSetupsTotal(totalsSheet, staffNameArr);

  return undefined;
}

/******************************************************************************/
