/*jslint browser:true, long:true, white:true*/
/*global PropertiesService, SpreadsheetApp*/

/**
 * @file Defines the <code><b>InitStatsTemplate</b></code> module.  This module
 * initializes the Stats Template spreadsheet's <b>Weekend Days</b> sheet with
 * A1Notation references to its <b>Imported Data</b> sheet.
 * <p>Before using this module, values must be set for the following {@linkcode
 * https://developers.google.com/apps-script/guides/properties script
 * properties}:
 * <ul>
 *  <li><b>yearlyStatsTemplateId</b>&nbsp;&ndash;&nbsp;the spreadsheet id for
 *  the yearly stats template</li>
 * </ul>
 * @author Kevin Griffin <kevin.griffin@gmail.com>
 */

/**
 * @namespace InitStatsTemplate
 */

// eslint-disable-next-line no-unused-vars
const InitStatsTemplate = (

  function (PropertiesService, SpreadsheetApp) {
    "use strict";

    /**
     * Sets up links from the Imported Data sheet to the Weekend Days sheet
     * @function initCodeMoves
     * @memberof InitStatsTemplate
     * @private
     * @param {Object} statsTemplate - The spreadsheet object
     */
    function initCodeMoves(statsTemplate) {
      const sheet = statsTemplate.getSheetByName("Weekend Days");
      Array.from({
        length: 12
      }).forEach(
        function (ignore, index) {
          const letter = String.fromCharCode(66 + index);
          const column = 1 + index;
          sheet.getRange(letter + "3:" + letter + "11").setValues(
            [
              [
                "=IF('Imported Data'!D" + column
                + "<1, 0, 'Imported Data'!D" + column + ")"
              ],
              [
                "=IF('Imported Data'!O" + column
                + "<1, 0, 'Imported Data'!O" + column + ")"
              ],
              [
                "=IF('Imported Data'!Z" + column
                + "<1, 0, 'Imported Data'!Z" + column + ")"
              ],

              [
                "=IF('Imported Data'!G" + column
                + "<1, 0, 'Imported Data'!G" + column + ")"
              ],
              [
                "=IF('Imported Data'!R" + column
                + "<1, 0, 'Imported Data'!R" + column + ")"
              ],
              [
                "=IF('Imported Data'!AC" + column
                + "<1, 0, 'Imported Data'!AC" + column + ")"
              ],

              [
                "=IF('Imported Data'!J" + column
                + "<1, 0, 'Imported Data'!J" + column + ")"
              ],
              [
                "=IF('Imported Data'!U" + column
                + "<1, 0, 'Imported Data'!U" + column + ")"
              ],
              [
                "=IF('Imported Data'!AF" + column
                + "<1, 0, 'Imported Data'!AF" + column + ")"
              ]
            ]
          );
        }
      );
    }

    /**
     * Sets up spreadsheet formulas for the Weekend Days sheet PE/MD cells
     * @function initPeMd
     * @memberof InitStatsTemplate
     * @private
     * @param {Object} statsTemplate - The spreadsheet object
     */
    function initPeMd(statsTemplate) {
      const sheet = statsTemplate.getSheetByName("Weekend Days");
      Array.from({
        length: 12
      }).forEach(
        function (ignore, index) {
          const letter = String.fromCharCode(66 + index);
          const column = 1 + index;
          sheet.getRange(letter + "14:" + letter + "22").setValues(
            [
              [
                "=SUM('Imported Data'!B" + column
                + ", 'Imported Data'!C" + column + ")"
              ],
              [
                "=SUM('Imported Data'!M" + column
                + ", 'Imported Data'!N" + column + ")"
              ],
              [
                "=SUM('Imported Data'!X" + column
                + ", 'Imported Data'!Y" + column + ")"
              ],

              [
                "=SUM('Imported Data'!E" + column
                + ", 'Imported Data'!F" + column + ")"
              ],
              [
                "=SUM('Imported Data'!P" + column
                + ", 'Imported Data'!Q" + column + ")"
              ],
              [
                "=SUM('Imported Data'!AA" + column
                + ", 'Imported Data'!AB" + column + ")"
              ],

              [
                "=SUM('Imported Data'!H" + column
                + ", 'Imported Data'!I" + column + ")"
              ],
              [
                "=SUM('Imported Data'!S" + column
                + ", 'Imported Data'!T" + column + ")"
              ],
              [
                "=SUM('Imported Data'!AD" + column
                + ", 'Imported Data'!AE" + column + ")"
              ]

            ]
          );
        }
      );
    }

    /**
     * Sets up spreadsheet formulas for the Weekend Days sheet bundles cells
     * @function initBundles
     * @memberof InitStatsTemplate
     * @private
     * @param {Object} statsTemplate - The spreadsheet object
     */
    function initBundles(statsTemplate) {
      const sheet = statsTemplate.getSheetByName("Weekend Days");
      Array.from({
        length: 12
      }).forEach(
        function (ignore, index) {
          const letter = String.fromCharCode(66 + index);
          const column = 1 + index;
          sheet.getRange(letter + "25:" + letter + "26").setValues(
            [
              [
                "=SUM('Imported Data'!K" + column
                + ",'Imported Data'!V" + column
                + ",'Imported Data'!AG" + column + ")"
              ],
              [
                "=SUM('Imported Data'!L" + column
                + ",'Imported Data'!W" + column
                + ",'Imported Data'!AH" + column + ")"
              ]
            ]
          );
        }
      );
    }

    /**
     * Sets up spreadsheet formulas for the Weekend Days sheet bundles cells
     * @function initAddToShipping
     * @memberof InitStatsTemplate
     * @private
     * @param {Object} statsTemplate - The spreadsheet object
     */
    function initAddToShipping(statsTemplate) {
      const sheet = statsTemplate.getSheetByName("Weekend Days");
      Array.from({
        length: 12
      }).forEach(
        function (ignore, index) {
          const letter = String.fromCharCode(66 + index);
          const column = 1 + index;
          sheet.getRange(letter + "28").setValues(
            [
              [
                "=IF('Imported Data'!AO" + column
                + "<1, 0, 'Imported Data'!AO" + column
                + ")"
              ]
            ]
          );
        }
      );
    }

    /**
     * Sets up spreadsheet formulas for the Weekend Days sheet updates cells
     * @function initUpdates
     * @memberof InitStatsTemplate
     * @private
     * @param {Object} statsTemplate - The spreadsheet object
     */
    function initUpdates(statsTemplate) {
      const sheet = statsTemplate.getSheetByName("Weekend Days");
      Array.from({
        length: 12
      }).forEach(
        function (ignore, index) {
          const letter = String.fromCharCode(66 + index);
          const column = 1 + index;
          sheet.getRange(letter + "30:" + letter + "32").setValues(
            [
              // AI Magic Updates total
              [
                "=IF('Imported Data'!AI" + column
                + "<1, 0, 'Imported Data'!AI" + column + ")"
              ],
              // AJ C/S Updates total
              [
                "=IF('Imported Data'!AJ" + column
                + "<1, 0, 'Imported Data'!AJ" + column + ")"
              ],
              // AK Exp Updates total
              [
                "=IF('Imported Data'!AK" + column
                + "<1, 0, 'Imported Data'!AK" + column + ")"
              ]
            ]
          );
        }
      );
    }

    /**
     * Sets up spreadsheet formulas for the Weekend Days sheet software support cells
     * @function initSoftwareSupport
     * @memberof InitStatsTemplate
     * @private
     * @param {Object} statsTemplate - The spreadsheet object
     */
    function initSoftwareSuppport(statsTemplate) {
      const sheet = statsTemplate.getSheetByName("Weekend Days");
      Array.from({
        length: 12
      }).forEach(
        function (ignore, index) {
          const letter = String.fromCharCode(66 + index);
          const column = 1 + index;
          sheet.getRange(letter + "38:" + letter + "40").setValues(
            [
              // AN HCIS Deletion total
              [
                "=IF('Imported Data'!AN" + column
                + "<1, 0, 'Imported Data'!AN" + column + ")"
              ],
              // AL Ring Deletion Total
              [
                "=IF('Imported Data'!AL" + column
                + "<1, 0, 'Imported Data'!AL" + column + ")"
              ],
              // AM Test Setup Total
              [
                "=IF('Imported Data'!AM" + column
                + "<1, 0, 'Imported Data'!AM" + column + ")"
              ]
            ]
          );
        }
      );
    }

    /**
     * Sets up spreadsheet formulas to link the <b>Imported Data</b> sheet to the
     * <b>Weekend Days</b> sheet and also sets up formulas to compile Weekend Days
     * sheet stats.
     * @function main
     * @memberof! InitStatsTemplate
     * @public
     * @param {Object} statsTemplate - Yearly Stats template spreadsheet object
     */
    // eslint-disable-next-line no-unused-vars
    function main(yearlyStatsFile) {

      // --------------------
      // 2020.06.14 - Allow ability to initize an existing yearly stats sheet (using RestoreImportedData) in addition to default init
      //              Also made change to: RestoreImportedData.gs
      /*
      const statsTemplate = SpreadsheetApp.openById(
        PropertiesService.getScriptProperties()
        .getProperty("yearlyStatsTemplateId")
      );
      */

      var statsTemplate;
      if (yearlyStatsFile) {
        statsTemplate = SpreadsheetApp.open(yearlyStatsFile);

      } else {
        statsTemplate = SpreadsheetApp.openById(
          PropertiesService.getScriptProperties()
          .getProperty("yearlyStatsTemplateId")
        );
      }
      // --------------------

      initCodeMoves(statsTemplate);
      initPeMd(statsTemplate);
      initBundles(statsTemplate);
      initAddToShipping(statsTemplate);
      initUpdates(statsTemplate);
      initSoftwareSuppport(statsTemplate);
    }

    return Object.freeze({
      main
    });

  }(PropertiesService, SpreadsheetApp));