/*jslint long:true, white:true*/

"use strict";

/**
 * @file This script initializes the Stats Template spreadsheet's "Weekend Days"
 * sheet with A1Notation references to its "Imported Data" sheet.
 * <p>Before running the script, values must be set for the following {@linkcode
 * https://developers.google.com/apps-script/guides/properties|script
 * properties}:
 * <ul>
 *  <li>yearlyStatsTemplateId</li>
 * </ul>
 * <p>Run this script with the {@linkcode
 * https://developers.google.com/apps-script/guides/v8-runtime|V8 Runtime}.
 * @author Kevin Griffin <kevin.griffin@gmail.com>
 */

var PropertiesService;
var SpreadsheetApp;

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
          ["=IF('Imported Data'!D" + column
            + "<1, 0, 'Imported Data'!D" + column + ")"
          ],
          ["=IF('Imported Data'!O" + column
            + "<1, 0, 'Imported Data'!O" + column + ")"
          ],
          ["=IF('Imported Data'!Z" + column
            + "<1, 0, 'Imported Data'!Z" + column + ")"
          ],

          ["=IF('Imported Data'!G" + column
            + "<1, 0, 'Imported Data'!G" + column + ")"
          ],
          ["=IF('Imported Data'!R" + column
            + "<1, 0, 'Imported Data'!R" + column + ")"
          ],
          ["=IF('Imported Data'!AC" + column
            + "<1, 0, 'Imported Data'!AC" + column + ")"
          ],

          ["=IF('Imported Data'!J" + column
            + "<1, 0, 'Imported Data'!J" + column + ")"
          ],
          ["=IF('Imported Data'!U" + column
            + "<1, 0, 'Imported Data'!U" + column + ")"
          ],
          ["=IF('Imported Data'!AF" + column
            + "<1, 0, 'Imported Data'!AF" + column + ")"
          ]
        ]
      );
    }
  );
}

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
          ["=SUM('Imported Data'!B" + column
            + ", 'Imported Data'!C" + column + ")"
          ],
          ["=SUM('Imported Data'!M" + column
            + ", 'Imported Data'!N" + column + ")"
          ],
          ["=SUM('Imported Data'!X" + column
            + ", 'Imported Data'!Y" + column + ")"
          ],

          ["=SUM('Imported Data'!E" + column
            + ", 'Imported Data'!F" + column + ")"
          ],
          ["=SUM('Imported Data'!P" + column
            + ", 'Imported Data'!Q" + column + ")"
          ],
          ["=SUM('Imported Data'!AA" + column
            + ", 'Imported Data'!AB" + column + ")"
          ],

          ["=SUM('Imported Data'!H" + column
            + ", 'Imported Data'!I" + column + ")"
          ],
          ["=SUM('Imported Data'!S" + column
            + ", 'Imported Data'!T" + column + ")"
          ],
          ["=SUM('Imported Data'!AD" + column
            + ", 'Imported Data'!AE" + column + ")"
          ]

        ]
      );
    }
  );
}

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
          ["=SUM('Imported Data'!K" + column
            + ",'Imported Data'!V" + column
            + ",'Imported Data'!AG" + column + ")"
          ],
          ["=SUM('Imported Data'!L" + column
            + ",'Imported Data'!W" + column
            + ",'Imported Data'!AH" + column + ")"
          ]
        ]
      );
    }
  );
}

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
          ["=IF('Imported Data'!AI" + column
            + "<1, 0, 'Imported Data'!AI" + column + ")"
          ],
          ["=IF('Imported Data'!AJ" + column
            + "<1, 0, 'Imported Data'!AJ" + column + ")"
          ],
          ["=IF('Imported Data'!AK" + column
            + "<1, 0, 'Imported Data'!AK" + column + ")"
          ]
        ]
      );
    }
  );
}

// eslint-disable-next-line no-unused-vars
function initStatsTemplateMain() {
  const statsTemplate = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties()
    .getProperty("yearlyStatsTemplateId")
  );
  initCodeMoves(statsTemplate);
  initPeMd(statsTemplate);
  initBundles(statsTemplate);
  initUpdates(statsTemplate);
}
