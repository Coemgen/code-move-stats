/*jslint long:true, white: true*/

"use strict";

/**
 * @file Driver for the monthlyRunMain function
 */
 
/** @global */
var yearlyStatsMain;

// eslint-disable-next-line no-unused-vars
function test() {
  const numMonths = 1;
  const startYear = 2019;
  const startMonth = 10;
  const monthArr = Array.from({
    "length": numMonths
  });
  monthArr.forEach(function (ignore, index) {
    monthlyRunMain(startYear, startMonth + index);
  });

  return undefined;
}
