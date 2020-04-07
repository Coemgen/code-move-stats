/*jslint long:true, white:true*/

"use strict";

/**
 * @file Mostly functions for getting user information.
 */

/* jshint ignore:start */
/**
 * Declare global variables to satisfy linter expectations for strict mode.
 * - note that re-declaring a var type global that already has a value does not
 * - affect its value.
 * - See: https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Statements/var#Description
 */
/* jshing ignore:end */
/** @global */
var AdminDirectory;
/** @global */
var GroupsApp;

function capitalizeFirstLetter(string) {
  return string.charAt(0).toUpperCase() + string.slice(1);
}

/**
 * Gets a Google user object for a given user ID.
 * @param {string} userkey A userID string e.g., kevin.griffin@gmail.com
 * @returns {object} JSON representing the Google user
 */
// eslint-disable-next-line no-unused-vars
function getUser(userKey) {
  const optionalArgs = {
    viewType: "domain_public"
  };
  const perPos = userKey.indexOf(".");
  const givenName = userKey.slice(0, perPos);
  const familyName = userKey.slice((perPos + 1), userKey.indexOf("@"));
  var userObj = {
    "name": {
      "familyName": capitalizeFirstLetter(familyName),
      "givenName": capitalizeFirstLetter(givenName)
    },
    "email": userKey
  };

  try {
    /*
     * To activate AdminDirectory:
     *   Resources (drop-down menu),
     *     Advanced Google Services,
     *     turn on Admin Directory API
     *
     * Authorize OAuth2 by viewing a user profile via this url
     * https://developers.google.com/admin-sdk/directory/v1/reference/users/get
     */
    userObj = AdminDirectory.Users.get(userKey, optionalArgs);
  } catch (e) {
    console.log(e);
  }

  return userObj;
}

/**
 * Takes a Google Groups email then returns an arry of objects including
 * group members' names and email addresses.
 * @param {string} email - Google Group Email
 * @returns {object[]} - Array of {name,email} objects
 */
// eslint-disable-next-line no-unused-vars
function getStaffObjArr(groupEmail) {

  const staffArr = GroupsApp.getGroupByEmail(groupEmail)
    .getUsers().map(function (user) {
      return user.getEmail();
    });

  return staffArr.map(function (staffEmail) {
    const staffMember = getUser(staffEmail);
    const staffName = staffMember.name.familyName
      + ","
      + staffMember.name.givenName;
    return {
      "name": staffName,
      "email": staffEmail
    };
  });
}

/**
 * Get a sorted array of staff names from a lost of staff objects
 * @param {object[]} staffObjArr - Array of {name, email} objects
 * @returns {string[]} - A sorted array of name strings
 */
// eslint-disable-next-line no-unused-vars
function getStaffNameArr(staffObjArr) {
  return staffObjArr.map((staffObj) => staffObj.name).sort();
}
