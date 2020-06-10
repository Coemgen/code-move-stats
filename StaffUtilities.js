/*jslint browser:true, devel:true, long:true, white:true*/
/*global AdminDirectory, GroupsApp*/

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

/**
 * @namespace Staff
 */

/**
 * The object type of the elements of the array returned by the getObjArr
 * function.
 * @typedef {object} UserData
 * @property {string} staffName - Family,Given MI
 * @property {string} staffEmail - mnemonic@mail.com
 */

// eslint-disable-next-line no-unused-vars
const Staff = (

  function (AdminDirectory, GroupsApp) {
    "use strict";

    /**
     * @function capitalizeFirstLetter
     * @memberof Staff
     * @private
     */
    function capitalizeFirstLetter(string) {
      return string.charAt(0).toUpperCase() + string.slice(1);
    }

    /**
     * Gets a Google user object for a given user ID.
     * @function getUser
     * @memberof Staff
     * @private
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
     * Takes a Google Groups email then returns an array of objects including
     * group members' names and email addresses.
     * @function getObjArr
     * @memberof! Staff
     * @public
     * @param {string} email - Google Group Email
     * @returns {UserData[]} - Array of UserData objects
     */
    // eslint-disable-next-line no-unused-vars
    function getObjArr(groupEmail) {
      const staffArr = GroupsApp.getGroupByEmail(groupEmail)
        .getUsers().map(function (user) {
          return user.getEmail();
        });

      return Object.freeze(staffArr.map(function (email) {
        const staffMember = getUser(email);
        const name = staffMember.name.familyName
          + ","
          + staffMember.name.givenName;
        return Object.freeze({
          name,
          email
        });
      }));
    }

    /**
     * Get a sorted array of staff names from an array of staff objects
     * @function getNameArr
     * @memberof! Staff
     * @public
     * @param {UserData[]} staffObjArr - Array of UserData objects
     * @returns {string[]} - A sorted array of name strings
     */
    // eslint-disable-next-line no-unused-vars
    function getNameArr(staffObjArr) {
      return Object.freeze(staffObjArr.map((staffObj) => staffObj.name).sort());
    }

    return Object.freeze({
      getObjArr,
      getNameArr
    });

  }(AdminDirectory, GroupsApp));