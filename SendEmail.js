/*jslint browser:true, long:true, white:true*/
/*global
DriveApp, GmailApp, PropertiesService, StaffUtilities
*/
// eslint-disable-next-line no-unused-vars
const SendEmail = (
  function (DriveApp, GmailApp, PropertiesService) {

    
    function getEmailAddressesFromGoogleContactGroup(groupContact){
      return StaffUtilities.getObjArr(groupContact).map((userObj) => userObj.email);
    }
    
    function getRecipients(typeEmail, flagTesting){
      
      if ( flagTesting )
        return ["jeburns@meditech.com"];
        //return ["jeburns@meditech.com", "kgriffin@meditech.com"];
      
      if ( typeEmail == "Weekend Code Move Count" )
      {
        // groupEmail: weekend-code-move-tracking-test-group@meditech.com
        var groupEmail = PropertiesService.getScriptProperties().getProperty("groupEmail");
        return getEmailAddressesFromGoogleContactGroup(groupEmail); // returns []
      }
        
      if ( typeEmail == "OHS Stat" )
      {
        // OHS Stat sheet is only for management
        return ["rhomsey@meditech.com"];
      }
      
      return []; // Default of zero recipients
    }
    
    function getMessageBody(typeEmail, spreadsheetFile){
      
      var body = '<p>Click the following link to access the sheet: <a href="{file.getUrl}">{file.getName}</a></p>\
                 <div><br></div>';
      
      switch (typeEmail)
      {
        case "Weekend Code Move Count":
          body += '<div>Hi everyone,<br><br>This is your reminder to use the Weekend Code Move Count Spreadsheet! If you have any questions about the seet, please contact James E Burns or Kevin Griffin. Thanks</div>';
          break;
          
        case "OHS Stat":
          body += '<div>If you have any questions, please contact James E Burns or Kevin Griffin. </div>';
          break
      }
      
      body = body.replace(/\{file.getName\}/g, spreadsheetFile.getName())
                 .replace(/\{file.getUrl\}/g, spreadsheetFile.getUrl());
      
      return body;
      
    }
    
    function main(spreadsheetFile, monthStr, flagTesting, typeEmail) {

      if ( (typeEmail !== "OHS Stat") && (typeEmail !== "Weekend Code Move Count") )
        return; // Abort if there is no email type
      
      
      var userEmailAddress = getRecipients(typeEmail, flagTesting);

      var recipients, subject, body, options;

      if (userEmailAddress.length > 0) {

        // Build Message Recipients
        recipients = (userEmailAddress.length == 1) ? userEmailAddress[0] : recipients = userEmailAddress.join(",");

        // Build Message Subject
        subject = (typeEmail == "OHS Stat") ? "REMINDER: Weekend Days OHS Stats for " + monthStr + " is available!" : "REMINDER: Weekend Code Move Count for " + monthStr + " is available!";
        //subject = 'MONTHLY: {file.getName} is now available for editing in Google Drive!';
        //subject = subject.replace(/\{file.getName\}/g, spreadsheetFile.getName());

        // Build Message Body
        body = getMessageBody(typeEmail, spreadsheetFile);
        body = body.replace(/\{getMonthInfo\}/g, monthStr);

        // Build Message Options
        options = {htmlBody: body};

        // ---- Use Gamil Service to send email(s) ----
        GmailApp.sendEmail(recipients, subject, body, options);
      }

    }
    return Object.freeze({
      main
    });

  }(DriveApp, GmailApp, PropertiesService));