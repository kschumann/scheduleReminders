/*-----------INITIATION---------------*/
function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  try{
  var ui = SpreadsheetApp.getUi(); 
  ui.createAddonMenu()
  .addItem('Settings', 'openSettings') 
  .addToUi(); 
  //console.info(Session.getTemporaryActiveUserKey());
  } catch(e){
   // console.error("onOpen Error: " + e + "-" + Session.getTemporaryActiveUserKey());
  }
}

function clearAll(){
  removeTrigger();
  removePropId();
}

/*Creates EmailReminders sheet and defines template*/
function createTemplate(){
  try{
    var ss = SpreadsheetApp.getActive();
    var sheetId = ss.getId();
    if(!ss.getSheetByName('EmailReminders')){
      ss.insertSheet('EmailReminders', 0);
    }
    var reminderSheet = ss.getSheetByName('EmailReminders');
    var reminderSheetId = reminderSheet.getSheetId();
    PropertiesService.getUserProperties().setProperty('reminderSheetId', reminderSheetId);
    reminderSheetId = PropertiesService.getUserProperties().getProperty('reminderSheetId');
    ss.setActiveSheet(reminderSheet); 
    var headerRange = reminderSheet.getRange(1, 1, 1, 11);
    var headerContent = 
        [["Event Subject",
          "Organizational Grouping (optional)",
          "Send To",
          "cc (optional)",
          "Event Date (optional)",
          "First Reminder Date",
          "First Reminder Message",
          "Second Reminder Date (optional)",
          "Second Reminder Message (optional)",
          "First Reminder Sent On",
          "Second Reminder Sent On"
         ]];    
    var explanatoryNotes = [[
      "This will be the subject line of your email. (max 250 chars)",
      "This name can be included in the email body using a formula.  See example",
      "This is who the email will be sent to. Comma separated list of individual emails or email groups.",
      "Add anyone who should be cc'd in the reminder, such as yourself or some other admin. Separate multiple emails using commas.",
      "Date of the event.  Used only for your reference and can be included in the body of your email. See example message.",
      "On this day, the Add-On will send the first reminder message to recipients indicated. Make sure this is a valid date. If your spreadsheet is not set up for a USA locale, use dd/mm/yyyy format instead of mm/dd/yyy.",
      "Message that will be included in the body of the email.  Can use formulas as in example.",
      "Include a second reminder date if desired. Make sure this is a valid date. If your spreadsheet is not set up for a USA locale, use dd/mm/yyyy format instead of mm/dd/yyy.",
      "Include a second reminder message if desired.",
      "Date-Time that first email was sent.  Will be filled in by Add-On after successful send.",
      "Date-Time that second email was sent.   Will be filled in by Add-On after successful send."
    ]];
    headerRange.setValues(headerContent).setFontWeight('900').setBackground('#d9d9d9').setNotes(explanatoryNotes);
    reminderSheet.setColumnWidths(1, 11, 100);
    reminderSheet.setColumnWidth(1, 200);
    reminderSheet.setColumnWidth(3, 200);
    reminderSheet.setColumnWidth(4, 200);
    reminderSheet.setColumnWidth(7, 300);
    reminderSheet.setColumnWidth(9, 300);
    reminderSheet.setRowHeight(1, 40);
    reminderSheet.getRange(1, 1, 100, 11).setWrap(true);
    
    //Set Example -----------------
    var exampleExists = false;
    var exampleRange = reminderSheet.getRange(2,1,1,11);  
    var values = exampleRange.getValues();
    for(var i=0; i<values[0].length;i++){
      if(values[0][i]){
        exampleExists = true;
        break
      }
    }
    if(!exampleExists){//Make sure that first data row is only overwritten if not already populated, to ensure that user data is not lost.
      var exampleContent = 
          [["Bring baked goods to class",
            "Joyce Family",
            "joy@example.com, john@example.com",
            "admin@hostexample.com",
            "11/20/2021",
            "11/17/2021",
            "=\"Hello \" & B2 & \",\n Don\'t forget that your turn to bring the baked goods is coming up on \" & MONTH(E2) & \"/\" & DAY(E2) & \"/\" & YEAR(E2)",
            "11/18/2020",
            "=\"Hello \" & B2 & \",\n This is your last reminder to bring the baked goods on \" & MONTH(E2) & \"/\" & DAY(E2) & \"/\" & YEAR(E2)",
            "",
            ""
           ]];  
      exampleRange.setValues(exampleContent).setFontStyle('italic');
    }
    console.info(sheetId + "-Template Created for sheet Id:" + reminderSheetId);    
  } catch (e){
    console.error(sheetId + '-createTemplate failed, error: ' + e);
  }

}

/*Set Trigger for Scheduled Run*/
function activateTrigger(hour) {
  var ss = SpreadsheetApp.getActive();
  var timeZone = ss.getSpreadsheetTimeZone();
  setTrigger(hour,timeZone);
  setPropertyId(hour);
}

/*-----------INTERFACES---------------*/

function openSettings(){
  var html = HtmlService.createHtmlOutputFromFile('Setup')
      .setTitle('Set Up Reminder Spreadsheet')
      .setWidth(500).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Manage Your Reminder Spreadsheet');
}

/*-----------TRIGGERS---------------*/
/*Sets Send Reminder Trigger and stores trigger Id*/

function setTrigger(hour,timeZone){
  try{
    var trigger = doesTriggerExist('sendReminders');
    if(!trigger){
      ScriptApp.newTrigger('sendReminders').timeBased().inTimezone(timeZone).everyDays(1).atHour(hour).create();
    }
    console.info("TriggerSet for the following hour: " + hour);
  } catch(e){
    console.error('setTrigger failed, error: ' + e);
  }
}

function removeTrigger(){//removes trigger if exists;
  try{
    var trigger = doesTriggerExist('sendReminders');
    if(trigger){
      ScriptApp.deleteTrigger(trigger);
    }
    console.info("Trigger Removed");
  } catch(e){
    console.error('removeTrigger failed, error: ' + e);
  }
}

function doesTriggerExist(functionName){//returns trigger if it exists, else null
  try{
    var trigger = null;
    var triggers = ScriptApp.getProjectTriggers();
    for(var i=0; i<triggers.length; i++){
      if(triggers[i].getHandlerFunction() == functionName){
        trigger = triggers[i];
        break
      }
    } 
    return trigger;
  } catch (e){
    console.error('doesTriggerExist failed, error: ' + e);
  }
}

function doesSendRemindersTriggerExist(){
  var trigger = doesTriggerExist('sendReminders');
  trigger = trigger ? true : false;
  return trigger;
}


// PropertiesService.getUserProperties().deleteProperty('sendRemindersTriggerId');
/*-----------USER PROPERTIES---------------*/
/*Adds Id of active doc as a user property - needed for trigger to call the correct doc*/
function setPropertyId(hour){
  try{
    var docId = SpreadsheetApp.getActive().getId();
    var triggerTime = formatTime(hour);
    PropertiesService.getUserProperties().setProperty('reminderDocId', docId);
    PropertiesService.getUserProperties().setProperty('triggerTime', triggerTime);
    return docId;
  } catch (e) {
    console.error(docId+'-setPropertyId failed, error: ' + e);
  }
}

/*Removes property Id*/
function removePropId(){
  try{
    var userKeys = PropertiesService.getUserProperties().getKeys();
    for(var i = 0; i<userKeys.length;i++){
      if(userKeys[i] == 'reminderDocId' || userKeys[i] == 'triggerTime' || userKeys[i] == 'reminderSheetId'){
        PropertiesService.getUserProperties().deleteProperty(userKeys[i]);
      } 
    }
  } catch (e){
    console.error('removePropId failed, error: ' + e);
  }
}

/*Checks whether user already has property set by script*/
function doesPropIdExist(){
  var userKeys = PropertiesService.getUserProperties().getKeys();
  var propIdExists = false;
  for(var i = 0; i<userKeys.length;i++){
    if(userKeys[i] == 'reminderDocId'){
      var propIdExists = true;
    } 
  }
  return propIdExists;
}

/*Gets document info for setup UI*/
function getDocProps(){
  try{
    var trigger = doesTriggerExist('sendReminders');
    trigger = trigger ? true : false;
    var savedDocId = PropertiesService.getUserProperties().getProperty('reminderDocId');
    try{
      var ss =  SpreadsheetApp.openById(savedDocId)
    } catch(e){
      PropertiesService.getUserProperties().setProperty('reminderDocId','');
      savedDocId = '';
    }
    if(trigger){
      var triggerTime = PropertiesService.getUserProperties().getProperty('triggerTime');
      if(!triggerTime){
        var triggerTime = "2:00am";   
      }
      
    } else { 
      var triggerTime = "None";
    }
    var currentDocId = SpreadsheetApp.getActive().getId();
    var emailSheet = SpreadsheetApp.getActive().getSheetByName('EmailReminders');
    var usingSaved = savedDocId == currentDocId ? true : false;  
    if (savedDocId){
      var ss =  SpreadsheetApp.openById(savedDocId);
      var savedSsName = ss.getName();
    } else {
      var savedSsName = "";}
    var docInfo = 
        {
          "savedDocId":savedDocId,
          "savedSsName":savedSsName,
          "usingSaved":usingSaved,
          "trigger":trigger,
          "triggerTime":triggerTime
        };
    return docInfo;
  } catch (e){
    console.error(savedDocId + '-getDocProps failed, error: ' + e);
  }
}



/*-----------SEND REMINDERS---------------*/
/*parse the designated sheet and send out emails*/
function sendReminders() {  
  try{
    var docId = PropertiesService.getUserProperties().getProperty('reminderDocId');
    var sheetId = PropertiesService.getUserProperties().getProperty('reminderSheetId');
   var ss = SpreadsheetApp.openById(docId);
   // console.log(docId + "-SendReminders Got Spreadsheet ID");
  //  console.log(docId + "-SendReminders Got Sheet ID: " + sheetId);
    var sheet = ss.getSheetByName('EmailReminders');
    if(sheet){
     // console.log(docId + " - Sheet Assigned by Name.");
    }
    if(sheetId){
      var sheets = ss.getSheets();
      for(var i=0; i<sheets.length; i++){
        if(sheets[i].getSheetId()==sheetId){
          sheet = sheets[i];
          var sheetName = sheet.getName();
         // console.log(docId + " - Sheet withId Assigned.");
        }
      }    
    }
    
    if(!sheet){
      throw "The Add-On was not able to locate the sheet that was initially set up.";
    } else{
      var sheetName = sheet.getName();
      //console.log(docId + "- SendReminders Sheet Name: " + sheetName);      
      var numbRows = sheet.getLastRow()-1;
      var allValues = sheet.getRange(2, 1, numbRows, 9).getValues();
      var firstDateValues = sheet.getRange(2, 6, numbRows, 1).getValues();
      var secondDateValues = sheet.getRange(2, 8, numbRows, 1).getValues();
      
      var timeZone = ss.getSpreadsheetTimeZone();
      if(!timeZone){
        throw "A timezone was not found in your GoogleSheet. Check your timezone settings under File=>Spreadsheet Settings";
      }
      var dateTime =   new Date();
      var todayFormatted = Utilities.formatDate(dateTime, timeZone, "YYYY-MM-dd");
      var scriptTimeZone = Session.getScriptTimeZone();
     // console.log(docId + "-todayformatted: " + todayFormatted + "-" + scriptTimeZone + ":scriptTime/sheetTime:" + timeZone); 
      
      /*Create Arrays of dates to loop through*/ 
      var firstReminderDates = createArray(firstDateValues);
      var secondReminderDates = createArray(secondDateValues);
    }
  } catch(e){
    console.error(docId + '-sendReminders-Setup failed, error: ' + e);
    catchError(e);
    clearAll();
    return;
  }
  if(sheet && numbRows>0){ 
    var numbEmailsSent = 0;
    /*Send email if today is the first reminder date*/
    try{
      if(!firstReminderDates){
        throw "No First Reminder Dates Found.";
      }
      for(var j = 0; j < firstReminderDates.length; j++){ 
        if(firstReminderDates[j]){
          var firstDateFormatted ="";
          try{
          var firstDateFormatted = Utilities.formatDate(firstReminderDates[j], timeZone, "YYYY-MM-dd");
          } catch(e){
           // console.error(e);
          }
          if(firstDateFormatted == todayFormatted){
            MailApp.sendEmail(
              allValues[j][2], //to:
              allValues[j][0], //subject:
              allValues[j][6], //body:
              {
                cc:allValues[j][3]
              });  
            var sentDate = new Date();
            sheet.getRange(j+2, 10).setValue(sentDate); 
            numbEmailsSent++;
            console.info("Email Sent - 1")
          }
        }
      }
    } catch(e){
      console.error(docId +'-sendReminders-1stDate failed, error: ' + e);
      catchError(e);
      clearAll();
      return;
    }
    
    /*Send email if today is the second reminder date*/
    try{
      if(!secondReminderDates){
        throw "No Second Reminder Dates Found.";
      }
      for(var k = 0; k < secondReminderDates.length; k++){
        if(secondReminderDates[k]){
          var secondDateFormatted = "";
          try{
          var secondDateFormatted = Utilities.formatDate(secondReminderDates[k], timeZone, "YYYY-MM-dd");   
          } catch(e){
          // console.error(e);
          }
          if(secondDateFormatted == todayFormatted){
            MailApp.sendEmail(
              allValues[k][2], //to:
              allValues[k][0], //subject:
              allValues[k][8], //body:
              {
                cc:allValues[k][3]
              }); 
            var sentDate = new Date();
            sheet.getRange(k+2, 11).setValue(sentDate);              
            numbEmailsSent++;
            console.info("Email Sent - 2")
          }
        }
      }
    } catch(e){
      console.error(docId + '-sendReminders-2ndDate failed, error: ' + e);      
      catchError(e);
      clearAll();
      return;
    }
    console.info("DocId: " + docId + "; SheetName: " + sheetName + "; Rows: " + numbRows + "; Timezone: " + timeZone + "; NumberEmailsSent: " + numbEmailsSent);
  }
}

/*-----------GENERAL UTILITY---------------*/
//Send Error Email
function catchError(e){
  var errorSubject = "Error Processing Send Reminder Data";
  var errorBody = "There was an error processing your data.  Ensure that you have at least one row of data and that all columns match the template in order, especially date fields and email fields (multiple emails must be separated by commas).  The error is as follows: " + e + ".";
  errorBody = errorBody + "\n\nThe dialy email send has been stopped.  Please correct data in your sheet and then use Settings in the Add-On menu to re-start. \n\n";
  var user = Session.getEffectiveUser().getEmail();
  MailApp.sendEmail(
    user, 
    errorSubject, 
    errorBody
  );
  console.log("Error email sent to user: " + user);
}


function formatTime(hour){
  try{
    if(hour == 12){
      var formatted = "Noon";
    } else if(hour == 0){
      var formatted = "Midnight";
    } else if(hour > 12) {
      var hourPM = hour - 12;
      var formatted = hourPM.toString() + ":00 pm";
    } else if(hour<12) {
      var formatted = hour.toString() + ":00am";
    }
    return formatted;
  } catch (e) {
    console.error('formattime failed, error: ' + e);
    
  }
}


//create single array from list
function createArray(array){
  var newArray = [];
  for(var i = 0; i < array.length; i++){
    newArray.push(array[i][0]);
  }
  return newArray;
}  
