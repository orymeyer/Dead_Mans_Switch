var sheet      = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
var ACTIVATION =sheet.getRange("B2").getValue();
var DEADLINE   =sheet.getRange("B3").getValue();
var STATUS     =sheet.getRange("B4").getValue();
  
var TO         =sheet.getRange("B7").getValue();
var SUB        =sheet.getRange("B8").getValue();
var BODY       =sheet.getRange("B9").getValue();

var EMAIL_ALERTS =sheet.getRange("B12").getValue();
var SMS_ALERTS   =sheet.getRange("B13").getValue();
var id           = ScriptProperties.getProperty("ID");
var alerttimes   = sheet.getRange("B14").getValue();

function onOpen() {
   var sheet = SpreadsheetApp.getActiveSpreadsheet();
   var id    = SpreadsheetApp.getActiveSpreadsheet().getId();
  
   var menu = [ 
   {name: "Step 1:Initialize & Start ", functionName: "createTrigger"},
   {name: "Step 2:Stop  ", functionName: "destroyTrigger"}
   ];  
  
   sheet.addMenu("DMS-Control", menu);
  
  if(ACTIVATION=="YES")
  {
   postponeDeadline();
  }
   Logger.log("DMS was opened.");
   
  ScriptProperties.setProperty("ID",id);
  
};

function checkActivation(){

  if(ACTIVATION != "YES")
  {
    //DMS is inactive.Activate the DMS
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.getRange("B2").setValue("YES");
  }
 }

function postponeDeadline(){
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 sheet.getRange("B4").setValue("Y");//set status as Y
}

function resetStatus(){
  var sheet = SpreadsheetApp.openById(id);
  sheet.getRange("B4").setValue("N");//set status as N
}

function createTrigger(){
//Set activation as Yes,function loops.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange("B2").setValue("YES");
  
  ScriptApp.newTrigger("check")
     .timeBased()
     .everyDays(1)
     .create();

  ScriptApp.newTrigger("alerts")
  .timeBased()
  .everyHours(1)
  .create();

}

function destroyTrigger(){

  var triggers = ScriptApp.getScriptTriggers();
  for(var i=0; i < triggers.length; i++) {
  ScriptApp.deleteTrigger(triggers[i]);
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange("B2").setValue("NO");//Set activation as NO.Dectivate the DMS.
  Logger.log("DMS destroyed.");
}

function sendMail(){
  //Send the email 
  MailApp.sendEmail(TO,SUB,BODY);  
}


function check(){
  
  var sheet = SpreadsheetApp.openById(id);
  var status =sheet.getRange("B4").getValue();
 
  if(status.toString()!="Y"){
    //Dead Mans Switch is Triggered
    sendMail();
    destroyTrigger();
    Logger.log("DMS is Triggered!")
 }
  
  else
  {
    //Reset Status
    resetStatus();
    Logger.log("Resetting status...");
  }
}

function calculatedeadline(){
  
  var date = new Date();
  var rem_hours = 24 - date.getHours();
  var rem_min = 60 - date.getMinutes();
  var d = 24+rem_hours + " Hours,"+rem_min+" Minutes";
  Logger.log(24+rem_hours + " Hours,"+rem_min+" Minutes");
  
  if(rem_hours>0)
  {
    d= "Open this sheet after "+date.getDate()+"/"+(date.getMonth()+1)+"/"+date.getYear();
    return d; 
  }
  else
  {
    return d;
  }
}

function alerts()
{
  var date = new Date();
  var rem_hours = 24 - date.getHours();
 if(EMAIL_ALERTS=="YES" && rem_hours<=alerttimes)
 {
  
   //send alert email
   subject="DMS-ALERT";
   body   ="This is to notify you that the DEAD MANS SWITCH you have set up will trigger in less than "+rem_hours+" hours"
          +"Open the DMS sheet to prevent this"+"You can deactivate the sheet from the DMS menu" ; 
   MailApp.sendEmail(TO, subject, body)
   Logger.log("Alert Email Sent");
 } 
  
  if(SMS_ALERTS="YES" && rem_hours<=alerttimes)
  {
   //Implement SMS alerts
   var time = new Date(); 
   var now   = new Date(time.getTime() + 10000);
   CalendarApp.createEvent(msg, now, now).addSmsReminder(0);

  }
}