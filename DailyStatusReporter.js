// Not being used
function _emailSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName("Daily Status Report")
  var file = DriveApp.getFileById('TBD') //open to the sheet you want and copy the ID from the url and put it here   
  
  MailApp.sendEmail("<Project Name>qa@<company>.com", "<Project Name> : Task Allocation for Today", "Hi All, Please find attached task allocation for today", { //email address, subject, message
     name: 'DSR - Apr 2019', //file name
     attachments: [file.getAs(MimeType.PDF)] //file to attach
 });
}

function emailSummary() {
  var recepient = ""
  var subject = "<Project Name> DSR : " + new Date().toDateString();
  var messagebody = generateDSR(); 
  MailApp.sendEmail(recepient, subject, "Daily Status Updates for <Project Name>",{htmlBody:messagebody});
}



//This function will pick up data from the existing range and send it as DSR over mail to dsr@<company>.com
function generateDSR() {
  
  // Generate DSR mail body
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Daily Status Report")
  var r = sheet.getLastRow();
  var c = sheet.getLastColumn();
  var dataRange = sheet.getRange(1,1,r, c);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  var emailBodyMessage = "<head><style>.blue {background: #80b3ff; } .done {background: #00e64d;} .pending {background: #ff5c33 } .inprogress {background: #ffad33 }  .onleave {background: #ad33ff }</style></head><body>Hi All,<br><br> Following is an integrated DSR for the <Project Name> project <br><br>";
  var messageBody = emailBodyMessage + "<table border = 1><tr class='blue' ><th>"+ data[0][0] + "</th><th>" + data[0][1] +"</th><th>" + data[0][2] + "</th><th>" + data[0][3] + "</th><th>"+ data[0][4] + "</th></tr>";
  
  for (var i = 1; i < data.length; i++) {
    var value = "<tr><td>" + data[i][0] + "</td><td>" + data[i][1] + "</td>";
    var status = "<td ";
    // Highlight based on completion status
    if (data[i][2].toString() == 'Complete')
        status += " class='done'>Complete</td><td>" + data[i][3] + "</td><td>" + data[i][4]+ "</td></tr>";
    else if(data[i][2].toString() =='Pending')
      status += "  class='pending'>Pending</td><td>" + data[i][3] + "</td><td>" + data[i][4]+ "</td></tr>";
    else if (data[i][2].toString() =='In Progress')
      status += "  class='inprogress'>In Progress</td><td>" + data[i][3] + "</td><td>" + data[i][4]+ "</td></tr>";
    else if (data[i][2].toString() =='On Leave')
      status += "  class='onleave'>On Leave</td><td>" + data[i][3] + "</td><td>" + data[i][4]+ "</td></tr>";
    value += status
    messageBody += value;
  }
  messageBody += "</table> <br> <br> <i>This is an auto-generated mail.Please do not reply to this mail.For any queries please reply to <Project Name>qa@<company>.com";
  Logger.log("Email message body in html %s",messageBody);
  return messageBody;
}

// Clear data
function clearData(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Daily Status Report")
  var r = sheet.getLastRow();
  var c = sheet.getLastColumn();
  var dataRange = sheet.getRange(2,1,r, c);
  dataRange.clearContent();
}


// TODO: Function to send individual reminders.
function onEdit() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Daily Status Report");
  Logger.log("Changes in %s spreadsheet",sheet.getName())
  
  var range = sheet.getActiveRange();
  var r = sheet.getLastRow();
  var c = sheet.getLastColumn();
  Logger.log("Data:  (row,column) (%s,%s)",r,c)
}

