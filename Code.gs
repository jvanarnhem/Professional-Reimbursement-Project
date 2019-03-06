var SPREADSHEET_ID = '1e44dVmTQRrXRS2pynzd0cPn-mjKUOtuU_pvOrFDJobY'; // Replace with your spreadsheet ID

var CACHE_PROP = CacheService.getPublicCache();
var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
var SETTINGS_SHEET = "_Settings";
var CACHE_SETTINGS = false;
var SETTINGS_CACHE_TTL = 900;
var cache = JSONCacheService();
var SETTINGS = getSettings();
var subSheet = ss.getSheetByName("Submissions");
var approvedSheet = ss.getSheetByName("Approved");
var compSheet = ss.getSheetByName("Completed");


function doGet(e) {
  var idVal = e.parameter.idNum;
  var step = e.parameter.st;
  var data = null;
  var template = null;
  if (idVal) {
    if(step == 1) {
       template = HtmlService.createTemplateFromFile('admin.html');
       data = subSheet.getDataRange().getValues();
    } else {
       template = HtmlService.createTemplateFromFile('finalForm.html');
       data = approvedSheet.getDataRange().getValues();
    }
    
    for (var i in data) {
      if (data[i][0] == idVal) {
        template.info = data[i];
        break;
      }
    };
  }
  else {
    var template = HtmlService.createTemplateFromFile('index.html');
  }
  var html = template.evaluate();
  return HtmlService.createHtmlOutput(html).setTitle("OFCS Professional Reimbursement Form");
}

function createFolder(folderName) {
  var parentFolderId = SETTINGS.ATTACH_FOLDER_ID;
  try {
    var parentFolder = DriveApp.getFolderById(parentFolderId);
    var folders = parentFolder.getFoldersByName(folderName);
    var folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = parentFolder.createFolder(folderName);
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }
    
    return {
      'folderId' : folder.getId()
    }
  } catch (e) {
    return {
      'error' : e.toString()
    }
  }
}

function uploadFile(base64Data, fileName, folderId) {
  try {
    var foldPath = "https://drive.google.com/drive/folders/"+folderId;
    CACHE_PROP.put('foldPath', foldPath, 3600);
    Logger.log(foldPath);
    var splitBase = base64Data.split(','), type = splitBase[0].split(';')[0]
    .replace('data:', '');
    var byteCharacters = Utilities.base64Decode(splitBase[1]);
    var ss = Utilities.newBlob(byteCharacters, type);
    ss.setName(fileName);
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFilesByName(fileName);
    var file;
    while (files.hasNext()) {
      // delete existing files with the same name.
      file = files.next();
      folder.removeFile(file);
    }
    file = folder.createFile(ss);
    
    return {
      'folderId' : folderId,
      'fileName' : file.getName()
    };
  } catch (e) {
    return {
      'error' : e.toString()
    };
  }
}
//***********************************************************************************************************************************************
function submitReport(data) { 
  var data3 = JSON.parse(data);
  var blding = data3.building;
  var adminemail = "ofcsdistrict@ofcs.net";
  if (blding == "HS") {
    adminemail = SETTINGS.HS_EMAIL;
  } else if (blding == "MS") {
    adminemail = SETTINGS.MS_EMAIL;
  } else if (blding == "OFIS") {
    adminemail = SETTINGS.IS_EMAIL;
  } else if (blding == "FL") {
    adminemail = SETTINGS.FL_EMAIL;
  } else if (blding == "ECC") {
    adminemail = SETTINGS.ECC_EMAIL;
  }
  else {
    adminemail = SETTINGS.DISTRICT_EMAIL;
  }
  try {
    var newStuff = [];
    var subNumber = +new Date();
    var timeStamp = new Date();
    newStuff.push(subNumber);
    newStuff.push(timeStamp);
    for (var i in data3) {
      if(i != data3.length) {
        if(typeof data3[i]=="object"){
          var dataString = data3[i].join(", ");
          newStuff.push(dataString);
        } else {
          newStuff.push(data3[i]);
        }
      }
    }
    //var thePath = CACHE_PROP.get('foldPath');
    //newStuff.push(thePath);
    newStuff.push("Pending");
    newStuff.push("");
    newStuff.push(adminemail);
    subSheet.appendRow(newStuff);
    
    var htmlBody = "<h2>A Professional Reimbursement Form was submitted. </h2>";
    htmlBody += '<p><strong>Click <a href="' + ScriptApp.getService().getUrl()
       + '?idNum=' + subNumber
       + '&st=1">on this link</a> to see full details of report and to approve/deny proposal.</strong></p>';
    htmlBody += "<p>&nbsp;</p>";
    htmlBody += "<h4>Summary: </h4>";
    htmlBody += "<p>Submitter: " + data3['firstName'] + " " + data3['lastName'] + "<br>";
    htmlBody += "Date submitted: " + timeStamp + "<br>";
    htmlBody += "Meeting information: " + data3['meetingInfo']  + "<br>";
    
 
 // CHANGE EMAIL ADDRESS HERE to "adminemail"  
    MailApp.sendEmail({
      to: 'jvanarnhem@ofcs.net',
      subject: "Professional Reimbursement Form Submission: "+data3['lastName'] + " #" + subNumber,
      htmlBody: htmlBody
    });
    return "Submission successful. You may close this window now.";
    
  } catch(err) {
    return "Something went wrong.";
  }
}

// moves a row from a sheet to another when a magic value is entered in a column
function moveCompleted() {
  var sheetNameToWatch = "Submissions";
  var sheetNameToMoveTheRowTo = "Approved";
  var columnNumberToWatch = 12; // column A = 0, B = 1, etc.
  var valueToWatch = "Pending";
    
  var data = subSheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    Logger.log(data[i][columnNumberToWatch]);
    if (data[i][columnNumberToWatch] != valueToWatch) {
      var targetSheet = ss.getSheetByName(sheetNameToMoveTheRowTo);
      var targetRange = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
      subSheet.getRange(i+1, 1, 1, subSheet.getLastColumn()).moveTo(targetRange);
      subSheet.deleteRow(i+1);
    }
  };
}

//***********************************************************************************************************************************************

function rejectSubmission (data) {
  Logger.log("rejected");
  var data3 = JSON.parse(data);
  // CHANGE TO SUBMITTER form.email + ",jtatman@ofcs.net";
 
  update(data3.subNum, 0, data3.adminComments)
  
  var htmlBody = "<p>The following Equivalent Activity Form was rejected by your administrator:</p>";
  htmlBody += "<p>Submitter: " + data3.lastName + ", " + data3.firstName + "<br>";
  htmlBody += "Date submitted: " + data3.timeStamp + "<br>";
  htmlBody += "Equivalent Activity: " + data3.activity + "<br>";
  htmlBody += "Hours completed: " + data3.hours + "</p>";
  htmlBody += "<p>Administrator Comments:" + data3.adminComments + "</p>";
  htmlBody += "<p><a href='" + data3.linkDoc + "' target='_blank'>Link to submitted document</a></p>";
  htmlBody += '<p>Unfortunately you will need to complete the application again if you wish to reapply.</p>';
  htmlBody += '<p>Click <a href="https://script.google.com/macros/s/AKfycbwN6xIs573yTVAV4ltiocMLlPxQ_Ap4r9LZEXWSg7est6lhZoQ/exec">here</a> to resubmit your application.</p>';
 
  MailApp.sendEmail({
    to: data3.email + ", " + SETTINGS.FINAL_EMAIL,
    subject: "Equivalent Activity Proposal not approved" + " #" + data3['subNum'],
    htmlBody: htmlBody
  });

  return "A notice of this rejection will be sent to the applicant. You may close this window now.";
}
//***********************************************************************************************************************************************
function approveSubmission (data) {
  Logger.log("approved");
  var data3 = JSON.parse(data);
  
  update(data3.subNum, 1, data3.adminComments);
  
  var htmlBody = "<p>The following Professional Reimbursement Form was approved by your administrator:</p>";
  htmlBody += "<p>Submitter: " + data3.lastName + ", " + data3.firstName + "<br>";
  htmlBody += "Date submitted: " + data3.timeStamp + "<br>";
  htmlBody += "Meeting: " + data3.meetinginfo + "<br>";
  htmlBody += "Dates of Meeting: " + data3.dates + "</p>";
  htmlBody += "Total Reimbursement: " + data3.grandtotal + "</p>";
  htmlBody += "<p>Administrator Comments:" + data3.adminComments + "</p>";
  htmlBody += '<p><strong><a href="' + ScriptApp.getService().getUrl() + '?idNum=' + data3.subNum 
           + '&st=2" target=_blank>Link to submit final costs and receipts.</a></strong></p>';
  htmlBody += "<p>You will not receive reimbursement until the final form is submitted and approved.</p>";
  htmlBody += "<p>If you have any questions, please contact James Tatman</p>";
 
  MailApp.sendEmail({
    to: data3.email + ", " + SETTINGS.FINAL_EMAIL,
    subject: "Professional Reimbursement Form Approval Notice" + " #" + data3['subNum'],
    htmlBody: htmlBody
  });
  
  return "Your submission is complete. You may close this window now.";
}
//***********************************************************************************************************************************************
function update (num, val, adminComments) {
  var data = subSheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
      if (data[i][0] == num) {
        if (val == 1) {
          subSheet.getRange(i+1, 32, 1, 1).setValue(adminComments);
          subSheet.getRange(i+1, 31, 1, 1).setValue("APPROVED").setBackground("#00FF00");
        } else {
          subSheet.getRange(i+1, 32, 1, 1).setValue(adminComments);
          subSheet.getRange(i+1, 31, 1, 1).setValue("Rejected").setBackground("red");
        }
        break;
      }
  };
  moveCompleted();
}
//***********************************************************************************************************************************************
function getSettings() { 
  if(CACHE_SETTINGS) {
    var settings = cache.get("_settings");
  }
  
  if(settings == undefined) {
    var sheet = ss.getSheetByName(SETTINGS_SHEET);
    var values = sheet.getDataRange().getValues();
  
    var settings = {};
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      settings[row[0]] = row[1];
    }
    
    cache.put("_settings", settings, SETTINGS_CACHE_TTL);
  }
  //Logger.log(settings);
  return settings;
}
function JSONCacheService() {
  var _cache = CacheService.getPublicCache();
  var _key_prefix = "_json#";
  
  var get = function(k) {
    var payload = _cache.get(_key_prefix+k);
    if(payload !== undefined) {
      JSON.parse(payload);
    }
    return payload
  }
  
  var put = function(k, d, t) {
    _cache.put(_key_prefix+k, JSON.stringify(d), t);
  }
  
  return {
    'get': get,
    'put': put
  }
}
//**************************************************************************************************************************************************
function testStuff() {
  Logger.log(SETTINGS.MS_EMAIL);
}