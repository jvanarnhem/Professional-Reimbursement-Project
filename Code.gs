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
    } else if (step == 2) {
       template = HtmlService.createTemplateFromFile('district.html');
       data = subSheet.getDataRange().getValues();
    } else if (step == 3) {
       template = HtmlService.createTemplateFromFile('finalForm.html');
       data = approvedSheet.getDataRange().getValues();
    }  else {
       template = HtmlService.createTemplateFromFile('finalApproval.html');
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
  var adminname = "Mr. Admin";
  if (blding == "HS") {
    adminemail = SETTINGS.HS_EMAIL;
    adminname = SETTINGS.HS_ADMIN;
  } else if (blding == "MS") {
    adminemail = SETTINGS.MS_EMAIL;
    adminname = SETTINGS.MS_ADMIN;
  } else if (blding == "OFIS") {
    adminemail = SETTINGS.IS_EMAIL;
    adminname = SETTINGS.IS_ADMIN;
  } else if (blding == "FL") {
    adminemail = SETTINGS.FL_EMAIL;
    adminname = SETTINGS.FL_ADMIN;
  } else if (blding == "ECC") {
    adminemail = SETTINGS.ECC_EMAIL;
    adminname = SETTINGS.ECC_ADMIN;
  }
  else {
    adminemail = SETTINGS.DISTRICT_EMAIL;
    adminname = SETTINGS.DISTRICT_ADMIN;
  }
  try {
    var newStuff = [];
    var subNumber = +new Date();
    var timeStamp = new Date();
    var nowDate = Utilities.formatDate(new Date(), "GMT-4", "MMM d, yyyy");
    newStuff.push(subNumber);
    newStuff.push(timeStamp);
    newStuff.push(nowDate);
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
    newStuff.push(adminname);
    newStuff.push(adminemail);
    subSheet.appendRow(newStuff);
    var column1 = subSheet.getRange("AE2:AE");
    var column2 = subSheet.getRange("C2:C");
    column1.setNumberFormat("MMM D, YYYY");
    column1.setNumberFormat("@STRING@");
    column2.setNumberFormat("@STRING@");
    var column3 = subSheet.getRange("P2:AC");
    column3.setNumberFormat("@STRING@");
    
    var htmlBody = "<h2>A Professional Reimbursement Form was submitted. </h2>";
    htmlBody += '<p><strong>Click <a href="' + ScriptApp.getService().getUrl()
       + '?idNum=' + subNumber
       + '&st=1">on this link</a> to see full details of report and to approve/deny proposal.</strong></p>';
    htmlBody += "<p>&nbsp;</p>";
    htmlBody += "<h4>Summary: </h4>";
    htmlBody += "<p>Submitter: " + data3['firstName'] + " " + data3['lastName'] + "<br>";
    htmlBody += "Date submitted: " + nowDate + "<br>";
    htmlBody += "Meeting information: " + data3['meetingInfo']  + "<br>";
    
 
 // CHANGE EMAIL ADDRESS HERE to "adminemail"  
    MailApp.sendEmail({
      to: 'jvanarnhem@ofcs.net',
      subject: "B Professional Reimbursement Form Submission: "+data3['lastName'] + " #" + subNumber,
      htmlBody: htmlBody
    });
    return "Submission successful. You may close this window now.";
    
  } catch(err) {
    return "Something went wrong in submit report.";
  }
}

// moves a row from a sheet to another when a magic value is entered in a column
function moveCompleted() {
  var sheetNameToWatch = "Submissions";
  var sheetNameToMoveTheRowTo = "Approved";
  var columnNumberToWatch = 31; // column A = 0, B = 1, etc.
  var valueToWatch = "Pending";
    
  var data = subSheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
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
  
  var data3 = JSON.parse(data);
  // CHANGE TO SUBMITTER form.email + ",jtatman@ofcs.net";
 
  update(data3.subNum, 0, data3.adminComments,"")
  moveRejected1();
  var htmlBody = "<p>The following Equivalent Activity Form was rejected by your administrator:</p>";
  htmlBody += "<p>Submitter: " + data3.lastName + ", " + data3.firstName + "<br>";
  htmlBody += "Date submitted: " + data3.timeStamp + "<br>";
  htmlBody += "Meeting: " + data3.meetingInfo + "<br>";
  htmlBody += "Dates of Meeting: " + data3.dates + "</p>";
  htmlBody += "Total Reimbursement: " + data3.grandtotal + "</p>";
  htmlBody += "<p>Administrator Comments:" + data3.adminComments + "</p>";
  //htmlBody += "<p><a href='" + data3.linkDoc + "' target='_blank'>Link to submitted document</a></p>";
  htmlBody += '<p>Unfortunately you will need to complete the application again if you wish to reapply.</p>';
  htmlBody += '<p>Click <a href="https://script.google.com/macros/s/AKfycbw3anWQOg-TVmlA2W0SpngyuXvc70UW6CplUTPwC4GKa9LRiRRR/exec">here</a> to resubmit your application.</p>';
 
  MailApp.sendEmail({
    to: data3.email + ", " + SETTINGS.FINAL_EMAIL,
    subject: "Equivalent Activity Proposal not approved" + " #" + data3['subNum'],
    htmlBody: htmlBody
  });

  return "A notice of this rejection will be sent to the applicant. You may close this window now.";
}
//***********************************************************************************************************************************************
function approveSubmission (data) {
  
  var data3 = JSON.parse(data);
  
  update(data3.subNum, 1, data3.adminComments,"");
  
  var htmlBody = "<h2>A Professional Reimbursement Form was submitted. </h2>";
  htmlBody += '<p><strong>Click <a href="' + ScriptApp.getService().getUrl()
       + '?idNum=' + data3.subNum
       + '&st=2">on this link</a> to see full details of report and to approve/deny proposal.</strong></p>';
  htmlBody += "<p>&nbsp;</p>";
  htmlBody += "<h4>Summary: </h4>";
  htmlBody += "<p>Submitter: " + data3.lastName + ", " + data3.firstName + "<br>";
  htmlBody += "Date submitted: " + data3.timeStamp + "<br>";
  htmlBody += "Dates of Meeting: " + data3.dates + "</p>";
  htmlBody += "Meeting information: " + data3.meetingInfo  + "<br>";
  htmlBody += "Total Reimbursement: " + data3.grandtotal + "</p>";
  htmlBody += "<p>Administrator Comments:" + data3.adminComments + "</p>";
 
  MailApp.sendEmail({
    to: data3.email + ", " + SETTINGS.FINAL_EMAIL,
    subject: "D Professional Reimbursement Form Approval Notice" + " #" + data3['subNum'],
    htmlBody: htmlBody
  });
  
  return "A notice of this rejection will be sent to the applicant. You may close this window now.";
}
//***********************************************************************************************************************************************

function rejectSubmission2 (data) {
  
  var data3 = JSON.parse(data);
  // CHANGE TO SUBMITTER form.email + ",jtatman@ofcs.net";
 
  update(data3.subNum, 0, data3.adminComments,"")
  moveRejected1();
  var htmlBody = "<p>The following Equivalent Activity Form was rejected by your administrator:</p>";
  htmlBody += "<p>Submitter: " + data3.lastName + ", " + data3.firstName + "<br>";
  htmlBody += "Date submitted: " + data3.timeStamp + "<br>";
  htmlBody += "Meeting: " + data3.meetingInfo + "<br>";
  htmlBody += "Dates of Meeting: " + data3.dates + "</p>";
  htmlBody += "Total Reimbursement: " + data3.grandtotal + "</p>";
  htmlBody += "<p>Building Administrator Comments:" + data3.adminComments + "</p>";
  htmlBody += "<p>District Administrator Comments:" + data3.districtComments + "</p>";
  //htmlBody += "<p><a href='" + data3.linkDoc + "' target='_blank'>Link to submitted document</a></p>";
  htmlBody += '<p>Unfortunately you will need to complete the application again if you wish to reapply.</p>';
  htmlBody += '<p>Click <a href="https://script.google.com/macros/s/AKfycbw3anWQOg-TVmlA2W0SpngyuXvc70UW6CplUTPwC4GKa9LRiRRR/exec">here</a> to resubmit your application.</p>';
 
  MailApp.sendEmail({
    to: data3.email + ", " + SETTINGS.FINAL_EMAIL,
    subject: "Equivalent Activity Proposal not approved" + " #" + data3['subNum'],
    htmlBody: htmlBody
  });

  return "A notice of this rejection will be sent to the applicant. You may close this window now.";
}
function moveRejected1() {
  var sheetNameToWatch = "Submissions";
  var sheetNameToMoveTheRowTo = "Rejected";
  var columnNumberToWatch = 31; // column A = 0, B = 1, etc.
  var valueToWatch1 = "Rejected - District";
  var valueToWatch2 = "Rejected - Building";
  
  var data = subSheet.getDataRange().getValues();
  //Logger.log(data[data.length-1][columnNumberToWatch].val);
  for (var i = data.length - 1; i >= 1; i--) {
    //Logger.log(data[i][columnNumberToWatch]);
    if (data[i][columnNumberToWatch] == valueToWatch1 || data[i][columnNumberToWatch] == valueToWatch2) {
      var targetSheet = ss.getSheetByName(sheetNameToMoveTheRowTo);
      var targetRange = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
      subSheet.getRange(i+1, 1, 1, subSheet.getLastColumn()).moveTo(targetRange);
      subSheet.deleteRow(i+1);
    }
  };
}
//***********************************************************************************************************************************************
function approveSubmission2 (data) {
 
  var data3 = JSON.parse(data);
  
  update(data3.subNum, 2, data3.districtComments, data3.route);
  var finalDoc = doMerge(data3.subNum, data3.lastName, SETTINGS.DESTINATION_FOLDER_ID, SETTINGS.FORM_TEMPLATE_ID, SPREADSHEET_ID);
  moveCompleted();
  var htmlBody = "<p>The following Professional Reimbursement Form was approved by your administrator:</p>";
  htmlBody += "<p>Submitter: " + data3.lastName + ", " + data3.firstName + "<br>";
  htmlBody += "Date submitted: " + data3.timeStamp + "<br>";
  htmlBody += "Meeting: " + data3.meetingInfo + "<br>";
  htmlBody += "Dates of Meeting: " + data3.dates + "</p>";
  htmlBody += "Total Reimbursement: " + data3.grandtotal + "</p>";
  htmlBody += "<p>Building Administrator Comments:" + data3.adminComments + "</p>";
  htmlBody += "<p>District Administrator Comments:" + data3.districtComments + "</p>";
  htmlBody += '<p><strong><a href="' + ScriptApp.getService().getUrl() + '?idNum=' + data3.subNum 
           + '&st=3" target=_blank>Link to submit final costs and receipts.</a></strong></p>';
  htmlBody += "<p>You will not receive reimbursement until the final form is submitted and approved.</p>";
  htmlBody += "<p>If you have any questions, please contact James Tatman</p>";
 
  MailApp.sendEmail({
    to: data3.email + ", " + SETTINGS.FINAL_EMAIL,
    subject: "Professional Reimbursement Form Approval Notice" + " #" + data3['subNum'],
    htmlBody: htmlBody,
    attachments: [finalDoc.getAs(MimeType.PDF)]
  });
  
  return "Your submission is complete. You may close this window now.";
}
//***********************************************************************************************************************************************

function submitReport2 (data) {
  
  var data3 = JSON.parse(data);
  
  //update(data3.subNum, 1, data3.adminComments);
  var thePath = CACHE_PROP.get('foldPath');
  
  update2(data3.subNum, 5, data3.finalComments, thePath, data3.signDate2, data3.finalsignature);
  updateCosts(data3.subNum, data3);
  
  var column4 = approvedSheet.getRange("AN2:AW");
  column4.setNumberFormat("$####.00");
  column4.setNumberFormat("@STRING@");
  var column5 = approvedSheet.getRange("AX2:AX");
  column5.setNumberFormat("MMM D, YYYY");
  column5.setNumberFormat("@STRING@");
  
  var htmlBody = "<h2>A Professional Reimbursement Form was submitted. </h2>";
  htmlBody += '<p><strong>Click <a href="' + ScriptApp.getService().getUrl()
       + '?idNum=' + data3.subNum
       + '&st=4">on this link</a> to review full details of report and to approve/deny reimbursement.</strong></p>';
  htmlBody += "<p>&nbsp;</p>";
  htmlBody += "<h4>Summary: </h4>";
  htmlBody += "<p>Submitter: " + data3.lastName + ", " + data3.firstName + "<br>";
  htmlBody += "Date submitted: " + data3.signDate2 + "<br>";
  htmlBody += "Dates of Meeting: " + data3.dates + "</p>";
  htmlBody += "Meeting information: " + data3.meetingInfo  + "<br>";
  htmlBody += "Total Reimbursement: " + data3.actGrandTotal + "</p>";
  htmlBody += "<p>Final Submission Comments:" + data3.finalComments + "</p>";
  htmlBody += "<p>Link to Documentation: <a href='" + thePath + "' target='_BLANK'>Click</a></p>";
  var adminemail = SETTINGS.DISTRICT_EMAIL;
  if(data3.route == "Building") {
    var blding = data3.building;
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
  }
  MailApp.sendEmail({
    to: "jvanarnhem@ofcs.net", //adminemail,
    subject: "Final Approval for Professional Reimbursement Form" + " #" + data3['subNum'] + " to:"+adminemail,
    htmlBody: htmlBody
  });
  
  return "Your submission is complete. You may close this window now.";
}

//***********************************************************************************************************************************************

function approveSubmission3 (data) {
  var data3 = JSON.parse(data);
  //Logger.log("hi");
  update3(data3.subNum, 1, data3.route);
  //var finalDoc = doMerge(data3.subNum, data3.lastName, SETTINGS.DESTINATION_FOLDER_ID, SETTINGS.FORM_TEMPLATE_ID, SPREADSHEET_ID);
  moveCompleted2();
  var htmlBody = "<p>The following Professional Reimbursement Form was approved for reimbursement:</p>";
  htmlBody += "<p>Submitter: " + data3.lastName + ", " + data3.firstName + "<br>";
  htmlBody += "Date submitted: " + data3.timeStamp + "<br>";
  htmlBody += "Meeting: " + data3.meetingInfo + "<br>";
  htmlBody += "Dates of Meeting: " + data3.dates + "</p>";
  htmlBody += "Total Reimbursement: " + data3.actGrandTotal + "</p>";
  
  htmlBody += "<p>If you have any questions, please contact James Tatman</p>";
 
  MailApp.sendEmail({
    to: data3.email,
    subject: "Professional Reimbursement Final Approval Notice" + " #" + data3['subNum'],
    htmlBody: htmlBody
  });
  
  return "Your submission is complete. You may close this window now.";
}
function moveCompleted2() {
  var sheetNameToWatch = "Approved";
  var sheetNameToMoveTheRowTo = "Completed";
  var columnNumberToWatch = 31; // column A = 0, B = 1, etc.
  var valueToWatch = "FINAL APPROVAL";
  
  var data = approvedSheet.getDataRange().getValues();
  //Logger.log(data[data.length-1][columnNumberToWatch].val);
  for (var i = data.length - 1; i >= 1; i--) {
    //Logger.log(data[i][columnNumberToWatch]);
    if (data[i][columnNumberToWatch] == valueToWatch) {
      var targetSheet = ss.getSheetByName(sheetNameToMoveTheRowTo);
      var targetRange = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
      approvedSheet.getRange(i+1, 1, 1, subSheet.getLastColumn()).moveTo(targetRange);
      approvedSheet.deleteRow(i+1);
    }
  };
}
//***********************************************************************************************************************************************
function rejectSubmission3 (data) {
 
  var data3 = JSON.parse(data);
  // CHANGE TO SUBMITTER form.email + ",jtatman@ofcs.net";
 
  update(data3.subNum, 0, data3.adminComments,"")
  
  var htmlBody = "<p>The following Equivalent Activity Form was rejected by your administrator:</p>";
  htmlBody += "<p>Submitter: " + data3.lastName + ", " + data3.firstName + "<br>";
  htmlBody += "Date submitted: " + data3.timeStamp + "<br>";
  htmlBody += "Meeting: " + data3.meetingInfo + "<br>";
  htmlBody += "Dates of Meeting: " + data3.dates + "</p>";
  htmlBody += "Total Reimbursement: " + data3.grandtotal + "</p>";
  htmlBody += "<p>Building Administrator Comments:" + data3.adminComments + "</p>";
  htmlBody += "<p>District Administrator Comments:" + data3.districtComments + "</p>";
  //htmlBody += "<p><a href='" + data3.linkDoc + "' target='_blank'>Link to submitted document</a></p>";
  htmlBody += '<p>Unfortunately you will need to complete the application again if you wish to reapply.</p>';
  htmlBody += '<p>Click <a href="https://script.google.com/macros/s/AKfycbwN6xIs573yTVAV4ltiocMLlPxQ_Ap4r9LZEXWSg7est6lhZoQ/exec">here</a> to resubmit your application.</p>';
 
  MailApp.sendEmail({
    to: data3.email + ", " + SETTINGS.FINAL_EMAIL,
    subject: "Equivalent Activity Proposal not approved" + " #" + data3['subNum'],
    htmlBody: htmlBody
  });

  return "A notice of this rejection will be sent to the applicant. You may close this window now.";
}
function moveRejected2() {
  var sheetNameToWatch = "Approved";
  var sheetNameToMoveTheRowTo = "Rejected";
  var columnNumberToWatch = 31; // column A = 0, B = 1, etc.
  var valueToWatch1 = "Rejected - District";
  var valueToWatch2 = "Rejected - Building";
  
  var data = approvedSheet.getDataRange().getValues();
  //Logger.log(data[data.length-1][columnNumberToWatch].val);
  for (var i = data.length - 1; i >= 1; i--) {
    //Logger.log(data[i][columnNumberToWatch]);
    if (data[i][columnNumberToWatch] == valueToWatch1 || data[i][columnNumberToWatch] == valueToWatch2) {
      var targetSheet = ss.getSheetByName(sheetNameToMoveTheRowTo);
      var targetRange = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
      approvedSheet.getRange(i+1, 1, 1, subSheet.getLastColumn()).moveTo(targetRange);
      approvedSheet.deleteRow(i+1);
    }
  };
}
//***********************************************************************************************************************************************

function update (num, val, adminComments, route) {
  var data = subSheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
      if (data[i][0] == num) {
        if (val == 1) {
          subSheet.getRange(i+1, 35, 1, 1).setValue(adminComments);
          subSheet.getRange(i+1, 32, 1, 1).setValue("BUILDING APPROVED").setBackground("#00FF00");
        } else if (val == 2) {
          subSheet.getRange(i+1, 36, 1, 1).setValue(adminComments);
          subSheet.getRange(i+1, 32, 1, 1).setValue("DISTRICT APPROVED").setBackground("#00FF00");
          subSheet.getRange(i+1, 37, 1, 1).setValue(route);
        } else if (val == -2) {
          subSheet.getRange(i+1, 36, 1, 1).setValue(adminComments);
          subSheet.getRange(i+1, 32, 1, 1).setValue("Rejected - District").setBackground("red");
        } else if (val == 5) {
          subSheet.getRange(i+1, 39, 1, 1).setValue(path);
        } else {
          subSheet.getRange(i+1, 35, 1, 1).setValue(adminComments);
          subSheet.getRange(i+1, 32, 1, 1).setValue("Rejected - Building").setBackground("red");
        }
        break;
      }
  };
  
}
//***********************************************************************************************************************************************
function update2 (num, val, finalComments, path, finalSign, finalDate) {
  var data = approvedSheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
      if (data[i][0] == num) {
        if (val == 5) {
          approvedSheet.getRange(i+1, 38, 1, 1).setValue(finalComments);
          approvedSheet.getRange(i+1, 39, 1, 1).setValue(path);
          approvedSheet.getRange(i+1, 50, 1, 1).setValue(finalSign);
          approvedSheet.getRange(i+1, 51, 1, 1).setValue(finalDate);
          
        } else {
          //approvedSheet.getRange(i+1, 35, 1, 1).setValue(adminComments);
          approvedSheet.getRange(i+1, 32, 1, 1).setValue("Rejected - Building").setBackground("red");
        }
        break;
      }
  };
  
}
//***********************************************************************************************************************************************
function update3 (num, val, route) {
  var data = approvedSheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
      if (data[i][0] == num) {
        if (val == 1) {
          approvedSheet.getRange(i+1, 32, 1, 1).setValue("FINAL APPROVAL").setBackground("#42f4eb");
          approvedSheet.getRange(i+1, 52, 1, 1).setValue(route); 
        } else {
          //approvedSheet.getRange(i+1, 35, 1, 1).setValue(adminComments);
          approvedSheet.getRange(i+1, 32, 1, 1).setValue("Rejected - FINAL").setBackground("red");
        }
        break;
      }
  };
  
}
//***********************************************************************************************************************************************
function updateCosts (num, data3) {
  var data = approvedSheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
      if (data[i][0] == num) {
          approvedSheet.getRange(i+1, 40, 1, 1).setValue(data3.actMileageTotal);
          approvedSheet.getRange(i+1, 41, 1, 1).setValue(SETTINGS.MILEAGE_RATE);
          approvedSheet.getRange(i+1, 42, 1, 1).setValue(data3.actMileageTotal*SETTINGS.MILEAGE_RATE);
          approvedSheet.getRange(i+1, 43, 1, 1).setValue(data3.actRegistrationTotal);
          approvedSheet.getRange(i+1, 44, 1, 1).setValue(data3.actParkingTotal);
          approvedSheet.getRange(i+1, 45, 1, 1).setValue(data3.actLodgingTotal);
          approvedSheet.getRange(i+1, 46, 1, 1).setValue(data3.actBreakfastTotal);
          approvedSheet.getRange(i+1, 47, 1, 1).setValue(data3.actLunchTotal);
          approvedSheet.getRange(i+1, 48, 1, 1).setValue(data3.actDinnerTotal);
          approvedSheet.getRange(i+1, 49, 1, 1).setValue(data3.actGrandTotal);
          break;
      }
  };
  
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