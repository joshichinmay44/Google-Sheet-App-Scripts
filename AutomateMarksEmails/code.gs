// @OnlyCurrentDoc

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Send Email')
      .addItem('Email Student Reports to Parents','emailDictPDFs')
      .addToUi();
}

function getEmailDict(ss) {
  var dictSheet = ss.getSheetByName("Test Emails");
  var recipientDict = {}
  if (!dictSheet) {
    ui.alert('Error', 'The "Recipients" tab was not found. Please create a sheet tab named "Recipients" with Names in Column A and Emails in Column B.', ui.ButtonSet.OK);
    return;
  }
  var dictData = dictSheet.getDataRange().getValues();
  // Loop starts at row index 1 to skip table headers
  for (var d = 1; d < dictData.length; d++) {
    var nameKey = dictData[d][0] ? dictData[d][0].toString().trim() : "";
    var emailValue = dictData[d][1] ? dictData[d][1].toString().trim() : "";
    
    if (nameKey !== "" && emailValue !== "") {
      recipientDict[nameKey] = emailValue; // Map "Student Name": "email@example.com"
    }
  }
  return recipientDict;
}

function emailDictPDFs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  // =======================================================
  // 1. DEFINE YOUR EXACT COLUMN MAP HERE (0 = A, 1 = B, 2 = C...)
  // =======================================================
  var NAME_COLUMN_INDEX = 2;   // Change this if Name is not in Column A
  var ID_COLUMN_INDEX = 1;     // Change this if ID is not in Column B
  var GRADES_START_INDEX = 3;  // Change this to the index where your grade columns begin
  
  // =======================================================
  // 2. DEFINE YOUR RECIPIENT DICTIONARY HERE
  // =======================================================
  // var recipientDict = {
  //   "Vedant B": "joshichinmay96@gmail.com",
  //   "Vedashree P": "joshichinmay96@gmail.com",
  //   "Tanvi C": "joshichinmay96@gmail.com",
  //   "Bhavesh K": "joshichinmay96@gmail.com"
  // };
  var student_absent = 0;

  var recipientDict = getEmailDict(ss);
  
  var sheets = ss.getSheets();
  // var sheetNames = sheets.map(function(s) { return s.getName(); })
  //                        .filter(function(name) { return name !== "SYS_TEMP_TEMPLATE"; });
  
  // var response = ui.prompt(
  //   'Select Data Source',
  //   'Available sheets:\n' + sheetNames.join(', ') + '\n\nEnter the exact name of the sheet to process:',
  //   ui.ButtonSet.OK_CANCEL
  // );
  
  // if (response.getSelectedButton() != ui.Button.OK) return;
  var dataSheet = ss.getActiveSheet();
  var targetSheetName = dataSheet.getName();
  if (!dataSheet) {
    ui.alert('Error', 'Sheet not found.', ui.ButtonSet.OK);
    return;
  }
  
  var data = dataSheet.getDataRange().getValues();
  var headers = data[0]; // Row 1 headers
  if (data.length <= 1) {
    ui.alert('Info', 'No data rows found.', ui.ButtonSet.OK);
    return;
  }
  
  var tempSheet = ss.getSheetByName("SYS_TEMP_TEMPLATE");
  if (!tempSheet) {
    tempSheet = ss.insertSheet("SYS_TEMP_TEMPLATE");
  }
  
  var emailCount = 0;
  
  for (var i = 1; i < data.length; i++) {
    var rowCells = data[i];
    if (!rowCells || rowCells[NAME_COLUMN_INDEX] === "") continue; 
    
    // Dynamically pull from your defined column mapping markers
    var studentName = rowCells[NAME_COLUMN_INDEX]; 
    var studentID   = rowCells[ID_COLUMN_INDEX];   
    
    var recipientEmail = recipientDict[studentName];
    if (!recipientEmail) {
      Logger.log("No email found in dictionary for: " + studentName);
      continue; 
    }
    
    tempSheet.clear();
    
    // --- LARGE LAYOUT DIMENSIONS & DESIGN CONTROLS ---
    tempSheet.setColumnWidth(1, 600); // Column A (Ultra-wide for subjects)
    tempSheet.setColumnWidth(2, 150); // Column B
    tempSheet.setColumnWidth(3, 150); // Column C
    tempSheet.setColumnWidth(4, 150); // Column D
    tempSheet.setColumnWidth(5, 150); // Column E
    tempSheet.setColumnWidth(6, 150); // Column F
    tempSheet.setColumnWidth(7, 150); 
    
    // Main Banner Title
    tempSheet.getRange("A1:B1").merge().setValue("Achievers Defence Academy Report Card - NDA 2026")
             .setFontSize(20).setFontWeight("bold").setHorizontalAlignment("center")
             .setBackground("#1a365d").setFontColor("#ffffff");
    tempSheet.setRowHeight(1, 45).setRowWid; 
    

    // Profile Metadata Fields
    tempSheet.getRange("A3").setValue("Test").setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");
    tempSheet.getRange("B3").setValue(targetSheetName).setFontSize(13).setHorizontalAlignment("center");
    tempSheet.getRange("A4").setValue("Student Name:").setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");
    tempSheet.getRange("B4").setValue(studentName).setFontSize(13).setHorizontalAlignment("center");
    tempSheet.getRange("A5").setValue("ID Number:").setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");
    tempSheet.getRange("B5").setValue(studentID).setFontSize(13).setHorizontalAlignment("center");
    
    tempSheet.setRowHeight(3, 25);
    tempSheet.setRowHeight(4, 25);
    
    // Performance Table Headers
    // tempSheet.getRange("A6").setValue("Assessment").setFontWeight("bold").setFontSize(14).setBackground("#e2e8f0");
    // tempSheet.getRange("B6").setValue("Score").setFontWeight("bold").setFontSize(14).setBackground("#e2e8f0");
    // tempSheet.setRowHeight(6, 30);
    
    var currentRow = 7;
    // Loops starting directly from your defined Grade start column index marker
    for (var j = GRADES_START_INDEX; j < headers.length; j++) {
      if (rowCells[j] == 'Absent'){
        student_absent = 1
        break;
      }
      else if (headers[j] && rowCells[j] !== "") {
        tempSheet.getRange("A" + currentRow).setValue(headers[j]).setFontSize(12).setFontWeight("bold");
        tempSheet.getRange("B" + currentRow).setValue(rowCells[j]).setFontSize(12).setHorizontalAlignment("center");
        tempSheet.setRowHeight(currentRow, 25); 
        currentRow++;
      }
    }
    
    tempSheet.getRange("A6:B" + (currentRow - 1)).setBorder(true, true, true, true, true, true);
    SpreadsheetApp.flush(); 
    
    var sheetId = tempSheet.getSheetId();
    var url = ss.getUrl().replace(/edit$/, '') + 'export?';
    var exportOptions = {
      exportFormat: 'pdf', format: 'pdf', size: 'letter', portrait: 'true',
      fitw: 'true', sheetnames: 'false', printtitle: 'false', gridlines: 'false', fzr: 'false', gid: sheetId
    };
    
    var urlParts = [];
    for (var key in exportOptions) { urlParts.push(key + '=' + exportOptions[key]); }
    var exportUrl = url + urlParts.join('&');
    
    var pdfResponse = UrlFetchApp.fetch(exportUrl, {
      headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    
    
    
    if (student_absent==0){
    var pdfBlob = pdfResponse.getBlob().setName(studentName + "_"+ targetSheetName +"_Report_Card.pdf");
    var emailSubject = "Achievers Academy Official Report Card - " + studentName + " for NDA 2026 "+  targetSheetName;
    var emailBody = "Dear Recipient,\n\nPlease find the official report card for " + studentName + " attached to this email as a PDF document.\n\nBest regards,\nAdministration";
    try {
      MailApp.sendEmail({
        to: recipientEmail,
        subject: emailSubject,
        body: emailBody,
        attachments: [pdfBlob],
         advancedArgs: {
          headers: {
            "X-Priority": "1",         // 1 = High / Urgent priority marker
            "X-MSMail-Priority": "High", // Compatibility marker for Microsoft Outlook
            "Importance": "high"       // Broad webmail client compatibility standard flag
          }
      }
    });
      emailCount++;
    } catch (e) {
      Logger.log("Failed sending email: " + e.message);
    }
    }
    else {
    var emailSubject = "Achievers Academy Official Report Card - " + studentName + " for "+  targetSheetName;
    var emailBody = "Dear Recipient,\n\nYour ward " + studentName + " was absent for NDA "+targetSheetName+".\n\nBest regards,\nAdministration";
    try {
      MailApp.sendEmail({
        to: recipientEmail,
        subject: emailSubject,
        body: emailBody,
         advancedArgs: {
          headers: {
            "X-Priority": "1",         // 1 = High / Urgent priority marker
            "X-MSMail-Priority": "High", // Compatibility marker for Microsoft Outlook
            "Importance": "high"       // Broad webmail client compatibility standard flag
          }
      }
    });
      emailCount++;
    } catch (e) {
      Logger.log("Failed sending email: " + e.message);
    }
    }
  }
  
  ss.deleteSheet(tempSheet);
  ui.alert('Complete', 'Emailed ' + emailCount + ' PDF report cards!', ui.ButtonSet.OK);
}

// Secondary safety trigger backup to force the menu open
function onEdit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuCheck = ss.getUi();
  try {
    ss.addMenu('Report Automation', [
      {name: 'Email Active Sheet Reports', functionName: 'emailDictPDFs'}
    ]);
  } catch(err) {}
}

