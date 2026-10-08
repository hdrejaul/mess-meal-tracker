// =====================================================================================
// 🚀 MESS MEAL MANAGEMENT SYSTEM - UNIVERSAL MASTER GOOGLE APPS SCRIPT (Code.gs)
// ⚡ 100% CLEAN MAIN SHEETS + CLOUD DEPOSIT LOGS (MAX LAST 2 MONTHS AUTO-RETENTION):
//    - Main monthly sheets (Jan..Dec) have 0 extra rows/columns (100% Clean Sheet)
//    - Automated "DepositLogs" tab records deposit receipts across all devices
//    - Auto-Prunes deposit history older than last 2 months automatically
//    - Dynamic Month Switching & All-Months Discovery
//    - Real-Time Live Sync with Web App
// =====================================================================================

var MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

// 📌 ১. চলতি বা নির্দিষ্ট মাসের শিট স্বয়ংক্রিয়ভাবে বের করা বা তৈরি করা
function getOrCreateSheet(ss, targetMonthName) {
  var monthName = targetMonthName;
  if (!monthName || monthName === "current") {
    monthName = MONTH_NAMES[new Date().getMonth()];
  }
  
  var sheet = ss.getSheetByName(monthName);
  if (sheet) return sheet;
  
  // নতুন মাস শুরু হলে আগের মাসের শিটটিকে হুবহু ক্লোন করে সব ফর্মুলাসহ ফ্রেশ নতুন শিট তৈরি করা
  var sheets = ss.getSheets();
  var templateSheet = null;
  
  // DepositLogs ছাড়া যেকোনো মূল মাসিক শিট খুঁজে নেওয়া
  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName();
    if (!sName.toLowerCase().includes("log") && !sName.toLowerCase().includes("deposit")) {
      templateSheet = sheets[i];
      break;
    }
  }
  if (!templateSheet) templateSheet = sheets[0];
  
  // হুবহু ফরম্যাট, রঙ ও সমস্ত ফর্মুলা অক্ষত রেখে নতুন শিট তৈরি
  var newSheet = templateSheet.copyTo(ss).setName(monthName);
  
  try {
    newSheet.getRange("B2:E32").clearContent(); // শুধুমাত্র মিলের ঘর রিসেট (ফর্মুলা অক্ষত)
    newSheet.getRange("K2:K5").setValue(0);      // মেম্বার জমার ঘর ০ করা (ফর্মুলা অক্ষত)
    newSheet.getRange("H10:I23").clearContent(); // ক্যাটারিং পেমেন্ট ক্লিয়ার
  } catch(e) {}
  
  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(1); // চলতি মাসের শিটটিকে সবার শুরুতে নিয়ে আসা
  return newSheet;
}

// 📌 ২. স্প্রেডশিটের সব সক্রিয় মাসের তালিকা পড়া
function getAllAvailableMonths(ss) {
  var sheets = ss.getSheets();
  var list = [];
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    // DepositLogs বা টেমপ্লেট শিট বাদ দিয়ে শুধুমাত্র মাসের নামের শিটগুলো নেওয়া
    if (!name.toLowerCase().includes("copy") && 
        !name.toLowerCase().includes("template") && 
        !name.toLowerCase().includes("depositlog") && 
        !name.toLowerCase().includes("log")) {
      list.push(name);
    }
  }
  return list;
}

// 📌 ৩. ক্লাউড ডিপোজিট লগ শিট তৈরি করা বা খুঁজে নেওয়া
function getOrCreateDepositLogsSheet(ss) {
  var logSheet = ss.getSheetByName("DepositLogs");
  if (!logSheet) {
    logSheet = ss.insertSheet("DepositLogs");
    logSheet.appendRow(["Timestamp", "Month", "Date", "Member", "Amount", "Note"]);
    logSheet.setFrozenRows(1);
    try {
      logSheet.getRange("A1:F1").setFontWeight("bold").setBackground("#e0f2fe");
    } catch(e){}
  }
  return logSheet;
}

// 📌 ৪. সর্বোচ্চ বিগত ২ মাসের ডাটা রেখে বাকি পুরোনো ডাটা স্বয়ংক্রিয়ভাবে মুছে ফেলা (Auto-Prune)
function pruneDepositLogs(logSheet) {
  try {
    var lastRow = logSheet.getLastRow();
    if (lastRow <= 1) return;
    
    var now = new Date();
    var curMonthIdx = now.getMonth();
    var prevMonthIdx = (curMonthIdx - 1 + 12) % 12;
    var allowedMonths = [MONTH_NAMES[curMonthIdx], MONTH_NAMES[prevMonthIdx]];
    
    var values = logSheet.getRange(2, 1, lastRow - 1, 6).getValues();
    var rowsToKeep = [];
    
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var rowMonth = String(row[1] || "").trim();
      var rowDateVal = row[2];
      
      var isRecent = false;
      // চলতি বা আগের মাসের নাম থাকলে
      if (rowMonth && allowedMonths.indexOf(rowMonth) !== -1) {
        isRecent = true;
      } else if (rowDateVal) {
        var d = new Date(rowDateVal);
        if (!isNaN(d.getTime())) {
          var diffDays = (now.getTime() - d.getTime()) / (1000 * 3600 * 24);
          if (diffDays <= 65) isRecent = true; // সর্বোচ্চ ৬৫ দিন (২ মাস)
        }
      }
      
      if (isRecent) {
        rowsToKeep.push(row);
      }
    }
    
    if (rowsToKeep.length !== values.length) {
      logSheet.getRange(2, 1, Math.max(values.length, 1), 6).clearContent();
      if (rowsToKeep.length > 0) {
        logSheet.getRange(2, 1, rowsToKeep.length, 6).setValues(rowsToKeep);
      }
    }
  } catch(err) {
    Logger.log("Prune error: " + err);
  }
}

// 📌 ৫. রিয়েল-টাইম অপটিমাইজড রিড হ্যান্ডলার (doGet)
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var params = e && e.parameter ? e.parameter : {};
  
  var reqMonth = params.month || "current";
  var sheet = getOrCreateSheet(ss, reqMonth);
  
  // ⚡ ব্যাচ রিড কল (Batch Read - Ultra Fast)
  var allValues = sheet.getRange(1, 1, 35, 12).getDisplayValues();
  
  // দৈনিক ৩১ দিনের মিল (A2:E32)
  var dailyMealsList = [];
  for (var r = 1; r <= 31; r++) {
    if (allValues[r]) {
      dailyMealsList.push(allValues[r].slice(0, 5));
    }
  }
  
  // সদস্যভিত্তিক হিসাব টেবিল (H2:L6 -> Cols 7..11)
  var summaryRange = [];
  for (var r = 1; r <= 5; r++) {
    if (allValues[r]) {
      summaryRange.push(allValues[r].slice(7, 12));
    }
  }
  
  // খাবার ওয়ালাকে দেওয়া পেমেন্ট লিস্ট (H10:I23 -> Cols 7..8)
  var cateringPayments = [];
  var paidAmount = 0;
  for (var r = 9; r <= 22; r++) {
    if (allValues[r]) {
      var hVal = allValues[r][7];
      var iVal = allValues[r][8];
      cateringPayments.push([hVal, iVal]);
      var parsedVal = parseFloat(iVal);
      if (!isNaN(parsedVal)) paidAmount += parsedVal;
    }
  }
  
  // বকেয়া টাকা (J12)
  var dueAmount = (allValues[11] && allValues[11][9] !== undefined) ? allValues[11][9] : 0;
  
  // 📜 DepositLogs শিট থেকে বিগত ২ মাসের জমার হিস্ট্রি রিড করা
  var depositHistory = [];
  try {
    var logSheet = ss.getSheetByName("DepositLogs");
    if (logSheet) {
      pruneDepositLogs(logSheet);
      var lastLogRow = logSheet.getLastRow();
      if (lastLogRow > 1) {
        var logValues = logSheet.getRange(2, 1, lastLogRow - 1, 6).getValues();
        for (var i = logValues.length - 1; i >= 0; i--) {
          var row = logValues[i];
          if (row[3] && row[4] !== "") {
            var rawDate = row[2];
            var formattedDate = "";
            if (rawDate instanceof Date) {
              formattedDate = Utilities.formatDate(rawDate, "Asia/Dhaka", "yyyy-MM-dd");
            } else {
              formattedDate = String(rawDate || "");
            }
            depositHistory.push({
              "timestamp": row[0] ? String(row[0]) : "",
              "month": row[1] ? String(row[1]) : "",
              "date": formattedDate,
              "member": String(row[3]),
              "amount": parseFloat(row[4]) || 0,
              "note": row[5] ? String(row[5]) : "মেস জমা"
            });
          }
        }
      }
    }
  } catch(err) {
    Logger.log("doGet depositHistory error: " + err);
  }
  
  var result = {
    "currentMonth": sheet.getName(),
    "allMonths": getAllAvailableMonths(ss),
    "summaryTable": summaryRange,
    "dueAmount": dueAmount,
    "paidAmount": paidAmount,
    "cateringPayments": cateringPayments,
    "dailyMealsList": dailyMealsList,
    "depositHistory": depositHistory
  };
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// 📌 ৬. ইউনিভার্সাল অ্যাকশন রাউটার (doPost)
function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = JSON.parse(e.postData.contents);
  var targetMonth = data.month || "current";
  var sheet = getOrCreateSheet(ss, targetMonth);
  
  // -------------------------------------------------------------
  // অ্যাকশন ১: মেস জমা এন্ট্রি (K কলাম আপডেট + DepositLogs এ ক্লাউড হিস্ট্রি)
  // -------------------------------------------------------------
  if (data.type === 'deposit') {
    var memberNames = sheet.getRange("H2:H5").getValues();
    var selectedMember = data.member;
    var depositVal = parseFloat(data.amount) || 0;
    var depositDate = data.date || Utilities.formatDate(new Date(), "Asia/Dhaka", "yyyy-MM-dd");
    var memberRow = -1;
    
    for (var m = 0; m < memberNames.length; m++) {
      var nameInSheet = memberNames[m][0].toString().trim();
      if (nameInSheet.toLowerCase().indexOf(selectedMember.toLowerCase()) !== -1 || selectedMember.toLowerCase().indexOf(nameInSheet.toLowerCase()) !== -1) {
        memberRow = 2 + m;
        break;
      }
    }
    
    var newTotalDeposit = depositVal;
    if (memberRow !== -1) {
      var currentDepositCell = sheet.getRange("K" + memberRow);
      var previousDeposit = parseFloat(currentDepositCell.getValue()) || 0;
      newTotalDeposit = previousDeposit + depositVal;
      currentDepositCell.setValue(newTotalDeposit);
    }
    
    // DepositLogs ট্যাবে স্বয়ংক্রিয়ভাবে লেনদেন সেভ করা (সর্বোচ্চ ২ মাস সংরক্ষিত থাকবে)
    try {
      var logSheet = getOrCreateDepositLogsSheet(ss);
      logSheet.appendRow([
        Utilities.formatDate(new Date(), "Asia/Dhaka", "yyyy-MM-dd HH:mm:ss"),
        sheet.getName(),
        depositDate,
        selectedMember,
        depositVal,
        "মেস জমা"
      ]);
      pruneDepositLogs(logSheet);
    } catch(err) {
      Logger.log("Deposit log error: " + err);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      "result": "success", 
      "member": selectedMember,
      "amount": depositVal,
      "date": depositDate,
      "newTotalDeposit": newTotalDeposit
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  // -------------------------------------------------------------
  // অ্যাকশন ২: নির্দিষ্ট জমার রেকর্ড ডিলিট করা (Delete Single Transaction)
  // -------------------------------------------------------------
  else if (data.type === 'delete_deposit_txn') {
    try {
      var logSheet = ss.getSheetByName("DepositLogs");
      if (logSheet) {
        var lastLogRow = logSheet.getLastRow();
        if (lastLogRow > 1) {
          var logValues = logSheet.getRange(2, 1, lastLogRow - 1, 6).getValues();
          for (var i = logValues.length - 1; i >= 0; i--) {
            var row = logValues[i];
            var mName = String(row[3] || '').trim();
            var dAmt = parseFloat(row[4]) || 0;
            if (mName.toLowerCase().indexOf(data.member.toLowerCase()) !== -1 && dAmt === parseFloat(data.amount)) {
              logSheet.deleteRow(2 + i);
              break;
            }
          }
        }
      }
      
      // অপশনাল: শিটের K কলাম থেকেও বাদ দেওয়া
      if (data.deductFromSheet) {
        var memberNames = sheet.getRange("H2:H5").getValues();
        for (var m = 0; m < memberNames.length; m++) {
          var nameInSheet = memberNames[m][0].toString().trim();
          if (nameInSheet.toLowerCase().indexOf(data.member.toLowerCase()) !== -1 || data.member.toLowerCase().indexOf(nameInSheet.toLowerCase()) !== -1) {
            var cellK = sheet.getRange("K" + (2 + m));
            var curK = parseFloat(cellK.getValue()) || 0;
            var updatedK = Math.max(0, curK - parseFloat(data.amount));
            cellK.setValue(updatedK);
            break;
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({"result": "success"})).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({"result": "error", "message": err.toString()})).setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  // -------------------------------------------------------------
  // অ্যাকশন ৩: খাবার ওয়ালাকে পেমেন্ট এন্ট্রি (H10:I23)
  // -------------------------------------------------------------
  else if (data.type === 'catering_pay') {
    var rangeH = sheet.getRange("H10:H23").getValues();
    var rangeI = sheet.getRange("I10:I23").getValues();
    var targetRow = -1;
    for (var i = 0; i < rangeH.length; i++) {
      if ((rangeH[i][0] === "" || rangeH[i][0] === null) && (rangeI[i][0] === "" || rangeI[i][0] === null)) {
        targetRow = 10 + i;
        break;
      }
    }
    if (targetRow !== -1) {
      sheet.getRange("H" + targetRow).setValue(data.date);
      sheet.getRange("I" + targetRow).setValue(data.amount);
      return ContentService.createTextOutput(JSON.stringify({"result": "success", "row": targetRow})).setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({"result": "error", "message": "ঘর পূর্ণ হয়ে গেছে!"})).setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  // -------------------------------------------------------------
  // অ্যাকশন ৪: নির্দিষ্ট মাসের ডাটা রিসেট বা শিট ডিলিট করা
  // -------------------------------------------------------------
  else if (data.type === 'delete_month') {
    var monthToDelete = data.targetMonth;
    var targetSheet = ss.getSheetByName(monthToDelete);
    if (!targetSheet) {
      return ContentService.createTextOutput(JSON.stringify({"result": "error", "message": "মাসটি পাওয়া যায়নি!"})).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (ss.getSheets().length > 1) {
      ss.deleteSheet(targetSheet);
      return ContentService.createTextOutput(JSON.stringify({"result": "success", "deletedMonth": monthToDelete})).setMimeType(ContentService.MimeType.JSON);
    } else {
      targetSheet.getRange("B2:E32").clearContent();
      targetSheet.getRange("K2:K5").setValue(0);
      targetSheet.getRange("H10:I23").clearContent();
      return ContentService.createTextOutput(JSON.stringify({"result": "success", "resetMonth": monthToDelete})).setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  // -------------------------------------------------------------
  // অ্যাকশন ৫: নতুন মাসের শিট তৈরি করা (Create Month)
  // -------------------------------------------------------------
  else if (data.type === 'create_month') {
    var newMonthName = data.targetMonth;
    var targetSheet = getOrCreateSheet(ss, newMonthName);
    return ContentService.createTextOutput(JSON.stringify({
      "result": "success", 
      "month": targetSheet.getName(),
      "allMonths": getAllAvailableMonths(ss)
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  // -------------------------------------------------------------
  // অ্যাকশন ৬: দৈনিক মিল এন্ট্রি
  // -------------------------------------------------------------
  else {
    var datesRange = sheet.getRange("A2:A32").getDisplayValues();
    var entryDateStr = data.date;
    var entryDay = parseInt(entryDateStr.split("-")[2], 10);
    var foundRow = -1;
    
    for (var r = 0; r < datesRange.length; r++) {
      var sheetDateCell = datesRange[r][0];
      if (sheetDateCell !== "" && sheetDateCell !== null && sheetDateCell !== undefined) {
        var cellStr = sheetDateCell.toString().trim();
        var match = cellStr.match(/^0?(\d{1,2})/);
        if (match && parseInt(match[1], 10) === entryDay) {
          foundRow = 2 + r;
          break;
        }
      }
    }
    
    if (foundRow === -1 && entryDay >= 1 && entryDay <= 31) {
      foundRow = 1 + entryDay;
    }
    
    if (foundRow !== -1) {
      sheet.getRange("B" + foundRow).setValue(data.Rejaul);
      sheet.getRange("C" + foundRow).setValue(data.Tafiqul);
      sheet.getRange("D" + foundRow).setValue(data.Rafiul);
      sheet.getRange("E" + foundRow).setValue(data.Samiul);
    } else {
      sheet.appendRow([data.date, data.Rejaul, data.Tafiqul, data.Rafiul, data.Samiul]);
    }
    
    return ContentService.createTextOutput(JSON.stringify({"result": "success", "row": foundRow})).setMimeType(ContentService.MimeType.JSON);
  }
}

// 📌 ৭. প্রতি মাসের শুরুতে স্বয়ংক্রিয়ভাবে নতুন মাসের শিট তৈরি করার ট্রিগার ফাংশন
function autoCreateMonthlySheetTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentMonth = MONTH_NAMES[new Date().getMonth()];
  getOrCreateSheet(ss, currentMonth);
}
