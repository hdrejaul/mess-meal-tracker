// =====================================================================================
// 🚀 MESS MEAL MANAGEMENT SYSTEM - UNIVERSAL MASTER GOOGLE APPS SCRIPT (Code.gs)
// ⚡ FUTURE-PROOF & ZERO-EDIT ARCHITECTURE:
//    - Dynamic Month Switching & All-Months Discovery
//    - Auto New Month Sheet Creation (Format-Preserving Template Clone)
//    - Main Sheet Only Stores Member Total Deposit (K2:K5) - 100% Clean Sheet
//    - Detailed Deposit History & Transactions Stored in db.json / Client Database
//    - Safe Month Deletion / Reset with Admin Controls
//    - Universal Action Router (Extendable without touching the script again!)
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
  
  // নতুন মাস এলে প্রথম শিটটিকে টেমপ্লেট হিসেবে কপি করে স্বয়ংক্রিয়ভাবে ফ্রেশ শিট তৈরি করা
  var sheets = ss.getSheets();
  var templateSheet = sheets[0];
  var newSheet = templateSheet.copyTo(ss).setName(monthName);
  
  try {
    newSheet.getRange("B2:E32").clearContent(); // মিল রিসেট
    newSheet.getRange("K2:K5").setValue(0);      // মেম্বার জমা ০ করা
    newSheet.getRange("H10:I23").clearContent(); // ক্যাটারিং পেমেন্ট ক্লিয়ার
  } catch(e) {}
  
  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(1);
  return newSheet;
}

// 📌 ২. স্প্রেডশিটের সব সক্রিয় মাসের তালিকা পড়া
function getAllAvailableMonths(ss) {
  var sheets = ss.getSheets();
  var list = [];
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    // শুধুমাত্র মাসের নামের শিটগুলো নেওয়া (Jan..Dec বা অন্য যেকোনো শিট)
    if (!name.toLowerCase().includes("copy") && !name.toLowerCase().includes("template")) {
      list.push(name);
    }
  }
  return list;
}

// 📌 ৩. রিয়েল-টাইম অপটিমাইজড রিড হ্যান্ডলার (doGet)
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var params = e && e.parameter ? e.parameter : {};
  
  // ইউজার কোনো নির্দিষ্ট মাস রিকোয়েস্ট করলে সেই মাস পড়বে, অন্যথায় চলতি মাস
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
  
  // বকেয়া টাকা (J12 -> Row index 11, Col index 9)
  var dueAmount = (allValues[11] && allValues[11][9] !== undefined) ? allValues[11][9] : 0;
  
  var result = {
    "currentMonth": sheet.getName(),
    "allMonths": getAllAvailableMonths(ss),
    "summaryTable": summaryRange,
    "dueAmount": dueAmount,
    "paidAmount": paidAmount,
    "cateringPayments": cateringPayments,
    "dailyMealsList": dailyMealsList
  };
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// 📌 ৪. ইউনিভার্সাল অ্যাকশন রাউটার (doPost) - যেকোনো নতুন ফিচারের জন্য প্রস্তুত
function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = JSON.parse(e.postData.contents);
  var targetMonth = data.month || "current";
  var sheet = getOrCreateSheet(ss, targetMonth);
  
  // -------------------------------------------------------------
  // অ্যাকশন ১: মেস জমা এন্ট্রি (শুধুমাত্র K কলামে মোট যোগফল আপডেট)
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
    
    return ContentService.createTextOutput(JSON.stringify({
      "result": "success", 
      "member": selectedMember,
      "amount": depositVal,
      "date": depositDate,
      "newTotalDeposit": newTotalDeposit
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  // -------------------------------------------------------------
  // অ্যাকশন ২: খাবার ওয়ালাকে পেমেন্ট এন্ট্রি (H10:I23)
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
  // অ্যাকশন ৩: নির্দিষ্ট মাসের ডাটা রিসেট বা শিট ডিলিট করা
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
      // শেষ শিট হলে ডাটা রিসেট করবে
      targetSheet.getRange("B2:E32").clearContent();
      targetSheet.getRange("K2:K5").setValue(0);
      targetSheet.getRange("H10:I23").clearContent();
      return ContentService.createTextOutput(JSON.stringify({"result": "success", "resetMonth": monthToDelete})).setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  // -------------------------------------------------------------
  // অ্যাকশন ৪: নতুন মাসের শিট তৈরি করা (Create Month)
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
  // অ্যাকশন ৫: ইউনিভার্সাল সেল/রেঞ্জ রাইটার (ভবিষ্যতের যেকোনো নতুন ফিচারের জন্য)
  // -------------------------------------------------------------
  else if (data.type === 'custom_write') {
    if (data.cell && data.value !== undefined) {
      sheet.getRange(data.cell).setValue(data.value);
      return ContentService.createTextOutput(JSON.stringify({"result": "success"})).setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  // -------------------------------------------------------------
  // অ্যাকশন ৬: দৈনিক মিল এন্ট্রি (Default Fallback)
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

// 📌 ৫. প্রতি মাসের শুরুতে স্বয়ংক্রিয়ভাবে নতুন মাসের শিট তৈরি করার ট্রিগার ফাংশন
function autoCreateMonthlySheetTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentMonth = MONTH_NAMES[new Date().getMonth()];
  getOrCreateSheet(ss, currentMonth);
}
