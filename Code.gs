// =================================================================
// 🚀 MESS MEAL MANAGEMENT SYSTEM - FINAL GOOGLE APPS SCRIPT (Code.gs)
// =================================================================

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 📌 ১. চলতি মাসের নাম অনুযায়ী অটো শিট ডিটেকশন (Jan, Feb, Oct ইত্যাদি)
  var monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  var currentMonthName = monthNames[new Date().getMonth()];
  
  var sheet = ss.getSheetByName(currentMonthName) || ss.getActiveSheet();
  
  // 📌 ২. সদস্যভিত্তিক হিসাব টেবিল পড়া (H2:L7 Range)
  var summaryRange = sheet.getRange("H2:L7").getValues();
  
  // 📌 ৩. বকেয়া টাকা সেল পড়া (J12)
  var dueAmount = sheet.getRange("J12").getValue();
  
  // 📌 ৪. খাবার ওয়ালাকে দেওয়া মোট পরিশোধ (I10 থেকে I23 পর্যন্ত ডাইনামিক যোগফল)
  var paidValues = sheet.getRange("I10:I23").getValues();
  var paidAmount = 0;
  for (var i = 0; i < paidValues.length; i++) {
    var val = parseFloat(paidValues[i][0]);
    if (!isNaN(val)) paidAmount += val;
  }
  
  // 📌 ৫. খাবার ওয়ালাকে দেওয়া পেমেন্ট লিস্ট (H10:I23 Range)
  var cateringPayments = sheet.getRange("H10:I23").getValues();

  // 📌 ৬. দৈনিক ৩০ দিনের মিলের হিসাব পড়া (A2:F32 Range)
  var dailyMealsList = sheet.getRange("A2:F32").getValues();

  var result = {
    "currentMonth": sheet.getName(),
    "summaryTable": summaryRange,
    "dueAmount": dueAmount,
    "paidAmount": paidAmount,
    "cateringPayments": cateringPayments,
    "dailyMealsList": dailyMealsList
  };
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  var currentMonthName = monthNames[new Date().getMonth()];
  
  var sheet = ss.getSheetByName(currentMonthName) || ss.getActiveSheet();
  var data = JSON.parse(e.postData.contents);
  
  // 📌 টাইপ ১: খাবার ওয়ালাকে পেমেন্ট এন্ট্রি (H10:H23 ও I10:I23 এর ফাঁকা ঘরে)
  if (data.type === 'catering_pay') {
    var rangeH = sheet.getRange("H10:H23").getValues();
    var rangeI = sheet.getRange("I10:I23").getValues();
    
    var targetRow = -1;
    for (var i = 0; i < rangeH.length; i++) {
      var cellH = rangeH[i][0];
      var cellI = rangeI[i][0];
      if ((cellH === "" || cellH === null) && (cellI === "" || cellI === null)) {
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
  
  // 📌 টাইপ ২: মেস জমা এন্ট্রি (সরাসরি H2:H6 মেম্বার খুঁজে K কলামের Total Deposit এ যোগ করা)
  else if (data.type === 'deposit') {
    var memberNames = sheet.getRange("H2:H6").getValues();
    var selectedMember = data.member; // e.g. "Rejaul", "Tafiqul", "Rafiul", "Samiul", "Kader"
    var depositVal = parseFloat(data.amount) || 0;
    var memberRow = -1;
    
    // H2:H6 থেকে মেম্বারের নাম মিলিয়ে রো বের করা
    for (var m = 0; m < memberNames.length; m++) {
      var nameInSheet = memberNames[m][0].toString().trim();
      if (nameInSheet.toLowerCase().indexOf(selectedMember.toLowerCase()) !== -1 || selectedMember.toLowerCase().indexOf(nameInSheet.toLowerCase()) !== -1) {
        memberRow = 2 + m; // Row 2, 3, 4, 5, or 6
        break;
      }
    }
    
    if (memberRow !== -1) {
      // K কলাম (Total Deposit) থেকে আগের জমার মান পড়া
      var currentDepositCell = sheet.getRange("K" + memberRow);
      var previousDeposit = parseFloat(currentDepositCell.getValue()) || 0;
      
      // আগের জমার সাথে নতুন জমা যোগ করে K কলামে আপডেট করা
      currentDepositCell.setValue(previousDeposit + depositVal);
      
      return ContentService.createTextOutput(JSON.stringify({
        "result": "success", 
        "member": selectedMember,
        "newTotalDeposit": previousDeposit + depositVal
      })).setMimeType(ContentService.MimeType.JSON);
    } else {
      // ব্যাকআপ: যদি মেম্বার না মেলে তবে আলাদা রো যুক্ত করবে
      sheet.appendRow([data.date, data.member, "Deposit", data.amount]);
      return ContentService.createTextOutput(JSON.stringify({"result": "success"})).setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  // 📌 টাইপ ৩: দৈনিক মিল এন্ট্রি (A2:A32 তারিখ মিলিয়ে B-F কলামে মেম্বার মিল বসানো)
  else {
    var datesRange = sheet.getRange("A2:A32").getValues();
    var entryDateStr = data.date; // e.g. "2026-10-02"
    var entryDay = parseInt(entryDateStr.split("-")[2], 10); // Day number e.g. 2
    var foundRow = -1;
    
    for (var r = 0; r < datesRange.length; r++) {
      var sheetDateCell = datesRange[r][0];
      if (sheetDateCell !== "" && sheetDateCell !== null && sheetDateCell !== undefined) {
        if (sheetDateCell instanceof Date) {
          if (sheetDateCell.getDate() === entryDay) {
            foundRow = 2 + r;
            break;
          }
        } else {
          var cellStr = sheetDateCell.toString().trim();
          if (parseInt(cellStr, 10) === entryDay) {
            foundRow = 2 + r;
            break;
          }
          if (cellStr.indexOf(entryDateStr) !== -1 || entryDateStr.indexOf(cellStr) !== -1) {
            foundRow = 2 + r;
            break;
          }
        }
      }
    }
    
    // ব্যাকআপ ফলব্যাক: তারিখের দিন সংখ্যা অনুযায়ী সরাসরি রো (যেমন ২ তারিখ মানে row 3)
    if (foundRow === -1 && entryDay >= 1 && entryDay <= 31) {
      foundRow = 1 + entryDay;
    }
    
    if (foundRow !== -1) {
      sheet.getRange("B" + foundRow).setValue(data.Rejaul);
      sheet.getRange("C" + foundRow).setValue(data.Tafiqul);
      sheet.getRange("D" + foundRow).setValue(data.Rafiul);
      sheet.getRange("E" + foundRow).setValue(data.Samiul);
      sheet.getRange("F" + foundRow).setValue(data.Kader);
    } else {
      sheet.appendRow([data.date, data.Rejaul, data.Tafiqul, data.Rafiul, data.Samiul, data.Kader]);
    }
    
    return ContentService.createTextOutput(JSON.stringify({"result": "success", "row": foundRow})).setMimeType(ContentService.MimeType.JSON);
  }
}
