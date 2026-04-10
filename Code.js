/* ====================================================================================================
   GETTING AND SYNCHRONIZING
==================================================================================================== */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index') // 'Index' is your HTML file name
    .setTitle('Aqua Wash Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- DATA FETCHING ---
function getSheetData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  values.shift(); // Remove header row
  return values;
}

// --- DATA SAVING ---
function addDataToSheet(sheetName, rowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  sheet.appendRow(rowData);
  return "Success";
}

// --- DATA DELETION ---
function deleteDataFromSheet(sheetName, primaryValue) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] == primaryValue) {
      sheet.deleteRow(i + 1);
      return "Deleted";
    }
  }
     }

/* ====================================================================================================
   SENDING API KEY (EMAIL)
==================================================================================================== */

function sendEmailApiKey(targetEmail, apiKey) {
  const subject = "Aqua Wash - Security API Key";
  const body = "Your security API key is: " + apiKey + ". Do not share this with anyone.";
  
  try {
    MailApp.sendEmail(targetEmail, subject, body);
    return "Success";
  } catch (e) {
    return "Error: " + e.toString();
  }
}


/* ====================================================================================================
   JOURNALIZING
==================================================================================================== */

function onEdit(e) {
  const range = e.range;
  const sheetName = range.getSheet().getName();
  
  if (sheetName === "Service Transaction" || sheetName === "Supplies & Other Transaction") {
    generateJournal();
  }
}

function generateJournal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const serviceSheet = ss.getSheetByName("Service Transaction");
  const otherSheet = ss.getSheetByName("Supplies & Other Transaction");
  const journalSheet = ss.getSheetByName("Journal(Database)");

  if (!serviceSheet || !otherSheet || !journalSheet) return;

  let journalEntries = [];

  // --- 1. PROCESS SERVICE TRANSACTIONS ---
  const serviceRows = serviceSheet.getLastRow();
  if (serviceRows >= 4) {
    const serviceData = serviceSheet.getRange("B4:J" + serviceRows).getValues();
    serviceData.forEach(row => {
      const idRef = String(row[0]); // Force ID to String (Plain Text)
      const date = row[1];        
      const cash = cleanAmount(row[6]); 
      const ar = cleanAmount(row[7]);
      const desc = row[8];
      
      if (!(date instanceof Date) || !idRef) return;

      if (ar < 0) {
        const amount = Math.abs(ar);
        journalEntries.push([date, desc, "Cash", idRef, amount, "", 1]);
        journalEntries.push([date, desc, "Accounts Receivable", idRef, "", amount, 1]);
      } else {
        if (cash > 0) {
          journalEntries.push([date, desc, "Cash", idRef, cash, "", 1]);
          journalEntries.push([date, desc, "Laundry Service Revenue", idRef, "", cash, 1]);
        }
        if (ar > 0) {
          journalEntries.push([date, desc, "Accounts Receivable", idRef, ar, "", 1]);
          journalEntries.push([date, desc, "Laundry Service Revenue", idRef, "", ar, 1]);
        }
      }
    });
  }

  // --- 2. PROCESS OTHER TRANSACTIONS ---
  const otherRows = otherSheet.getLastRow();
  if (otherRows >= 4) {
    const otherData = otherSheet.getRange("B4:I" + otherRows).getValues();
    otherData.forEach(row => {
      const idRef = String(row[0]); // Force ID to String (Plain Text)
      const date = row[1];
      const desc = row[4];
      const debitAcc = row[5];
      const creditAcc = row[6];
      const amount = cleanAmount(row[7]);
      
      if (!(date instanceof Date) || !amount) return;

      journalEntries.push([date, desc, debitAcc, idRef, amount, "", 0]);
      journalEntries.push([date, desc, creditAcc, idRef, "", amount, 0]);
    });
  }

  // --- 3. SORTING ---
  journalEntries.sort((a, b) => {
    if (a[0].getTime() !== b[0].getTime()) return a[0] - b[0];
    return a[6] - b[6];
  });

  // --- 4. PREPARE OUTPUT ---
  const finalValues = [];
  const finalColors = [];

  journalEntries.forEach(entry => {
    const isDebit = entry[4] !== ""; 
    const rowColor = isDebit ? "#CFE2F3" : "#A4C2F4";
    finalValues.push([entry[0], entry[1], entry[2], entry[3], entry[4], entry[5]]);
    finalColors.push([rowColor, rowColor, rowColor, rowColor, rowColor, rowColor]);
  });

  // --- 5. WRITE & FORMAT ---
  const lastJournalRow = journalSheet.getLastRow();
  if (lastJournalRow >= 3) {
    journalSheet.getRange(3, 1, lastJournalRow, 6).clearContent().setBackground(null);
  }
  
  if (finalValues.length > 0) {
    const targetRange = journalSheet.getRange(3, 1, finalValues.length, 6);
    
    // Set Column D (ID Ref) to Plain Text format specifically
    journalSheet.getRange(3, 4, finalValues.length, 1).setNumberFormat("@");
    // Set Column A (Date) to a standard Date format
    journalSheet.getRange(3, 1, finalValues.length, 1).setNumberFormat("dd/mm/yyyy");
    
    targetRange.setValues(finalValues);
    targetRange.setBackgrounds(finalColors);
  }
}

/**
 * Ensures amounts are pure numbers for math formulas
 */
function cleanAmount(val) {
  if (typeof val === 'number') return val;
  const cleaned = String(val).replace(/[^\d.-]/g, '');
  return cleaned ? parseFloat(cleaned) : 0;
}
