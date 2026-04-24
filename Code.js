/* ====================================================================================================
   GETTING AND SYNCHRONIZING
==================================================================================================== */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index') // 'Index' is your HTML file name
    .setTitle('Aqua Wash Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- SERVICE ID FETCHING ---
function getExistingServiceIds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Service Transaction'); 
  
  if (!sheet) {
    return []; // Fixed: return is now inside the function block
  }

   const lastRow = sheet.getLastRow();
  if (lastRow < 4) return [];
  
  // Get all IDs starting from row 4 to avoid headers
  const data = sheet.getRange("B4:B" + sheet.getLastRow()).getValues(); 
  return data.flat().filter(id => id !== "");
}

// --- CUSTOMER CREDIT FETCHING ---
function getUnsettledTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Service Transaction');
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return [];

  // Fetch B to J (covering ID through Balance/Description)
  const data = sheet.getRange("B4:J" + lastRow).getValues();
  
  // 1. Find all IDs that already have a "Settlement" row
  const settledIds = new Set();
  data.forEach(row => {
    const serviceType = String(row[3]).trim(); // Column E
    const serviceId = String(row[0]).trim();  // Column B
    if (serviceType === "Settlement") {
      settledIds.add(serviceId);
    }
  });

  // 2. Filter: Must be a Credit, must have a balance > 0, and must NOT be in settledIds
  const unsettled = data.filter(row => {
    const id = String(row[0]).trim();
    const status = String(row[5]); // Column G
    const balance = cleanAmount(row[7]); // Column I
    
    // Valid credit = has "Credit" in status, has money owed, and hasn't been settled yet
    const isCredit = status.includes("Credit");
    const hasOwed = balance > 0;
    const notYetSettled = !settledIds.has(id);

    return id !== "" && isCredit && hasOwed && notYetSettled;
  });

  console.log("Found unsettled rows: " + unsettled.length);
  return unsettled; 
}

// --- OTHER TRANSACTION FETCHING --- //
function saveOtherTransaction(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Supplies & Other Transaction");
  
  if (!sheet) return "Error: Sheet Not Found";

  // 1. Find the true last row with data in Column B (Transaction No.)
  const values = sheet.getRange("B:B").getValues();
  let lastRow = 0;
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "") {
      lastRow = i + 1;
      break;
    }
  }
  
  // 2. Insert row to prevent overwriting existing formulas below the table
  sheet.insertRowAfter(lastRow);
  const destRow = lastRow + 1;
  
  // 3. Write data starting at Column B (Index 2) spanning 8 columns
  const targetRange = sheet.getRange(destRow, 2, 1, 8);
  targetRange.setValues([payload]);
  
  // 4. Force strict formatting to ensure ledger formulas don't break
  sheet.getRange(destRow, 2).setNumberFormat("@"); // Force Column B (ID) to Plain Text
  sheet.getRange(destRow, 3).setNumberFormat("dd/MM/yyyy"); // Force Column C (Date) to strict format
  
  return "Success";
}

// --- MASTERDATA FETCHING ---
function getSheetData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Matches 'Entries(MasterData)'
  if (sheetName === "Entries(MasterData)") {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const rawData = sheet.getDataRange().getValues();
  
  // Skip the first 7 rows
  const dataAfterHeader = rawData.slice(7); 
  
  // SMARTER FILTER: 
  // 1. Ignore if Column A is empty
  // 2. Ignore if Column B is empty (Real transactions always have a description)
  // 3. Ignore if the row is one of your Category Headers
  const filteredRows = dataAfterHeader.filter(row => {
    const colA = String(row[0]).trim();
    const colB = String(row[1]).trim();
    
    // List of category headers to ignore specifically
    const categories = [
      "SUPPLIES ON CASH & SUPPLIES ON CREDIT",
      "EQUIPMENT / MACHINARY OF LAUNDRY",
      "ACCOUNTS PAYABLE / LIABILITIES",
      "AQUA WASH LAUNDRY SHOP'S EQUITY",
      "EXPENSES",
      "INCOME SUMMARY / CLOSING NOMINAL ACCTS"
    ];

    return colA !== "" && colB !== "" && !categories.includes(colA);
  });

  // Return the cleaned data
  return filteredRows.map(row => [row[0], row[2], row[4], row[1]]);
}

  // 2. Matches 'Customers List(MasterData)'
  if (sheetName === "Customers List(MasterData)") {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    const rawData = sheet.getDataRange().getValues();
    const dataOnly = rawData.slice(2); 
    // Returns only Column A (Index 0)
    return dataOnly.filter(row => row[0] !== "").map(row => [row[0]]);
  }

  // 3. DEFAULT: Service Pricing and Supplies Costing
  // These only skip 1 row (the header)
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  values.shift(); 
  return values;
}

// --- FINANCIAL STATEMENT FETCHING ---
function updateAndFetchFinancials(month, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trialSheet = ss.getSheetByName("Auto Adjusted Trial Balance");
  const finSheet = ss.getSheetByName("Financial Statements");

  if (!trialSheet || !finSheet) return { error: "Sheets not found" };

  // 1. Set the Date in Trial Balance (Triggers your sheet formulas)
  trialSheet.getRange("B3").setValue(month);
  trialSheet.getRange("D3").setValue(year);
  
  // 2. Wait a moment for Google Sheets to recalculate
  SpreadsheetApp.flush();

  // 3. Fetch data from "Financial Statements" sheet
  // ADJUST THESE CELL REFERENCES TO MATCH YOUR ACTUAL SHEET LAYOUT
  return {
    period: month + " " + year,
    revenue: finSheet.getRange("C10").getValue(), // Change C10 to your Revenue cell
    expenses: finSheet.getRange("C20").getValue(), // Change C20 to your Total Expenses cell
    netIncome: finSheet.getRange("C25").getValue(), // Change C25 to your Net Income cell
    assets: finSheet.getRange("F10").getValue(),    // Change F10 to your Total Assets cell
    liabilities: finSheet.getRange("F20").getValue(),// Change F20 to your Total Liabilities cell
    equity: finSheet.getRange("F25").getValue()     // Change F25 to your Total Equity cell
  };
}

// --- DATA SAVING ---
function addDataToSheet(sheetName, rowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Sheet Not Found";
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
   DASHBOARD INPUTS (BRIDGE)
==================================================================================================== */

function addServiceOrderToSheet(sheetName, payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return "Sheet Not Found";

  // 1. Find the last row with actual data in Column B (Service ID)
  const values = sheet.getRange("B:B").getValues();
  let lastRow = 0;
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "") {
      lastRow = i + 1;
      break;
    }
  }
  
  // 2. Insert a fresh row after the last data to ensure space exists
  sheet.insertRowAfter(lastRow);
  
  const destRow = lastRow + 1;
  
  // 3. Write payload starting from Column B (Index 2)
  // Payload: [ID, Date, Name, Service Type, Price, Status, Cash, Balance, Description]
  sheet.getRange(destRow, 2, 1, payload.length).setValues([payload]);
  
  return "Success";
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
   ACCOUNTING MODULES
==================================================================================================== */

function onEdit(e) {
  const range = e.range;
  const sheetName = range.getSheet().getName();
  
  if (sheetName === "Service Transaction" || sheetName === "Supplies & Other Transaction") {
    generateJournal();
  }

  if (sheetName === "Service Transaction" || sheetName === "Supplies & Other Transaction") {
    const today = new Date();
    
    // Only attempt to run closing logic during the first 5 days of a new month
    if (today.getDate() <= 5) {
      console.log("Start of month detected. Checking for pending closing entries...");
      runAutoMonthlyClosing();
    }
  }

}

/* ====================================================================================================
   JOURNALIZING 
==================================================================================================== */


function generateJournal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const serviceSheet = ss.getSheetByName("Service Transaction");
  const otherSheet = ss.getSheetByName("Supplies & Other Transaction");
  const journalSheet = ss.getSheetByName("Journal(Database)");

  if (!serviceSheet || !otherSheet || !journalSheet) return;

  let journalEntries = [];

  // --- 1. PROCESS SERVICE TRANSACTIONS ---
  // Range B to M (Columns 2 to 13)
  const serviceRows = serviceSheet.getLastRow();
  if (serviceRows >= 4) {
    const serviceData = serviceSheet.getRange("B4:M" + serviceRows).getValues();
    serviceData.forEach(row => {
      const idRef = String(row[0]);    // Col B: Ref #
      const date = row[1];             // Col C: Date
      const cashAmt = cleanAmount(row[7]); // Col I: Cash Received
      const paymentMode = row[8];      // Col J: Payment Mode (Cash on Hand/Bank)
      const arAmt = cleanAmount(row[9]);   // Col K: Credit Balance / AR
      const desc = row[11];            // Col M: Description
      
      if (!(date instanceof Date) || !idRef) return;

      // Handle Collections (Settlement / Negative AR)
      if (arAmt < 0) {
        const amount = Math.abs(arAmt);
        const cashAccount = (paymentMode === "Cash in Bank") ? "Cash in Bank" : "Cash on Hand";
        journalEntries.push([date, desc, cashAccount, idRef, amount, "", 1]);
        journalEntries.push([date, desc, "Accounts Receivable", idRef, "", amount, 1]);
      } else {
        // Record Cash Portion
        if (cashAmt > 0) {
          const cashAccount = (paymentMode === "Cash in Bank") ? "Cash in Bank" : "Cash on Hand";
          journalEntries.push([date, desc, cashAccount, idRef, cashAmt, "", 1]);
          journalEntries.push([date, desc, "Laundry Service Revenue", idRef, "", cashAmt, 1]);
        }
        // Record Credit Portion (AR)
        if (arAmt > 0) {
          journalEntries.push([date, desc, "Accounts Receivable", idRef, arAmt, "", 1]);
          journalEntries.push([date, desc, "Laundry Service Revenue", idRef, "", arAmt, 1]);
        }
      }
    });
  }

  // --- 2. PROCESS OTHER TRANSACTIONS ---
  // Range B to I
  const otherRows = otherSheet.getLastRow();
  if (otherRows >= 4) {
    const otherData = otherSheet.getRange("B4:I" + otherRows).getValues();
    otherData.forEach(row => {
      const idRef = String(row[0]);  // Col B: Ref #
      const date = row[1];           // Col C: Date
      const desc = row[4];           // Col F: Description
      const debitAcc = row[5];       // Col G: Account (Debit)
      const creditAcc = row[6];      // Col H: Source (Credit)
      const amount = cleanAmount(row[7]); // Col I: Amount
      
      if (!(date instanceof Date) || !amount) return;

      journalEntries.push([date, desc, debitAcc, idRef, amount, "", 0]);
      journalEntries.push([date, desc, creditAcc, idRef, "", amount, 0]);
    });
  }

  // --- 3. SORTING (By Date, then by Debit/Credit order) ---
  journalEntries.sort((a, b) => {
    if (a[0].getTime() !== b[0].getTime()) return a[0] - b[0];
    return a[6] - b[6];
  });

  // --- 4. PREPARE OUTPUT ---
  const finalValues = [];
  const finalColors = [];

  journalEntries.forEach(entry => {
    const isDebit = entry[4] !== ""; 
    // RESTORED COLORS: Light Blue for Debits, Cornflower Blue for Credits
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
    journalSheet.getRange(3, 4, finalValues.length, 1).setNumberFormat("@"); // Ref # as Text
    journalSheet.getRange(3, 1, finalValues.length, 1).setNumberFormat("dd/mm/yyyy");
    
    targetRange.setValues(finalValues);
    targetRange.setBackgrounds(finalColors); // Applies the restored colors
    
    // Applies the borders to match the clean table look
    targetRange.setBorder(true, true, true, true, true, true, "#999999", SpreadsheetApp.BorderStyle.SOLID);
  }
}

/* ====================================================================================================
   AUTO CLOSING ENTRIES (MONTHLY)
==================================================================================================== */

function runAutoMonthlyClosing() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const serviceSheet = ss.getSheetByName("Service Transaction");
  const otherSheet = ss.getSheetByName("Supplies & Other Transaction");

  if (!serviceSheet || !otherSheet) return "Error: Sheets not found.";

  // 1. DETERMINE TARGET MONTH (We close the PREVIOUS month based on today's date)
  const today = new Date();
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth(); // 0-indexed (0 = Jan, 3 = April, etc.)
  
  // The date we are checking for (e.g., If today is April 1, target is March)
  const targetDate = new Date(currentYear, currentMonth - 1, 1);
  const targetMonthNum = targetDate.getMonth();
  const targetYearNum = targetDate.getFullYear();
  
  // The official date of the closing entries (1st day of the CURRENT month)
  const closingDate = new Date(currentYear, currentMonth, 1);
  
  // Format a unique Reference Prefix for the closing batch (e.g., XCL-032026)
  const refPrefix = "XCL-" + (targetMonthNum + 1).toString().padStart(2, '0') + targetYearNum;

  // 2. CHECK IF ALREADY CLOSED TO PREVENT DUPLICATES
  const otherLastRow = otherSheet.getLastRow();
  if (otherLastRow >= 4) {
    const existingRefs = otherSheet.getRange("B4:B" + otherLastRow).getValues();
    for (let i = 0; i < existingRefs.length; i++) {
      if (String(existingRefs[i][0]).startsWith(refPrefix)) {
        return "Process aborted: Closing entries for " + (targetMonthNum + 1) + "/" + targetYearNum + " already exist.";
      }
    }
  }

  // 3. INITIALIZE TEMPORARY ACCOUNT TOTALS
  let totalLaundryRevenue = 0;
  let totalOtherIncome = 0;
  let expenses = {}; 
  let totalDrawing = 0;

  // 4. FETCH LAUNDRY SERVICE REVENUE (From Service Transaction -> Price Column)
  const serviceRows = serviceSheet.getLastRow();
  if (serviceRows >= 4) {
    // Range B to M. B=0, C=1 (Date), G=5 (Price)
    const serviceData = serviceSheet.getRange("B4:M" + serviceRows).getValues();
    
    serviceData.forEach(row => {
      let dateVal = row[1];
      let price = cleanAmount(row[5]); // Col G: Price / Service Availed
      
      if (dateVal instanceof Date && dateVal.getMonth() === targetMonthNum && dateVal.getFullYear() === targetYearNum) {
        totalLaundryRevenue += price;
      }
    });
  }

  // 5. FETCH EXPENSES, OTHER INCOME, AND DRAWINGS (From Supplies & Other Transaction)
  if (otherLastRow >= 4) {
    // Range B to I. B=0, C=1 (Date), G=5 (Debit), H=6 (Credit), I=7 (Amount)
    const otherData = otherSheet.getRange("B4:I" + otherLastRow).getValues();
    
    otherData.forEach(row => {
      let dateVal = row[1];
      let debitAcc = String(row[5]).trim();
      let creditAcc = String(row[6]).trim();
      let amount = cleanAmount(row[7]);

      if (dateVal instanceof Date && dateVal.getMonth() === targetMonthNum && dateVal.getFullYear() === targetYearNum) {
        
        // Other Income (Usually a Credit)
        if (creditAcc === "Other Income") {
          totalOtherIncome += amount;
        }
        
        // Expenses (Usually a Debit)
        if (debitAcc.includes("Expense") || debitAcc.includes("Cost") || debitAcc.includes("Fee") || debitAcc.includes("Charge")) {
          if (!expenses[debitAcc]) expenses[debitAcc] = 0;
          expenses[debitAcc] += amount;
        }
        
        // Withdrawals (Debit)
        if (debitAcc === "Owner's Drawing") {
          totalDrawing += amount;
        }
      }
    });
  }

  // 6. PREPARE THE CLOSING ENTRY ROWS
  let newRows = [];
  let totalRevenues = totalLaundryRevenue + totalOtherIncome;
  let totalExpenses = 0;
  
  const party = "Aqua Wash Laundry Shop";
  const paymentMode = "N/A"; // Enforcing NO CASH condition
  let entryCount = 1;

  // A. Close Revenues (Debit Revenue, Credit Income Summary)
  if (totalLaundryRevenue > 0) {
    newRows.push([refPrefix + "-" + entryCount++, closingDate, party, "Revenue Closing", "To record closing of revenue account", "Laundry Service Revenue", "Income Summary", totalLaundryRevenue, paymentMode]);
  }
  if (totalOtherIncome > 0) {
    newRows.push([refPrefix + "-" + entryCount++, closingDate, party, "Other Income Closing", "To record closing of other revenue", "Other Income", "Income Summary", totalOtherIncome, paymentMode]);
  }

  // B. Close Expenses (Debit Income Summary, Credit Expenses)
  for (const [expName, expAmount] of Object.entries(expenses)) {
    if (expAmount > 0) {
      newRows.push([refPrefix + "-" + entryCount++, closingDate, party, "Expense Closing", "To record closing of " + expName.toLowerCase(), "Income Summary", expName, expAmount, paymentMode]);
      totalExpenses += expAmount;
    }
  }

  // C. Close Owner's Drawing Directly to Capital (Debit Capital, Credit Drawing)
  if (totalDrawing > 0) {
    newRows.push([refPrefix + "-D", closingDate, party, "Withdrawal Closing", "To record closing of drawing account", "Owner's Capital", "Owner's Drawing", totalDrawing, paymentMode]);
  }

  // D. Close Income Summary to Capital (Net Income / Loss)
  let netIncome = totalRevenues - totalExpenses;
  if (netIncome > 0) {
    newRows.push([refPrefix + "-NI", closingDate, party, "Increase in Capital", "To record net income closed to capital", "Income Summary", "Owner's Capital", netIncome, paymentMode]);
  } else if (netIncome < 0) {
    newRows.push([refPrefix + "-NL", closingDate, party, "Decrease in Capital", "To record net loss closed to capital", "Owner's Capital", "Income Summary", Math.abs(netIncome), paymentMode]);
  }

  // 7. INJECT INTO SHEET & TRIGGER JOURNAL
  if (newRows.length > 0) {
    const destRow = otherSheet.getLastRow() + 1;
    
    // Writing to B through J (9 Columns)
    const targetRange = otherSheet.getRange(destRow, 2, newRows.length, 9);
    targetRange.setValues(newRows);
    
    // Force strict formatting
    otherSheet.getRange(destRow, 2, newRows.length, 1).setNumberFormat("@"); // Refs to Plain Text
    otherSheet.getRange(destRow, 3, newRows.length, 1).setNumberFormat("dd/MM/yyyy"); // Format Date
    
    // Auto-update the Journal
    generateJournal();
    
    return "Success: " + newRows.length + " closing entries generated for " + (targetMonthNum + 1) + "/" + targetYearNum;
  } else {
    return "No records found to close for " + (targetMonthNum + 1) + "/" + targetYearNum;
  }
}


function cleanAmount(val) {
  if (typeof val === 'number') return val;
  const cleaned = String(val).replace(/[^\d.-]/g, '');
  return cleaned ? parseFloat(cleaned) : 0;
}
