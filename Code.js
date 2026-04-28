/* ====================================================================================================
   GETTING AND SYNCHRONIZING
==================================================================================================== */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index') // 'Index' is your HTML file name
    .setTitle('Aqua Wash Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

const AUTHORIZED_EMAILS = [
  "jamandreprince@gmail.com",         // Replace with actual allowed emails
  "princejohnley19@gmail.com" 
];

function checkAuth() {
  const userEmail = Session.getActiveUser().getEmail();
  // Allow if email is in whitelist, or if email is blank (sometimes happens in dev environment)
  if (userEmail !== "" && !AUTHORIZED_EMAILS.includes(userEmail)) {
    return false;
  }
  return true;
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

// --- CUSTOMER PAID FETCHING ---
function getInvoiceTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stSheet = ss.getSheetByName('Service Transaction');
  if (!stSheet) return [];

  const stLastRow = Math.max(stSheet.getLastRow(), 4);
  // Using getDisplayValues to keep the exact Date formatting from the sheet
  const stData = stSheet.getRange("B4:J" + stLastRow).getDisplayValues(); 
  
  let groupedInvoices = {};

  stData.forEach(row => {
    const ref = String(row[0]).trim();
    if (!ref) return;

    const date = String(row[1]).trim();
    const surname = String(row[2]).trim();
    const firstName = String(row[3]).trim();
    const fullName = `${firstName} ${surname}`.trim();
    
    const type = String(row[5]).trim();
    
    // Clean formatting (removes commas and currency symbols) before calculating
    const price = parseFloat(row[6].replace(/[^0-9.-]+/g,"")) || 0;
    const payment = parseFloat(row[7].replace(/[^0-9.-]+/g,"")) || 0;

    // Initialize the group if this Ref# is seen for the first time
    if (!groupedInvoices[ref]) {
      groupedInvoices[ref] = {
        refId: ref,
        fullName: fullName,
        date: '',
        serviceType: '',
        servicePrice: 0,
        totalPayment: 0,
        balance: 0
      };
    }

    // Add every cash received (including Settlements) to the Total Payment
    groupedInvoices[ref].totalPayment += payment;

    // If it's NOT a settlement, grab the original details (Date, Type, Price)
    if (type !== 'Settlement' && !groupedInvoices[ref].serviceType) {
      groupedInvoices[ref].date = date;
      groupedInvoices[ref].serviceType = type;
      groupedInvoices[ref].servicePrice = price;
    }
  });

  const allTransactions = [];

  // Calculate balances and push to the final array
  for (const ref in groupedInvoices) {
    const inv = groupedInvoices[ref];
    
    // Calculate balance and round to 2 decimals to prevent floating-point bugs
    inv.balance = Math.round((inv.servicePrice - inv.totalPayment) * 100) / 100;
    
    // Only include it if it has a valid original service (avoids orphaned settlements)
    if (inv.serviceType) {
      allTransactions.push(inv);
    }
  }

  return allTransactions;
}

// --- CUSTOMER CREDIT FETCHING ---
function getUnsettledTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Pre-fetch original Names from Service Transaction to separate Surname and First Name
  const stSheet = ss.getSheetByName('Service Transaction');
  let nameMap = {};
  if (stSheet) {
    const stLastRow = stSheet.getLastRow();
    if (stLastRow >= 4) {
      // Fetch B4:E (Ref #, Date, Surname, FirstName)
      const stData = stSheet.getRange("B4:E" + stLastRow).getValues();
      stData.forEach(row => {
        const ref = String(row[0]).trim();
        if (ref) {
          nameMap[ref] = {
            surname: String(row[2]).trim(),
            firstName: String(row[3]).trim()
          };
        }
      });
    }
  }

  // 2. Fetch Truth/Balances from Balance Tracker(Checker)
  const btSheet = ss.getSheetByName('Balance Tracker(Checker)');
  if (!btSheet) return [];
  
  const btLastRow = btSheet.getLastRow();
  if (btLastRow < 3) return [];

  // CHANGED: Now fetching from A to F
  // Array Indices map: A (0), B (1), C (2), D (3), E (4), F (5)
  const btData = btSheet.getRange("A3:F" + btLastRow).getValues();
  const unsettled = [];

  // Helper function to safely handle "(35.00)" accounting formats as true Math
  function getCleanNumber(val) {
    if (typeof val === 'number') return val;
    let s = String(val).replace(/,/g, '').trim();
    if (s.startsWith('(') && s.endsWith(')')) s = "-" + s.slice(1, -1);
    let n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  }

  btData.forEach(row => {
    const ref = String(row[0]).trim();       // Col A: Ref #
    const status = String(row[5]).trim();    // Col F: Status

    // Only target Outstanding Balances 
    if (ref && status === "With Outstanding Balance") {
      const ar = Math.abs(getCleanNumber(row[3]));       // Col D: AR
      const payment = Math.abs(getCleanNumber(row[4]));  // Col E: Payment
      const balance = ar - payment; // Ensures precision logic difference

      if (balance > 0) {
        // Fallback: grab directly from Col B and C in Tracker if needed
        let surname = String(row[1]).trim();
        let firstName = String(row[2]).trim();

        // Give priority to the perfectly split names mapped from Service Transaction
        if (nameMap[ref] && (nameMap[ref].surname || nameMap[ref].firstName)) {
          surname = nameMap[ref].surname;
          firstName = nameMap[ref].firstName;
        } 

        const fullName = `${firstName} ${surname}`.trim();

        unsettled.push({
          refId: ref,
          surname: surname,
          firstName: firstName,
          fullName: fullName,
          balance: balance
        });
      }
    }
  });

  return unsettled;
}

/* ====================================================================================================
   SUPPLIES FETCHING AND PROCESSING 
==================================================================================================== */
function getSuppliesInventory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Balance Tracker(Checker)");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  
  // Range: Start row 3, Col 8 (H), 13 rows down, 8 columns wide (H to O)
  const data = sheet.getRange(3, 8, lastRow - 2, 8).getValues();
  
  return data.map(row => ({
    id: row[0],
    brand: row[1],
    costPerUnit: row[2],
    qty: row[3],
    totalPhysical: row[4], // Col L
    totalJournal: row[5],  // Col M
    status: row[6],        // Col N
    adjustment: row[7]     // Col O
  })).filter(item => item.id !== "");
}

/**
 * PROCESS: Main Entry Point for Supplies
 */
function processSuppliesTransaction(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName("Supplies & Other Transaction");
  const today = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");

  // 1. Calculate final amount based on "Supplies Used" logic
  let finalAmount = payload.amount;
  if (payload.type === "Supplies Used" && payload.isUnusedMode) {
    // Logic: Beginning (Journal) Balance - Ending (Unused) Input
    finalAmount = payload.journalBalance - payload.amount;
  }

  // 2. RECORD TO: Supplies & Other Transaction (Segmented Write)
  recordToSuppliesLogSheet(transSheet, payload.id, today, payload.party, payload.type, finalAmount, payload.payment);

  // 3. UPDATE: Balance Tracker (Checker) - Update the Quantity
  updateSuppliesInventoryQty(payload.id, payload.type, payload.qtyChange, finalAmount);

  return payload.id;
}

/**
 * LOGIC: Segmented Write to skip Columns F, G, H
 */
function recordToSuppliesLogSheet(sheet, id, date, party, type, amount, payment) {
  const leftSide = [[id, date, party, type]]; // B to E
  const rightSide = [[amount, payment]];     // I to J

  const colB = sheet.getRange("B:B").getValues();
  let targetRow = 0;
  for (let i = 3; i < colB.length; i++) {
    if (colB[i][0] === "" || colB[i][0] === null) { targetRow = i + 1; break; }
  }
  if (targetRow === 0) targetRow = colB.length + 1;

  // Segmented Write
  sheet.getRange(targetRow, 2, 1, 4).setValues(leftSide);
  sheet.getRange(targetRow, 9, 1, 2).setValues(rightSide);
}

/**
 * UPDATE: Adjust Quantity in Balance Tracker
 */
function updateSuppliesInventoryQty(id, type, qtyChange, amount) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Balance Tracker(Checker)");
  const data = sheet.getRange("H3:H" + sheet.getLastRow()).getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === id) {
      let targetRow = i + 3;
      let currentQty = sheet.getRange(targetRow, 11).getValue() || 0; // Col K
      
      // If used, subtract. If purchased, add.
      let newQty = (type === "Supplies Used") ? currentQty - qtyChange : currentQty + qtyChange;
      sheet.getRange(targetRow, 11).setValue(newQty); 
      break;
    }
  }
}

/* ====================================================================================================
   EQUIPMENT PROCESSING AND FETCHING 
==================================================================================================== */
// --- PROCESSING EQUIPMENT 
function processEquipmentTransaction(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName("Supplies & Other Transaction");
  const balSheet = ss.getSheetByName("Balance Tracker(Checker)");
  
  // 1. Generate the unique ID or use existing for sales
  let transactionId = generateEqptId(transSheet, data);

  // 2. Standardize Date (dd/mm/yyyy)
  const today = new Date();
  const formattedDate = Utilities.formatDate(today, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");

  // 3. RECORD TO: Supplies & Other Transaction
  recordToEqptOtherSheet(transSheet, transactionId, formattedDate, data);

  // 4. RECORD TO: Balance Tracker (Acquisitions Only)
  if (data.type === "Equipment Acquired" || data.type === "Eqpt. Acquired on Credit") {
    recordToBalanceSheet(balSheet, transactionId, formattedDate, data);
  }

  return transactionId;
}

/**
 * LOGIC FOR: Supplies & Other Transaction
 * Handles multi-row (compound) entries for sales.
 */
function recordToEqptOtherSheet(transSheet, id, date, data) {
  let leftSideRows = [];  // Columns B, C, D, E
  let rightSideRows = []; // Columns I, J
   const internalParty = "Aqua Wash Laundry Shop";

  // Prepare Data Rows
  if (data.type === "Equipment Sold") {
    leftSideRows.push([id, date, data.party, data.type]);
    rightSideRows.push([data.amount, data.payment]);
    leftSideRows.push([id, date, internalParty, "Accumulated Removal"]);
    rightSideRows.push([data.accDep, "N/A"]);
    
    let diff = data.amount - data.carrying;
    let glType = diff >= 0 ? "Gain on Equipment Sale" : "Loss on Equipment Sale";
    leftSideRows.push([id, date, internalParty, glType]);
    rightSideRows.push([Math.abs(diff), data.payment]);
  } else {
    leftSideRows.push([id, date, data.party, data.type]);
    rightSideRows.push([data.amount, data.payment]);
  }

  // 1. FIND TARGET ROW (Gap filling or Append)
  const colB = sheet.getRange("B:B").getValues();
  let targetRow = 0; 
  for (let i = 3; i < colB.length; i++) {
    if (colB[i][0] === "" || colB[i][0] === null) {
      targetRow = i + 1;
      break;
    }
  }
  // If no empty row was found in the entire existing sheet
  if (targetRow === 0) targetRow = colB.length + 1;

  // 2. SAFETY CHECK: Insert rows if we are going beyond the sheet limit
  const maxRows = sheet.getMaxRows();
  const neededRows = targetRow + leftSideRows.length - 1;
  if (neededRows > maxRows) {
    sheet.insertRowsAfter(maxRows, neededRows - maxRows);
  }

  // 3. SEGMENTED WRITE: Preserves formulas in F, G, H
  sheet.getRange(targetRow, 2, leftSideRows.length, 4).setValues(leftSideRows); // B to E
  sheet.getRange(targetRow, 9, rightSideRows.length, 2).setValues(rightSideRows); // I to J
}

/**
 * LOGIC FOR: Balance Tracker(Checker)
 * Only records the initial acquisition data.
 */
function recordToBalanceSheet(sheet, id, date, data) {
  const colQ = sheet.getRange("Q:Q").getValues();
  let targetRow = 0;
  for (let i = 2; i < colQ.length; i++) {
    if (colQ[i][0] === "" || colQ[i][0] === null) {
      targetRow = i + 1;
      break;
    }
  }
  if (targetRow === 0) targetRow = colQ.length + 1;

  const maxRows = sheet.getMaxRows();
  if (targetRow > maxRows) {
    sheet.insertRowsAfter(maxRows, 1);
  }

  // Write: ID(Q), Date(R), Cost(S-null), Scrap(T), Life(U)
  sheet.getRange(targetRow, 17, 1, 5).setValues([[id, date, null, data.scrapValue, data.usefulLife]]);
}

/**
 * HELPER: Generate EQPT ID
 */
function generateEqptId(sheet, data) {
  if (data.type === "Equipment Sold") return data.eqptId;

  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return "EQPT-0001"; // Default if sheet is empty

  const ids = sheet.getRange("B4:B" + lastRow).getValues().flat();
  let maxId = 0;
  ids.forEach(val => {
    if (String(val).startsWith("EQPT-")) {
      let num = parseInt(val.replace("EQPT-", ""), 10);
      if (num > maxId) maxId = num;
    }
  });
  return "EQPT-" + String(maxId + 1).padStart(4, '0');
}

// --- EQUIPMENT FETCHING

function getEquipmentInventory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Balance Tracker(Checker)");
  if (!sheet) return [];

  // Q=17, R=18, S=19, W=23, X=24, AB=28
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  
  const range = sheet.getRange(3, 17, lastRow - 2, 12).getValues(); 
  
  return range.map(row => {
    // Force dates into exact dd/mm/yyyy format
    let rawDate = row[1];
    let formattedDate = rawDate;
    if (rawDate instanceof Date) {
      let dd = String(rawDate.getDate()).padStart(2, '0');
      let mm = String(rawDate.getMonth() + 1).padStart(2, '0');
      let yyyy = rawDate.getFullYear();
      formattedDate = `${dd}/${mm}/${yyyy}`;
    }

    return {
      id: row[0],          // Col Q 
      date: formattedDate, // Col R (Formatted)
      cost: row[2],        // Col S 
      accDep: row[6],      // Col W (Needed for Accumulated Removal)
      carrying: row[7],    // Col X 
      status: row[11]      // Col AB 
    };
  });
}


/* ====================================================================================================
   OTHER TRANSACTIONS FETCHING
==================================================================================================== */
function saveOtherTransaction(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Supplies & Other Transaction");
  const tAccountSheet = ss.getSheetByName("T-Accounts(Database2)");
  
  if (!sheet || !tAccountSheet) return "Error: Required Sheets Not Found";

  // --- 1. EXTRACT PAYLOAD DATA ---
  const tId = payload[0];
  const tWith = payload[2];
  const tType = payload[3];
  const tAmt = parseFloat(payload[7]) || 0;
  const tPayMode = payload[8];

  // --- 2. GET LIVE BALANCES FROM T-ACCOUNTS ---
  // D3 = Cash on Hand, H3 = Cash in Bank, AB3 = Accounts Payable
  const cashOnHandBal = parseFloat(tAccountSheet.getRange("D3").getValue()) || 0;
  const cashInBankBal = parseFloat(tAccountSheet.getRange("H3").getValue()) || 0;
  const accountsPayableBal = parseFloat(tAccountSheet.getRange("AB3").getValue()) || 0;

  // --- 3. SPECIFIC RESTRICTIONS PER YOUR RULES ---

  // Rule: Debt Payment validation against Accounts Payable (Cell AB3)
  if (tType === "Debt Payment") {
    if (tAmt > accountsPayableBal) {
      return `ERROR: The Debt Payment (₱${tAmt.toFixed(2)}) exceeds the current Accounts Payable balance (₱${accountsPayableBal.toFixed(2)}).`;
    }
  }

  // Rule: Transfer validations (Cells D3 and H3)
  if (tType === "Cash on hand to Cash in bank") {
    if (tAmt > cashOnHandBal) {
      return `ERROR: Transfer amount (₱${tAmt.toFixed(2)}) exceeds available Cash on Hand (₱${cashOnHandBal.toFixed(2)}).`;
    }
  }

  if (tType === "Cash in bank to Cash on hand") {
    if (tAmt > cashInBankBal) {
      return `ERROR: Transfer amount (₱${tAmt.toFixed(2)}) exceeds available Cash in Bank (₱${cashInBankBal.toFixed(2)}).`;
    }
  }

  // --- 4. GENERAL CASH OUTFLOW VALIDATION ---
  // This checks any transaction (like Supplies Payment) that reduces cash
  const masterSheet = ss.getSheetByName("Entries(MasterData)");
  let creditAcc = "";
  
  if (masterSheet) {
    const mData = masterSheet.getDataRange().getValues();
    for (let i = 7; i < mData.length; i++) {
      if (String(mData[i][0]).trim() === String(tType).trim()) {
        creditAcc = String(mData[i][4]).trim(); // Column E: Source (Credit)
        break;
      }
    }
  }

  // Check general outflow based on Payment Mode or Credit Account
  if (creditAcc === "Cash on Hand" || tPayMode === "Cash on Hand") {
    // Only check if it's not a transfer already handled above
    if (tType !== "Cash on hand to Cash in bank" && tAmt > cashOnHandBal) {
      return `ERROR: Insufficient Cash on Hand. Current Balance: ₱${cashOnHandBal.toFixed(2)}`;
    }
  }

  if (creditAcc === "Cash in Bank" || tPayMode === "Cash in Bank") {
    if (tType !== "Cash in bank to Cash on hand" && tAmt > cashInBankBal) {
      return `ERROR: Insufficient Cash in Bank. Current Balance: ₱${cashInBankBal.toFixed(2)}`;
    }
  }

  // --- 5. EXECUTE SAVE ---
  const bValues = sheet.getRange("B4:B").getValues();
  let destRow = -1;
  
  for (let i = 0; i < bValues.length; i++) {
    if (String(bValues[i][0]).trim() === "") {
      destRow = i + 4;
      break;
    }
  }

  if (destRow === -1) {
    destRow = sheet.getLastRow() + 1;
  }

  sheet.getRange(destRow, 2, 1, 9).setValues([payload]);
  sheet.getRange(destRow, 2).setNumberFormat("@"); 
  sheet.getRange(destRow, 3).setNumberFormat("dd/MM/yyyy"); 

  return "Success";
}

/* ====================================================================================================
   MASTERDATA FETCHING
==================================================================================================== */
function getSheetData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Entries(MasterData)'
  if (sheetName === "Entries(MasterData)") {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    const rawData = sheet.getDataRange().getValues();
    
    // Skip the first 7 rows
    const dataAfterHeader = rawData.slice(7); 
    
    // Clean data and filter out category headers
    const filteredRows = dataAfterHeader.filter(row => {
      const colA = String(row[0]).trim();
      const colB = String(row[1]).trim();
      
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

    return filteredRows.map(row => [row[0], row[2], row[4], row[1]]);
  }

  // 2. 'Records List(MasterData)' SHEET
  const recordsSheet = ss.getSheetByName("Records List(MasterData)");
  if (!recordsSheet) return [];
  
  const rawData = recordsSheet.getDataRange().getValues();
  // We slice starting at index 3 (Row 4 in Sheets) to skip the 3 header rows
  const dataRows = rawData.slice(3); 

  // --- TARGET: CUSTOMERS ---
  if (sheetName === "Customers List(MasterData)" || sheetName === "Customers") {
    return dataRows
      .filter(row => String(row[0]).trim() !== "" || String(row[1]).trim() !== "") 
      .map(row => {
        const surname = String(row[0]).trim();
        const firstName = String(row[1]).trim();
        const fullName = `${firstName} ${surname}`.trim(); 
        
        // Return all three items!
        // [0] = Full Name, [1] = First Name, [2] = Surname
        return [fullName, firstName, surname]; 
      });
  }
  
  // --- TARGET: SERVICES ---
  if (sheetName === "Service Pricing(MasterData)" || sheetName === "Services") {
    return dataRows
      .filter(row => String(row[7]).trim() !== "") // Col H is Index 7 (Type)
      .map(row => [row[7], row[8]]); // Col H (7) and Col I (8)
  }
  
  // --- TARGET: SUPPLIES ---
  if (sheetName === "Supplies Costing(MasterData)" || sheetName === "Supplies") {
    return dataRows
      .filter(row => String(row[11]).trim() !== "") // Col L is Index 11 (Brandname)
      .map(row => [row[11], row[12]]); // Col L (11) and Col M (12)
  }

  // DEFAULT FALLBACK: Just return the raw values if sheet names don't match
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  values.shift(); 
  return values;
}

/* ====================================================================================================
   FINANCIAL STATEMENT FETCHING
==================================================================================================== */
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

// --- DATA SAVING MASTERDATA ---
function addDataToSheet(targetCategory, rowData) {
  if (!checkAuth()) return "Unauthorized: Your email is not whitelisted.";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Handle "Other Transactions"
  if (targetCategory === "Entries(MasterData)") {
    const sheet = ss.getSheetByName("Entries(MasterData)");
    if (!sheet) return "Sheet Not Found";
    sheet.appendRow(rowData);
    return "Success";
  }

  // Handle the Merged "Records List" Sheet
  const sheet = ss.getSheetByName("Records List(MasterData)");
  if (!sheet) return "Sheet Not Found";

  let startCol;
  if (targetCategory.includes("Customer")) {
    startCol = 1;  // Col A
  } 
  else if (targetCategory.includes("Service")) {
    startCol = 8;  // Col H
  } 
  else if (targetCategory.includes("Supplies")) {
    startCol = 11; // Col K (ID Column)

    // --- AUTO ID GENERATION LOGIC ---
    // Scan Column K from Row 4 downwards to find the highest SUPS number
    const idData = sheet.getRange(4, 11, sheet.getLastRow()).getValues();
    let maxNum = 0;
    
    for (let i = 0; i < idData.length; i++) {
      let currentId = String(idData[i][0]).trim();
      if (currentId.startsWith("SUPS-")) {
        // Extract the number part (e.g., "0003" -> 3)
        let num = parseInt(currentId.replace("SUPS-", ""), 10);
        if (num > maxNum) maxNum = num;
      }
    }
    
    // Generate new ID (e.g., if max is 3, next is "SUPS-0004")
    let newId = "SUPS-" + String(maxNum + 1).padStart(4, '0');
    
    // Push the new ID to the beginning of the rowData array
    // rowData goes from [Brandname, Cost] -> [ID, Brandname, Cost]
    rowData.unshift(newId);
  } 
  else {
    return "Unknown Category";
  }

  // Find the first empty row in THAT specific column starting from Row 4
  const data = sheet.getRange(4, startCol, sheet.getLastRow()).getValues();
  let targetRow = 4;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === "") {
      targetRow = i + 4;
      break;
    }
    if (i === data.length - 1) targetRow = sheet.getLastRow() + 1;
  }

  // Set the values in the specific columns
  sheet.getRange(targetRow, startCol, 1, rowData.length).setValues([rowData]);
  return "Success";
}


// --- DATA DELETION MASTERDATA ---
function deleteDataFromSheet(targetCategory, primaryValue) {
  if (!checkAuth()) return "Unauthorized: Your email is not whitelisted.";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = targetCategory.includes("Entries") ? "Entries(MasterData)" : "Records List(MasterData)";
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Sheet Not Found";

  const data = sheet.getDataRange().getValues();
  let searchColIndex = 0; 
  
  if (targetCategory.includes("Service")) searchColIndex = 7; // Col H
  else if (targetCategory.includes("Supplies")) searchColIndex = 11; // Col L (Brandname)

  for (let i = data.length - 1; i >= 0; i--) {
    let isMatch = false;

    // CUSTOMER MATCHING LOGIC
    if (targetCategory.includes("Customer")) {
      // Recreate the Firstname + Surname string to match what the frontend sent
      const surname = String(data[i][0]).trim();    // Col A
      const firstName = String(data[i][1]).trim();  // Col B
      const fullNameInSheet = `${firstName} ${surname}`.trim();
      
      if (fullNameInSheet === String(primaryValue).trim()) {
        isMatch = true;
      }
    } 
    // ALL OTHER CATEGORIES MATCHING LOGIC
    else {
      if (String(data[i][searchColIndex]).trim() === String(primaryValue).trim()) {
        isMatch = true;
      }
    }

    // IF A MATCH IS FOUND, DELETE IT
    if (isMatch) {
      if (sheetName === "Records List(MasterData)") {
        // Clear specific number of columns based on category
        let numCols = 2; 
        if (targetCategory.includes("Supplies")) {
          // If Supply, clear Col K, L, M (ID, Brand, Cost)
          searchColIndex = 10; // Shift back to Col K to clear the ID too
          numCols = 3;
        }
        
        // Clear the specific cells instead of deleting the whole row
        sheet.getRange(i + 1, searchColIndex + 1, 1, numCols).clearContent();
      } else {
        sheet.deleteRow(i + 1);
      }
      return "Deleted";
    }
  }
  return "Not Found";
}

/* ====================================================================================================
   DASHBOARD INPUTS (BRIDGE) SERVICE CUTEII (ADD/DELETE)
==================================================================================================== */

function addServiceOrderToSheet(sheetName, payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return "Sheet Not Found";

  // 1. Find the last row with data in Column B
  const values = sheet.getRange("B:B").getValues();
  let lastRow = 3; 
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "") {
      lastRow = i + 1;
      break;
    }
  }
  
  const destRow = Math.max(lastRow + 1, 4);
  
  // 2. Insert the row
  sheet.insertRowAfter(lastRow);
  
  // 3. Write the payload
  sheet.getRange(destRow, 1, 1, payload.length).setValues([payload]);

  // --- NEW: MERGE E AND F FOR THE NEW ROW ---
  // Column E is index 5, and we want to merge 2 columns across (E and F)
  sheet.getRange(destRow, 5, 1, 2).mergeAcross();

  // Optional: Center the text in the merged cell
  sheet.getRange(destRow, 5).setHorizontalAlignment("left");
  
  return "Success";
}

function deleteOrderById(refId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Service Transaction');
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return "No data to delete.";

  const data = sheet.getRange("B1:B" + lastRow).getValues();
  
  // Backward loop is essential when deleting rows to maintain index accuracy
  for (let i = data.length - 1; i >= 0; i--) {
    if (String(data[i][0]).trim() === String(refId).trim()) {
      sheet.deleteRow(i + 1);
    }
  }
  return "All records for " + refId + " deleted.";
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
