/**
 * Automated Journalizing for Aqua Wash Laundry Shop
 * Now includes color coding for Debit and Credit rows.
 */

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
  const journalSheet = ss.getSheetByName("Journal");

  if (!serviceSheet || !otherSheet || !journalSheet) return;

  let journalEntries = [];

  // --- 1. PROCESS SERVICE TRANSACTIONS ---
  const serviceRows = serviceSheet.getLastRow();
  if (serviceRows >= 4) {
    const serviceData = serviceSheet.getRange("B4:J" + serviceRows).getValues();
    serviceData.forEach(row => {
      const idRef = row[0]; const date = row[1] ? new Date(row[1]) : ""; const cash = row[6]; 
      const ar = row[7]; const desc = row[8];
      if (!date || !idRef) return;

      if (ar < 0) {
        const amount = Math.abs(ar);
        journalEntries.push([date, desc, "Cash", idRef, amount, "", 1]); // Debit
        journalEntries.push([date, desc, "Accounts Receivable", idRef, "", amount, 1]); // Credit
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
      const idRef = row[0]; const date = row[1] ? new Date(row[1]) : ""; const desc = row[4];
      const debitAcc = row[5]; const creditAcc = row[6]; const amount = row[7];
      if (!date || !amount) return;

      journalEntries.push([date, desc, debitAcc, idRef, amount, "", 0]);
      journalEntries.push([date, desc, creditAcc, idRef, "", amount, 0]);
    });
  }

  // --- 3. SORTING ---
  journalEntries.sort((a, b) => {
    const dateA = new Date(a[0]);
    const dateB = new Date(b[0]);
    if (dateA.getTime() !== dateB.getTime()) return dateA - dateB;
    return a[6] - b[6];
  });

  // --- 4. PREPARE OUTPUT & COLORS ---
  const finalValues = [];
  const finalColors = [];

  journalEntries.forEach(entry => {
    // Column 4 is Debit, Column 5 is Credit (0-indexed)
    const isDebit = entry[4] !== ""; 
    const rowColor = isDebit ? "#CFE2F3" : "#A4C2F4";
    
    finalValues.push([entry[0], entry[1], entry[2], entry[3], entry[4], entry[5]]);
    // Apply the same color to all 6 columns of the row
    finalColors.push([rowColor, rowColor, rowColor, rowColor, rowColor, rowColor]);
  });

  // --- 5. WRITE TO SHEET ---
  const lastJournalRow = journalSheet.getLastRow();
  // Clear everything (content + formatting) from row 3 down
  if (lastJournalRow >= 3) {
    journalSheet.getRange(3, 1, lastJournalRow, 6).clearContent().setBackground(null);
  }
  
  if (finalValues.length > 0) {
    const targetRange = journalSheet.getRange(3, 1, finalValues.length, 6);
    targetRange.setValues(finalValues);
    targetRange.setBackgrounds(finalColors);
  }
}
