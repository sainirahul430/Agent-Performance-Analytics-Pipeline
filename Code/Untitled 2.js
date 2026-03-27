function debugCallIDs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CDR_Data');
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const callIdIdx = headers.findIndex(h => 
    h && h.toString().toLowerCase().replace(/\s|_/g, '') === 'callid'
  );
  
  Logger.log('CallID column index: ' + callIdIdx);
  Logger.log('CallID header: ' + headers[callIdIdx]);
  Logger.log('---');
  
  // Get first 10 rows of CallID data
  const data = sheet.getRange(2, callIdIdx + 1, Math.min(10, sheet.getLastRow() - 1), 1).getValues();
  
  data.forEach((row, i) => {
    const val = row[0];
    Logger.log(`Row ${i+2}: "${val}" | Type: ${typeof val} | Length: ${String(val).length}`);
  });
  
  Logger.log('---');
  Logger.log('Total rows: ' + (sheet.getLastRow() - 1));
}




function debugDuplicates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CDR_Data');
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const callIdIdx = headers.findIndex(h => 
    h && h.toString().toLowerCase().replace(/\s|_/g, '') === 'callid'
  );
  
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, callIdIdx + 1, lastRow - 1, 1).getValues();
  
  // Track CallIDs
  const seenCallIds = new Set();
  const duplicates = [];
  
  data.forEach((row, idx) => {
    const rawCallId = row[0];
    if (!rawCallId) return;
    
    const callId = String(rawCallId).trim();
    
    if (seenCallIds.has(callId)) {
      duplicates.push({
        row: idx + 2,
        callId: callId
      });
    }
    seenCallIds.add(callId);
  });
  
  Logger.log('=== DUPLICATE CHECK ===');
  Logger.log('Total rows in sheet: ' + (lastRow - 1));
  Logger.log('Unique CallIDs: ' + seenCallIds.size);
  Logger.log('Duplicate count: ' + duplicates.length);
  Logger.log('---');
  
  if (duplicates.length > 0) {
    Logger.log('First 10 duplicates:');
    duplicates.slice(0, 10).forEach(d => {
      Logger.log(`Row ${d.row}: ${d.callId}`);
    });
  } else {
    Logger.log('NO DUPLICATES FOUND in raw data!');
    Logger.log('The inflated numbers must be coming from elsewhere.');
  }
}