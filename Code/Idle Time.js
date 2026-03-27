/**
 * Imports Ozonetel Agent CDR reports from GitHub and parses the 
 * Python-style dictionary strings into a structured sheet.
 */
function importOzonetelAgentReports() {
  const githubSourceUrl = "https://raw.githubusercontent.com/John-Cena-DEV/ozonetel_agents/refs/heads/main/report.csv";
  
  // Fetch the raw CSV data
  const fetchResponse = UrlFetchApp.fetch(githubSourceUrl);
  const rawCsvContent = fetchResponse.getContentText();
  const csvDataGrid = Utilities.parseCsv(rawCsvContent);

  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let targetSheet = activeSpreadsheet.getSheetByName("Agent Performance");
  
  // Create sheet if it doesn't exist, otherwise clear it
  if (!targetSheet) {
    targetSheet = activeSpreadsheet.insertSheet("Agent Performance");
  } else {
    targetSheet.clearContents();
  }

  const finalOutputTable = [];
  const uniqueColumnKeys = new Set();

  // Phase 1: Clean and Parse Data
  for (let rowIndex = 1; rowIndex < csvDataGrid.length; rowIndex++) {
    let rawJsonString = csvDataGrid[rowIndex][0]; // Extract the 'details' column

    if (!rawJsonString) continue;

    // Sanitize Python-style formatting to valid JSON
    const sanitizedJson = rawJsonString
      .replace(/None/g, "null")
      .replace(/True/g, "true")
      .replace(/False/g, "false")
      .replace(/'/g, '"')
      .replace(/,\s*}/g, "}");

    try {
      const parsedDataObj = JSON.parse(sanitizedJson);
      csvDataGrid[rowIndex].parsedData = parsedDataObj;
      
      // Collect all unique keys to build headers
      Object.keys(parsedDataObj).forEach(key => uniqueColumnKeys.add(key));
    } catch (error) {
      Logger.log("Skipping invalid row " + rowIndex + ": " + error.message);
    }
  }

  // Phase 2: Map Data to Columns
  const dynamicHeaders = Array.from(uniqueColumnKeys);
  finalOutputTable.push(dynamicHeaders);

  for (let rowIndex = 1; rowIndex < csvDataGrid.length; rowIndex++) {
    const rowObj = csvDataGrid[rowIndex].parsedData;
    if (!rowObj) continue;

    // Ensure every row has a value (or empty string) for every header
    const rowValues = dynamicHeaders.map(header => rowObj[header] ?? "");
    finalOutputTable.push(rowValues);
  }

  // Phase 3: Write to Sheet
  if (finalOutputTable.length > 0) {
    targetSheet.getRange(1, 1, finalOutputTable.length, dynamicHeaders.length).setValues(finalOutputTable);
  }
}





/**
 * Process "agent performance" sheet:
 * - Sums TotalIdleTime, TotalLoginDuration, totalWrapupTime per agent
 *   across ALL rows (Blended + Manual combined)
 * - Maps agent names to canonical names via teamMapping
 * - Returns { data: { canonicalName: {idleTime, loginTime, wrapupTime} }, date: 'dd-MMM-yy' }
 */
function processAgentPerformanceData(teamMapping) {
  Logger.log('=== Processing Agent Performance Data ===');

  // Guard: if teamMapping is null/invalid, return empty
  if (!teamMapping || !teamMapping.canonicalMap) {
    Logger.log('⚠️  teamMapping not available — skipping perf data.');
    return { data: {}, date: '' };
  }




  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const perfSheet =
    ss.getSheetByName('agent performance') ||
    ss.getSheetByName('Agent Performance') ||
    ss.getSheetByName('Agent performance');

  if (!perfSheet) {
    Logger.log('⚠️  Agent performance sheet not found — skipping perf columns.');
    return { data: {}, date: '' };
  }

  const lastRow = perfSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No data rows in agent performance sheet.');
    return { data: {}, date: '' };
  }

  // Read headers (lowercased for flexible matching)
  const rawHeaders = perfSheet
    .getRange(1, 1, 1, perfSheet.getLastColumn())
    .getValues()[0];
  const headers = rawHeaders.map(h => (h ? h.toString().toLowerCase().trim() : ''));

  const data = perfSheet
    .getRange(2, 1, lastRow - 1, perfSheet.getLastColumn())
    .getValues();

  // Column index lookup (handles minor spelling variations)
  const agentNameIdx = headers.findIndex(h => h === 'agentname');
  const callDateIdx  = headers.findIndex(h => h === 'calldate');
  const idleIdx      = headers.findIndex(h => h === 'totalidletime');
  // AFTER this line:
const wrapupIdx    = headers.findIndex(h => h === 'totalwrapuptime' || h === 'wrapuptime');


  const loginIdx     = headers.findIndex(
    h => h === 'totalloginduration' || h.startsWith('totalloginduratio')
  );


  const pauseIdx     = headers.findIndex(
    h => h === 'totalpausetime' || h === 'pausetime' || h === 'pause_time'
  );

  Logger.log(
    `Col indices → agentName:${agentNameIdx} callDate:${callDateIdx} ` +
    `idle:${idleIdx} login:${loginIdx} wrapup:${wrapupIdx} pause:${pauseIdx}`
  );

  if (agentNameIdx === -1) {
    Logger.log('❌ AgentName column not found in agent performance sheet.');
    return { data: {}, date: '' };
  }

  // Build a date label from the first data row (e.g. "02-Mar-26")
  let dateLabel = '';
  if (callDateIdx !== -1 && data.length > 0) {
    const dv = data[0][callDateIdx];
    if (dv instanceof Date) {
      dateLabel = Utilities.formatDate(dv, Session.getScriptTimeZone(), 'dd-MMM-yy');
    } else if (typeof dv === 'string') {
      // e.g. "2026-03-02 00:00:00" → take the date part only
      dateLabel = dv.split(' ')[0];
    } else if (typeof dv === 'number') {
      // Excel serial date
      const d = new Date((dv - 25569) * 86400 * 1000);
      dateLabel = Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd-MMM-yy');
    }
  }

  // Aggregate per canonical agent
  const agentPerf = {};

  data.forEach(row => {
    const rawName = row[agentNameIdx];
    if (!rawName) return;

    const canonical = getCanonicalName(rawName.toString().trim(), teamMapping);

    if (!agentPerf[canonical]) {
      agentPerf[canonical] = { idleTime: 0, loginTime: 0, wrapupTime: 0, pauseTime: 0 };

    }

    if (idleIdx   !== -1) agentPerf[canonical].idleTime   += parsePerfTime(row[idleIdx]);
    if (loginIdx  !== -1) agentPerf[canonical].loginTime  += parsePerfTime(row[loginIdx]);
    if (wrapupIdx !== -1) agentPerf[canonical].wrapupTime += parsePerfTime(row[wrapupIdx]);
    if (pauseIdx  !== -1) agentPerf[canonical].pauseTime  += parsePerfTime(row[pauseIdx]);

  });

  Logger.log(
    `Agent performance loaded for ${Object.keys(agentPerf).length} canonical agents.` +
    ` Date label: "${dateLabel}"`
  );
  return { data: agentPerf, date: dateLabel };
}

/**
 * Parse a time value to total seconds.
 * Handles: Date objects, "HH:MM:SS" strings, and numeric (fraction-of-day or raw seconds).
 */
function parsePerfTime(value) {
  if (value === null || value === undefined || value === '') return 0;

  if (value instanceof Date) {
    return value.getHours() * 3600 + value.getMinutes() * 60 + value.getSeconds();
  }
  if (typeof value === 'string') {
    if (value.includes(':')) return hmsToSeconds(value);
    const n = parseFloat(value);
    if (!isNaN(n)) return n < 1 ? Math.round(n * 86400) : Math.round(n);
  }
  if (typeof value === 'number') {
    return value < 1 ? Math.round(value * 86400) : Math.round(value);
  }
  return 0;
}











