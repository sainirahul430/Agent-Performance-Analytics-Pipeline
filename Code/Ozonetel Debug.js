/**
 * Look up calls and talktime for a specific agent, broken down by time bucket
 * HOW TO USE: Change AGENT_NAME below, then Run this function
 */
function debugAgentTimeBuckets() {
  const AGENT_NAME = 'Bhavesh'; // ← Change this to any agent name (partial match works)

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ozonetelSheet = ss.getSheetByName('CDR_Data');

  if (!ozonetelSheet) {
    Logger.log('ERROR: CDR_Data sheet not found');
    return;
  }

  const lastRow = ozonetelSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No data in CDR_Data sheet');
    return;
  }

  const headers = ozonetelSheet
    .getRange(1, 1, 1, ozonetelSheet.getLastColumn())
    .getValues()[0];

  const data = ozonetelSheet
    .getRange(2, 1, lastRow - 1, ozonetelSheet.getLastColumn())
    .getValues();

  // ── Column detection ──
  const agentNameIdx = headers.findIndex(h =>
    h && h.toString().toLowerCase().replace(/\s/g, '') === 'agentname'
  );
  const callDateIdx = headers.findIndex(h =>
    h && h.toString().toLowerCase().includes('calldate')
  );
  const startTimeIdx = headers.findIndex(h =>
    h && h.toString().toLowerCase().includes('starttime')
  );
  const durationIdx = headers.findIndex(h =>
    h && h.toString().toLowerCase() === 'duration'
  );
  const callIdIdx = headers.findIndex(h =>
    h && h.toString().toLowerCase().replace(/\s|_/g, '') === 'callid'
  );
  const typeIdx = headers.findIndex(h =>
    h && h.toString().toLowerCase() === 'type'
  );
  const statusIdx = headers.findIndex(h =>
    h && h.toString().toLowerCase() === 'status'
  );

  Logger.log('=== COLUMN CHECK ===');
  Logger.log('agentName col : ' + (agentNameIdx !== -1 ? headers[agentNameIdx] : '❌ NOT FOUND'));
  Logger.log('callDate col  : ' + (callDateIdx   !== -1 ? headers[callDateIdx]  : '❌ NOT FOUND'));
  Logger.log('startTime col : ' + (startTimeIdx  !== -1 ? headers[startTimeIdx] : '❌ NOT FOUND'));
  Logger.log('duration col  : ' + (durationIdx   !== -1 ? headers[durationIdx]  : '❌ NOT FOUND'));
  Logger.log('callId col    : ' + (callIdIdx     !== -1 ? headers[callIdIdx]    : '❌ NOT FOUND'));

  if (agentNameIdx === -1 || callDateIdx === -1 || startTimeIdx === -1 || durationIdx === -1) {
    Logger.log('❌ Cannot continue — required columns missing');
    return;
  }

  // ── Date range ──
  const dateConfig = getDateRangeConfig();
  const fromDate = new Date(dateConfig.fromDate);
  const toDate   = new Date(dateConfig.toDate);
  fromDate.setHours(0, 0, 0, 0);
  toDate.setHours(23, 59, 59, 999);

  Logger.log('\n=== DATE RANGE ===');
  Logger.log('From : ' + fromDate.toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }));
  Logger.log('To   : ' + toDate.toLocaleString('en-IN',   { timeZone: 'Asia/Kolkata' }));

  // ── Time bucket map ──
  const hourToBucket = {
    8: '08-09 AM', 9: '09-10 AM', 10: '10-11 AM', 11: '11-12 PM',
    12: '12-01 PM', 13: '01-02 PM', 14: '02-03 PM', 15: '03-04 PM',
    16: '04-05 PM', 17: '05-06 PM', 18: '06-07 PM',
    19: '07-08 PM', 20: '08-09 PM', 21: '09-10 PM'
  };

  const bucketOrder = [
    '08-09 AM','09-10 AM','10-11 AM','11-12 PM',
    '12-01 PM','01-02 PM','02-03 PM','03-04 PM',
    '04-05 PM','05-06 PM','06-07 PM','07-08 PM',
    '08-09 PM','09-10 PM'
  ];

  // ── Accumulators ──
  const bucketData   = {}; // { bucket: { calls, duration } }
  const seenCallIds  = new Set();
  const matchedNames = new Set();

  let totalRows     = 0;
  let matchedRows   = 0;
  let skippedDupe   = 0;
  let skippedDate   = 0;
  let skippedBucket = 0;

  bucketOrder.forEach(b => { bucketData[b] = { calls: 0, duration: 0 }; });

  // ── Main loop ──
  data.forEach(row => {
    const rawAgent = row[agentNameIdx];
    if (!rawAgent) return;

    totalRows++;

    // Partial, case-insensitive match
    if (!rawAgent.toString().toLowerCase().includes(AGENT_NAME.toLowerCase())) return;

    matchedRows++;
    matchedNames.add(rawAgent.toString().trim());

    // Deduplicate by CallID (global, not per-agent — change to per-agent if needed)
    if (callIdIdx !== -1) {
      const callId = String(row[callIdIdx]).trim();
      if (seenCallIds.has(callId)) { skippedDupe++; return; }
      seenCallIds.add(callId);
    }

    // Parse timestamp
    const callDateValue  = row[callDateIdx];
    const startTimeValue = row[startTimeIdx];
    let timestamp;

    if (callDateValue instanceof Date) {
      timestamp = new Date(callDateValue.getTime());
    } else if (typeof callDateValue === 'string') {
      timestamp = new Date(callDateValue);
    } else if (typeof callDateValue === 'number') {
      timestamp = new Date((callDateValue - 25569) * 86400 * 1000);
    }

    if (!timestamp || isNaN(timestamp.getTime())) {
      Logger.log('⚠️  Bad timestamp on row: ' + JSON.stringify(row.slice(0, 5)));
      return;
    }

    // Overlay startTime for correct hour
    if (typeof startTimeValue === 'string' && startTimeValue.includes(':')) {
      const p = startTimeValue.split(':');
      timestamp.setHours(+p[0] || 0, +p[1] || 0, +p[2] || 0);
    } else if (startTimeValue instanceof Date) {
      timestamp.setHours(
        startTimeValue.getHours(),
        startTimeValue.getMinutes(),
        startTimeValue.getSeconds()
      );
    }

    // Date range check
    const dateOnly = new Date(timestamp);
    dateOnly.setHours(0, 0, 0, 0);
    if (dateOnly < fromDate || dateOnly > toDate) { skippedDate++; return; }

    // Bucket
    const hour   = timestamp.getHours();
    const bucket = hourToBucket[hour];
    if (!bucket) { skippedBucket++; return; }

    // Duration
    let duration = 0;
    const d = row[durationIdx];
    if (typeof d === 'number') {
      duration = d < 1 ? Math.round(d * 86400) : d;
    } else if (typeof d === 'string') {
      duration = d.includes(':') ? hmsToSeconds(d) : Number(d) || 0;
    } else if (d instanceof Date) {
      duration = d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds();
    }

    bucketData[bucket].calls++;
    bucketData[bucket].duration += duration;
  });

  // ── Print results ──
  Logger.log('\n=== AGENT MATCH INFO ===');
  Logger.log('Searched for : "' + AGENT_NAME + '"');
  Logger.log('Exact names matched in sheet: ' + Array.from(matchedNames).join(' | '));
  Logger.log('Total rows scanned : ' + totalRows);
  Logger.log('Rows matching agent: ' + matchedRows);
  Logger.log('Skipped (duplicate CallID) : ' + skippedDupe);
  Logger.log('Skipped (out of date range): ' + skippedDate);
  Logger.log('Skipped (outside buckets)  : ' + skippedBucket);

  Logger.log('\n=== TIME BUCKET BREAKDOWN ===');
  Logger.log('Bucket       | Calls |   Talktime');
  Logger.log('-------------|-------|------------');

  let grandCalls    = 0;
  let grandDuration = 0;

  bucketOrder.forEach(bucket => {
    const { calls, duration } = bucketData[bucket];
    if (calls > 0) {
      Logger.log(
        bucket.padEnd(12) + ' | ' +
        String(calls).padStart(5) + ' | ' +
        secondsToHMS(duration)
      );
      grandCalls    += calls;
      grandDuration += duration;
    }
  });

  Logger.log('-------------|-------|------------');
  Logger.log('TOTAL        | ' + String(grandCalls).padStart(5) + ' | ' + secondsToHMS(grandDuration));
  Logger.log('=====================================');
}
