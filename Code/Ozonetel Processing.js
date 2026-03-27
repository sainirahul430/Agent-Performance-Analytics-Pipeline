function processOzonetelDataForTeamSheets(teamData, timeBuckets, teamMapping) {
  Logger.log('=== Processing Ozonetel Data for Team Sheets ===');

  // 🔍 Kiran tracking variables
  let kiranTotal = 0;
  let kiranSkippedDupe = 0;
  let kiranSkippedNoAgent = 0;
  let kiranSkippedTimestamp = 0;
  let kiranSkippedRange = 0;
  let kiranProcessed = 0;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ozonetelSheet = ss.getSheetByName('CDR_Data');

  if (!ozonetelSheet) {
    Logger.log('CDR_Data sheet not found');
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

  // ===============================
  // COLUMN INDICES
  // ===============================
  const agentNameIdx = headers.findIndex(h =>
    h && h.toString().toLowerCase().replace(/\s/g, '') === 'agentname'
  );

  const typeIdx = headers.findIndex(h =>
    h && h.toString().toLowerCase() === 'type'
  );

  const durationIdx = headers.findIndex(h =>
    h && h.toString().toLowerCase() === 'duration'
  );

  const statusIdx = headers.findIndex(h =>
    h && h.toString().toLowerCase() === 'status'
  );

  const startTimeIdx = headers.findIndex(h =>
    h && h.toString().toLowerCase().includes('starttime')
  );

  const callDateIdx = headers.findIndex(h =>
    h && h.toString().toLowerCase().includes('calldate')
  );

  const callIdIdx = headers.findIndex(h =>
    h && h.toString().toLowerCase().replace(/\s|_/g, '') === 'callid'
  );

  const endTimeIdx = headers.findIndex(h =>
    h && h.toString().toLowerCase().replace(/\s|_/g, '').includes('endtime')
  );




  if (
    agentNameIdx === -1 ||
    typeIdx === -1 ||
    durationIdx === -1 ||
    statusIdx === -1 ||
    callDateIdx === -1 ||
    startTimeIdx === -1 ||
    callIdIdx === -1
  ) {
    Logger.log('ERROR: Required columns missing');
    return;
  }

  // ===============================
  // DATE RANGE
  // ===============================
  const dateConfig = getDateRangeConfig();
  const fromDate = new Date(dateConfig.fromDate);
  const toDate = new Date(dateConfig.toDate);
  fromDate.setHours(0, 0, 0, 0);
  toDate.setHours(23, 59, 59, 999);

  // ===============================
  // TIME BUCKET MAP
  // ===============================
  const hourToBucket = {
    8: '08-09 AM', 9: '09-10 AM', 10: '10-11 AM', 11: '11-12 PM',
    12: '12-01 PM', 13: '01-02 PM', 14: '02-03 PM', 15: '03-04 PM',
    16: '04-05 PM', 17: '05-06 PM', 18: '06-07 PM',
    19: '07-08 PM', 20: '08-09 PM', 21: '09-10 PM'
  };

function splitDurationAcrossBuckets(startTs, endTs, hourToBucket) {
  const splits = {};
  if (!startTs || !endTs || endTs <= startTs) {
    const bucket = hourToBucket[startTs.getHours()];
    if (bucket) splits[bucket] = (endTs - startTs) / 1000 || 0;
    return splits;
  }
  let cursor = new Date(startTs.getTime());
  while (cursor < endTs) {
    const hour = cursor.getHours();
    const bucket = hourToBucket[hour];
    const slotEnd = new Date(cursor);
    slotEnd.setMinutes(59, 59, 999);
    const segmentEnd = slotEnd < endTs ? slotEnd : endTs;
    const seconds = Math.round((segmentEnd - cursor) / 1000);
    if (bucket && seconds > 0) {
      splits[bucket] = (splits[bucket] || 0) + seconds;
    }
    cursor = new Date(slotEnd.getTime() + 1);
  }
  return splits;
}


  // ===============================
  // UNIQUE CALL TRACKER
  // ===============================

const seenCallIdsByAgent = new Map(); // Track CallIDs per agent

  let processedCount = 0;
  let totalDuration = 0;
  let skipped = {
    duplicateCallId: 0,
    noTimestamp: 0,
    outOfRange: 0,
    invalidBucket: 0
  };

  // ===============================
  // MAIN LOOP
  // ===============================
  data.forEach((row, rowIndex) => {
    const rawCallId = row[callIdIdx];
    if (!rawCallId) return;
    
    const callId = String(rawCallId).trim();
    
const agentName = row[agentNameIdx];
    if (!agentName) return;

    // Get canonical agent for duplicate check
    const canonicalAgent = getCanonicalName(agentName, teamMapping);
    if (!canonicalAgent) return;
    
    // Check if Kiran
    const isKiranCall = canonicalAgent.toLowerCase().includes('kiran');
    if (isKiranCall) kiranTotal++;

    // 🚫 SKIP DUPLICATE CALLS - PER AGENT
    if (!seenCallIdsByAgent.has(canonicalAgent)) {
      seenCallIdsByAgent.set(canonicalAgent, new Set());
    }

    const agentCallIds = seenCallIdsByAgent.get(canonicalAgent);
    if (agentCallIds.has(callId)) {
      skipped.duplicateCallId++;
      if (isKiranCall) kiranSkippedDupe++;
      return;
    }
    agentCallIds.add(callId);

    // ---- TIMESTAMP PARSING ----
    const callDateValue = row[callDateIdx];
    const startTimeValue = row[startTimeIdx];
    let timestamp;

    if (callDateValue instanceof Date) {
      timestamp = new Date(callDateValue.getTime());
    } else if (typeof callDateValue === 'string') {
      timestamp = new Date(callDateValue);
    } else if (typeof callDateValue === 'number') {
      timestamp = new Date((callDateValue - 25569) * 86400 * 1000);
    }

    if (!timestamp || isNaN(timestamp)) {
      skipped.noTimestamp++;
      if (isKiranCall) kiranSkippedTimestamp++;
      return;
    }

    if (typeof startTimeValue === 'string') {
      const p = startTimeValue.split(':');
      timestamp.setHours(+p[0] || 0, +p[1] || 0, +p[2] || 0);
    } else if (startTimeValue instanceof Date) {
      timestamp.setHours(
        startTimeValue.getHours(),
        startTimeValue.getMinutes(),
        startTimeValue.getSeconds()
      );
    }

    const dateOnly = new Date(timestamp);
    dateOnly.setHours(0, 0, 0, 0);

    if (dateOnly < fromDate || dateOnly > toDate) {
      skipped.outOfRange++;
      if (isKiranCall) kiranSkippedRange++;
      return;
    }


  

    // ---- DURATION ----
    let duration = 0;
    const d = row[durationIdx];

    if (typeof d === 'number') {
      duration = d < 1 ? Math.round(d * 86400) : d;
    } else if (typeof d === 'string') {
      duration = d.includes(':') ? hmsToSeconds(d) : Number(d) || 0;
    } else if (d instanceof Date) {
      duration = d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds();
    }

    // ---- CALL TYPE ----
    const type = (row[typeIdx] || '').toString().toLowerCase();
    const status = (row[statusIdx] || '').toString().toLowerCase();

    const isInbound = type === 'inbound';
    const isOutbound = type === 'manual';
    const isAnswered = status === 'answered';
    const isDialer = type === 'progressive';

    // ---- AGENT / TEAM ----
    
    const team = getAgentTeam(agentName, teamMapping);
    
    // Track if processed for Kiran
    if (canonicalAgent && canonicalAgent.toLowerCase().includes('kiran')) {
      kiranProcessed++;
    }

    if (!teamData[team]) teamData[team] = {};

    if (!teamData[team][canonicalAgent]) {
      teamData[team][canonicalAgent] = {
        totalCalls: 0,
        totalDuration: 0,
        dialerCalls: 0,
        inboundCalls: 0,
        outboundCalls: 0,
        answeredCalls: 0,
        ozonetelCalls: 0,
        ozonetelDuration: 0,
        buckets: {}
      };

      timeBuckets.forEach(b => {
        teamData[team][canonicalAgent].buckets[b] = {
          calls: 0,
          duration: 0,
          inboundCalls: 0
        };
      });
    }

    // ---- AGGREGATION ----
    const agentObj = teamData[team][canonicalAgent];

    agentObj.ozonetelCalls++;
    agentObj.ozonetelDuration += duration;

    if (isInbound) agentObj.inboundCalls++;
    if (isOutbound) agentObj.outboundCalls++;
    if (isAnswered) agentObj.answeredCalls++;
    if (isDialer) agentObj.dialerCalls++;

// Parse end timestamp
let endTimestamp = null;
if (endTimeIdx !== -1) {
  const endTimeValue = row[endTimeIdx];
  endTimestamp = new Date(timestamp.getTime());
  if (typeof endTimeValue === 'string' && endTimeValue.includes(':')) {
    const p = endTimeValue.split(':');
    endTimestamp.setHours(+p[0] || 0, +p[1] || 0, +p[2] || 0);
  } else if (endTimeValue instanceof Date) {
    endTimestamp.setHours(
      endTimeValue.getHours(),
      endTimeValue.getMinutes(),
      endTimeValue.getSeconds()
    );
  }
  if (endTimestamp < timestamp) {
    endTimestamp.setDate(endTimestamp.getDate() + 1); // midnight crossover
  }
}
// Fallback if no end time column
if (!endTimestamp) {
  endTimestamp = new Date(timestamp.getTime() + duration * 1000);
}

// Split duration across buckets
const bucketSplits = splitDurationAcrossBuckets(timestamp, endTimestamp, hourToBucket);
Object.entries(bucketSplits).forEach(([b, splitDuration], i) => {
  if (agentObj.buckets[b]) {
    if (i === 0) agentObj.buckets[b].calls++;                        // count call only in start bucket
    agentObj.buckets[b].duration += splitDuration;
    if (isInbound && i === 0) agentObj.buckets[b].inboundCalls++;   // same for inbound
  }
});

    totalDuration += duration;
    processedCount++;
  });

  // ===============================
  // SUMMARY
  // ===============================
  Logger.log('=== Ozonetel Processing Summary ===');
  Logger.log('🔍 KIRAN BREAKDOWN:');
  Logger.log('  Total Kiran rows seen: ' + kiranTotal);
  Logger.log('  Skipped (duplicate): ' + kiranSkippedDupe);
  Logger.log('  Skipped (no agent): ' + kiranSkippedNoAgent);
  Logger.log('  Skipped (bad timestamp): ' + kiranSkippedTimestamp);
  Logger.log('  Skipped (out of range): ' + kiranSkippedRange);
  Logger.log('  Successfully processed: ' + kiranProcessed);
  Logger.log('---');
const totalUniqueCallIds = Array.from(seenCallIdsByAgent.values())
  .reduce((sum, set) => sum + set.size, 0);
Logger.log('Unique Call IDs processed: ' + totalUniqueCallIds);
  Logger.log('Processed calls: ' + processedCount);
  Logger.log('Total duration: ' + secondsToHMS(totalDuration));
  Logger.log('Skipped duplicates: ' + skipped.duplicateCallId);
  Logger.log('Skipped no timestamp: ' + skipped.noTimestamp);
  Logger.log('Skipped out of range: ' + skipped.outOfRange);
  Logger.log('Skipped invalid bucket: ' + skipped.invalidBucket);
  Logger.log('====================================');
}