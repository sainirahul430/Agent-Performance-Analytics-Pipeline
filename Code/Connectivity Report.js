function createCompleteCallSummary_Tata_Ozonetel_WhatsApp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const tataSheet = ss.getSheetByName('Call Records Data');
  const ozonetelSheet = ss.getSheetByName('CDR_Data');
  const whatsappSheet = ss.getSheetByName('Whatsapp Call');

  if (!ozonetelSheet) {
    Browser.msgBox('Sheet "CDR_Data" not found for Ozonetel');
    return;
  }

  const summarySheetName = 'Connectivity Report';
  let summarySheet = ss.getSheetByName(summarySheetName);
  if (summarySheet) summarySheet.clear();
  else summarySheet = ss.insertSheet(summarySheetName);

  summarySheet.getRange('A:Z')
    .setFontFamily('Helvetica Neue')
    .setFontSize(13);

  let rowPointer = 1;
  let grandTotalCalls = 0;
  let grandTotalDuration = 0;

  // ---------- TATA ----------
  if (tataSheet) {
    const tataResult = writeVendorBlock({
      sheet: summarySheet,
      rowPointer,
      vendorName: 'TATA',
      sourceSheet: tataSheet,
      countUnique: true,
      mapping: {
        id: 'id',
        direction: 'direction',
        status: 'status',
        dialer: 'call_hint',
        duration: 'total_call_duration',
        agent: 'agent_name_cleaned'
      }
    });
    rowPointer = tataResult.rowPointer;
    grandTotalCalls += tataResult.totalCalls;
    grandTotalDuration += tataResult.totalDuration;
    rowPointer += 3;
  }

  // ---------- OZONETEL ----------
  if (ozonetelSheet) {
    const ozonetelResult = writeVendorBlock({
      sheet: summarySheet,
      rowPointer,
      vendorName: 'OZONETEL',
      sourceSheet: ozonetelSheet,
      countUnique: true,
      mapping: {
        id: 'CallID',
        direction: 'Type',
        status: 'Status',
        dialer: 'CampaignName',
        duration: 'TalkTime',
        agent: 'AgentName'
      }
    });
    rowPointer = ozonetelResult.rowPointer;
    grandTotalCalls += ozonetelResult.totalCalls;
    grandTotalDuration += ozonetelResult.totalDuration;
    rowPointer += 3;
  }

  // ---------- WHATSAPP ----------
  if (whatsappSheet) {
    const whatsappResult = writeWhatsAppBlock({
      sheet: summarySheet,
      rowPointer,
      sourceSheet: whatsappSheet
    });
    rowPointer = whatsappResult.rowPointer;
    grandTotalCalls += whatsappResult.totalCalls;
    grandTotalDuration += whatsappResult.totalDuration;
    rowPointer += 3;
  }

  // ---------- GRAND TOTAL ----------
  summarySheet.getRange(rowPointer, 1, 1, 3).merge();
  summarySheet.getRange(rowPointer, 1)
    .setValue('GRAND TOTAL (ALL VENDORS)')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('center')
    .setFontColor('#ffffff')
    .setBackground('#8E8E93');
  summarySheet.setRowHeight(rowPointer, 42);
  rowPointer++;

  summarySheet.getRange(rowPointer, 1, 1, 3).setValues([[
    grandTotalCalls,
    secondsToSheetDuration(grandTotalDuration),
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM dd, yyyy')
  ]]);
  summarySheet.getRange(rowPointer, 1).setNumberFormat('0');
  summarySheet.getRange(rowPointer, 2).setNumberFormat('[hh]:mm:ss');
  summarySheet.getRange(rowPointer, 3).setHorizontalAlignment('center');

  Browser.msgBox('Summary created successfully');
}

// ------------------ Vendor Block ------------------
function writeVendorBlock({ sheet, rowPointer, vendorName, sourceSheet, mapping, countUnique = true }) {
  const lastRow = sourceSheet.getLastRow();
  const lastCol = sourceSheet.getLastColumn();

  if (lastRow < 2) return { rowPointer, totalCalls: 0, totalDuration: 0 };

  const headers = sourceSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const normalizedHeaders = headers.map(h => h.toString().trim().toLowerCase());
  const idx = {};

  Object.keys(mapping).forEach(k => {
    const colName = mapping[k].toLowerCase();
    idx[k] = normalizedHeaders.indexOf(colName);
    if (idx[k] === -1) {
      Browser.msgBox(`Column "${mapping[k]}" not found for ${vendorName}\nDetected: ${headers.join(', ')}`);
      throw new Error('Missing column');
    }
  });

  const rawData = sourceSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // ---- DEDUPLICATION BY CALL ID ----
  // For both TATA and OZONETEL: keep only the first row per unique CallID
  let data;
  if (countUnique) {
    const seenIds = new Set();
    data = rawData.filter(row => {
      const callId = row[idx.id]?.toString().trim();
      if (!callId || callId === '') return false;   // skip blank IDs
      if (seenIds.has(callId)) return false;        // skip duplicate
      seenIds.add(callId);
      return true;
    });
    Logger.log(`${vendorName}: ${rawData.length} raw rows → ${data.length} after dedup by CallID`);
  } else {
    data = rawData.filter(row => {
      const callId = row[idx.id]?.toString().trim();
      return callId && callId !== '';
    });
  }
  // ----------------------------------

  let totalCallsCount = 0;
  let inbound = 0;
  let outbound = 0;
  let dialer = 0;
  let answered = 0;
  let totalDuration = 0;
  const agentOutboundCount = {};

  // Now iterate over DEDUPLICATED data — all metrics are clean
  data.forEach(row => {
    totalCallsCount++;

    const direction = row[idx.direction]?.toString().toLowerCase() || '';
    const status    = row[idx.status]?.toString().toLowerCase().trim() || '';
    const dialerVal = row[idx.dialer];
    const agent     = row[idx.agent]?.toString().trim() || '';
    const isOutbound = direction.includes('outbound') || direction.includes('manual');

    if (direction.includes('inbound')) inbound++;
    else if (isOutbound) outbound++;

    if (dialerVal && dialerVal.toString().trim() !== '') dialer++;

    if (status === 'answered' && agent !== '') answered++;

    totalDuration += parseDuration(row[idx.duration], vendorName);

    if (isOutbound && agent !== '') {
      agentOutboundCount[agent] = (agentOutboundCount[agent] || 0) + 1;
    }
  });

  const totalCalls  = totalCallsCount;   // already deduplicated above
  const connectivity = totalCalls ? ((answered / totalCalls) * 100).toFixed(2) + '%' : '0%';

  const activeAgents = Object.values(agentOutboundCount).filter(count => count >= 1).length;
  const avgAnswered  = activeAgents > 0 ? Math.round(answered  / activeAgents) : 0;
  const avgDials     = activeAgents > 0 ? Math.round(outbound  / activeAgents) : 0;

  // ... rest of the function (heading, table headers, values) stays exactly the same


 // ---- DEBUG ----
  Logger.log(`\n=== ${vendorName} - Agent Outbound Counts ===`);
  Object.entries(agentOutboundCount)
    .sort((a, b) => b[1] - a[1])
    .forEach(([agent, count]) => {
      Logger.log(`${count >= 1 ? '✅ ACTIVE' : '❌ NOT ACTIVE'} | ${agent} → ${count} outbound calls`);
    });
  Logger.log(`Total Active Agents: ${activeAgents}`);
  // ---- END DEBUG ----


  // ---------- Heading ----------
  sheet.getRange(rowPointer, 1, 1, 10).merge();
  sheet.getRange(rowPointer, 1)
    .setValue(vendorName + ' SUMMARY')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('center')
    .setFontColor('#ffffff')
    .setBackground(vendorName === 'TATA' ? '#0A84FF' : '#FF9F0A');
  sheet.setRowHeight(rowPointer, 42);
  rowPointer++;

  // ---------- Table headers (9 columns now) ----------
const tableHeaders = [
  'Total Calls', 'Outbound', 'Inbound', 'Dialer',
  'Answered', 'Total Talktime', 'Connectivity %',
  'Active Agents', 'Avg Answered / Agent', 'Avg Dials / Agent'
];
sheet.getRange(rowPointer, 1, 1, 10).setValues([tableHeaders]);
sheet.getRange(rowPointer, 1, 1, 10)
  .setFontWeight('bold')
  .setBackground('#F2F2F7')
  .setHorizontalAlignment('center');
  sheet.getRange(rowPointer, 10).setWrap(true);  
  rowPointer++;

  // ---------- Values ----------
sheet.getRange(rowPointer, 1, 1, 10).setValues([[
  totalCalls, outbound, inbound, dialer,
  answered,
  secondsToSheetDuration(totalDuration),
  connectivity,
  activeAgents,
  avgAnswered,
  avgDials        // ← new
]]);
sheet.getRange(rowPointer, 1, 1, 5).setNumberFormat('0');
sheet.getRange(rowPointer, 6).setNumberFormat('[hh]:mm:ss');
sheet.getRange(rowPointer, 7).setHorizontalAlignment('center');
sheet.getRange(rowPointer, 8).setNumberFormat('0');
sheet.getRange(rowPointer, 9).setNumberFormat('0');
sheet.getRange(rowPointer, 10).setNumberFormat('0');  // ← new
  rowPointer++;

  return { rowPointer, totalCalls, totalDuration };
}

// ------------------ WhatsApp Block ------------------
function writeWhatsAppBlock({ sheet, rowPointer, sourceSheet }) {
  const lastRow = sourceSheet.getLastRow();
  const lastCol = sourceSheet.getLastColumn();

  if (lastRow < 2) return { rowPointer, totalCalls: 0, totalDuration: 0 };

  const headers = sourceSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const normalizedHeaders = headers.map(h => h.toString().trim().toLowerCase());

  const timestampIdx = normalizedHeaders.findIndex(h => h.includes('timestamp'));
  const agentIdx = normalizedHeaders.findIndex(h => h.includes('agent_name'));
  const durationIdx = normalizedHeaders.findIndex(h => h.includes('total duration'));

  if (timestampIdx === -1 || agentIdx === -1 || durationIdx === -1) {
    Browser.msgBox(`Required WhatsApp columns not found.\nDetected headers: ${headers.join(', ')}`);
    return { rowPointer, totalCalls: 0, totalDuration: 0 };
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let totalDials = 0;
  let totalDuration = 0;

  const data = sourceSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  data.forEach(row => {
    const timestamp = row[timestampIdx];
    const agent = row[agentIdx];
    const duration = row[durationIdx];

    if (!timestamp || !agent || agent.toString().trim() === '') return;

    let rowDate = (timestamp instanceof Date) ? new Date(timestamp) : new Date(timestamp.toString());
    rowDate.setHours(0, 0, 0, 0);

    if (rowDate.getTime() === today.getTime()) {
      totalDials++;
      totalDuration += parseDuration(duration, 'WHATSAPP');
    }
  });

  // ---------- Heading ----------
  sheet.getRange(rowPointer, 1, 1, 3).merge();
  sheet.getRange(rowPointer, 1)
    .setValue('WHATSAPP SUMMARY (TODAY)')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('center')
    .setFontColor('#ffffff')
    .setBackground('#25D366');
  sheet.setRowHeight(rowPointer, 42);
  rowPointer++;

  const tableHeaders = ['Total Dials', 'Total Talktime', 'Date'];
  sheet.getRange(rowPointer, 1, 1, 3).setValues([tableHeaders]);
  sheet.getRange(rowPointer, 1, 1, 3)
    .setFontWeight('bold')
    .setBackground('#F2F2F7')
    .setHorizontalAlignment('center');
  rowPointer++;

  const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MMM dd, yyyy');
  sheet.getRange(rowPointer, 1, 1, 3).setValues([[totalDials, secondsToSheetDuration(totalDuration), todayStr]]);
  sheet.getRange(rowPointer, 1).setNumberFormat('0');
  sheet.getRange(rowPointer, 2).setNumberFormat('[hh]:mm:ss');
  sheet.getRange(rowPointer, 3).setHorizontalAlignment('center');
  rowPointer++;

  return { rowPointer, totalCalls: totalDials, totalDuration };
}

// ------------------ Utilities ------------------
function parseDuration(val, vendorName = '') {
  if (!val) return 0;

  const strVal = val.toString().trim();
  if (strVal === '' || strVal === '0') return 0;

  if (typeof val === 'number' && val > 0 && val < 1) return Math.round(val * 86400);
  if (typeof val === 'number' && val >= 1) return Math.round(val);

  if (strVal.includes(':')) {
    const parts = strVal.split(':').map(p => parseInt(p) || 0);
    if (parts.length === 3) return parts[0]*3600 + parts[1]*60 + parts[2];
    if (parts.length === 2) return parts[0]*60 + parts[1];
  }

  const numVal = parseFloat(strVal);
  if (!isNaN(numVal) && numVal > 0) return Math.round(numVal);

  return 0;
}

function secondsToSheetDuration(seconds) {
  if (!seconds || seconds === 0) return 0;
  return seconds / 86400;
}