/**
 * ============================================================
 *  CANONICAL AGENT REPORT  (v3 – Team Grouped)
 *
 *  Output columns: Agent Name | Login Time | Idle Time | Wrapup Time
 *  Layout:
 *    ┌─ TEAM NAME header row (coloured)
 *    │   Agent 1
 *    │   Agent 2
 *    │   Team Subtotal row
 *    ├─ NEXT TEAM ...
 *    └─ GRAND TOTAL row
 * ============================================================
 */


// ────────────────────────────────────────────────────────────
//  ENTRY POINT
// ────────────────────────────────────────────────────────────
function generateCanonicalAgentReport() {
  const wb = SpreadsheetApp.getActiveSpreadsheet();

  // Step 1 – Load Team Mapping (raw name → canonical + team)
  const teamLookup = loadTeamLookup(wb);
  // teamLookup = {
  //   nameMap : { 'lowercase raw name': 'Canonical Name' },
  //   teamMap : { 'Canonical Name': 'Team Name' }
  // }

  // Step 2 – Aggregate Agent Performance sheet
  const aggResult   = aggregateByCanonical(wb, teamLookup.nameMap);
  const perfBuckets = aggResult.buckets;     // { canonicalName: { secLogin, secIdle, secWrapup } }
  const reportDate  = aggResult.callDate;    // "07-Mar-26"

  // Step 3 – Attach team to each bucket
  attachTeamsToBuckets(perfBuckets, teamLookup.teamMap);

  // Step 4 – Write output
  const destSheet = fetchOrCreateSheet(wb, 'Canonical Report');
  renderTeamGroupedReport(destSheet, perfBuckets, reportDate);

  Logger.log('✅  Canonical Report refreshed!');
  try { SpreadsheetApp.getUi().alert('✅  Canonical Report refreshed!'); } catch (_) {}
}


// ════════════════════════════════════════════════════════════
//  A – Load Team Mapping
//  Column A: Raw Agent Name | Column B: Team | Column C: Canonical Name
// ════════════════════════════════════════════════════════════
function loadTeamLookup(wb) {
  const mappingSheet =
    wb.getSheetByName('Team Mapping') ||
    wb.getSheetByName('team mapping') ||
    wb.getSheetByName('TeamMapping');

  const nameMap = {};   // { 'lowercase raw': 'Canonical Name' }
  const teamMap = {};   // { 'Canonical Name': 'Team Name' }

  if (!mappingSheet) {
    Logger.log('⚠️  Team Mapping sheet not found.');
    return { nameMap, teamMap };
  }

  const gridRows = mappingSheet.getDataRange().getValues();

  for (let r = 1; r < gridRows.length; r++) {
    const rawAgentName  = (gridRows[r][0] || '').toString().trim();   // Col A
    const teamName      = (gridRows[r][1] || '').toString().trim();   // Col B
    const canonicalName = (gridRows[r][2] || '').toString().trim();   // Col C

    if (rawAgentName && canonicalName) {
      nameMap[rawAgentName.toLowerCase()] = canonicalName;
      if (teamName) teamMap[canonicalName] = teamName;
    }
  }

  Logger.log(`Loaded ${Object.keys(nameMap).length} name entries, ${Object.keys(teamMap).length} team entries.`);
  return { nameMap, teamMap };
}


// ════════════════════════════════════════════════════════════
//  B – Aggregate Agent Performance (Login + Idle + Wrapup only)
// ════════════════════════════════════════════════════════════
function aggregateByCanonical(wb, nameMap) {
  const perfSheet =
    wb.getSheetByName('Agent Performance') ||
    wb.getSheetByName('agent performance') ||
    wb.getSheetByName('Agent performance');

  if (!perfSheet || perfSheet.getLastRow() < 2) {
    Logger.log('⚠️  Agent Performance sheet missing or empty.');
    return { buckets: {}, callDate: '' };
  }

  const lastCol    = perfSheet.getLastColumn();
  const lastRow    = perfSheet.getLastRow();
  const headerVals = perfSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const dataGrid   = perfSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const normHdrs   = headerVals.map(h => h.toString().toLowerCase().replace(/\s/g, ''));

  // Column finder
  function findCol(exact, fallbacks) {
    let idx = normHdrs.indexOf(exact);
    if (idx !== -1) return idx;
    if (fallbacks) for (const fb of fallbacks) {
      idx = normHdrs.findIndex(h => h.startsWith(fb));
      if (idx !== -1) return idx;
    }
    return -1;
  }

  const COL = {
    agentName : findCol('agentname'),
    callDate  : findCol('calldate'),
    idleTime  : findCol('totalidletime'),
    loginDur  : findCol('totalloginduration', ['totallogindu']),
    wrapupTime: findCol('totalwrapuptime'),
    pauseTime : findCol('totalpausetime', ['pausetime', 'pause_time']),
  };

  Logger.log('Columns → ' + JSON.stringify(COL));
  if (COL.agentName === -1) {
    Logger.log('❌  AgentName column not found.');
    return { buckets: {}, callDate: '' };
  }

  // Date label from first row
  let callDate = '';
  if (COL.callDate !== -1 && dataGrid.length > 0) {
    const dv = dataGrid[0][COL.callDate];
    if (dv instanceof Date) {
      callDate = Utilities.formatDate(dv, Session.getScriptTimeZone(), 'dd-MMM-yy');
    } else if (typeof dv === 'string') {
      callDate = dv.split(' ')[0];
    } else if (typeof dv === 'number') {
      callDate = Utilities.formatDate(
        new Date((dv - 25569) * 86400 * 1000),
        Session.getScriptTimeZone(), 'dd-MMM-yy'
      );
    }
  }

  const buckets = {};

  dataGrid.forEach(dataRow => {
    const rawName = (dataRow[COL.agentName] || '').toString().trim();
    if (!rawName) return;

    const canonical = nameMap[rawName.toLowerCase()] || rawName;

    if (!buckets[canonical]) {
buckets[canonical] = { secLogin: 0, secIdle: 0, secWrapup: 0, secPause: 0, teamName: '' };
    }

    const bkt = buckets[canonical];
    if (COL.loginDur   !== -1) bkt.secLogin  += cellToSeconds(dataRow[COL.loginDur]);
    if (COL.idleTime   !== -1) bkt.secIdle   += cellToSeconds(dataRow[COL.idleTime]);
    if (COL.wrapupTime !== -1) bkt.secWrapup += cellToSeconds(dataRow[COL.wrapupTime]);
    if (COL.pauseTime !== -1) bkt.secPause += cellToSeconds(dataRow[COL.pauseTime]);

  });

  Logger.log(`Aggregated ${Object.keys(buckets).length} canonical agents. Date: "${callDate}"`);
  return { buckets, callDate };
}


// ════════════════════════════════════════════════════════════
//  C – Attach team names to each bucket
// ════════════════════════════════════════════════════════════
function attachTeamsToBuckets(perfBuckets, teamMap) {
  Object.keys(perfBuckets).forEach(canonicalName => {
    perfBuckets[canonicalName].teamName = teamMap[canonicalName] || 'Unassigned';
  });
}


// ════════════════════════════════════════════════════════════
//  D – Render team-grouped report
// ════════════════════════════════════════════════════════════
function renderTeamGroupedReport(destSheet, perfBuckets, reportDate) {
  const dateSuffix = reportDate ? ` (${reportDate})` : '';

  // ── Group agents by team ──────────────────────────────────────────────────
  const teamGroups = {};
  Object.keys(perfBuckets).forEach(agentName => {
    const tn = perfBuckets[agentName].teamName;
    if (!teamGroups[tn]) teamGroups[tn] = [];
    teamGroups[tn].push(agentName);
  });

  const sortedTeams = Object.keys(teamGroups).sort();
  sortedTeams.forEach(tn => teamGroups[tn].sort());

  const outputGrid = [];
  const rowMeta    = [];

  // Column header row
  outputGrid.push([
    'Agent Name',
    `Login Time${dateSuffix}`,
    `Idle Time${dateSuffix}`,
    `Wrapup Time${dateSuffix}`,
    `Pause Time${dateSuffix}`
  ]);
  rowMeta.push({ type: 'colheader' });

  // Grand total accumulators
  let grandLogin = 0, grandIdle = 0, grandWrapup = 0, grandPause = 0;

  sortedTeams.forEach(teamName => {
    const agentsInTeam = teamGroups[teamName];

    // Spacers
    outputGrid.push(['', '', '', '', '']);
    rowMeta.push({ type: 'spacer' });

    outputGrid.push(['', '', '', '', '']);
    rowMeta.push({ type: 'spacer' });

    // Team title
    outputGrid.push([teamName, '', '', '', '']);
    rowMeta.push({ type: 'teamheader', team: teamName });

    // Repeated column headers
    outputGrid.push([
      'Agent Name',
      `Login Time${dateSuffix}`,
      `Idle Time${dateSuffix}`,
      `Wrapup Time${dateSuffix}`,
      `Pause Time${dateSuffix}`
    ]);
    rowMeta.push({ type: 'colheader' });

    // Agent rows
    let subtotalLogin = 0, subtotalIdle = 0, subtotalWrapup = 0, subtotalPause = 0;

    agentsInTeam.forEach(agentName => {
      const bkt = perfBuckets[agentName];
      outputGrid.push([
        agentName,
        hmsOrBlank(bkt.secLogin),
        hmsOrBlank(bkt.secIdle),
        hmsOrBlank(bkt.secWrapup),
        hmsOrBlank(bkt.secPause)
      ]);
      rowMeta.push({ type: 'agent', team: teamName, idleSec: bkt.secIdle });

      subtotalLogin  += bkt.secLogin;
      subtotalIdle   += bkt.secIdle;
      subtotalWrapup += bkt.secWrapup;
      subtotalPause  += bkt.secPause;
    });

    grandLogin  += subtotalLogin;
    grandIdle   += subtotalIdle;
    grandWrapup += subtotalWrapup;
    grandPause  += subtotalPause;

    // Team subtotal row
    outputGrid.push([
      `${teamName} — Total`,
      hmsOrBlank(subtotalLogin),
      hmsOrBlank(subtotalIdle),
      hmsOrBlank(subtotalWrapup),
      hmsOrBlank(subtotalPause)
    ]);
    rowMeta.push({ type: 'subtotal', team: teamName });
  });

  // Grand total row
  outputGrid.push([
    'GRAND TOTAL',
    hmsOrBlank(grandLogin),
    hmsOrBlank(grandIdle),
    hmsOrBlank(grandWrapup),
    hmsOrBlank(grandPause)
  ]);
  rowMeta.push({ type: 'grandtotal' });

  // Write to sheet
  const totalRows = outputGrid.length;
  const totalCols = 5;
  destSheet.getRange(1, 1, totalRows, totalCols).setValues(outputGrid);

  // Apply formatting
  applyTeamGroupedStyle(destSheet, outputGrid, rowMeta, totalCols);
}


// ════════════════════════════════════════════════════════════
//  E – Formatting
// ════════════════════════════════════════════════════════════

// Distinct pastel colours cycling per team
const TEAM_PALETTE = [
  { header: '#F3F4F6', text: '#111827' } // clean grey
];

function applyTeamGroupedStyle(destSheet, outputGrid, rowMeta, totalCols) {
  const teamColourIndex = {};  // { teamName: index into TEAM_PALETTE }
  let colourCursor = 0;

  rowMeta.forEach((meta, zeroIdx) => {
    const sheetRow = zeroIdx + 1;  // 1-based
    const range    = destSheet.getRange(sheetRow, 1, 1, totalCols);

if (meta.type === 'colheader') {

  range
    .setBackground('#D9D9D9')
    .setFontColor('#111111')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  destSheet.setRowHeight(sheetRow, 32);
}  else if (meta.type === 'teamheader') {

  const teamRange = destSheet.getRange(sheetRow, 1, 1, totalCols);
  teamRange.merge();   // <-- this centers across the full width

  teamRange
    .setBackground('#EFEFEF')
    .setFontColor('#000000')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  destSheet.setRowHeight(sheetRow, 34);
}

 else if (meta.type === 'agent') {

  range.setBackground('#FFFFFF');

  destSheet.getRange(sheetRow, 1)
    .setHorizontalAlignment('left')
    .setFontSize(10);

  destSheet.getRange(sheetRow, 2, 1, 3)
    .setHorizontalAlignment('center')
    .setFontFamily('Verdana')
    .setFontSize(15);

  destSheet.setRowHeight(sheetRow, 24);

  const IDLE_THRESHOLD_SEC = 5 * 3600;

if ((meta.idleSec || 0) > IDLE_THRESHOLD_SEC) {
  destSheet.getRange(sheetRow, 3)
    .setFontColor('#C00000')
    .setFontWeight('bold');
}
}   else if (meta.type === 'subtotal') {

  range
    .setBackground('#E6E6E6')
    .setFontWeight('bold')
    .setFontSize(10);

  destSheet.getRange(sheetRow, 1).setHorizontalAlignment('center');

  destSheet.getRange(sheetRow, 2, 1, 3)
    .setHorizontalAlignment('center')
    .setFontFamily('Courier New');

  destSheet.setRowHeight(sheetRow, 26);
}
  });

  // Column widths
  destSheet.setColumnWidth(1, 180);   // Agent Name
  destSheet.setColumnWidth(2, 160);   // Login Time
  destSheet.setColumnWidth(3, 160);   // Idle Time
  destSheet.setColumnWidth(4, 160);   // Wrapup Time

  // Freeze header row only (cannot freeze columns alongside merged cells)
  destSheet.setFrozenRows(1);

  // Outer border
  destSheet.getRange(1, 1, outputGrid.length, totalCols)
    .setBorder(true, true, true, true, false, false,
      '#D0D0D0', SpreadsheetApp.BorderStyle.SOLID);


// ── Conditional formatting for Idle Time column

const idleRange = destSheet.getRange(2, 3, destSheet.getLastRow() - 1, 1);

const rules = [];

// good idle (light green)
rules.push(
  SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(7200)
    .setBackground('#E8F5E9')
    .setFontColor('#1B5E20')
    .setRanges([idleRange])
    .build()
);

// medium idle (light amber)
rules.push(
  SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(7200, 18000)
    .setBackground('#FFF4E5')
    .setFontColor('#8A4B00')
    .setRanges([idleRange])
    .build()
);

// high idle (light red)
rules.push(
  SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(18000)
    .setBackground('#FDECEA')
    .setFontColor('#B00020')
    .setRanges([idleRange])
    .build()
);

destSheet.setConditionalFormatRules(rules);

}

/**
 * Blend a hex colour toward white by `factor` (0 = original, 1 = white).
 * Used to generate light agent-row tints from team header colours.
 */
function blendWithWhite(hexColour, factor) {
  try {
    const r = parseInt(hexColour.slice(1, 3), 16);
    const g = parseInt(hexColour.slice(3, 5), 16);
    const b = parseInt(hexColour.slice(5, 7), 16);
    const blend = c => Math.round(c + (255 - c) * factor);
    return '#' + [blend(r), blend(g), blend(b)]
      .map(v => v.toString(16).padStart(2, '0')).join('');
  } catch (_) {
    return '#F5F5F5';
  }
}


// ════════════════════════════════════════════════════════════
//  F – Helpers
// ════════════════════════════════════════════════════════════

function fetchOrCreateSheet(wb, sheetTitle) {
  let sh = wb.getSheetByName(sheetTitle);
  if (!sh) {
    sh = wb.insertSheet(sheetTitle);
  } else {
    sh.clearContents();
    sh.clearFormats();
    // Remove any merged cells first
    try { sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart(); } catch (_) {}
  }
  return sh;
}

function cellToSeconds(cellVal) {
  if (cellVal === null || cellVal === undefined || cellVal === '') return 0;
  if (cellVal instanceof Date) {
    return cellVal.getHours() * 3600 + cellVal.getMinutes() * 60 + cellVal.getSeconds();
  }
  if (typeof cellVal === 'string') {
    const t = cellVal.trim();
    if (t.includes(':')) return hmsStringToSec(t);
    const n = parseFloat(t);
    return isNaN(n) ? 0 : n < 1 ? Math.round(n * 86400) : Math.round(n);
  }
  if (typeof cellVal === 'number') {
    return cellVal < 1 ? Math.round(cellVal * 86400) : Math.round(cellVal);
  }
  return 0;
}

function hmsStringToSec(hmsStr) {
  const parts = hmsStr.split(':').map(Number);
  if (parts.length === 3) return parts[0] * 3600 + parts[1] * 60 + parts[2];
  if (parts.length === 2) return parts[0] * 60 + parts[1];
  return 0;
}

function hmsOrBlank(totalSec) {
  if (!totalSec || isNaN(totalSec)) return '00:00:00';
  totalSec = Math.round(Math.abs(totalSec));
  const h = Math.floor(totalSec / 3600);
  const m = Math.floor((totalSec % 3600) / 60);
  const s = totalSec % 60;
  return [h, m, s].map(u => String(u).padStart(2, '0')).join(':');
}