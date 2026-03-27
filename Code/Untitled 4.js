/**
 * Look up calls and talktime for a specific agent, broken down by time bucket.
 * Results are written to a sheet named "Agent Lookup".
 *
 * HOW TO USE:
 *   Run getAgentTimeBucketReport() — it will prompt you for a name.
 *   OR call it directly: getAgentTimeBucketReport('Ravi Kumar')
 */
function getAgentTimeBucketReport(agentNameInput) {

  // ── 1. Get agent name ──────────────────────────────────────────────────────
  if (!agentNameInput) {
    const name = Browser.inputBox('Agent Lookup', 'Enter agent name (partial match works):', Browser.Buttons.OK_CANCEL);
    if (!name || name === 'cancel') return;
    agentNameInput = name.trim();
  }

  if (!agentNameInput) {
    Browser.msgBox('No agent name entered.');
    return;
  }

  // ── 2. Setup ───────────────────────────────────────────────────────────────
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const teamMapping = loadTeamMappingFromSheet();
  const searchLower = agentNameInput.toLowerCase();

  const timeBuckets = [
    '08-09 AM','09-10 AM','10-11 AM','11-12 PM',
    '12-01 PM','01-02 PM','02-03 PM','03-04 PM',
    '04-05 PM','05-06 PM','06-07 PM','07-08 PM',
    '08-09 PM','09-10 PM'
  ];

  const hourToBucket = {
     8:'08-09 AM',  9:'09-10 AM', 10:'10-11 AM', 11:'11-12 PM',
    12:'12-01 PM', 13:'01-02 PM', 14:'02-03 PM', 15:'03-04 PM',
    16:'04-05 PM', 17:'05-06 PM', 18:'06-07 PM', 19:'07-08 PM',
    20:'08-09 PM', 21:'09-10 PM'
  };

  const agentStats = {};

  function initAgent(name) {
    if (!agentStats[name]) {
      agentStats[name] = { _totals: { calls: 0, duration: 0 } };
      timeBuckets.forEach(b => { agentStats[name][b] = { calls: 0, duration: 0 }; });
    }
  }

  function isMatch(rawName) {
    if (!rawName) return false;
    const canonical = getCanonicalName(rawName, teamMapping);
    return (
      rawName.toLowerCase().includes(searchLower) ||
      canonical.toLowerCase().includes(searchLower) ||
      cleanAgentName(rawName).toLowerCase().includes(searchLower)
    );
  }

  function addRecord(rawName, bucket, duration) {
    if (!bucket || !timeBuckets.includes(bucket)) return;
    const canonical = getCanonicalName(rawName, teamMapping);
    initAgent(canonical);
    agentStats[canonical][bucket].calls    += 1;
    agentStats[canonical][bucket].duration += (duration || 0);
    agentStats[canonical]._totals.calls    += 1;
    agentStats[canonical]._totals.duration += (duration || 0);
  }

  // ── 3. Process main call records (Tata Smartflo) ───────────────────────────
  const dataSheet = ss.getActiveSheet();
  const headers   = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  const lastRow   = dataSheet.getLastRow();

  if (lastRow > 1) {
    const data             = dataSheet.getRange(2, 1, lastRow - 1, dataSheet.getLastColumn()).getValues();
    const agentNameIdx     = headers.indexOf('agent_name_cleaned');
    const agentMissedIdx   = headers.indexOf('agents_missed_cleaned');
    const callDurationIdx  = headers.indexOf('total_call_duration');
    const timeIdx          = headers.findIndex(h => h && h.toString().toLowerCase() === 'time');
    const effectiveTimeIdx = timeIdx !== -1 ? timeIdx : 11;

    data.forEach(row => {
      const agentName   = row[agentNameIdx]   || '';
      const agentMissed = row[agentMissedIdx] || '';
      const candidate   = agentName || agentMissed;
      if (!isMatch(candidate)) return;

      let duration = 0;
      const dv = row[callDurationIdx];
      if (dv) {
        if (typeof dv === 'number')                          duration = dv > 0 && dv < 1 ? Math.round(dv * 86400) : dv;
        else if (typeof dv === 'string' && dv.includes(':')) duration = hmsToSeconds(dv);
      }

      const hour   = getTimeBucket(row[effectiveTimeIdx]);
      const bucket = (hour >= 8 && hour <= 21) ? hourToBucket[hour] : null;
      addRecord(candidate, bucket, duration);
    });
  }

  // ── 4. Process WhatsApp calls ──────────────────────────────────────────────
  (function processWA() {
    const waSheet = ss.getSheetByName('Whatsapp Call');
    if (!waSheet || waSheet.getLastRow() < 2) return;
    const waHeaders = waSheet.getRange(1, 1, 1, waSheet.getLastColumn()).getValues()[0];
    const waData    = waSheet.getRange(2, 1, waSheet.getLastRow() - 1, waSheet.getLastColumn()).getValues();
    const tsIdx     = waHeaders.findIndex(h => h && h.toString().toLowerCase() === 'timestamp');
    const agIdx     = waHeaders.findIndex(h => h && h.toString().toLowerCase().includes('agent') && h.toString().toLowerCase().includes('name'));
    const durIdx    = waHeaders.findIndex(h => h && h.toString().toLowerCase().includes('total duration'));
    if (tsIdx === -1 || agIdx === -1) return;

    const dateConfig = getDateRangeConfig();
    const fromDate   = new Date(dateConfig.fromDate); fromDate.setHours(0,0,0,0);
    const toDate     = new Date(dateConfig.toDate);   toDate.setHours(23,59,59,999);

    waData.forEach(row => {
      const agentName = row[agIdx];
      if (!isMatch(agentName)) return;
      let ts;
      const tv = row[tsIdx];
      if (tv instanceof Date)          ts = new Date(tv.getTime());
      else if (typeof tv === 'string') ts = new Date(tv);
      else if (typeof tv === 'number') ts = new Date((tv - 25569) * 86400000);
      if (!ts || isNaN(ts.getTime()))  return;
      const tsDay = new Date(ts); tsDay.setHours(0,0,0,0);
      if (tsDay < fromDate || tsDay > toDate) return;
      const bucket = hourToBucket[ts.getHours()] || null;
      let duration = 0;
      if (durIdx !== -1 && row[durIdx]) {
        const dv = row[durIdx];
        if (dv instanceof Date)                              duration = dv.getHours()*3600 + dv.getMinutes()*60 + dv.getSeconds();
        else if (typeof dv === 'string' && dv.includes(':')) duration = hmsToSeconds(dv);
        else if (typeof dv === 'number')                     duration = dv < 1 ? Math.round(dv * 86400) : dv;
      }
      addRecord(agentName, bucket, duration);
    });
  })();

  // ── 5. Process Avyukta calls ───────────────────────────────────────────────
  (function processAvyukta() {
    const avSheet = ss.getSheetByName('Avyukta');
    if (!avSheet || avSheet.getLastRow() < 2) return;
    const avHeaders = avSheet.getRange(1, 1, 1, avSheet.getLastColumn()).getValues()[0];
    const avData    = avSheet.getRange(2, 1, avSheet.getLastRow() - 1, avSheet.getLastColumn()).getValues();
    const cdIdx     = avHeaders.findIndex(h => h && h.toString().toLowerCase() === 'call_date');
    const fnIdx     = avHeaders.findIndex(h => h && h.toString().toLowerCase() === 'full_name');
    const lenIdx    = avHeaders.findIndex(h => h && h.toString().toLowerCase() === 'length_in_sec');
    if (cdIdx === -1 || fnIdx === -1) return;

    const dateConfig = getDateRangeConfig();
    const fromDate   = new Date(dateConfig.fromDate); fromDate.setHours(0,0,0,0);
    const toDate     = new Date(dateConfig.toDate);   toDate.setHours(23,59,59,999);

    avData.forEach(row => {
      const agentName = row[fnIdx];
      if (!isMatch(agentName)) return;
      const duration = lenIdx !== -1 && row[lenIdx] > 0 ? row[lenIdx] : 0;
      if (!duration) return;
      let cd;
      const cv = row[cdIdx];
      if (cv instanceof Date)          cd = new Date(cv.getTime());
      else if (typeof cv === 'string') cd = new Date(cv);
      else if (typeof cv === 'number') cd = new Date((cv - 25569) * 86400000);
      if (!cd || isNaN(cd.getTime()))  return;
      const cdDay = new Date(cd); cdDay.setHours(0,0,0,0);
      if (cdDay < fromDate || cdDay > toDate) return;
      const bucket = hourToBucket[cd.getHours()] || null;
      addRecord(agentName, bucket, duration);
    });
  })();

  // ── 6. Check results ───────────────────────────────────────────────────────
  if (Object.keys(agentStats).length === 0) {
    Browser.msgBox('No data found for agent matching "' + agentNameInput + '".');
    return;
  }

  // ── 7. Write output sheet ──────────────────────────────────────────────────
  let outSheet = ss.getSheetByName('Agent Lookup');
  if (outSheet) outSheet.clear();
  else outSheet = ss.insertSheet('Agent Lookup');

  outSheet.getRange('A:Z').setFontFamily('Helvetica Neue').setFontSize(12);

  let row = 1;

  Object.entries(agentStats).forEach(([canonicalName, bucketData]) => {

    // Agent header
    outSheet.getRange(row, 1, 1, timeBuckets.length + 3).merge();
    outSheet.getRange(row, 1)
      .setValue(canonicalName + '  |  Team: ' + getAgentTeam(canonicalName, teamMapping))
      .setFontWeight('bold').setFontSize(14)
      .setBackground('#007AFF').setFontColor('#ffffff')
      .setHorizontalAlignment('center');
    outSheet.setRowHeight(row, 40);
    row++;

    // Column headers
    const colHeaders = ['Metric', ...timeBuckets, 'TOTAL'];
    outSheet.getRange(row, 1, 1, colHeaders.length).setValues([colHeaders]);
    outSheet.getRange(row, 1, 1, colHeaders.length)
      .setFontWeight('bold').setBackground('#F5F5F7')
      .setFontColor('#1D1D1F').setHorizontalAlignment('center');
    outSheet.setRowHeight(row, 32);
    row++;

    // Calls row
    const callsRow = ['Calls'];
    let totalCalls = 0;
    timeBuckets.forEach(b => {
      const c = bucketData[b].calls;
      callsRow.push(c > 0 ? c : '');
      totalCalls += c;
    });
    callsRow.push(totalCalls);
    outSheet.getRange(row, 1, 1, colHeaders.length).setValues([callsRow]);
    outSheet.getRange(row, 1).setFontWeight('bold');
    outSheet.getRange(row, 2, 1, colHeaders.length - 1).setHorizontalAlignment('center').setNumberFormat('0');
    timeBuckets.forEach((b, i) => {
      const c = bucketData[b].calls;
      const cell = outSheet.getRange(row, i + 2);
      if (c === 0) return;
      if      (c < 5)  cell.setBackground('#FFECEC').setFontColor('#E65100');
      else if (c < 15) cell.setBackground('#F1F8E9').setFontColor('#558B2F');
      else if (c < 30) cell.setBackground('#E8F5E9').setFontColor('#2E7D32');
      else             cell.setBackground('#C8E6C9').setFontColor('#1B5E20');
    });
    outSheet.getRange(row, colHeaders.length).setFontWeight('bold').setBackground('#E3F2FD').setFontColor('#0D47A1');
    row++;

    // Talktime row
    const ttRow = ['Talktime'];
    let totalDuration = 0;
    timeBuckets.forEach(b => {
      const d = bucketData[b].duration;
      ttRow.push(d > 0 ? secondsToHMS(d) : '');
      totalDuration += d;
    });
    ttRow.push(secondsToHMS(totalDuration));
    outSheet.getRange(row, 1, 1, colHeaders.length).setValues([ttRow]);
    outSheet.getRange(row, 1).setFontWeight('bold');
    outSheet.getRange(row, 2, 1, colHeaders.length - 1).setHorizontalAlignment('center');
    timeBuckets.forEach((b, i) => {
      const d = bucketData[b].duration;
      const cell = outSheet.getRange(row, i + 2);
      if (d === 0) return;
      if      (d < 300)  cell.setBackground('#FFECEC').setFontColor('#E65100');
      else if (d < 900)  cell.setBackground('#F1F8E9').setFontColor('#558B2F');
      else if (d < 1800) cell.setBackground('#E8F5E9').setFontColor('#2E7D32');
      else               cell.setBackground('#C8E6C9').setFontColor('#1B5E20');
    });
    outSheet.getRange(row, colHeaders.length).setFontWeight('bold').setBackground('#E3F2FD').setFontColor('#0D47A1');
    row++;

    row += 2; // spacer
  });

  // Column widths
  outSheet.setColumnWidth(1, 110);
  for (let c = 2; c <= timeBuckets.length + 2; c++) outSheet.setColumnWidth(c, 105);
  outSheet.setFrozenColumns(1);

  // Timestamp
  outSheet.getRange(row, 1)
    .setValue('✓ Looked up "' + agentNameInput + '" on ' + new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }))
    .setFontColor('#0f9d58').setFontWeight('bold');

  ss.setActiveSheet(outSheet);
  Browser.msgBox('Done! Found ' + Object.keys(agentStats).length + ' matching agent(s). Check the "Agent Lookup" sheet.');
}