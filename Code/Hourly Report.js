// ===== CONFIGURATION — only sheet name, webhook & label needed =====
const REPORT_CONFIG = [
  {
    sheetName: 'Sales Team',
    webhookUrl: 'https://chat.googleapis.com/v1/spaces/AAQAG1cWOtk/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=y-zRy3i2I55kLskJ6Y61XJgl5mw9YiJhMtsznNopSLM',
    label: 'Sales Team – HOURLY PERFORMANCE',
  },
  {
    sheetName: 'HT Done Team',
    webhookUrl: 'https://chat.googleapis.com/v1/spaces/AAQAYrMDbMU/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=yWxMdq6gNO8mV8-624qi3RRtt-G8mPBQlrOTLBrrqXs',
    label: 'HT Done Team – HOURLY PERFORMANCE',
  },
  {
    sheetName: 'URoots Sales Team',
    webhookUrl: 'https://chat.googleapis.com/v1/spaces/AAQAI2Uq6xo/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=8btKYNiy0j2mZ55FnJWm9YmIBD_zoDcVV_45cwLszdc',
    label: 'URoots Sales Team – HOURLY PERFORMANCE',
  },
  {
    sheetName: 'URoots Confirmation Team',
    webhookUrl: 'https://chat.googleapis.com/v1/spaces/AAQA0IaYnfI/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=tbiuLbpmGbWFUv9P6M107tU7yst87tR57LoDr7J9P6c',
    label: 'URoots Confirmation Team – HOURLY PERFORMANCE',
  },
  {
    sheetName: 'Incoming Dept.',
    webhookUrl: 'https://chat.googleapis.com/v1/spaces/AAQAG1cWOtk/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=y-zRy3i2I55kLskJ6Y61XJgl5mw9YiJhMtsznNopSLM',
    label: 'Incoming Dept – HOURLY PERFORMANCE',
  },
  {
    sheetName: 'Welcome Team',
    webhookUrl: 'https://chat.googleapis.com/v1/spaces/AAQAL6s5j8c/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=i5oJ-Y3Bex39jy6g4Y-6mBHXCmKEMphScz50-UdrOVY',
    label: 'Welcome Team – HOURLY PERFORMANCE',
  },
  {
    sheetName: 'Datealignment Team',
    webhookUrl: 'https://chat.googleapis.com/v1/spaces/AAQAgHe8HPU/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=DxbMQtedJPAyi8bCg_lGOtqnTPxpV3fV2xNqvywoi3s',
    label: 'Datealignment Team – HOURLY PERFORMANCE',
  },
];


// ===== 4-HOUR WINDOWS =====
// Define your windows as arrays of header labels exactly as they appear in your sheet.
// Adjust these to match your actual column names.
const FOUR_HOUR_WINDOWS = [
  {
    label: '9 AM – 1 PM',
    startHour: 9,
    endHour: 13,
    buckets: ['9 AM-10 AM', '10 AM-11 AM', '11 AM-12 PM', '12 PM-1 PM']
  },
  {
    label: '1 PM – 5 PM',
    startHour: 13,
    endHour: 17,
    buckets: ['1 PM-2 PM', '2 PM-3 PM', '3 PM-4 PM', '4 PM-5 PM']
  },
  {
    label: '5 PM – 9 PM',
    startHour: 17,
    endHour: 21,
    buckets: ['5 PM-6 PM', '6 PM-7 PM', '7 PM-8 PM', '8 PM-9 PM']
  },
];


// ===== GET CURRENT 4-HOUR WINDOW =====
function getCurrentWindow() {
  const now = new Date();
  const currentHour = now.getHours(); // 0–23

  for (const window of FOUR_HOUR_WINDOWS) {
    if (currentHour >= window.startHour && currentHour < window.endHour) {
      return window;
    }
  }

  // If outside all defined windows, return null (no report sent)
  return null;
}


// ===== MAIN — point your trigger here =====
function sendAllHourlyReports() {
  const currentWindow = getCurrentWindow();

  if (!currentWindow) {
    Logger.log('⏸ Outside of all defined 4-hour windows. No report sent.');
    return;
  }

  Logger.log(`📦 Current window: ${currentWindow.label}`);
  REPORT_CONFIG.forEach(config => sendReportForTeam(config, currentWindow));
}


// ===== DYNAMIC ROW FINDER =====
function findRowContaining(sheet, keyword, searchCol = 1) {
  const data = sheet.getRange(1, searchCol, sheet.getLastRow(), 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase().includes(keyword.toLowerCase())) {
      return i + 1;
    }
  }
  return -1;
}

function findTableEndRow(sheet, startRow) {
  const lastRow = sheet.getLastRow();
  let endRow = startRow;
  let seenData = false;

  for (let r = startRow + 1; r <= lastRow; r++) {
    const val = sheet.getRange(r, 1).getValue();
    if (val !== '' && val !== null) {
      seenData = true;
      endRow = r;
    } else if (seenData) {
      break;
    }
  }
  return endRow;
}


// ===== CORE LOGIC =====
function sendReportForTeam(config, currentWindow) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(config.sheetName);

    if (!sheet) {
      Logger.log(`Sheet not found: ${config.sheetName}`);
      return;
    }

    // --- Dynamically locate the two sections ---
    const callsHeaderRow = findRowContaining(sheet, 'HOURLY CALLS');
    const talkHeaderRow  = findRowContaining(sheet, 'HOURLY TALKTIME');

    if (callsHeaderRow === -1 || talkHeaderRow === -1) {
      Logger.log(`Could not find section headers in: ${config.sheetName}`);
      return;
    }

    const callsDataStart = callsHeaderRow + 1;
    const talkDataStart  = talkHeaderRow  + 1;

    const callsDataEnd = findTableEndRow(sheet, callsDataStart);
    const talkDataEnd  = findTableEndRow(sheet, talkDataStart);

    Logger.log(`${config.sheetName} → Calls rows ${callsDataStart}–${callsDataEnd} | Talk rows ${talkDataStart}–${talkDataEnd}`);

    const lastCol = sheet.getLastColumn();

    const callsData = sheet
      .getRange(callsDataStart, 1, callsDataEnd - callsDataStart + 1, lastCol)
      .getDisplayValues();

    const talkData = sheet
      .getRange(talkDataStart, 1, talkDataEnd - talkDataStart + 1, lastCol)
      .getDisplayValues();

   

// ===== FILTER HEADERS TO CURRENT 4-HOUR WINDOW =====

// 1️⃣ Read all hour headers from sheet (skip Agent column)
const allHeaders = callsData[0].slice(1);

// 2️⃣ Extract start hour from header like "08-09 AM", "01-02 PM"
function extractStartHour(header) {
  const match = header.match(/(\d{1,2})\s*-\s*\d{1,2}\s*(AM|PM)/i);
  if (!match) return null;

  let hour = parseInt(match[1], 10);
  const period = match[2].toUpperCase();

  if (period === 'PM' && hour !== 12) hour += 12;
  if (period === 'AM' && hour === 12) hour = 0;

  return hour;
}

// 3️⃣ Keep only headers that fall inside current 4-hour window
const headers = allHeaders.filter(h => {
  if (!h) return false;
  const startHour = extractStartHour(h.toString());
  return (
    startHour !== null &&
    startHour >= currentWindow.startHour &&
    startHour < currentWindow.endHour
  );
});

// 4️⃣ Safety log (keep this)
if (headers.length === 0) {
  Logger.log(`⚠️ No matching hour columns for window ${currentWindow.label} in: ${config.sheetName}`);
  Logger.log(`Sheet headers found: ${allHeaders.join(' | ')}`);
  return;
}

    const agents = callsData.slice(1).map(r => r[0]).filter(a => a && a !== 'TEAM TOTAL');

    // ===== BUILD LOOKUP =====
    const map = {};
    for (let i = 1; i < callsData.length; i++) {
      const agent = callsData[i][0];
      if (!agent || agent === 'TEAM TOTAL') continue;
      headers.forEach((h) => {
        const colIndex = allHeaders.indexOf(h) + 1; // +1 because col 0 is agent name
        map[agent + '|' + h] = map[agent + '|' + h] || {};
        map[agent + '|' + h].calls = callsData[i][colIndex] || '0';
      });
    }
    for (let i = 1; i < talkData.length; i++) {
      const agent = talkData[i][0];
      if (!agent || agent === 'TEAM TOTAL') continue;
      headers.forEach((h) => {
        const colIndex = allHeaders.indexOf(h) + 1;
        map[agent + '|' + h] = map[agent + '|' + h] || {};
        map[agent + '|' + h].talk = talkData[i][colIndex] || '00:00:00';
      });
    }

    // ===== SORT AGENTS BY TOTAL TALKTIME (descending) =====
    agents.sort((a, b) => {
      const getTotalTalk = (agent) => {
        let totalSeconds = 0;
        headers.forEach(h => {
          const talk = (map[agent + '|' + h]?.talk) || '00:00:00';
          const parts = talk.split(':').map(Number);
          if (parts.length === 3) totalSeconds += parts[0] * 3600 + parts[1] * 60 + parts[2];
        });
        return totalSeconds;
      };
      return getTotalTalk(b) - getTotalTalk(a);
    });

    // ===== BUILD TABLE =====
    const BUCKET_WIDTH = 18;
    const table = [];

    const header1 = ['Agent'];
    headers.forEach(h => header1.push(h.padEnd(BUCKET_WIDTH, ' ')));
    table.push(header1);

    const header2 = [''];
    headers.forEach(() => header2.push('Calls   |   TT'.padEnd(BUCKET_WIDTH, ' ')));
    table.push(header2);

    agents.forEach(agent => {
      const row = [agent];
      headers.forEach(h => {
        const cell = map[agent + '|' + h] || {};
        const d = (cell.calls || '0').padStart(3, ' ');
        const t = (cell.talk  || '00:00:00').padEnd(8, ' ');
        row.push(`${d} | ${t}`.padEnd(BUCKET_WIDTH, ' '));
      });
      table.push(row);
    });

    // ===== CALCULATE WIDTHS =====
    const widths = [];
    table.forEach(r =>
      r.forEach((c, i) => widths[i] = Math.max(widths[i] || 0, c.length))
    );

    // ===== RENDER TEXT =====
    const now = new Date();
    const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd MMM yyyy');
    const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'hh:mm a');

    let text = `${config.label}\n`;
    text += `📅 ${dateStr}   🕐 ${timeStr}   📦 Window: ${currentWindow.label}\n\n`;

    table.forEach((row, i) => {
      text += row.map((c, j) => c.padEnd(widths[j], ' ')).join('   ') + '\n';
      if (i === 1) text += widths.map(w => '-'.repeat(w)).join('   ') + '\n';
    });

    // ===== SEND TO GCHAT =====
    UrlFetchApp.fetch(config.webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: '```\n' + text.trim() + '\n```' })
    });

    Logger.log(`✅ Sent: ${config.sheetName}`);

  } catch (e) {
    Logger.log(`❌ Error (${config.sheetName}): ${e.message}`);
  }
}