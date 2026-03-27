// ============================================================================
// TEAM CONFIGURATION
// ============================================================================

const SALES_SHEET_ID = "1Dad77vTHkxTh8jPag6h3l_Ygxv6RWmXvMx3lIk58_Zk";

const SALES_ALERT_MIN_CALLS = 150;
const SALES_ALERT_MIN_TALKTIME_MINUTES = 120;

const SALES_TEAMS = [
  {
    name: "HT Done Team",
    sheetName: "HT Done Team",
    webhookUrl: "//https://chat.googleapis.com/v1/spaces/AAQAYrMDbMU/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=VBNW6I-5ndVEzKPEquE5Y7Dx8vihUFhrMuGFmijG6fA"
  },
  {
    name: "URoots Sales Team",
    sheetName: "URoots Sales Team",
    webhookUrl: "//https://chat.googleapis.com/v1/spaces/AAQAj-5GaIo/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=Fdwcy_s6HSXCqXcmlOyw92Uk5qTQWkNy4mkYIV6Zu7s"
  },
  {
    name: "Sales Team",
    sheetName: "Sales Team",
    webhookUrl: "https://chat.googleapis.com/v1/spaces/oyQZRiAAAAE/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=p_GrcZfuy_qil9ZAAXyUmOfi19q2rJ4O40-ROkjoYUQ"
  },
  {
    name: "Datealignment Team",
    sheetName: "Datealignment Team",
    webhookUrl: "//https://chat.googleapis.com/v1/spaces/AAQAyvbZ5mw/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=qtQV6UP92IyZpdRw7ZHdk24zNNl-FjraJLHChOQ2DMM"
  },
  {
    name: "Welcome Team",
    sheetName: "Welcome Team",
    webhookUrl: "//https://chat.googleapis.com/v1/spaces/AAQAL6s5j8c/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=atwKo-WjpSrW0TV72h5YotX0uB5F5DRkZbp6iJpi_ag"
  },
  {
    name: "URoots Confirmation Team",
    sheetName: "URoots Confirmation Team",
    webhookUrl: "//https://chat.googleapis.com/v1/spaces/AAQA0IaYnfI/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=p30kCHY7S4994krXSJHuLOJdgn2LE4ZkoMQw_Wd7Ad4"
  },
  {
    name: "Incoming Dept.",
    sheetName: "Incoming Dept.",
    webhookUrl: "//https://chat.googleapis.com/v1/spaces/AAQAG1cWOtk/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=S4-lDKzOyBJd3XGhZBGsZZVghsv-tJn-pr4mZjbiVyg"
  }
];

// ============================================================================
// MAIN — Send All Reports
// ============================================================================

function salesReport_sendAllReports() {
  Logger.log("=== Starting 5PM alert reports for all teams ===");
  let successCount = 0;
  let failCount = 0;

  SALES_TEAMS.forEach(team => {
    try {
      Logger.log(`Sending: ${team.name}`);
      salesReport_sendTeamReport(team);
      successCount++;
      Logger.log(`✅ ${team.name} - SUCCESS`);
      Utilities.sleep(1000);
    } catch (error) {
      failCount++;
      Logger.log(`❌ ${team.name} - FAILED: ${error.message}`);
      salesReport_sendErrorNotification(team.name, error.message);
    }
  });

  Logger.log(`=== Done: ${successCount} success, ${failCount} failed ===`);
}

// ============================================================================
// SEND — Single Team Report
// ============================================================================

function salesReport_sendTeamReport(team) {
  const data = salesReport_getSheetData(SALES_SHEET_ID, team.sheetName);
  const message = salesReport_buildGroupedCard(team.name, data, true);
  salesReport_sendToChat(team.webhookUrl, message);
}

function salesReport_sendRegularTeamReport(team) {
  const data = salesReport_getSheetData(SALES_SHEET_ID, team.sheetName);
  const message = salesReport_buildGroupedCard(team.name, data, false);
  salesReport_sendToChat(team.webhookUrl, message);
}

// ============================================================================
// DATA — Fetch Sheet
// ============================================================================

function salesReport_getSheetData(sheetId, sheetName) {
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);

  const range = sheet.getRange(7, 1, 14, 15);
  return {
    values: range.getValues(),
    display: range.getDisplayValues()
  };
}

// ============================================================================
// CORE LOGIC — Group Agents into Tiers
// ============================================================================

function salesReport_groupAgents(values, display) {
  const groups = {
    critical: [],   // 🔴 Very low — needs coaching
    monitor: [],    // 🟡 Below both targets
    push: [],       // 🟢 Close — one target met
    onTrack: []     // ✅ Meets requirement
  };

  values.forEach((row, i) => {
    const name = row[0];
    const calls = row[1];
    const talktimeStr = display[i][9];
    const mins = salesReport_parseTimeToMinutes(talktimeStr);

    if (!name || calls <= 0) return;

    const meetsCalls = calls >= SALES_ALERT_MIN_CALLS;
    const meetsTalktime = mins >= SALES_ALERT_MIN_TALKTIME_MINUTES;

    if (meetsCalls || meetsTalktime) {
      groups.onTrack.push({ name, calls, talktimeStr });
    } else if (calls < 40 && mins < 90) {
      groups.critical.push({ name, calls, talktimeStr });
    } else if (calls < SALES_ALERT_MIN_CALLS && mins < SALES_ALERT_MIN_TALKTIME_MINUTES) {
      groups.monitor.push({ name, calls, talktimeStr });
    } else {
      groups.push.push({ name, calls, talktimeStr });
    }
  });

  return groups;
}

// ============================================================================
// CARD BUILDER — Grouped cardsV2 Format
// ============================================================================

function salesReport_buildGroupedCard(teamName, data, is5PM) {
  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMMM dd, yyyy");
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "hh:mm a");

  const groups = salesReport_groupAgents(data.values, data.display);

  const alertLabel = is5PM ? "– 5 PM ALERT" : "– REPORT";

  const cards = [
    // Summary card
    {
      cardId: "summary",
      card: {
        header: {
          title: `📊 ${teamName.toUpperCase()} ${alertLabel}`,
          subtitle: `${dateStr} • ${timeStr}`
        },
        sections: [{
          widgets: [{
            textParagraph: {
              text:
                `<b>Target:</b> ${SALES_ALERT_MIN_CALLS} calls OR ${SALES_ALERT_MIN_TALKTIME_MINUTES} mins talktime<br>` +
                `✅ On Track: <b>${groups.onTrack.length}</b>  ` +
                `🟢 Push: <b>${groups.push.length}</b>  ` +
                `🟡 Monitor: <b>${groups.monitor.length}</b>  ` +
                `🔴 Critical: <b>${groups.critical.length}</b>`
            }
          }]
        }]
      }
    },

    // ✅ On Track
    salesReport_groupCard(
      "✅ On Track – Meeting Target",
      "Calls OR talktime target met",
      "00aa00",
      groups.onTrack
    ),

    // 🟢 Push
    salesReport_groupCard(
      "🟢 Close – Final Push",
      "One target nearly met",
      "44bb44",
      groups.push
    ),

    // 🟡 Monitor
    salesReport_groupCard(
      "🟡 Below Target – Monitor",
      "Both targets below threshold",
      "ffcc00",
      groups.monitor
    ),

    // 🔴 Critical
    salesReport_groupCard(
      "🔴 Critical – Immediate Action",
      "Far below target — needs coaching",
      "ff0000",
      groups.critical
    )

  ].filter(Boolean); // removes null cards (empty groups)

  return { cardsV2: cards };
}

function salesReport_groupCard(title, subtitle, color, agents) {
  if (!agents.length) return null;

  return {
    cardId: title,
    card: {
      header: {
        title: title,
        subtitle: subtitle,
        imageUrl: `https://singlecolorimage.com/get/${color}/60x60`,
        imageType: "CIRCLE"
      },
      sections: [{
        widgets: agents.map(a => ({
          textParagraph: {
            text: `<b>${a.name}</b> — ${a.calls} calls | ${a.talktimeStr}`
          }
        }))
      }]
    }
  };
}

// ============================================================================
// HELPER — Parse Time String to Minutes
// ============================================================================

function salesReport_parseTimeToMinutes(timeStr) {
  if (!timeStr || timeStr === "") return 0;
  const parts = timeStr.toString().split(":");
  if (parts.length === 3) {
    return (parseInt(parts[0]) || 0) * 60 +
           (parseInt(parts[1]) || 0) +
           (parseInt(parts[2]) || 0) / 60;
  }
  return 0;
}

// ============================================================================
// SEND — Post to Google Chat
// ============================================================================

function salesReport_sendToChat(webhookUrl, message) {
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(message),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(webhookUrl, options);
  if (response.getResponseCode() !== 200) {
    throw new Error("Chat API error: " + response.getContentText());
  }
}

// ============================================================================
// ERROR NOTIFICATION
// ============================================================================

function salesReport_sendErrorNotification(teamName, errorMessage) {
  const firstTeam = SALES_TEAMS.find(t => t.webhookUrl && !t.webhookUrl.startsWith("//"));
  if (!firstTeam) return;

  try {
    UrlFetchApp.fetch(firstTeam.webhookUrl, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ text: `⚠️ Error for ${teamName}: ${errorMessage}` }),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log("Failed to send error notification: " + e);
  }
}

// ============================================================================
// SCHEDULED TRIGGERS
// ============================================================================

function salesReport_sendMorningReport()   { salesReport_sendRegularReports(); }
function salesReport_sendAfternoonReport() { salesReport_sendRegularReports(); }
function salesReport_send5PMAlert()        { salesReport_sendAllReports(); }

function salesReport_sendRegularReports() {
  Logger.log("=== Sending regular reports ===");
  let successCount = 0, failCount = 0;

  SALES_TEAMS.forEach(team => {
    try {
      salesReport_sendRegularTeamReport(team);
      successCount++;
      Logger.log(`✅ ${team.name}`);
      Utilities.sleep(1000);
    } catch (error) {
      failCount++;
      Logger.log(`❌ ${team.name}: ${error.message}`);
    }
  });

  Logger.log(`=== Done: ${successCount} success, ${failCount} failed ===`);
}

// ============================================================================
// TEST — Single Team
// ============================================================================

function salesReport_testSingleTeam() {
  const TEST_TEAM_NAME = "Sales Team"; // 👈 Change to test any team

  // Available:
  // "HT Done Team" | "URoots Sales Team" | "Sales Team"
  // "Datealignment Team" | "Welcome Team"
  // "URoots Confirmation Team" | "Incoming Dept."

  const team = SALES_TEAMS.find(t => t.name === TEST_TEAM_NAME);

  if (!team) {
    Logger.log(`❌ Team not found: "${TEST_TEAM_NAME}"`);
    Logger.log("Available: " + SALES_TEAMS.map(t => t.name).join(", "));
    return;
  }

  Logger.log(`🧪 Testing: ${team.name}`);
  Logger.log(`   Sheet tab : ${team.sheetName}`);
  Logger.log(`   Webhook   : ${team.webhookUrl.substring(0, 60)}...`);

  try {
    salesReport_sendTeamReport(team);
    Logger.log("✅ Done! Check Google Chat for the grouped cards.");
  } catch (error) {
    Logger.log(`❌ Failed: ${error.message}`);
  }
}