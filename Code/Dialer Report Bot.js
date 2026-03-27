/**
 * MULTI-TEAM SALES REPORTS
 * Automatically sends reports to all TEAMS1 3 times daily
 */

// ============================================================================
// TEAM CONFIGURATION - ALL TEAMS1 IN ONE SHEET!
// ============================================================================

// Your Google Sheet ID (same for all TEAMS1 - they're all tabs in one sheet)
const SHEET_ID1 = "1Dad77vTHkxTh8jPag6h3l_Ygxv6RWmXvMx3lIk58_Zk";

// Each team has their own tab and webhook
const TEAMS1 = [
 {
   name: "HT Done Team",
   sheetName: "HT Done Team",                // Exact tab name (check bottom of sheet)
   webhookUrl: "https://chat.googleapis.com/v1/spaces/AAQAYrMDbMU/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=VBNW6I-5ndVEzKPEquE5Y7Dx8vihUFhrMuGFmijG6fA"      // From Google Chat → Manage webhooks
 },
 {
   name: "URoots Sales Team",
   sheetName: "URoots Sales Team",
   webhookUrl: "https://chat.googleapis.com/v1/spaces/AAQAI2Uq6xo/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=8btKYNiy0j2mZ55FnJWm9YmIBD_zoDcVV_45cwLszdc"
 },
 {
   name: "Sales Team",
   sheetName: "Sales Team",
   webhookUrl: "https://chat.googleapis.com/v1/spaces/AAQAG1cWOtk/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=y-zRy3i2I55kLskJ6Y61XJgl5mw9YiJhMtsznNopSLM"  // ✅ Already configured!
 },
 {
   name: "Datealignment Team",
   sheetName: "Datealignment Team",
   webhookUrl: "https://chat.googleapis.com/v1/spaces/AAQAyvbZ5mw/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=qtQV6UP92IyZpdRw7ZHdk24zNNl-FjraJLHChOQ2DMM"
 },
 {
   name: "Welcome Team",
   sheetName: "Welcome Team",
   webhookUrl: "https://chat.googleapis.com/v1/spaces/AAQAL6s5j8c/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=atwKo-WjpSrW0TV72h5YotX0uB5F5DRkZbp6iJpi_ag"
 },
 {
   name: "URoots Confirmation Team",
   sheetName: "URoots Confirmation Team",
   webhookUrl: "https://chat.googleapis.com/v1/spaces/AAQA0IaYnfI/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=p30kCHY7S4994krXSJHuLOJdgn2LE4ZkoMQw_Wd7Ad4"
 },
 {
   name: "Incoming Dept.",
   sheetName: "Incoming Dept.",
   webhookUrl: "https://chat.googleapis.com/v1/spaces/AAQAG1cWOtk/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=y-zRy3i2I55kLskJ6Y61XJgl5mw9YiJhMtsznNopSLM"
 }
];



// ============================================================================
// MAIN FUNCTION - Sends Reports to ALL TEAMS1
// ============================================================================

function sendAllReports() {
  Logger.log("=== Starting reports for all TEAMS1 ===");
  
  let successCount = 0;
  let failCount = 0;
  
  TEAMS1.forEach(team => {
    try {
      Logger.log(`Sending report for: ${team.name}`);
      sendTeamReport(team);
      successCount++;
      Logger.log(`✅ ${team.name} - SUCCESS`);
      
      // Wait 1 second between TEAMS1 to avoid rate limits
      Utilities.sleep(1000);
      
    } catch (error) {
      failCount++;
      Logger.log(`❌ ${team.name} - FAILED: ${error.message}`);
      sendErrorNotification(team.name, error.message);
    }
  });
  
  Logger.log(`=== Complete: ${successCount} success, ${failCount} failed ===`);
}

// ============================================================================
// Send Report for Single Team
// ============================================================================

function sendTeamReport(team) {
  // Get data from team's tab (using global SHEET_ID1)
  const data = getSheetData(SHEET_ID1, team.sheetName);
  
  // Format the message
  const message = formatReport(team.name, data);
  
  // Send to team's Google Chat
  sendToChat(team.webhookUrl, message);
}

// ============================================================================
// Get Data from Sheet
// ============================================================================

function getSheetData(sheetId, sheetName) {
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error(`Sheet not found: ${sheetName}`);
  }
  
  const lastCol = 15;
  const startRow = 7;
  const maxRows = 50;
  
  // Get a larger range to find TEAM TOTAL
  const fullRange = sheet.getRange(startRow, 1, maxRows, lastCol);
  const allValues = fullRange.getValues();
  const allDisplay = fullRange.getDisplayValues();
  
  // Find where TEAM TOTAL row is
  let teamTotalIndex = -1;
  for (let i = 0; i < allValues.length; i++) {
    const firstCell = allValues[i][0];
    if (firstCell && firstCell.toString().toUpperCase().includes("TEAM TOTAL")) {
      teamTotalIndex = i;
      break;
    }
  }
  
  // Only return data up to (but not including) TEAM TOTAL
  let agentValues, agentDisplay;
  if (teamTotalIndex > 0) {
    agentValues = allValues.slice(0, teamTotalIndex);
    agentDisplay = allDisplay.slice(0, teamTotalIndex);
  } else {
    agentValues = allValues.slice(0, 14);
    agentDisplay = allDisplay.slice(0, 14);
  }
  
  return {
    agentValues: agentValues,
    agentDisplay: agentDisplay
  };
}

// ============================================================================
// Format the Report
// ============================================================================

function formatReport(teamName, data) {
  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMMM dd, yyyy");
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "hh:mm a");
  
  // Column indices
  const nameCol = 0;
  const totalCallsCol = 1;
  const talktimeCol = 9;
  
  // Build agent list
  const agentList = formatAgentList(data.agentValues, data.agentDisplay, nameCol, totalCallsCol, talktimeCol);
  
  const message = {
    cards: [{
      header: {
        title: `📊 ${teamName.toUpperCase()} REPORT`,
        subtitle: dateStr + " • " + timeStr
      },
      sections: [
        {
          header: "👥 Agent Performance",
          widgets: [{
            textParagraph: {
              text: agentList
            }
          }]
        }
      ]
    }]
  };
  
  return message;
}

// ============================================================================
// Format Agent List
// ============================================================================

function formatAgentList(agentValues, agentDisplay, nameCol, callsCol, talktimeCol) {
  let text = "";
  
  agentValues.forEach((row, index) => {
    if (row[callsCol] > 0) {
      const name = row[nameCol];
      const calls = row[callsCol];
      const talktime = agentDisplay[index][talktimeCol];
      
      text += name + ": " + calls + " calls | " + talktime + "\n\n";
    }
  });
  
  return text || "No data available";
}

// ============================================================================
// Send to Google Chat
// ============================================================================

function sendToChat(webhookUrl, message) {
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
// Error Notification
// ============================================================================

function sendErrorNotification(teamName, errorMessage) {
  // Send error to the first configured team's chat
  const firstTeam = TEAMS1.find(t => t.webhookUrl.includes("chat.googleapis.com"));
  
  if (!firstTeam) return;
  
  const message = {
    text: `⚠️ Error sending report for ${teamName}: ${errorMessage}`
  };
  
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(message),
    muteHttpExceptions: true
  };
  
  try {
    UrlFetchApp.fetch(firstTeam.webhookUrl, options);
  } catch (e) {
    Logger.log("Failed to send error notification: " + e);
  }
}

// ============================================================================
// TEST FUNCTIONS
// ============================================================================

function testSingleTeam() {
  // Test with Sales Team (already configured)
  const salesTeam = TEAMS1.find(t => t.name === "Sales Team");
  
  if (!salesTeam) {
    Logger.log("Sales Team not found in config");
    return;
  }
  
  Logger.log("Testing Sales Team...");
  sendTeamReport(salesTeam);
  Logger.log("✅ Test complete! Check Google Chat.");
}

function testAllConfiguredTEAMS1() {
  // Test only TEAMS1 that have webhook URLs configured
  Logger.log("Testing all configured TEAMS1...");
  
  TEAMS1.forEach(team => {
    if (team.webhookUrl && !team.webhookUrl.includes("YOUR_WEBHOOK")) {
      try {
        Logger.log(`Testing: ${team.name}`);
        sendTeamReport(team);
        Logger.log(`✅ ${team.name} - SUCCESS`);
        Utilities.sleep(1000);
      } catch (error) {
        Logger.log(`❌ ${team.name} - FAILED: ${error.message}`);
      }
    } else {
      Logger.log(`⏭️ Skipping ${team.name} - not configured yet`);
    }
  });
  
  Logger.log("✅ Test complete!");
}

// ============================================================================
// SCHEDULED FUNCTIONS (for triggers)
// ============================================================================

function sendMorningReport() {
  sendAllReports();
}

function sendAfternoonReport() {
  sendAllReports();
}

function sendEveningReport() {
  sendAllReports();
}

// ============================================================================
// HELPER: List Team Configuration Status
// ============================================================================

function checkConfiguration() {
  Logger.log("=== TEAM CONFIGURATION STATUS ===");
  Logger.log(`Sheet ID: ${SHEET_ID1}`);
  Logger.log("");
  
  TEAMS1.forEach(team => {
    const hasWebhook = team.webhookUrl && !team.webhookUrl.includes("YOUR_WEBHOOK");
    
    const status = hasWebhook ? "✅ READY" : "⚠️ NEEDS WEBHOOK";
    
    Logger.log(`${team.name}: ${status}`);
    Logger.log(`  Tab: "${team.sheetName}"`);
    if (!hasWebhook) Logger.log(`  Missing: Webhook URL`);
  });
  
  Logger.log("=================================");
}





// ============================================================================
// EXACT TIME TRIGGER SYSTEM
// Fires sendAllReports at EXACTLY 1:00 PM, 4:00 PM, 6:00 PM IST daily
// ============================================================================

const REPORT_HOURS_IST = [13, 16, 18]; // 1 PM, 4 PM, 6 PM in 24hr

// ============================================================================
// STEP 1: Run this ONCE manually to bootstrap the whole system
// ============================================================================

function bootstrapTriggers() {
  deleteAllExactTriggers(); // Clean slate

  REPORT_HOURS_IST.forEach(hour => {
    scheduleNextTriggerForHour(hour);
  });

  Logger.log("✅ Bootstrap complete! Triggers scheduled for 1 PM, 4 PM, 6 PM IST.");
}

// ============================================================================
// STEP 2: This is what actually fires at each scheduled time
// It sends the report, then reschedules itself for the same time next day
// ============================================================================

function sendReportAndReschedule() {
  // Get the current trigger's hour so we know which time slot fired
  const now = new Date();
  const istOffset = 5.5 * 60 * 60 * 1000; // IST = UTC+5:30
  const istNow = new Date(now.getTime() + istOffset);
  const currentHourIST = istNow.getUTCHours();

  Logger.log(`⏰ Triggered at IST hour: ${currentHourIST}`);

  // Send the reports
  try {
    sendAllReports();
    Logger.log("✅ Reports sent successfully.");
  } catch (e) {
    Logger.log("❌ Failed to send reports: " + e.message);
  }

  // Find the closest matching configured hour and reschedule for next day
  const matchedHour = REPORT_HOURS_IST.reduce((prev, curr) => {
    return Math.abs(curr - currentHourIST) < Math.abs(prev - currentHourIST) ? curr : prev;
  });

  Logger.log(`🔁 Rescheduling for hour ${matchedHour} IST tomorrow...`);
  scheduleNextTriggerForHour(matchedHour);
}

// ============================================================================
// CORE HELPER: Schedules a one-time trigger for a specific IST hour (next occurrence)
// ============================================================================

function scheduleNextTriggerForHour(hourIST) {
  const now = new Date();
  const istOffset = 5.5 * 60 * 60 * 1000;
  const nowIST = new Date(now.getTime() + istOffset);

  // Build target time: today at hourIST:00:00 IST
  const target = new Date(Date.UTC(
    nowIST.getUTCFullYear(),
    nowIST.getUTCMonth(),
    nowIST.getUTCDate(),
    hourIST - 5,   // Convert IST hour to UTC: subtract 5 hours
    30,            // subtract 30 minutes (for the :30 in UTC+5:30)
    0, 0
  ));

  // If that time has already passed today, schedule for tomorrow
  if (target <= now) {
    target.setUTCDate(target.getUTCDate() + 1);
  }

  ScriptApp.newTrigger("sendReportAndReschedule")
    .timeBased()
    .at(target)
    .create();

  const label = `${hourIST > 12 ? hourIST - 12 : hourIST}:00 ${hourIST >= 12 ? "PM" : "AM"} IST`;
  Logger.log(`📅 One-time trigger set for: ${label} on ${target.toUTCString()}`);
}

// ============================================================================
// HELPER: Delete all triggers created by this system (to avoid duplicates)
// ============================================================================

function deleteAllExactTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "sendReportAndReschedule") {
      ScriptApp.deleteTrigger(trigger);
      count++;
    }
  });
  Logger.log(`🗑️ Deleted ${count} existing exact trigger(s)`);
}

// ============================================================================
// HELPER: Check what triggers are currently scheduled
// ============================================================================

function listAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  Logger.log(`=== Active Triggers (${triggers.length}) ===`);
  triggers.forEach((t, i) => {
    Logger.log(`${i + 1}. Function: ${t.getHandlerFunction()}`);
  });
}