/**
 * Archive TODAY's team summaries at 11 PM to Yesterday_Archive sheet
 * Latest date appears at the top
 */
function archiveYesterdaySummaries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get or create Yesterday_Archive sheet
  let archiveSheet = ss.getSheetByName('Yesterday_Archive');
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet('Yesterday_Archive');
    
    // Set up headers - NOW WITH ALL COLUMNS
    const headers = [
      'Archive Date',
      'Team Name',
      'Agent Name',
      'Total Calls',
      'Inbound',
      'Outbound',
      'Dialer',
      'Answered',
      'WhatsApp Calls',
      'Avyukta Calls',
      'Ozonetel Calls',
      'Total Talktime',
      'WhatsApp Time',
      'Avyukta Time',
      'Ozonetel Time',
      'Avg Duration'
    ];
    archiveSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    archiveSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');
    archiveSheet.setFrozenRows(1);
  }
  
  // Get TODAY's date (the data we're archiving)
  const today = new Date();
  const archiveDate = Utilities.formatDate(today, 'Asia/Kolkata', 'yyyy-MM-dd');
  
  Logger.log(`Archiving data for: ${archiveDate}`);
  
  // Get team mapping to find all team sheets
  const teamMapping = loadTeamMappingFromSheet();
  if (!teamMapping) {
    Logger.log('No team mapping found');
    return;
  }
  
  // Get unique teams
  const teams = [...new Set(Object.values(teamMapping.teamMap))];
  teams.push('Unassigned'); // Include unassigned
  
  const rowsToArchive = [];
  
  // Loop through each team sheet
  teams.forEach((teamName, teamIndex) => {
    const teamSheet = ss.getSheetByName(teamName);
    
    if (!teamSheet) {
      Logger.log(`Team sheet not found: ${teamName}`);
      return;
    }
    
    // Find the summary section
    const summaryStartRow = 6; // First data row after headers
    let currentRow = summaryStartRow;
    
    // Read until we hit "TEAM TOTAL" row
    while (currentRow <= teamSheet.getLastRow()) {
      const agentName = teamSheet.getRange(currentRow, 1).getValue();
      
      if (!agentName || agentName === '') {
        break;
      }
      
      if (agentName.toString().includes('TEAM TOTAL')) {
        break;
      }
      
      // Extract the summary data (14 columns now - up to Avg Duration)
      const rowData = teamSheet.getRange(currentRow, 1, 1, 14).getValues()[0];
      
      // Clean agent name (remove emojis)
      let cleanAgentName = rowData[0].toString()
        .replace(/🔥/g, '')
        .replace(/🏆/g, '')
        .replace(/⚠️/g, '')
        .trim();
      
      // Build archive row: [Date, Team, Agent, ...stats]
      const archiveRow = [
        archiveDate,
        teamName,
        cleanAgentName,
        rowData[1],  // Total Calls
        rowData[2],  // Inbound
        rowData[3],  // Outbound
        rowData[4],  // Dialer
        rowData[5],  // Answered
        rowData[6],  // WhatsApp Calls
        rowData[7],  // Avyukta Calls
        rowData[8],  // Ozonetel Calls
        rowData[9],  // Total Talktime
        rowData[10], // WhatsApp Time
        rowData[11], // Avyukta Time
        rowData[12], // Ozonetel Time
        rowData[13]  // Avg Duration
      ];
      
      rowsToArchive.push(archiveRow);
      
      currentRow++;
    }
    
    // ⭐ ADD BLANK ROW AFTER EACH TEAM (except the last team)
    if (teamIndex < teams.length - 1) {
      rowsToArchive.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']); // Empty row with 16 columns
    }
  });
  
if (rowsToArchive.length > 0) {
    // ⭐ KEY CHANGE: Insert at row 2 (right after headers) instead of at the end
    const insertRow = 2;
    
    // Insert blank rows to make space
    archiveSheet.insertRowsAfter(1, rowsToArchive.length);
    
    // ⭐ FIX: Clear formatting from the newly inserted rows
    archiveSheet.getRange(insertRow, 1, rowsToArchive.length, 16)
      .setBackground('#ffffff')
      .setFontColor('#000000')
      .setFontWeight('normal');
    
    // Write data starting from row 2
    archiveSheet.getRange(insertRow, 1, rowsToArchive.length, 16).setValues(rowsToArchive);
    
    // Format the duration columns (columns 12, 13, 14, 15, 16 in archive sheet)
    archiveSheet.getRange(insertRow, 12, rowsToArchive.length, 5)
      .setNumberFormat('[hh]:mm:ss');
    
    // Format call count columns as numbers (columns 4-11)
    archiveSheet.getRange(insertRow, 4, rowsToArchive.length, 8)
      .setNumberFormat('0');
    
    // Center align data columns
    archiveSheet.getRange(insertRow, 4, rowsToArchive.length, 13)
      .setHorizontalAlignment('center');
    
    // ---------------------- APPLE-STYLE COLOR CODING ----------------------
    
    // Color coding for Total Calls (Column 4) and Total Talktime (Column 12)
    const totalCallsRange = archiveSheet.getRange(insertRow, 4, rowsToArchive.length, 1);
    const totalTalktimeRange = archiveSheet.getRange(insertRow, 12, rowsToArchive.length, 1);
    
    // Get the values
    const totalCallsValues = totalCallsRange.getValues();
    const totalTalktimeValues = totalTalktimeRange.getValues();
    
    // Function to map value to Apple-style color (green gradient)
    function getAppleGreenColor(value, maxValue) {
      if (value === '' || value === 0) return '#ffffff'; // blank or 0 = white
      const intensity = Math.min(Math.round((value / maxValue) * 200), 200); // 0-200
      return `rgb(${255 - intensity}, 255, ${255 - intensity})`; // Light green to green
    }
    
    // Find max for scaling
    const maxCalls = Math.max(...totalCallsValues.flat().map(v => Number(v) || 0));
    
    const maxTalktimeSeconds = Math.max(...totalTalktimeValues.flat().map(v => {
      if (!v) return 0;
      if (typeof v === 'number') return Math.round(v * 24 * 60 * 60);
      return 0;
    }));
    
    // Apply colors to Total Calls
    const callsColors = totalCallsValues.map(row => [getAppleGreenColor(Number(row[0] || 0), maxCalls)]);
    totalCallsRange.setBackgrounds(callsColors);
    
    // Apply colors to Total Talktime safely
    const talktimeColors = totalTalktimeValues.map(row => {
      let v = row[0];
      if (!v) return ['#ffffff'];
      
      let seconds = 0;
      
      if (typeof v === 'number') {
        // It's a time value in Google Sheets (fraction of a day)
        seconds = Math.round(v * 24 * 60 * 60);
      } else if (typeof v === 'string') {
        // String like "hh:mm:ss"
        const parts = v.split(':').map(Number);
        seconds = parts[0]*3600 + parts[1]*60 + parts[2];
      }
      
      return [getAppleGreenColor(seconds, maxTalktimeSeconds)];
    });
    totalTalktimeRange.setBackgrounds(talktimeColors);
    
    Logger.log(`✓ Archived ${rowsToArchive.length} agent summaries for ${archiveDate}`);
    
    // Add success message at the top (row 2 + number of rows)
    const messageRow = insertRow + rowsToArchive.length;
    archiveSheet.getRange(messageRow, 1, 1, 3).merge();
    archiveSheet.getRange(messageRow, 1).setValue(
      `✓ Archived ${rowsToArchive.length} records for ${archiveDate} at ${new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' })}`
    );
    archiveSheet.getRange(messageRow, 1)
      .setFontWeight('bold')
      .setFontColor('#0f9d58')
      .setBackground('#d9ead3');
    
    // Add a blank row for separation
    archiveSheet.insertRowAfter(messageRow);
}


 else {
    Logger.log('⚠ No data to archive');
  }
}