function fetchCallRecords() {
 // Configuration
 const AUTH_TOKEN = PropertiesService.getScriptProperties().getProperty('Token');
 const API_URL = 'https://api-smartflo.tatateleservices.com/v1/call/records';
  
 // Get date range from config
 const dateConfig = getDateRangeConfig();
 const fromDate = dateConfig.fromDate;
 const toDate = dateConfig.toDate;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 sheet.clear();
  let allRecords = [];
 let page = 1;
 let totalCount = 0;
  try {
   // Show progress
   sheet.getRange('A1').setValue('Fetching records...');
   SpreadsheetApp.flush();
  
   // Fetch all pages
   while (true) {
     const url = `${API_URL}?limit=1000&page=${page}&to_date=${encodeURIComponent(toDate)}&from_date=${encodeURIComponent(fromDate)}`;
    
     const options = {
       'method': 'GET',
       'headers': {
         'accept': 'application/json',
         'Authorization': AUTH_TOKEN
       },
       'muteHttpExceptions': true
     };
    
     Logger.log(`Fetching page ${page}...`);
     const response = UrlFetchApp.fetch(url, options);
     const responseCode = response.getResponseCode();
    
     if (responseCode !== 200) {
       throw new Error('API Error (HTTP ' + responseCode + '): ' + response.getContentText());
     }
    
     const data = JSON.parse(response.getContentText());
    
     // Store total count from first page
     if (page === 1) {
       totalCount = data.count || 0;
       Logger.log(`Total records available: ${totalCount}`);
     }
    
     // Get records from response
     const records = data.results || data.data || [];
    
     if (records.length === 0) {
       break; // No more records
     }
    
     allRecords = allRecords.concat(records);
     Logger.log(`Fetched ${allRecords.length} of ${totalCount} records`);
    
     // Check if we've got all records
     if (allRecords.length >= totalCount) {
       break;
     }
    
     page++;
    
     // Safety check to prevent infinite loops
     if (page > 100) {
       Logger.log('Safety limit reached (100 pages)');
       break;
     }
    
     // Small delay to avoid rate limiting
     Utilities.sleep(300);
   }
  
   // Clear the sheet again for actual data
   sheet.clear();
  
   if (allRecords.length === 0) {
     sheet.getRange('A1').setValue('No records found');
     return;
   }
  
   Logger.log(`Total records fetched: ${allRecords.length}`);
  
   // Extract headers from the first record
   const originalHeaders = Object.keys(allRecords[0]);
  
   // Add custom headers for cleaned agent names
   const headers = [...originalHeaders, 'agents_missed', 'agent_name_cleaned', 'agents_missed_cleaned'];
  
   // Write headers
   sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
   sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
   sheet.getRange(1, 1, 1, headers.length).setBackground('#4285f4').setFontColor('#ffffff');
  
   // Prepare data rows
   const rows = allRecords.map(record => {
     const rowData = originalHeaders.map(header => {
       const value = record[header];
       return value !== null && value !== undefined ? value : '';
     });
    
     // Extract agent name from missed_agents field
     const agentMissed = extractAgentName(record.missed_agents);
     rowData.push(agentMissed);
    
     // Add cleaned names
     const cleanedAgentName = cleanAgentName(record.agent_name);
     const cleanedAgentMissed = cleanAgentName(agentMissed);
    
     rowData.push(cleanedAgentName);
     rowData.push(cleanedAgentMissed);
    
     return rowData;
   });
  
   // Write data in batches
   const BATCH_SIZE = 1000;
   for (let i = 0; i < rows.length; i += BATCH_SIZE) {
     const batchRows = rows.slice(i, Math.min(i + BATCH_SIZE, rows.length));
     sheet.getRange(i + 2, 1, batchRows.length, headers.length).setValues(batchRows);
     Logger.log(`Written rows ${i + 1} to ${i + batchRows.length}`);
     SpreadsheetApp.flush();
   }
  
   // Freeze header row
   sheet.setFrozenRows(1);
  
   Logger.log(`SUCCESS: Imported ${allRecords.length} records!`);

   // ⭐ ADD THIS - Store last fetch time
PropertiesService.getScriptProperties().setProperty('LAST_FETCH_TIME', new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }));
PropertiesService.getScriptProperties().setProperty('LAST_FETCH_FROM_DATE', fromDate);
PropertiesService.getScriptProperties().setProperty('LAST_FETCH_TO_DATE', toDate);
// ⭐ END ADD
  
   // Add a completion indicator in the sheet
   const lastRow = sheet.getLastRow();
   sheet.getRange(lastRow + 2, 1).setValue(`✓ Successfully imported ${allRecords.length} records on ${new Date().toLocaleString()}`);
   sheet.getRange(lastRow + 2, 1).setFontWeight('bold').setFontColor('#0f9d58');
  
 } catch (error) {
   Logger.log('Error: ' + error.toString());
   sheet.clear();
   sheet.getRange('A1').setValue('Error occurred:');
   sheet.getRange('A2').setValue(error.toString());
 }
}


function getDateRangeConfig() {
  // Check if custom dates are set in script properties
  const customFrom = PropertiesService.getScriptProperties().getProperty('CUSTOM_FROM_DATE');
  const customTo = PropertiesService.getScriptProperties().getProperty('CUSTOM_TO_DATE');
  
  if (customFrom && customTo) {
    Logger.log('Using custom date range: ' + customFrom + ' to ' + customTo);
    return {
      fromDate: customFrom,
      toDate: customTo
    };
  }
  
  // Default to today
  Logger.log('Using today\'s date range');
  const today = new Date();
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, '0');
  const day = String(today.getDate()).padStart(2, '0');
  const dateString = `${year}-${month}-${day}`;
  
return {
   fromDate: `${dateString} 00:00:00`,
   toDate: `${dateString} 23:59:59`
 };

}


/**
* Extract agent name from missed_agents field
*/
function extractAgentName(missedAgents) {
 if (!missedAgents) {
   return '';
 }
  try {
   if (typeof missedAgents === 'object' && missedAgents.name) {
     return missedAgents.name;
   }
  
   if (typeof missedAgents === 'string') {
     const nameMatch = missedAgents.match(/name=([^,}]+)/);
     if (nameMatch && nameMatch[1]) {
       return nameMatch[1].trim();
     }
   }
  
   if (Array.isArray(missedAgents) && missedAgents.length > 0) {
     return extractAgentName(missedAgents[0]);
   }
  
   return '';
 } catch (e) {
   Logger.log('Error extracting agent name: ' + e.toString());
   return '';
 }
}


/**
* Clean agent name by removing extensions, TTBS, and other suffixes
* This standardizes names so the same person is counted together
*/
function cleanAgentName(name) {
 if (!name || name === '') {
   return '';
 }
  let cleaned = name.toString().trim();
  // Remove common patterns:
 // - Extension suffix (e.g., "Name-Extension", "Name - Extension")
 cleaned = cleaned.replace(/[-\s]*Extension$/i, '');
  // - TTBS suffix (e.g., "Name-TTBS", "Name - TTBS")

  // - Any hyphen followed by digits (e.g., "Name-123")
 cleaned = cleaned.replace(/[-\s]*\d+$/, '');
  // - Parenthetical information (e.g., "Name (Extension)")
 cleaned = cleaned.replace(/\s*\([^)]*\)$/, '');
  // - Trailing hyphens or spaces
 cleaned = cleaned.replace(/[-\s]+$/, '');
  // Convert to title case for consistency
 cleaned = cleaned.split(' ')
   .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
   .join(' ');
  return cleaned;
}


/**
* Convert seconds to HH:MM:SS format
*/
function secondsToHMS(seconds) {
  if (!seconds || seconds < 0) {
    return '';
  }
  
  const hrs = Math.floor(seconds / 3600);
  const mins = Math.floor((seconds % 3600) / 60);
  const secs = Math.floor(seconds % 60);
  
  return `${String(hrs).padStart(2, '0')}:${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
}



/**
* Get time bucket index from time value
* Returns index 0-23 for the hour bucket, or -1 if invalid
*/
function getTimeBucket(timeValue) {
  if (!timeValue) {
    return -1;
  }
  
  try {
    let hour = -1;
    
    // Handle Date objects - MOST COMMON CASE
    if (timeValue instanceof Date) {
      hour = timeValue.getHours();
      return (hour >= 0 && hour < 24) ? hour : -1;
    }
    
    // Handle object that might be a date
    if (typeof timeValue === 'object' && timeValue.getHours) {
      hour = timeValue.getHours();
      return (hour >= 0 && hour < 24) ? hour : -1;
    }
    
    // Handle string values
    if (typeof timeValue === 'string') {
      const timeStr = timeValue.trim();
      
      // Extract time from "YYYY-MM-DD HH:MM:SS" format
      const dateTimeMatch = timeStr.match(/(\d{4})-(\d{2})-(\d{2})\s+(\d{1,2}):(\d{2}):(\d{2})/);
      if (dateTimeMatch) {
        hour = parseInt(dateTimeMatch[4], 10);
        return (hour >= 0 && hour < 24) ? hour : -1;
      }
      
      // Format: HH:MM:SS or HH:MM (24-hour)
      const match24 = timeStr.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?/);
      if (match24) {
        hour = parseInt(match24[1], 10);
        return (hour >= 0 && hour < 24) ? hour : -1;
      }
      
      // Format: Time part after space
      if (timeStr.includes(' ')) {
        const parts = timeStr.split(' ');
        for (let part of parts) {
          const match = part.match(/^(\d{1,2}):(\d{2})/);
          if (match) {
            hour = parseInt(match[1], 10);
            return (hour >= 0 && hour < 24) ? hour : -1;
          }
        }
      }
    }
    
    // Handle number values (Excel serial numbers)
    if (typeof timeValue === 'number') {
      // Just the hour (0-23)
      if (timeValue >= 0 && timeValue < 24) {
        hour = Math.floor(timeValue);
      }
      // Fraction of a day (0 to 1)
      else if (timeValue > 0 && timeValue < 1) {
        hour = Math.floor(timeValue * 24);
      }
      // Excel serial date (days since 1900) - extract time portion
      else if (timeValue > 1) {
        const fraction = timeValue - Math.floor(timeValue);
        hour = Math.floor(fraction * 24);
      }
      
      return (hour >= 0 && hour < 24) ? hour : -1;
    }
    
    return -1;
    
  } catch (e) {
    Logger.log('Error parsing time "' + timeValue + '": ' + e.toString());
    return -1;
  }
}

/**
* TEAM MAPPING SHEET CONFIGURATION
* Configure the sheet name and column names where team mapping is stored
*/
function getTeamMappingConfig() {
 return {
   sheetName: 'Team Mapping',  // Name of the sheet with agent-team mapping
   agentColumn: 'Agent Name',  // Column header for agent names
   teamColumn: 'Team'          // Column header for team names
 };
}


/**
* Load team mapping from a sheet
* Expected format: Sheet should have columns for Agent Name and Team
*/
/**
* Load team mapping from a sheet
* Expected format: Sheet should have columns for Agent Name, Team, and optionally Canonical Name
*/
function loadTeamMappingFromSheet() {
 const config = getTeamMappingConfig();
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const teamSheet = ss.getSheetByName(config.sheetName);
  
 if (!teamSheet) {
   Logger.log(`Team mapping sheet "${config.sheetName}" not found. Please create it.`);
   return null;
 }
  
 const lastRow = teamSheet.getLastRow();
 const lastCol = teamSheet.getLastColumn();
  
 if (lastRow < 2) {
   Logger.log('Team mapping sheet is empty.');
   return null;
 }
  
 const data = teamSheet.getRange(1, 1, lastRow, lastCol).getValues();
 const headers = data[0];
  
 // Find column indices
 const agentColIdx = headers.findIndex(h => h && h.toString().toLowerCase() === config.agentColumn.toLowerCase());
 const teamColIdx = headers.findIndex(h => h && h.toString().toLowerCase() === config.teamColumn.toLowerCase());
 const canonicalColIdx = headers.findIndex(h => h && h.toString().toLowerCase() === 'canonical name');
  
 if (agentColIdx === -1 || teamColIdx === -1) {
   Logger.log(`Could not find "${config.agentColumn}" or "${config.teamColumn}" columns in the team mapping sheet.`);
   Logger.log(`Found headers: ${headers.join(', ')}`);
   return null;
 }
  
 // Build mapping object with both team and canonical name
 const mapping = {
   teamMap: {},
   canonicalMap: {}
 };
 
for (let i = 1; i < data.length; i++) {
  const agentName = data[i][agentColIdx];
  const teamName = data[i][teamColIdx];
  const canonicalName = canonicalColIdx !== -1 ? data[i][canonicalColIdx] : null;

  if (agentName && teamName) {
    const originalAgent = agentName.toString().trim();
    const cleanedAgent = cleanAgentName(originalAgent);

    // Canonical name: explicit > fallback
    const canonical =
      canonicalName && canonicalName.toString().trim() !== ''
        ? canonicalName.toString().trim()
        : cleanedAgent;

    // 🔒 Identity mapping (ONLY original agent name)
    mapping.teamMap[originalAgent] = teamName.toString().trim();
    mapping.canonicalMap[originalAgent] = canonical;

    // 🧭 Optional lookup helper (NOT identity)
    mapping.lookupMap = mapping.lookupMap || {};
    mapping.lookupMap[cleanedAgent] = originalAgent;
  }
}

  
 Logger.log(`Loaded ${Object.keys(mapping.teamMap).length} agent-team mappings from sheet.`);
 return mapping;
}



function getCanonicalName(agentName, teamMapping) {
  if (!teamMapping || !teamMapping.canonicalMap) return agentName;
  if (!agentName) return agentName;
  // 1. Exact match
  if (teamMapping.canonicalMap[agentName]) {
    return teamMapping.canonicalMap[agentName];
  }
  // 2. Case-insensitive match
  const agentLower = agentName.toLowerCase();
  for (const [mappedAgent, canonical] of Object.entries(teamMapping.canonicalMap)) {
    if (mappedAgent.toLowerCase() === agentLower) return canonical;
  }
  // 3. ✅ NEW: Try via lookupMap (cleaned name → original → canonical)
  if (teamMapping.lookupMap) {
    const cleanedInput = cleanAgentName(agentName);
    const originalKey = teamMapping.lookupMap[cleanedInput];
    if (originalKey && teamMapping.canonicalMap[originalKey]) {
      return teamMapping.canonicalMap[originalKey];
    }
  }
  return agentName;
}
function getAgentTeam(agentName, teamMapping) {
  if (!teamMapping || !teamMapping.teamMap) return 'Unassigned';
  if (!agentName) return 'Unassigned';
  // 1. Exact match
  if (teamMapping.teamMap[agentName]) return teamMapping.teamMap[agentName];
  // 2. Case-insensitive match
  const agentLower = agentName.toLowerCase();
  for (const [mappedAgent, team] of Object.entries(teamMapping.teamMap)) {
    if (mappedAgent.toLowerCase() === agentLower) return team;
  }
  // 3. ✅ NEW: Try via lookupMap (cleaned name → original → team)
  if (teamMapping.lookupMap) {
    const cleanedInput = cleanAgentName(agentName);
    const originalKey = teamMapping.lookupMap[cleanedInput];
    if (originalKey && teamMapping.teamMap[originalKey]) {
      return teamMapping.teamMap[originalKey];
    }
  }
  return 'Unassigned';
}







/**
 * Helper function to convert HH:MM:SS to seconds
 */
function hmsToSeconds(hms) {
  if (!hms || hms === '00:00:00') return 0;
  
  const parts = hms.split(':');
  if (parts.length !== 3) return 0;
  
  const hours = parseInt(parts[0], 10) || 0;
  const minutes = parseInt(parts[1], 10) || 0;
  const seconds = parseInt(parts[2], 10) || 0;
  
  return hours * 3600 + minutes * 60 + seconds;
}


/**
 * Process WhatsApp call data and add to team data structure
 * Only includes calls within the configured date range
 */
function processWhatsAppData(teamData, timeBuckets, teamMapping) {
  Logger.log('=== Starting WhatsApp Data Processing ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const whatsappSheet = ss.getSheetByName('Whatsapp Call');
  
  if (!whatsappSheet) {
    Logger.log('ERROR: Whatsapp Call sheet not found!');
    return;
  }
  
  const lastRow = whatsappSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No data in Whatsapp Call sheet');
    return;
  }
  
  const headers = whatsappSheet.getRange(1, 1, 1, whatsappSheet.getLastColumn()).getValues()[0];
  const data = whatsappSheet.getRange(2, 1, lastRow - 1, whatsappSheet.getLastColumn()).getValues();
  
  // Find column indices
  const timestampIdx = headers.findIndex(h => h && h.toString().toLowerCase() === 'timestamp');
  const agentNameIdx = headers.findIndex(h => h && h.toString().toLowerCase().includes('agent') && h.toString().toLowerCase().includes('name'));
  const durationIdx = headers.findIndex(h => h && h.toString().toLowerCase().includes('total duration'));
  const timeBucketIdx = headers.findIndex(h => h && h.toString().toLowerCase().includes('time bucket'));
  
  Logger.log('Column indices - Timestamp: ' + timestampIdx + ', Agent: ' + agentNameIdx + ', Duration: ' + durationIdx);
  
  if (agentNameIdx === -1 || timestampIdx === -1) {
    Logger.log('ERROR: Required columns not found!');
    return;
  }
  
  // Get date range filter
  const dateConfig = getDateRangeConfig();
  const fromDate = new Date(dateConfig.fromDate);
  const toDate = new Date(dateConfig.toDate);
  
  // Normalize to start/end of day for comparison
  fromDate.setHours(0, 0, 0, 0);
  toDate.setHours(23, 59, 59, 999);
  
  Logger.log('Date filter: ' + fromDate.toDateString() + ' to ' + toDate.toDateString());
  
  let processedCount = 0;
  let totalDuration = 0;
  let skippedReasons = {
    noAgent: 0,
    noTimestamp: 0,
    outOfRange: 0,
    invalidBucket: 0
  };
  
  const hourToBucket = {
    8: '08-09 AM', 9: '09-10 AM', 10: '10-11 AM', 11: '11-12 PM',
    12: '12-01 PM', 13: '01-02 PM', 14: '02-03 PM', 15: '03-04 PM',
    16: '04-05 PM', 17: '05-06 PM', 18: '06-07 PM', 19: '07-08 PM'
  };
  
  data.forEach((row, index) => {
    const agentName = row[agentNameIdx];
    const timestampValue = row[timestampIdx];
    
    if (!agentName) {
      skippedReasons.noAgent++;
      return;
    }
    
    if (!timestampValue) {
      skippedReasons.noTimestamp++;
      return;
    }
    
    // Parse timestamp
    let timestamp;
    if (timestampValue instanceof Date) {
      timestamp = new Date(timestampValue.getTime());
    } else if (typeof timestampValue === 'string') {
      timestamp = new Date(timestampValue);
    } else if (typeof timestampValue === 'number') {
      timestamp = new Date((timestampValue - 25569) * 86400 * 1000);
    } else {
      skippedReasons.noTimestamp++;
      return;
    }
    
    if (isNaN(timestamp.getTime())) {
      skippedReasons.noTimestamp++;
      return;
    }
    
    // Create a date-only version for comparison
    const timestampDateOnly = new Date(timestamp);
    timestampDateOnly.setHours(0, 0, 0, 0);
    
    // Check if date is within range
    if (timestampDateOnly < fromDate || timestampDateOnly > toDate) {
      skippedReasons.outOfRange++;
      return;
    }
    
    // Get hour and bucket
    const hour = timestamp.getHours();
    const bucket = hourToBucket[hour];
    
    // Skip if invalid bucket
    if (!bucket || !timeBuckets.includes(bucket)) {
      skippedReasons.invalidBucket++;
      return;
    }
    
    // Get canonical name and team
    const canonicalAgent = getCanonicalName(agentName, teamMapping);
    const team = getAgentTeam(agentName, teamMapping);
    
    // Initialize team and agent if needed
    if (!teamData[team]) {
      teamData[team] = {};
    }
    
    if (!teamData[team][canonicalAgent]) {
      teamData[team][canonicalAgent] = {};
      timeBuckets.forEach(b => {
        teamData[team][canonicalAgent][b] = 0;
      });
    }
    
    // Get duration - UPDATED PARSING
    let duration = 0;
    if (durationIdx !== -1 && row[durationIdx]) {
      const durationValue = row[durationIdx];
      
      if (durationValue instanceof Date) {
        // Duration is stored as Date object - extract time portion
        const hours = durationValue.getHours();
        const minutes = durationValue.getMinutes();
        const seconds = durationValue.getSeconds();
        duration = hours * 3600 + minutes * 60 + seconds;
      } else if (typeof durationValue === 'string' && durationValue.includes(':')) {
        // Duration is in HH:MM:SS string format
        duration = hmsToSeconds(durationValue);
      } else if (typeof durationValue === 'number') {
        // Duration is a number (could be seconds or Excel decimal days)
        if (durationValue < 1) {
          // Likely a fraction of a day (Excel format)
          duration = Math.round(durationValue * 24 * 60 * 60);
        } else {
          // Likely already in seconds
          duration = parseFloat(durationValue);
        }
      }
    }
    
    // Add to the bucket
    teamData[team][canonicalAgent][bucket] += duration;
    totalDuration += duration;
    processedCount++;
    
    if (processedCount <= 5) {
      Logger.log(`Processed: ${canonicalAgent} (${team}) - ${bucket} - ${duration}s`);
    }
  });
  
  Logger.log('=== WhatsApp Processing Summary ===');
  Logger.log('Total records: ' + data.length);
  Logger.log('Processed: ' + processedCount);
  Logger.log('Total duration: ' + totalDuration + ' seconds (' + secondsToHMS(totalDuration) + ')');
  Logger.log('Skipped - No agent: ' + skippedReasons.noAgent);
  Logger.log('Skipped - No timestamp: ' + skippedReasons.noTimestamp);
  Logger.log('Skipped - Out of range: ' + skippedReasons.outOfRange);
  Logger.log('Skipped - Invalid bucket: ' + skippedReasons.invalidBucket);
  Logger.log('====================================');
}







/**
 * Process WhatsApp call data for call counts (not duration)
 * Only includes calls within the configured date range
 */
function processWhatsAppDataForCalls(teamData, timeBuckets, teamMapping) {
  Logger.log('=== Starting WhatsApp Call Count Processing ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const whatsappSheet = ss.getSheetByName('Whatsapp Call');
  
  if (!whatsappSheet) {
    Logger.log('ERROR: Whatsapp Call sheet not found!');
    return;
  }
  
  const lastRow = whatsappSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No data in Whatsapp Call sheet');
    return;
  }
  
  const headers = whatsappSheet.getRange(1, 1, 1, whatsappSheet.getLastColumn()).getValues()[0];
  const data = whatsappSheet.getRange(2, 1, lastRow - 1, whatsappSheet.getLastColumn()).getValues();
  
  // Find column indices
  const timestampIdx = headers.findIndex(h => h && h.toString().toLowerCase() === 'timestamp');
  const agentNameIdx = headers.findIndex(h => h && h.toString().toLowerCase().includes('agent') && h.toString().toLowerCase().includes('name'));
  const timeBucketIdx = headers.findIndex(h => h && h.toString().toLowerCase().includes('time bucket'));
  
  if (agentNameIdx === -1 || timestampIdx === -1) {
    Logger.log('ERROR: Required columns not found!');
    return;
  }
  
  // Get date range filter
  const dateConfig = getDateRangeConfig();
  const fromDate = new Date(dateConfig.fromDate);
  const toDate = new Date(dateConfig.toDate);
  
  // Normalize to start/end of day
  fromDate.setHours(0, 0, 0, 0);
  toDate.setHours(23, 59, 59, 999);
  
  Logger.log('Date filter: ' + fromDate.toDateString() + ' to ' + toDate.toDateString());
  
  let processedCount = 0;
  
  const hourToBucket = {
    8: '08-09 AM', 9: '09-10 AM', 10: '10-11 AM', 11: '11-12 PM',
    12: '12-01 PM', 13: '01-02 PM', 14: '02-03 PM', 15: '03-04 PM',
    16: '04-05 PM', 17: '05-06 PM', 18: '06-07 PM', 19: '07-08 PM'
  };
  
  data.forEach((row, index) => {
    const agentName = row[agentNameIdx];
    const timestampValue = row[timestampIdx];
    
    if (!agentName || !timestampValue) return;
    
    // Parse timestamp
    let timestamp;
    if (timestampValue instanceof Date) {
      timestamp = new Date(timestampValue.getTime());
    } else if (typeof timestampValue === 'string') {
      timestamp = new Date(timestampValue);
    } else if (typeof timestampValue === 'number') {
      timestamp = new Date((timestampValue - 25569) * 86400 * 1000);
    } else {
      return;
    }
    
    if (isNaN(timestamp.getTime())) return;
    
    // Create date-only version for comparison
    const timestampDateOnly = new Date(timestamp);
    timestampDateOnly.setHours(0, 0, 0, 0);
    
    // Check if date is within range
    if (timestampDateOnly < fromDate || timestampDateOnly > toDate) {
      return;
    }
    
    // Get hour and bucket
    const hour = timestamp.getHours();
    const bucket = hourToBucket[hour];
    
    if (!bucket || !timeBuckets.includes(bucket)) return;
    
    // Get canonical name and team
    const canonicalAgent = getCanonicalName(agentName, teamMapping);
    const team = getAgentTeam(agentName, teamMapping);
    
    // Initialize if needed
    if (!teamData[team]) {
      teamData[team] = {};
    }
    
    if (!teamData[team][canonicalAgent]) {
      teamData[team][canonicalAgent] = {};
      timeBuckets.forEach(b => {
        teamData[team][canonicalAgent][b] = 0;
      });
    }
    
    // Increment call count
    teamData[team][canonicalAgent][bucket] += 1;
    processedCount++;
  });
  
  Logger.log('WhatsApp call count processed: ' + processedCount);
  Logger.log('====================================');
}






/**
 * Process WhatsApp data for team sheets with hourly buckets
 */
function processWhatsAppDataForTeamSheets(teamData, timeBuckets, teamMapping) {
  Logger.log('=== Processing WhatsApp Data for Team Sheets ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const whatsappSheet = ss.getSheetByName('Whatsapp Call');
  
  if (!whatsappSheet) {
    Logger.log('Whatsapp Call sheet not found');
    return;
  }
  
  const lastRow = whatsappSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No data in Whatsapp Call sheet');
    return;
  }
  
  const headers = whatsappSheet.getRange(1, 1, 1, whatsappSheet.getLastColumn()).getValues()[0];
  const data = whatsappSheet.getRange(2, 1, lastRow - 1, whatsappSheet.getLastColumn()).getValues();
  
  // Find column indices
  const timestampIdx = headers.findIndex(h => h && h.toString().toLowerCase() === 'timestamp');
  const agentNameIdx = headers.findIndex(h => h && h.toString().toLowerCase().includes('agent') && h.toString().toLowerCase().includes('name'));
  const durationIdx = headers.findIndex(h => h && h.toString().toLowerCase().includes('total duration'));
  
  if (agentNameIdx === -1 || timestampIdx === -1) {
    Logger.log('Required columns not found in WhatsApp sheet');
    return;
  }
  
  // Get date range filter
  const dateConfig = getDateRangeConfig();
  const fromDate = new Date(dateConfig.fromDate);
  const toDate = new Date(dateConfig.toDate);
  fromDate.setHours(0, 0, 0, 0);
  toDate.setHours(23, 59, 59, 999);
  
  // Map hours to bucket labels
  const hourToBucket = {
    8: '08-09 AM', 9: '09-10 AM', 10: '10-11 AM', 11: '11-12 PM',
    12: '12-01 PM', 13: '01-02 PM', 14: '02-03 PM', 15: '03-04 PM',
    16: '04-05 PM', 17: '05-06 PM', 18: '06-07 PM', 19: '07-08 PM',
    20: '08-09 PM', 21: '09-10 PM'
  };
  
  let processedCount = 0;
  
  data.forEach(row => {
    const agentName = row[agentNameIdx];
    const timestampValue = row[timestampIdx];
    
    if (!agentName || !timestampValue) return;
    
    // Parse timestamp
    let timestamp;
    if (timestampValue instanceof Date) {
      timestamp = new Date(timestampValue.getTime());
    } else if (typeof timestampValue === 'string') {
      timestamp = new Date(timestampValue);
    } else if (typeof timestampValue === 'number') {
      timestamp = new Date((timestampValue - 25569) * 86400 * 1000);
    } else {
      return;
    }
    
    if (isNaN(timestamp.getTime())) return;
    
    // Check date range
    const timestampDateOnly = new Date(timestamp);
    timestampDateOnly.setHours(0, 0, 0, 0);
    
    if (timestampDateOnly < fromDate || timestampDateOnly > toDate) {
      return;
    }
    
    // Get hour and bucket
    const hour = timestamp.getHours();
    const bucket = hourToBucket[hour];
    
    // Get canonical name and team
    const canonicalAgent = getCanonicalName(agentName, teamMapping);
    const team = getAgentTeam(agentName, teamMapping);
    
    // Initialize if needed
    if (!teamData[team]) {
      teamData[team] = {};
    }
    
if (!teamData[team][canonicalAgent]) {
  teamData[team][canonicalAgent] = {
    totalCalls: 0,
    totalDuration: 0,
    inboundCalls: 0,
    outboundCalls: 0,
    dialerCalls: 0,
    answeredCalls: 0,
    whatsappCalls: 0,
    whatsappDuration: 0,
    avyuktaCalls: 0,
    avyuktaDuration: 0,
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
    
    // Get duration
    let duration = 0;
    if (durationIdx !== -1 && row[durationIdx]) {
      const durationValue = row[durationIdx];
      
      if (durationValue instanceof Date) {
        const hours = durationValue.getHours();
        const minutes = durationValue.getMinutes();
        const seconds = durationValue.getSeconds();
        duration = hours * 3600 + minutes * 60 + seconds;
      } else if (typeof durationValue === 'string' && durationValue.includes(':')) {
        duration = hmsToSeconds(durationValue);
      } else if (typeof durationValue === 'number') {
        if (durationValue < 1) {
          duration = Math.round(durationValue * 24 * 60 * 60);
        } else {
          duration = parseFloat(durationValue);
        }
      }
    }
    
  // Add to team data totals
// Add to team data totals (combined)
teamData[team][canonicalAgent].whatsappCalls += 1;           // Track WhatsApp separately
teamData[team][canonicalAgent].whatsappDuration += duration; // Track WhatsApp separately

    // Add to bucket if valid
    if (bucket && timeBuckets.includes(bucket)) {
      teamData[team][canonicalAgent].buckets[bucket].calls += 1;
      teamData[team][canonicalAgent].buckets[bucket].duration += duration;
    }
    
    processedCount++;
  });
  
  Logger.log(`WhatsApp data processed: ${processedCount} calls`);
}


/**
 * Process Avyukta call data and add to team data structure
 * Only includes calls within the configured date range
 */
function processAvyuktaData(teamData, timeBuckets, teamMapping) {
  Logger.log('=== Starting Avyukta Data Processing ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const avyuktaSheet = ss.getSheetByName('Avyukta');
  
  if (!avyuktaSheet) {
    Logger.log('ERROR: Avyukta sheet not found!');
    return;
  }
  
  const lastRow = avyuktaSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No data in Avyukta sheet');
    return;
  }
  
  const headers = avyuktaSheet.getRange(1, 1, 1, avyuktaSheet.getLastColumn()).getValues()[0];
  const data = avyuktaSheet.getRange(2, 1, lastRow - 1, avyuktaSheet.getLastColumn()).getValues();
  
  // Find column indices
  const callDateIdx = headers.findIndex(h => h && h.toString().toLowerCase() === 'call_date');
  const fullNameIdx = headers.findIndex(h => h && h.toString().toLowerCase() === 'full_name');
  const lengthInSecIdx = headers.findIndex(h => h && h.toString().toLowerCase() === 'length_in_sec');
  
  Logger.log('Column indices - Call Date: ' + callDateIdx + ', Full Name: ' + fullNameIdx + ', Length: ' + lengthInSecIdx);
  
  if (callDateIdx === -1 || fullNameIdx === -1 || lengthInSecIdx === -1) {
    Logger.log('ERROR: Required columns not found!');
    return;
  }
  
  // Get date range filter
  const dateConfig = getDateRangeConfig();
  const fromDate = new Date(dateConfig.fromDate);
  const toDate = new Date(dateConfig.toDate);
  
  // Normalize to start/end of day for comparison
  fromDate.setHours(0, 0, 0, 0);
  toDate.setHours(23, 59, 59, 999);
  
  Logger.log('Date filter: ' + fromDate.toDateString() + ' to ' + toDate.toDateString());
  
  let processedCount = 0;
  let totalDuration = 0;
  let skippedReasons = {
    noAgent: 0,
    noCallDate: 0,
    outOfRange: 0,
    invalidBucket: 0,
    zeroDuration: 0
  };
  
  const hourToBucket = {
    8: '08-09 AM', 9: '09-10 AM', 10: '10-11 AM', 11: '11-12 PM',
    12: '12-01 PM', 13: '01-02 PM', 14: '02-03 PM', 15: '03-04 PM',
    16: '04-05 PM', 17: '05-06 PM', 18: '06-07 PM', 19: '07-08 PM',
    20: '08-09 PM', 21: '09-10 PM'
  };
  
  data.forEach((row, index) => {
    const agentName = row[fullNameIdx];
    const callDateValue = row[callDateIdx];
    const duration = row[lengthInSecIdx];
    
    // Skip if no agent name
    if (!agentName) {
      skippedReasons.noAgent++;
      return;
    }
    
    // Skip if no call date
    if (!callDateValue) {
      skippedReasons.noCallDate++;
      return;
    }
    
    // Skip if duration is 0 or invalid
    if (!duration || duration <= 0) {
      skippedReasons.zeroDuration++;
      return;
    }
    
    // Parse call_date (format: 2026-02-04 10:05:06)
    let callDate;
    if (callDateValue instanceof Date) {
      callDate = new Date(callDateValue.getTime());
    } else if (typeof callDateValue === 'string') {
      callDate = new Date(callDateValue);
    } else if (typeof callDateValue === 'number') {
      callDate = new Date((callDateValue - 25569) * 86400 * 1000);
    } else {
      skippedReasons.noCallDate++;
      return;
    }
    
    if (isNaN(callDate.getTime())) {
      skippedReasons.noCallDate++;
      return;
    }
    
    // Create a date-only version for comparison
    const callDateOnly = new Date(callDate);
    callDateOnly.setHours(0, 0, 0, 0);
    
    // Check if date is within range
    if (callDateOnly < fromDate || callDateOnly > toDate) {
      skippedReasons.outOfRange++;
      return;
    }
    
    // Get hour and bucket
    const hour = callDate.getHours();
    const bucket = hourToBucket[hour];
    
    // Skip if invalid bucket
    if (!bucket || !timeBuckets.includes(bucket)) {
      skippedReasons.invalidBucket++;
      return;
    }
    
    // Get canonical name and team
    const canonicalAgent = getCanonicalName(agentName, teamMapping);
    const team = getAgentTeam(agentName, teamMapping);
    
    // Initialize team and agent if needed
    if (!teamData[team]) {
      teamData[team] = {};
    }
    
    if (!teamData[team][canonicalAgent]) {
      teamData[team][canonicalAgent] = {};
      timeBuckets.forEach(b => {
        teamData[team][canonicalAgent][b] = 0;
      });
    }
    
    // Add duration to the bucket
    teamData[team][canonicalAgent][bucket] += duration;
    totalDuration += duration;
    processedCount++;
    
    if (processedCount <= 5) {
      Logger.log(`Processed: ${canonicalAgent} (${team}) - ${bucket} - ${duration}s at ${callDate}`);
    }
  });
  
  Logger.log('=== Avyukta Processing Summary ===');
  Logger.log('Total records: ' + data.length);
  Logger.log('Processed: ' + processedCount);
  Logger.log('Total duration: ' + totalDuration + ' seconds (' + secondsToHMS(totalDuration) + ')');
  Logger.log('Skipped - No agent: ' + skippedReasons.noAgent);
  Logger.log('Skipped - No call date: ' + skippedReasons.noCallDate);
  Logger.log('Skipped - Out of range: ' + skippedReasons.outOfRange);
  Logger.log('Skipped - Invalid bucket: ' + skippedReasons.invalidBucket);
  Logger.log('Skipped - Zero duration: ' + skippedReasons.zeroDuration);
  Logger.log('====================================');
}


/**
 * Process Avyukta call data for call counts (not duration)
 * Only includes calls within the configured date range
 */
function processAvyuktaDataForCalls(teamData, timeBuckets, teamMapping) {
  Logger.log('=== Starting Avyukta Call Count Processing ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const avyuktaSheet = ss.getSheetByName('Avyukta');
  
  if (!avyuktaSheet) {
    Logger.log('ERROR: Avyukta sheet not found!');
    return;
  }
  
  const lastRow = avyuktaSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No data in Avyukta sheet');
    return;
  }
  
  const headers = avyuktaSheet.getRange(1, 1, 1, avyuktaSheet.getLastColumn()).getValues()[0];
  const data = avyuktaSheet.getRange(2, 1, lastRow - 1, avyuktaSheet.getLastColumn()).getValues();
  
  // Find column indices
  const callDateIdx = headers.findIndex(h => h && h.toString().toLowerCase() === 'call_date');
  const fullNameIdx = headers.findIndex(h => h && h.toString().toLowerCase() === 'full_name');
  
  if (callDateIdx === -1 || fullNameIdx === -1) {
    Logger.log('ERROR: Required columns not found!');
    return;
  }
  
  // Get date range filter
  const dateConfig = getDateRangeConfig();
  const fromDate = new Date(dateConfig.fromDate);
  const toDate = new Date(dateConfig.toDate);
  
  // Normalize to start/end of day
  fromDate.setHours(0, 0, 0, 0);
  toDate.setHours(23, 59, 59, 999);
  
  Logger.log('Date filter: ' + fromDate.toDateString() + ' to ' + toDate.toDateString());
  
  let processedCount = 0;
  
  const hourToBucket = {
    8: '08-09 AM', 9: '09-10 AM', 10: '10-11 AM', 11: '11-12 PM',
    12: '12-01 PM', 13: '01-02 PM', 14: '02-03 PM', 15: '03-04 PM',
    16: '04-05 PM', 17: '05-06 PM', 18: '06-07 PM', 19: '07-08 PM',
    20: '08-09 PM', 21: '09-10 PM'
  };
  
  data.forEach((row, index) => {
    const agentName = row[fullNameIdx];
    const callDateValue = row[callDateIdx];
    
    if (!agentName || !callDateValue) return;
    
    // Parse call date
    let callDate;
    if (callDateValue instanceof Date) {
      callDate = new Date(callDateValue.getTime());
    } else if (typeof callDateValue === 'string') {
      callDate = new Date(callDateValue);
    } else if (typeof callDateValue === 'number') {
      callDate = new Date((callDateValue - 25569) * 86400 * 1000);
    } else {
      return;
    }
    
    if (isNaN(callDate.getTime())) return;
    
    // Create date-only version for comparison
    const callDateOnly = new Date(callDate);
    callDateOnly.setHours(0, 0, 0, 0);
    
    // Check if date is within range
    if (callDateOnly < fromDate || callDateOnly > toDate) {
      return;
    }
    
    // Get hour and bucket
    const hour = callDate.getHours();
    const bucket = hourToBucket[hour];
    
    if (!bucket || !timeBuckets.includes(bucket)) return;
    
    // Get canonical name and team
    const canonicalAgent = getCanonicalName(agentName, teamMapping);
    const team = getAgentTeam(agentName, teamMapping);
    
    // Initialize if needed
    if (!teamData[team]) {
      teamData[team] = {};
    }
    
    if (!teamData[team][canonicalAgent]) {
      teamData[team][canonicalAgent] = {};
      timeBuckets.forEach(b => {
        teamData[team][canonicalAgent][b] = 0;
      });
    }
    
    // Increment call count
    teamData[team][canonicalAgent][bucket] += 1;
    processedCount++;
  });
  
  Logger.log('Avyukta call count processed: ' + processedCount);
  Logger.log('====================================');
}


/**
 * Process Avyukta data for team sheets with hourly buckets
 */
function processAvyuktaDataForTeamSheets(teamData, timeBuckets, teamMapping) {
  Logger.log('=== Processing Avyukta Data for Team Sheets ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const avyuktaSheet = ss.getSheetByName('Avyukta');
  
  if (!avyuktaSheet) {
    Logger.log('Avyukta sheet not found');
    return;
  }
  
  const lastRow = avyuktaSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No data in Avyukta sheet');
    return;
  }
  
  const headers = avyuktaSheet.getRange(1, 1, 1, avyuktaSheet.getLastColumn()).getValues()[0];
  const data = avyuktaSheet.getRange(2, 1, lastRow - 1, avyuktaSheet.getLastColumn()).getValues();
  
  // Find column indices
  const callDateIdx = headers.findIndex(h => h && h.toString().toLowerCase() === 'call_date');
  const fullNameIdx = headers.findIndex(h => h && h.toString().toLowerCase() === 'full_name');
  const lengthInSecIdx = headers.findIndex(h => h && h.toString().toLowerCase() === 'length_in_sec');
  
  if (callDateIdx === -1 || fullNameIdx === -1 || lengthInSecIdx === -1) {
    Logger.log('Required columns not found in Avyukta sheet');
    return;
  }
  
  // Get date range filter
  const dateConfig = getDateRangeConfig();
  const fromDate = new Date(dateConfig.fromDate);
  const toDate = new Date(dateConfig.toDate);
  fromDate.setHours(0, 0, 0, 0);
  toDate.setHours(23, 59, 59, 999);
  
  // Map hours to bucket labels
  const hourToBucket = {
    8: '08-09 AM', 9: '09-10 AM', 10: '10-11 AM', 11: '11-12 PM',
    12: '12-01 PM', 13: '01-02 PM', 14: '02-03 PM', 15: '03-04 PM',
    16: '04-05 PM', 17: '05-06 PM', 18: '06-07 PM', 19: '07-08 PM',
    20: '08-09 PM', 21: '09-10 PM'
  };
  
  let processedCount = 0;
  
  data.forEach(row => {
    const agentName = row[fullNameIdx];
    const callDateValue = row[callDateIdx];
    const duration = row[lengthInSecIdx];
    
    if (!agentName || !callDateValue || !duration || duration <= 0) return;
    
    // Parse call date
    let callDate;
    if (callDateValue instanceof Date) {
      callDate = new Date(callDateValue.getTime());
    } else if (typeof callDateValue === 'string') {
      callDate = new Date(callDateValue);
    } else if (typeof callDateValue === 'number') {
      callDate = new Date((callDateValue - 25569) * 86400 * 1000);
    } else {
      return;
    }
    
    if (isNaN(callDate.getTime())) return;
    
    // Check date range
    const callDateOnly = new Date(callDate);
    callDateOnly.setHours(0, 0, 0, 0);
    
    if (callDateOnly < fromDate || callDateOnly > toDate) {
      return;
    }
    
    // Get hour and bucket
    const hour = callDate.getHours();
    const bucket = hourToBucket[hour];
    
    // Get canonical name and team
    const canonicalAgent = getCanonicalName(agentName, teamMapping);
    const team = getAgentTeam(agentName, teamMapping);
    
    // Initialize if needed
    if (!teamData[team]) {
      teamData[team] = {};
    }
    
if (!teamData[team][canonicalAgent]) {
  teamData[team][canonicalAgent] = {
    totalCalls: 0,
    totalDuration: 0,
    inboundCalls: 0,
    outboundCalls: 0,
    dialerCalls: 0,
    answeredCalls: 0,
    whatsappCalls: 0,
    whatsappDuration: 0,
    avyuktaCalls: 0,
    avyuktaDuration: 0,
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
    
    // Add to team data totals
    teamData[team][canonicalAgent].avyuktaCalls += 1;
    teamData[team][canonicalAgent].avyuktaDuration += duration;
    
    // Add to bucket if valid
    if (bucket && timeBuckets.includes(bucket)) {
      teamData[team][canonicalAgent].buckets[bucket].calls += 1;
      teamData[team][canonicalAgent].buckets[bucket].duration += duration;
    }
    
    processedCount++;
  });
  
  Logger.log(`Avyukta data processed: ${processedCount} calls`);
}


/**
 * Create individual sheets for each team with total calls and talktime
 * Includes hourly buckets, call type breakdown, WhatsApp data, Avyukta data, and Ozonetel data
 */
function createTeamWiseSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getActiveSheet();
  
  // Load team mapping
  const teamMapping = loadTeamMappingFromSheet();

const perfResult  = processAgentPerformanceData(teamMapping);
const agentPerfData = perfResult.data;   // { canonicalName: {idleTime,loginTime,wrapupTime} }
const perfDate    = perfResult.date;      // "02-Mar-26"

// Build column headers that include the date
const perfIdleHeader   = perfDate ? `Idle Time\n(${perfDate})`  : 'Idle Time';
const perfLoginHeader  = perfDate ? `Login Time\n(${perfDate})` : 'Login Time';
const perfWrapupHeader = perfDate ? `Wrapup Time\n(${perfDate})`:'Wrapup Time';
const perfPauseHeader  = perfDate ? `Pause Time\n(${perfDate})`  : 'Pause Time';


  
  if (!teamMapping) {
    const config = getTeamMappingConfig();
    Browser.msgBox(
      'Team Mapping Sheet Not Found',
      `Please create a sheet named "${config.sheetName}" with columns:\n` +
      `- ${config.agentColumn}\n` +
      `- ${config.teamColumn}\n\n` +
      `Then populate it with your agent names and their teams.`,
      Browser.Buttons.OK
    );
    return;
  }
  
  // Get all data
  const lastRow = dataSheet.getLastRow();
  const lastCol = dataSheet.getLastColumn();
  
  if (lastRow < 2) {
    Browser.msgBox('No data found. Please fetch records first.');
    return;
  }
  
  const headers = dataSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = dataSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  
  // Find column indices
  const cleanedAgentNameIdx = headers.indexOf('agent_name_cleaned');
  const cleanedAgentMissedIdx = headers.indexOf('agents_missed_cleaned');
  const callDurationIdx = headers.indexOf('total_call_duration');
  const directionIdx = headers.indexOf('direction');
  const statusIdx = headers.indexOf('status');
  const callHintIdx = headers.indexOf('call_hint');
  
  // Find time column
  const timeColName = 'time';
  let timeIdx = headers.findIndex(h => h && h.toString().toLowerCase() === timeColName);
  if (timeIdx === -1) {
    timeIdx = 11; // Column L as fallback
  }
  
  if (cleanedAgentNameIdx === -1 || cleanedAgentMissedIdx === -1 || callDurationIdx === -1) {
    Browser.msgBox('Required columns not found. Please run fetchCallRecords first.');
    return;
  }
  
  // Define time buckets
  const timeBuckets = [
    '08-09 AM', '09-10 AM', '10-11 AM', '11-12 PM', '12-01 PM', 
    '01-02 PM', '02-03 PM', '03-04 PM', '04-05 PM', '05-06 PM', 
    '06-07 PM', '07-08 PM', '08-09 PM', '09-10 PM'
  ];
  
  // Map hours to bucket labels
  const hourToBucket = {
    8: '08-09 AM', 9: '09-10 AM', 10: '10-11 AM', 11: '11-12 PM',
    12: '12-01 PM', 13: '01-02 PM', 14: '02-03 PM', 15: '03-04 PM',
    16: '04-05 PM', 17: '05-06 PM', 18: '06-07 PM', 19: '07-08 PM',
    20: '08-09 PM', 21: '09-10 PM'
  };
  
  // Collect data by team with hourly buckets
  const teamData = {};
  
  // Process main call records data
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const agentName = row[cleanedAgentNameIdx];
    const agentMissed = row[cleanedAgentMissedIdx];
    const timeValue = row[timeIdx];
    
    let duration = 0;
    const durationValue = row[callDurationIdx];
    
    if (durationValue) {
      if (typeof durationValue === 'number') {
        if (durationValue > 0 && durationValue < 1) {
          duration = Math.round(durationValue * 24 * 60 * 60);
        } else {
          duration = durationValue;
        }
      } else if (typeof durationValue === 'string' && durationValue.includes(':')) {
        duration = hmsToSeconds(durationValue);
      } else if (typeof durationValue === 'string') {
        duration = parseFloat(durationValue) || 0;
      }
    }
    
    const direction = directionIdx !== -1 ? row[directionIdx] : '';
    const status = statusIdx !== -1 ? row[statusIdx] : '';
    
    const hour = getTimeBucket(timeValue);
    const bucket = (hour >= 8 && hour <= 21) ? hourToBucket[hour] : null;
    
    // Process both agents
const agents = [];

// Priority: answered / cleaned agent
if (agentName) {
  agents.push(agentName);
}
// Fallback only if no answered agent
else if (agentMissed) {
  agents.push(agentMissed);
}

const share = 1; // never split

    
    for (let j = 0; j < agents.length; j++) {
      const agent = agents[j];
      const canonicalAgent = getCanonicalName(agent, teamMapping);
      const team = getAgentTeam(agent, teamMapping);
      
      if (!teamData[team]) {
        teamData[team] = {};
      }
      
      if (!teamData[team][canonicalAgent]) {
        teamData[team][canonicalAgent] = {
          totalCalls: 0,
          totalDuration: 0,
          inboundCalls: 0,
          outboundCalls: 0,
          dialerCalls: 0,
          answeredCalls: 0,
          whatsappCalls: 0,
          whatsappDuration: 0,
          avyuktaCalls: 0,
          avyuktaDuration: 0,
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
      
      // Update totals
      teamData[team][canonicalAgent].totalCalls += share;
      teamData[team][canonicalAgent].totalDuration += duration * share;
      
      const isAnswered = status && status.toString().toLowerCase().includes('answered');
      const callHint = callHintIdx !== -1 && row[callHintIdx] ? row[callHintIdx].toString().toLowerCase() : '';
      const isDialer = callHint.includes('dialer');
      const dir = direction?.toString().toLowerCase();
      
      if (isAnswered) {
        teamData[team][canonicalAgent].answeredCalls += share;
      }
      
      if (isDialer) {
        teamData[team][canonicalAgent].dialerCalls += share;
      }
      
      if (dir?.includes('inbound')) {
        teamData[team][canonicalAgent].inboundCalls += share;
      } else if (dir?.includes('outbound')) {
        teamData[team][canonicalAgent].outboundCalls += share;
      }
      
      if (bucket && timeBuckets.includes(bucket)) {
        teamData[team][canonicalAgent].buckets[bucket].calls += share;
        teamData[team][canonicalAgent].buckets[bucket].duration += duration * share;
        
        if (dir && dir.includes('inbound')) {
          teamData[team][canonicalAgent].buckets[bucket].inboundCalls += share;
        }
      }
    }
  }
  
  // Process WhatsApp data
  processWhatsAppDataForTeamSheets(teamData, timeBuckets, teamMapping);
  
  // Process Avyukta data
  processAvyuktaDataForTeamSheets(teamData, timeBuckets, teamMapping);
  
  // Process Ozonetel data
  processOzonetelDataForTeamSheets(teamData, timeBuckets, teamMapping);
  
  // Sort teams alphabetically, but put 'Unassigned' at the end
  const sortedTeams = Object.keys(teamData).sort((a, b) => {
    if (a === 'Unassigned') return 1;
    if (b === 'Unassigned') return -1;
    return a.localeCompare(b);
  });
  
  // Create a sheet for each team
  sortedTeams.forEach(teamName => {
    const agents = teamData[teamName];
    const sheetName = `${teamName}`;
    
    // Create or clear sheet
    let teamSheet = ss.getSheetByName(sheetName);
    if (teamSheet) {
      teamSheet.clear();
    } else {
      teamSheet = ss.insertSheet(sheetName);
    }
    
    // Set default font size for the sheet
teamSheet.getRange("A:Z").setFontSize(13);
teamSheet.getRange("A:Z").setFontFamily("Helvetica Neue");
teamSheet.setRowHeights(1, teamSheet.getMaxRows(), 32); // Add breathing room

    let currentRow = 1;
    
    // ============= QUOTE SECTION =============
    const todaysQuote = getTodaysQuote();
    
    teamSheet.getRange(currentRow, 1, 1, 16).merge();
    teamSheet.getRange(currentRow, 1).setValue('"' + todaysQuote + '"');
    teamSheet.getRange(currentRow, 1)
      .setFontFamily('Georgia')
      .setFontSize(14)
      .setFontStyle('italic')
      .setFontColor('#37474F')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true)
      .setBackground('#FAFAFA');
    
    teamSheet.setRowHeight(currentRow, 48);
    
    teamSheet.getRange(currentRow, 1).setBorder(
      true, true, true, true,
      false, false,
      '#E0E0E0',
      SpreadsheetApp.BorderStyle.SOLID
    );
    currentRow++;
    currentRow++;
    
    // ============= SUMMARY SECTION =============
    
    // Team header
    teamSheet.getRange(currentRow, 1, 1, 11).merge();
    teamSheet.getRange(currentRow, 1).setValue(teamName + ' - SUMMARY');
    teamSheet.getRange(currentRow, 1)
      .setFontWeight('bold')
      .setFontSize(14)
      .setHorizontalAlignment('center');
    
if (teamName === 'Unassigned') {
  teamSheet.getRange(currentRow, 1).setBackground('#8E8E93').setFontColor('#ffffff');
} else {
  teamSheet.getRange(currentRow, 1).setBackground('#007AFF').setFontColor('#ffffff');
}
teamSheet.setRowHeight(currentRow, 40); // Add this line
    currentRow++;
    currentRow++;
    
    // Date range info
    const dateConfig = getDateRangeConfig();
    const lastFetchTime = PropertiesService.getScriptProperties().getProperty('LAST_FETCH_TIME') || 'Never';
    const lastFetchFrom = PropertiesService.getScriptProperties().getProperty('LAST_FETCH_FROM_DATE') || dateConfig.fromDate;
    const lastFetchTo = PropertiesService.getScriptProperties().getProperty('LAST_FETCH_TO_DATE') || dateConfig.toDate;
    
    teamSheet.getRange(currentRow, 1, 1, 11).merge();
    teamSheet.getRange(currentRow, 1).setValue(`Data Range: ${lastFetchFrom} to ${lastFetchTo}  |  Calls Data Till: ${lastFetchTime}`);
teamSheet.getRange(currentRow, 1)
  .setFontSize(11)
  .setHorizontalAlignment('center')
  .setBackground('#F5F5F7')
  .setFontColor('#86868B')
  .setFontWeight('normal'); // Changed from bold
    currentRow++;
    
    // Summary column headers
    const summaryHeaders = [
      'Agent Name', 
      'Total Calls',
      'Inbound', 
      'Outbound', 
      'Dialer', 
      'Answered', 
      'WhatsApp Calls',
      'Avyukta Calls',
      'Ozonetel Calls',
      'Total Talktime (Inc. WA & Avyukta & Ozonetel)',
      'WhatsApp Time',
      'Avyukta Time',
      'Ozonetel Time',
      'Avg Duration',
      'Connectivity %',
// ▼ NEW PERF COLUMNS ▼
      perfIdleHeader,
      perfLoginHeader,
      perfWrapupHeader,
      perfPauseHeader
    ];

    
teamSheet.getRange(currentRow, 1, 1, 19).setValues([summaryHeaders]);
teamSheet.getRange(currentRow, 1, 1, 19)
  .setWrap(true)
  .setFontWeight('600')
  .setFontSize(13)
  .setBackground('#F5F5F7')
  .setFontColor('#1D1D1F')
  .setHorizontalAlignment('center');
teamSheet.setRowHeight(currentRow, 40); // Add this line
    currentRow++;
    
    // Sort agents by total talktime including WhatsApp, Avyukta, and Ozonetel (descending)
    const sortedAgents = Object.entries(agents)
      .sort((a, b) => {
        const aTotalDuration = (a[1].totalDuration || 0) + (a[1].whatsappDuration || 0) + (a[1].avyuktaDuration || 0) + (a[1].ozonetelDuration || 0);
        const bTotalDuration = (b[1].totalDuration || 0) + (b[1].whatsappDuration || 0) + (b[1].avyuktaDuration || 0) + (b[1].ozonetelDuration || 0);
        return bTotalDuration - aTotalDuration;
      });
    
    // Agent summary data
    let teamTotalCalls = 0;
    let teamTotalDuration = 0;
    let isFirstAgent = true;
    
    sortedAgents.forEach(([agentName, stats]) => {
      const whatsappCalls = stats.whatsappCalls || 0;
      const whatsappDuration = stats.whatsappDuration || 0;
      const avyuktaCalls = stats.avyuktaCalls || 0;
      const avyuktaDuration = stats.avyuktaDuration || 0;
      const ozonetelCalls = stats.ozonetelCalls || 0;
      const ozonetelDuration = stats.ozonetelDuration || 0;
      
      const totalCalls = (stats.totalCalls || 0) + whatsappCalls +avyuktaCalls+ozonetelCalls;
      const totalDuration = (stats.totalDuration || 0) + whatsappDuration + avyuktaDuration + ozonetelDuration;
      
      const answered = stats.answeredCalls || 0;
      const avgDuration = answered > 0 ? totalDuration / answered : 0;
      const connectivity = totalCalls > 0 ? (answered / totalCalls) * 100 : 0;

const perfStats = agentPerfData[agentName] || { idleTime: 0, loginTime: 0, wrapupTime: 0, pauseTime: 0 };

const row = [
      agentName,
      Math.round(totalCalls),
      Math.round(stats.inboundCalls || 0),
      Math.round(stats.outboundCalls || 0),
      Math.round(stats.dialerCalls || 0),
      Math.round(stats.answeredCalls || 0),
      Math.round(whatsappCalls),
      Math.round(avyuktaCalls),
      Math.round(ozonetelCalls),
      secondsToSheetDuration(totalDuration),
      secondsToSheetDuration(whatsappDuration),
      secondsToSheetDuration(avyuktaDuration),
      secondsToSheetDuration(ozonetelDuration),
      secondsToSheetDuration(avgDuration),
      connectivity.toFixed(1) + '%',
   //     ▼ NEW PERF COLUMNS ▼
      secondsToSheetDuration(perfStats.idleTime),
      secondsToSheetDuration(perfStats.loginTime),
      secondsToSheetDuration(perfStats.wrapupTime),
      secondsToSheetDuration(perfStats.pauseTime || 0)  // ← col 19
    ];
 
    teamSheet.getRange(currentRow, 1, 1, 19).setValues([row]);
    teamSheet.getRange(currentRow, 2, 1, 8).setNumberFormat('0');
    teamSheet.getRange(currentRow, 2, 1, 13).setHorizontalAlignment('center');
    teamSheet.getRange(currentRow, 10, 1, 5).setNumberFormat('[hh]:mm:ss;[hh]:mm:ss;""');
  //    Format the 3 new perf time columns the same way
    teamSheet.getRange(currentRow, 16, 1, 4).setNumberFormat('[hh]:mm:ss;[hh]:mm:ss;""');
    teamSheet.getRange(currentRow, 16, 1, 4).setHorizontalAlignment('center');
      
// Conditional formatting for total talktime
if (totalDuration > 0) {
  if (totalDuration < 3600) {
    teamSheet.getRange(currentRow, 10).setBackground('#FFECEC').setFontColor('#E65100');
  } else if (totalDuration < 7200) {
    teamSheet.getRange(currentRow, 10).setBackground('#F1F8E9').setFontColor('#558B2F');
  } else if (totalDuration < 14400) {
    teamSheet.getRange(currentRow, 10).setBackground('#E8F5E9').setFontColor('#2E7D32');
  } else {
    teamSheet.getRange(currentRow, 10).setBackground('#C8E6C9').setFontColor('#1B5E20');
    teamSheet.getRange(currentRow, 10).setBorder(false, false, false, true, false, false, '#34C759', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    const currentName = teamSheet.getRange(currentRow, 1).getValue();
    teamSheet.getRange(currentRow, 1).setValue(currentName + ' 🔥');
  }
}

      // ← ADD HERE
if (perfStats.idleTime > 0) {
        const idleCell = teamSheet.getRange(currentRow, 16);
        if (perfStats.idleTime > 7200) {
          idleCell.setBackground('#FFECEC').setFontColor('#C62828');
          if (perfStats.idleTime > 18000) idleCell.setFontWeight('bold'); // bold if > 5 hrs
        } else if (perfStats.idleTime > 3600) {
          idleCell.setBackground('#FFF8E1').setFontColor('#E65100');
        } else {
          idleCell.setBackground('#C8E6C9').setFontColor('#1B5E20');
        }
      }
      // ← END ADD

      // Trophy for top agent
      if (isFirstAgent && totalDuration > 0) {
        const topAgentName = teamSheet.getRange(currentRow, 1).getValue();
        teamSheet.getRange(currentRow, 1).setValue(topAgentName + ' 🏆 ');
        isFirstAgent = false;
      }
      
      teamTotalCalls += totalCalls;
      teamTotalDuration += totalDuration;
      
      currentRow++;
    });
    
    // Calculate team totals
    let teamInbound = 0, teamOutbound = 0, teamDialer = 0, teamAnswered = 0;
    let teamWhatsappCalls = 0, teamWhatsappDuration = 0;
    let teamAvyuktaCalls = 0, teamAvyuktaDuration = 0;
    let teamOzonetelCalls = 0, teamOzonetelDuration = 0;
    
    sortedAgents.forEach(([_, stats]) => {
      teamInbound += stats.inboundCalls;
      teamOutbound += stats.outboundCalls;
      teamDialer += stats.dialerCalls;
      teamAnswered += stats.answeredCalls;
      teamWhatsappCalls += stats.whatsappCalls || 0;
      teamWhatsappDuration += stats.whatsappDuration || 0;
      teamAvyuktaCalls += stats.avyuktaCalls || 0;
      teamAvyuktaDuration += stats.avyuktaDuration || 0;
      teamOzonetelCalls += stats.ozonetelCalls || 0;
      teamOzonetelDuration += stats.ozonetelDuration || 0;
    });
    
    const teamAvgDuration = teamAnswered > 0 ? teamTotalDuration / teamAnswered : 0;
    
    // Calculate peak buckets
    let peakCallBucket = '';
    let peakCallCount = 0;
    let peakDurationBucket = '';
    let peakDuration = 0;
    
    const teamBucketCalls = {};
    const teamBucketDuration = {};
    const teamBucketInbound = {};
    const teamConnectivity = teamTotalCalls > 0 ? (teamAnswered / teamTotalCalls) * 100 : 0;
    timeBuckets.forEach(bucket => {
      teamBucketCalls[bucket] = 0;
      teamBucketDuration[bucket] = 0;
      teamBucketInbound[bucket] = 0;
    });
    
    Object.values(agents).forEach(stats => {
      timeBuckets.forEach(bucket => {
        teamBucketCalls[bucket] += stats.buckets[bucket].calls;
        teamBucketDuration[bucket] += stats.buckets[bucket].duration;
        teamBucketInbound[bucket] += stats.buckets[bucket].inboundCalls || 0;
      });
    });
    
    timeBuckets.forEach(bucket => {
      if (teamBucketInbound[bucket] > peakCallCount) {
        peakCallCount = teamBucketInbound[bucket];
        peakCallBucket = bucket;
      }
      if (teamBucketDuration[bucket] > peakDuration) {
        peakDuration = teamBucketDuration[bucket];
        peakDurationBucket = bucket;
      }
    });
    
       //     Sum perf columns across all agents
     let teamIdleTotal = 0, teamLoginTotal = 0, teamWrapupTotal = 0, teamPauseTotal=0;
     sortedAgents.forEach(([agentName]) => {
       const ps = agentPerfData[agentName] || {};
       teamIdleTotal   += ps.idleTime   || 0;
       teamLoginTotal  += ps.loginTime  || 0;
       teamWrapupTotal += ps.wrapupTime || 0;
       teamPauseTotal  += ps.pauseTime  || 0;  // ← ADD THIS

     });
  
     const totalRow = [
       'TEAM TOTAL',
       Math.round(teamTotalCalls),
       Math.round(teamInbound),
       Math.round(teamOutbound),
       Math.round(teamDialer),
       Math.round(teamAnswered),
       Math.round(teamWhatsappCalls),
       Math.round(teamAvyuktaCalls),
       Math.round(teamOzonetelCalls),
       secondsToSheetDuration(teamTotalDuration),
       secondsToSheetDuration(teamWhatsappDuration),
       secondsToSheetDuration(teamAvyuktaDuration),
       secondsToSheetDuration(teamOzonetelDuration),
       secondsToSheetDuration(teamAvgDuration),
       teamConnectivity.toFixed(1) + '%',
       //   ▼ NEW PERF COLUMNS ▼
       secondsToSheetDuration(teamIdleTotal),
       secondsToSheetDuration(teamLoginTotal),
       secondsToSheetDuration(teamWrapupTotal),
       secondsToSheetDuration(teamPauseTotal)  // ← ADD THIS

     ];
  
     teamSheet.getRange(currentRow, 1, 1, 19).setValues([totalRow]);
     teamSheet.getRange(currentRow, 1, 1, 19)
       .setFontWeight('600').setFontSize(14).setBackground('#FFFFFF').setHorizontalAlignment('center');
     teamSheet.getRange(currentRow, 1, 1, 19).setBorder(
       true, false, false, false, false, false, '#007AFF', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
     );
     teamSheet.getRange(currentRow, 2, 1, 8).setNumberFormat('0');
     teamSheet.getRange(currentRow, 10, 1, 5).setNumberFormat('[hh]:mm:ss;[hh]:mm:ss;""');
     teamSheet.getRange(currentRow, 16, 1, 4).setNumberFormat('[hh]:mm:ss;[hh]:mm:ss;""');
     teamSheet.getRange(currentRow, 16, 1, 4).setHorizontalAlignment('center');
    

    
    currentRow++;
    currentRow += 2;
    
    // ============= HOURLY CALL BREAKDOWN =============
// Get today's date
const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy");

const callHeaderWidth = timeBuckets.length + 2;

teamSheet.getRange(currentRow, 1, 1, callHeaderWidth).merge();

teamSheet.getRange(currentRow, 1)
  .setValue(teamName + ' - HOURLY CALLS (' + today + ')')
  .setFontWeight('600')
  .setFontSize(18)
  .setHorizontalAlignment('center')
  .setBackground('#34C759')
  .setFontColor('#ffffff');

teamSheet.setRowHeight(currentRow, 40);

currentRow++;
    
    const callHeaders = ['Agent', ...timeBuckets, 'Total'];
    teamSheet.getRange(currentRow, 1, 1, callHeaders.length).setValues([callHeaders]);
teamSheet.getRange(currentRow, 1, 1, callHeaders.length)
  .setFontWeight('600')
  .setBackground('#F5F5F7')
  .setFontColor('#1D1D1F')
  .setHorizontalAlignment('center');
teamSheet.setRowHeight(currentRow, 36); // Add this line
    currentRow++;
    
    const teamCallTotals = {};
    timeBuckets.forEach(b => teamCallTotals[b] = 0);
    let teamCallGrandTotal = 0;
    
    const sortedAgentsByCalls = Object.entries(agents)
      .sort((a, b) => b[1].totalCalls - a[1].totalCalls);
    
    sortedAgentsByCalls.forEach(([agentName, stats]) => {
      const row = [agentName];
      let rowTotal = 0;
      
      timeBuckets.forEach(bucket => {
        const count = stats.buckets[bucket].calls;
        if (count > 0) {
          row.push(Math.round(count));
          rowTotal += count;
          teamCallTotals[bucket] += count;
        } else {
          row.push('');
        }
      });
      
      row.push(Math.round(rowTotal));
      teamCallGrandTotal += rowTotal;
      
      teamSheet.getRange(currentRow, 1, 1, callHeaders.length).setValues([row]);
      teamSheet.getRange(currentRow, 2, 1, callHeaders.length - 1).setNumberFormat('0');
      teamSheet.getRange(currentRow, 2, 1, callHeaders.length - 1).setHorizontalAlignment('center');
      
      for (let j = 1; j < callHeaders.length - 1; j++) {
        const cell = teamSheet.getRange(currentRow, j + 1);
        const value = row[j];
        
if (value > 0) {
  if (value < 5) {
    cell.setBackground('#FFECEC').setFontColor('#E65100');
  } else if (value < 15) {
    cell.setBackground('#F1F8E9').setFontColor('#558B2F');
  } else if (value < 30) {
    cell.setBackground('#E8F5E9').setFontColor('#2E7D32');
  } else {
    cell.setBackground('#C8E6C9').setFontColor('#1B5E20');
  }
}
      }
      
      currentRow++;
    });
    
    const teamCallTotalsRow = ['TEAM TOTAL'];
    timeBuckets.forEach(bucket => {
      teamCallTotalsRow.push(Math.round(teamCallTotals[bucket]));
    });
    teamCallTotalsRow.push(Math.round(teamCallGrandTotal));
    
    teamSheet.getRange(currentRow, 1, 1, callHeaders.length).setValues([teamCallTotalsRow]);
teamSheet.getRange(currentRow, 1, 1, callHeaders.length)
  .setFontWeight('600')
  .setFontSize(14)
  .setBackground('#FFFFFF');
teamSheet.getRange(currentRow, 1, 1, callHeaders.length).setBorder(
  true, false, false, false,
  false, false,
  '#34C759',
  SpreadsheetApp.BorderStyle.SOLID_MEDIUM
);
    teamSheet.getRange(currentRow, 2, 1, callHeaders.length - 1).setNumberFormat('0');
    teamSheet.getRange(currentRow, 2, 1, callHeaders.length - 1).setHorizontalAlignment('center');
    currentRow++;
    
    currentRow += 2;
    
    // ============= HOURLY TALKTIME BREAKDOWN =============
// Get today's date
const today2 = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy");

const talktimeHeaderWidth = timeBuckets.length + 2;

teamSheet.getRange(currentRow, 1, 1, talktimeHeaderWidth).merge();

teamSheet.getRange(currentRow, 1)
  .setValue(teamName + ' - HOURLY TALKTIME (' + today2 + ')')
  .setFontWeight('600')
  .setFontSize(18)
  .setHorizontalAlignment('center')
  .setBackground('#FF3B30')
  .setFontColor('#ffffff');

teamSheet.setRowHeight(currentRow, 40);

currentRow++;
    
const talktimeHeaders = (teamName === 'URoots Sales Team' || teamName === 'TTBS Team')
  ? ['Agent', ...timeBuckets, 'Total', 'Status']
  : ['Agent', ...timeBuckets, 'Total'];
    teamSheet.getRange(currentRow, 1, 1, talktimeHeaders.length).setValues([talktimeHeaders]);
teamSheet.getRange(currentRow, 1, 1, talktimeHeaders.length)
  .setFontWeight('600')
  .setBackground('#F5F5F7')
  .setFontColor('#1D1D1F')
  .setHorizontalAlignment('center');
teamSheet.setRowHeight(currentRow, 36); // Add this line
    currentRow++;
    
    const teamTalktimeTotals = {};
    timeBuckets.forEach(b => teamTalktimeTotals[b] = 0);
    let teamTalktimeGrandTotal = 0;
    
    const sortedAgentsByTalktime = Object.entries(agents)
      .sort((a, b) => {
        const aDuration = (b[1].totalDuration || 0) + (b[1].whatsappDuration || 0) + (b[1].avyuktaDuration || 0) + (b[1].ozonetelDuration || 0);
        const bDuration = (a[1].totalDuration || 0) + (a[1].whatsappDuration || 0) + (a[1].avyuktaDuration || 0) + (a[1].ozonetelDuration || 0);
        return aDuration - bDuration;
      });
    
    sortedAgentsByTalktime.forEach(([agentName, stats]) => {
      const row = [agentName];
      let rowTotal = 0;
      
      timeBuckets.forEach(bucket => {
        const duration = stats.buckets[bucket].duration;
        row.push(secondsToHMS(duration));
        rowTotal += duration;
        teamTalktimeTotals[bucket] += duration;
      });
      
      const totalAgentDuration = (stats.totalDuration || 0) + (stats.whatsappDuration || 0) + (stats.avyuktaDuration || 0) + (stats.ozonetelDuration || 0);
      row.push(secondsToHMS(totalAgentDuration));
      teamTalktimeGrandTotal += totalAgentDuration;

      // Status: Full Day if >= 2.5 hrs (9000 seconds), else Half Day
const statusLabel = totalAgentDuration >= 9000 ? 'Full Day' : 'Half Day';
const isTrackedAgent = teamName === 'TTBS Team' ? ['Urwah', 'Vamika'].includes(agentName) : true;
if (teamName === 'URoots Sales Team' || teamName === 'TTBS Team') {
  row.push(isTrackedAgent ? statusLabel : '');
}
      
      teamSheet.getRange(currentRow, 1, 1, talktimeHeaders.length).setValues([row]);
      teamSheet.getRange(currentRow, 2, 1, talktimeHeaders.length - 1).setHorizontalAlignment('center');

      // Color the Status cell
if ((teamName === 'URoots Sales Team' || teamName === 'TTBS Team') && isTrackedAgent) {
  const statusCell = teamSheet.getRange(currentRow, talktimeHeaders.length);
  if (statusLabel === 'Full Day') {
    statusCell.setBackground('#C8E6C9').setFontColor('#1B5E20').setFontWeight('bold');
  } else {
    statusCell.setBackground('#FFECEC').setFontColor('#C62828').setFontWeight('bold');
  }
  statusCell.setHorizontalAlignment('center');
}
      
      for (let j = 1; j < talktimeHeaders.length - 1; j++) {
        const cell = teamSheet.getRange(currentRow, j + 1);
        const value = row[j];
        
        if (value && value !== '' && value !== '00:00:00') {
          const seconds = hmsToSeconds(value);
          
if (seconds < 300) {
  cell.setBackground('#FFECEC').setFontColor('#E65100');
} else if (seconds < 900) {
  cell.setBackground('#F1F8E9').setFontColor('#558B2F');
} else if (seconds < 1800) {
  cell.setBackground('#E8F5E9').setFontColor('#2E7D32');
} else {
  cell.setBackground('#C8E6C9').setFontColor('#1B5E20');
  cell.setBorder(false, false, false, true, false, false, '#34C759', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  cell.setValue(value); 
}
        }
      }
      
      currentRow++;
    });
    
    const teamTalktimeTotalsRow = ['TEAM TOTAL'];
    timeBuckets.forEach(bucket => {
      teamTalktimeTotalsRow.push(secondsToHMS(teamTalktimeTotals[bucket]));
    });
    teamTalktimeTotalsRow.push(secondsToHMS(teamTalktimeGrandTotal));
    if (teamName === 'URoots Sales Team' || teamName === 'TTBS Team') teamTalktimeTotalsRow.push(''); // Status column — blank for team total
    
    teamSheet.getRange(currentRow, 1, 1, talktimeHeaders.length).setValues([teamTalktimeTotalsRow]);
teamSheet.getRange(currentRow, 1, 1, talktimeHeaders.length)
  .setFontWeight('600')
  .setFontSize(14)
  .setBackground('#FFFFFF');
teamSheet.getRange(currentRow, 1, 1, talktimeHeaders.length).setBorder(
  true, false, false, false,
  false, false,
  '#FF3B30',
  SpreadsheetApp.BorderStyle.SOLID_MEDIUM
);
    teamSheet.getRange(currentRow, 2, 1, talktimeHeaders.length - 1).setHorizontalAlignment('center');
    currentRow++;
    
    // Format columns
teamSheet.setColumnWidth(1, 220);
for (let i = 2; i <= Math.max(callHeaders.length, talktimeHeaders.length); i++) {
  teamSheet.setColumnWidth(i, 110);
}
    // Add timestamp
    currentRow += 2;
    teamSheet.getRange(currentRow, 1).setValue(`✓ Updated on ${new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' })}`);
    teamSheet.getRange(currentRow, 1).setFontWeight('bold').setFontColor('#0f9d58');
  });
  
  Logger.log(`Created ${sortedTeams.length} team sheets.`);
  Browser.msgBox('Success', `Created ${sortedTeams.length} team-wise sheets with WhatsApp, Avyukta, and Ozonetel data!`, Browser.Buttons.OK);
}




function secondsToSheetDuration(seconds) {
  return seconds > 0 ? seconds / 86400 : 0;
}




function runDailyReport() {
  fetchCallRecords();
  Utilities.sleep(5000);
  importCDRFromGitHub();
  Utilities.sleep(5000);
  
   // Wait 5 seconds for data to settle
  createTeamWiseSheets();

  Utilities.sleep(5000);
  createCompleteCallSummary_Tata_Ozonetel_WhatsApp();
}




/**
 * Get a random quote from the Daily Quotes sheet
 * Uses a seed based on today's date so the same quote shows all day
 */
function getTodaysQuote() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quotesSheet = ss.getSheetByName('Daily Quotes');
  
  if (!quotesSheet) {
    return 'Daily Quotes sheet not found.';
  }
  
  // Get all quotes
  const lastRow = quotesSheet.getLastRow();
  if (lastRow < 1) {
    return 'No quotes available.';
  }
  
  const quotes = quotesSheet.getRange(1, 1, lastRow, 1).getValues();
  const quoteList = quotes.map(row => row[0]).filter(q => q && q.toString().trim() !== '');
  
  if (quoteList.length === 0) {
    return 'No quotes available.';
  }
  
  // Get today's date as a seed (YYYY-MM-DD format)
  const today = new Date();
  const dateString = today.getFullYear() + '-' + 
                     String(today.getMonth() + 1).padStart(2, '0') + '-' + 
                     String(today.getDate()).padStart(2, '0');
  
  // Simple hash function to convert date string to a number
  let hash = 0;
  for (let i = 0; i < dateString.length; i++) {
    hash = ((hash << 5) - hash) + dateString.charCodeAt(i);
    hash = hash & hash; // Convert to 32-bit integer
  }
  
  // Use hash to get consistent index for today
  const index = Math.abs(hash) % quoteList.length;
  
  return quoteList[index];
}





function debugAgentLookup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getActiveSheet();
  const teamMapping = loadTeamMappingFromSheet();
  
  // Get first 50 rows of data
  const headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  const data = dataSheet.getRange(2, 1, Math.min(50, dataSheet.getLastRow() - 1), dataSheet.getLastColumn()).getValues();
  
  const cleanedAgentNameIdx = headers.indexOf('agent_name_cleaned');
  const agentNameIdx = headers.indexOf('agent_name');
  
  Logger.log('=== AGENT LOOKUP DEBUG ===');
  Logger.log('Team Mapping Keys: ' + Object.keys(teamMapping.teamMap).join(', '));
  Logger.log('');
  
  const uniqueAgents = {};
  
  for (let i = 0; i < data.length; i++) {
    const originalName = data[i][agentNameIdx];
    const cleanedName = data[i][cleanedAgentNameIdx];
    
    if (originalName && !uniqueAgents[originalName]) {
      uniqueAgents[originalName] = true;
      
      const team = getAgentTeam(originalName, teamMapping);
      const canonical = getCanonicalName(originalName, teamMapping);
      
      Logger.log(`Original: "${originalName}" | Cleaned: "${cleanedName}" | Team: ${team} | Canonical: ${canonical}`);
    }
  }
}




function debugUnassignedAgents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getActiveSheet();
  const teamMapping = loadTeamMappingFromSheet();
  
  const headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  const lastRow = dataSheet.getLastRow();
  const data = dataSheet.getRange(2, 1, lastRow - 1, dataSheet.getLastColumn()).getValues();
  
  const cleanedAgentNameIdx = headers.indexOf('agent_name_cleaned');
  const cleanedAgentMissedIdx = headers.indexOf('agents_missed_cleaned');
  
  const unassigned = {};
  const assigned = {};

  for (let i = 0; i < data.length; i++) {
    const agentName = data[i][cleanedAgentNameIdx] || data[i][cleanedAgentMissedIdx];
    if (!agentName) continue;
    
    const team = getAgentTeam(agentName, teamMapping);
    const canonical = getCanonicalName(agentName, teamMapping);
    const cleanedInput = cleanAgentName(agentName);
    const lookupKey = teamMapping.lookupMap ? teamMapping.lookupMap[cleanedInput] : 'N/A';
    
    if (team === 'Unassigned') {
      if (!unassigned[agentName]) {
        unassigned[agentName] = {
          cleanedInput: cleanedInput,
          lookupKey: lookupKey || 'NOT FOUND IN LOOKUPMAP',
          canonical: canonical
        };
      }
    } else {
      if (!assigned[agentName]) assigned[agentName] = team;
    }
  }

  Logger.log('=== UNASSIGNED AGENTS ===');
  Logger.log('Total unique unassigned names: ' + Object.keys(unassigned).length);
  for (const [name, info] of Object.entries(unassigned)) {
    Logger.log(`"${name}" → cleaned: "${info.cleanedInput}" → lookupKey: "${info.lookupKey}"`);
  }
  
  Logger.log('\n=== LOOKUPMAP KEYS (from Team Mapping sheet) ===');
  if (teamMapping.lookupMap) {
    for (const [key, val] of Object.entries(teamMapping.lookupMap)) {
      Logger.log(`"${key}" → "${val}"`);
    }
  }
  
  Logger.log('\n=== TEAMMAP KEYS (exact originals in mapping sheet) ===');
  for (const key of Object.keys(teamMapping.teamMap)) {
    Logger.log(`"${key}"`);
  }
}









function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('★★ QHT Reporting Team ★★')
    .addItem('Just Breathe','Welcome to QHT')
    .addSeparator()

    .addToUi();
}