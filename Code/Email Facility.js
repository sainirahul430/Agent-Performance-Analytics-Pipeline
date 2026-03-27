/**
 * MULTI-TEAM EMAIL CONFIGURATION
 * Send multiple team summaries in a single email
 */
function getMultiTeamEmailConfig() {
  return {
    // Format: 'Email description': { recipients: ['emails'], teams: ['Team Names'] }
    
    'QHT Management': {
      recipients: ['drankur@qhtclinic.com','shipra@qhtclinic.com'],
      teams: ['Sales Team','Datealignment Team','HT Done Team','Incoming Dept.','URoots Confirmation Team','URoots Sales Team','Welcome Team']
    },
    
    'QHT Analytics Team': {
      recipients: ['rahulcreative@qhtclinic.com','nitin@qhtclinic.com'],
      teams: ['Sales Team','Datealignment Team','HT Done Team','Incoming Dept.','URoots Confirmation Team','URoots Sales Team','Welcome Team']
    },
    
    'QHT Management-Sales & DA': {
      recipients: ['drakansha@qhtclinic.com','rajshree@qhtclinic.com'],
      teams: ['Sales Team', 'Datealignment Team']
    }
    
    // Add more combinations as needed
  };
}

/**
 * Send multiple team summaries in a single consolidated email
 */
function sendMultiTeamSummary(configName, teamList, recipients) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dateConfig = getDateRangeConfig();
  
  // Get last fetch time
  const lastFetchTime = PropertiesService.getScriptProperties().getProperty('LAST_FETCH_TIME') || 'Just now';
  const reportDate = new Date().toLocaleDateString('en-IN', { 
    weekday: 'long', 
    year: 'numeric', 
    month: 'long', 
    day: 'numeric' 
  });
  
  // Get today's quote (from first team sheet)
  const firstTeamSheet = ss.getSheetByName(teamList[0]);
  const todaysQuote = firstTeamSheet ? firstTeamSheet.getRange('A1').getValue().toString().replace(/"/g, '') : '';
  
  // Build consolidated HTML
  let allTeamTables = '';
  
  teamList.forEach((teamName, index) => {
    const teamSheet = ss.getSheetByName(teamName);
    
    if (!teamSheet) {
      Logger.log(`Sheet not found for team: ${teamName}`);
      return;
    }
    
    // ⭐ DYNAMIC RANGE DETECTION
    const lastRow = teamSheet.getLastRow();
    const data = teamSheet.getRange(1, 1, lastRow, 10).getValues();
    
    let summaryStartRow = -1;
    let summaryEndRow = -1;
    
    // Find "Agent Name" header
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === 'Agent Name') {
        summaryStartRow = i;
        break;
      }
    }
    
    // Find "TEAM TOTAL" row
    for (let i = summaryStartRow + 1; i < data.length; i++) {
      if (data[i][0] === 'TEAM TOTAL') {
        summaryEndRow = i;
        break;
      }
    }
    
    if (summaryStartRow === -1 || summaryEndRow === -1) {
      Logger.log(`Could not find summary for team: ${teamName}`);
      return;
    }
    
    const summaryData = data.slice(summaryStartRow, summaryEndRow + 1);
    const backgrounds = teamSheet.getRange(summaryStartRow + 1, 1, summaryEndRow - summaryStartRow + 1, 10).getBackgrounds();
    const fontColors = teamSheet.getRange(summaryStartRow + 1, 1, summaryEndRow - summaryStartRow + 1, 10).getFontColors();
    const fontWeights = teamSheet.getRange(summaryStartRow + 1, 1, summaryEndRow - summaryStartRow + 1, 10).getFontWeights();
    
    // Build table for this team
    let htmlTable = '';
    
    // Add team header
    allTeamTables += `
      <div style="margin-top: ${index > 0 ? '40px' : '0'};">
        <h2 style="color: #667eea; font-size: 24px; font-weight: 700; margin: 0 0 20px 0; padding-bottom: 10px; border-bottom: 3px solid #667eea;">
          ${teamName}
        </h2>
    `;
    
    // Build table rows
    for (let i = 0; i < summaryData.length; i++) {
      const isHeader = (i === 0);
      const isTotal = summaryData[i][0] === 'TEAM TOTAL';
      
      htmlTable += '<tr>';
      
      for (let j = 0; j < summaryData[i].length; j++) {
        let value = summaryData[i][j];
        const bg = backgrounds[i] ? backgrounds[i][j] : '#ffffff';
        const color = fontColors[i] ? fontColors[i][j] : '#212121';
        const weight = fontWeights[i] && fontWeights[i][j] === 'bold' ? 'bold' : 'normal';
        
// ⭐ Convert decimal time values to HH:MM:SS
if ((j >= 9 && j <= 13) && typeof value === 'number' && value > 0) {
  const totalSeconds = Math.round(value * 24 * 60 * 60);
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;
  value = `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
} 
else if ((j >= 9 && j <= 13) && (value === 0 || value === '')) {
  value = '';
}

// Format call count numbers
if (j >= 1 && j <= 8 && typeof value === 'number') {
  value = Math.round(value);
}

        
        // Clean agent names
        if (j === 0 && typeof value === 'string') {
          value = value.replace(/[🏆🔥⚠️]/g, '').trim();
        }
        
        const cellTag = isHeader ? 'th' : 'td';
        
        let cellStyle = `
          background-color: ${bg || '#ffffff'};
          color: ${color || '#212121'};
          font-weight: ${weight};
          padding: ${isHeader ? '16px 12px' : '12px'};
          text-align: ${j === 0 ? 'left' : 'center'};
          border-bottom: 1px solid #e0e0e0;
          font-size: ${isHeader ? '13px' : '14px'};
        `;
        
        if (isTotal) {
          cellStyle += 'background: #fff9c4 !important; font-weight: bold;';
        }
        
        htmlTable += `<${cellTag} style="${cellStyle}">${value}</${cellTag}>`;
      }
      
      htmlTable += '</tr>';
    }
    
    allTeamTables += `
        <div style="background: #ffffff; border-radius: 8px; overflow: hidden; border: 1px solid #e0e0e0;">
          <table style="width: 100%; border-collapse: collapse;">
            ${htmlTable}
          </table>
        </div>
      </div>
    `;
  });
  
  // Create consolidated email
  const subject = `📊  Agent Input Matrix for ${configName} (${dateConfig.fromDate.split(' ')[0]})`;
  
  const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="margin: 0; padding: 20px; font-family: 'Segoe UI', Roboto, Arial, sans-serif; background: #f5f5f5;">
  
  <div style="max-width: 1200px; margin: 0 auto; background: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 12px rgba(0,0,0,0.1);">
    
    <!-- Header Section -->
    <div style="background: #ffffff; padding: 40px 30px; text-align: center; border-bottom: 3px solid #667eea;">
      <h1 style="color: #667eea; font-size: 32px; font-weight: 700; margin: 0 0 10px 0; letter-spacing: -0.5px;">
        Report for ${configName}
      </h1>
      <div style="margin-top: 16px; padding-top: 16px; border-top: 1px solid #e0e0e0;">
<p style="
  display: inline-block;
  padding: 10px 18px;
  font-size: 18px;
  font-weight: 600;
  border-radius: 999px;
  background: #1f1f1f;
  color: #ffd54f;
  margin: 0;
">
  ⚡ Built by QHT Analytics Team
</p>

      </div>


      <p style="color: #9e9e9e; font-size: 14px; margin: 10px 0 0 0;">
        ${teamList.length} Team${teamList.length > 1 ? 's' : ''} Included
      </p>
    </div>
    
    <!-- Quote Section -->
    ${todaysQuote ? `
    ` : ''}
    
    <!-- Info Cards Section -->
    <div style="padding: 30px; background: #ffffff;">
      <table style="width: 100%; border-spacing: 0;">
        <tr>
          <td style="width: 50%; padding-right: 10px;">
            <div style="background: #fafafa; padding: 20px; border-radius: 8px; border: 1px solid #e0e0e0;">
              <div style="color: #5f6368; font-size: 12px; font-weight: 600; text-transform: uppercase; margin-bottom: 8px; letter-spacing: 0.5px;">📅 Report Date</div>
              <div style="color: #202124; font-size: 16px; font-weight: 600;">${reportDate}</div>
            </div>
          </td>
          <td style="width: 50%; padding-left: 10px;">
            <div style="background: #fafafa; padding: 20px; border-radius: 8px; border: 1px solid #e0e0e0;">
              <div style="color: #5f6368; font-size: 12px; font-weight: 600; text-transform: uppercase; margin-bottom: 8px; letter-spacing: 0.5px;">🕐 Last Updated</div>
              <div style="color: #202124; font-size: 16px; font-weight: 600;">${lastFetchTime}</div>
            </div>
          </td>
        </tr>
      </table>
    </div>
    
    <!-- All Team Tables -->
    <div style="padding: 0 30px 30px 30px;">
      ${allTeamTables}
    </div>
    
    <!-- Footer Section -->
    <div style="background: #fafafa; padding: 24px 30px; text-align: center; border-top: 1px solid #e0e0e0;">
      <p style="color: #757575; font-size: 13px; margin: 0 0 8px 0;">
        📊 This is an automated consolidated report from your QHT Analytics Team
      </p>
      <p style="color: #9e9e9e; font-size: 12px; margin: 0;">
        For detailed hourly breakdown, please access the <a href="${SpreadsheetApp.getActiveSpreadsheet().getUrl()}" style="color: #667eea; text-decoration: none; font-weight: 600;">full dashboard</a>
      </p>
      
    </div>
    
  </div>
  
</body>
</html>
  `;
  
  // Send email
  try {
    MailApp.sendEmail({
      to: recipients.join(','),
      subject: subject,
      htmlBody: htmlBody
    });
    Logger.log(`Consolidated email sent to ${recipients.join(', ')} with teams: ${teamList.join(', ')}`);
  } catch (error) {
    Logger.log(`Error sending consolidated email: ${error.toString()}`);
  }
}

/**
 * Send all configured multi-team summaries
 */
function sendAllMultiTeamSummaries() {
  const config = getMultiTeamEmailConfig();
  
  let sentCount = 0;
  
  Object.keys(config).forEach(configName => {
    const setup = config[configName];
    sendMultiTeamSummary(configName, setup.teams, setup.recipients);
    sentCount++;
    
    // Delay to avoid hitting email limits
    Utilities.sleep(2000);
  });
  
  Logger.log(`Successfully sent ${sentCount} consolidated emails`);
  Browser.msgBox('Success', `Sent ${sentCount} consolidated team reports!`, Browser.Buttons.OK);
}