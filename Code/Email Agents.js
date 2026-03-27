function sendDailyAgentSummaryEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const emailMapping = loadAgentEmailMapping();
  const teamMapping = loadTeamMappingFromSheet();

  if (!emailMapping || !teamMapping) return;

  const teams = Object.keys(teamMapping.teamMap)
    .map(agent => getAgentTeam(agent, teamMapping));

  const uniqueTeams = [...new Set(teams)].filter(t => t !== 'Unassigned');

  let sentCount = 0;

  uniqueTeams.forEach(teamName => {
    const sheet = ss.getSheetByName(teamName);
    if (!sheet) return;

    const data = sheet.getDataRange().getDisplayValues();

    let headerRowIndex = data.findIndex(r => r[0] === 'Agent Name');
    if (headerRowIndex === -1) return;

    const headers = data[headerRowIndex];

    for (let r = headerRowIndex + 1; r < data.length; r++) {
      const row = data[r];
      const agentName = row[0];

      if (!agentName || agentName.includes('TEAM TOTAL')) break;

      const cleanName = agentName.replace(/🏆|🔥|⚠️/g, '').trim();
      const email = emailMapping[cleanName];
      if (!email) continue;

      const summaryTable = buildAgentSummaryTable(headers, row);

      sendAgentSummaryEmail(
        cleanName,
        email,
        teamName,
        summaryTable,
        row
      );

      sentCount++;
    }
  });

  Logger.log(`Daily summaries sent: ${sentCount}`);
}


function loadAgentEmailMapping() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Agent Emails'); // exact name
  if (!sheet) {
    throw new Error('Agent Email Mapping sheet not found');
  }

  const data = sheet.getDataRange().getValues();
  const mapping = {};

  for (let i = 1; i < data.length; i++) {
    const agent = String(data[i][0]).trim();
    const email = String(data[i][1]).trim();

    if (agent && email) {
      mapping[agent] = email;
    }
  }

  return mapping;
}



function sendAgentSummaryEmail(
  agentName,
  email,
  teamName,
  summaryTable,
  row
) {
  const dateConfig = getDateRangeConfig();
  const dateLabel = Utilities.formatDate(
  new Date(dateConfig.fromDate),
  "IST",
  "d MMM yyyy"
);

  const subject = `📊 Daily Performance Summary – ${dateLabel}`;

const totalCalls  = row[1];
const inbound     = row[2];
const outbound    = row[3];
const answered    = row[5];

const waCalls     = row[6];
const avyukta     = row[7];
const ozonetel    = row[8];

const totalTalktime = row[9];
const WhatsAppTalktime= row[10];
const OzonetelTalktime= row[12];

const avgDuration   = row[13];


  const highlightStrip = `
    <span style="font-weight:600;color:#1d1d1f;">${totalCalls} calls</span>
    <span style="color:#6e6e73;"> • </span>
    <span>${answered} answered</span>
    <span style="color:#6e6e73;"> • </span>
    <span>${totalTalktime} talktime</span>
    <span style="color:#6e6e73;"> • </span>
    <span>${avgDuration} avg</span>
  `;

  const htmlBody = `
  <div style="
    background:#ffffff;
    padding:32px 16px;
    font-family:-apple-system,BlinkMacSystemFont,'SF Pro Text','SF Pro Display',Helvetica,Arial,sans-serif;
    color:#1d1d1f;
  ">
    <div style="max-width:900px;margin:0 auto;">

      <h1 style="
        margin:0 0 8px 0;
        font-size:28px;
        font-weight:600;
        letter-spacing:-0.02em;
      ">
         Daily Performance Summary
      </h1>

      <p style="
        margin:0 0 24px 0;
        font-size:15px;
        color:#6e6e73;
      ">
        ${highlightStrip}
      </p>

      <p style="margin:0 0 8px 0;font-size:15px;color:#6e6e73;">
        Hi <strong style="color:#1d1d1f;">${agentName}</strong>,
      </p>

      <p style="margin:0 0 20px 0;font-size:15px;color:#6e6e73;">
        Here is your performance summary for
        <strong style="color:#1d1d1f;">${dateLabel}</strong>.
      </p>

      <p style="margin:0 0 28px 0;font-size:15px;color:#6e6e73;">
        <strong style="color:#1d1d1f;">${teamName}</strong> 
      </p>

<!-- Tabular Summary -->
<div style="
  border-top:1px solid #d2d2d7;
  border-bottom:1px solid #d2d2d7;
  padding:8px 0;
">
  <table style="
    width:100%;
    border-collapse:collapse;
    font-family:-apple-system,BlinkMacSystemFont,'SF Pro Text','SF Pro Display',Helvetica,Arial,sans-serif;
    font-size:14px;
  ">
    <tbody>

      <tr>
        <td style="padding:12px 0;color:#6e6e73;">Agent</td>
        <td style="padding:12px 0;font-weight:500;color:#1d1d1f;">${agentName}</td>
      </tr>

      <tr style="border-top:1px solid #ededf0;">
        <td style="padding:12px 0;color:#6e6e73;">Total Calls</td>
        <td style="padding:12px 0;color:#1d1d1f;">${totalCalls}</td>
      </tr>

      <tr style="border-top:1px solid #ededf0;">
        <td style="padding:12px 0;color:#6e6e73;">Answered</td>
        <td style="padding:12px 0;color:#1d1d1f;">${answered}</td>
      </tr>

      <tr style="border-top:1px solid #ededf0;">
        <td style="padding:12px 0;color:#6e6e73;">Inbound / Outbound</td>
        <td style="padding:12px 0;color:#1d1d1f;">${inbound} / ${outbound}</td>
      </tr>

      <tr style="border-top:1px solid #ededf0;">
        <td style="padding:12px 0;color:#6e6e73;">WhatsApp / Avyukta / Ozonetel</td>
        <td style="padding:12px 0;color:#1d1d1f;">${waCalls} / ${avyukta} / ${ozonetel}</td>
      </tr>

      <tr style="border-top:1px solid #ededf0;">
        <td style="padding:12px 0;color:#6e6e73;">Total Talktime</td>
        <td style="padding:12px 0;color:#1d1d1f;">${totalTalktime}</td>
      </tr>


      <tr style="border-top:1px solid #ededf0;">
        <td style="padding:12px 0;color:#6e6e73;">WA Talktime</td>
        <td style="padding:12px 0;color:#1d1d1f;">${WhatsAppTalktime}</td>
      </tr>


      <tr style="border-top:1px solid #ededf0;">
        <td style="padding:12px 0;color:#6e6e73;">Ozonetel Talktime</td>
        <td style="padding:12px 0;color:#1d1d1f;">${OzonetelTalktime}</td>
      </tr>


      <tr style="border-top:1px solid #ededf0;">
        <td style="padding:12px 0;color:#6e6e73;">Avg Duration</td>
        <td style="padding:12px 0;color:#1d1d1f;">${avgDuration}</td>
      </tr>

    </tbody>
  </table>
</div>


      <p style="margin-top:32px;font-size:13px;color:#86868b;">
        This is an automated daily summary from QHT Performance System. 
      </p>

    </div>
  </div>
  `;

  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: htmlBody,
    body: `Hi ${agentName},\n\nYour daily performance summary for ${dateLabel} is ready.`
  });
}



function buildAgentSummaryTable(headers, row) {
  return `
    <table style="width:100%;border-collapse:collapse;font-size:14px;">
      <thead>
        <tr>
          ${headers.map(h => `
            <th style="
              text-align:left;
              padding:12px 14px;
              font-weight:500;
              color:#6e6e73;
              border-bottom:1px solid #d2d2d7;
              white-space:nowrap;
            ">
              ${h}
            </th>
          `).join('')}
        </tr>
      </thead>
      <tbody>
        <tr>
          ${row.map(v => `
            <td style="
              padding:14px;
              color:#1d1d1f;
              border-bottom:1px solid #f5f5f7;
              white-space:nowrap;
            ">
              ${v || '—'}
            </td>
          `).join('')}
        </tr>
      </tbody>
    </table>
  `;
}
