function sendConnectivityReportToGChat() {
  const WEBHOOK_URL = 'https://chat.googleapis.com/v1/spaces/AAQAapfqgQ8/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=kuZfcjtuURpFBEPBSCDIMum9YXeSJjNvHpDjiB1Iaaw';

  const ss = SpreadsheetApp.openById('1Dad77vTHkxTh8jPag6h3l_Ygxv6RWmXvMx3lIk58_Zk');
  const sheet = ss.getSheetByName('Connectivity Report');

  // ─── Column layout (9 cols): A=Total, B=Outbound, C=Inbound, D=Dialer,
  //                              E=Answered, F=Talktime, G=Connectivity%,
  //                              H=Active Agents, I=Avg Answered/Agent

// 🔵 TATA  (data row → 3)
const tataTotal      = sheet.getRange('A3').getValue();
const tataOutbound   = sheet.getRange('B3').getValue();
const tataInbound    = sheet.getRange('C3').getValue();
const tataAnswered   = sheet.getRange('E3').getValue();
const tataTalktime   = sheet.getRange('F3').getDisplayValue();
const tataConn       = sheet.getRange('G3').getDisplayValue();
const tataActive     = sheet.getRange('H3').getValue();
const tataAvgAns     = sheet.getRange('I3').getValue();
const tataAvgDials   = sheet.getRange('J3').getValue();   // ← new

// 🟠 OZONETEL  (data row → 9)
const ozTotal        = sheet.getRange('A9').getValue();
const ozOutbound     = sheet.getRange('B9').getValue();
const ozInbound      = sheet.getRange('C9').getValue();
const ozAnswered     = sheet.getRange('E9').getValue();
const ozTalktime     = sheet.getRange('F9').getDisplayValue();
const ozConn         = sheet.getRange('G9').getDisplayValue();
const ozActive       = sheet.getRange('H9').getValue();
const ozAvgAns       = sheet.getRange('I9').getValue();
const ozAvgDials     = sheet.getRange('J9').getValue();   // ← new

  // 🟢 WHATSAPP  (data row → 15)
  const waDials        = sheet.getRange('A15').getValue();
  const waTalk         = sheet.getRange('B15').getDisplayValue();

  // ⚫ GRAND TOTAL  (data row → 20)
  const grandTotal     = sheet.getRange('A20').getValue();
  const grandTalk      = sheet.getRange('B20').getDisplayValue();
  const reportDate     = sheet.getRange('C20').getDisplayValue();
  const grandAnswered  = tataAnswered + ozAnswered;
  const grandActive    = tataActive + ozActive;

  const message =
`📊 *Connectivity Report — ${reportDate}*


🔵 *TATA*
Total Calls: *${tataTotal}*   |   Outbound: *${tataOutbound}*   |   Inbound: *${tataInbound}*
Answered: *${tataAnswered}*   |   Connectivity: *${tataConn}*
Talktime: *${tataTalktime}*
Active Agents: *${tataActive}*   |   Avg Answered/Agent: *${tataAvgAns}*   |   Avg Dials/Agent: *${tataAvgDials}*


🟠 *OZONETEL*
Total Calls: *${ozTotal}*   |   Outbound: *${ozOutbound}*   |   Inbound: *${ozInbound}*
Answered: *${ozAnswered}*   |   Connectivity: *${ozConn}*
Talktime: *${ozTalktime}*
Active Agents: *${ozActive}*   |   Avg Answered/Agent: *${ozAvgAns}*   |   Avg Dials/Agent: *${ozAvgDials}*


🟢 *WHATSAPP*
Total Dials: *${waDials}*   |   Talktime: *${waTalk}*


⚫ *GRAND TOTAL*
Total Calls: *${grandTotal}*   |   Answered: *${grandAnswered}*
Total Active Agents: *${grandActive}*
Total Talktime: *${grandTalk}*`;

  UrlFetchApp.fetch(WEBHOOK_URL, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ text: message })
  });
}