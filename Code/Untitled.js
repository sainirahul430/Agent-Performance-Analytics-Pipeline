/****************************************************
 * CONFIGURATION
 ****************************************************/
/**
 * @OnlyCurrentDoc false
 */

const REPORT_CONFIG1 = [
  {
    sheetName: 'Sales Team',
    webhookUrl: 'https://chat.googleapis.com/v1/spaces/AAQAG1cWOtk/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=y-zRy3i2I55kLskJ6Y61XJgl5mw9YiJhMtsznNopSLM',
    label: 'Sales Team – FULL SHEET REPORT'
  },
  {
    sheetName: 'HT Done Team',
    webhookUrl: 'https://chat.googleapis.com/v1/spaces/AAQAYrMDbMU/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=yWxMdq6gNO8mV8-624qi3RRtt-G8mPBQlrOTLBrrqXs',
    label: 'HT Done Team – FULL SHEET REPORT'
  },
  {
    sheetName: 'URoots Sales Team',
    webhookUrl: 'https://chat.googleapis.com/v1/spaces/AAQAI2Uq6xo/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=8btKYNiy0j2mZ55FnJWm9YmIBD_zoDcVV_45cwLszdc',
    label: 'URoots Sales Team – FULL SHEET REPORT'
  },
  {
    sheetName: 'URoots Confirmation Team',
    webhookUrl: 'https://chat.googleapis.com/v1/spaces/AAQA0IaYnfI/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=tbiuLbpmGbWFUv9P6M107tU7yst87tR57LoDr7J9P6c',
    label: 'URoots Confirmation Team – FULL SHEET REPORT'
  },
  {
    sheetName: 'Incoming Dept.',
    webhookUrl: 'https://chat.googleapis.com/v1/spaces/AAQAG1cWOtk/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=y-zRy3i2I55kLskJ6Y61XJgl5mw9YiJhMtsznNopSLM',
    label: 'Incoming Dept – FULL SHEET REPORT'
  },
  {
    sheetName: 'Welcome Team',
    webhookUrl: 'https://chat.googleapis.com/v1/spaces/AAQAL6s5j8c/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=i5oJ-Y3Bex39jy6g4Y-6mBHXCmKEMphScz50-UdrOVY',
    label: 'Welcome Team – FULL SHEET REPORT'
  },
  {
    sheetName: 'Datealignment Team',
    webhookUrl: 'https://chat.googleapis.com/v1/spaces/AAQAgHe8HPU/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=DxbMQtedJPAyi8bCg_lGOtqnTPxpV3fV2xNqvywoi3s',
    label: 'Datealignment Team – FULL SHEET REPORT'
  }
];


/****************************************************
 * MAIN ENTRY POINT — RUN THIS FUNCTION ONLY
 ****************************************************/
function sendAllFullSheetReports() {
  for (var i = 0; i < REPORT_CONFIG1.length; i++) {
    var config = REPORT_CONFIG1[i];
    if (!config || !config.sheetName) continue;
    exportSheetAndSend(config);
  }
}


function exportSheetAndSend(config) {
  if (!config || !config.sheetName) {
    Logger.log('exportSheetAndSend: missing config. Run sendAllFullSheetReports instead.');
    return;
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName(config.sheetName);

    if (!sourceSheet) {
      Logger.log('Sheet not found: ' + config.sheetName);
      return;
    }

    // Direct link to the specific tab
    var sheetUrl = ss.getUrl() + '#gid=' + sourceSheet.getSheetId();

    var now = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'dd MMM yyyy hh:mm a'
    );

    var message =
      '📊 ' + config.label + '\n' +
      '📅 ' + now + '\n\n' +
      '📎 Full sheet link:\n' + sheetUrl;

    UrlFetchApp.fetch(config.webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: message })
    });

    Logger.log('Sent: ' + config.sheetName);

  } catch (e) {
    Logger.log('Error (' + config.sheetName + '): ' + e.message);
  }
}