function importCDRFromGitHub() {
  const url = "https://raw.githubusercontent.com/John-Cena-DEV/cdr-fetcher/refs/heads/main/cdr_data_master.csv";

  try {
    const response = UrlFetchApp.fetch(url);

    if (response.getResponseCode() !== 200) {
      throw new Error("Failed to fetch data");
    }

    const text = response.getContentText();
    const rows = Utilities.parseCsv(text);

    if (!rows || rows.length < 2) {
      throw new Error("CSV format invalid or empty");
    }

    let rawJson = rows[1][0];

    // 🔍 Debug (first run only)
    Logger.log("RAW START:");
    Logger.log(rawJson.substring(0, 500));

    // =========================
    // 🧹 CLEANING START
    // =========================
    rawJson = rawJson.trim();

    // Remove outer quotes if entire JSON is wrapped
    if (rawJson.startsWith('"') && rawJson.endsWith('"')) {
      rawJson = rawJson.slice(1, -1);
    }

    // Fix escaped quotes
    rawJson = rawJson.replace(/\\"/g, '"');

    // Convert Python literals → JSON
    rawJson = rawJson
      .replace(/\bNone\b/g, "null")
      .replace(/\bTrue\b/g, "true")
      .replace(/\bFalse\b/g, "false");

    // Convert single-quoted keys → double-quoted
    rawJson = rawJson.replace(/([{,]\s*)'([^']+?)'\s*:/g, '$1"$2":');

    // Convert single-quoted string values → double-quoted
    rawJson = rawJson.replace(/:\s*'([^']*?)'/g, ': "$1"');

    // Remove trailing commas
    rawJson = rawJson
      .replace(/,\s*}/g, "}")
      .replace(/,\s*]/g, "]");

    // =========================
    // 🧠 PARSE JSON SAFELY
    // =========================
    let jsonData;
    try {
      jsonData = JSON.parse(rawJson);
    } catch (e) {
      Logger.log("❌ JSON PARSE FAILED:");
      Logger.log(rawJson.substring(0, 1000));
      throw e;
    }

    if (!Array.isArray(jsonData) || jsonData.length === 0) {
      throw new Error("Parsed JSON is empty or not an array");
    }

    // =========================
    // 🔄 FLATTEN JSON
    // =========================
    function flattenObject(obj, prefix = "") {
      return Object.keys(obj).reduce((acc, k) => {
        const pre = prefix ? prefix + "_" : "";
        if (typeof obj[k] === "object" && obj[k] !== null) {
          Object.assign(acc, flattenObject(obj[k], pre + k));
        } else {
          acc[pre + k] = obj[k];
        }
        return acc;
      }, {});
    }

    const flatData = jsonData.map(obj => flattenObject(obj));

    // =========================
    // 📊 PREPARE SHEET DATA
    // =========================
    const headers = Object.keys(flatData[0]);
    const output = [headers];

    flatData.forEach(obj => {
      output.push(headers.map(h => obj[h] ?? ""));
    });

    // =========================
    // 📄 WRITE TO SHEET
    // =========================
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("CDR_Data");

    if (!sheet) {
      sheet = ss.insertSheet("CDR_Data");
    }

    sheet.clear();
    sheet.getRange(1, 1, output.length, headers.length).setValues(output);

    Logger.log("✅ Data imported successfully");

  } catch (err) {
    Logger.log("❌ ERROR: " + err.message);
    throw err;
  }
}