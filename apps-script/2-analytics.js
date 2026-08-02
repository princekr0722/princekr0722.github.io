/**
 * Analytics logger — writes each event to a tab named after its `type`.
 *
 * NOT named doPost: this project also serves the chat assistant, and Apps
 * Script allows only one doPost per project (all files share one global
 * namespace, so a second definition silently overrides the first). 1-router.js
 * owns doPost and dispatches here. See 1-router.js.
 */
function handleAnalytics(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(5000); // Wait up to 5s for other concurrent requests

  try {
    const requestData = JSON.parse(e.postData.contents);
    const type = requestData.type || "General_Logs";
    const payload = requestData.payload || {};

    // Open Spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(type);

    // --- CONFIGURATION ---
    // 1. Priority Headers: These will always be the first columns (Left -> Right)
    const PRIORITY_HEADERS = [
      "timestamp",
      "uid",
      "session_id",
      "ip_address",
      "page_url",
      "screen_res",
      "type",
    ];

    // 2. Create Sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(type);
      // Initialize with Priority + Raw_Payload + userAgent (as explicitly requested at end)
      // Any other custom fields will be appended after these dynamically
      const initialHeaders = [...PRIORITY_HEADERS, "Raw_Payload", "userAgent"];
      sheet.appendRow(initialHeaders);
      sheet.setFrozenRows(1);
    }

    // 3. Map Payload to Columns
    try {
      // Get current headers from the sheet to match payload keys to columns
      const lastCol = sheet.getLastColumn();
      let sheetHeaders = [];

      if (lastCol > 0) {
        sheetHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      } else {
        // Fallback if sheet is empty but exists
        sheetHeaders = [...PRIORITY_HEADERS, "Raw_Payload", "userAgent"];
        sheet.appendRow(sheetHeaders);
      }

      // 4. Dynamic Column Creation
      // Check if payload has keys that aren't in the headers yet
      const payloadKeys = Object.keys(payload);
      const newKeys = payloadKeys.filter((key) => !sheetHeaders.includes(key));

      if (newKeys.length > 0) {
        // Add new headers to the end of the first row
        const startCol = sheetHeaders.length + 1;
        sheet.getRange(1, startCol, 1, newKeys.length).setValues([newKeys]);

        // Update local header reference so we can map data correctly immediately
        sheetHeaders = [...sheetHeaders, ...newKeys];
      }

      // 5. Construct Row Data
      // Map the header list to values in the payload
      const rowData = sheetHeaders.map((header) => {
        // Special Handling for specific columns
        if (header === "timestamp") {
          const ts = payload.timestamp || new Date();
          const date = new Date(ts);

          // Convert to IST offset
          const istOffset = 5.5 * 60; // IST = UTC+5:30 in minutes
          const istDate = new Date(date.getTime() + istOffset * 60 * 1000);

          // Extract parts
          const YYYY = istDate.getUTCFullYear();
          const MM = String(istDate.getUTCMonth() + 1).padStart(2, "0"); // months 0-11
          const DD = String(istDate.getUTCDate()).padStart(2, "0");
          const HH = String(istDate.getUTCHours()).padStart(2, "0");
          const min = String(istDate.getUTCMinutes()).padStart(2, "0");
          const sec = String(istDate.getUTCSeconds()).padStart(2, "0");
          const ms = String(istDate.getUTCMilliseconds()).padStart(3, "0");

          const formatted = `${YYYY}/${MM}/${DD} ${HH}:${min}:${sec}:${ms}`;
          return formatted;
        }

        if (header === "type") return type;

        // Get value from payload
        let val = payload[header];

        // Safety: Stringify objects (like nested JSON) so they fit in one cell
        if (typeof val === "object" && val !== null) {
          return JSON.stringify(val);
        }

        // Return value or empty string if undefined
        return val !== undefined ? val : "";
      });

      // 6. Append the Row
      sheet.appendRow(rowData);
    } catch (innerError) {
      // --- FAILSAFE ERROR HANDLING ---
      // If column mapping crashes, dump data to "Raw_Payload" column to prevent data loss.
      console.error("Column mapping failed: " + innerError);

      // We construct a simple row targeting the indices we know generally exist
      // timestamp (0), ..., type (6), Raw_Payload (7)
      // We fill the first few known columns and dump the rest in Raw_Payload
      const timestamp = new Date();
      const safePayload = JSON.stringify(payload);
      const errorMsg = "ERROR: " + innerError.toString();

      // Append a safe row: [Timestamp, "", "", ..., Type, Payload, ErrorMessage]
      sheet.appendRow([
        timestamp,
        "",
        "",
        "",
        "",
        "",
        type,
        safePayload,
        errorMsg,
      ]);
    }

    return createJSONOutput({
      status: "success",
      message: `Data logged to ${type}`,
    });
  } catch (error) {
    // Catch-all for top-level errors (e.g., JSON parse fail)
    return createJSONOutput({ status: "error", message: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

function createJSONOutput(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(
    ContentService.MimeType.JSON,
  );
}

/** Health check. Reached via 1-router.js — see the note on handleAnalytics. */
function analyticsHealth(e) {
  return createJSONOutput({ status: "active", message: "API is online." });
}
