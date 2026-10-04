// SmartSites Google Sheets receiver
// 1. Open your Google Sheet.
// 2. Go to Extensions > Apps Script.
// 3. Paste this code into the Apps Script editor.
// 4. Replace the sheet name below if needed.
// 5. Deploy > New deployment > Web app.
// 6. Set "Who has access" to Anyone.
// 7. Copy the Web app URL into QUOTE_SHEET_URL in quote.html.

const SHEET_NAME = "Requests";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || "{}");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

    if (!sheet) {
      throw new Error('Sheet "' + SHEET_NAME + '" was not found.');
    }

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Date", "Estimated Total", "Selections"]);
    }

    const selections = data.selections || {};

    const formattedSelections = Object.entries(selections)
      .map(([question, answer]) => question + ": " + answer)
      .join(" | ");

    sheet.appendRow([
      new Date(),
      data.total || "$0",
      formattedSelections
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: error.message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}