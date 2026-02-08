// =============================================================
// SUPER BOWL LX PREDICTION GAME - Google Apps Script
// =============================================================
// INSTRUCTIONS:
// 1. Open your Google Sheet
// 2. Go to Extensions > Apps Script
// 3. Delete any existing code and paste this entire file
// 4. Click "Deploy" > "New deployment"
// 5. Type = "Web app"
// 6. Execute as: "Me"
// 7. Who has access: "Anyone"
// 8. Click "Deploy" and authorize when prompted
// 9. Copy the Web App URL and provide it back
// =============================================================

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);

    // Create header row if sheet is empty
    if (sheet.getLastRow() === 0) {
      var headers = [
        "Timestamp",
        "Name",
        "Q1: Winner",
        "Q2: First TD Team",
        "Q3: First TD Scorer",
        "Q4: Opening Song",
        "Q5: First Commercial",
        "Q6: Coin Toss",
        "Q7: Anthem Length",
        "Q8: Guest Performer",
        "Q9: Gatorade Color",
        "Q10: Total FGs",
        "Q11: First Penalty",
        "Q12: Songs Performed",
        "Q13: MVP",
        "Q14: Winning Margin",
        "Q15: DtMF Performed",
        "Tiebreaker: Seahawks Score",
        "Tiebreaker: Patriots Score"
      ];
      sheet.appendRow(headers);

      // Bold the header row
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    // Build the row
    var row = [
      new Date().toLocaleString("en-US", { timeZone: "America/New_York" }),
      data.name
    ];

    // Add answers for Q1-Q15
    for (var i = 1; i <= 15; i++) {
      row.push(data.answers[i] ? data.answers[i].answer : "");
    }

    // Add tiebreaker scores
    row.push(data.tiebreaker ? data.tiebreaker.seahawks : 0);
    row.push(data.tiebreaker ? data.tiebreaker.patriots : 0);

    sheet.appendRow(row);

    // Auto-resize columns on first few entries
    if (sheet.getLastRow() <= 3) {
      sheet.autoResizeColumns(1, row.length);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success", row: sheet.getLastRow() }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: "success", submissions: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var headers = data[0];
    var submissions = [];

    for (var i = 1; i < data.length; i++) {
      var entry = {};
      for (var j = 0; j < headers.length; j++) {
        entry[headers[j]] = data[i][j];
      }
      submissions.push(entry);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success", submissions: submissions }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
