// =============================================================
// SUPER BOWL LX PREDICTION GAME - Google Apps Script
// =============================================================
// INSTRUCTIONS:
// 1. Open your Google Sheet
// 2. Go to Extensions > Apps Script
// 3. Delete any existing code and paste this entire file
// 4. Click "Deploy" > "Manage deployments"
// 5. Edit the existing deployment and choose "New version"
// 6. Click "Deploy" to update
//
// NOTE: You MUST create a new version for changes to take effect.
// =============================================================

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var data = JSON.parse(e.postData.contents);

    // ---- Answer Key POST ----
    if (data.type === 'answer_key') {
      var akSheet = ss.getSheetByName('Answer Key');
      if (!akSheet) {
        akSheet = ss.insertSheet('Answer Key');
        akSheet.appendRow(['Question', 'Answer Index', 'Answer Text', 'Locked At']);
        akSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
        akSheet.setFrozenRows(1);
      }

      var qNum = data.question;
      var ansIdx = data.answerIndex;
      var ansText = data.answerText;
      var existing = akSheet.getDataRange().getValues();

      // Check if this question already has an answer
      var found = false;
      for (var i = 1; i < existing.length; i++) {
        if (existing[i][0] == qNum) {
          // Update existing row
          akSheet.getRange(i + 1, 2).setValue(ansIdx);
          akSheet.getRange(i + 1, 3).setValue(ansText);
          akSheet.getRange(i + 1, 4).setValue(new Date().toLocaleString('en-US', { timeZone: 'America/New_York' }));
          found = true;
          break;
        }
      }

      if (!found) {
        akSheet.appendRow([
          qNum,
          ansIdx,
          ansText,
          new Date().toLocaleString('en-US', { timeZone: 'America/New_York' })
        ]);
      }

      return ContentService
        .createTextOutput(JSON.stringify({ status: 'success', action: 'answer_key_saved' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ---- Regular Submission POST ----
    var sheet = ss.getSheetByName('Sheet1') || ss.getSheets()[0];

    if (sheet.getLastRow() === 0) {
      var headers = [
        "Timestamp", "Name",
        "Q1: Winner", "Q2: First TD Team", "Q3: First TD Scorer",
        "Q4: Opening Song", "Q5: First Commercial", "Q6: Coin Toss",
        "Q7: Anthem Length", "Q8: Guest Performer", "Q9: Gatorade Color",
        "Q10: Total FGs", "Q11: First Penalty", "Q12: Songs Performed",
        "Q13: MVP", "Q14: Winning Margin", "Q15: DtMF Performed",
        "Tiebreaker: Seahawks Score", "Tiebreaker: Patriots Score"
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    var row = [
      new Date().toLocaleString("en-US", { timeZone: "America/New_York" }),
      data.name
    ];

    for (var i = 1; i <= 15; i++) {
      row.push(data.answers[i] ? data.answers[i].answer : "");
    }

    row.push(data.tiebreaker ? data.tiebreaker.seahawks : 0);
    row.push(data.tiebreaker ? data.tiebreaker.patriots : 0);

    sheet.appendRow(row);

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
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result = { status: 'success' };

    // Always return answer key
    var akSheet = ss.getSheetByName('Answer Key');
    var answerKey = {};
    if (akSheet && akSheet.getLastRow() > 1) {
      var akData = akSheet.getDataRange().getValues();
      for (var i = 1; i < akData.length; i++) {
        answerKey[akData[i][0]] = {
          answerIndex: akData[i][1],
          answerText: akData[i][2],
          lockedAt: akData[i][3]
        };
      }
    }
    result.answerKey = answerKey;

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
