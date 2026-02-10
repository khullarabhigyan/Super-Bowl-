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

function getSubmissionsSheet(ss) {
  // Try common names, then fall back to first sheet that isn't Answer Key
  var names = ['Sheet1', 'Submissions', 'sheet1'];
  for (var n = 0; n < names.length; n++) {
    var s = ss.getSheetByName(names[n]);
    if (s) return s;
  }
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() !== 'Answer Key') return sheets[i];
  }
  return sheets[0];
}

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var data = JSON.parse(e.postData.contents);

    // ---- Answer Key POST ----
    if (data.type === 'answer_key') {
      var akSheet = ss.getSheetByName('Answer Key');
      if (!akSheet) {
        akSheet = ss.insertSheet('Answer Key');
        akSheet.appendRow(['Question', 'Answer Indices', 'Answer Texts', 'Locked At']);
        akSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
        akSheet.setFrozenRows(1);
      }

      // Check if headers need updating (old format had singular names)
      var headerRow = akSheet.getRange(1, 1, 1, 4).getValues()[0];
      if (headerRow[1] === 'Answer Index') {
        akSheet.getRange(1, 2).setValue('Answer Indices');
        akSheet.getRange(1, 3).setValue('Answer Texts');
      }

      var qNum = data.question;
      // Support both single and multiple answers
      var ansIndices = data.answerIndices || [data.answerIndex];
      var ansTexts = data.answerTexts || [data.answerText];
      var existing = akSheet.getDataRange().getValues();

      var indicesStr = ansIndices.join(',');
      var textsStr = ansTexts.join('||');

      var found = false;
      for (var i = 1; i < existing.length; i++) {
        if (existing[i][0] == qNum) {
          akSheet.getRange(i + 1, 2).setValue(indicesStr);
          akSheet.getRange(i + 1, 3).setValue(textsStr);
          akSheet.getRange(i + 1, 4).setValue(new Date().toLocaleString('en-US', { timeZone: 'America/New_York' }));
          found = true;
          break;
        }
      }

      if (!found) {
        akSheet.appendRow([
          qNum,
          indicesStr,
          textsStr,
          new Date().toLocaleString('en-US', { timeZone: 'America/New_York' })
        ]);
      }

      return ContentService
        .createTextOutput(JSON.stringify({ status: 'success', action: 'answer_key_saved' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ---- Regular Submission POST ----
    var sheet = getSubmissionsSheet(ss);

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

    // Return answer key
    var akSheet = ss.getSheetByName('Answer Key');
    var answerKey = {};
    if (akSheet && akSheet.getLastRow() > 1) {
      var akData = akSheet.getDataRange().getValues();
      for (var i = 1; i < akData.length; i++) {
        var rawIndices = String(akData[i][1]);
        var rawTexts = String(akData[i][2]);

        // Parse multi-answer format
        var indices = rawIndices.split(',').map(function(x) { return parseInt(x.trim()); });
        var texts = rawTexts.split('||');

        answerKey[akData[i][0]] = {
          answerIndices: indices,
          answerTexts: texts,
          lockedAt: akData[i][3]
        };
      }
    }
    result.answerKey = answerKey;

    // Return submissions
    var sheet = getSubmissionsSheet(ss);
    var submissions = [];
    if (sheet && sheet.getLastRow() > 1) {
      var data = sheet.getDataRange().getValues();
      var headers = data[0];
      for (var j = 1; j < data.length; j++) {
        var entry = {};
        for (var k = 0; k < headers.length; k++) {
          entry[headers[k]] = data[j][k];
        }
        submissions.push(entry);
      }
    }
    result.submissions = submissions;

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
