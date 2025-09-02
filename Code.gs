const SHEET_ID = '여기에_스프레드시트_ID';

// JSON 출력 유틸
function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// GET: ?sheet=시트이름&range=A1:C3
function doGet(e) {
  try {
    const sheetName = (e && e.parameter && e.parameter.sheet) || 'Sheet1';
    const range = (e && e.parameter && e.parameter.range) || 'A1:C10';

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error("시트 없음: " + sheetName);

    const values = sheet.getRange(range).getDisplayValues();
    return jsonOut({ok:true, sheet:sheetName, values});
  } catch (err) {
    return jsonOut({ok:false, error:String(err)});
  }
}

// POST: body = {sheet:"Sheet2", mode:"append", values:["a","b"]}
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    const sheetName = data.sheet || 'Sheet1';

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error("시트 없음: " + sheetName);

    if (data.mode === 'append') {
      sheet.appendRow(data.values);
      return jsonOut({ok:true, sheet:sheetName});
    } else if (data.mode === 'update') {
      const rng = sheet.getRange(data.range);
      rng.setValues(data.values);
      return jsonOut({ok:true, sheet:sheetName});
    } else {
      return jsonOut({ok:false, error:"unknown mode"});
    }
  } catch (err) {
    return jsonOut({ok:false, error:String(err)});
  }
}
