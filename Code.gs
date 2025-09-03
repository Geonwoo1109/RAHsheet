const SHEET_ID = '1Ae7cFXb6Q6jS6elzE_3xFgsBUsCquBvwW3FK4iHZGIQ';

// JSON 출력 유틸
function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// GET: ?sheet=시트명&range=A1:D10
function doGet(e) {
  try {
    const sheetName = e.parameter.sheet || 'Sheet1';
    const rangeA1 = e.parameter.range || 'A1:C10';
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error("시트 없음: " + sheetName);

    const range = sheet.getRange(rangeA1);
    const values = range.getDisplayValues();

    // 병합 범위를 확인하고 확장값 채우기
    const mergedRanges = range.getMergedRanges();
    mergedRanges.forEach(mr => {
      const val = mr.getCell(1,1).getDisplayValue();
      const r = mr.getRow() - range.getRow() + 1;
      const c = mr.getColumn() - range.getColumn() + 1;
      const numRows = mr.getNumRows();
      const numCols = mr.getNumColumns();
      for (let i = 0; i < numRows; i++) {
        for (let j = 0; j < numCols; j++) {
          if (values[r-1+i] && values[r-1+i][c-1+j] !== undefined) {
            values[r-1+i][c-1+j] = val;
          }
        }
      }
    });

    return jsonOut({ok:true, sheet:sheetName, range:rangeA1, values});
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
      // appendRow는 Range를 반환하지 않으므로, 마지막 행 Range를 가져와 스타일 지정
      sheet.appendRow(data.values);
      const lastRow = sheet.getLastRow();
      const range = sheet.getRange(lastRow, 1, 1, data.values.length);

      if (data.backgrounds) {
        range.setBackgrounds([data.backgrounds]);
      }
      if (data.fontColors) {
        range.setFontColors([data.fontColors]);
      }
      if (data.bold) {
        range.setFontWeight("bold");
      }

      range.setHorizontalAlignment("center");
      range.setVerticalAlignment("middle");

      return jsonOut({ok:true, sheet:sheetName});
    } 
    else if (data.mode === 'update') {
      const range = sheet.getRange(data.range);
      range.setValues(data.values);

      if (data.backgrounds) {
        rng.setBackgrounds(data.backgrounds);
      }
      if (data.fontColors) {
        rng.setFontColors(data.fontColors);
      }
      if (data.bold) {
        rng.setFontWeight("bold");
      }

      range.setHorizontalAlignment("center");
      range.setVerticalAlignment("middle");

      return jsonOut({ok:true, sheet:sheetName});
    }
    else if (data.mode == "mergeWrite") {
      const range = sheet.getRange(data.range);
      if (data.backgrounds) {
        range.setBackgrounds(data.backgrounds);
      }
      range.getCell(1,1).setValue(data.value);    // 값 입력 (대표칸)
      range.merge();
      
      range.setHorizontalAlignment("center");
      range.setVerticalAlignment("middle");
      
      return jsonOut({ok:true, sheet:sheetName, action:"mergeWrite"});
    }
    else {
      return jsonOut({ok:false, error:"unknown mode"});
    }
  } catch (err) {
    return jsonOut({ok:false, error:String(err)});
  }
}

