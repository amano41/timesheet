function doPost(e) {

  const params = JSON.parse(e.postData.getDataAsString());
  const props = PropertiesService.getScriptProperties().getProperties();

  // 身元確認
  const token = props.VERIFICATION_TOKEN;
  if (params.token != token) {
    throw new Error("Invalid Token");
  }

  const text = params.text.toLowerCase();
  Logger.log(text);

  const mode = parseMode(text);
  const date = parseDate(text);
  Logger.log("mode: " + mode);
  Logger.log("date: " + date);

  const sheet = getSheet(props.SPREADSHEET_ID, date.getFullYear());
  Logger.log("sheet ID: " + sheet.getSheetId());

  writeTimestamp(sheet, mode, date);
}

const parseMode = function (text) {
  if (text.match(/hello/)) {
    return "hello"
  }
  else if (text.match(/bye/)) {
    return "bye"
  }
  else {
    throw new Error("Unknown Mode");
  }
}

const parseDate = function (text) {

  const today = new Date();
  let year, month, day, hour, minute;

  // 日付
  const date_regex = /(\d{4}\/)?(\d{1,2})\/(\d{1,2})/;
  const date_match = text.match(date_regex);
  if (date_match) {
    year = date_match[1] ? parseInt(date_match[1]) : today.getFullYear();
    month = parseInt(date_match[2]);
    day = parseInt(date_match[3]);
  }
  else if (text.match(/tod(ay)?/)) {
    year = today.getFullYear();
    month = today.getMonth() + 1;  // getMonth() は 0-11 を返す
    day = today.getDate();
    Logger.log(day)
  }
  else if (text.match(/yes(terday)?/)) {
    const yesterday = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);
    year = yesterday.getFullYear();
    month = yesterday.getMonth() + 1;
    day = yesterday.getDate();
  }
  else if (text.match(/tom(orrow)?/)) {
    const tomorrow = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1);
    year = tomorrow.getFullYear();
    month = tomorrow.getMonth() + 1;
    day = tomorrow.getDate();
  }
  else {
    year = today.getFullYear();
    month = today.getMonth() + 1;
    day = today.getDate();
  }

  if (year < 1970 || month < 1 || month > 12 || day < 1 || day > 31) {
    throw new Error("Invalid Date: " + date_match[0]);
  }

  // 時刻
  const time_regex = /(\d{1,2})\s*:\s*(\d{1,2})(\s*:\s*\d{1,2})?/;
  const time_match = text.match(time_regex);
  if (time_match) {
    hour = parseInt(time_match[1]);
    minute = parseInt(time_match[2]);
  }
  else {
    hour = today.getHours();
    minute = today.getMinutes();
  }

  if (hour < 0 || hour > 23 || minute < 0 || minute > 59) {
    throw new Error("Invalid Time: " + time_match[0]);
  }

  return new Date(year, month - 1, day, hour, minute);
}

const getSheet = function (spreadsheetId, year) {
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  let sheet = spreadsheet.getSheetByName(year);
  if (!sheet) {
    sheet = spreadsheet.insertSheet();
    sheet.setName(year);

    // 見出し
    sheet.getRange(1, 1).setValue("日付");
    sheet.getRange(1, 2).setValue("出勤");
    sheet.getRange(1, 3).setValue("退勤");
    sheet.getRange(1, 4).setValue("開始時間");
    sheet.getRange(1, 5).setValue("終了時間");
    sheet.getRange(1, 6).setValue("勤務時間");

    // 1月1日
    sheet.getRange(2, 1).setValue(year + "/01/01");

    // 勤務時間計算用の数式
    sheet.getRange(2, 4).setFormula("=IF(B2<>\"\", TIME(HOUR(B2),FLOOR(MINUTE(B2),5),0), IF(WEEKDAY(A2,2)<6, TIME(9,0,0), \"\"))");
    sheet.getRange(2, 5).setFormula("=IF(C2<>\"\", TIME(HOUR(C2),CEILING(MINUTE(C2),5),0), IF(WEEKDAY(A2,2)<6, TIME(17,45,0), \"\"))");
    sheet.getRange(2, 6).setFormula("=IF(AND(E2<>\"\",D2<>\"\",(E2-D2)*60*24<>8*60+45), E2-D2, \"\")");
  }
  return sheet;
}

const writeTimestamp = function (sheet, mode, date) {

  // 1月1日から指定日までの経過日数
  const d1 = new Date(date.getFullYear(), 0, 1, 0, 0, 0)
  const d2 = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0);
  const date_diff = parseInt((d2.getTime() - d1.getTime()) / (1000 * 60 * 60 * 24));

  const row = date_diff + 2;
  const col = mode === "hello" ? 2 : 3;
  const date_str = Utilities.formatDate(date, "Asia/Tokyo", "yyyy/MM/dd");
  const time_str = Utilities.formatDate(date, "Asia/Tokyo", "HH:mm");

  // 日付と時刻を記録
  sheet.getRange(row, 1).setValue(date_str);
  sheet.getRange(row, col).setValue(time_str);
  Logger.log(Utilities.formatString("[%d, %d] = %s", row, 1, date_str));
  Logger.log(Utilities.formatString("[%d, %d] = %s", row, col, time_str));

  // 日付をオートフィル
  let src = sheet.getRange(2, 1);
  let dest = sheet.getRange(2, 1, date_diff + 1);
  src.autoFill(dest, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  // 数式をオートフィル
  src = sheet.getRange(2, 4, 1, 3);
  dest = sheet.getRange(2, 4, date_diff + 1, 3)
  src.autoFill(dest, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}
