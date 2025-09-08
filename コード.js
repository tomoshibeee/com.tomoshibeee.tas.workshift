function init() {
  const title = "初期設定";
  const prompt = "対象年月は設定しましたか?";

  const response = Browser.msgBox(title, prompt, Browser.Buttons.YES_NO);
  if (response == "yes") {
    setEnvVariable();
    Browser.msgBox("完了しました。")
  }
}

function setEnvVariable() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("BAR_COLOR", "red");
}

function copyAllWorkerMonthlyTimeTable() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let workerSheet = spreadsheet.getSheetByName("Worker");
  let range = workerSheet.getRange("E:E");
  var values = range.getValues();
  for (let i = 0; i < values.length; i++) {
    if (i == 0) {
      continue;
    }
    if (values[i][0] == "") {
      break;
    }
    copyWorkerMonthlyTimeTable(values[i][0]);
  }

  const title = "出勤予定";
  const prompt = "出勤予定を入力してください";
  const response = Browser.msgBox(title, prompt, Browser.Buttons.YES_NO);

}

function copyWorkerMonthlyTimeTable(newSheetName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  var templateSheetName = 'X0000:name';
  var templateSheet = spreadsheet.getSheetByName(templateSheetName);

  if (sheetExists(newSheetName)) {
    Logger.log(`シート「${newSheetName}」という名前のシートは存在するため、作成しませんでした。`);
    return;
  }

  if (templateSheet) {
    var newSheet = templateSheet.copyTo(spreadsheet);
    newSheet.setName(newSheetName);
    var range = newSheet.getRange("G1");
    range.setValue(newSheetName);
    newSheet.activate();
    Logger.log('シート「' + templateSheetName + '」の完全な複製を「' + newSheetName + '」という名前で作成しました。');
  } else {
    Logger.log('エラー: テンプレートシート "' + templateSheetName + '" が見つかりませんでした。');
  }
}

function removeAllWorkerMonthlyTimeTable() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let workerSheet = spreadsheet.getSheetByName("Worker");
  let range = workerSheet.getRange("E:E");
  var values = range.getValues();
  for (let i = 0; i < values.length; i++) {
    if (i == 0) {
      continue;
    }
    if (values[i][0] == "") {
      break;
    }
    var sheet = spreadsheet.getSheetByName(values[i][0]);
    if (sheet) {
      spreadsheet.deleteSheet(sheet);
    }
  }
}

function createWorkshift(date) {
  const title = "シフト表作成";
  const prompt = "シフト表を作成してもよろしいですか?";

  const response = Browser.msgBox(title, prompt, Browser.Buttons.YES_NO);
  if (response == "yes") {
    Browser.msgBox("test")
  }
}

function makeWorkShiftDiagram() {
  const sheetName = getActiveSheetName()
  if (!checkDayFormat(sheetName)) {
    const title = "シフト表作成";
    const prompt = "シフト表を作成できませんでした";
    const response = Browser.msgBox(title, prompt, Browser.Buttons.OK);
    return 
  }

  // --- user settings ---
  const startRow = 3
  const startColumn = "K";
  const colorDefault = "white"
  // --- user settings ---

  let sheet = SpreadsheetApp.getActiveSheet();

  // シートを初期化
  let rangeColumn = null;
  rangeColumn = sheet.getRange(`${startColumn}${startRow+1}:BR`);
  rangeColumn.setBackground(colorDefault)
  rangeColumn.setBorder(false,false,false,false,false,false);

  var lastRow = sheet.getMaxRows();
  sheet.showRows(1, lastRow);

  // TODO : シート名を確認
  let range = sheet.getRange("A:I");
  let values = range.getValues();
  for (let i = 0; i < values.length; i++) {
      if (i < startRow) {
        continue;
      }
      if (values[i][0] == "") {
        rangeColumn = sheet.getRange(`$A${startRow+1}:H${i}`);
        rangeColumn.setBorder(true,true,true,true,true, true,"black", SpreadsheetApp.BorderStyle.SOLID);
        rangeColumn = sheet.getRange(`${startColumn}${startRow+1}:BR${i}`);
        rangeColumn.setBorder(true,true,true,true,true, true,"black", SpreadsheetApp.BorderStyle.SOLID);
        break;
      }
      Logger.log(`名前:${values[i][4]}`);
      makeWorkShiftRow(i+1, values[i][5], values[i][6])
  }
}

function columnNumberToLetter(columnNumber) {
  var letter = '';
  while (columnNumber > 0) {
    var modulo = (columnNumber - 1) % 26;
    letter = String.fromCharCode(65 + modulo) + letter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return letter;
}

function makeWorkShiftRow(row, startTime, endTime) {
  // --- user settings ---
  const startColumn = "K";
  // --- user settings ---
  
  const posA = 65;
  const codeStartColumn = startColumn.charCodeAt(0)-posA;

  const props = PropertiesService.getScriptProperties();
  const color = props.getProperty("BAR_COLOR");

  let sheet = SpreadsheetApp.getActiveSheet();

  if (startTime != "" && startTime != null && endTime != "" && endTime != null) {
    const startHH = startTime.getHours();
    const startMM = startTime.getMinutes();
    const endHH   = endTime.getHours();
    const endMM   = endTime.getMinutes();

    const indexStart = (startHH*2 + startMM/30) + codeStartColumn + 1;
    const colStart = columnNumberToLetter(indexStart);

    const indexEnd   = (endHH*2   + endMM/30)  + codeStartColumn;
    const colEnd = columnNumberToLetter(indexEnd);

    // Logger.log(`${row}行目 ${startHH}:${startMM}~${endHH}:${endMM}`);
    // Logger.log(`時刻インデックス ${indexStart}~${indexEnd}`)
    // Logger.log(`背景色塗りつぶし ${colStart}${row}:${colEnd}${row}`);

    let range = sheet.getRange(`${colStart}${row}:${colEnd}${row}`);
    range.setBackground(color)
  } else {
    sheet.hideRow(sheet.getRange(`A${row}`));
  }    
}

function exportDailyWorkShift(date) {
  var newSheetName = date.getDate();

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheetName = "0";

  if (sheetExists(newSheetName)) {
    Logger.log(`シート ${newSheetName} は既に存在しています。処理を中止します。`);
    return;
  }

  var templateSheet = spreadsheet.getSheetByName(templateSheetName);
  if (!templateSheet) {
    throw new Error(`テンプレートシート ${templateSheetName} が存在しません。`);
  }
  var newSheet = templateSheet.copyTo(spreadsheet);
  newSheet.setName(newSheetName);

  if (date instanceof Date) {
    var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd");
    newSheet.getRange("C1").setValue(formattedDate);
  } else {
    newSheet.getRange("C1").setValue(date);
  }

  // ✅ シフト表の作成（仮置き）

  Logger.log(`${date} 分のシート 「${newSheetName}」 を作成しました。`);
}

function exportMonthlyWorkShift() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheetName = "0";
  var sheet = spreadsheet.getSheetByName(templateSheetName);
  var range = sheet.getRange("C1");

  var startOfMonth = range.getValue();
  var endOfLastMonth = new Date(startOfMonth);
  endOfLastMonth.setDate(startOfMonth.getDate()-1);

  for (let i = 1; i <= 31; i++) {
    var date = new Date(endOfLastMonth);
    date.setDate(endOfLastMonth.getDate()+i);
    exportDailyWorkShift(date)
  }
}

function removeMonthlyWorkShift() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  for (let i = 1; i <= 31; i++) {
    var sheetName = i;
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      spreadsheet.deleteSheet(sheet);
      Logger.log(`シート 「${sheetName}」 を作成しました。`);
    }
  }
}

function sheetExists(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(sheetName) !== null;
}

function checkDayFormat(sheetName) {
  // var regex = /^Day([1-9]\d*)$/;
  var regex = /^([0-9]\d*)$/;
  var match = sheetName.match(regex);
  
  if (match) {
    Logger.log(sheetName + " is valid, number = " + match[1]);
    return true;
  } else {
    Logger.log(sheetName + " is invalid");
    return false;
  }
}

function getActiveSheetName() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var sheetName = sheet.getName();
  return sheetName;
}