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

function copyAllWorkerSheet() {
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
    copyWorkerSheet(values[i][0]);
  }

  const title = "出勤予定";
  const prompt = "出勤予定を入力してください";
  const response = Browser.msgBox(title, prompt, Browser.Buttons.YES_NO);

}

function copyWorkerSheet(newSheetName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  var templateSheetName = 'X0000:name';
  var templateSheet = spreadsheet.getSheetByName(templateSheetName);

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

function createWorkshift(date) {
  const title = "シフト表作成";
  const prompt = "シフト表を作成してもよろしいですか?";

  const response = Browser.msgBox(title, prompt, Browser.Buttons.YES_NO);
  if (response == "yes") {
    Browser.msgBox("test")
  }
}

function makeWorkShiftDiagram() {
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
  }    
}


function getActiveSheetName() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var sheetName = sheet.getName();
  return sheetName;
}