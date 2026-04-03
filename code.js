// ══════════════════════════════════
// Jim 計分板 - Google Apps Script
// ══════════════════════════════════

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  // 允許跨來源請求
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  const action = e.parameter.action;
  let result;

  try {
    switch (action) {
      case 'getScore':
        result = getScore();
        break;
      case 'addRecord':
        result = addRecord(e.parameter);
        break;
      case 'getRecords':
        result = getRecords();
        break;
      case 'getBehaviors':
        result = getBehaviors();
        break;
      case 'getRewards':
        result = getRewards();
        break;
      case 'redeemReward':
        result = redeemReward(e.parameter);
        break;
      case 'checkPin':
        result = checkPin(e.parameter.pin);
        break;
      case 'getSettings':
        result = getSettings();
        break;
      case 'updateBehavior':
        result = updateBehavior(e.parameter);
        break;
      case 'addBehavior':
        result = addBehavior(e.parameter);
        break;
      case 'deleteBehavior':
        result = deleteBehavior(e.parameter);
        break;
      case 'updateReward':
        result = updateReward(e.parameter);
        break;
      case 'addReward':
        result = addReward(e.parameter);
        break;
      case 'updatePin':
        result = updatePin(e.parameter.newPin);
        break;
      case 'adjustScore':
        result = adjustScore(e.parameter.score);
        break;
      default:
        result = { success: false, message: '未知指令：' + action };
    }
  } catch (err) {
    result = { success: false, message: err.toString() };
  }

  output.setContent(JSON.stringify(result));
  return output;
}

// ── 取得總分 ──
function getScore() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');
  const score = Number(sheet.getRange('B3').getValue());
  return { success: true, score: score };
}

// ── 新增紀錄（修正版）──
function addRecord(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recordSheet = ss.getSheetByName('紀錄');
  const settingSheet = ss.getSheetByName('設定');

  // 確保分數是數字
  const scoreChange = Number(params.score);
  if (isNaN(scoreChange)) {
    return { success: false, message: '分數格式錯誤：' + params.score };
  }

  // 取得目前總分
  const currentScore = Number(settingSheet.getRange('B3').getValue()) || 0;
  const newScore = currentScore + scoreChange;

  // 寫入紀錄
  const now = new Date();
  const timeStr = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd HH:mm');
  recordSheet.appendRow([
    timeStr,
    String(params.name || ''),
    scoreChange,
    String(params.type || ''),
    String(params.note || '')
  ]);

  // 更新總分
  settingSheet.getRange('B3').setValue(newScore);

  // 強制寫入（避免快取問題）
  SpreadsheetApp.flush();

  return { success: true, newScore: newScore };
}

// ── 取得最近紀錄 ──
function getRecords() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('紀錄');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, records: [] };

  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const records = [];

  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0]) {
      records.push({
        time: data[i][0],
        name: data[i][1],
        score: Number(data[i][2]),
        type: data[i][3],
        note: data[i][4]
      });
    }
    if (records.length >= 30) break;
  }

  return { success: true, records: records };
}

// ── 取得行為項目 ──
function getBehaviors() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('行為項目');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, behaviors: [] };

  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const behaviors = [];

  for (let i = 0; i < data.length; i++) {
    if (data[i][0]) {
      behaviors.push({
        index: i + 2, // 對應到實際的 row 號碼
        name: data[i][0],
        score: Number(data[i][1]),
        type: data[i][2]
      });
    }
  }

  return { success: true, behaviors: behaviors };
}

// ── 取得獎勵項目 ──
function getRewards() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('獎勵項目');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, rewards: [] };

  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const rewards = [];

  for (let i = 0; i < data.length; i++) {
    if (data[i][0]) {
      rewards.push({
        index: i + 2,
        name: data[i][0],
        score: Number(data[i][1]),
        status: data[i][2]
      });
    }
  }

  return { success: true, rewards: rewards };
}

// ── 兌換獎勵 ──
function redeemReward(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingSheet = ss.getSheetByName('設定');
  const recordSheet = ss.getSheetByName('紀錄');

  const currentScore = Number(settingSheet.getRange('B3').getValue()) || 0;
  const cost = Number(params.cost);

  if (currentScore < cost) {
    return { success: false, message: '分數不足！還差 ' + (cost - currentScore) + ' 分' };
  }

  const newScore = currentScore - cost;
  settingSheet.getRange('B3').setValue(newScore);

  const now = new Date();
  const timeStr = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd HH:mm');
  recordSheet.appendRow([timeStr, '🎁 兌換：' + params.name, -cost, '兌換', '']);

  SpreadsheetApp.flush();

  return { success: true, newScore: newScore };
}

// ── 驗證 PIN ──
function checkPin(pin) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');
  const savedPin = String(sheet.getRange('B2').getValue());
  return { success: String(pin) === savedPin };
}

// ── 取得設定 ──
function getSettings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');
  return {
    success: true,
    pin: String(sheet.getRange('B2').getValue()),
    score: Number(sheet.getRange('B3').getValue())
  };
}

// ── 更新行為項目 ──
function updateBehavior(params) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('行為項目');
  const row = parseInt(params.index);
  sheet.getRange(row, 1).setValue(String(params.name));
  sheet.getRange(row, 2).setValue(Number(params.score));
  sheet.getRange(row, 3).setValue(String(params.type));
  SpreadsheetApp.flush();
  return { success: true };
}

// ── 新增行為項目 ──
function addBehavior(params) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('行為項目');
  sheet.appendRow([String(params.name), Number(params.score), String(params.type)]);
  SpreadsheetApp.flush();
  return { success: true };
}

// ── 刪除行為項目 ──
function deleteBehavior(params) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('行為項目');
  sheet.deleteRow(parseInt(params.index));
  SpreadsheetApp.flush();
  return { success: true };
}

// ── 更新獎勵項目 ──
function updateReward(params) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('獎勵項目');
  const row = parseInt(params.index);
  sheet.getRange(row, 1).setValue(String(params.name));
  sheet.getRange(row, 2).setValue(Number(params.score));
  sheet.getRange(row, 3).setValue(String(params.status));
  SpreadsheetApp.flush();
  return { success: true };
}

// ── 新增獎勵項目 ──
function addReward(params) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('獎勵項目');
  sheet.appendRow([String(params.name), Number(params.score), '上架']);
  SpreadsheetApp.flush();
  return { success: true };
}

// ── 更新 PIN ──
function updatePin(newPin) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');
  sheet.getRange('B2').setValue(String(newPin));
  SpreadsheetApp.flush();
  return { success: true };
}

// ── 直接調整總分 ──
function adjustScore(score) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');
  sheet.getRange('B3').setValue(Number(score));
  SpreadsheetApp.flush();
  return { success: true };
}
