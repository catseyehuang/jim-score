const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const action = e.parameter.action;
  let result;

  try {
    switch(action) {
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
        result = { success: false, message: '未知指令' };
    }
  } catch(err) {
    result = { success: false, message: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getScore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('設定');
  const score = sheet.getRange('B3').getValue();
  return { success: true, score: score };
}

function addRecord(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recordSheet = ss.getSheetByName('紀錄');
  const settingSheet = ss.getSheetByName('設定');

  const score = parseInt(params.score);
  const now = new Date();
  const timeStr = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd HH:mm');

  recordSheet.appendRow([
    timeStr,
    params.name,
    score,
    params.type,
    params.note || ''
  ]);

  const currentScore = settingSheet.getRange('B3').getValue();
  settingSheet.getRange('B3').setValue(currentScore + score);

  return { success: true, newScore: currentScore + score };
}

function getRecords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('紀錄');
  const data = sheet.getDataRange().getValues();
  const records = [];

  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0]) {
      records.push({
        time: data[i][0],
        name: data[i][1],
        score: data[i][2],
        type: data[i][3],
        note: data[i][4]
      });
    }
    if (records.length >= 30) break;
  }

  return { success: true, records: records };
}

function getBehaviors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('行為項目');
  const data = sheet.getDataRange().getValues();
  const behaviors = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      behaviors.push({
        index: i,
        name: data[i][0],
        score: data[i][1],
        type: data[i][2]
      });
    }
  }

  return { success: true, behaviors: behaviors };
}

function getRewards() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('獎勵項目');
  const data = sheet.getDataRange().getValues();
  const rewards = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      rewards.push({
        index: i,
        name: data[i][0],
        score: data[i][1],
        status: data[i][2]
      });
    }
  }

  return { success: true, rewards: rewards };
}

function redeemReward(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingSheet = ss.getSheetByName('設定');
  const recordSheet = ss.getSheetByName('紀錄');

  const currentScore = settingSheet.getRange('B3').getValue();
  const cost = parseInt(params.cost);

  if (currentScore < cost) {
    return { success: false, message: '分數不足！' };
  }

  const newScore = currentScore - cost;
  settingSheet.getRange('B3').setValue(newScore);

  const now = new Date();
  const timeStr = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd HH:mm');
  recordSheet.appendRow([timeStr, '🎁 兌換：' + params.name, -cost, '兌換', '']);

  return { success: true, newScore: newScore };
}

function checkPin(pin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('設定');
  const savedPin = sheet.getRange('B2').getValue().toString();
  return { success: pin === savedPin };
}

function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('設定');
  return {
    success: true,
    pin: sheet.getRange('B2').getValue(),
    score: sheet.getRange('B3').getValue()
  };
}

function updateBehavior(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('行為項目');
  const row = parseInt(params.index) + 1;
  sheet.getRange(row, 1).setValue(params.name);
  sheet.getRange(row, 2).setValue(parseInt(params.score));
  sheet.getRange(row, 3).setValue(params.type);
  return { success: true };
}

function addBehavior(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('行為項目');
  sheet.appendRow([params.name, parseInt(params.score), params.type]);
  return { success: true };
}

function updateReward(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('獎勵項目');
  const row = parseInt(params.index) + 1;
  sheet.getRange(row, 1).setValue(params.name);
  sheet.getRange(row, 2).setValue(parseInt(params.score));
  sheet.getRange(row, 3).setValue(params.status);
  return { success: true };
}

function addReward(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('獎勵項目');
  sheet.appendRow([params.name, parseInt(params.score), '上架']);
  return { success: true };
}

function updatePin(newPin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('設定');
  sheet.getRange('B2').setValue(newPin);
  return { success: true };
}

function adjustScore(score) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('設定');
  sheet.getRange('B3').setValue(parseInt(score));
  return { success: true };
}
