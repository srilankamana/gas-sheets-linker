/**
 * ===================================================================================
 * ★★★ グローバル設定 ★★★
 * シート名や列名を、環境に合わせて設定してください。
 * ===================================================================================
 */

// --- メインシート（推薦情報一覧） ---
const MAIN_SHEET_NAME = '推薦一覧';
const TRIGGER_COLUMN_NAME = '推薦済みフラグ'; // トリガーとなる列
const URL_COLUMN_MAIN_NAME = '推薦URL';      // メインシートのURL列
const NAME_COLUMN_MAIN_NAME = '求職者名';    // メインシートの候補者名列
const SSID_COLUMN_MAIN_NAME = 'スプシ';      // メインシートの連携先ID列

// ★追加: 処理済みかどうかを記録する列（メインシートに追加してください）
const SYNC_STATUS_COLUMN_NAME = '連携ステータス'; 


// --- 連携先シート（A社推薦管理） ---
const TARGET_SHEET_NAME = 'A社推薦管理';
const URL_COLUMN_TARGET_NAME = '推薦URL（自動入力）'; 
const NAME_COLUMN_TARGET_NAME = '候補者名'; 
const STATUS_COLUMN_TARGET_NAME = '推薦ステータス'; 

// --- 書き込むステータスのテキスト ---
const COMPLETED_STATUS_TEXT = '推薦済';


/**
 * ===================================================================================
 * トリガー関数
 * ===================================================================================
 */

/**
 * 【手動編集用】編集イベントのトリガー関数
 * 人間が手入力した時にも、即座に反応するように残しておきます。
 */
function myOnEditTrigger(e) {
  if (!e) return;
  const sheet = e.source.getActiveSheet();
  
  // シート名とトリガー列のチェック
  if (sheet.getName() !== MAIN_SHEET_NAME) return;
  if (e.range.getRow() < 2) return;

  try {
    const colMap = getColumnNumbersByName(sheet, [TRIGGER_COLUMN_NAME]);
    if (e.range.getColumn() !== colMap[TRIGGER_COLUMN_NAME]) return;
    
    // 値が入力されたら処理実行
    if (e.range.getValue()) {
      syncRow(sheet, e.range.getRow());
    }
  } catch (err) {
    console.error(err.message);
  }
}

/**
 * 【自動連携用】時間主導型トリガー関数
 * これを「トリガー」設定で「10分おき」などに設定します。
 * 未処理の行（フラグがあるのに、連携ステータスが空の行）をまとめて処理します。
 */
function processUnsyncedRows() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) {
    console.error('シートが見つかりません: ' + MAIN_SHEET_NAME);
    return;
  }

  // 必要な列番号を取得
  let colMap;
  try {
    colMap = getColumnNumbersByName(sheet, [TRIGGER_COLUMN_NAME, SYNC_STATUS_COLUMN_NAME]);
  } catch (e) {
    console.error(e.message);
    return;
  }

  const triggerColIdx = colMap[TRIGGER_COLUMN_NAME] - 1; // 配列用インデックス
  const statusColIdx = colMap[SYNC_STATUS_COLUMN_NAME] - 1;
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // 全データを取得してメモリ上でチェック（高速化）
  const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  const values = dataRange.getValues();
  const statusColNum = colMap[SYNC_STATUS_COLUMN_NAME];

  // 1行ずつチェック
  for (let i = 0; i < values.length; i++) {
    const rowNum = i + 2;
    const triggerVal = values[i][triggerColIdx]; // 推薦フラグ
    const statusVal = values[i][statusColIdx];   // 連携ステータス

    // 「推薦フラグがあり」かつ「まだ連携済になっていない」行だけを処理
    if (triggerVal && statusVal !== '連携済') {
      console.log('未連携の行を検出しました: ' + rowNum + '行目');
      
      // 同期処理を実行
      const success = syncRecommendationStatus(sheet, rowNum);
      
      // 成功したら、メインシートに「連携済」とハンコを押す
      if (success) {
        sheet.getRange(rowNum, statusColNum).setValue('連携済');
        SpreadsheetApp.flush(); // 確実に保存
      }
    }
  }
}

/**
 * 1行分の処理を行うラッパー関数（手動トリガー用）
 */
function syncRow(sheet, row) {
  let colMap;
  try {
    colMap = getColumnNumbersByName(sheet, [SYNC_STATUS_COLUMN_NAME]);
    const success = syncRecommendationStatus(sheet, row);
    if (success) {
       sheet.getRange(row, colMap[SYNC_STATUS_COLUMN_NAME]).setValue('連携済');
    }
  } catch (e) {
    console.error(e.message);
  }
}

/**
 * ===================================================================================
 * メインロジック
 * ===================================================================================
 */

/**
 * 2つのシートのステータスを同期するメイン関数
 * @returns {boolean} 成功したら true を返す
 */
function syncRecommendationStatus(sheet, row) {
  let mainColMap;
  try {
    mainColMap = getColumnNumbersByName(sheet, [URL_COLUMN_MAIN_NAME, NAME_COLUMN_MAIN_NAME, SSID_COLUMN_MAIN_NAME]);
  } catch(err) {
    console.error('処理停止: ' + err.message);
    return false;
  }

  const recUrl = sheet.getRange(row, mainColMap[URL_COLUMN_MAIN_NAME]).getValue();
  const candidateName = sheet.getRange(row, mainColMap[NAME_COLUMN_MAIN_NAME]).getValue();
  const ssInfo = sheet.getRange(row, mainColMap[SSID_COLUMN_MAIN_NAME]).getValue();

  if (!recUrl || !candidateName || !ssInfo) {
    console.warn('処理停止: 必須情報が空欄です (行: ' + row + ')');
    return false;
  }

  const ssId = getSpreadsheetIdFromUrl(ssInfo);
  if (!ssId) return false;

  let companySS;
  try {
    companySS = SpreadsheetApp.openById(ssId);
  } catch (e) {
    console.error('スプレッドシートが開けません: ' + e.message);
    return false;
  }

  const companySheet = companySS.getSheetByName(TARGET_SHEET_NAME);
  if (!companySheet) {
    console.error(TARGET_SHEET_NAME + ' が見つかりません');
    return false;
  }

  // 連携先シートの列特定
  let targetColMap;
  try {
    targetColMap = getColumnNumbersByName(companySheet, [URL_COLUMN_TARGET_NAME, NAME_COLUMN_TARGET_NAME, STATUS_COLUMN_TARGET_NAME]);
  } catch(err) {
    console.error('連携先シートエラー: ' + err.message);
    return false;
  }

  const lastRow = companySheet.getLastRow();
  if (lastRow < 2) return false;

  const values = companySheet.getRange(2, 1, lastRow - 1, companySheet.getLastColumn()).getValues();
  const baseUrl = String(recUrl).trim();
  const baseName = String(candidateName).trim();
  
  const urlIdx = targetColMap[URL_COLUMN_TARGET_NAME] - 1;
  const nameIdx = targetColMap[NAME_COLUMN_TARGET_NAME] - 1;
  let targetRow = null;

  for (let i = 0; i < values.length; i++) {
    const rowUrl = String(values[i][urlIdx] || '').trim();
    const rowName = String(values[i][nameIdx] || '').trim();

    if (rowUrl === baseUrl && rowName === baseName) {
      targetRow = i + 2;
      break;
    }
  }

  if (!targetRow) {
    console.log('一致する行が見つかりませんでした: ' + baseName);
    return false;
  }

  // 書き込みとフィルタ
  const statusColNum = targetColMap[STATUS_COLUMN_TARGET_NAME];
  companySheet.getRange(targetRow, statusColNum).setValue(COMPLETED_STATUS_TEXT);
  SpreadsheetApp.flush();
  applyFilterToHideCompleted(companySheet, statusColNum);
  
  console.log('同期成功: ' + baseName);
  return true; // 成功
}

/**
 * ===================================================================================
 * ユーティリティ
 * ===================================================================================
 */

function getColumnNumbersByName(sheet, colNames) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colMap = {};
  const notFound = [];
  for (const name of colNames) {
    const index = headers.indexOf(name);
    if (index !== -1) colMap[name] = index + 1;
    else notFound.push(name);
  }
  if (notFound.length > 0) throw new Error('列が見つかりません: ' + notFound.join(', '));
  return colMap;
}

function applyFilterToHideCompleted(sheet, statusColNum) {
  const criteria = SpreadsheetApp.newFilterCriteria().whenCellEmpty().build();
  const filter = sheet.getFilter();
  if (filter) {
    filter.removeColumnFilterCriteria(statusColNum);
    filter.setColumnFilterCriteria(statusColNum, criteria);
  } else {
    const dataRange = sheet.getDataRange();
    if (dataRange.getNumRows() > 1) {
      dataRange.createFilter().setColumnFilterCriteria(statusColNum, criteria);
    }
  }
  SpreadsheetApp.flush();
}

function getSpreadsheetIdFromUrl(value) {
  const str = String(value).trim();
  if (str.indexOf('/') === -1 && str.length > 10) return str;
  const match = str.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}
