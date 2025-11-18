/**
 * ===================================================================================
 * ★★★ グローバル設定（定数） ★★★
 * ここのシート名や「列名」を、あなたのスプレッドシートの1行目と
 * 一字一句同じになるように設定してください。
 * ===================================================================================
 */

// --- メインシート（このシートで編集をトリガーする） ---
const MAIN_SHEET_NAME = '推薦一覧';
const TRIGGER_COLUMN_NAME = '推薦済みフラグ'; // トリガーとなる列
const URL_COLUMN_MAIN_NAME = '推薦URL'; // メインシートのURL列
const NAME_COLUMN_MAIN_NAME = '求職者名'; // メインシートの候補者名列
const SSID_COLUMN_MAIN_NAME = 'スプシ'; // メインシートの連携先ID列


// --- 連携先シート（ステータスを書き込むシート） ---
const TARGET_SHEET_NAME = 'A社推薦管理';
const URL_COLUMN_TARGET_NAME = '推薦URL（自動入力）'; // 連携先のURL列
const NAME_COLUMN_TARGET_NAME = '候補者名'; // 連携先の候補者名列
const STATUS_COLUMN_TARGET_NAME = '推薦ステータス'; // 連携先のステータス列

// --- 書き込むステータスのテキスト ---
const COMPLETED_STATUS_TEXT = '推薦済';


/**
 * ===================================================================================
 * メインロジック
 * ===================================================================================
 */

/**
 * 編集イベントのトリガー関数（インストール型）
 * 手動実行時のガード処理や、シート名・列名のチェックを行います。
 * @param {Object} e - Google Sheets が発行するイベントオブジェクト
 */
function myOnEditTrigger(e) {
  // eオブジェクトがない場合（手動実行など）は何もしない
  if (!e) {
    return;
  }

  const sheet = e.source.getActiveSheet();
  const range = e.range;

  // 1. 対象シートのチェック
  if (sheet.getName() !== MAIN_SHEET_NAME) {
    return;
  }

  // 2. ヘッダー行（1行目）の編集は無視
  const row = range.getRow();
  if (row < 2) {
    return;
  }

  // 3. トリガーとなる列かを「列名」でチェック
  const col = range.getColumn();
  let triggerColNum = -1;
  try {
    // メインシートのヘッダーからトリガー列の列番号を取得
    const mainColMap = getColumnNumbersByName(sheet, [TRIGGER_COLUMN_NAME]);
    triggerColNum = mainColMap[TRIGGER_COLUMN_NAME];

  } catch (err) {
    // ヘッダーが見つからないエラーはログに残すが、処理は停止する
    console.error(err.message);
    return;
  }
  
  if (col !== triggerColNum) {
    return; // トリガー列以外の編集は無視
  }

  // 4. トリガー列が空欄にされた場合は何もしない
  const newValue = range.getValue();
  if (!newValue) {
    return;
  }

  // すべての条件をクリアした場合、実処理を実行
  syncRecommendationStatus(sheet, row);
}

/**
 * 2つのシートのステータスを同期するメイン関数
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - メインシートのオブジェクト
 * @param {number} row - 編集された行番号
 */
function syncRecommendationStatus(sheet, row) {
  
  // --- 1. メインシートから照合キーを取得 ---
  let mainColMap;
  try {
    // 必要な列名から列番号を一括取得
    const requiredCols = [URL_COLUMN_MAIN_NAME, NAME_COLUMN_MAIN_NAME, SSID_COLUMN_MAIN_NAME];
    mainColMap = getColumnNumbersByName(sheet, requiredCols);
  } catch(err) {
    console.error('処理停止(sync): ' + err.message);
    return;
  }

  const recUrl = sheet.getRange(row, mainColMap[URL_COLUMN_MAIN_NAME]).getValue();
  const candidateName = sheet.getRange(row, mainColMap[NAME_COLUMN_MAIN_NAME]).getValue();
  const ssInfo = sheet.getRange(row, mainColMap[SSID_COLUMN_MAIN_NAME]).getValue();

  // キーとなる情報（URL, 氏名, 連携先ID）が空欄なら停止
  if (!recUrl || !candidateName || !ssInfo) {
    console.warn('処理停止(sync): 必須情報（URL/氏名/連携先ID）のいずれかが空欄です');
    return;
  }

  // --- 2. 連携先スプレッドシートを開く ---
  const ssId = getSpreadsheetIdFromUrl(ssInfo);
  if (!ssId) {
    console.warn('処理停止(sync): スプレッドシートIDを抽出できませんでした。');
    return;
  }

  let companySS;
  try {
    // セキュリティ制限のため、この操作にはインストール型トリガーが必須
    companySS = SpreadsheetApp.openById(ssId);
  } catch (e) {
    console.error('スプレッドシートが開けません ID: ' + ssId + ' エラー: ' + e.message);
    return;
  }

  // --- 3. 連携先シートを取得 ---
  const companySheet = companySS.getSheetByName(TARGET_SHEET_NAME);
  if (!companySheet) {
    console.error(TARGET_SHEET_NAME + ' シートが見つかりません: ' + ssId);
    return;
  }

  const lastRow = companySheet.getLastRow();
  if (lastRow < 2) return; // データ行なし

  // --- 4. 連携先シートの列番号を取得 ---
  let targetColMap;
  try {
    const requiredCols = [URL_COLUMN_TARGET_NAME, NAME_COLUMN_TARGET_NAME, STATUS_COLUMN_TARGET_NAME];
    targetColMap = getColumnNumbersByName(companySheet, requiredCols);
  } catch(err) {
    console.error('処理停止(sync): 連携先シートでエラー ' + err.message);
    return;
  }

  // --- 5. 連携先シートの全データを走査して照合 ---
  const lastCol = companySheet.getLastColumn();
  const values = companySheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // 照合キーを準備 (前後の空白などを除去)
  const baseUrl = String(recUrl).trim();
  const baseName = String(candidateName).trim();

  let targetRow = null;

  // getValues()で取得した配列は0から始まるため、列番号から-1したインデックスでアクセス
  const urlIndexTarget = targetColMap[URL_COLUMN_TARGET_NAME] - 1;
  const nameIndexTarget = targetColMap[NAME_COLUMN_TARGET_NAME] - 1;

  // 全データをループ
  for (let i = 0; i < values.length; i++) {
    const companyUrl = values[i][urlIndexTarget];
    const companyName = values[i][nameIndexTarget];

    if (!companyUrl || !companyName) continue; // どちらかが空欄ならスキップ

    const rowUrl = String(companyUrl).trim();
    const rowName = String(companyName).trim();

    // ★ URLと氏名が両方一致したら
    if (rowUrl === baseUrl && rowName === baseName) {
      targetRow = i + 2; // 実際の行番号（2行目スタート）
      break; // 一致したらループを抜ける
    }
  }

  // --- 6. 一致した場合の処理 ---
  if (!targetRow) {
    console.log('一致する行が見つかりませんでした: [氏名]' + baseName + ' / [URL]' + baseUrl);
    return;
  }

  // 6a. ステータス列に「推薦済」と書き込む
  const statusColNum = targetColMap[STATUS_COLUMN_TARGET_NAME];
  companySheet.getRange(targetRow, statusColNum).setValue(COMPLETED_STATUS_TEXT);
  console.log(targetRow + '行目の ' + STATUS_COLUMN_TARGET_NAME + ' に ' + COMPLETED_STATUS_TEXT + ' と書き込みました。');
 
  // 6b. スプレッドシートに変更を即時反映させる
  SpreadsheetApp.flush();
 
  // 6c. A列が空欄のみ表示するフィルタを再適用
  applyFilterToHideCompleted(companySheet, statusColNum);
}

/**
 * ===================================================================================
 * ユーティリティ（補助）関数
 * ===================================================================================
 */

/**
 * シートの1行目（ヘッダー）を読み取り、指定された列名（複数）に対応する
 * 列番号（1から始まる）を、マップ（{列名: 列番号, ...}）として返す。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {string[]} colNames - 探したい列名の配列
 * @returns {Object} 列名:列番号 のマップ
 * @throws {Error} 必要な列名が見つからない場合にエラーを投げる
 */
function getColumnNumbersByName(sheet, colNames) {
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colMap = {};
  const notFound = [];

  // 必要な列名（colNames）ごとにループ
  for (const name of colNames) {
    const index = headerRow.indexOf(name);
    
    if (index !== -1) {
      // 見つかった場合（indexは0から始まるので、列番号は+1する）
      colMap[name] = index + 1;
    } else {
      // 見つからなかった場合
      notFound.push(name);
    }
  }

  // 1つでも見つからない列名があれば、エラーを投げて処理を停止
  if (notFound.length > 0) {
    throw new Error('シート "' + sheet.getName() + '" に必要な列ヘッダー "' + notFound.join(', ') + '" が見つかりません。');
  }

  return colMap;
}

/**
 * 完了した行を非表示にするため、フィルタを再適用する関数
 * 既存のフィルタがある場合はルールのみ更新し、表形式のシートでのエラーを回避します。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {number} statusColNum - ステータス列の列番号
 */
function applyFilterToHideCompleted(sheet, statusColNum) {
  // 1. フィルタの基準（A列が空欄）を作成
  const criteria = SpreadsheetApp.newFilterCriteria()
    .whenCellEmpty()
    .build();
  
  // 2. シートに既にあるフィルタを取得
  const filter = sheet.getFilter();
  
  if (filter) {
    // 3a. もしフィルタが「既にある」なら、ルール(基準)だけを更新
    // 念のため、対象列の古いルールをクリア
    filter.removeColumnFilterCriteria(statusColNum);
    // 新しいルールを適用
    filter.setColumnFilterCriteria(statusColNum, criteria);
    
  } else {
    // 3b. もしフィルタが「ない」なら、データ範囲に新しく作成する
    const dataRange = sheet.getDataRange(); // シート全体ではなく、データ範囲
    if (dataRange.getNumRows() > 1) { // データが1行以上ある場合
      const newFilter = dataRange.createFilter();
      newFilter.setColumnFilterCriteria(statusColNum, criteria);
    }
  }
  
  // 最終的な変更を反映
  SpreadsheetApp.flush();
}

/**
 * K列に入力されたURLまたはIDから、スプレッドシートIDのみを抽出する
 * @param {string} value - K列に入力された値
 * @returns {string|null} スプレッドシートID、またはnull
 */
function getSpreadsheetIdFromUrl(value) {
  const str = String(value).trim();

  // スラッシュ無し＆それなりの長さ → IDが直接入力されているとみなす
  if (str.indexOf('/') === -1 && str.length > 10) {
    return str;
  }

  // URL形式からIDを正規表現で抽出
  const match = str.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}

/**
 * デバッグ（テスト）用関数
 * メインシートで処理したい行をアクティブにしてからこの関数を実行すると、
 * トリガー設定なしで1行分だけ強制的に実行できる。
 */
function debugProcessActiveRow() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) {
    console.log(MAIN_SHEET_NAME + ' が見つかりません');
    return;
  }
  const row = sheet.getActiveCell().getRow();
  if (row < 2) {
    console.log('データ行（2行目以降）を選択してください');
    return;
  }
  
  console.log('--- デバッグ実行: ' + row + '行目を処理します ---');
  syncRecommendationStatus(sheet, row);
  console.log('--- デバッグ実行 完了 ---');
}
