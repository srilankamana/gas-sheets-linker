/**
 * ===================================================================================
 * グローバル設定（定数）
 * 実行前に、ここのシート名や列番号を実際の環境に合わせてください。
 * ===================================================================================
 */

// メインシート（このシートで編集をトリガーする）
const MAIN_SHEET_NAME = '推薦情報一覧';   // Q列を編集するシートの名前
const TRIGGER_COLUMN = 17;              // トリガーとなる列（Q列）
const URL_COLUMN_MAIN = 2;              // メインシートのURL列（B列）
const NAME_COLUMN_MAIN = 9;             // メインシートの候補者名列（I列）
const SSID_COLUMN_MAIN = 11;            // メインシートの連携先ID列（K列）

// 連携先シート（ステータスを書き込むシート）
const TARGET_SHEET_NAME = 'A社推薦管理'; // 連携先のシート名
const URL_COLUMN_TARGET = 4;           // 連携先のURL列（D列） ※配列インデックスは -1
const NAME_COLUMN_TARGET = 11;         // 連携先の候補者名列（K列） ※配列インデックスは -1
const STATUS_COLUMN_TARGET = 1;        // 連携先のステータス列（A列）
const COMPLETED_STATUS_TEXT = '推薦済';  // 書き込むステータスのテキスト


/**
 * ===================================================================================
 * メインロジック
 * ===================================================================================
 */

/**
 * 編集イベントのトリガー関数（インストール型）
 * 'onEdit'という名前ではなく、トリガーに設定するために 'myOnEditTrigger' などの
 * 固有の名前にしています。
 * @param {Object} e - Google Sheets が発行するイベントオブジェクト
 */
function myOnEditTrigger(e) {
  // eオブジェクトがない場合（デバッグ実行など）は何もしない
  if (!e) return;

  const sheet = e.source.getActiveSheet();
  const range = e.range;

  // 1. 対象シートのチェック
  if (sheet.getName() !== MAIN_SHEET_NAME) return;

  // 2. ヘッダー行（1行目）の編集は無視
  const row = range.getRow();
  if (row < 2) return;

  // 3. トリガーとなる列（Q列）以外の編集は無視
  const col = range.getColumn();
  if (col !== TRIGGER_COLUMN) return;

  // 4. Q列が空欄にされた場合は何もしない
  const newValue = range.getValue();
  if (!newValue) return;

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
  const recUrl = sheet.getRange(row, URL_COLUMN_MAIN).getValue();
  const candidateName = sheet.getRange(row, NAME_COLUMN_MAIN).getValue();
  const ssInfo = sheet.getRange(row, SSID_COLUMN_MAIN).getValue();

  // キーとなる情報（URL, 氏名, 連携先ID）が空欄なら停止
  if (!recUrl || !candidateName || !ssInfo) {
    Logger.log('処理停止: B列(URL), I列(氏名), K列(SSID) のいずれかが空欄です');
    return;
  }

  // --- 2. 連携先スプレッドシートを開く ---
  const ssId = getSpreadsheetIdFromUrl(ssInfo);
  if (!ssId) {
    Logger.log('処理停止: K列からスプレッドシートIDを抽出できませんでした。');
    return;
  }

  let companySS;
  try {
    // セキュリティ制限のため、この操作にはインストール型トリガーが必須
    companySS = SpreadsheetApp.openById(ssId);
  } catch (e) {
    Logger.log('スプレッドシートが開けません ID: ' + ssId + ' エラー: ' + e.message);
    return;
  }

  // --- 3. 連携先シートを取得 ---
  const companySheet = companySS.getSheetByName(TARGET_SHEET_NAME);
  if (!companySheet) {
    Logger.log(TARGET_SHEET_NAME + ' シートが見つかりません: ' + ssId);
    return;
  }

  const lastRow = companySheet.getLastRow();
  if (lastRow < 2) return; // データ行なし

  const lastCol = companySheet.getLastColumn();

  // --- 4. 連携先シートの全データを走査して照合 ---
  const values = companySheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // 照合キーを準備 (前後の空白などを除去)
  const baseUrl = String(recUrl).trim();
  const baseName = String(candidateName).trim();

  let targetRow = null;

  // getValues()で取得した配列は0から始まるため、列番号から-1したインデックスでアクセス
  const urlIndexTarget = URL_COLUMN_TARGET - 1;
  const nameIndexTarget = NAME_COLUMN_TARGET - 1;

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

  // --- 5. 一致した場合の処理 ---
  if (!targetRow) {
    Logger.log('一致する行が見つかりませんでした: [氏名]' + baseName + ' / [URL]' + baseUrl);
    return;
  }

  // 5a. ステータス列に「推薦済」と書き込む
  companySheet.getRange(targetRow, STATUS_COLUMN_TARGET).setValue(COMPLETED_STATUS_TEXT);
 
  // 5b. スプレッドシートに変更を即時反映させる (フィルタ適用のタイミング問題対策)
  SpreadsheetApp.flush();
 
  // 5c. A列が空欄のみ表示するフィルタを再適用
  applyFilterToHideCompleted(companySheet);
}

/**
 * ===================================================================================
 * ユーティリティ（補助）関数
 * ===================================================================================
 */

/**
 * 完了した行を非表示にするため、フィルタを再適用する関数
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 */
function applyFilterToHideCompleted(sheet) {
  // この「削除」→「再作成」のシンプルなロジックが、
  // Google Sheetsのタイミング問題を回避するのに有効だった。
  
  // 1. 既存のフィルタを削除
  const existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return; // シートが空なら何もしない

  // 2. フィルタをシート全体に再作成
  sheet.getRange(1, 1, lastRow, lastCol).createFilter();

  // 3. フィルタに「ステータス列が空欄のセルのみ表示」のルールを適用
  sheet.getFilter().setColumnFilterCriteria(
    STATUS_COLUMN_TARGET, // A列
    SpreadsheetApp.newFilterCriteria()
      .whenCellEmpty()
      .build()
  );
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
    Logger.log(MAIN_SHEET_NAME + ' が見つかりません');
    return;
  }
  const row = sheet.getActiveCell().getRow();
  if (row < 2) {
    Logger.log('データ行（2行目以降）を選択してください');
    return;
  }
  syncRecommendationStatus(sheet, row);
}