/**
 * logger.gs
 * 操作・エラーなどのログを記録するユーティリティ
 */

const LOG_SHEET_NAME = "LOG_Operations";

/**
 * ログを記録する（汎用）
 * @param {string} type - 'INFO' | 'ERROR'
 * @param {string} category - 例: 'Staff', 'GPT', 'Form'
 * @param {string} message - 操作内容やエラー内容
 */
function logEvent(type, category, message) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    sheet.appendRow(["日時", "ユーザー", "タイプ", "カテゴリ", "内容"]);
  }

  const timestamp = new Date();
  const user = Session.getActiveUser().getEmail() || "匿名";
  sheet.appendRow([timestamp, user, type, category, message]);
}

/**
 * 成功操作ログを記録
 * @param {string} category - 処理対象（例: 'Staff', 'Customer'）
 * @param {string} message - 実行内容（例: '登録成功: 山田 太郎'）
 */
function logInfo(category, message) {
  logEvent("INFO", category, message);
}

/**
 * エラーログを記録
 * @param {string} category - 処理対象
 * @param {string} message - エラー内容
 */
function logError(category, message) {
  logEvent("ERROR", category, message);
}
/**
 * 高精度で完全防御型のシート取得関数
 * ・未定義チェック
 * ・シート存在チェック
 * ・シート名の正規化
 * ・エラー時ログ記録＋ユーザーフレンドリーな例外表示
 * ・今後の拡張にも対応
 */
function getSheet(sheetNameRaw) {
  const sheetName = (sheetNameRaw || "").toString().trim();

  // ✅ 空文字または未定義の場合
  if (!sheetName || sheetName.length === 0) {
    const msg = `❌ シート名が未指定です。sheetName="${sheetNameRaw}"`;
    logError("getSheet", msg);
    throw new Error("内部エラー：連携対象のシートが正しく設定されていません。");
  }

  // ✅ シートIDの取得（存在確認）
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (!ss) {
    const msg = `❌ スプレッドシートが見つかりません（ID: ${SPREADSHEET_ID}）`;
    logError("getSheet", msg);
    throw new Error("データベースとの接続に失敗しました。管理者に連絡してください。");
  }

  // ✅ シート取得
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    const existingSheets = ss.getSheets().map(s => s.getName()).join(", ");
    const msg = `❌ シート "${sheetName}" が見つかりません。\n存在するシート: [${existingSheets}]`;
    logError("getSheet", msg);
    throw new Error(`データシート「${sheetName}」が見つかりません。\n入力フォームが壊れている可能性があります。`);
  }

  // ✅ 正常時ログ（開発時用、コメントアウト可）
  // logInfo("getSheet", `✅ シート "${sheetName}" を取得しました`);

 sheet;
}
// シートが存在しない場合に自動作成する
function getOrCreateLogSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName); // ✅ 先に宣言しておく

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);      // ✅ 再代入なので let でOK
    logInfo("logger", `🛠️ ログ用シート "${sheetName}" を新規作成`);
  }

  return sheet;
}
// ✅ グローバル定数 CONFIG が未定義なら定義（重複防止）
if (typeof CONFIG === "undefined") {
  const CONFIG = {
    SPREADSHEET_ID: "1pJq2KT4WSfRIFdjGhZxm2m8aYfeq1g2KbUVLxPEL8TI"
  };
}

// ✅ logError が未定義なら安全版を定義（上書きしない）
if (typeof logError === "undefined") {
  function logError(context, message) {
    const safeContext = context || "⚠️ unknown";
    const safeMessage = message || "⚠️ エラー内容が未定義です";
    console.error(`[${safeContext}] ❌ ${safeMessage}`);
  }
}

// ✅ logInfo も定義されてなければ定義
if (typeof logInfo === "undefined") {
  function logInfo(context, message) {
    const safeContext = context || "ℹ️ unknown";
    const safeMessage = message || "ℹ️ 情報なし";
    console.log(`[${safeContext}] ✅ ${safeMessage}`);
  }
}

// ✅ Errorの拡張（将来用）：デフォルトメッセージ補完
function safeThrow(error, context) {
  try {
    const message = (error && error.message) ? error.message : "❌ 未知のエラーが発生しました";
    logError(context || "safeThrow", message);
    throw new Error(message);
  } catch (e) {
    logError("safeThrow", e.message || "❌ 再スロー中に不明なエラー");
    throw e;
  }
}

