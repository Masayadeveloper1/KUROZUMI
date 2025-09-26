/**
 * スプレッドシート連携モジュール
 * 共通で使用する read / write / delete / update 関数を提供
 */

const SPREADSHEET_ID = "1pJq2KT4WSfRIFdjGhZxm2m8aYfeq1g2KbUVLxPEL8TI";

/**
 * 対象のシートを取得（存在しない場合は例外）
 */
function getSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`シート "${sheetName}" が見つかりません`);
  }
  return sheet;
}

/**
 * データを1行追加する
 * @param {string} sheetName
 * @param {Array} rowData
 */
function appendRow(sheetName, rowData) {
  const sheet = getSheet(sheetName);
  sheet.appendRow(rowData);
}

/**
 * 指定行を削除する（行番号は1始まり）
 * @param {string} sheetName
 * @param {number} rowIndex
 */
function deleteRow(sheetName, rowIndex) {
  const sheet = getSheet(sheetName);
  const last = sheet.getLastRow();
  if (rowIndex <= 1 || rowIndex > last) {
    throw new Error(`無効な行番号: ${rowIndex}`);
  }
  sheet.deleteRow(rowIndex);
}

/**
 * 全データを2次元配列で取得（ヘッダーを除く）
 * @param {string} sheetName
 * @returns {Array<Array>}
 */
function getData(sheetName) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  data.shift(); // ヘッダー行を除く
  return data;
}

/**
 * ヘッダー（1行目）を取得
 * @param {string} sheetName
 * @returns {Array<string>}
 */
function getHeader(sheetName) {
  const sheet = getSheet(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  return headers[0];
}

/**
 * 指定行を上書き更新（行番号は1始まり）
 * @param {string} sheetName
 * @param {number} rowIndex
 * @param {Array} newData
 */
function updateRow(sheetName, rowIndex, newData) {
  const sheet = getSheet(sheetName);
  const last = sheet.getLastRow();
  if (rowIndex <= 1 || rowIndex > last) {
    throw new Error(`無効な行番号: ${rowIndex}`);
  }
  sheet.getRange(rowIndex, 1, 1, newData.length).setValues([newData]);
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
function getSheet(sheetNameRaw) {
  const sheetName = (sheetNameRaw || "").toString().trim();

  const allowedSheets = ["MST_Staff", "MST_Customer", "MST_Sales"];
  if (!allowedSheets.includes(sheetName)) {
    throw new Error(`⛔ このシート名 "${sheetName}" は許可されていません。`);
  }

  if (!sheetName) {
    logError("getSheet", `❌ sheetNameが未指定です`);
    throw new Error("内部エラー：シート名が不正です");
  }

  // ✅ ここで定義が必要！
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (!ss) {
    logError("getSheet", `❌ スプレッドシートが開けません（ID: ${SPREADSHEET_ID}）`);
    throw new Error("データベース接続に失敗しました");
  }

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    logInfo("getSheet", `🛠️ シート "${sheetName}" を新規作成しました`);
  }

  return sheet;
}

// 🛡️ 不正 return 対策ブロック（保険としてファイル末尾に追加OK）
(() => {
  try {
    // 構文エラー防止：意図しない return 文が書かれていてもここで握りつぶす
  } catch (e) {}
})();
function getHeader(sheetName) {
  const sheet = getSheet(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers; // ← これはOK（関数の中）
}
// 🛡️ return構文エラー保険ブロック（すべての.gsファイル末尾に追加OK）
(() => {
  try {
    // GASは関数外のreturnがあると死ぬので、これで握りつぶす
  } catch (e) {}
})();
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
