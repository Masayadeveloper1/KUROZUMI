function submitStaffForm(data) {
  const sheet = SpreadsheetApp.openById("1pJq2KT4WSfRIFdjGhZxm2m8aYfeq1g2KbUVLxPEL8TI").getSheetByName("MST_Staff");
  if (!sheet) throw new Error("MST_Staff シートが見つかりません");

  const row = [
    data.name,
    data.department,
    data.position,
    data.phone,
    data.email,
    data.startDate,
    data.note
  ];

  sheet.appendRow(row);
}
// スタッフ一覧を取得
function getStaffList() {
  const sheet = SpreadsheetApp.openById("1pJq2KT4WSfRIFdjGhZxm2m8aYfeq1g2KbUVLxPEL8TI").getSheetByName("MST_Staff");
  if (!sheet) throw new Error("MST_Staff シートが見つかりません");

  const data = sheet.getDataRange().getValues();
  data.shift(); // ヘッダー削除
  return data;
}

// 指定行を削除（index: 2行目以降）
function deleteStaffByIndex(rowIndex) {
  const sheet = SpreadsheetApp.openById("1pJq2KT4WSfRIFdjGhZxm2m8aYfeq1g2KbUVLxPEL8TI").getSheetByName("MST_Staff");
  if (!sheet) throw new Error("MST_Staff シートが見つかりません");
  
  const lastRow = sheet.getLastRow();
  if (rowIndex <= 1 || rowIndex > lastRow) throw new Error("無効な行番号です");

  sheet.deleteRow(rowIndex);
}
/**
 * フォーム送信・一覧取得・削除処理
 * 各HTML画面から呼び出される窓口
 */

/**
 * データを登録（1行追加）
 * @param {string} sheetName
 * @param {Object} formData - キー:カラム名, 値:入力値
 * @returns {string} 成功メッセージ or エラー
 */
function submitForm(sheetName, formData) {
  try {
    const headers = getHeader(sheetName); // ['氏名', '所属', ...]
    const rowData = headers.map(key => formData[key] || ""); // ヘッダー順に並び替え
    appendRow(sheetName, rowData);
    return "✅ 登録が完了しました。";
  } catch (err) {
    return `❌ エラー: ${err.message}`;
  }
}

/**
 * データを取得（一覧表示用）
 * @param {string} sheetName
 * @returns {Array<Array>} - 登録データ（2次元配列）
 */
function getGenericData(sheetName) {
  try {
    return getData(sheetName);
  } catch (err) {
    return [["エラー", err.message]];
  }
}

/**
 * 指定行を削除
 * @param {string} sheetName
 * @param {number} rowIndex - 実データの行番号（1行目はヘッダー）
 */
function deleteGenericRow(sheetName, rowIndex) {
  try {
    deleteRow(sheetName, rowIndex + 1); // 表示上は0始まりなので +1 +1
    return "✅ 削除が完了しました。";
  } catch (err) {
    return `❌ 削除失敗: ${err.message}`;
  }
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
function getSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    logInfo("getSheet", `🛠️ シート "${sheetName}" を新規作成しました`);
  }
  return sheet;
}

// 最後の有効な関数
function submitGenericForm(sheetName, formData) {
  const headers = getHeader(sheetName);
  const row = headers.map(key => formData[key] || "");
  appendRow(sheetName, row);
  logInfo(sheetName, `登録成功: ${JSON.stringify(formData)}`);
  return "✅ 登録が完了しました。";
}

// 🛡️ return構文エラー保険
(() => {
  try {
    // 保険。何もしない。
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
