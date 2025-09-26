function doGet(e) {
  const page = e.parameter.page || 'Start';
  return HtmlService.createTemplateFromFile(page).evaluate()
    .setTitle("業務支援ランチャー")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
/**
 * GAS のメインエントリーポイント
 * URLパラメータ "page" によって表示ページを切り替える
 */
function doGet(e) {
  const page = (e.parameter.page || "index").trim();
  const html = renderPage(page);
  return html.setTitle("業務支援アプリ").addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * HTMLファイルをテンプレートとして読み込む
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
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
function getSheet(sheetNameRaw) {
  const sheetName = (sheetNameRaw || "").toString().trim();

  if (!sheetName) {
    const msg = `❌ シート名が未指定です（sheetName="${sheetNameRaw}"）`;
    logError("getSheet", msg);
    throw new Error("内部エラー：シート名が不正です。");
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (!ss) {
    const msg = `❌ スプレッドシートが見つかりません（ID: ${SPREADSHEET_ID}）`;
    logError("getSheet", msg);
    throw new Error("データベースに接続できません。管理者に連絡してください。");
  }

  let sheet = ss.getSheetByName(sheetName);

  // ✅ 自動作成オプション（任意）
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    logInfo("getSheet", `🛠️ シート "${sheetName}" を新規作成しました`);
  }

 sheet;
}
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

   sheet;
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
