// ✅ ログ関数（開発＆本番用に使用）
function logError(context, message) {
  console.error(`[${context}] ❌ ${message}`);
}
function logInfo(context, message) {
  console.log(`[${context}] ✅ ${message}`);
}

// ✅ 安全・統一的なシート取得関数
function getSheet(sheetNameRaw) {
  const sheetName = (sheetNameRaw || "").toString().trim();

  if (!sheetName) {
    logError("getSheet", `❌ シート名が未指定です`);
    throw new Error("内部エラー：連携対象のシートが正しく設定されていません。");
  }

  let ss;
  try {
    ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); // ← config.gsから取得
  } catch (e) {
    logError("getSheet", `❌ スプレッドシートの取得に失敗：${e.message}`);
    throw new Error("データベース接続に失敗しました。");
  }

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    const existing = ss.getSheets().map(s => s.getName()).join(", ");
    logError("getSheet", `❌ シート "${sheetName}" は存在しません。候補: ${existing}`);
    throw new Error(`「${sheetName}」というシートは存在しません。`);
  }

  return sheet;
}

// ✅ シートがなければ自動作成（主にログ用途）
function getOrCreateLogSheet(sheetName) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    logInfo("utils", `🛠️ ログ用シート "${sheetName}" を新規作成`);
  }

  return sheet;
}

// ✅ 安全策：CONFIG が未定義な場合のみ定義（最終保険）
if (typeof CONFIG === "undefined") {
  const CONFIG = {
    SPREADSHEET_ID: "1pJq2KT4WSfRIFdjGhZxm2m8aYfeq1g2KbUVLxPEL8TI"
  };
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
