/**
 * 📍 メインルーター関数 (?page=〇〇でテンプレート切替)
 */
function doGet(e) {
  const DEFAULT_PAGE = "start"; // 安全なデフォルトページ
  let page = DEFAULT_PAGE;

  try {
    if (e?.parameter?.page && typeof e.parameter.page === "string") {
      const trimmedPage = e.parameter.page.trim();

      if (isValidPageName(trimmedPage)) {
        page = trimmedPage;
      } else {
        logWarn("doGet", `⚠️ 無効なページ名が指定されました: "${trimmedPage}". デフォルトにフォールバック。`);
      }
    } else {
      logInfo("doGet", "📩 クエリ 'page' が未指定または不正。デフォルトにフォールバック。");
    }
  } catch (err) {
    safeThrow(err, "doGet");
  }

  return renderPage(page);
}

/**
 * 📄 指定ページのHTMLテンプレートを安全に返す
 * @param {string} pageName 
 * @returns {HtmlOutput}
 */
function renderPage(pageName) {
  try {
    if (!isValidPageName(pageName)) {
      throw new Error(`無効なページ名が指定されました: "${pageName}"`);
    }

    const htmlTemplate = HtmlService.createTemplateFromFile(pageName);
    return htmlTemplate.evaluate()
      .setTitle("アプリ")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  } catch (err) {
    logError("renderPage", `❌ ページ "${pageName}" の読み込みに失敗: ${err.message}`);
    return HtmlService.createHtmlOutput(`<h1>ページが見つかりません: ${sanitizeHtml(pageName)}</h1>`);
  }
}

/**
 * 🧩 HTMLテンプレート用 include ヘルパー
 * @param {string} filename 
 * @returns {string}
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (err) {
    logError("include", `❌ includeファイル "${filename}" が見つかりません: ${err.message}`);
    return `<!-- include失敗: ${sanitizeHtml(filename)} -->`;
  }
}

/**
 * ✅ 有効なページ名かどうかを検証
 * @param {string} name 
 * @returns {boolean}
 */
function isValidPageName(name) {
  const VALID_PAGE_REGEX = /^[a-zA-Z0-9_-]+$/;
  return typeof name === "string" && name.length > 0 && VALID_PAGE_REGEX.test(name);
}

/**
 * 🧼 危険な文字をサニタイズ（XSS予防）
 * @param {string} str 
 * @returns {string}
 */
function sanitizeHtml(str) {
  return String(str).replace(/[&<>"'`=\/]/g, function (s) {
    return {
      '&': "&amp;",
      '<': "&lt;",
      '>': "&gt;",
      '"': "&quot;",
      "'": "&#39;",
      '/': "&#x2F;",
      '=': "&#x3D;",
      '`': "&#x60;"
    }[s];
  });
}

/**
 * 🧱 グローバル設定定数
 */
const CONFIG = (() => {
  return {
    SPREADSHEET_ID: "1pJq2KT4WSfRIFdjGhZxm2m8aYfeq1g2KbUVLxPEL8TI",
  };
})();

/**
 * 📕 エラーログ（未定義なら定義）
 */
if (typeof logError === "undefined") {
  function logError(context, message) {
    const tag = `[${context || "unknown"}]`;
    console.error(`${tag} ❌ ${message || "エラー内容なし"}`);
  }
}

/**
 * 📗 情報ログ
 */
if (typeof logInfo === "undefined") {
  function logInfo(context, message) {
    const tag = `[${context || "unknown"}]`;
    console.log(`${tag} ✅ ${message || "情報なし"}`);
  }
}

/**
 * 📙 警告ログ
 */
if (typeof logWarn === "undefined") {
  function logWarn(context, message) {
    const tag = `[${context || "unknown"}]`;
    console.warn(`${tag} ⚠️ ${message || "警告内容なし"}`);
  }
}

/**
 * 🔒 安全なエラー再スロー
 */
if (typeof safeThrow === "undefined") {
  function safeThrow(error, context) {
    try {
      const msg = (error && error.message) ? error.message : "❌ 未知のエラーが発生";
      logError(context || "safeThrow", msg);
      throw new Error(msg);
    } catch (e) {
      logError("safeThrow", e.message || "再スロー中に不明なエラー");
      throw e;
    }
  }
}
/**
 * 🧱 CONFIGの安全な定義（他ファイルとの競合防止）
 */
(function defineGlobalConfig() {
  if (typeof globalThis.CONFIG === "undefined") {
    globalThis.CONFIG = {
      SPREADSHEET_ID: "1pJq2KT4WSfRIFdjGhZxm2m8aYfeq1g2KbUVLxPEL8TI",
    };
    logInfo("CONFIG", "✅ CONFIGをグローバルに定義しました");
  } else {
    logInfo("CONFIG", "⛔ CONFIGは既に定義されています（スキップ）");
  }
})();
/**
 * 🧱 CONFIGの安全な定義（他ファイルとの競合防止）
 */
(function defineGlobalConfig() {
  if (typeof globalThis.CONFIG === "undefined") {
    globalThis.CONFIG = {
      SPREADSHEET_ID: "1pJq2KT4WSfRIFdjGhZxm2m8aYfeq1g2KbUVLxPEL8TI",
    };
    logInfo("CONFIG", "✅ CONFIGをグローバルに定義しました");
  } else {
    logInfo("CONFIG", "⛔ CONFIGは既に定義されています（スキップ）");
  }
})();
