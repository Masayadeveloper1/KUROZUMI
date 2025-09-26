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
