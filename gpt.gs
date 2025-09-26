const OPENAI_API_KEY = 'ここにあなたのOpenAI APIキーを入力'; // 必須！

function askGPT(userQuestion) {
  const url = "https://api.openai.com/v1/chat/completions";

  const payload = {
    model: "gpt-3.5-turbo",
    messages: [
      { role: "system", content: "あなたは親切な業務アシスタントです。フォーム入力で困っている人に簡潔かつわかりやすく答えてください。" },
      { role: "user", content: userQuestion }
    ],
    temperature: 0.7,
    max_tokens: 500
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${OPENAI_API_KEY}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();

  if (responseCode !== 200) {
    throw new Error(`APIエラー (${responseCode})`);
  }

  const json = JSON.parse(response.getContentText());
  const reply = json.choices?.[0]?.message?.content;
  return reply || "回答が取得できませんでした。";
}
/**
 * OpenAI GPT連携モジュール（チャット形式）
 */

function callGPT(promptText) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error("OpenAI APIキーが設定されていません");

  const url = "https://api.openai.com/v1/chat/completions";

  const payload = {
    model: "gpt-4",
    messages: [
      { role: "system", content: "あなたは親切で正確な業務支援アシスタントです。" },
      { role: "user", content: promptText }
    ],
    temperature: 0.7
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${apiKey}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const result = JSON.parse(response.getContentText());

  if (result.error) {
    throw new Error(result.error.message);
  }

  return result.choices[0].message.content.trim();
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
  let sheet = ss.getSheetByName(sheetName); // ✅ 最初に sheet を宣言

  if (!sheet) {
    sheet = ss.insertSheet(sheetName); // ✅ 再代入OKなので let で問題なし
    logInfo("getSheet", `🛠️ シート "${sheetName}" を新規作成しました`);
  }

  return sheet; // ✅ 必ず返す
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

