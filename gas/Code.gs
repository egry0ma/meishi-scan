/**
 * 名刺スキャン 受信用 Google Apps Script
 *
 * スプレッドシートに紐づいたスクリプトとして使う。ウェブアプリ
 * （実行ユーザー：自分、アクセス：全員）としてデプロイする。
 *
 * doPost の流れ：
 *   トークン照合 → カードIDで重複確認 → 画像をDriveに保存 → LLMで抽出 → スプシに追加
 *
 * 書き込み専用。スプシやDriveの既存データを返す処理は持たない。
 *
 * スクリプトプロパティ：
 *   PROVIDER             claude | gemini（既定：claude）
 *   ANTHROPIC_API_KEY    PROVIDER=claude のとき必須
 *   GEMINI_API_KEY       PROVIDER=gemini のとき必須
 *   CLAUDE_MODEL         任意（既定：claude-sonnet-5-5）
 *   GEMINI_MODEL         任意（既定：gemini-3.6-flash）
 *   GEMINI_THINKING_LEVEL 任意（既定：low。"none" で送らない）
 *   TOKEN / FOLDER_ID / SHEET_ID  setup() が作成する
 */

const SHEET_NAME = '名刺';
const DEFAULT_CLAUDE_MODEL = 'claude-sonnet-5-5';
const DEFAULT_GEMINI_MODEL = 'gemini-3.6-flash';
// 3MB 程度の画像まで受け付ける（base64 で約 4MB）
const MAX_IMAGE_B64_LENGTH = 4 * 1024 * 1024;

// LLM に抽出させる項目（キー, 見出し）
const FIELDS = [
  ['name', '氏名'],
  ['name_kana', '氏名（読み）'],
  ['company', '会社名'],
  ['department', '部署'],
  ['title', '役職'],
  ['postal_code', '郵便番号'],
  ['address', '住所'],
  ['phone', '電話'],
  ['mobile', '携帯'],
  ['fax', 'FAX'],
  ['email', 'メール'],
  ['url', 'URL'],
  ['notes', 'その他'],
];

const HEADERS = ['登録日時', 'カードID', '撮影日時']
  .concat(FIELDS.map(function (f) { return f[1]; }))
  .concat(['画像', 'LLM']);
const CARD_ID_COL = 2; // B列

const PROMPT = [
  'この画像は1枚の名刺です。名刺に印刷されている情報を読み取り、指定のJSON形式で返してください。',
  '- 名刺に書かれていない項目は空文字にする。推測で補わない。',
  '- 同じ項目に複数の値がある場合（電話番号が2つなど）は " / " でつなぐ。',
  '- name_kana は名刺にふりがな・ローマ字表記があればそれを入れる。無ければ空文字。',
  '- phone は固定電話、mobile は携帯電話（070/080/090 など、または Mobile/携帯 の表記）。',
  '- 電話番号・郵便番号は名刺の表記どおり（ハイフンを含む）に書く。',
  '- notes には上の項目に入らない情報（資格、SNS、営業時間など）を簡潔に入れる。',
  '- 名刺でない画像の場合は、すべて空文字にし、notes に「名刺ではない」と書く。',
].join('\n');

/* ------------------------------------------------------------------ */
/* セットアップ                                                         */
/* ------------------------------------------------------------------ */

/**
 * 初回に1回だけエディタから実行する。
 * TOKEN / FOLDER_ID / SHEET_ID を作り、シートと見出し行を用意する。
 * 既に値がある項目は上書きしない（何度実行しても安全）。
 */
function setup() {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('スプレッドシートの「拡張機能 → Apps Script」から作ったスクリプトで実行してください。');
  }

  if (!props.getProperty('SHEET_ID')) {
    props.setProperty('SHEET_ID', ss.getId());
  }
  if (!props.getProperty('FOLDER_ID')) {
    const folder = DriveApp.createFolder('名刺スキャン画像');
    props.setProperty('FOLDER_ID', folder.getId());
  }
  if (!props.getProperty('TOKEN')) {
    props.setProperty('TOKEN', newToken_());
  }
  if (!props.getProperty('PROVIDER')) {
    props.setProperty('PROVIDER', 'claude');
  }

  getSheet_();

  Logger.log('セットアップ完了');
  Logger.log('TOKEN: ' + props.getProperty('TOKEN'));
  Logger.log('このトークンを撮影ページの「設定」に入力してください。');
}

/** トークンを作り直す（漏れた疑いがあるとき）。撮影ページの設定も更新すること。 */
function rotateToken() {
  const token = newToken_();
  PropertiesService.getScriptProperties().setProperty('TOKEN', token);
  Logger.log('新しい TOKEN: ' + token);
}

function newToken_() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
}

/* ------------------------------------------------------------------ */
/* ウェブアプリ                                                         */
/* ------------------------------------------------------------------ */

/** 動作確認用。データは返さない。 */
function doGet() {
  return json_({ ok: true, service: 'meishi-scan' });
}

function doPost(e) {
  let reservedId = null;
  let savedFile = null;
  try {
    const req = parseRequest_(e);
    const props = PropertiesService.getScriptProperties();

    const token = props.getProperty('TOKEN');
    if (!token) return json_({ ok: false, error: 'not_configured' });
    if (!safeEqual_(String(req.token || ''), token)) {
      return json_({ ok: false, error: 'unauthorized' });
    }

    // 接続テスト（トークンの確認だけ。LLMは呼ばない）
    if (req.action === 'ping') return json_({ ok: true, pong: true });

    const cardId = String(req.cardId || '');
    if (!/^[A-Za-z0-9-]{8,64}$/.test(cardId)) {
      return json_({ ok: false, error: 'bad_card_id' });
    }
    const mimeType = req.mimeType === 'image/png' ? 'image/png' : 'image/jpeg';
    const imageB64 = String(req.image || '');
    if (!imageB64 || imageB64.length > MAX_IMAGE_B64_LENGTH) {
      return json_({ ok: false, error: 'bad_image' });
    }

    // 重複確認（登録済み、または別の実行が処理中か）
    const dup = reserve_(cardId);
    if (dup === 'exists') return json_({ ok: true, duplicate: true, cardId: cardId });
    if (dup === 'processing') return json_({ ok: false, error: 'processing', cardId: cardId });
    reservedId = cardId;

    // 画像をDriveに保存（フォルダは非公開のまま）
    const folderId = props.getProperty('FOLDER_ID');
    if (!folderId) throw new Error('FOLDER_ID が未設定です。setup() を実行してください。');
    const ext = mimeType === 'image/png' ? 'png' : 'jpg';
    const stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
    const blob = Utilities.newBlob(Utilities.base64Decode(imageB64), mimeType, stamp + '_' + cardId + '.' + ext);
    savedFile = DriveApp.getFolderById(folderId).createFile(blob);

    // LLMで抽出
    const extracted = extract_(imageB64, mimeType, props);

    // スプシに追加
    const capturedAt = parseDate_(req.capturedAt);
    const row = [new Date(), cardId, capturedAt]
      .concat(FIELDS.map(function (f) { return clean_(extracted.fields[f[0]]); }))
      .concat([savedFile.getUrl(), extracted.model]);
    const added = appendRow_(cardId, row);
    if (!added) {
      // 他の実行が先に書き込んでいた
      savedFile.setTrashed(true);
      savedFile = null;
      return json_({ ok: true, duplicate: true, cardId: cardId });
    }
    savedFile = null; // 成功したのでファイルは残す

    return json_({
      ok: true,
      cardId: cardId,
      name: clean_(extracted.fields.name),
      company: clean_(extracted.fields.company),
    });
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    if (savedFile) {
      try { savedFile.setTrashed(true); } catch (ignore) { /* noop */ }
    }
    return json_({ ok: false, error: String((err && err.message) || err) });
  } finally {
    if (reservedId) CacheService.getScriptCache().remove('proc:' + reservedId);
  }
}

function parseRequest_(e) {
  if (!e || !e.postData || !e.postData.contents) throw new Error('empty_body');
  try {
    return JSON.parse(e.postData.contents);
  } catch (err) {
    throw new Error('bad_json');
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 長さと内容を一定時間で比較する */
function safeEqual_(a, b) {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function parseDate_(s) {
  if (!s) return '';
  const d = new Date(String(s));
  return isNaN(d.getTime()) ? '' : d;
}

/* ------------------------------------------------------------------ */
/* スプレッドシート                                                     */
/* ------------------------------------------------------------------ */

function getSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('SHEET_ID が未設定です。setup() を実行してください。');
  const ss = SpreadsheetApp.openById(id);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findCard_(sheet, cardId) {
  const last = sheet.getLastRow();
  if (last < 2) return false;
  return !!sheet.getRange(2, CARD_ID_COL, last - 1, 1)
    .createTextFinder(cardId).matchEntireCell(true).findNext();
}

/**
 * 処理の予約。登録済みなら 'exists'、他で処理中なら 'processing'、
 * 予約できたら null を返す。並行送信での二重登録を防ぐ。
 */
function reserve_(cardId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const cache = CacheService.getScriptCache();
    if (cache.get('proc:' + cardId)) return 'processing';
    if (findCard_(getSheet_(), cardId)) return 'exists';
    cache.put('proc:' + cardId, '1', 600);
    return null;
  } finally {
    lock.releaseLock();
  }
}

/** 1行追加する。既に同じカードIDがあれば追加せず false を返す。 */
function appendRow_(cardId, row) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getSheet_();
    if (findCard_(sheet, cardId)) return false;
    const r = sheet.getLastRow() + 1;
    const formats = row.map(function (v) {
      return v instanceof Date ? 'yyyy/mm/dd hh:mm:ss' : '@'; // 文字列は書式なしテキスト（先頭の0を保つ）
    });
    sheet.getRange(r, 1, 1, row.length).setNumberFormats([formats]).setValues([row]);
    return true;
  } finally {
    lock.releaseLock();
  }
}

/** 文字列を整え、数式として解釈されないようにする */
function clean_(v) {
  if (v === null || v === undefined) return '';
  let s = String(v).replace(/\s+$/g, '').replace(/^\s+/g, '').slice(0, 1000);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

/* ------------------------------------------------------------------ */
/* LLM                                                                  */
/* ------------------------------------------------------------------ */

function extract_(imageB64, mimeType, props) {
  const provider = String(props.getProperty('PROVIDER') || 'claude').trim().toLowerCase();
  if (provider === 'gemini') return callGemini_(imageB64, mimeType, props);
  if (provider === 'claude') return callClaude_(imageB64, mimeType, props);
  throw new Error('PROVIDER は claude か gemini を指定してください: ' + provider);
}

function jsonSchema_() {
  const properties = {};
  FIELDS.forEach(function (f) { properties[f[0]] = { type: 'string' }; });
  return {
    type: 'object',
    properties: properties,
    required: FIELDS.map(function (f) { return f[0]; }),
    additionalProperties: false,
  };
}

// サーバー側フォールバック（拒否時に別モデルで再実行）に対応するモデル
const CLAUDE_FALLBACK_MODELS = ['claude-sonnet-5-5', 'claude-opus-5-5', 'claude-opus-5', 'claude-fable-5-1'];

function callClaude_(imageB64, mimeType, props) {
  const key = props.getProperty('ANTHROPIC_API_KEY');
  if (!key) throw new Error('ANTHROPIC_API_KEY が未設定です。');
  const model = props.getProperty('CLAUDE_MODEL') || DEFAULT_CLAUDE_MODEL;

  // temperature などのサンプリングパラメータは送らない（新しいモデル・APIバージョンでは400になる）
  const body = {
    model: model,
    max_tokens: 8000,
    output_config: {
      effort: 'low', // 単純な抽出なので思考は少なめ
      format: { type: 'json_schema', schema: jsonSchema_() },
    },
    messages: [{
      role: 'user',
      content: [
        { type: 'image', source: { type: 'base64', media_type: mimeType, data: imageB64 } },
        { type: 'text', text: PROMPT },
      ],
    }],
  };
  const headers = {
    'x-api-key': key,
    'anthropic-version': '2026-10-21',
  };
  if (CLAUDE_FALLBACK_MODELS.indexOf(model) >= 0) {
    body.fallbacks = 'default';
    headers['anthropic-beta'] = 'server-side-fallback-2026-07-01';
  }

  const res = fetchWithRetry_('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: headers,
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });
  const data = JSON.parse(res.getContentText());

  if (data.stop_reason === 'refusal') throw new Error('Claude が処理を拒否しました');
  if (data.stop_reason === 'max_tokens') throw new Error('Claude の出力が途中で切れました');
  const text = (data.content || [])
    .filter(function (b) { return b.type === 'text'; })
    .map(function (b) { return b.text; })
    .join('');
  return { fields: parseFields_(text), model: data.model || model };
}

function callGemini_(imageB64, mimeType, props) {
  const key = props.getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('GEMINI_API_KEY が未設定です。');
  const model = props.getProperty('GEMINI_MODEL') || DEFAULT_GEMINI_MODEL;
  const thinkingLevel = props.getProperty('GEMINI_THINKING_LEVEL') || 'low';

  const properties = {};
  FIELDS.forEach(function (f) { properties[f[0]] = { type: 'STRING' }; });

  // Gemini 3 系は temperature 等を既定値のまま使うことが推奨されているため送らない
  const generationConfig = {
    responseMimeType: 'application/json',
    responseSchema: {
      type: 'OBJECT',
      properties: properties,
      required: FIELDS.map(function (f) { return f[0]; }),
      propertyOrdering: FIELDS.map(function (f) { return f[0]; }),
    },
  };
  if (thinkingLevel !== 'none') {
    generationConfig.thinkingConfig = { thinkingLevel: thinkingLevel };
  }

  const body = {
    contents: [{
      role: 'user',
      parts: [
        { inlineData: { mimeType: mimeType, data: imageB64 } },
        { text: PROMPT },
      ],
    }],
    generationConfig: generationConfig,
  };

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
    encodeURIComponent(model) + ':generateContent';
  const res = fetchWithRetry_(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': key }, // キーはURLに入れない（ログに残さないため）
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });
  const data = JSON.parse(res.getContentText());

  if (data.promptFeedback && data.promptFeedback.blockReason) {
    throw new Error('Gemini がブロックしました: ' + data.promptFeedback.blockReason);
  }
  const cand = (data.candidates || [])[0];
  if (!cand || !cand.content) {
    throw new Error('Gemini の応答が空です' + (cand && cand.finishReason ? ': ' + cand.finishReason : ''));
  }
  const text = (cand.content.parts || [])
    .filter(function (p) { return typeof p.text === 'string' && !p.thought; })
    .map(function (p) { return p.text; })
    .join('');
  return { fields: parseFields_(text), model: data.modelVersion || model };
}

function parseFields_(text) {
  let obj;
  try {
    obj = JSON.parse(text);
  } catch (err) {
    // 念のため、前後に余計な文字があっても {...} を取り出す
    const m = String(text).match(/\{[\s\S]*\}/);
    if (!m) throw new Error('LLMの応答をJSONとして読めません');
    obj = JSON.parse(m[0]);
  }
  const out = {};
  FIELDS.forEach(function (f) { out[f[0]] = obj && obj[f[0]] != null ? String(obj[f[0]]) : ''; });
  return out;
}

/** 429 / 5xx / 529 は少し待って再試行する */
function fetchWithRetry_(url, options) {
  const maxAttempts = 3;
  for (let i = 1; ; i++) {
    const res = UrlFetchApp.fetch(url, options);
    const code = res.getResponseCode();
    if (code >= 200 && code < 300) return res;
    const retryable = code === 429 || code >= 500;
    if (!retryable || i >= maxAttempts) {
      throw new Error('LLM API エラー ' + code + ': ' + res.getContentText().slice(0, 500));
    }
    Utilities.sleep(1000 * Math.pow(2, i));
  }
}
