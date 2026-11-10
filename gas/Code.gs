/**
 * 名刺スキャン 受信用 Google Apps Script
 *
 * スプレッドシートに紐づいたスクリプトとして使う。ウェブアプリ
 * （実行ユーザー：自分、アクセス：全員）としてデプロイする。
 *
 * 既存のプロジェクトに別ファイル（例：meishi.gs）として追加しても動くよう、
 * 中身はすべて MeishiScan の中に閉じ込めている。グローバルに出る名前は
 * setup / rotateToken / doGet / doPost の4つだけ。
 * （同じプロジェクトの .gs ファイルは名前を共有するため、同名の定数はエラーに、
 *   同名の関数は黙って上書きになる）
 *
 * 書き込み先は既存の「名簿(AI)」シートの末尾。列は次の8列（見出しが違うと書き込まない）：
 *   A タイムスタンプ / B 連絡先受領者氏名 / C 会社名 / D 氏名 / E 氏名(フリガナ)
 *   F 電話番号 / G Email / H エビデンス（Drive の画像リンク）
 *
 * doPost の流れ：
 *   トークン照合 → カードIDで重複確認（Drive のファイル名で判定）→ 画像をDriveに保存
 *   → LLMで抽出 → シートの末尾に追加
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
 *   TARGET_SHEET         任意（既定：名簿(AI)）
 *   TOKEN / FOLDER_ID / SHEET_ID  setup() が作成する
 */

/** 初回に1回だけエディタから実行する（詳細は MeishiScan 内の setup_ を参照） */
function setup() { return MeishiScan.setup(); }

/** トークンを作り直す（漏れた疑いがあるとき） */
function rotateToken() { return MeishiScan.rotateToken(); }

function doGet(e) { return MeishiScan.doGet(e); }

function doPost(e) { return MeishiScan.doPost(e); }

const MeishiScan = (function () {
  const DEFAULT_TARGET_SHEET = '名簿(AI)';
  const DEFAULT_CLAUDE_MODEL = 'claude-sonnet-5-5';
  const DEFAULT_GEMINI_MODEL = 'gemini-3.6-flash';
  // 3MB 程度の画像まで受け付ける（base64 で約 4MB）
  const MAX_IMAGE_B64_LENGTH = 4 * 1024 * 1024;

  // 書き込み先シートの見出し（この順・この列数であることを確認してから書き込む）
  const HEADERS = ['タイムスタンプ', '連絡先受領者氏名', '会社名', '氏名', '氏名(フリガナ)', '電話番号', 'Email', 'エビデンス'];

  // LLM に抽出させる項目
  const FIELD_KEYS = ['company', 'name', 'name_kana', 'phone', 'mobile', 'email'];

  const PROMPT = [
    'この画像は1枚の名刺です。名刺に印刷されている情報を読み取り、指定のJSON形式で返してください。',
    '- company：会社名・団体名（部署名は含めない）。',
    '- name：氏名。姓と名の間は半角スペース1つ。',
    '- name_kana：氏名の読み。次の順で決める。',
    '  1. 名刺にふりがな・フリガナ・ローマ字の読みが印刷されていれば、その表記のまま書く。',
    '  2. 印刷されていなければ、メールアドレスの @ より前に氏名のローマ字（例：taro.yamada、t-yamada、yamada_t）が',
    '     含まれていて、名刺の漢字の氏名と対応すると判断できる場合だけ、その読みをカタカナで書く。',
    '     姓と名の間は半角スペース1つ。姓か名の片方しか読み取れなければ、読み取れた方だけ書く。',
    '  3. それ以外（info@ や部署名のアドレス、イニシャルだけ、氏名と対応しない文字列など）は空文字にする。',
    '     漢字の氏名だけから読みを推測しない。',
    '- phone：固定電話の番号。mobile：携帯電話の番号（070/080/090 や Mobile・携帯 の表記のもの）。FAX は含めない。',
    '- email：メールアドレス。',
    '- 電話番号は名刺の表記どおり（ハイフンを含む）に書く。',
    '- 同じ項目に複数の値がある場合は ", " でつなぐ。',
    '- 名刺に書かれていない項目は空文字にする。推測で補わない。',
    '- 名刺でない画像の場合は、すべて空文字にする。',
  ].join('\n');

  /* ---------------------------------------------------------------- */
  /* セットアップ                                                       */
  /* ---------------------------------------------------------------- */

  /**
   * 初回に1回だけエディタから実行する。
   * TOKEN / FOLDER_ID / SHEET_ID を作り、書き込み先シートの見出しを確認する。
   * 既に値がある項目は上書きしない（何度実行しても安全）。
   */
  function setup_() {
    const props = PropertiesService.getScriptProperties();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error('スプレッドシートの「拡張機能 → Apps Script」から作ったスクリプトで実行してください。');
    }

    const currentSheetId = props.getProperty('SHEET_ID');
    if (!currentSheetId) {
      props.setProperty('SHEET_ID', ss.getId());
    } else if (currentSheetId !== ss.getId()) {
      // 既存のプロジェクトに追加した場合、別用途の SHEET_ID が残っていることがある
      throw new Error('スクリプトプロパティ SHEET_ID が、このスプレッドシートとは別の ID になっています（' +
        currentSheetId + '）。意図したものでなければ SHEET_ID を削除してから setup() を再実行してください。');
    }

    // 書き込み先シートと見出しを確認する（問題があればここでエラーになる）
    getSheet_();

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

    Logger.log('セットアップ完了（書き込み先：' + targetSheetName_() + '）');
    Logger.log('TOKEN: ' + props.getProperty('TOKEN'));
    Logger.log('このトークンを撮影ページの「設定」に入力してください。');
  }

  /** トークンを作り直す（漏れた疑いがあるとき）。撮影ページの設定も更新すること。 */
  function rotateToken_() {
    const token = newToken_();
    PropertiesService.getScriptProperties().setProperty('TOKEN', token);
    Logger.log('新しい TOKEN: ' + token);
  }

  function newToken_() {
    return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  }

  /* ---------------------------------------------------------------- */
  /* ウェブアプリ                                                       */
  /* ---------------------------------------------------------------- */

  /** 動作確認用。データは返さない。 */
  function doGet_() {
    return json_({ ok: true, service: 'meishi-scan' });
  }

  function doPost_(e) {
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
      const receiver = clean_(String(req.receiver || '').slice(0, 50));
      if (!receiver) return json_({ ok: false, error: 'no_receiver' });
      const mimeType = req.mimeType === 'image/png' ? 'image/png' : 'image/jpeg';
      const imageB64 = String(req.image || '');
      if (!imageB64 || imageB64.length > MAX_IMAGE_B64_LENGTH) {
        return json_({ ok: false, error: 'bad_image' });
      }
      const folderId = props.getProperty('FOLDER_ID');
      if (!folderId) throw new Error('FOLDER_ID が未設定です。setup() を実行してください。');
      const folder = DriveApp.getFolderById(folderId);

      // 重複確認（登録済み、または別の実行が処理中か）
      const dup = reserve_(cardId, folder);
      if (dup === 'exists') return json_({ ok: true, duplicate: true, cardId: cardId });
      if (dup === 'processing') return json_({ ok: false, error: 'processing', cardId: cardId });
      reservedId = cardId;

      // 画像をDriveに保存（フォルダは非公開のまま）。ファイル名のカードIDで重複を判定する
      const ext = mimeType === 'image/png' ? 'png' : 'jpg';
      const stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
      const blob = Utilities.newBlob(Utilities.base64Decode(imageB64), mimeType, stamp + '_' + cardId + '.' + ext);
      savedFile = folder.createFile(blob);

      // LLMで抽出
      const f = extract_(imageB64, mimeType, props).fields;
      if (!f.company && !f.name && !f.email && !f.phone && !f.mobile) {
        throw new Error('名刺の情報を読み取れませんでした（名刺が写っているか確認してください）');
      }

      // シートの末尾に追加
      const capturedAt = parseDate_(req.capturedAt) || new Date();
      const phones = [f.phone, f.mobile].filter(function (s) { return s; }).join(', ');
      const row = [
        capturedAt,
        receiver,
        clean_(f.company),
        clean_(f.name),
        clean_(f.name_kana),
        clean_(phones),
        clean_(f.email),
        savedFile.getUrl(),
      ];
      appendRow_(row);
      savedFile = null; // 成功したのでファイルは残す（重複判定にも使う）

      return json_({ ok: true, cardId: cardId, name: clean_(f.name), company: clean_(f.company) });
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
    if (!s) return null;
    const d = new Date(String(s));
    return isNaN(d.getTime()) ? null : d;
  }

  /* ---------------------------------------------------------------- */
  /* スプレッドシート                                                   */
  /* ---------------------------------------------------------------- */

  function targetSheetName_() {
    return PropertiesService.getScriptProperties().getProperty('TARGET_SHEET') || DEFAULT_TARGET_SHEET;
  }

  /** 見出しの比較用（全角・半角の括弧や空白の違いは無視する） */
  function normalize_(s) {
    return String(s).normalize('NFKC').replace(/\s+/g, '');
  }

  /** 書き込み先シートを返す。シートが無い、または見出しが想定と違う場合はエラー（シートは作らない） */
  function getSheet_() {
    const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    if (!id) throw new Error('SHEET_ID が未設定です。setup() を実行してください。');
    const name = targetSheetName_();
    const sheet = SpreadsheetApp.openById(id).getSheetByName(name);
    if (!sheet) throw new Error('シート「' + name + '」が見つかりません。');
    const actual = sheet.getRange(1, 1, 1, HEADERS.length).getDisplayValues()[0];
    const ok = HEADERS.every(function (h, i) { return normalize_(h) === normalize_(actual[i]); });
    if (!ok) {
      throw new Error('シート「' + name + '」の1行目の見出しが想定と違うため書き込みません。' +
        '想定：' + HEADERS.join(' / ') + '　実際：' + actual.join(' / '));
    }
    return sheet;
  }

  /** Drive のフォルダに、このカードIDの画像（ゴミ箱以外）があれば登録済み */
  function isRegistered_(folder, cardId) {
    // cardId は英数字とハイフンだけ（doPost で確認済み）なので検索式に入れても安全
    return folder.searchFiles("title contains '" + cardId + "' and trashed = false").hasNext();
  }

  /**
   * 処理の予約。登録済みなら 'exists'、他で処理中なら 'processing'、
   * 予約できたら null を返す。並行送信での二重登録を防ぐ。
   */
  function reserve_(cardId, folder) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const cache = CacheService.getScriptCache();
      if (cache.get('proc:' + cardId)) return 'processing';
      if (isRegistered_(folder, cardId)) return 'exists';
      cache.put('proc:' + cardId, '1', 600);
      return null;
    } finally {
      lock.releaseLock();
    }
  }

  /**
   * シートの末尾に1行追加する。
   * A列（タイムスタンプ）が最後に入っている行の次から、A〜H列がすべて空の行を探して書き込む。
   * （シートの下の方に入っている無関係な値で、行が飛んだり上書きしたりしないようにする）
   */
  function appendRow_(row) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const sheet = getSheet_();
      const width = HEADERS.length;
      const lastRow = Math.max(sheet.getLastRow(), 1);
      const colA = sheet.getRange(1, 1, lastRow, 1).getDisplayValues();
      let r = 1;
      for (let i = colA.length - 1; i >= 0; i--) {
        if (colA[i][0] !== '') { r = i + 1; break; }
      }
      r += 1;
      // A〜H がすべて空の行まで進める
      if (r <= lastRow) {
        const block = sheet.getRange(r, 1, lastRow - r + 1, width).getDisplayValues();
        let i = 0;
        while (i < block.length && block[i].some(function (v) { return v !== ''; })) i++;
        r += i;
      }
      const formats = row.map(function (v) {
        return v instanceof Date ? 'yyyy/mm/dd hh:mm:ss' : '@'; // 文字列は書式なしテキスト（先頭の0を保つ）
      });
      sheet.getRange(r, 1, 1, width).setNumberFormats([formats]).setValues([row]);
      return r;
    } finally {
      lock.releaseLock();
    }
  }

  /** 文字列を整え、数式として解釈されないようにする */
  function clean_(v) {
    if (v === null || v === undefined) return '';
    let s = String(v).replace(/^\s+|\s+$/g, '').slice(0, 1000);
    if (/^[=+\-@]/.test(s)) s = "'" + s;
    return s;
  }

  /* ---------------------------------------------------------------- */
  /* LLM                                                                */
  /* ---------------------------------------------------------------- */

  function extract_(imageB64, mimeType, props) {
    const provider = String(props.getProperty('PROVIDER') || 'claude').trim().toLowerCase();
    if (provider === 'gemini') return callGemini_(imageB64, mimeType, props);
    if (provider === 'claude') return callClaude_(imageB64, mimeType, props);
    throw new Error('PROVIDER は claude か gemini を指定してください: ' + provider);
  }

  function jsonSchema_() {
    const properties = {};
    FIELD_KEYS.forEach(function (k) { properties[k] = { type: 'string' }; });
    return {
      type: 'object',
      properties: properties,
      required: FIELD_KEYS.slice(),
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
    FIELD_KEYS.forEach(function (k) { properties[k] = { type: 'STRING' }; });

    // Gemini 3 系は temperature 等を既定値のまま使うことが推奨されているため送らない
    const generationConfig = {
      responseMimeType: 'application/json',
      responseSchema: {
        type: 'OBJECT',
        properties: properties,
        required: FIELD_KEYS.slice(),
        propertyOrdering: FIELD_KEYS.slice(),
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
    FIELD_KEYS.forEach(function (k) { out[k] = obj && obj[k] != null ? String(obj[k]).trim() : ''; });
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

  return { setup: setup_, rotateToken: rotateToken_, doGet: doGet_, doPost: doPost_ };
})();
