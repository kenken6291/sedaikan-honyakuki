/**
 * 昭和・平成・令和・Z翻訳機 - GASバックエンド
 * ------------------------------------------------------
 * シート構成:
 *   Members  : email, nickname, passwordHash, sns, role, status, agreedAt, createdAt
 *   Sessions : token, email, createdAt, expiresAt
 *   Events   : eventId, title, description, datetime, place, capacity, deadline,
 *              ownerEmail, ownerNickname, status, createdAt
 *   Participants : eventId, email, nickname, sns, count, status(joined/declined), updatedAt
 *
 * PropertiesService に設定する値:
 *   SPREADSHEET_ID   : このアプリ専用スプレッドシートのID
 *   GEMINI_API_KEY   : 翻訳機能で使うGemini APIキー（Google AI Studioで取得）
 *   GEMINI_MODEL     : 既定のGeminiモデル名（任意。例: gemini-4-hydrogen）
 *                      未設定なら汎用の最新モデル（FALLBACK_GEMINI_MODELS）を使用
 *
 * Geminiモデルの決定順:
 *   1. 画面の「AIモデル」で選ばれたモデル
 *   2. スクリプトプロパティ GEMINI_MODEL
 *   3. 汎用の最新モデル FALLBACK_GEMINI_MODELS を先頭から順に試す
 *   指定モデルが存在しない（廃止・名前違い）場合は自動で汎用の最新モデルへ切り替える
 *   ADMIN_PASSWORD   : 管理者ログイン用パスワード(平文ではなくハッシュを推奨。ここではハッシュ保存)
 */

// ===== 共通設定 =====
const SHEET_MEMBERS = 'Members';
const SHEET_SESSIONS = 'Sessions';
const SHEET_EVENTS = 'Events';
const SHEET_PARTICIPANTS = 'Participants';
const SESSION_TTL_MS = 1000 * 60 * 60 * 24 * 14; // 14日間

// ===== Geminiモデル設定 =====
// 指定なし・該当なしの時に使う汎用最新モデル（先頭から順に試す）
//   gemini-flash-latest : Googleが常に最新の汎用モデルへ自動で差し替えるエイリアス
//   gemini-4-hydrogen   : 2026年11月時点の汎用最新モデル（「日々の暮らしをサポート」）
//   gemini-3.8-flash    : 最終的な保険
const FALLBACK_GEMINI_MODELS = ['gemini-flash-latest', 'gemini-4-hydrogen', 'gemini-3.8-flash'];
const GEMINI_MODEL_PATTERN = /^gemini-[a-z0-9][a-z0-9.\-]{0,60}$/;
const GEMINI_RETRY_CODES = [429, 500, 503];
const GEMINI_MAX_ATTEMPTS = 3;

function getSS_() {
  const props = PropertiesService.getScriptProperties();
  const ssId = props.getProperty('SPREADSHEET_ID');
  return SpreadsheetApp.openById(ssId);
}

function getSheet_(name) {
  const ss = getSS_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

// ===== エントリポイント =====
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, message: 'sedaikan-honyakuki API' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result;
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    switch (action) {
      case 'register':
        result = handleRegister_(body);
        break;
      case 'login':
        result = handleLogin_(body);
        break;
      case 'requestPasswordReset':
        result = handleRequestPasswordReset_(body);
        break;
      case 'changePassword':
        result = handleChangePassword_(body);
        break;
      case 'updateProfile':
        result = handleUpdateProfile_(body);
        break;
      case 'getProfile':
        result = handleGetProfile_(body);
        break;
      case 'translate':
        result = handleTranslate_(body);
        break;
      case 'listEvents':
        result = handleListEvents_(body);
        break;
      case 'createEvent':
        result = handleCreateEvent_(body);
        break;
      case 'updateEvent':
        result = handleUpdateEvent_(body);
        break;
      case 'deleteEvent':
        result = handleDeleteEvent_(body);
        break;
      case 'joinEvent':
        result = handleJoinEvent_(body);
        break;
      case 'leaveEvent':
        result = handleLeaveEvent_(body);
        break;
      case 'getParticipants':
        result = handleGetParticipants_(body);
        break;
      case 'adminLogin':
        result = handleAdminLogin_(body);
        break;
      case 'adminListMembers':
        result = handleAdminListMembers_(body);
        break;
      case 'adminUpdateMemberStatus':
        result = handleAdminUpdateMemberStatus_(body);
        break;
      case 'adminListEvents':
        result = handleAdminListEvents_(body);
        break;
      case 'adminDeleteEvent':
        result = handleAdminDeleteEvent_(body);
        break;
      default:
        result = { ok: false, error: '不明なアクションです: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: 'サーバーエラー: ' + err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== ユーティリティ =====
function sha256_(text) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return rawHash.map(function (b) {
    const v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function generateTempPassword_() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789';
  let pw = '';
  for (let i = 0; i < 10; i++) {
    pw += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return pw;
}

function generateToken_() {
  return Utilities.getUuid();
}

function findRowByValue_(sheet, colIndex, value) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colIndex]) === String(value)) {
      return { rowIndex: i + 1, row: data[i] };
    }
  }
  return null;
}

function getAllRows_(sheet) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    headers.forEach(function (h, idx) {
      obj[h] = data[i][idx];
    });
    obj._rowIndex = i + 1;
    rows.push(obj);
  }
  return rows;
}

function ensureHeaders_() {
  const membersSheet = getSheet_(SHEET_MEMBERS);
  if (membersSheet.getLastRow() === 0) {
    membersSheet.appendRow(['email', 'nickname', 'passwordHash', 'sns', 'role', 'status', 'agreedAt', 'createdAt']);
  }
  const sessionsSheet = getSheet_(SHEET_SESSIONS);
  if (sessionsSheet.getLastRow() === 0) {
    sessionsSheet.appendRow(['token', 'email', 'createdAt', 'expiresAt']);
  }
  const eventsSheet = getSheet_(SHEET_EVENTS);
  if (eventsSheet.getLastRow() === 0) {
    eventsSheet.appendRow(['eventId', 'title', 'description', 'datetime', 'place', 'capacity', 'deadline', 'ownerEmail', 'ownerNickname', 'status', 'createdAt']);
  }
  const participantsSheet = getSheet_(SHEET_PARTICIPANTS);
  if (participantsSheet.getLastRow() === 0) {
    participantsSheet.appendRow(['eventId', 'email', 'nickname', 'sns', 'count', 'status', 'updatedAt']);
  }
}

// ===== セッション検証 =====
function validateSession_(token) {
  if (!token) return null;
  const sheet = getSheet_(SHEET_SESSIONS);
  const found = findRowByValue_(sheet, 0, token);
  if (!found) return null;
  const expiresAt = new Date(found.row[3]).getTime();
  if (Date.now() > expiresAt) return null;
  return { email: found.row[1] };
}

function getMemberByEmail_(email) {
  const sheet = getSheet_(SHEET_MEMBERS);
  const found = findRowByValue_(sheet, 0, email);
  if (!found) return null;
  return {
    email: found.row[0],
    nickname: found.row[1],
    passwordHash: found.row[2],
    sns: found.row[3],
    role: found.row[4],
    status: found.row[5],
    agreedAt: found.row[6],
    createdAt: found.row[7],
    _rowIndex: found.rowIndex
  };
}

// ===== 会員登録 =====
function handleRegister_(body) {
  ensureHeaders_();
  const email = String(body.email || '').trim().toLowerCase();
  const nickname = String(body.nickname || '').trim();
  const sns = String(body.sns || '').trim();
  const agreed = body.agreed === true;

  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return { ok: false, error: 'メールアドレスの形式が正しくありません。' };
  }
  if (!nickname) {
    return { ok: false, error: 'ニックネームを入力してください。' };
  }
  if (!agreed) {
    return { ok: false, error: '免責事項および注意事項への同意が必要です。' };
  }
  if (getMemberByEmail_(email)) {
    return { ok: false, error: 'このメールアドレスは既に登録されています。' };
  }

  const tempPassword = generateTempPassword_();
  const passwordHash = sha256_(tempPassword);
  const now = new Date();

  const sheet = getSheet_(SHEET_MEMBERS);
  sheet.appendRow([email, nickname, passwordHash, sns, 'member', 'active', now.toISOString(), now.toISOString()]);

  try {
    MailApp.sendEmail({
      to: email,
      subject: '【昭和・平成・令和・Z翻訳機】仮パスワードのお知らせ',
      body:
        nickname + ' 様\n\n' +
        '会員登録ありがとうございます。\n' +
        '以下の仮パスワードでログインし、必ずパスワードを変更してください。\n\n' +
        '会員番号（メールアドレス）: ' + email + '\n' +
        '仮パスワード: ' + tempPassword + '\n\n' +
        '※本メールは送信専用です。心当たりのない場合は破棄してください。'
    });
  } catch (err) {
    return { ok: false, error: 'メール送信に失敗しました: ' + err.message };
  }

  return { ok: true, message: '登録が完了しました。仮パスワードをメールに送信しました。' };
}

// ===== ログイン =====
function handleLogin_(body) {
  ensureHeaders_();
  const email = String(body.email || '').trim().toLowerCase();
  const password = String(body.password || '');

  const member = getMemberByEmail_(email);
  if (!member) {
    return { ok: false, error: '会員番号またはパスワードが正しくありません。' };
  }
  if (member.status !== 'active') {
    return { ok: false, error: 'このアカウントは利用停止中です。' };
  }
  if (sha256_(password) !== member.passwordHash) {
    return { ok: false, error: '会員番号またはパスワードが正しくありません。' };
  }

  const token = generateToken_();
  const now = new Date();
  const expiresAt = new Date(now.getTime() + SESSION_TTL_MS);
  const sessionsSheet = getSheet_(SHEET_SESSIONS);
  sessionsSheet.appendRow([token, email, now.toISOString(), expiresAt.toISOString()]);

  return {
    ok: true,
    token: token,
    profile: { email: member.email, nickname: member.nickname, sns: member.sns, role: member.role }
  };
}

// ===== パスワード再発行 =====
function handleRequestPasswordReset_(body) {
  ensureHeaders_();
  const email = String(body.email || '').trim().toLowerCase();
  const member = getMemberByEmail_(email);
  if (!member) {
    // メール存在の有無を外部に漏らさないため、成功メッセージのみ返す
    return { ok: true, message: '登録されている場合、仮パスワードを送信しました。' };
  }

  const tempPassword = generateTempPassword_();
  const passwordHash = sha256_(tempPassword);

  const sheet = getSheet_(SHEET_MEMBERS);
  sheet.getRange(member._rowIndex, 3).setValue(passwordHash);

  try {
    MailApp.sendEmail({
      to: email,
      subject: '【昭和・平成・令和・Z翻訳機】仮パスワード再発行のお知らせ',
      body:
        member.nickname + ' 様\n\n' +
        'パスワード再発行のリクエストを受け付けました。\n' +
        '以下の仮パスワードでログインし、必ずパスワードを変更してください。\n\n' +
        '会員番号（メールアドレス）: ' + email + '\n' +
        '仮パスワード: ' + tempPassword + '\n\n' +
        '※心当たりのない場合は本メールを破棄し、念のためサポートまでご連絡ください。'
    });
  } catch (err) {
    return { ok: false, error: 'メール送信に失敗しました: ' + err.message };
  }

  return { ok: true, message: '登録されている場合、仮パスワードを送信しました。' };
}

// ===== パスワード変更 =====
function handleChangePassword_(body) {
  const session = validateSession_(body.token);
  if (!session) return { ok: false, error: 'セッションが無効です。再度ログインしてください。' };

  const newPassword = String(body.newPassword || '');
  if (newPassword.length < 8) {
    return { ok: false, error: 'パスワードは8文字以上で設定してください。' };
  }

  const member = getMemberByEmail_(session.email);
  if (!member) return { ok: false, error: '会員情報が見つかりません。' };

  const sheet = getSheet_(SHEET_MEMBERS);
  sheet.getRange(member._rowIndex, 3).setValue(sha256_(newPassword));

  return { ok: true, message: 'パスワードを変更しました。' };
}

// ===== プロフィール更新・取得 =====
function handleUpdateProfile_(body) {
  const session = validateSession_(body.token);
  if (!session) return { ok: false, error: 'セッションが無効です。再度ログインしてください。' };

  const nickname = String(body.nickname || '').trim();
  const sns = String(body.sns || '').trim();
  if (!nickname) return { ok: false, error: 'ニックネームを入力してください。' };

  const member = getMemberByEmail_(session.email);
  if (!member) return { ok: false, error: '会員情報が見つかりません。' };

  const sheet = getSheet_(SHEET_MEMBERS);
  sheet.getRange(member._rowIndex, 2).setValue(nickname);
  sheet.getRange(member._rowIndex, 4).setValue(sns);

  return { ok: true, message: 'プロフィールを更新しました。' };
}

function handleGetProfile_(body) {
  const session = validateSession_(body.token);
  if (!session) return { ok: false, error: 'セッションが無効です。再度ログインしてください。' };

  const member = getMemberByEmail_(session.email);
  if (!member) return { ok: false, error: '会員情報が見つかりません。' };

  return {
    ok: true,
    profile: { email: member.email, nickname: member.nickname, sns: member.sns, role: member.role }
  };
}

// ===== 翻訳機能 =====
function handleTranslate_(body) {
  const text = String(body.text || '').trim();
  const targetGen = String(body.targetGeneration || '');
  const validGens = ['昭和', '平成', '令和', 'Z'];

  if (!text) return { ok: false, error: '翻訳したい言葉を入力してください。' };
  if (validGens.indexOf(targetGen) === -1) return { ok: false, error: '変換先の世代を選択してください。' };

  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');
  if (!apiKey) return { ok: false, error: '翻訳機能が設定されていません（管理者にお問い合わせください）。' };

  const prompt =
    'あなたは日本の世代別言葉づかい（昭和・平成・令和・Z世代）の翻訳家です。\n' +
    '次の言葉を「' + targetGen + '世代」風の言い回しに変換してください。\n' +
    '変換後の言葉のみを1〜2文程度で出力し、説明や前置きは不要です。\n\n' +
    '変換したい言葉: ' + text;

  const requestBody = {
    contents: [{ parts: [{ text: prompt }] }]
  };

  // 指定モデル → 汎用の最新モデル（順番に）で、該当があるまで試す
  const model = resolveGeminiModel_(body.model);
  const candidates = [model].concat(FALLBACK_GEMINI_MODELS.filter(function (m) { return m !== model; }));

  try {
    let usedModel = model;
    let res = null;
    for (let i = 0; i < candidates.length; i++) {
      usedModel = candidates[i];
      res = callGeminiModel_(usedModel, apiKey, requestBody);
      if (!isModelNotFound_(res.code, res.text)) break;
      console.warn('Geminiモデル「' + usedModel + '」に該当がないため次の候補に切り替えます');
    }

    let json;
    try {
      json = JSON.parse(res.text);
    } catch (parseErr) {
      return { ok: false, error: '翻訳処理でエラーが発生しました（HTTP ' + res.code + '）' };
    }

    if (res.code !== 200 || json.error) {
      const msg = json.error && json.error.message ? json.error.message : ('HTTP ' + res.code);
      return { ok: false, error: '翻訳処理でエラーが発生しました: ' + msg };
    }

    const candidate = json.candidates && json.candidates[0];
    const parts = (candidate && candidate.content && candidate.content.parts) || [];
    // 思考(thought)パートを除いた本文だけをつなげる
    const resultText = parts
      .filter(function (p) { return p && p.text && !p.thought; })
      .map(function (p) { return p.text; })
      .join('')
      .trim();

    if (resultText) {
      return { ok: true, result: resultText, model: usedModel, requestedModel: model, fallback: usedModel !== model };
    }
    return { ok: false, error: '翻訳結果の取得に失敗しました。' };
  } catch (err) {
    return { ok: false, error: '翻訳処理でエラーが発生しました: ' + err.message };
  }
}

// ===== Geminiモデル決定・呼び出し =====
/**
 * 使用するGeminiモデル名を決定する。
 * 画面での指定 → スクリプトプロパティ → 汎用の最新モデル の順。不正な名前は無視する。
 */
function resolveGeminiModel_(requested) {
  const req = String(requested || '').trim().toLowerCase();
  if (req && GEMINI_MODEL_PATTERN.test(req)) return req;

  const prop = String(PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || '').trim().toLowerCase();
  if (prop && GEMINI_MODEL_PATTERN.test(prop)) return prop;

  return FALLBACK_GEMINI_MODELS[0];
}

/**
 * Gemini APIを1モデルに対して呼び出す。混雑(429/500/503)時は待ち時間を伸ばしながら再試行する。
 */
function callGeminiModel_(model, apiKey, requestBody) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent';
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };

  let code = 0;
  let text = '';
  for (let attempt = 1; attempt <= GEMINI_MAX_ATTEMPTS; attempt++) {
    const response = UrlFetchApp.fetch(url, options);
    code = response.getResponseCode();
    text = response.getContentText();
    if (code === 200) break;
    if (GEMINI_RETRY_CODES.indexOf(code) === -1 || attempt === GEMINI_MAX_ATTEMPTS) break;
    Utilities.sleep(2000 * attempt);
  }
  return { code: code, text: text };
}

/**
 * モデルが存在しない（廃止・名前違い）ことを示す応答かどうか。
 */
function isModelNotFound_(code, text) {
  if (code === 404) return true;
  if (code === 400 && /model|NOT_FOUND|not found|not supported/i.test(text)) {
    return /not found|is not supported|NOT_FOUND|unknown model|invalid model/i.test(text);
  }
  return false;
}

// ===== イベント機能 =====
function handleListEvents_(body) {
  ensureHeaders_();
  const sheet = getSheet_(SHEET_EVENTS);
  const rows = getAllRows_(sheet);

  // 募集終了判定を都度更新
  const participantsSheet = getSheet_(SHEET_PARTICIPANTS);
  const participantRows = getAllRows_(participantsSheet);

  const events = rows.map(function (ev) {
    const joinedCount = participantRows
      .filter(function (p) { return p.eventId === ev.eventId && p.status === 'joined'; })
      .reduce(function (sum, p) { return sum + (Number(p.count) || 0); }, 0);

    let status = ev.status;
    const now = new Date();
    const deadline = ev.deadline ? new Date(ev.deadline) : null;
    const isExpired = deadline && now > deadline;
    const isFull = ev.capacity && joinedCount >= Number(ev.capacity);

    if (status !== 'cancelled') {
      status = (isExpired || isFull) ? 'closed' : 'open';
    }

    return {
      eventId: ev.eventId,
      title: ev.title,
      description: ev.description,
      datetime: ev.datetime,
      place: ev.place,
      capacity: ev.capacity,
      deadline: ev.deadline,
      ownerNickname: ev.ownerNickname,
      status: status,
      joinedCount: joinedCount,
      createdAt: ev.createdAt
    };
  });

  events.sort(function (a, b) {
    return new Date(b.createdAt) - new Date(a.createdAt);
  });

  return { ok: true, events: events };
}

function handleCreateEvent_(body) {
  const session = validateSession_(body.token);
  if (!session) return { ok: false, error: 'セッションが無効です。再度ログインしてください。' };

  const member = getMemberByEmail_(session.email);
  if (!member) return { ok: false, error: '会員情報が見つかりません。' };

  const title = String(body.title || '').trim();
  const description = String(body.description || '').trim();
  const datetime = String(body.datetime || '').trim();
  const place = String(body.place || '').trim();
  const capacity = Number(body.capacity || 0);
  const deadline = String(body.deadline || '').trim();

  if (!title || !datetime) {
    return { ok: false, error: 'イベント名と日時は必須です。' };
  }

  const eventId = Utilities.getUuid();
  const now = new Date();
  const sheet = getSheet_(SHEET_EVENTS);
  sheet.appendRow([eventId, title, description, datetime, place, capacity, deadline, member.email, member.nickname, 'open', now.toISOString()]);

  return { ok: true, eventId: eventId, message: 'イベントを登録しました。' };
}

function handleUpdateEvent_(body) {
  const session = validateSession_(body.token);
  if (!session) return { ok: false, error: 'セッションが無効です。再度ログインしてください。' };

  const member = getMemberByEmail_(session.email);
  if (!member) return { ok: false, error: '会員情報が見つかりません。' };

  const sheet = getSheet_(SHEET_EVENTS);
  const found = findRowByValue_(sheet, 0, body.eventId);
  if (!found) return { ok: false, error: 'イベントが見つかりません。' };

  const ownerEmail = found.row[7];
  if (ownerEmail !== member.email && member.role !== 'admin') {
    return { ok: false, error: 'このイベントを編集する権限がありません。' };
  }

  const rowIndex = found.rowIndex;
  if (body.title !== undefined) sheet.getRange(rowIndex, 2).setValue(String(body.title).trim());
  if (body.description !== undefined) sheet.getRange(rowIndex, 3).setValue(String(body.description).trim());
  if (body.datetime !== undefined) sheet.getRange(rowIndex, 4).setValue(String(body.datetime).trim());
  if (body.place !== undefined) sheet.getRange(rowIndex, 5).setValue(String(body.place).trim());
  if (body.capacity !== undefined) sheet.getRange(rowIndex, 6).setValue(Number(body.capacity));
  if (body.deadline !== undefined) sheet.getRange(rowIndex, 7).setValue(String(body.deadline).trim());

  return { ok: true, message: 'イベントを更新しました。' };
}

function handleDeleteEvent_(body) {
  const session = validateSession_(body.token);
  if (!session) return { ok: false, error: 'セッションが無効です。再度ログインしてください。' };

  const member = getMemberByEmail_(session.email);
  if (!member) return { ok: false, error: '会員情報が見つかりません。' };

  const sheet = getSheet_(SHEET_EVENTS);
  const found = findRowByValue_(sheet, 0, body.eventId);
  if (!found) return { ok: false, error: 'イベントが見つかりません。' };

  const ownerEmail = found.row[7];
  if (ownerEmail !== member.email && member.role !== 'admin') {
    return { ok: false, error: 'このイベントを削除する権限がありません。' };
  }

  sheet.deleteRow(found.rowIndex);

  // 関連する参加者データも削除
  const participantsSheet = getSheet_(SHEET_PARTICIPANTS);
  const data = participantsSheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === body.eventId) {
      participantsSheet.deleteRow(i + 1);
    }
  }

  return { ok: true, message: 'イベントを削除しました。' };
}

// ===== 参加・不参加 =====
function handleJoinEvent_(body) {
  const session = validateSession_(body.token);
  if (!session) return { ok: false, error: 'セッションが無効です。再度ログインしてください。' };

  const member = getMemberByEmail_(session.email);
  if (!member) return { ok: false, error: '会員情報が見つかりません。' };

  const eventId = body.eventId;
  const count = Math.max(1, Number(body.count || 1));

  const eventsSheet = getSheet_(SHEET_EVENTS);
  const eventFound = findRowByValue_(eventsSheet, 0, eventId);
  if (!eventFound) return { ok: false, error: 'イベントが見つかりません。' };

  const capacity = Number(eventFound.row[5]) || 0;
  const deadline = eventFound.row[6] ? new Date(eventFound.row[6]) : null;
  if (deadline && new Date() > deadline) {
    return { ok: false, error: '募集期限を過ぎているため参加できません。' };
  }

  const participantsSheet = getSheet_(SHEET_PARTICIPANTS);
  const participantRows = getAllRows_(participantsSheet);
  const currentJoined = participantRows
    .filter(function (p) { return p.eventId === eventId && p.status === 'joined' && p.email !== member.email; })
    .reduce(function (sum, p) { return sum + (Number(p.count) || 0); }, 0);

  if (capacity > 0 && currentJoined + count > capacity) {
    return { ok: false, error: '募集人数の上限を超えるため参加できません。' };
  }

  const existing = participantRows.filter(function (p) { return p.eventId === eventId && p.email === member.email; })[0];
  const now = new Date().toISOString();

  if (existing) {
    participantsSheet.getRange(existing._rowIndex, 3).setValue(member.nickname);
    participantsSheet.getRange(existing._rowIndex, 4).setValue(member.sns);
    participantsSheet.getRange(existing._rowIndex, 5).setValue(count);
    participantsSheet.getRange(existing._rowIndex, 6).setValue('joined');
    participantsSheet.getRange(existing._rowIndex, 7).setValue(now);
  } else {
    participantsSheet.appendRow([eventId, member.email, member.nickname, member.sns, count, 'joined', now]);
  }

  return { ok: true, message: '参加登録しました。' };
}

function handleLeaveEvent_(body) {
  const session = validateSession_(body.token);
  if (!session) return { ok: false, error: 'セッションが無効です。再度ログインしてください。' };

  const member = getMemberByEmail_(session.email);
  if (!member) return { ok: false, error: '会員情報が見つかりません。' };

  const eventId = body.eventId;
  const participantsSheet = getSheet_(SHEET_PARTICIPANTS);
  const participantRows = getAllRows_(participantsSheet);
  const existing = participantRows.filter(function (p) { return p.eventId === eventId && p.email === member.email; })[0];

  const now = new Date().toISOString();
  if (existing) {
    participantsSheet.getRange(existing._rowIndex, 6).setValue('declined');
    participantsSheet.getRange(existing._rowIndex, 7).setValue(now);
  } else {
    participantsSheet.appendRow([eventId, member.email, member.nickname, member.sns, 0, 'declined', now]);
  }

  return { ok: true, message: '不参加に登録しました。' };
}

function handleGetParticipants_(body) {
  const eventId = body.eventId;
  const participantsSheet = getSheet_(SHEET_PARTICIPANTS);
  const rows = getAllRows_(participantsSheet);

  const participants = rows
    .filter(function (p) { return p.eventId === eventId && p.status === 'joined'; })
    .map(function (p) {
      return { nickname: p.nickname, sns: p.sns, count: p.count };
    });

  return { ok: true, participants: participants };
}

// ===== 管理者機能 =====
function handleAdminLogin_(body) {
  const props = PropertiesService.getScriptProperties();
  const adminPasswordHash = props.getProperty('ADMIN_PASSWORD_HASH');
  const password = String(body.password || '');

  if (!adminPasswordHash || sha256_(password) !== adminPasswordHash) {
    return { ok: false, error: '管理者パスワードが正しくありません。' };
  }

  const token = generateToken_();
  const props2 = PropertiesService.getScriptProperties();
  const sessions = JSON.parse(props2.getProperty('ADMIN_SESSIONS') || '{}');
  sessions[token] = Date.now() + SESSION_TTL_MS;
  props2.setProperty('ADMIN_SESSIONS', JSON.stringify(sessions));

  return { ok: true, token: token };
}

function validateAdminSession_(token) {
  if (!token) return false;
  const props = PropertiesService.getScriptProperties();
  const sessions = JSON.parse(props.getProperty('ADMIN_SESSIONS') || '{}');
  const expiresAt = sessions[token];
  return expiresAt && Date.now() < expiresAt;
}

function handleAdminListMembers_(body) {
  if (!validateAdminSession_(body.adminToken)) {
    return { ok: false, error: '管理者セッションが無効です。' };
  }
  ensureHeaders_();
  const sheet = getSheet_(SHEET_MEMBERS);
  const rows = getAllRows_(sheet);
  const members = rows.map(function (m) {
    return { email: m.email, nickname: m.nickname, sns: m.sns, role: m.role, status: m.status, createdAt: m.createdAt };
  });
  return { ok: true, members: members };
}

function handleAdminUpdateMemberStatus_(body) {
  if (!validateAdminSession_(body.adminToken)) {
    return { ok: false, error: '管理者セッションが無効です。' };
  }
  const sheet = getSheet_(SHEET_MEMBERS);
  const found = findRowByValue_(sheet, 0, body.email);
  if (!found) return { ok: false, error: '会員が見つかりません。' };

  const newStatus = body.status === 'active' ? 'active' : 'suspended';
  sheet.getRange(found.rowIndex, 6).setValue(newStatus);

  return { ok: true, message: '会員ステータスを更新しました。' };
}

function handleAdminListEvents_(body) {
  if (!validateAdminSession_(body.adminToken)) {
    return { ok: false, error: '管理者セッションが無効です。' };
  }
  ensureHeaders_();
  const eventsSheet = getSheet_(SHEET_EVENTS);
  const eventRows = getAllRows_(eventsSheet);
  const participantsSheet = getSheet_(SHEET_PARTICIPANTS);
  const participantRows = getAllRows_(participantsSheet);

  const events = eventRows.map(function (ev) {
    const joinedCount = participantRows.filter(function (p) {
      return p.eventId === ev.eventId && p.status === 'joined';
    }).length;
    return {
      eventId: ev.eventId,
      title: ev.title,
      datetime: ev.datetime,
      place: ev.place,
      ownerEmail: ev.ownerEmail,
      ownerNickname: ev.ownerNickname,
      status: ev.status,
      joinedCount: joinedCount,
      createdAt: ev.createdAt,
    };
  });
  events.sort(function (a, b) { return new Date(b.createdAt) - new Date(a.createdAt); });

  return { ok: true, events: events };
}

function handleAdminDeleteEvent_(body) {
  if (!validateAdminSession_(body.adminToken)) {
    return { ok: false, error: '管理者セッションが無効です。' };
  }
  const sheet = getSheet_(SHEET_EVENTS);
  const found = findRowByValue_(sheet, 0, body.eventId);
  if (!found) return { ok: false, error: 'イベントが見つかりません。' };

  sheet.deleteRow(found.rowIndex);

  // 関連する参加者データも削除
  const participantsSheet = getSheet_(SHEET_PARTICIPANTS);
  const data = participantsSheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === body.eventId) {
      participantsSheet.deleteRow(i + 1);
    }
  }

  return { ok: true, message: 'イベントを強制削除しました。' };
}

/**
 * 初回セットアップ用: 管理者パスワードを設定するヘルパー。
 * スクリプトエディタから手動で一度だけ実行してください。
 * 実行後、このままだと平文がスクリプト内に残るため、実行後は値を消すか関数を削除すること。
 */
function setupAdminPassword_DO_NOT_DEPLOY() {
  const plainPassword = 'CHANGE_ME_TO_A_STRONG_PASSWORD';
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD_HASH', sha256_(plainPassword));
}
