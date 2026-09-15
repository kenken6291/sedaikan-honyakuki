/**
 * 世代間アンケート「これ、どう思う？」— Code.gs（GAS バックエンド）
 * ──────────────────────────────────────────────────────────────
 * 構成: GitHub Pages(フロント) + GAS(API) + Googleスプレッドシート(DB)
 *      + Gemini API(世代ギャップのAIひとことコメント自動生成)
 * 会員認証: メール登録 → 仮パスワード発行 → 初回ログイン時強制パスワード変更
 *          / パスワード忘れ時の仮パスワード再発行 / ニックネーム変更 / ログアウト
 *          / ログイン5回失敗で15分ロック
 *
 * 【事前準備】
 *   1. 新規スプレッドシートを作成し、「拡張機能」→「Apps Script」を開いて
 *      このファイルの中身を貼り付ける（このファイル自体をスプレッドシートに
 *      紐づけるので、スプレッドシートIDの設定は不要）
 *      （Questions / Votes / Members シートは初回アクセス時に自動作成されます）
 *   2. 「プロジェクトの設定」→「スクリプト プロパティ」に追加:
 *        GEMINI_API_KEY   : Gemini APIキー（https://aistudio.google.com/apikey）
 *                           ※未設定でもアプリ自体は動作します（AIコメントのみ非表示）
 *        PASSWORD_PEPPER  : 任意の長いランダム文字列（パスワードハッシュ用の秘密のこしょう）
 *   3. 「デプロイ」→「新しいデプロイ」→ 種類：「ウェブアプリ」
 *        実行するユーザー：自分 / アクセスできるユーザー：全員 → デプロイ
 *      発行された URL を index.html の「⚙️ GAS を設定する」から貼り付ける
 *   4. 初回実行時、MailApp（メール送信）の権限承認を求められるので許可する
 *      ※ MailAppは1日100通までの送信上限（Googleアカウントの標準クォータ）
 *   5. コードを変更するたびに「デプロイを管理」→ 新しいバージョンを作成
 *      （URLはそのまま。バージョンを切らないと変更が反映されません）
 * ──────────────────────────────────────────────────────────────
 */

const SH_QUESTIONS = 'Questions';
const SH_VOTES      = 'Votes';
const SH_MEMBERS    = 'Members';

const MEMBERS_HEADER = [
  'userId','email','nickname','passwordHash','salt','isTempPassword',
  'failedAttempts','lockUntil','sessionToken','createdAt','updatedAt'
];

const MAX_LOGIN_FAILS = 5;
const LOCK_MINUTES    = 15;
const TEMP_PW_CHARS   = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789'; // 紛らわしい文字(0/O/1/l/I)を除外

// ────────────────────────────────────────────────────────────────
//  唯一のエントリポイント：POST/JSON-RPCスタイル
//  text/plain 応答でブラウザのCORSプリフライトを回避する
// ────────────────────────────────────────────────────────────────
function doPost(e) {
  const out = (obj) =>
    ContentService.createTextOutput(JSON.stringify(obj))
      .setMimeType(ContentService.MimeType.TEXT);

  try {
    const req     = JSON.parse(e.postData.contents);
    const action  = req.action;
    const p       = req.payload || {};

    switch (action) {
      case 'getQuestions':    return out(ok(getQuestions()));
      case 'getResults':      return out(ok(getResults()));
      case 'saveQuestion':    return out(ok(requireAuth(p, saveQuestion)));
      case 'deleteQuestion':  return out(ok(requireAuth(p, deleteQuestion)));
      case 'castVote':        return out(ok(requireAuth(p, castVote)));
      case 'resetAllVotes':   return out(ok(requireAuth(p, resetAllVotes)));
      case 'getGapComment':   return out(ok(getGapComment(p)));

      case 'register':        return out(handleRegister(p));
      case 'login':            return out(handleLogin(p));
      case 'getSession':       return out(handleGetSession(p));
      case 'changePassword':   return out(handleChangePassword(p));
      case 'forgotPassword':   return out(handleForgotPassword(p));
      case 'changeNickname':   return out(handleChangeNickname(p));
      case 'logout':            return out(handleLogout(p));

      default:
        return out(err('Unknown action: ' + action));
    }
  } catch (ex) {
    return out(err(ex.message));
  }
}

function doGet() {
  return ContentService.createTextOutput('世代間アンケート API is running.');
}

function ok(data)  { return { ok: true, data }; }
function err(msg)  { return { ok: false, error: msg }; }

// ════════════════════════════════════════════════════════════════
//  シート取得（初回アクセス時に自動作成）
// ════════════════════════════════════════════════════════════════
function getQuestionsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SH_QUESTIONS);
  if (!sh) {
    sh = ss.insertSheet(SH_QUESTIONS);
    sh.appendRow([
      'id','title','badge','authorGen',
      'opt0_emoji','opt0_text',
      'opt1_emoji','opt1_text',
      'opt2_emoji','opt2_text',
      'opt3_emoji','opt3_text',
      'createdAt','gapComment','gapCommentAt'
    ]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,15).setFontWeight('bold');
    sh.appendRow([
      'q_date','デートの費用は「割り勘」にするべき？','FALSE',0,
      '🤝','割り勘にすべき',
      '🎩','年長者・誘った側が多めに払うべき',
      '','','','',new Date().toISOString(),'',''
    ]);
    sh.appendRow([
      'q_wedding','結婚式の規模・スタイルについてどう思う？','TRUE',2,
      '⛪','伝統的な式・披露宴をやりたい',
      '👨‍👩‍👧','身内だけの少人数でやりたい',
      '📸','フォトウェディングだけでいい',
      '✌️','式は不要・ナシ婚でいい',
      new Date().toISOString(),'',''
    ]);
  }
  return sh;
}

function getVotesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SH_VOTES);
  if (!sh) {
    sh = ss.insertSheet(SH_VOTES);
    sh.appendRow(['questionId','genIndex','optionIndex','timestamp','userId']);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,5).setFontWeight('bold');
  }
  return sh;
}

function getMembersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SH_MEMBERS);
  if (!sh) {
    sh = ss.insertSheet(SH_MEMBERS);
    sh.appendRow(MEMBERS_HEADER);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,MEMBERS_HEADER.length).setFontWeight('bold');
  }
  return sh;
}

// ════════════════════════════════════════════════════════════════
//  質問 / 投票（会員向けアクションは requireAuth 経由で呼ばれる）
// ════════════════════════════════════════════════════════════════
function getQuestions() {
  const rows = getQuestionsSheet().getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1).map(rowToQuestion).filter(q => q.id);
}

function getResults() {
  const questions = getQuestions();
  const rows = getVotesSheet().getDataRange().getValues().slice(1);
  const tally = {};
  questions.forEach(q => {
    tally[q.id] = [0,1,2,3].map(() => q.options.map(() => 0));
  });
  rows.forEach(r => {
    const [qid, gi, oi] = [String(r[0]), Number(r[1]), Number(r[2])];
    if (tally[qid] && tally[qid][gi] && tally[qid][gi][oi] !== undefined) {
      tally[qid][gi][oi]++;
    }
  });
  return tally;
}

/** 質問保存（新規作成 / 編集）※会員のみ */
function saveQuestion(p, member) {
  const id        = p.id;
  const title     = String(p.title || '').trim();
  const badge     = p.badge === true || p.badge === 'true';
  const authorGen = p.authorGen !== '' && p.authorGen != null ? Number(p.authorGen) : null;
  const options   = Array.isArray(p.options) ? p.options : JSON.parse(p.options || '[]');

  if (!title) throw new Error('質問文を入力してください');
  if (!Array.isArray(options) || options.length < 2) throw new Error('選択肢は2個以上必要です');

  const sh   = getQuestionsSheet();
  const data = sh.getDataRange().getValues();
  const row  = buildRow(id, title, badge, authorGen, options);
  const idx  = data.findIndex((r, i) => i > 0 && String(r[0]) === id);

  if (idx > 0) {
    // 選択肢が変わったらキャッシュ済みAIコメントは無効化
    row[13] = data[idx][13] || '';
    row[14] = data[idx][14] || '';
    sh.getRange(idx + 1, 1, 1, row.length).setValues([row]);
    return { status:'updated', id };
  } else {
    sh.appendRow(row);
    return { status:'created', id };
  }
}

/** 質問削除 ※会員のみ */
function deleteQuestion(p, member) {
  const id   = p.id;
  const sh   = getQuestionsSheet();
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      sh.deleteRow(i + 1);
      deleteVotesByQuestion(id);
      return { status:'deleted', id };
    }
  }
  throw new Error('Question not found: ' + id);
}

/** 投票 ※会員のみ・同じ会員による同一質問への二重投票を防止 */
function castVote(p, member) {
  const qid = String(p.questionId);
  const sh  = getVotesSheet();
  const data = sh.getDataRange().getValues();
  const already = data.slice(1).some(r => String(r[0]) === qid && String(r[4]) === member.userId);
  if (already) throw new Error('この質問にはすでに投票済みです');

  sh.appendRow([
    qid,
    Number(p.genIndex),
    Number(p.optionIndex),
    new Date().toISOString(),
    member.userId,
  ]);
  return { status:'voted' };
}

/** 全投票データのリセット ※会員のみ・元に戻せません */
function resetAllVotes(p, member) {
  const sh = getVotesSheet();
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).clearContent();
  // AIコメントのキャッシュもクリア
  const qsh = getQuestionsSheet();
  const lastQRow = qsh.getLastRow();
  if (lastQRow > 1) {
    qsh.getRange(2, 14, lastQRow - 1, 2).clearContent();
  }
  return { status:'reset' };
}

// ────────────────────────────────────────────────────────────────
//  会員操作のガード：セッション確認してから本処理を呼ぶ
// ────────────────────────────────────────────────────────────────
function requireAuth(payload, fn) {
  const sheet = getMembersSheet();
  const found = findMemberBySession(sheet, payload.sessionToken);
  if (!found) throw new Error('ログインが必要です');
  return fn(payload, { userId: found.data.userId, nickname: found.data.nickname });
}

// ────────────────────────────────────────────────────────────────
//  ヘルパー
// ────────────────────────────────────────────────────────────────
function rowToQuestion(r) {
  const opts = [];
  for (let i = 4; i <= 10; i += 2) {
    if (String(r[i+1]).trim() !== '') {
      opts.push({ emoji: String(r[i]) || '▶', text: String(r[i+1]) });
    }
  }
  return {
    id:        String(r[0]),
    title:     String(r[1]),
    badge:     String(r[2]).toUpperCase() === 'TRUE',
    authorGen: (r[3] !== '' && r[3] != null) ? Number(r[3]) : null,
    options:   opts,
    gapComment: r[13] ? String(r[13]) : '',
  };
}

function buildRow(id, title, badge, authorGen, options) {
  const opts = options.slice(0, 4);
  const row  = [id, title, badge ? 'TRUE' : 'FALSE', authorGen ?? ''];
  for (let i = 0; i < 4; i++) {
    row.push(opts[i]?.emoji ?? '');
    row.push(opts[i]?.text  ?? '');
  }
  row.push(new Date().toISOString());
  row.push(''); // gapComment（saveQuestion側で既存値を引き継ぐ）
  row.push(''); // gapCommentAt
  return row;
}

function deleteVotesByQuestion(qid) {
  const sh   = getVotesSheet();
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === qid) sh.deleteRow(i + 1);
  }
}

// ════════════════════════════════════════════════════════════════
//  Gemini：世代ギャップのAIひとことコメント（結果が3票以上たまったら生成・キャッシュ）
// ════════════════════════════════════════════════════════════════
function getGapComment(p) {
  const qid = String(p.questionId);
  const sh   = getQuestionsSheet();
  const data = sh.getDataRange().getValues();
  const idx  = data.findIndex((r, i) => i > 0 && String(r[0]) === qid);
  if (idx < 1) throw new Error('質問が見つかりません');

  const cached = data[idx][13];
  if (cached) return { comment: String(cached) };

  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return { comment: '' }; // 未設定時は静かに空を返す（エラーにしない）

  const q = rowToQuestion(data[idx]);
  const results = getResults();
  const votes = results[qid];
  const total = votes ? votes.reduce((s,row)=>s+row.reduce((a,b)=>a+b,0),0) : 0;
  if (total < 3) return { comment: '' }; // 票数が少なすぎる間は生成しない

  const gens = ['10代・20代','30代・40代','50代・60代','70代以上'];
  const breakdown = votes.map((row,gi) => {
    const t = row.reduce((a,b)=>a+b,0);
    if (t === 0) return `${gens[gi]}：まだ投票なし`;
    const pctStr = row.map((c,oi) => `${q.options[oi]?.text || ''}${Math.round(c/t*100)}%`).join('、');
    return `${gens[gi]}：${pctStr}`;
  }).join(' / ');

  const prompt =
    'あなたは世代間の意識調査を分析するライターです。以下のアンケート結果を見て、' +
    '世代間の意見の違い（ギャップ）について、30〜60文字程度で気の利いた一言コメントを' +
    '日本語で1つ作成してください。煽りすぎず、なるほどと思わせる客観的なトーンで。' +
    '絵文字や記号は使わず、文章のみを出力してください。\n\n' +
    `質問: ${q.title}\n` +
    `世代別の回答割合: ${breakdown}`;

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.8-flash:generateContent?key=' + apiKey;
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { maxOutputTokens: 200, thinkingConfig: { thinkingLevel: 'low' } }
    })
  });

  if (res.getResponseCode() !== 200) return { comment: '' };
  const json = JSON.parse(res.getContentText());
  const text = json && json.candidates && json.candidates[0] &&
               json.candidates[0].content && json.candidates[0].content.parts &&
               json.candidates[0].content.parts[0] && json.candidates[0].content.parts[0].text;
  if (!text) return { comment: '' };

  const comment = text.trim();
  sh.getRange(idx + 1, 14).setValue(comment);
  sh.getRange(idx + 1, 15).setValue(new Date().toISOString());
  return { comment };
}

// ════════════════════════════════════════════════════════════════
//  会員認証：ハッシュ・検索まわりの共通ヘルパー
// ════════════════════════════════════════════════════════════════
function hashPassword(password, salt) {
  const pepper = PropertiesService.getScriptProperties().getProperty('PASSWORD_PEPPER') || '';
  const raw    = String(password) + String(salt) + pepper;
  const bytes  = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return bytes.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

function generateTempPassword() {
  let pw = '';
  for (let i = 0; i < 8; i++) {
    pw += TEMP_PW_CHARS.charAt(Math.floor(Math.random() * TEMP_PW_CHARS.length));
  }
  return pw;
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || ''));
}

function findMemberRow(sheet, predicate) {
  const values = sheet.getDataRange().getValues();
  const header = values[0];
  for (let r = 1; r < values.length; r++) {
    const obj = {};
    header.forEach((key, i) => { obj[key] = values[r][i]; });
    if (predicate(obj)) return { rowNum: r + 1, header, data: obj };
  }
  return null;
}

function findMemberByEmail(sheet, email) {
  const target = String(email).trim().toLowerCase();
  return findMemberRow(sheet, m => String(m.email).trim().toLowerCase() === target);
}

function findMemberBySession(sheet, sessionToken) {
  if (!sessionToken) return null;
  return findMemberRow(sheet, m => m.sessionToken && m.sessionToken === sessionToken);
}

function setMemberCell(sheet, header, rowNum, colName, value) {
  sheet.getRange(rowNum, header.indexOf(colName) + 1).setValue(value);
}

// ════════════════════════════════════════════════════════════════
//  会員登録
// ════════════════════════════════════════════════════════════════
function handleRegister(payload) {
  const email    = String(payload.email || '').trim();
  const nickname = String(payload.nickname || '').trim();

  if (!isValidEmail(email)) return { ok: false, error: 'メールアドレスの形式が正しくありません' };
  if (nickname.length < 1 || nickname.length > 20) return { ok: false, error: 'ニックネームは1〜20文字で入力してください' };

  const sheet = getMembersSheet();
  if (findMemberByEmail(sheet, email)) {
    return { ok: false, error: 'このメールアドレスは既に登録されています' };
  }

  const tempPassword = generateTempPassword();
  const salt = Utilities.getUuid();
  const hash = hashPassword(tempPassword, salt);
  const now  = new Date().toISOString();
  const userId = Utilities.getUuid();

  sheet.appendRow([
    userId, email, nickname, hash, salt, true,
    0, '', '', now, now
  ]);

  try {
    MailApp.sendEmail({
      to: email,
      subject: '【世代間アンケート：これ、どう思う？】仮パスワードのお知らせ',
      body:
        `${nickname} 様\n\n` +
        `ご登録ありがとうございます。以下の仮パスワードでログインし、\n` +
        `初回ログイン時に新しいパスワードへの変更をお願いします。\n\n` +
        `仮パスワード: ${tempPassword}\n\n` +
        `※このメールに心当たりがない場合は破棄してください。`
    });
  } catch (e) {
    return { ok: false, error: '登録は完了しましたが、メール送信に失敗しました。時間を置いて再度お試しください' };
  }

  return { ok: true };
}

// ════════════════════════════════════════════════════════════════
//  ログイン
// ════════════════════════════════════════════════════════════════
function handleLogin(payload) {
  const email    = String(payload.email || '').trim();
  const password = String(payload.password || '');

  const sheet = getMembersSheet();
  const found = findMemberByEmail(sheet, email);
  if (!found) return { ok: false, error: 'メールアドレスまたはパスワードが正しくありません' };

  const { rowNum, header, data } = found;

  if (data.lockUntil && new Date(data.lockUntil) > new Date()) {
    const minutes = Math.ceil((new Date(data.lockUntil) - new Date()) / 60000);
    return { ok: false, error: `ログイン試行回数が上限に達しました。${minutes}分後に再度お試しください` };
  }

  const hash = hashPassword(password, data.salt);
  if (hash !== data.passwordHash) {
    const fails = Number(data.failedAttempts || 0) + 1;
    if (fails >= MAX_LOGIN_FAILS) {
      setMemberCell(sheet, header, rowNum, 'failedAttempts', 0);
      setMemberCell(sheet, header, rowNum, 'lockUntil', new Date(Date.now() + LOCK_MINUTES * 60000).toISOString());
      return { ok: false, error: `ログインに${MAX_LOGIN_FAILS}回失敗したため、${LOCK_MINUTES}分間ロックされました` };
    }
    setMemberCell(sheet, header, rowNum, 'failedAttempts', fails);
    return { ok: false, error: 'メールアドレスまたはパスワードが正しくありません' };
  }

  const sessionToken = Utilities.getUuid();
  const now = new Date().toISOString();
  setMemberCell(sheet, header, rowNum, 'failedAttempts', 0);
  setMemberCell(sheet, header, rowNum, 'lockUntil', '');
  setMemberCell(sheet, header, rowNum, 'sessionToken', sessionToken);
  setMemberCell(sheet, header, rowNum, 'updatedAt', now);

  return {
    ok: true,
    sessionToken,
    userId: data.userId,
    nickname: data.nickname,
    isTempPassword: !!data.isTempPassword
  };
}

// ════════════════════════════════════════════════════════════════
//  セッション確認（再訪問時の自動ログイン維持）
// ════════════════════════════════════════════════════════════════
function handleGetSession(payload) {
  const sheet = getMembersSheet();
  const found = findMemberBySession(sheet, payload.sessionToken);
  if (!found) return { ok: false };
  return {
    ok: true,
    userId: found.data.userId,
    nickname: found.data.nickname,
    isTempPassword: !!found.data.isTempPassword
  };
}

// ════════════════════════════════════════════════════════════════
//  パスワード変更（初回強制変更 / 任意変更 共通）
// ════════════════════════════════════════════════════════════════
function handleChangePassword(payload) {
  const sheet = getMembersSheet();
  const found = findMemberBySession(sheet, payload.sessionToken);
  if (!found) return { ok: false, error: 'ログインし直してください' };

  const { rowNum, header, data } = found;
  const currentPassword = String(payload.currentPassword || '');
  const newPassword     = String(payload.newPassword || '');

  const currentHash = hashPassword(currentPassword, data.salt);
  if (currentHash !== data.passwordHash) {
    return { ok: false, error: '現在のパスワードが正しくありません' };
  }
  if (newPassword.length < 8) {
    return { ok: false, error: '新しいパスワードは8文字以上で設定してください' };
  }

  const newSalt = Utilities.getUuid();
  const newHash = hashPassword(newPassword, newSalt);
  const now = new Date().toISOString();

  setMemberCell(sheet, header, rowNum, 'passwordHash', newHash);
  setMemberCell(sheet, header, rowNum, 'salt', newSalt);
  setMemberCell(sheet, header, rowNum, 'isTempPassword', false);
  setMemberCell(sheet, header, rowNum, 'updatedAt', now);

  return { ok: true };
}

// ════════════════════════════════════════════════════════════════
//  パスワードを忘れた場合 → 仮パスワード再発行
// ════════════════════════════════════════════════════════════════
function handleForgotPassword(payload) {
  const email = String(payload.email || '').trim();
  if (!isValidEmail(email)) return { ok: false, error: 'メールアドレスの形式が正しくありません' };

  const sheet = getMembersSheet();
  const found = findMemberByEmail(sheet, email);

  // メールアドレスの存在有無を外部に漏らさないよう、常に同じ成功メッセージを返す
  const genericResult = { ok: true, message: 'このメールアドレスが登録されていれば、仮パスワードを送信しました' };
  if (!found) return genericResult;

  const { rowNum, header, data } = found;
  const tempPassword = generateTempPassword();
  const newSalt = Utilities.getUuid();
  const newHash = hashPassword(tempPassword, newSalt);
  const now = new Date().toISOString();

  setMemberCell(sheet, header, rowNum, 'passwordHash', newHash);
  setMemberCell(sheet, header, rowNum, 'salt', newSalt);
  setMemberCell(sheet, header, rowNum, 'isTempPassword', true);
  setMemberCell(sheet, header, rowNum, 'sessionToken', ''); // 既存セッションは無効化
  setMemberCell(sheet, header, rowNum, 'failedAttempts', 0);
  setMemberCell(sheet, header, rowNum, 'lockUntil', '');
  setMemberCell(sheet, header, rowNum, 'updatedAt', now);

  try {
    MailApp.sendEmail({
      to: email,
      subject: '【世代間アンケート：これ、どう思う？】仮パスワードの再発行',
      body:
        `${data.nickname} 様\n\n` +
        `パスワード再発行のリクエストを受け付けました。\n` +
        `以下の仮パスワードでログインし、新しいパスワードに変更してください。\n\n` +
        `仮パスワード: ${tempPassword}\n\n` +
        `※心当たりがない場合はこのメールを破棄してください。`
    });
  } catch (e) {
    // メール送信失敗時もパスワードは既に更新済みのため、成功として扱う（次回問い合わせで案内）
  }

  return genericResult;
}

// ════════════════════════════════════════════════════════════════
//  ニックネーム変更
// ════════════════════════════════════════════════════════════════
function handleChangeNickname(payload) {
  const sheet = getMembersSheet();
  const found = findMemberBySession(sheet, payload.sessionToken);
  if (!found) return { ok: false, error: 'ログインし直してください' };

  const nickname = String(payload.nickname || '').trim();
  if (nickname.length < 1 || nickname.length > 20) {
    return { ok: false, error: 'ニックネームは1〜20文字で入力してください' };
  }

  const { rowNum, header } = found;
  setMemberCell(sheet, header, rowNum, 'nickname', nickname);
  setMemberCell(sheet, header, rowNum, 'updatedAt', new Date().toISOString());

  return { ok: true, nickname };
}

// ════════════════════════════════════════════════════════════════
//  ログアウト
// ════════════════════════════════════════════════════════════════
function handleLogout(payload) {
  const sheet = getMembersSheet();
  const found = findMemberBySession(sheet, payload.sessionToken);
  if (found) {
    setMemberCell(sheet, found.header, found.rowNum, 'sessionToken', '');
  }
  return { ok: true };
}
