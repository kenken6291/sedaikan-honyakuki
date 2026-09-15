/**
 * auth.js — 会員認証（登録・ログイン・パスワード変更・ログアウト）
 * ──────────────────────────────────────────────────────────────
 * index.html のインラインスクリプトより先に読み込むこと。
 * GAS の接続先URLは、このアプリの「⚙️ GAS を設定する」で保存した
 * localStorage の値（sedai_gas_url）を共有して使う。
 *
 * 他スクリプトからは window.Auth 経由で参照する:
 *   Auth.isLoggedIn()      → true/false
 *   Auth.getUserId()       → ログイン中ユーザーのID（未ログイン時はnull）
 *   Auth.getNickname()     → ニックネーム
 *   Auth.getSessionToken() → GAS APIに渡すセッショントークン
 *   Auth.requireLogin()    → 未ログインならログインモーダルを開いてfalseを返す
 * ──────────────────────────────────────────────────────────────
 */

"use strict";

const LS_GAS_URL_AUTH   = "sedai_gas_url";        // index.html 側と同じキー
const SESSION_TOKEN_KEY = "sedai_session_token_v1";

const Auth = (() => {
  let sessionToken = localStorage.getItem(SESSION_TOKEN_KEY) || null;
  let nickname     = null;
  let userId       = null;
  let onLoginChangeCallbacks = [];

  function getGasUrl() {
    return localStorage.getItem(LS_GAS_URL_AUTH) || "";
  }

  /* ---------- GAS APIラッパー（index.htmlのgasCallと共通仕様） ---------- */
  async function callApi(action, payload) {
    const gasUrl = getGasUrl();
    if (!gasUrl) throw new Error("GAS Web App の URL が未設定です");
    const res = await fetch(gasUrl, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ action, payload: payload || {} })
    });
    if (!res.ok) throw new Error("HTTP " + res.status);
    return res.json();
  }

  function notifyLoginChange() {
    onLoginChangeCallbacks.forEach(cb => cb());
  }

  /* ---------- 初期化: 保存済みトークンの検証 ---------- */
  async function init() {
    bindEvents();
    if (!getGasUrl() || !sessionToken) { renderHeader(); return; }
    try {
      const res = await callApi("getSession", { sessionToken });
      if (res.ok) {
        nickname = res.nickname;
        userId   = res.userId;
        renderHeader();
        notifyLoginChange();
        if (res.isTempPassword) openForceChangeModal();
      } else {
        clearSession();
      }
    } catch (e) {
      console.warn("セッション確認に失敗:", e);
    }
  }

  function clearSession() {
    sessionToken = null; nickname = null; userId = null;
    localStorage.removeItem(SESSION_TOKEN_KEY);
    renderHeader();
    notifyLoginChange();
  }

  function setSession(token, nick) {
    sessionToken = token; nickname = nick;
    localStorage.setItem(SESSION_TOKEN_KEY, token);
    renderHeader();
    notifyLoginChange();
  }

  /* ---------- ヘッダー表示切り替え ---------- */
  function renderHeader() {
    const guest = document.getElementById("authGuestArea");
    const user  = document.getElementById("authUserArea");
    if (!guest || !user) return;
    if (sessionToken) {
      guest.hidden = true;
      user.hidden  = false;
      document.getElementById("userNicknameLabel").textContent = nickname || "";
    } else {
      guest.hidden = false;
      user.hidden  = true;
    }
  }

  /* ============================================================
     パスワード表示/非表示トグル（共通ユーティリティ）
     ============================================================ */
  function bindPasswordToggle(toggleId, inputId) {
    const toggle = document.getElementById(toggleId);
    const input  = document.getElementById(inputId);
    if (!toggle || !input) return;
    toggle.addEventListener("click", () => {
      const show = input.type === "password";
      input.type = show ? "text" : "password";
      toggle.textContent = show ? "🙈 隠す" : "👁 表示";
    });
  }

  /* ============================================================
     認証モーダル（ログイン / 登録タブ）
     ============================================================ */
  function openAuthModal(tab = "login") {
    switchAuthTab(tab);
    clearAuthErrors();
    openModal("authModal");
  }
  function closeAuthModalFn() { closeModal("authModal"); }

  function switchAuthTab(tab) {
    document.getElementById("authTabLogin").classList.toggle("active", tab === "login");
    document.getElementById("authTabRegister").classList.toggle("active", tab === "register");
    document.getElementById("loginForm").hidden    = tab !== "login";
    document.getElementById("registerForm").hidden = tab !== "register";
  }

  function clearAuthErrors() {
    ["loginError","registerError","forgotError","forceChangeError","accountError"].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.textContent = "";
    });
  }

  async function handleLoginSubmit(e) {
    e.preventDefault();
    const email    = document.getElementById("loginEmail").value.trim();
    const password = document.getElementById("loginPassword").value;
    const btn = document.getElementById("btnLoginSubmit");
    const errEl = document.getElementById("loginError");
    errEl.textContent = "";
    btn.disabled = true; btn.textContent = "⏳ ログイン中...";

    try {
      const res = await callApi("login", { email, password });
      if (!res.ok) { errEl.textContent = res.error || "ログインに失敗しました"; return; }
      userId = res.userId;
      setSession(res.sessionToken, res.nickname);
      closeAuthModalFn();
      document.getElementById("loginForm").reset();
      if (res.isTempPassword) {
        openForceChangeModal();
      } else {
        showToast(`👋 おかえりなさい、${res.nickname}さん`);
      }
    } catch (err) {
      errEl.textContent = err.message || "通信エラーが発生しました";
    } finally {
      btn.disabled = false; btn.textContent = "ログイン";
    }
  }

  async function handleRegisterSubmit(e) {
    e.preventDefault();
    const email = document.getElementById("registerEmail").value.trim();
    const nick  = document.getElementById("registerNickname").value.trim();
    const btn = document.getElementById("btnRegisterSubmit");
    const errEl = document.getElementById("registerError");
    errEl.textContent = "";
    btn.disabled = true; btn.textContent = "⏳ 送信中...";

    try {
      const res = await callApi("register", { email, nickname: nick });
      if (!res.ok) { errEl.textContent = res.error || "登録に失敗しました"; return; }
      document.getElementById("registerForm").reset();
      showToast("📩 仮パスワードをメールで送信しました。届いたパスワードでログインしてください");
      switchAuthTab("login");
      document.getElementById("loginEmail").value = email;
    } catch (err) {
      errEl.textContent = err.message || "通信エラーが発生しました";
    } finally {
      btn.disabled = false; btn.textContent = "仮パスワードをメールで受け取る";
    }
  }

  /* ============================================================
     パスワードを忘れた場合
     ============================================================ */
  function openForgotModal() {
    document.getElementById("forgotEmail").value = document.getElementById("loginEmail").value;
    document.getElementById("forgotError").textContent = "";
    document.getElementById("forgotSuccess").hidden = true;
    closeModal("authModal");
    openModal("forgotModal");
  }

  async function handleForgotSubmit(e) {
    e.preventDefault();
    const email = document.getElementById("forgotEmail").value.trim();
    const btn = document.getElementById("btnForgotSubmit");
    const errEl = document.getElementById("forgotError");
    errEl.textContent = "";
    btn.disabled = true; btn.textContent = "⏳ 送信中...";

    try {
      const res = await callApi("forgotPassword", { email });
      if (!res.ok) { errEl.textContent = res.error || "送信に失敗しました"; return; }
      document.getElementById("forgotForm").hidden = true;
      document.getElementById("forgotSuccess").hidden = false;
    } catch (err) {
      errEl.textContent = err.message || "通信エラーが発生しました";
    } finally {
      btn.disabled = false; btn.textContent = "仮パスワードを再発行する";
    }
  }

  /* ============================================================
     初回ログイン時の強制パスワード変更
     ============================================================ */
  function openForceChangeModal() {
    document.getElementById("forceChangeError").textContent = "";
    document.getElementById("forceCurrentPassword").value = "";
    document.getElementById("forceNewPassword").value = "";
    document.getElementById("forceNewPasswordConfirm").value = "";
    openModal("forceChangeModal"); // 閉じるボタンなし＝完了するまで閉じられない
  }

  async function handleForceChangeSubmit(e) {
    e.preventDefault();
    const current = document.getElementById("forceCurrentPassword").value;
    const next1   = document.getElementById("forceNewPassword").value;
    const next2   = document.getElementById("forceNewPasswordConfirm").value;
    const errEl = document.getElementById("forceChangeError");
    errEl.textContent = "";

    if (next1 !== next2) { errEl.textContent = "新しいパスワードが一致しません"; return; }

    const btn = document.getElementById("btnForceChangeSubmit");
    btn.disabled = true; btn.textContent = "⏳ 変更中...";
    try {
      const res = await callApi("changePassword", {
        sessionToken, currentPassword: current, newPassword: next1
      });
      if (!res.ok) { errEl.textContent = res.error || "変更に失敗しました"; return; }
      closeModal("forceChangeModal");
      showToast("✅ パスワードを変更しました");
    } catch (err) {
      errEl.textContent = err.message || "通信エラーが発生しました";
    } finally {
      btn.disabled = false; btn.textContent = "パスワードを変更する";
    }
  }

  /* ============================================================
     アカウント設定（ニックネーム変更・パスワード変更）
     ============================================================ */
  function openAccountModal() {
    document.getElementById("accountError").textContent = "";
    document.getElementById("accountNickname").value = nickname || "";
    document.getElementById("accountCurrentPassword").value = "";
    document.getElementById("accountNewPassword").value = "";
    document.getElementById("accountNewPasswordConfirm").value = "";
    openModal("accountModal");
  }

  async function handleAccountNicknameSubmit(e) {
    e.preventDefault();
    const nick = document.getElementById("accountNickname").value.trim();
    const errEl = document.getElementById("accountError");
    errEl.textContent = "";
    const btn = document.getElementById("btnAccountNicknameSubmit");
    btn.disabled = true; btn.textContent = "⏳ 保存中...";
    try {
      const res = await callApi("changeNickname", { sessionToken, nickname: nick });
      if (!res.ok) { errEl.textContent = res.error || "変更に失敗しました"; return; }
      nickname = res.nickname;
      renderHeader();
      notifyLoginChange();
      showToast("✅ ニックネームを変更しました");
    } catch (err) {
      errEl.textContent = err.message || "通信エラーが発生しました";
    } finally {
      btn.disabled = false; btn.textContent = "ニックネームを保存";
    }
  }

  async function handleAccountPasswordSubmit(e) {
    e.preventDefault();
    const current = document.getElementById("accountCurrentPassword").value;
    const next1   = document.getElementById("accountNewPassword").value;
    const next2   = document.getElementById("accountNewPasswordConfirm").value;
    const errEl = document.getElementById("accountError");
    errEl.textContent = "";

    if (!current || !next1) { errEl.textContent = "現在のパスワードと新しいパスワードを入力してください"; return; }
    if (next1 !== next2) { errEl.textContent = "新しいパスワードが一致しません"; return; }

    const btn = document.getElementById("btnAccountPasswordSubmit");
    btn.disabled = true; btn.textContent = "⏳ 変更中...";
    try {
      const res = await callApi("changePassword", { sessionToken, currentPassword: current, newPassword: next1 });
      if (!res.ok) { errEl.textContent = res.error || "変更に失敗しました"; return; }
      document.getElementById("accountCurrentPassword").value = "";
      document.getElementById("accountNewPassword").value = "";
      document.getElementById("accountNewPasswordConfirm").value = "";
      showToast("✅ パスワードを変更しました");
    } catch (err) {
      errEl.textContent = err.message || "通信エラーが発生しました";
    } finally {
      btn.disabled = false; btn.textContent = "パスワードを変更";
    }
  }

  /* ============================================================
     ログアウト
     ============================================================ */
  async function handleLogout() {
    try { await callApi("logout", { sessionToken }); } catch (e) { /* 失敗してもローカルはクリアする */ }
    clearSession();
    showToast("👋 ログアウトしました");
  }

  /* ============================================================
     未ログイン時のガード
     ============================================================ */
  function requireLogin() {
    if (sessionToken) return true;
    showToast("⚠️ ログインが必要です");
    openAuthModal("login");
    return false;
  }

  /* ============================================================
     イベントバインド
     ============================================================ */
  function bindEvents() {
    document.getElementById("btnOpenAuth").addEventListener("click", () => openAuthModal("login"));
    document.getElementById("btnCloseAuth").addEventListener("click", closeAuthModalFn);
    document.getElementById("authModal").addEventListener("click", (e) => {
      if (e.target === e.currentTarget) closeAuthModalFn();
    });
    document.getElementById("authTabLogin").addEventListener("click", () => switchAuthTab("login"));
    document.getElementById("authTabRegister").addEventListener("click", () => switchAuthTab("register"));
    document.getElementById("loginForm").addEventListener("submit", handleLoginSubmit);
    document.getElementById("registerForm").addEventListener("submit", handleRegisterSubmit);
    document.getElementById("linkForgotPassword").addEventListener("click", (e) => { e.preventDefault(); openForgotModal(); });

    document.getElementById("btnCloseForgot").addEventListener("click", () => closeModal("forgotModal"));
    document.getElementById("forgotModal").addEventListener("click", (e) => {
      if (e.target === e.currentTarget) closeModal("forgotModal");
    });
    document.getElementById("forgotForm").addEventListener("submit", handleForgotSubmit);

    // 強制変更モーダルは閉じるボタンを設けない（完了するまで操作させる）
    document.getElementById("forceChangeForm").addEventListener("submit", handleForceChangeSubmit);

    document.getElementById("btnAccount").addEventListener("click", openAccountModal);
    document.getElementById("btnCloseAccount").addEventListener("click", () => closeModal("accountModal"));
    document.getElementById("accountModal").addEventListener("click", (e) => {
      if (e.target === e.currentTarget) closeModal("accountModal");
    });
    document.getElementById("accountNicknameForm").addEventListener("submit", handleAccountNicknameSubmit);
    document.getElementById("accountPasswordForm").addEventListener("submit", handleAccountPasswordSubmit);

    document.getElementById("btnLogout").addEventListener("click", handleLogout);

    bindPasswordToggle("btnToggleLoginPw", "loginPassword");
    bindPasswordToggle("btnToggleForceCurrentPw", "forceCurrentPassword");
    bindPasswordToggle("btnToggleForceNewPw", "forceNewPassword");
    bindPasswordToggle("btnToggleForceNewPwConfirm", "forceNewPasswordConfirm");
    bindPasswordToggle("btnToggleAccountCurrentPw", "accountCurrentPassword");
    bindPasswordToggle("btnToggleAccountNewPw", "accountNewPassword");
    bindPasswordToggle("btnToggleAccountNewPwConfirm", "accountNewPasswordConfirm");
  }

  /* ============================================================
     共通モーダル開閉ヘルパー（このアプリの .modal-bg / .open 方式に合わせる）
     ============================================================ */
  function openModal(id, focusId) {
    document.getElementById(id).classList.add("open");
    if (focusId) setTimeout(() => document.getElementById(focusId)?.focus(), 50);
  }
  function closeModal(id) {
    document.getElementById(id).classList.remove("open");
  }

  function showToast(msg) {
    if (typeof window.showToast === "function") { window.showToast(msg); return; }
  }

  return {
    init,
    isLoggedIn: () => !!sessionToken,
    getSessionToken: () => sessionToken,
    getNickname: () => nickname,
    getUserId: () => userId,
    requireLogin,
    openAuthModal,
    onLoginChange: (cb) => onLoginChangeCallbacks.push(cb),
  };
})();

document.addEventListener("DOMContentLoaded", () => Auth.init());
