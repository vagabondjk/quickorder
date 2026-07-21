/* ===================================================================
   퀵오더 — Gmail 연동 (브라우저에서 직접, 서버 없음)
   구글 로그인(OAuth) → 메일 읽기 / 첨부 받기 / 메일 보내기
   =================================================================== */
"use strict";
const GMAIL = (() => {
  const SCOPES = "https://www.googleapis.com/auth/gmail.readonly https://www.googleapis.com/auth/gmail.send";
  const API = "https://gmail.googleapis.com/gmail/v1/users/me";
  const TKEY = "qo_gmail_token";     // 토큰을 기기에 보관 → 로그인 유지
  let tokenClient = null, accessToken = null, tokenExp = 0, clientId = null;

  // 저장해 둔 토큰 불러오기 (아직 유효하면 재로그인 불필요)
  (function loadToken() {
    try {
      const s = JSON.parse(localStorage.getItem(TKEY) || "null");
      if (s && s.token && s.exp && Date.now() < s.exp) { accessToken = s.token; tokenExp = s.exp; }
    } catch (e) {}
  })();
  function saveToken() {
    try { localStorage.setItem(TKEY, JSON.stringify({ token: accessToken, exp: tokenExp })); } catch (e) {}
  }
  function clearToken() { try { localStorage.removeItem(TKEY); } catch (e) {} }

  /* ---------- 인증 ---------- */
  function gsiLoaded() {
    return !!(window.google && google.accounts && google.accounts.oauth2 &&
              google.accounts.oauth2.initTokenClient);
  }
  function ready() { return !!clientId && !!tokenClient; }
  function signedIn() { return !!accessToken && Date.now() < tokenExp; }

  // 클라이언트 ID 저장 + (가능하면) 토큰 클라이언트 생성
  function init(cid) {
    if (cid !== undefined) clientId = (cid || "").trim();
    return ensureInit();
  }
  function ensureInit() {
    if (!clientId) return false;
    if (!gsiLoaded()) return false;
    try {
      if (!tokenClient) {
        tokenClient = google.accounts.oauth2.initTokenClient({
          client_id: clientId, scope: SCOPES, callback: () => {},
        });
      }
      return true;
    } catch (e) { return false; }
  }
  // 구글 라이브러리(GSI)가 늦게 로드돼도 최대 ms 밀리초까지 기다렸다 준비
  async function waitReady(ms = 9000) {
    const t0 = Date.now();
    while (Date.now() - t0 < ms) {
      if (ensureInit()) return true;
      await new Promise(r => setTimeout(r, 150));
    }
    return ensureInit();
  }

  function signIn(interactive = true) {
    return new Promise(async (res, rej) => {
      if (!clientId) return rej(new Error("클라이언트 ID가 없습니다. 설정에서 입력·저장하세요."));
      if (!ensureInit()) {
        const ok = await waitReady();
        if (!ok) return rej(new Error("구글 로그인 라이브러리를 불러오지 못했습니다.\n인터넷 연결 또는 광고차단/보안 확장프로그램을 확인하세요."));
      }
      if (!tokenClient) return rej(new Error("구글 로그인 준비가 안 됐어요. 설정에서 클라이언트 ID를 확인하세요."));
      tokenClient.callback = resp => {
        if (resp.error) return rej(new Error(resp.error_description || resp.error));
        accessToken = resp.access_token;
        tokenExp = Date.now() + (Number(resp.expires_in || 3600) - 60) * 1000;
        saveToken();                       // 기기에 저장 → 다음에 재로그인 안 해도 됨
        res(accessToken);
      };
      // 처음 한 번만 동의창(consent), 이후에는 조용히(prompt:"") 갱신
      try { tokenClient.requestAccessToken({ prompt: interactive ? "consent" : "" }); }
      catch (e) { rej(e); }
    });
  }
  async function token() {
    if (signedIn()) return accessToken;
    return await signIn(!accessToken);   // 처음엔 동의창, 이후엔 조용히 갱신
  }
  function signOut() {
    if (accessToken && window.google) try { google.accounts.oauth2.revoke(accessToken, () => {}); } catch (e) {}
    accessToken = null; tokenExp = 0; clearToken();
  }
  function hasToken() { return !!accessToken; }

  async function api(path, opts) {
    const t = await token();
    const r = await fetch(API + path, {
      ...(opts || {}),
      headers: { Authorization: "Bearer " + t, ...((opts || {}).headers || {}) },
    });
    if (!r.ok) {
      let d = {}; try { d = await r.json(); } catch (e) {}
      const m = (d.error && (d.error.message || d.error.status)) || ("HTTP " + r.status);
      if (r.status === 401) { accessToken = null; }
      throw new Error(m);
    }
    return await r.json();
  }

  /* ---------- 유틸 ---------- */
  const b64urlToBytes = s => {
    const b = atob(s.replace(/-/g, "+").replace(/_/g, "/"));
    const u = new Uint8Array(b.length);
    for (let i = 0; i < b.length; i++) u[i] = b.charCodeAt(i);
    return u;
  };
  const bytesToB64 = buf => {
    const u = new Uint8Array(buf); let s = "";
    for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000));
    return btoa(s);
  };
  const utf8B64 = str => btoa(String.fromCharCode(...new TextEncoder().encode(str)));
  const mimeWord = s => "=?UTF-8?B?" + utf8B64(s) + "?=";
  function decodeText(part) {
    if (!part || !part.body || !part.body.data) return "";
    try { return new TextDecoder("utf-8").decode(b64urlToBytes(part.body.data)); } catch (e) { return ""; }
  }
  function headerOf(msg, name) {
    const hs = (msg.payload && msg.payload.headers) || [];
    const h = hs.find(x => x.name.toLowerCase() === name.toLowerCase());
    return h ? h.value : "";
  }
  function walkParts(part, out) {
    if (!part) return out;
    out.push(part);
    (part.parts || []).forEach(p => walkParts(p, out));
    return out;
  }
  function bodyText(msg) {
    const parts = walkParts(msg.payload, []);
    const plain = parts.filter(p => p.mimeType === "text/plain" && !(p.filename || "")).map(decodeText).join("\n").trim();
    if (plain) return plain;
    const html = parts.filter(p => p.mimeType === "text/html" && !(p.filename || "")).map(decodeText).join("\n");
    return html
      .replace(/<(script|style)[^>]*>[\s\S]*?<\/\1>/gi, " ")
      .replace(/<br\s*\/?>/gi, "\n").replace(/<\/(p|div|tr|li|h[1-6])\s*>/gi, "\n")
      .replace(/<[^>]+>/g, "")
      .replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"')
      .replace(/[ \t]+/g, " ").replace(/\n{3,}/g, "\n\n").trim();
  }
  function fmtDate(internalDate) {
    const d = new Date(Number(internalDate));
    const p = n => String(n).padStart(2, "0");
    return `${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
  }

  /* ---------- 메일 목록 (엑셀 첨부만) ----------
     senders  : 발신 메일/도메인 (일치하면 후보)
     keywords : 첨부명·제목·본문에 이 단어가 있으면 후보
     union    : true면 발신자+키워드 합집합(회신용), false면 발신자 우선(발주서용)
  ------------------------------------------------------------------ */
  async function listMails({ days = 7, senders = [], keywords = [], exclude = [], union = false, scanText = false, max = 40, onProgress } = {}) {
    const q = `newer_than:${days}d has:attachment (filename:xlsx OR filename:xlsm)`;
    const listed = await api(`/messages?q=${encodeURIComponent(q)}&maxResults=${max}`);
    const ids = (listed.messages || []).map(m => m.id);
    const fromHits = [], nameHits = [];
    for (let i = 0; i < ids.length; i++) {
      if (onProgress) onProgress(i + 1, ids.length);
      let msg;
      try { msg = await api(`/messages/${ids[i]}?format=full`); } catch (e) { continue; }
      const subject = headerOf(msg, "Subject");
      const from = headerOf(msg, "From");
      const date = fmtDate(msg.internalDate);
      const fromLc = from.toLowerCase();
      const fromOk = senders.length ? senders.some(s => fromLc.includes(String(s).trim().toLowerCase())) : false;
      let body = null;
      const parts = walkParts(msg.payload, []);
      for (const p of parts) {
        const fn = p.filename || "";
        if (!fn || !/\.xls[xm]$/i.test(fn)) continue;
        if (exclude.length && exclude.some(x => fn.includes(x))) continue;
        const att = p.body && p.body.attachmentId;
        if (!att) continue;
        const rec = { id: msg.id, attachmentId: att, filename: fn, subject, from, date,
                      body: (body === null ? (body = bodyText(msg)) : body).slice(0, 500) };
        if (fromOk) { fromHits.push(rec); continue; }
        let nameOk = keywords.length ? keywords.some(k => fn.includes(k)) : false;
        if (!nameOk && keywords.length && scanText) {
          if (body === null) body = bodyText(msg);
          const hay = subject + " " + body;
          nameOk = keywords.some(k => hay.includes(k));
        }
        if (nameOk) nameHits.push(rec);
      }
    }
    const items = union ? fromHits.concat(nameHits) : (fromHits.length ? fromHits : nameHits);
    return items;
  }

  /* ---------- 첨부 내려받기 ---------- */
  async function getAttachment(messageId, attachmentId) {
    const d = await api(`/messages/${messageId}/attachments/${attachmentId}`);
    return b64urlToBytes(d.data).buffer;
  }

  /* ---------- 메일 보내기 (첨부 포함) ---------- */
  async function send({ to, subject, body, attachments = [] }) {
    const bd = "----qo" + Date.now().toString(36);
    const L = [];
    L.push(`To: ${to}`);
    L.push(`Subject: ${mimeWord(subject)}`);
    L.push("MIME-Version: 1.0");
    L.push(`Content-Type: multipart/mixed; boundary="${bd}"`);
    L.push("");
    L.push(`--${bd}`);
    L.push('Content-Type: text/plain; charset="UTF-8"');
    L.push("Content-Transfer-Encoding: base64");
    L.push("");
    L.push(utf8B64(body));
    for (const a of attachments) {
      L.push(`--${bd}`);
      L.push(`Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; name="${mimeWord(a.filename)}"`);
      L.push("Content-Transfer-Encoding: base64");
      L.push(`Content-Disposition: attachment; filename="${mimeWord(a.filename)}"`);
      L.push("");
      L.push(bytesToB64(a.data));
    }
    L.push(`--${bd}--`);
    const raw = btoa(unescape(encodeURIComponent(L.join("\r\n"))))
      .replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
    return await api("/messages/send", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ raw }),
    });
  }

  async function profile() { return await api("/profile"); }

  return { init, ensureInit, waitReady, gsiLoaded, ready, signedIn, hasToken, token, signIn, signOut, listMails, getAttachment, send, profile };
})();
