/* ===================================================================
   퀵오더 — Gmail 연동 (브라우저에서 직접, 서버 없음)
   구글 로그인(OAuth) → 메일 읽기 / 첨부 받기 / 메일 보내기
   =================================================================== */
"use strict";
const GMAIL = (() => {
  const SCOPES = "https://www.googleapis.com/auth/gmail.readonly https://www.googleapis.com/auth/gmail.send " +
                 "https://www.googleapis.com/auth/drive.appdata https://www.googleapis.com/auth/drive";  // drive = 읽기+쓰기(되쓰기용)
  const API = "https://gmail.googleapis.com/gmail/v1/users/me";
  /* 공용 메일 도메인 — 업체 회사 도메인인지 가릴 때 쓴다 */
  const FREEMAIL = /^(gmail|googlemail|naver|daum|hanmail|nate|kakao|hanmir|empas|korea|outlook|hotmail|live|msn|yahoo|icloud|me|proton|protonmail|aol|qq|163)\./i;
  // 회사별로 키를 분리 (같은 도메인에 여러 회사를 올려도 로그인이 섞이지 않게)
  const TKEY = () => CONFIG.ls("qo_gmail_token4");    // 토큰을 기기에 보관 → 로그인 유지
  const GKEY = () => CONFIG.ls("qo_gmail_granted4");  // 권한 승인 이력(토큰 만료 후 동의창 반복 방지)
  const HKEY = () => CONFIG.ls("qo_gmail_hint");      // 계정 이메일 힌트(재로그인 시 계정 선택 건너뛰기)
  let tokenClient = null, accessToken = null, tokenExp = 0, clientId = null;

  /* 저장해 둔 토큰 불러오기 (아직 유효하면 재로그인 불필요)
     ★ 이 파일이 로드되는 시점엔 아직 누가 로그인했는지 모른다. 그래서 앞사람의 토큰이
       메모리에 올라와, 다른 업체로 들어가도 '구글 메일 연결됨' 으로 보였다.
       업체가 정해진 뒤 reloadToken() 을 다시 불러 그 업체 것으로 맞춘다. */
  function reloadToken() {
    accessToken = null; tokenExp = 0;
    try {
      const s = JSON.parse(localStorage.getItem(TKEY()) || "null");
      if (s && s.token && s.exp && Date.now() < s.exp) { accessToken = s.token; tokenExp = s.exp; }
    } catch (e) {}
  }
  reloadToken();
  function saveToken() {
    try {
      localStorage.setItem(TKEY(), JSON.stringify({ token: accessToken, exp: tokenExp }));
      localStorage.setItem(GKEY(), "1");   // 승인 이력 기록(토큰 만료돼도 유지)
    } catch (e) {}
  }
  // 이 기기에서 한 번이라도 구글 권한을 승인했는지 (동의창 반복 방지용)
  function granted() { try { return localStorage.getItem(GKEY()) === "1"; } catch (e) { return false; } }
  function clearToken() { try { localStorage.removeItem(TKEY()); localStorage.removeItem(GKEY()); } catch (e) {} }
  function withTimeout(p, ms) {
    return new Promise((res, rej) => {
      const t = setTimeout(() => rej(new Error("timeout")), ms);
      p.then(v => { clearTimeout(t); res(v); }, e => { clearTimeout(t); rej(e); });
    });
  }

  /* ---------- 인증 ---------- */
  function gsiLoaded() {
    return !!(window.google && google.accounts && google.accounts.oauth2 &&
              google.accounts.oauth2.initTokenClient);
  }
  function ready() { return !!clientId && !!tokenClient; }
  /* 구글이 돌려주는 영문 오류를 무슨 뜻인지 알아볼 수 있게 바꾼다.
     access_denied 는 대개 '비밀번호가 틀렸다' 가 아니라, 그 구글 프로젝트가 아직
     테스트 상태라 등록된 사람만 들어올 수 있다는 뜻이다. */
  /* 클라이언트 ID 앞의 숫자가 곧 구글 프로젝트 번호다.
     '게시했는데 왜 막히지' 는 대개 다른 프로젝트를 게시한 경우라, 번호를 눈에 보이게 둔다. */
  function projectNo(cid) { return String(cid || clientId || "").split("-")[0] || "?"; }
  function explain(msg) {
    const m = String(msg || "");
    if (/access_denied|has not completed|verification/i.test(m))
      return "이 구글 프로젝트가 아직 '테스트' 상태라 등록된 계정만 들어올 수 있습니다.\n\n" +
             "▶ 지금 쓰고 있는 구글 프로젝트 번호: " + projectNo() + "\n" +
             "   관리자에게 이 번호를 알려주고 '이 프로젝트'를 게시해 달라고 하세요.\n" +
             "   콘솔 → API 및 서비스 → OAuth 동의 화면(대상) → [앱 게시] → 프로덕션\n" +
             "   ※ 게시했는데도 이 창이 뜨면 다른 프로젝트를 게시한 것입니다.\n\n" +
             "다른 방법 — 회사 전용 구글 프로젝트를 쓰려면\n" +
             "  설정 → 구글 메일·드라이브 연결 → 클라이언트 ID 에 그 프로젝트 ID 를 넣고 [저장]\n\n" +
             "(원문: " + m + ")";
    if (/popup_closed|취소/i.test(m)) return "로그인 창이 닫혔습니다. 다시 시도하세요.";
    return m;
  }
  function signedIn() { return !!accessToken && Date.now() < tokenExp; }
  /* ★ '오늘' 은 오늘 날짜여야 한다.
     newer_than:1d 는 '최근 24시간' 이라, 지금이 15:00 이면 어제 15:00~24:00 메일이 함께 걸린다.
     실제로 어제 15:41 메일이 '오늘' 에 떴다. 하루짜리는 오늘 0시 이후로 못박는다.
     (3일·7일은 이름 그대로 '최근 N일' 이므로 그대로 둔다) */
  function sinceQuery(days) {
    const d = Number(days) || 0;
    if (d === 1) {
      const t = new Date(); t.setHours(0, 0, 0, 0);
      return `after:${Math.floor(t.getTime() / 1000)}`;
    }
    return `newer_than:${d}d`;
  }

  /* 클라이언트 ID 저장 + (가능하면) 토큰 클라이언트 생성
     ★ ID 가 바뀌면 tokenClient 를 버려야 한다. ensureInit 은 없을 때만 만들기 때문에,
       버리지 않으면 새 ID 를 저장해도 계속 옛 프로젝트로 로그인 창이 뜬다. */
  function init(cid) {
    if (cid !== undefined) {
      const v = (cid || "").trim();
      if (v !== clientId) { clientId = v; tokenClient = null; }
    }
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

  // mode: "" = 조용한 갱신(화면 안 뜸) / "select_account" = 계정 선택 / "consent" = 최초 동의
  function signIn(mode) {
    return new Promise(async (res, rej) => {
      if (!clientId) return rej(new Error("클라이언트 ID가 없습니다. 설정에서 입력·저장하세요."));
      if (!ensureInit()) {
        const ok = await waitReady();
        if (!ok) return rej(new Error("구글 로그인 라이브러리를 불러오지 못했습니다.\n인터넷 연결 또는 광고차단/보안 확장프로그램을 확인하세요."));
      }
      if (!tokenClient) return rej(new Error("구글 로그인 준비가 안 됐어요. 설정에서 클라이언트 ID를 확인하세요."));
      tokenClient.callback = resp => {
        if (resp.error) return rej(new Error(explain(resp.error_description || resp.error)));
        accessToken = resp.access_token;
        tokenExp = Date.now() + (Number(resp.expires_in || 3600) - 60) * 1000;
        saveToken();                       // 기기에 저장 + 승인이력 기록
        res(accessToken);
      };
      try { tokenClient.error_callback = e => rej(new Error(explain((e && (e.message || e.type)) || "로그인 취소"))); } catch (e) {}
      // hint(계정 이메일)를 주면 계정 선택 단계를 건너뛰고, 세션이 살아있으면 화면 없이 조용히 통과
      const hint = (function () { try { return localStorage.getItem(HKEY()) || ""; } catch (e) { return ""; } })();
      const cfg = { prompt: mode === undefined ? "" : mode };
      if (hint) cfg.hint = hint;
      try { tokenClient.requestAccessToken(cfg); }
      catch (e) { rej(e); }
    });
  }
  /* 토큰 얻기: 유효하면 그대로, 만료면 '조용한 갱신' 먼저 → 안 되면 계정선택
     (예전엔 만료될 때마다 consent(동의창)를 띄워서 매번 로그인하라고 나왔음) */
  /* 토큰 얻기. ※ 구글 GSI는 prompt:""라도 '팝업'을 열기 때문에, 반드시 사용자의
     '직접 클릭' 안에서 호출해야 한다. (자동 호출하면 브라우저가 팝업을 차단함) */
  async function token(allowInteractive = true) {
    if (signedIn()) return accessToken;
    if (!allowInteractive) throw new Error("NEED_LOGIN");
    // 이미 승인한 적 있으면 prompt:"" → 동의창 없이 팝업이 잠깐 떴다 닫힘
    return await signIn(granted() ? "" : "consent");
  }
  const needLogin = () => !signedIn();
  function signOut() {
    if (accessToken && window.google) try { google.accounts.oauth2.revoke(accessToken, () => {}); } catch (e) {}
    accessToken = null; tokenExp = 0; clearToken();
  }
  function hasToken() { return !!accessToken; }
  /* 계정 바꾸기 — 힌트가 남아 있으면 구글이 계정 선택 단계를 건너뛰어 버려서,
     '다른 계정으로 바꾸고 싶은데 늘 같은 계정으로 들어가지는' 일이 생긴다. 힌트부터 지운다. */
  async function switchAccount() {
    try { localStorage.removeItem(HKEY()); } catch (e) {}
    clearToken();                       // 떠나는 계정의 토큰을 그 자리에서 지운다
    accessToken = null; tokenExp = 0;
    return await signIn("select_account");
  }
  /* 저장 자리(회사·계정)가 바뀐 뒤 지금 토큰을 그 자리에 다시 써넣는다.
     ★ 로그인 시점엔 아직 '앞 계정' 자리라, 여기서 옮기지 않으면 새 계정 토큰이
       앞 계정 칸에 남는다 — 다음에 열 때 남의 드라이브가 보이는 원인이었다. */
  function persistToken() { if (accessToken) saveToken(); }
  function dropStored() { clearToken(); }   // 메모리는 그대로, 저장된 것만 지운다

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
    const q = `${sinceQuery(days)} has:attachment (filename:xlsx OR filename:xlsm)`;
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
                      ts: Number(msg.internalDate) || 0,
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

  /* ---------- 본문 메일 목록 (첨부 없어도 됨) — CS 문의용 ----------
     listMails 는 '엑셀 첨부가 있는 메일'만 본다. CS는 첨부 없이 본문만 오는
     고객/고객센터 문의가 대부분이라 별도로 검색한다. 첨부가 있으면 함께 준다. */
  async function listTextMails({ days = 7, keywords = [], senders = [], exclude = [], max = 30, onProgress } = {}) {
    const qp = [sinceQuery(days), "-in:chats"];
    const or = [];
    if (keywords.length) or.push("(" + keywords.map(k => `"${String(k).replace(/"/g, "")}"`).join(" OR ") + ")");
    if (senders.length) or.push("(" + senders.map(s => `from:${String(s).trim()}`).join(" OR ") + ")");
    if (or.length) qp.push(or.length > 1 ? "(" + or.join(" OR ") + ")" : or[0]);
    const listed = await api(`/messages?q=${encodeURIComponent(qp.join(" "))}&maxResults=${max}`);
    const ids = (listed.messages || []).map(m => m.id);
    const out = [];
    for (let i = 0; i < ids.length; i++) {
      if (onProgress) onProgress(i + 1, ids.length);
      let msg;
      try { msg = await api(`/messages/${ids[i]}?format=full`); } catch (e) { continue; }
      const subject = headerOf(msg, "Subject");
      const from = headerOf(msg, "From");
      const body = bodyText(msg) || "";
      if (exclude.length && exclude.some(x => (subject + " " + body).includes(x))) continue;
      const atts = walkParts(msg.payload, [])
        .filter(p => p.filename && /\.xls[xm]$/i.test(p.filename) && p.body && p.body.attachmentId)
        .map(p => ({ filename: p.filename, attachmentId: p.body.attachmentId }));
      out.push({ id: msg.id, subject, from, date: fmtDate(msg.internalDate),
                 ts: Number(msg.internalDate) || 0, body: body.slice(0, 1500), atts });
    }
    return out;
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

  async function profile() {
    const p = await api("/profile");
    // 로그인한 계정 이메일을 힌트로 저장 → 다음 재로그인 때 계정 선택 안 뜨게
    try { if (p && p.emailAddress) localStorage.setItem(HKEY(), p.emailAddress); } catch (e) {}
    return p;
  }

  /* ---------- 업체명으로 메일 검색 → 받는사람 후보 주소 뽑기 ---------- */
  function extractEmails(str) {
    if (!str) return [];
    const m = String(str).match(/[^\s<>,"();:]+@[^\s<>,"();:]+\.[^\s<>,"();:]+/g);
    return m ? m.map(s => s.replace(/[.,;]+$/, "")) : [];
  }
  async function searchAddresses({ query, max = 15, days = 365 } = {}) {
    if (!query) return [];
    const q = `${query} ${sinceQuery(days)}`;
    let listed;
    try { listed = await api(`/messages?q=${encodeURIComponent(q)}&maxResults=${max}`); }
    catch (e) { return []; }
    const ids = (listed.messages || []).map(m => m.id);
    let self = "";
    try { self = ((await profile()).emailAddress || "").toLowerCase(); } catch (e) {}
    // 내 회사 도메인은 받는사람 후보가 아니다 (우리끼리 주고받은 주소가 섞여 올라온다).
    // 다만 gmail 같은 공용 메일이면 도메인으로 거를 수 없으니 그대로 둔다.
    const selfDom = (self.split("@")[1] || "").toLowerCase();
    const ownDom = FREEMAIL.test(selfDom + ".") ? "" : selfDom;
    const tally = {};
    await Promise.all(ids.map(async id => {
      let msg;
      try {
        msg = await api(`/messages/${id}?format=metadata&metadataHeaders=From&metadataHeaders=To&metadataHeaders=Cc&metadataHeaders=Reply-To`);
      } catch (e) { return; }
      for (const h of ["From", "To", "Cc", "Reply-To"]) {
        for (const addr of extractEmails(headerOf(msg, h))) {
          const lc = addr.toLowerCase();
          if (!lc || lc === self) continue;
          if (ownDom && (lc.split("@")[1] || "") === ownDom) continue;
          if (/(no[-_.]?reply|mailer-daemon|postmaster|donotreply|notification)/i.test(lc)) continue;
          const t = tally[lc] = tally[lc] || { count: 0, from: 0 };
          t.count++;
          if (h === "From" || h === "Reply-To") t.from++;   // 보낸 사람이면 더 확실한 담당자
        }
      }
    }));
    return Object.keys(tally)
      .map(email => ({ email, count: tally[email].count, fromCount: tally[email].from }))
      .sort((a, b) => b.count - a.count);
  }

  /* ---------- 구글 드라이브 (앱 전용 숨김 폴더 appDataFolder) ---------- */
  function driveErr(status, body) {
    let m = "";
    try { m = (JSON.parse(body).error || {}).message || ""; } catch (e) {}
    if (status === 403 && /has not been used|is disabled|Drive API/i.test(body || ""))
      return new Error("구글 드라이브 API가 아직 켜져 있지 않습니다.\n구글 클라우드 콘솔에서 'Google Drive API'를 사용 설정하세요.");
    if (status === 401 || status === 403)
      return new Error("드라이브 접근 권한이 없습니다. 설정에서 구글을 다시 연결(재로그인)하고 권한을 허용하세요.");
    return new Error("드라이브 오류: " + (m || ("HTTP " + status)));
  }
  async function driveFind(name) {
    const t = await token();
    const q = encodeURIComponent(`name='${name}'`);
    const r = await fetch(`https://www.googleapis.com/drive/v3/files?spaces=appDataFolder&q=${q}&fields=files(id,modifiedTime)&orderBy=modifiedTime desc`,
      { headers: { Authorization: "Bearer " + t } });
    if (!r.ok) throw driveErr(r.status, await r.text());
    const d = await r.json();
    return (d.files && d.files[0]) ? d.files[0].id : null;
  }
  async function driveDownload(fileId) {
    const t = await token();
    const r = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`,
      { headers: { Authorization: "Bearer " + t } });
    if (!r.ok) throw driveErr(r.status, await r.text());
    return await r.text();
  }
  /* ---------- 드라이브: 내 파일 읽기(발주서 끌어오기) ---------- */
  const DRIVE = "https://www.googleapis.com/drive/v3";
  const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const GSHEET_MIME = "application/vnd.google-apps.spreadsheet";
  async function driveApiGet(path) {
    const t = await token();
    const r = await fetch(DRIVE + path, { headers: { Authorization: "Bearer " + t } });
    if (!r.ok) throw driveErr(r.status, await r.text());
    return await r.json();
  }
  // 공유링크/ID 문자열에서 파일 ID 뽑기
  function driveIdFromLink(s) {
    s = String(s || "").trim();
    const m = s.match(/\/d\/([a-zA-Z0-9_-]{10,})/) || s.match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
    if (m) return m[1];
    if (/^[a-zA-Z0-9_-]{10,}$/.test(s)) return s;
    return null;
  }
  async function driveFileInfo(fileId) {
    return await driveApiGet(`/files/${encodeURIComponent(fileId)}?fields=id,name,mimeType,modifiedTime,parents&supportsAllDrives=true`);
  }
  /* 실제 상위 폴더들을 따라 올라가 경로를 만든다 (파일탐색기처럼 상위 이동하려고) */
  async function driveAncestors(folderId, max = 8) {
    const chain = [];
    let id = folderId;
    for (let i = 0; i < max && id; i++) {
      let info;
      try { info = await driveFileInfo(id); } catch (e) { break; }   // 권한 없는 상위면 중단
      if (!info || !info.id) break;
      chain.unshift({ id: info.id, name: info.name });
      id = (info.parents && info.parents[0]) || null;
    }
    return chain;
  }
  // 엑셀/시트 파일 검색 (이름으로, 최근 수정순)
  async function driveSearch(q, max = 30) {
    const types = `(mimeType='${XLSX_MIME}' or mimeType='${GSHEET_MIME}' or mimeType='application/vnd.ms-excel')`;
    let qq = `${types} and trashed=false`;
    if (q && q.trim()) qq += ` and name contains '${q.trim().replace(/'/g, "\\'")}'`;
    const d = await driveApiGet(`/files?q=${encodeURIComponent(qq)}` +
      `&fields=files(id,name,mimeType,modifiedTime)&orderBy=modifiedTime desc&pageSize=${max}` +
      `&supportsAllDrives=true&includeItemsFromAllDrives=true`);
    return d.files || [];
  }
  // 폴더 안의 하위폴더 + 엑셀/시트 파일 목록 (폴더 먼저, 이름순)
  const FOLDER_MIME = "application/vnd.google-apps.folder";
  async function driveListFolder(folderId, max = 200) {
    const types = `(mimeType='${FOLDER_MIME}' or mimeType='${XLSX_MIME}' or mimeType='${GSHEET_MIME}' or mimeType='application/vnd.ms-excel')`;
    const qq = `'${folderId || "root"}' in parents and trashed=false and ${types}`;
    const d = await driveApiGet(`/files?q=${encodeURIComponent(qq)}` +
      `&fields=files(id,name,mimeType,modifiedTime)&orderBy=folder,name&pageSize=${max}` +
      `&supportsAllDrives=true&includeItemsFromAllDrives=true`);
    return d.files || [];
  }
  const isFolder = f => f && f.mimeType === FOLDER_MIME;

  // 남이 공유해준 항목(공유 문서함) — '내 드라이브' 목록에는 안 나오므로 따로 조회
  async function driveListShared(max = 200) {
    const types = `(mimeType='${FOLDER_MIME}' or mimeType='${XLSX_MIME}' or mimeType='${GSHEET_MIME}' or mimeType='application/vnd.ms-excel')`;
    const qq = `sharedWithMe=true and trashed=false and ${types}`;
    const d = await driveApiGet(`/files?q=${encodeURIComponent(qq)}` +
      `&fields=files(id,name,mimeType,modifiedTime)&orderBy=folder,name&pageSize=${max}` +
      `&supportsAllDrives=true&includeItemsFromAllDrives=true`);
    return d.files || [];
  }

  // 파일을 엑셀 바이너리로 받기 (구글시트면 xlsx로 변환해서)
  async function driveFetchExcel(fileId) {
    const info = await driveFileInfo(fileId);
    const t = await token();
    const url = info.mimeType === GSHEET_MIME
      ? `${DRIVE}/files/${encodeURIComponent(fileId)}/export?mimeType=${encodeURIComponent(XLSX_MIME)}`
      : `${DRIVE}/files/${encodeURIComponent(fileId)}?alt=media&supportsAllDrives=true`;
    const r = await fetch(url, { headers: { Authorization: "Bearer " + t } });
    if (!r.ok) throw driveErr(r.status, await r.text());
    const buf = await r.arrayBuffer();
    let name = info.name || "drive";
    if (!/\.xls[xm]$/i.test(name)) name += ".xlsx";
    return { buf, name, modifiedTime: info.modifiedTime };
  }

  /* 드라이브의 기존 파일에 '그대로 덮어쓰기'(내용 갱신) — 송장취합 결과를 원본 양식에 되쓰기 */
  async function driveUpdateFile(fileId, arrayBuffer) {
    const t = await token();
    const r = await fetch(`https://www.googleapis.com/upload/drive/v3/files/${encodeURIComponent(fileId)}?uploadType=media&supportsAllDrives=true&fields=id,name,modifiedTime`, {
      method: "PATCH",
      headers: { Authorization: "Bearer " + t, "Content-Type": XLSX_MIME },
      body: arrayBuffer,
    });
    if (!r.ok) throw driveErr(r.status, await r.text());
    return await r.json();
  }

  async function driveUpload(name, content, fileId) {
    const t = await token();
    if (fileId) {
      const r = await fetch(`https://www.googleapis.com/upload/drive/v3/files/${fileId}?uploadType=media&fields=id`,
        { method: "PATCH", headers: { Authorization: "Bearer " + t, "Content-Type": "application/json" }, body: content });
      if (r.status === 404) return driveUpload(name, content, null);   // 원본이 지워졌으면 새로 생성
      if (!r.ok) throw driveErr(r.status, await r.text());
      return (await r.json()).id;
    }
    const boundary = "qoBd" + Math.random().toString(36).slice(2);
    const meta = JSON.stringify({ name, parents: ["appDataFolder"] });
    const body = `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${meta}\r\n` +
                 `--${boundary}\r\nContent-Type: application/json\r\n\r\n${content}\r\n--${boundary}--`;
    const r = await fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id",
      { method: "POST", headers: { Authorization: "Bearer " + t, "Content-Type": "multipart/related; boundary=" + boundary }, body });
    if (!r.ok) throw driveErr(r.status, await r.text());
    return (await r.json()).id;
  }

  return { init, ensureInit, waitReady, gsiLoaded, ready, signedIn, hasToken, token, signIn, signOut, switchAccount, persistToken, dropStored, reloadToken, projectNo, _sinceQuery: sinceQuery, listMails, listTextMails, getAttachment, send, profile,
           searchAddresses, driveFind, driveDownload, driveUpload, granted,
           driveIdFromLink, driveFileInfo, driveSearch, driveFetchExcel, driveListFolder, driveListShared, driveAncestors, driveUpdateFile, needLogin };
})();
