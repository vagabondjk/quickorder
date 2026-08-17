/* ===================================================================
   퀵오더 — 사용 승인(월별 비밀번호) 잠금  [웹 전용 비밀번호]
   "[JK전용] 퀵오더웹 비밀번호 2년치.txt" 의 비밀번호로 로그인.
   * 소스에는 비밀번호의 해시만 있어, 코드를 봐도 비밀번호를 알 수 없음.
   * 관리자 마스터 비밀번호는 만료 없이 통과(저장하지 않음).
   * 월 비밀번호로 들어오면 그 달 동안 이 기기에서 재입력 불필요.
   =================================================================== */
"use strict";
const LOCK = (() => {
  /* 회사별 비밀번호 — qo-config.js 의 lock 값에 따라 갈린다.
       null  → 아래 기본(랩노마드) 비밀번호 사용
       false → 잠금 없음 (바로 사용)
       {salt, master, months} → 그 회사 전용 비밀번호 */
  const LK = (typeof CONFIG !== "undefined") ? CONFIG.lock : null;
  const OFF = LK === false;
  const SALT = (LK && LK.salt) || "047ec6b655a0775637a17f233ce04bd4";
  const MASTER = OFF ? "" : ((LK && LK.master) ||
    "4e3782b5293b1e04564a07aa343a25638ff682519529454c5898d0d984907c35");
  const MONTHS = OFF ? {} : ((LK && LK.months) || {
    "202607": "e75224ee2d3606ca4cd87a87aaef042d92614ba1995e973c9d6bcbe1b7197f68",
    "202608": "869abde9ba917da3ac189c77dba267fccf9a6918a84c164ad052ea2f3a85776d",
    "202609": "cbe76ea6b6c40a2111332fb483004fc21a319f1fed4833af3d23ae8dc705e33f",
    "202610": "0feac02788359c0b4c2a5b11abe13f07fb7e4475d1d319e79f4fa39062fa065b",
    "202611": "ac62c71028f6e5d58d566562f78bf0c6de4b41538e7832bbca40769168beed1f",
    "202612": "7c066a3f172c383fd0814476d6dd3b206634ea1855bc938ffb3cd91c517e0679",
    "202701": "fb7d13374438440127f935c06cd72766de3aec4d6ea77d298fa9d3e0e0dd9658",
    "202702": "dc07b35af0802eb8ff345d7ef0a605762fde63029b4ce6a4d5d0b27c2f2cee40",
    "202703": "8e8e97e1354e1d085fbe1c6805931015a43c9fe749fbc4b6a2bc4ceb77a2b335",
    "202704": "552051b106e7da18918de7d3151c78880a433cdd87e1c824373bcded57a780da",
    "202705": "f24f6d2164a74e4907f75e7d5cc67e4abc06efa292be76ed09e8414d831af57c",
    "202706": "ad8fa5e4e8dd36bb05f38c27aa843bb3080daca3d8d1f8118aa801eb49361659",
    "202707": "e3220a841ffd431c055763755e71d256693e906d0b9a4c81620a247602b150b7",
    "202708": "3dcfb41ac36cc5a6c16fc1985064d2a58fa626deaa345d0dd08a562ef5f45b0b",
    "202709": "8c89c8350a7b8968d75b98877cbde08f10b387cee2eefc7bfa17a6dd29cf6304",
    "202710": "f98d4126d7e19a503f9989f5b235c1b895b6bc6184429b970b94c235923d228a",
    "202711": "d1ebedc0f42593036edefb3848377ef92f02421533e7fc93d4cc3c910be320d5",
    "202712": "eb8d2c7de5caa1023caf004613a5c23fc4787bc433a1fb0f060ee91bcbb4b2a9",
    "202801": "b631edb75d612bf7b51a23198d91d143264cc91427328b32fd10c48ac4279b22",
    "202802": "4aac6206280bee4a6163d82a6fe24fb25e804d56d7c8d35ec64fb8706ca8e351",
    "202803": "0ba5b6fb1c0722396a3bcb8bfd54d0354b9792d1ff4e42496a32d302d3ed7c1d",
    "202804": "5eca2575eaa9808db946ed3fb06a014b02844e133af6e4be8540571b9da9661e",
    "202805": "1c52a9a2380fca585e064ad127e212327655e31bb872fb7f9504860aa3f4bbef",
    "202806": "a186e069936eb801004e2976d21003b307d9804aff5d0a242a7b0f1693949385",
  });

  async function sha256Hex(str) {
    const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(str));
    return [...new Uint8Array(buf)].map(b => b.toString(16).padStart(2, "0")).join("");
  }
  const hashPw = pw => sha256Hex(SALT + String(pw).trim().toUpperCase());
  function currentYm() {
    const d = new Date();
    return "" + d.getFullYear() + String(d.getMonth() + 1).padStart(2, "0");
  }
  function deviceId() {
    let id = null;
    try { id = localStorage.getItem(CONFIG.lsBase("qo_device")); } catch (e) {}
    if (!id) {
      const r = crypto.getRandomValues(new Uint8Array(16));
      id = [...r].map(b => b.toString(16).padStart(2, "0")).join("");
      try { localStorage.setItem(CONFIG.lsBase("qo_device"), id); } catch (e) {}
    }
    return id;
  }
  const configured = () => !!MASTER || Object.keys(MONTHS).length > 0;

  /* ── 업체명(아이디) 승인 ─────────────────────────────────────────
     한 주소를 여러 회사가 같이 쓰기 때문에, 아무나 들어오지 못하게 업체명을 받는다.
     승인은 마스터가 발급한 '승인코드' 로 한 번만 하면 그 기기에 기억된다.

     ※ 서버가 없는 정적 사이트라, 소스를 읽을 줄 아는 사람은 이 검사를 우회할 수 있다.
       월별 비밀번호와 같은 한계다. 실수·무단 사용을 막는 문턱으로 쓴다.
       진짜 자료는 각 회사의 구글 계정으로 갈려 있어, 남의 계정 자료는 못 본다. */
  const MASTER_ID = "7167ccfda76104466cce86b23ad897c6473a5342175fb6b2ac32c68e6d2bce7e";
  const MASTER_PW = "27cbec30fa8a607bceb852c4b41ef7c3c494edb033dcd2b25ef5d3339dd35f4b";
  const normId = v => String(v == null ? "" : v).trim().replace(/\s+/g, "");
  const idHash = v => sha256Hex(SALT + "|id|" + normId(v));
  /* 업체명 + 그 달 → 승인번호 (마스터가 발급해 업체에 알려준다)
     ★ 달마다 번호가 바뀐다. 그래서 '막는 방법 = 다음 달 번호를 안 주는 것' 이 된다.
       서버가 없어 마스터의 중지 스위치가 업체 기기까지 전달될 수 없기 때문에,
       유효기간으로 통제하는 방식을 택했다. (2026-08 결정) */
  async function approvalCode(company, ym) {
    const h = await sha256Hex(SALT + "|approve|" + normId(company) + "|" + (ym || currentYm()));
    return h.slice(0, 10).toUpperCase();
  }
  /* 이 달의 마지막 순간. 로그인 유지도 여기서 끊어야 번호를 다시 받게 된다. */
  function monthEnd() {
    const d = new Date();
    return new Date(d.getFullYear(), d.getMonth() + 1, 1).getTime() - 1;
  }
  const APPROVED_KEY = () => CONFIG.lsBase("qo_approved");
  /* 마스터로 들어온 경우는 이 탭에서만 유지해야 해서 sessionStorage 에 둔다.
     둘 다 있으면 세션(이 탭) 것이 먼저다. */
  const sstore = () => { try { return window.sessionStorage; } catch (e) { return null; } };
  const readAt = (st, k) => { try { return st ? st.getItem(k) : null; } catch (e) { return null; } };
  async function savedApproval() {
    for (const raw of [readAt(sstore(), APPROVED_KEY()), readAt(localStorage, APPROVED_KEY())]) {
      try {
        const s2 = JSON.parse(raw || "null");
        if (!s2 || !s2.company) continue;
        if (s2.token !== await sha256Hex(deviceId() + SALT + "|ok|" + normId(s2.company))) continue;
        return s2.company;
      } catch (e) {}
    }
    return null;
  }
  async function saveApproval(company, tabOnly) {
    try {
      const v = JSON.stringify({
        company: normId(company), token: await sha256Hex(deviceId() + SALT + "|ok|" + normId(company)) });
      (tabOnly ? sstore() : localStorage).setItem(APPROVED_KEY(), v);
    } catch (e) {}
  }
  function clearApproval() {
    try { localStorage.removeItem(APPROVED_KEY()); } catch (e) {}
    try { sstore().removeItem(APPROVED_KEY()); } catch (e) {}
  }
  async function isMaster(id, pw) {
    return (await idHash(id)) === MASTER_ID && (await hashPw(pw)) === MASTER_PW;
  }
  /* 마스터로 들어왔는지 — 앱 안의 '마스터' 메뉴를 띄울지 판단한다 */
  /* ★ 마스터는 '이 탭에서만' 유지한다 (sessionStorage).
     예전엔 localStorage 라 한 번 들어가면 새 창·새 탭도 계속 마스터로 열렸다.
     승인번호를 발급하는 관리자 권한이라, 창을 새로 열면 다시 확인받는 게 맞다. */
  const MKEY = () => CONFIG.lsBase("qo_master");
  const mstore = () => { try { return window.sessionStorage; } catch (e) { return null; } };
  async function saveMaster() {
    try { mstore().setItem(MKEY(), await sha256Hex(deviceId() + SALT + "|master|")); } catch (e) {}
    try { localStorage.removeItem(MKEY()); } catch (e) {}   // 예전에 남은 것 정리
  }
  async function isMasterSession() {
    try {
      const st = mstore(); if (!st) return false;
      return st.getItem(MKEY()) === await sha256Hex(deviceId() + SALT + "|master|");
    } catch (e) { return false; }
  }
  function clearMaster() {
    try { mstore().removeItem(MKEY()); } catch (e) {}
    try { localStorage.removeItem(MKEY()); } catch (e) {}
  }

  /* v1.3.7~1.3.9 에서는 잠금 상태가 계정별 키(…_u<해시>)에 저장됐다.
     v1.4.0 에서 계정과 무관한 키로 옮겼는데, 그 사이 마스터로 들어와 있던 사람은
     표시가 옛 키에 남아 마스터 메뉴가 사라져 보였다. 기기ID 도 같이 갈렸을 수 있어
     남아 있는 옛 기기ID 를 전부 대입해 본다. 한 번 옮기면 다시 안 돈다. */
  function legacyKeys(name) {
    const base = CONFIG.lsBase(name), out = [];
    try {
      for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i);
        if (k !== base && k.indexOf(base + "_u") === 0) out.push(k);
      }
    } catch (e) {}
    return out;
  }
  /* 옛 버전이 localStorage 에 남긴 마스터 표시를 지운다.
     남겨두면 창을 새로 열 때마다 마스터로 들어가진다 (이제는 탭 단위가 원칙이다). */
  async function migrateMaster() {
    try {
      localStorage.removeItem(MKEY());
      legacyKeys("qo_master").forEach(k => { try { localStorage.removeItem(k); } catch (e) {} });
    } catch (e) {}
  }

  /* ── 사용 중지 명단 ──────────────────────────────────────────────
     서버가 없으므로 앱과 같은 주소에 놓인 roster.json 을 읽어 판정한다.
     같은 출처라 로그인도 CORS 도 필요 없다.
     ★ 규칙: 명단에 '사용 중지' 로 적힌 업체만 막는다.
       파일이 없거나·못 읽거나·명단에 없으면 통과시킨다. 새로 승인한 업체가
       배포를 기다리느라 못 들어오는 일이 없어야 하고, 인터넷이 흔들린다고
       멀쩡한 업체를 잠그면 안 되기 때문이다. */
  let rosterCache = null;
  async function roster() {
    if (rosterCache) return rosterCache;
    try {
      const url = "roster.json?ts=" + Math.floor(Date.now() / 60000);   // 1분 단위로만 새로 받는다
      const ctl = typeof AbortController !== "undefined" ? new AbortController() : null;
      const t = setTimeout(() => { try { ctl && ctl.abort(); } catch (e) {} }, 4000);
      const r = await fetch(url, { cache: "no-cache", signal: ctl ? ctl.signal : undefined });
      clearTimeout(t);
      if (!r.ok) throw new Error("no roster");
      rosterCache = await r.json();
    } catch (e) { rosterCache = { companies: [] }; }
    return rosterCache;
  }
  /* 막힌 업체면 true. 판단이 안 서면 false(통과) */
  async function blocked(company) {
    try {
      const n = normId(company); if (!n) return false;
      const r = await roster();
      const hit = (r.companies || []).find(c => normId(c.name) === n);
      return !!(hit && hit.on === false);
    } catch (e) { return false; }
  }
  const BLOCK_MSG = "승인받지 않은 아이디입니다.\n관리자에게 문의하세요.";
  /* ★ 승인명단(구글 시트) 확인을 로그인 화면에서 한다.
     예전엔 앱에 들어간 뒤 확인해서 화면이 잠깐 보였다가 팝업이 떴다 — 들어가기 전에 막는다.
     · 그 업체가 쓰던 구글 계정의 토큰으로 읽는다. 토큰이 없으면 통과시킨다
       (여기서 로그인 창을 띄우면 안 된다 — 사용자가 누른 적이 없다).
     · 못 읽어도 통과. 인터넷이 흔들린다고 멀쩡한 업체를 잠그면 안 된다. */
  async function blockedBySheet(company) {
    try {
      const id = (CONFIG && CONFIG.rosterSheetId) || "";
      if (!id || typeof GMAIL === "undefined") return false;
      CONFIG.useCompany(company);
      let last = "";
      try { last = localStorage.getItem(CONFIG.lsCompany("qo_last_account")) || ""; } catch (e) {}
      CONFIG.useAccount(last);
      GMAIL.reloadToken();
      if (!GMAIL.signedIn()) return false;
      const tabs = await GMAIL.sheetTabs(id);
      const rows = await GMAIL.sheetRead(id, tabs[0] || "시트1");
      if (!rows || rows.length < 2) return false;
      const hit = rows.slice(1).find(r => normId(r[0]) === normId(company));
      return !!(hit && /중지|off|false/i.test(String(hit[1] || "")));
    } catch (e) { return false; }
  }

  async function verify(pw) {
    const h = await hashPw(pw);
    if (MASTER && h === MASTER) return "master";
    const ym = currentYm();
    if (MONTHS[ym] && h === MONTHS[ym]) return "month";
    return false;
  }
  /* 예전에는 7일 슬라이딩이었다. 승인번호가 달마다 바뀌므로 로그인 유지도 그 달까지다 —
     안 그러면 번호가 바뀌어도 계속 들어와 있어서 유효기간이 무의미해진다. */
  const signExp = exp => sha256Hex(deviceId() + SALT + "|exp|" + exp);

  async function isUnlocked() {
    if (!configured()) return true;
    const tabOnly = !!readAt(sstore(), CONFIG.lsBase("qo_lock"));
    for (const raw of [readAt(sstore(), CONFIG.lsBase("qo_lock")), readAt(localStorage, CONFIG.lsBase("qo_lock"))]) {
      try {
        const s = JSON.parse(raw || "null");
        if (!s || !s.exp || Date.now() >= s.exp) continue;      // 없거나 이 달이 지남
        if (s.token !== await signExp(s.exp)) continue;
        await saveUnlock(tabOnly);   // 슬라이딩: 열 때마다 만료를 다시 7일 뒤로 연장
        return true;
      } catch (e) {}
    }
    return false;
  }
  async function saveUnlock(tabOnly) {
    const exp = monthEnd();
    try {
      const v = JSON.stringify({ exp, token: await signExp(exp) });
      (tabOnly ? sstore() : localStorage).setItem(CONFIG.lsBase("qo_lock"), v);
    } catch (e) {}
  }
  function signOut() {
    try { localStorage.removeItem(CONFIG.lsBase("qo_lock")); } catch (e) {}
    try { sstore().removeItem(CONFIG.lsBase("qo_lock")); } catch (e) {}
    clearMaster(); clearApproval();
  }

  /* ---------- 잠금 화면 UI ---------- */
  function injectStyle() {
    if (document.getElementById("qo-lock-style")) return;
    const st = document.createElement("style");
    st.id = "qo-lock-style";
    st.textContent = `
      #qo-lock{position:fixed;inset:0;z-index:2147483647;display:flex;align-items:center;justify-content:center;
        background:linear-gradient(160deg,#eef2ff,#f8fafc);font-family:-apple-system,BlinkMacSystemFont,"맑은 고딕",sans-serif}
      #qo-lock .card{width:min(92vw,360px);background:#fff;border-radius:20px;padding:28px 24px;
        box-shadow:0 12px 40px rgba(30,50,120,.18);text-align:center}
      #qo-lock .lk{font-size:40px;margin-bottom:6px}
      #qo-lock h2{margin:0 0 4px;font-size:19px;color:#1e2b6b}
      #qo-lock p{margin:0 0 18px;font-size:13px;color:#6b7280;line-height:1.5}
      #qo-lock input{width:100%;box-sizing:border-box;padding:14px 14px;font-size:17px;text-align:center;
        border:2px solid #dfe3ee;border-radius:12px;outline:none;letter-spacing:1px}
      #qo-lock input:focus{border-color:#4f6ef7}
      #qo-lock button{width:100%;margin-top:12px;padding:14px;font-size:16px;font-weight:700;color:#fff;
        background:#4f6ef7;border:none;border-radius:12px;cursor:pointer}
      #qo-lock button:disabled{opacity:.5}
      #qo-lock .msg{min-height:18px;margin-top:12px;font-size:13px;color:#e5484d;font-weight:600}
      #qo-lock .ver{margin-top:14px;font-size:12px;color:#9aa3b2;letter-spacing:.5px}
      #qo-lock .signup{margin-top:14px;font-size:12.5px}
      #qo-lock .signup a{color:#4f6ef7;font-weight:700;text-decoration:none}
      #qo-lock #qo-signup{margin-top:12px;padding-top:12px;border-top:1px solid #eef0f6}
      #qo-lock #qo-signup input{font-size:14px;padding:11px;text-align:left}
      #qo-lock #qo-signup button{margin-top:9px;padding:12px;font-size:14px}
    `;
    document.head.appendChild(st);
  }
  function buildOverlay() {
    const root = document.createElement("div");
    root.id = "qo-lock";
    root.innerHTML =
      "<div class=\"card\">" +
      "<div class=\"lk\">🔒</div>" +
      "<h2>퀵오더 사용 승인</h2>" +
      "<p>업체명과 <b>이번 달</b> 승인번호를 입력하세요.<br>" + (CONFIG.adminLabel || "관리자") + "에게 전달받습니다.</p>" +
      "<input id=\"qo-lock-id\" type=\"text\" autocomplete=\"username\" autocapitalize=\"off\" " +
        "autocorrect=\"off\" placeholder=\"업체명\" style=\"letter-spacing:0;margin-bottom:8px\">" +
      "<input id=\"qo-lock-pw\" type=\"password\" inputmode=\"text\" autocomplete=\"current-password\" " +
        "autocapitalize=\"characters\" autocorrect=\"off\" placeholder=\"승인번호\">" +
      "<button id=\"qo-lock-go\">확인</button>" +
      "<div class=\"msg\" id=\"qo-lock-msg\"></div>" +
      "<div id=\"qo-lock-extra\"></div>" +
      "<div class=\"signup\"><a href=\"#\" id=\"qo-signup-open\">처음이신가요? 사용 신청하기</a></div>" +
      "<div id=\"qo-signup\" style=\"display:none\">" +
        "<input id=\"su-name\" type=\"text\" placeholder=\"업체명\" autocapitalize=\"off\" style=\"letter-spacing:0;margin-bottom:7px\">" +
        "<input id=\"su-mail\" type=\"email\" placeholder=\"이메일 (승인번호를 받을 주소)\" autocapitalize=\"off\" style=\"letter-spacing:0;margin-bottom:7px\">" +
        "<input id=\"su-tel\" type=\"text\" placeholder=\"담당자 연락처\" autocapitalize=\"off\" style=\"letter-spacing:0\">" +
        "<button id=\"qo-signup-go\">신청서 보내기</button>" +
        "<div class=\"msg\" id=\"su-msg\" style=\"color:var(--muted)\"></div>" +
      "</div>" +
      /* 최신본을 받았는지 로그인 화면에서 바로 확인할 수 있게 버전을 보여준다 */
      "<div class=\"ver\">" + (typeof APP_VER !== "undefined" ? "v" + APP_VER : "") + "</div>" +
      "</div>";
    return root;
  }
  function ensureUnlocked() {
    return new Promise(resolve => {
      if (!configured()) return resolve(true);
      injectStyle();
      const root = buildOverlay();
      const attach = () => document.body.appendChild(root);
      if (document.body) attach(); else document.addEventListener("DOMContentLoaded", attach);
      const idIn = root.querySelector("#qo-lock-id");
      const input = root.querySelector("#qo-lock-pw");
      const btn = root.querySelector("#qo-lock-go");
      const msg = root.querySelector("#qo-lock-msg");
      const extra = root.querySelector("#qo-lock-extra");

      isUnlocked().then(async ok => {
        const approved = await savedApproval();
        if (ok && approved && (await blocked(approved) || await blockedBySheet(approved))) {
          signOut(); ok = false;                       // 중지된 업체는 기억된 로그인도 풀어버린다
          msg.style.color = "#e5484d"; msg.textContent = BLOCK_MSG;
        }
        if (ok && approved) { root.remove(); return resolve(true); }
        if (approved) idIn.value = approved;
        setTimeout(() => { try { (approved ? input : idIn).focus(); } catch (e) {} }, 100);

        let busy = false;
        const BAD = "업체명 또는 승인번호가 틀렸습니다.\n승인번호는 달마다 바뀝니다 — 이번 달 번호인지 확인해 주세요.";
        async function go() {
          if (busy) return; busy = true; btn.disabled = true; msg.textContent = "";
          const id = idIn.value, pw = input.value;
          /* 빨간 글씨만으로는 못 보고 지나친다 — 팝업으로도 알린다 */
          const fail = () => { msg.style.color = "#e5484d"; msg.textContent = BAD;
            input.value = ""; busy = false; btn.disabled = false;
            try { alert(BAD); } catch (e) {}
            try { input.focus(); } catch (e) {} };
          if (!normId(id) || !String(pw).trim()) return fail();

          // ① 마스터 — 앱으로 들여보내고, 안에서 '마스터' 메뉴가 뜬다
          if (await isMaster(id, pw)) {
            /* 관리자 권한이라 흔적을 기기에 남기지 않는다 — 창을 새로 열면 다시 확인받는다 */
            await saveMaster(); await saveApproval(id, true); await saveUnlock(true);
            root.remove(); return resolve(true);
          }
          clearMaster();

          if (await blocked(id)) { msg.style.color = "#e5484d"; msg.textContent = BLOCK_MSG;
            try { alert(BLOCK_MSG); } catch (e) {}
            input.value = ""; busy = false; btn.disabled = false; return; }

          /* ② 업체명 + 승인번호 — 이게 곧 비밀번호다. 두 단계로 나누지 않는다.
             안내문을 통째로 붙여넣는 일이 잦아, 그 안에 번호가 들어 있으면 받아준다.
             (번호를 모르면 여전히 못 들어온다 — 관대할 뿐 느슨하지 않다) */
          const typed = normId(pw).toUpperCase(), want = await approvalCode(id);
          if (typed === want || typed.indexOf(want) >= 0) {
            if (await blockedBySheet(id)) {
              msg.style.color = "#e5484d"; msg.textContent = BLOCK_MSG;
              try { alert(BLOCK_MSG); } catch (e) {}
              input.value = ""; busy = false; btn.disabled = false; return;
            }
            await saveApproval(id); await saveUnlock();
            root.remove(); return resolve(true);
          }
          /* ③ 예비 경로 — 예전 월별 비밀번호도 그대로 받는다.
             이미 승인된 업체가 승인번호를 아직 못 받았을 때 갑자기 막히지 않게. */
          const okName = await savedApproval();
          if (normId(okName) === normId(id) && await verify(pw)) {
            await saveUnlock(); root.remove(); return resolve(true);
          }
          fail();
        }
        btn.addEventListener("click", go);
        input.addEventListener("keydown", e => { if (e.key === "Enter") go(); });
        idIn.addEventListener("keydown", e => { if (e.key === "Enter") input.focus(); });

        /* ── 사용 신청 ──────────────────────────────────────────────
           서버가 없으니 신청 내용을 '신청 코드' 한 줄로 만들어 준다.
           예전엔 메일 앱을 열었는데, 모바일에서 잘 막히고 업체가 쓰는 메일이
           회사 메일이 아닌 경우도 있어 카톡·문자로도 보낼 수 있게 바꿨다.
           마스터 화면에서 이 코드를 붙여넣으면 신청서로 펼쳐진다. */
        function makeCode(o) {
          const bytes = new TextEncoder().encode(JSON.stringify(o));
          let bin = ""; bytes.forEach(b => { bin += String.fromCharCode(b); });
          return "QO1-" + btoa(bin).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
        }
        const su = root.querySelector("#qo-signup");
        const suLink = root.querySelector("#qo-signup-open");
        const suMsg = root.querySelector("#su-msg");
        suLink.addEventListener("click", e => {
          e.preventDefault();
          /* 구글폼을 세팅해 뒀으면 그걸 연다 — 업체는 폼만 채우면 되고,
             마스터는 응답 시트를 퀵오더 안에서 바로 본다. */
          const formUrl = (CONFIG && CONFIG.signupFormUrl) || "";
          if (formUrl) { window.open(formUrl, "_blank", "noopener"); return; }
          const open = su.style.display === "none";
          su.style.display = open ? "block" : "none";
          suLink.textContent = open ? "신청 접기" : "처음이신가요? 사용 신청하기";
          if (open) setTimeout(() => { try { root.querySelector("#su-name").focus(); } catch (e2) {} }, 60);
        });
        root.querySelector("#qo-signup-go").addEventListener("click", () => {
          const name = root.querySelector("#su-name").value.trim();
          const mail = root.querySelector("#su-mail").value.trim();
          const tel = root.querySelector("#su-tel").value.trim();
          if (!name || !mail) {
            suMsg.style.color = "#e5484d";
            suMsg.textContent = "업체명과 이메일은 꼭 넣어주세요.";
            return;
          }
          const code = makeCode({ n: name, e: mail, t: tel });
          const out = root.querySelector("#su-out");
          const cp = root.querySelector("#su-copy");
          out.value = code; out.style.display = "block"; cp.style.display = "block";
          suMsg.style.color = "#6b7280";
          suMsg.textContent = "이 코드를 " + ((CONFIG && CONFIG.adminLabel) || "관리자") +
            " 에게 보내주세요. 카톡·문자·메일 아무거나 괜찮습니다.\n승인되면 승인번호를 받으시게 됩니다.";
          cp.onclick = async () => {
            try { await navigator.clipboard.writeText(code); cp.textContent = "복사됨 ✓"; }
            catch (e2) { out.select(); try { document.execCommand("copy"); cp.textContent = "복사됨 ✓"; } catch (e3) {} }
            setTimeout(() => { cp.textContent = "코드 복사"; }, 1800);
          };
        });
      });
    });
  }

  const ready = migrateMaster().then(ensureUnlocked);
  /* 앱 안에서 마스터로 들어오기 — 저장된 표시가 없어도(기기를 바꿨거나 표시가 지워졌어도)
     아이디·비밀번호만 맞으면 언제든 마스터 메뉴를 열 수 있다. */
  async function signInMaster(id, pw) {
    if (!await isMaster(id, pw)) return false;
    await saveMaster();          // 이 탭에서만. 업체로 들어와 있던 상태는 건드리지 않는다
    return true;
  }

  return { ready, ensureUnlocked, isUnlocked, verify, signOut, configured, currentYm, signInMaster,
           approvalCode, savedApproval, clearApproval, isMaster, blocked, monthEnd,
           isMasterSession, clearMaster,
           company: () => {
             for (const raw of [readAt(sstore(), APPROVED_KEY()), readAt(localStorage, APPROVED_KEY())]) {
               try { const c = (JSON.parse(raw || "null") || {}).company; if (c) return c; } catch (e) {}
             }
             return "";
           } };
})();
