/* ===================================================================
   퀵오더 — 사용 승인(월별 비밀번호) 잠금  [웹 전용 비밀번호]
   "[JK전용] 퀵오더웹 비밀번호 2년치.txt" 의 비밀번호로 로그인.
   * 소스에는 비밀번호의 해시만 있어, 코드를 봐도 비밀번호를 알 수 없음.
   * 관리자 마스터 비밀번호는 만료 없이 통과(저장하지 않음).
   * 월 비밀번호로 들어오면 그 달 동안 이 기기에서 재입력 불필요.
   =================================================================== */
"use strict";
const LOCK = (() => {
  const SALT = "047ec6b655a0775637a17f233ce04bd4";
  const MASTER = "4e3782b5293b1e04564a07aa343a25638ff682519529454c5898d0d984907c35";
  const MONTHS = {
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
  };

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
    try { id = localStorage.getItem("qo_device"); } catch (e) {}
    if (!id) {
      const r = crypto.getRandomValues(new Uint8Array(16));
      id = [...r].map(b => b.toString(16).padStart(2, "0")).join("");
      try { localStorage.setItem("qo_device", id); } catch (e) {}
    }
    return id;
  }
  const configured = () => !!MASTER || Object.keys(MONTHS).length > 0;

  async function verify(pw) {
    const h = await hashPw(pw);
    if (MASTER && h === MASTER) return "master";
    const ym = currentYm();
    if (MONTHS[ym] && h === MONTHS[ym]) return "month";
    return false;
  }
  const unlockToken = ym => sha256Hex(deviceId() + SALT + ym + (MONTHS[ym] || ""));

  async function isUnlocked() {
    if (!configured()) return true;
    try {
      const s = JSON.parse(localStorage.getItem("qo_lock") || "null");
      if (!s || s.month !== currentYm()) return false;
      return s.token === await unlockToken(s.month);
    } catch (e) { return false; }
  }
  async function saveUnlock() {
    const ym = currentYm();
    try { localStorage.setItem("qo_lock", JSON.stringify({ month: ym, token: await unlockToken(ym) })); } catch (e) {}
  }
  function signOut() { try { localStorage.removeItem("qo_lock"); } catch (e) {} }

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
      "<p>이번 달 비밀번호를 입력하세요.<br>관리자(JK)에게 매달 전달받습니다.</p>" +
      "<input id=\"qo-lock-pw\" type=\"password\" inputmode=\"text\" autocomplete=\"off\" " +
        "autocapitalize=\"characters\" autocorrect=\"off\" placeholder=\"비밀번호\">" +
      "<button id=\"qo-lock-go\">확인</button>" +
      "<div class=\"msg\" id=\"qo-lock-msg\"></div>" +
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
      const input = root.querySelector("#qo-lock-pw");
      const btn = root.querySelector("#qo-lock-go");
      const msg = root.querySelector("#qo-lock-msg");

      isUnlocked().then(ok => {
        if (ok) { root.remove(); return resolve(true); }
        setTimeout(() => { try { input.focus(); } catch (e) {} }, 100);
        if (!MONTHS[currentYm()] && MASTER)
          msg.textContent = "이번 달 비밀번호가 없습니다. 관리자에게 문의하세요.";
        let busy = false;
        async function go() {
          if (busy) return; busy = true; btn.disabled = true; msg.textContent = "";
          const kind = await verify(input.value);
          if (kind === "month") { await saveUnlock(); root.remove(); resolve(true); }
          else if (kind === "master") { root.remove(); resolve(true); }
          else { msg.textContent = "비밀번호가 올바르지 않습니다."; input.value = ""; input.focus(); busy = false; btn.disabled = false; }
        }
        btn.addEventListener("click", go);
        input.addEventListener("keydown", e => { if (e.key === "Enter") go(); });
      });
    });
  }

  const ready = ensureUnlocked();
  return { ready, ensureUnlocked, isUnlocked, verify, signOut, configured, currentYm };
})();
