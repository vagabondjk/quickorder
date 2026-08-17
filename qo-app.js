/* ===================================================================
   퀵오더 앱 — 화면 동작 + 기기 저장(IndexedDB)
   =================================================================== */
"use strict";
// 구글 클라이언트 ID (기본 내장) — github 주소에서만 작동하게 묶여 있어 공개돼도 안전.
// 기기·주소가 바뀌어도 다시 입력할 필요가 없다.
const DEFAULT_CLIENT_ID = CONFIG.clientId;   // 회사별 구글 프로젝트 (qo-config.js)
async function clientId() { return (await DB.get("gmailClientId", "")) || DEFAULT_CLIENT_ID; }

const CHK = '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"><path d="M5 13l4 4L19 7"/></svg>';

// 여러 이메일(쉼표·세미콜론·줄바꿈·공백 구분) → 정리된 배열
function parseEmails(str) {
  const seen = new Set(), out = [];
  for (const raw of String(str || "").split(/[,;\s]+/)) {
    const e = raw.trim();
    if (!e) continue;
    const key = e.toLowerCase();
    if (!seen.has(key)) { seen.add(key); out.push(e); }
  }
  return out;
}
function invalidEmails(list) { return list.filter(e => !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e)); }
// 업체명 정리: 끝에 붙은 "발주양식/발주서/양식" 제거 → "디에스피_발주양식" → "디에스피"
function cleanVendor(n) {
  const s = String(n == null ? "" : n).replace(/[_\s]*(발주\s*양식|발주\s*서|양식)\s*$/g, "").trim();
  return s || String(n == null ? "" : n).trim();
}
// "@dsp.com, onekglobal.co.kr" → ["dsp.com","onekglobal.co.kr"]
function parseDomains(str) {
  const out = [];
  for (const raw of String(str || "").split(/[,;\s]+/)) {
    const d = raw.trim().replace(/^@/, "").toLowerCase();
    if (d && !out.includes(d)) out.push(d);
  }
  return out;
}
const DT_SHOW = 3;   // 날짜 칩은 최근 3개만 보이고 나머지는 "더보기"
const EYE = '미리보기';   // 미리보기 버튼 라벨(텍스트)

const $ = id => document.getElementById(id);
const esc = s => String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

/* ---------------- 엑셀 미리보기 (어떤 파일이든) ---------------- */
async function openPreview(buf, title) {
  const m = $("pvmodal");
  m.classList.add("on");
  $("pv-modal-title").textContent = title || "미리보기";
  $("pv-modal-sub").textContent = "읽는 중…";
  $("pv-modal-table").innerHTML = "";
  $("pv-modal-foot").textContent = "";
  try {
    const wb = await QO.loadWorkbook(buf.slice(0));
    // 시트가 여럿이면 골라서 볼 수 있게 한다 (공급가표처럼 4장짜리 파일이 있다)
    const all = QO.previewSheets(wb, 2000).filter(s => s.columns.length || s.rows.length);
    const big = QO.previewAny(wb, 2000);
    if (!all.length) { $("pv-modal-sub").textContent = ""; $("pv-modal-foot").textContent = "표시할 내용이 없어요."; return; }
    let cur = Math.max(0, all.findIndex(s => s.name === big.sheet));

    const drawTab = () => {
      const bar = $("pv-modal-tabs");
      if (!bar) return;
      if (all.length < 2) { bar.innerHTML = ""; bar.style.display = "none"; return; }
      bar.style.display = "flex";
      bar.innerHTML = all.map((s, i) =>
        `<button class="seg${i === cur ? " on" : ""}" data-i="${i}">${esc(s.name)} <span style="opacity:.6">${s.rows.length}</span></button>`).join("");
      bar.querySelectorAll("button").forEach(b => b.onclick = () => { cur = Number(b.dataset.i); drawTab(); drawSheet(); });
    };
    const drawSheet = () => {
      const pv = all[cur];
      $("pv-modal-sub").textContent = `시트: ${pv.name} · 전체 ${pv.rows.length}건` +
        (all.length > 1 ? ` · (${all.length}개 시트 중 ${cur + 1}번째)` : "");
      if (!pv.columns.length && !pv.rows.length) {
        $("pv-modal-table").innerHTML = ""; $("pv-modal-foot").textContent = "이 시트에는 내용이 없어요."; return;
      }
      if (!pv.rows.length) {
        $("pv-modal-sub").textContent = `시트: ${pv.name} · 빈 양식(내용 없음)`;
        $("pv-modal-foot").innerHTML = "ℹ️ 이 시트는 <b>빈 양식(템플릿)</b>이라 채워진 내용이 없습니다. 아래는 열(항목) 목록입니다.";
      }
      let h = "<tr>" + pv.columns.map((c, i) => `<th>${esc(c || "열" + (i + 1))}</th>`).join("") + "</tr>";
      pv.rows.forEach(row => {
        h += "<tr>" + pv.columns.map((_, i) => {
          const v = row[i] == null ? "" : row[i];
          const num = /^[0-9,.\-]+$/.test(v) && v !== "";
          return `<td${num ? ' class="num"' : ""}>${esc(v)}</td>`;
        }).join("") + "</tr>";
      });
      $("pv-modal-table").innerHTML = h;
      if (pv.rows.length) $("pv-modal-foot").textContent = `전체 ${pv.rows.length}건`;
    };
    drawTab(); drawSheet();
  } catch (e) {
    $("pv-modal-sub").textContent = "";
    $("pv-modal-foot").textContent = "⚠ 미리보기 실패: " + e.message;
  }
}
$("pv-modal-close").onclick = () => $("pvmodal").classList.remove("on");
$("pvmodal").onclick = e => { if (e.target === $("pvmodal")) $("pvmodal").classList.remove("on"); };
$("sab-preview").onclick = () => { if (S.sabBuf) openPreview(S.sabBuf, "송장취합양식"); };

/* ---------------- 기기 저장소 (IndexedDB) ---------------- */
const DB = (() => {
  let db = null;
  function open() {
    return new Promise((res, rej) => {
      if (db) return res(db);
      const rq = indexedDB.open(CONFIG.dbName, 1);   // 회사별로 저장소 분리 (qo-config.js)
      rq.onupgradeneeded = e => {
        const d = e.target.result;
        if (!d.objectStoreNames.contains("forms")) d.createObjectStore("forms", { keyPath: "name" });
        if (!d.objectStoreNames.contains("kv")) d.createObjectStore("kv", { keyPath: "k" });
      };
      rq.onsuccess = e => { db = e.target.result; res(db); };
      rq.onerror = e => rej(e.target.error);
    });
  }
  async function tx(store, mode, fn) {
    const d = await open();
    return new Promise((res, rej) => {
      const t = d.transaction(store, mode), s = t.objectStore(store);
      const r = fn(s);
      t.oncomplete = () => res(r && r.result !== undefined ? r.result : r);
      t.onerror = () => rej(t.error);
    });
  }
  let afterWrite = () => {}, suspended = false;
  const fire = () => { if (!suspended) { try { afterWrite(); } catch (e) {} } };
  return {
    listForms: () => tx("forms", "readonly", s => s.getAll()),
    putForm: async f => { const r = await tx("forms", "readwrite", s => s.put(f)); fire(); return r; },
    delForm: async n => { const r = await tx("forms", "readwrite", s => s.delete(n)); fire(); return r; },
    get: async (k, dflt) => { const v = await tx("kv", "readonly", s => s.get(k)); return v && v.v !== undefined ? v.v : dflt; },
    set: async (k, v) => { const r = await tx("kv", "readwrite", s => s.put({ k, v })); fire(); return r; },
    onChange(fn) { afterWrite = fn; },       // 데이터가 바뀔 때마다 호출 (자동 업로드용)
    suspend(b) { suspended = b; },           // 복원 중 자동업로드 방지
    /* 로그인 계정이 바뀌면 저장소 이름도 바뀐다 → 열려 있던 것을 닫고 다시 연다.
       이걸 안 하면 계정을 바꿔도 앞 회사 데이터를 계속 보게 된다. */
    reopen() { try { if (db) db.close(); } catch (e) {} db = null; return open(); },
  };
})();

/* ---------------- 공통 상태 ---------------- */
const S = {
  orderWb: null, orderBuf: null, orderName: "",
  brands: [], dateAll: [], dateSel: [], dateHeader: null,
  pv: null, pvAll: false,
  forms: [], brandVendor: {}, vendorEmails: {}, vendorSent: {}, vendorDomains: {}, sel: {},
  invEmails: "", invSent: [],
  sabBuf: null, sabName: "", sabDrive: null, reps: [],
};

function msg(el, kind, text) { const m = $(el); m.className = "msg" + (kind ? " show " + kind : ""); m.textContent = text; }
function busy(btnId, lblId, on, text) {
  const b = $(btnId); b.classList.toggle("loading", on); b.disabled = on;
  if (text) $(lblId).textContent = text;
}
function download(buf, filename) {
  const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}
/* 파일을 '공유하기'(휴대폰 기본 공유 시트 → 카카오톡 등 선택). 안 되면 다운로드로 폴백. */
async function shareFile(buf, filename) {
  try {
    const file = new File([buf.slice ? buf.slice(0) : buf], filename,
      { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      await navigator.share({ files: [file], title: filename });
      return true;
    }
  } catch (e) { if (e && e.name === "AbortError") return false; }  // 사용자가 취소
  // 공유 미지원(주로 PC) → 다운로드 후 안내
  download(buf, filename);
  alert("이 브라우저는 파일 바로 공유를 지원하지 않아 다운로드했어요.\n저장된 파일을 카카오톡 대화창에 첨부해 보내주세요.\n(휴대폰에서는 '공유' 버튼으로 카카오톡에 바로 보낼 수 있습니다)");
  return false;
}
function bindDrop(id, cb) {
  const el = $(id);
  ["dragover", "dragenter"].forEach(e => el.addEventListener(e, ev => { ev.preventDefault(); el.classList.add("hi"); }));
  ["dragleave", "drop"].forEach(e => el.addEventListener(e, ev => { ev.preventDefault(); el.classList.remove("hi"); }));
  el.addEventListener("drop", ev => {
    const f = [...ev.dataTransfer.files].filter(x => /\.xls[xm]$/i.test(x.name));
    if (f.length) cb(f);
  });
}
let readFile = f => new Promise((res, rej) => {
  const r = new FileReader();
  r.onload = () => res(r.result); r.onerror = () => rej(r.error);
  r.readAsArrayBuffer(f);
});

/* =====================================================================
   로딩 표시 — 엑셀을 읽고 쓰는 데 몇 초씩 걸리는데 화면이 그대로라
   멈춘 줄 알고 다시 누르는 일이 있었다. 오래 걸리는 일에는 달리는 사람을 띄운다.

   버튼마다 붙이지 않고 '느린 일' 자체를 감싼다 —
   파일 읽기 · 엑셀 열기/저장 · 드라이브/메일 통신. 어디서 부르든 자동으로 뜬다.
   겹쳐 불려도 세어서 마지막 하나가 끝날 때 닫는다.
   250ms 안에 끝나면 아예 뜨지 않는다 (깜빡임 방지).
   ===================================================================== */
const BUSY = (() => {
  let n = 0, timer = null, label = "";
  const el = () => document.getElementById("busy");
  function paint() {
    const b = el(); if (!b) return;
    const s = document.getElementById("busy-sub");
    if (s) s.textContent = label || "";
    b.classList.add("on");
  }
  /* 진행률 표시 — 실제로 셀 수 있는 작업에서만 부른다.
     모르는 작업에 가짜 숫자를 띄우면 '멈춘 것처럼' 보여서 더 나쁘다. */
  function progress(done, total) {
    const b = el(); if (!b) return;
    const box = b.querySelector(".busybox"), bar = b.querySelector(".busytrack i"), p = document.getElementById("busy-pct");
    if (!total || total < 0) { if (box) box.classList.remove("pct"); if (p) p.textContent = ""; if (bar) bar.style.width = ""; return; }
    const v = Math.max(0, Math.min(100, Math.round((done / total) * 100)));
    if (box) box.classList.add("pct");
    if (bar) bar.style.width = v + "%";
    if (p) p.textContent = v + "%" + (total > 1 ? `  (${done}/${total})` : "");
  }
  function start(txt) {
    n++; label = txt || label;
    if (timer || el() && el().classList.contains("on")) return;
    timer = setTimeout(() => { timer = null; if (n > 0) paint(); }, 250);
  }
  function end() {
    n = Math.max(0, n - 1);
    if (n) return;
    label = "";
    if (timer) { clearTimeout(timer); timer = null; }
    const b = el(); if (b) b.classList.remove("on");
    progress(0, 0);                       // 다음 작업이 옛 퍼센트를 물려받지 않게

  }
  /* 함수 하나를 '로딩 뜨는 함수'로 바꿔 끼운다 */
  function wrap(get, set, txt) {
    const f = get();
    if (typeof f !== "function" || f.__busy) return;
    const g = function (...a) {
      start(txt);
      let r;
      try { r = f.apply(this, a); }
      catch (e) { end(); throw e; }
      if (r && typeof r.then === "function") return r.then(v => { end(); return v; }, e => { end(); throw e; });
      end(); return r;
    };
    g.__busy = true;
    set(g);
  }
  function hook(obj, name, txt) {
    if (!obj) return;
    wrap(() => obj[name], v => { obj[name] = v; }, txt);
  }
  return { start, end, progress, hook, wrap };
})();
// 파일 읽기
BUSY.wrap(() => readFile, v => { readFile = v; }, "파일 읽는 중");
// 엑셀 열기·저장 (제일 오래 걸린다)
if (typeof QO !== "undefined") {
  BUSY.hook(QO, "loadWorkbook", "엑셀 읽는 중");
  BUSY.hook(QO, "saveWorkbook", "엑셀 만드는 중");
  BUSY.hook(QO, "convert", "발주서 만드는 중");
  BUSY.hook(QO, "collectInvoices", "송장 맞추는 중");
}
// 드라이브·메일 통신
if (typeof GMAIL !== "undefined") {
  ["driveFetchExcel", "driveListFolder", "driveListShared", "driveSearch", "driveUpdateFile"]
    .forEach(k => BUSY.hook(GMAIL, k, "구글 드라이브 여는 중"));
  ["listMails", "listTextMails", "getAttachment", "searchAddresses", "send"]
    .forEach(k => BUSY.hook(GMAIL, k, "메일 확인하는 중"));
}

/* =====================================================================
   메일 문구 — 발주서·송장·정산·CS 가 같이 쓴다.
   기본 문구를 두되 사용자가 고치면 기억한다. {회사}·{업체} 처럼 중괄호로 값이 들어간다.
   ===================================================================== */
const MAILTPL = (() => {
  const DEF = {
    order:   { name: "발주서", subject: "[{회사}] {날짜}_발주서 송부",
               body: "안녕하세요, {회사}입니다.\n\n발주서 송부드립니다.\n확인 부탁드립니다.\n\n감사합니다.",
               vars: ["회사", "업체", "날짜", "건수"] },
    invoice: { name: "송장 취합본", subject: "[{회사}] {날짜}_송장 취합본 송부",
               body: "안녕하세요, {회사}입니다.\n\n송장 취합본 송부드립니다.\n확인 부탁드립니다.\n\n감사합니다.",
               vars: ["회사", "업체", "날짜"] },
    settle:  { name: "정산서", subject: "[{회사}] {정산월} 정산서 - {업체}",
               body: "안녕하세요, {회사}입니다.\n\n{정산월} 정산 내역을 보내드립니다.\n정산기간 {정산기간} (출고완료된 주문건 기준)\n\n{요약}\n\n자세한 내역은 첨부 파일을 확인해주세요.\n감사합니다.",
               vars: ["회사", "업체", "정산월", "정산기간", "날짜", "건수", "수량", "지급액", "요약"] },
    cs:      { name: "CS 요청", subject: "[{회사}] CS 요청 {건수}건 - {날짜}",
               body: "안녕하세요, {회사}입니다.\n\n아래 CS 건 확인 부탁드립니다. (총 {건수}건)\n",
               vars: ["회사", "업체", "날짜", "건수"] },
  };
  let saved = {};
  async function load() { saved = (await DB.get("mailTemplates", {})) || {}; }
  /* 저장 키 — 공통은 "order", 업체별은 "order@@플라스머".
     업체마다 문구가 달라야 할 때가 있어서, 공통 문구 위에 업체 것을 덮어쓴다. */
  const vkey = (k, v) => k + "@@" + String(v == null ? "" : v).trim();
  const hasVendor = (k, v) => !!(v && saved[vkey(k, v)]);
  const get = (k, v) => Object.assign({}, DEF[k], saved[k] || {}, (v && saved[vkey(k, v)]) || {});
  /* {키} 를 값으로 바꾼다. 모르는 키는 그대로 둔다(지워버리면 왜 빠졌는지 알 수 없다). */
  function fill(s, vars) {
    return String(s == null ? "" : s).replace(/\{([^}\s]+)\}/g,
      (m, k) => (vars && vars[k] !== undefined && vars[k] !== null ? String(vars[k]) : m));
  }
  /* ★ 문구에 회사 이름이 {회사} 가 아니라 글자로 박혀 있는 경우가 있다
     (예전에 저장해 둔 문구, 다른 데서 복사해 온 문구).
     그대로 두면 베타브릭스가 보내는 메일에 '랩노마드' 가 찍혀 나간다 — 사고다.
     그래서 이 배포본의 원래 회사 이름은 로그인한 업체 이름으로 바꿔서 내보낸다. */
  function swapCo(str) {
    const home = String((typeof CONFIG !== "undefined" && CONFIG.homeCompany) || "").trim();
    const co = String((typeof CONFIG !== "undefined" && CONFIG.company) || "").trim();
    if (!home || !co || home === co) return str;
    return String(str == null ? "" : str).split(home).join(co);
  }
  /* vendor 를 주면 그 업체 전용 문구를 먼저 쓴다 (없으면 공통) */
  function render(kind, vars, vendor) {
    const t = get(kind, vendor !== undefined ? vendor : (vars && vars.업체));
    return { subject: swapCo(fill(t.subject, vars)), body: swapCo(fill(t.body, vars)) };
  }

  let cur = null;
  function open(kind, sample, vendor, after, real) {
    const v = (vendor || "").trim();
    cur = { kind, vendor: v, sample: sample || {}, after, real: !!real };
    const d = DEF[kind], t = get(kind, v);
    $("tpl-title").textContent = v ? `${v} — ${d.name} 메일 내용` : `${d.name} 메일 내용 (공통)`;
    $("tpl-sub").textContent = v
      ? `${v} 에게 보낼 때만 이 문구를 씁니다. (되돌리기를 누르면 공통 문구로 돌아갑니다)`
      : "여기서 고친 제목·본문을 앞으로 계속 씁니다.";
    /* ★ 편집창에 {회사} 가 아니라 실제 값을 넣어 보여준다.
       '{회사}입니다' 를 보여주면 무슨 글이 나갈지 가늠이 안 된다 — 그래서 미리보기도 없앴다.
       대신 저장할 때 값을 다시 {키} 로 되돌려, 회사·업체가 바뀌어도 문구가 따라간다. */
    $("tpl-subject").value = fill(swapCo(t.subject), cur.sample);
    $("tpl-body").value = fill(swapCo(t.body), cur.sample);
    $("tpl-vars").textContent = "회사·업체·날짜는 보낼 때 그 값으로 자동으로 바뀝니다"
      + (cur.real ? "" : " (숫자는 예시입니다)");
    $("tplmodal").classList.add("on");
  }
  /* 실제 값 → {키} 되돌리기. 짧은 값(숫자 등)은 엉뚱한 곳이 바뀔 수 있어 3글자 이상만.
     긴 값부터 바꿔야 '랩노마드' 가 '랩노마드물류' 안에서 잘려나가지 않는다. */
  function unfill(str, vars) {
    let out = String(str == null ? "" : str);
    Object.keys(vars || {})
      .filter(k => String(vars[k] == null ? "" : vars[k]).length >= 3)
      .sort((a, b) => String(vars[b]).length - String(vars[a]).length)
      .forEach(k => { out = out.split(String(vars[k])).join("{" + k + "}"); });
    return out;
  }
  function close() { $("tplmodal").classList.remove("on"); cur = null; }

  $("tpl-close").onclick = close;
  $("tplmodal").onclick = e => { if (e.target === $("tplmodal")) close(); };
  /* 되돌리기 — 업체별 창이면 공통 문구로, 공통 창이면 기본 문구로 */
  $("tpl-reset").onclick = () => {
    if (!cur) return;
    const base = cur.vendor ? get(cur.kind) : DEF[cur.kind];
    $("tpl-subject").value = fill(swapCo(base.subject), cur.sample);
    $("tpl-body").value = fill(swapCo(base.body), cur.sample);
  };
  $("tpl-save").onclick = async () => {
    if (!cur) return;
    const val = { subject: unfill($("tpl-subject").value, cur.sample),
                  body: unfill($("tpl-body").value, cur.sample) };
    if (cur.vendor) {
      const common = get(cur.kind);
      // 공통과 똑같아졌으면 업체 전용을 지운다 (안 그러면 공통을 고쳐도 이 업체만 옛 문구가 남는다)
      if (val.subject === common.subject && val.body === common.body) delete saved[vkey(cur.kind, cur.vendor)];
      else saved[vkey(cur.kind, cur.vendor)] = val;
    } else {
      saved[cur.kind] = val;
    }
    const after = cur.after;
    await DB.set("mailTemplates", saved);
    close();
    if (typeof after === "function") after();
  };
  /* 각 탭의 '✏️ 메일내용 수정' 버튼 (공통) */
  document.querySelectorAll("[data-mailtpl]").forEach(b => b.onclick = () => {
    const k = b.dataset.mailtpl;
    open(k, sampleVars(k), b.dataset.mailtplVendor || "");
  });
  function sampleVars(kind) {
    const co = (typeof CONFIG !== "undefined" && CONFIG.company) || "우리회사";
    const base = { 회사: co, 업체: "○○업체", 날짜: QO.fmtDate(QO.todayStr()), 건수: "12", 수량: "18" };
    if (kind === "settle") {
      const p = (window.ST && ST.result() && ST.result().period) || null;
      base.정산월 = (p && p.label) || "2026년 7월";
      base.정산기간 = (window.ST && ST.periodRange && ST.periodRange()) || "2026-07-01 ~ 2026-07-31";
      base.지급액 = "1,234,000원";
      base.요약 = "· 건수: 12건 (수량 18개)\n· 공급가 합계: 1,200,000원\n· 지급액: 1,234,000원 (부가세 포함)";
    }
    return base;
  }
  return { load, get, render, open, hasVendor, sampleVars };
})();

/* 업체별 '메일내용 수정' 버튼 — 발주서·송장취합·정산이 같이 쓴다.
   업체 전용 문구가 저장돼 있으면 라벨로 티를 낸다 (안 그러면 왜 문구가 다른지 알 수 없다). */
function mailtplLabel(kind, vendor) {
  return MAILTPL.hasVendor(kind, vendor) ? "✏️ 메일내용 수정 <b>*</b>" : "✏️ 메일내용 수정";
}
function bindMailtplBtn(btn, kind, vendor, extraVars) {
  if (!btn) return;
  btn.title = "이 업체에게 보낼 메일 제목·본문을 따로 정합니다";
  btn.onclick = () => {
    // 업체 카드에서 열면 그 업체의 진짜 숫자로 미리 보여준다 (견본 금액이면 오해를 산다)
    const sample = Object.assign(MAILTPL.sampleVars(kind), { 업체: vendor }, extraVars || {});
    MAILTPL.open(kind, sample, vendor, () => { btn.innerHTML = mailtplLabel(kind, vendor); }, true);
  };
}

/* 버전 표기 — qo-version.js 한 곳에서 읽어 배지와 푸터에 같이 넣는다.
   손으로 여러 곳을 고치다 하나를 빠뜨리는 일을 막으려고 이렇게 해뒀다. */
(function showVersion() {
  const v = (typeof APP_VER !== "undefined" && APP_VER) ? "v" + APP_VER : "";
  if (!v) return;
  const a = $("app-ver"), b = $("app-ver-foot");
  if (a) a.textContent = v;
  if (b) b.textContent = v;
})();

/* ---------------- 탭 ---------------- */
const TABS = ["o", "i", "c", "s"];
TABS.forEach(t => { const b = $("tab-" + t); if (b) b.onclick = () => switchTab(t); });
/* ③CS·교환반품 / ④정산 — 실제 데이터로 검증하기 전까지 화면에서 감춘다.
   켤 때는 아래 값을 true 로만 바꾸면 된다 (코드는 그대로 살아 있음). */
const SHOW_CS = false;      // ③ CS·교환반품 — 아직 실제 데이터로 검증 안 함
const SHOW_SETTLE = true;   // ④ 정산 — 2026-08-02 실제 7월 데이터(768행)로 검증하고 켬
if (!SHOW_CS) { const b = $("tab-c"); if (b) b.style.display = "none"; }
if (!SHOW_SETTLE) { const b = $("tab-s"); if (b) b.style.display = "none"; }
/* 감춘 탭을 빼고 번호를 다시 매긴다 — 안 하면 ① ② ④ 처럼 건너뛴 채로 보인다.
   라벨 안의 배지(<span class="dot">)는 건드리지 않도록 첫 텍스트 노드만 고친다. */
(function renumberTabs() {
  const NUM = ["①", "②", "③", "④", "⑤"];
  let n = 0;
  TABS.forEach(t => {
    const b = $("tab-" + t);
    if (!b || b.style.display === "none") return;
    const node = b.firstChild;
    if (node && node.nodeType === 3 && n < NUM.length)
      node.nodeValue = node.nodeValue.replace(/^\s*[①②③④⑤]?\s*/, NUM[n] + " ");
    n++;
  });
})();
function switchTab(t) {
  TABS.forEach(x => {
    const p = $("pane-" + x), b = $("tab-" + x);
    if (p) p.classList.toggle("on", x === t);
    if (b) b.classList.toggle("on", x === t);
  });
  window.scrollTo(0, 0);
  // 탭이 처음 열릴 때 해당 모듈이 목록을 그린다 (모듈이 없으면 무시)
  try { if (t === "c" && window.CS) CS.onShow(); } catch (e) {}
  try { if (t === "s" && window.ST) ST.onShow(); } catch (e) {}
}

/* =================================================================
   ① 발주서 변환
   ================================================================= */
$("f-order").addEventListener("change", function () { if (this.files[0]) setOrder(this.files[0]); });
bindDrop("drop-order", f => setOrder(f[0]));

async function setOrder(file) {
  msg("msg-o", "", "");
  try {
    S.orderBuf = await readFile(file);
    S.orderName = file.name;
    $("order-name").textContent = "📄 " + file.name;
    $("drop-order").classList.add("on");
    const wb = await QO.loadWorkbook(S.orderBuf.slice(0));
    S.brands = QO.listBrands(wb);
    S.dateSel = [];
    await loadDates();
    await drawPreview();
    buildVendorBrands();
    refreshO();
  } catch (e) { msg("msg-o", "err", "파일을 읽지 못했어요: " + e.message); }
}

/* =================================================================
   구글 드라이브 파일 선택기 (폴더 탐색 + 검색 + 링크) — 발주서/업체양식/송장양식/회신 공용
   ================================================================= */
const DRV = { multiple: false, onPick: null, path: [], sel: new Map(), home: null, files: [] };
const GSHEET = "application/vnd.google-apps.spreadsheet";
const GFOLDER = "application/vnd.google-apps.folder";

$("drv-close").onclick = () => $("drvmodal").classList.remove("on");
$("drvmodal").onclick = e => { if (e.target === $("drvmodal")) $("drvmodal").classList.remove("on"); };
$("drv-search").onclick = () => drvSearch($("drv-q").value);
$("drv-q").addEventListener("keydown", e => { if (e.key === "Enter") drvSearch($("drv-q").value); });
$("drv-link-go").onclick = async () => {
  const id = GMAIL.driveIdFromLink($("drv-link").value);
  if (!id) { $("drv-msg").textContent = "⚠ 링크에서 파일 ID를 못 찾았어요. 드라이브 공유 링크를 그대로 붙여넣어 주세요."; return; }
  try {
    const info = await GMAIL.driveFileInfo(id);
    await drvPick([{ id, name: info.name, mimeType: info.mimeType }]);
  } catch (e) { $("drv-msg").textContent = "⚠ " + e.message; }
};
/* 아무것도 안 고르면 '이 폴더 전부', 고른 게 있으면 '그것만'.
   버튼 글자가 무슨 일이 일어날지 그대로 말해주게 한다 (예전엔 빈 선택이면 조용히 아무 일도 안 했다). */
function drvDoneLabel() {
  const d = $("drv-done"); if (!d || !DRV.multiple) return;
  const n = (DRV.files || []).length;
  d.textContent = DRV.sel.size ? `선택 완료 (${DRV.sel.size}개)`
    : n ? `전체 가져오기 (${n}개)` : "선택 완료";
  d.style.opacity = (DRV.sel.size || n) ? "" : ".55";
}
$("drv-done").onclick = async () => {
  if (DRV.sel.size) return drvPick([...DRV.sel.values()]);
  const fs = (DRV.files || []).map(f => ({ id: f.id, name: f.name, mimeType: f.mimeType }));
  if (!fs.length) { $("drv-msg").textContent = "⚠ 이 폴더에는 가져올 엑셀 파일이 없어요."; return; }
  if (fs.length > 1 && !confirm(`이 폴더의 파일 ${fs.length}개를 모두 가져올까요?`)) return;
  await drvPick(fs);
};


/* opts: { key, title, sub, multiple, onPick(files) } — files: [{id,name,mimeType}]
   key: 용도별 기본 폴더 저장용 (order/tpl/sab/rep) → 다음부터 그 폴더가 바로 열림 */
async function openDrivePicker(opts) {
  DRV.multiple = !!opts.multiple; DRV.onPick = opts.onPick; DRV.sel = new Map();
  DRV.key = opts.key || "";
  $("drv-title").textContent = opts.title || "구글 드라이브에서 가져오기";
  /* 어느 구글 계정의 드라이브를 보고 있는지 함께 적는다.
     이게 안 보여서 '베타브릭스로 로그인했는데 랩노마드 드라이브가 보인다' 를
     한참 뒤에야 알아챘다. */
  const acct = CONFIG.account ? ` · 👤 ${CONFIG.account}` : "";
  $("drv-sub").textContent = (opts.sub || (opts.multiple
    ? "폴더 안에서 파일을 여러 개 고를 수 있어요." : "폴더 안에서 파일을 고르세요.")) + acct;
  $("drv-done").style.display = opts.multiple ? "" : "none";
  $("drv-done").textContent = "선택 완료";
  $("drv-msg").textContent = ""; $("drv-q").value = ""; $("drv-link").value = "";
  $("drv-list").innerHTML = "";
  $("drvmodal").classList.add("on");
  // 로그인 안 돼 있으면: 조용한 갱신(팝업 없음) 시도 → 실패하면 '로그인' 버튼을 보여준다.
  // (자동으로 팝업을 띄우면 브라우저가 "Failed to open popup window"로 막아버림)
  if (GMAIL.needLogin()) { drvNeedLogin(); return; }   // 자동 팝업 금지 → 버튼으로 유도
  drvStart();
}
/* 로그인 필요 화면 — 버튼을 '직접 클릭'해야 팝업이 안 막힘 */
function drvNeedLogin() {
  const box = $("drv-list"); box.innerHTML = "";
  const d = document.createElement("div"); d.className = "empty";
  d.innerHTML = "구글 드라이브를 보려면 로그인이 필요합니다.<br>아래 버튼을 눌러주세요.<br><br>";
  const b = document.createElement("button");
  b.className = "btn-go"; b.style.cssText = "width:auto;padding:11px 20px;font-size:14px";
  b.textContent = "구글 로그인";
  b.onclick = () => {                    // 직접 클릭 → 팝업 차단 안 됨(중간에 await 없음)
    $("drv-msg").textContent = "로그인 창을 여는 중…";
    GMAIL.signIn("select_account")
      .then(() => { $("drv-msg").textContent = ""; updateGmailWho(); drvStart(); })
      .catch(e => {
        const m = e.message || "";
        $("drv-msg").textContent = /popup/i.test(m)
          ? "⚠ 브라우저가 로그인 창(팝업)을 막았습니다.\n주소창 오른쪽의 '팝업 차단됨' 아이콘을 눌러 이 사이트의 팝업을 허용한 뒤 다시 눌러주세요."
          : "⚠ " + m;
      });
  };
  d.appendChild(b); box.appendChild(d);
}
/* 최상위 폴더 — 회사 작업 폴더가 '내 드라이브' 한참 안쪽에 있어서,
   매번 위에서부터 파고들지 않게 여기를 바닥으로 삼는다.
   여기가 정해지면 경로 표시가 이 폴더에서 시작하고 '상위' 버튼도 그 위로 안 올라간다.
   ★ 버튼으로 지정하지 않는다 — 파일을 처음 고르면 '내 드라이브 바로 아래' 폴더가
     자동으로 최상위가 된다. 탭(발주·송장·정산·CS)이 전부 같은 값을 쓴다. */
const DRV_ROOT = { id: "root", name: "내 드라이브" };
const drvBase = () => (DRV.home && DRV.home.id ? { id: DRV.home.id, name: DRV.home.name } : Object.assign({}, DRV_ROOT));
/* 최상위가 정해져 있으면 조상 사슬을 거기서 잘라낸다 (최상위 밖의 폴더면 그대로 둔다) */
function drvTrim(chain) {
  const path = [Object.assign({}, DRV_ROOT)].concat(chain || []);
  if (!(DRV.home && DRV.home.id)) return path;
  const i = path.findIndex(p => p.id === DRV.home.id);
  return i >= 0 ? path.slice(i) : path;
}
/* 시작 위치: ①마지막에 파일 고른 폴더 → ②최상위 폴더 → ③내 드라이브 */
async function drvStart() {
  const all = await DB.get("driveFolders", {});
  DRV.home = all[":home"] || null;
  const last = DRV.key ? all[DRV.key + ":last"] : null;
  const go = (last && last.id) ? last : (DRV.home && DRV.home.id ? DRV.home : null);
  if (go) {
    // 실제 드라이브 상위 폴더들을 따라 경로를 만든다 → '상위' 버튼이 제대로 동작
    DRV.path = [drvBase()];
    if (go.id !== DRV.path[0].id) DRV.path.push({ id: go.id, name: go.name });
    drvOpen(go.id);
    GMAIL.driveAncestors(go.id).then(chain => {
      if (chain && chain.length) {
        DRV.path = drvTrim(chain);
        drvCrumb();
        $("drv-up").style.display = DRV.path.length > 1 ? "" : "none";
      }
    }).catch(() => {});
  } else {
    DRV.path = [Object.assign({}, DRV_ROOT)];
    drvOpen("root");
  }
}
function drvCrumb() {
  const c = $("drv-crumb"); c.innerHTML = "";
  DRV.path.forEach((p, i) => {
    if (i) c.appendChild(document.createTextNode(" › "));
    const b = document.createElement("b"); b.textContent = p.name;
    b.onclick = () => { DRV.path = DRV.path.slice(0, i + 1); drvOpen(p.id, true); };
    c.appendChild(b);
  });
}
$("drv-up").onclick = () => {
  if (DRV.path.length <= 1) return;
  DRV.path.pop();
  const parent = DRV.path[DRV.path.length - 1];
  drvOpen(parent.id);
};
/* 드라이브 화면과 같은 순서: 폴더 먼저 → 이름순 (숫자 (1),(2)… 자연스럽게) */
const drvSort = arr => arr.slice().sort((a, b) => {
  const fa = a.mimeType === GFOLDER, fb = b.mimeType === GFOLDER;
  if (fa !== fb) return fa ? -1 : 1;
  return String(a.name || "").localeCompare(String(b.name || ""), "ko", { numeric: true, sensitivity: "base" });
});
async function drvOpen(folderId) {
  drvCrumb();
  $("drv-up").style.display = DRV.path.length > 1 ? "" : "none";
  const box = $("drv-list"); box.innerHTML = '<div class="empty">불러오는 중…</div>';
  try {
    let items;
    if (folderId === "shared") items = drvSort(await GMAIL.driveListShared(200));
    else {
      items = drvSort(await GMAIL.driveListFolder(folderId, 200));
      // 최상위엔 '공유 문서함'(남이 공유해준 폴더)도 맨 위에
      if (folderId === "root") items = [{ id: "shared", name: "공유 문서함", mimeType: GFOLDER }].concat(items);
    }
    drvRender(items, true);
  } catch (e) {
    if (/popup|로그인|권한|401|403|NEED_LOGIN/i.test(e.message || "")) { drvNeedLogin(); $("drv-msg").textContent = ""; }
    else { box.innerHTML = ""; $("drv-msg").textContent = "⚠ " + e.message; }
  }
}
async function drvSearch(q) {
  if (!q || !q.trim()) { DRV.path = [drvBase()]; return drvOpen(DRV.path[0].id); }
  const box = $("drv-list"); box.innerHTML = '<div class="empty">찾는 중…</div>';
  $("drv-crumb").innerHTML = `<span>검색: “${esc(q)}”</span>`;
  try { drvRender(await GMAIL.driveSearch(q, 50), false); }
  catch (e) { box.innerHTML = ""; $("drv-msg").textContent = "⚠ " + e.message; }
}
function drvRender(items, allowFolders) {
  const box = $("drv-list"); box.innerHTML = "";
  // 지금 화면의 파일(폴더 제외). 아무것도 안 고르고 눌렀을 때 통째로 가져오는 데 쓴다.
  DRV.files = items.filter(f => f.mimeType !== GFOLDER);
  drvDoneLabel();
  if (!items.length) { box.innerHTML = '<div class="empty">엑셀/시트 파일이 없어요.</div>'; return; }
  items.forEach(f => {
    const folder = f.mimeType === GFOLDER;
    if (folder && !allowFolders) return;
    const el = document.createElement("div");
    el.className = "drvrow" + (DRV.sel.has(f.id) ? " on" : "");
    const when = f.modifiedTime ? new Date(f.modifiedTime).toLocaleString("ko-KR",
      { month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" }) : "";
    el.innerHTML = `<span class="ic">${folder ? "📁" : (f.mimeType === GSHEET ? "📊" : "📄")}</span>
      <span class="nm"><b>${esc(f.name)}</b><span>${folder ? "폴더" : "수정 " + esc(when)}</span></span>`;
    el.onclick = () => {
      if (folder) { DRV.path.push({ id: f.id, name: f.name }); return drvOpen(f.id); }
      if (DRV.multiple) {
        if (DRV.sel.has(f.id)) DRV.sel.delete(f.id); else DRV.sel.set(f.id, f);
        el.classList.toggle("on", DRV.sel.has(f.id));
        drvDoneLabel();
        $("drv-msg").textContent = "";
      } else drvPick([f]);
    };
    box.appendChild(el);
  });
}
async function drvPick(files) {
  $("drv-msg").textContent = "가져오는 중…";
  try {
    if (DRV.onPick) await DRV.onPick(files);
    await drvRememberFolder();          // 실제로 파일을 고른 폴더만 기억한다
    $("drvmodal").classList.remove("on");
  } catch (e) { $("drv-msg").textContent = "⚠ " + e.message; }
}
/* 파일을 고른 그 폴더를 기억 → 다음에 열면 거기서 시작.
   (둘러보기만 한 폴더까지 기억하면, 잠깐 다른 데를 봤을 뿐인데 시작 위치가 바뀐다)
   같이 '최상위 폴더'도 정한다 — 지금 경로에서 내 드라이브 바로 아래 폴더.
   버튼을 눌러 지정할 필요 없이, 한 번 파일을 고르면 다음부터 거기가 바닥이 된다. */
async function drvRememberFolder() {
  const cur = DRV.path[DRV.path.length - 1];
  if (!cur || cur.id === "root" || cur.id === "shared") return;
  const all = await DB.get("driveFolders", {});
  if (DRV.key) all[DRV.key + ":last"] = { id: cur.id, name: cur.name };
  if (!(all[":home"] && all[":home"].id)) {
    // path[0] 이 '내 드라이브'면 그 다음이 작업 폴더. 이미 최상위 안에서 시작했으면 path[0] 이 곧 최상위다.
    const top = DRV.path[0] && DRV.path[0].id === "root" ? DRV.path[1] : DRV.path[0];
    if (top && top.id && top.id !== "root" && top.id !== "shared") {
      all[":home"] = { id: top.id, name: top.name };
      DRV.home = all[":home"];
    }
  }
  await DB.set("driveFolders", all);
}

/* --- ① 쇼핑몰 주문 파일 --- */
// 폴더 탐색으로 골라서 바로 변환 (고른 파일은 '바로 가져오기' 대상으로도 자동 지정)
// ※ 별도 버튼은 없앴다 — '지정 파일 · 다른 파일로 변경' 으로 같은 일을 한다.
//    설정 등 다른 데서 부를 수 있게 함수는 남겨 둔다.
function pickOrderFromDrive() {
  return openDrivePicker({
    key: "order", title: "드라이브에서 발주서 가져오기", multiple: false,
    onPick: async files => {
      const f = files[0];
      const r = await GMAIL.driveFetchExcel(f.id);
      await DB.set("driveOrderFile", { id: f.id, name: r.name });
      await setOrderFromBuf(r.buf, r.name);
      drawDriveRecent();
      msg("msg-o", "ok", `✔ 드라이브에서 가져왔어요: ${r.name}`);
    },
  });
}
if ($("drive-order")) $("drive-order").onclick = pickOrderFromDrive;
/* '바로 가져올 파일' 지정/변경.
   ※ '변경'은 지정만 바꾸고 끝내면 안 된다 — 지정은 8월인데 아래 불러온 파일은 7월인
     상태가 되어 헷갈린다. 바꿨으면 그 파일을 바로 불러온다. */
function pickOrderPin(alsoFetch) {
  openDrivePicker({
    key: "order", title: alsoFetch ? "다른 발주서 파일로 변경" : "바로 가져올 발주서 파일 지정", multiple: false,
    onPick: async files => {
      const f = files[0];
      await DB.set("driveOrderFile", { id: f.id, name: f.name });
      drawDriveRecent();
      if (!alsoFetch) { msg("msg-o", "ok", `✔ 바로 가져올 발주서로 지정했어요: ${f.name}`); return; }
      msg("msg-o", "", "");
      try {
        await ensureGmail();
        const r = await GMAIL.driveFetchExcel(f.id);
        await setOrderFromBuf(r.buf, r.name);
        msg("msg-o", "ok", `✔ ${r.name} (으)로 바꿔서 가져왔어요 · 아래에서 날짜 고르고 변환하세요`);
      } catch (e) { msg("msg-o", "err", "가져오기 실패: " + e.message); }
    },
  });
}
$("order-pin-set").onclick = () => pickOrderPin(false);
$("order-pin-change").onclick = () => pickOrderPin(true);
$("order-pin-clear").onclick = async () => {
  await DB.set("driveOrderFile", null);
  drawDriveRecent();
  msg("msg-o", "ok", "지정을 해제했어요.");
};
// 지정한 발주서를 한 번에(폴더 탐색 없이) 최신본으로 가져와 바로 변환 프로세스로
$("drive-again").onclick = async function () {
  const f = await DB.get("driveOrderFile", null);
  if (!f || !f.id) { pickOrderPin(true); return; }   // 지정된 게 없으면 드라이브에서 고르게
  this.disabled = true; const orig = this.innerHTML; this.textContent = "가져오는 중…";
  msg("msg-o", "", "");
  try {
    await ensureGmail();
    const r = await GMAIL.driveFetchExcel(f.id);
    await setOrderFromBuf(r.buf, r.name);          // ← 이후는 업로드와 동일한 변환 프로세스
    msg("msg-o", "ok", `✔ 최신본을 가져왔어요: ${r.name} · 아래에서 날짜 고르고 변환하세요`);
  } catch (e) { msg("msg-o", "err", "가져오기 실패: " + e.message); }
  finally { this.disabled = false; this.innerHTML = orig; }
};
async function drawDriveRecent() {
  const f = await DB.get("driveOrderFile", null);
  const has = !!(f && f.id);
  // 퀵로딩 버튼은 항상 보인다. 지정 파일이 없으면 누를 때 드라이브에서 고르게 한다
  // (지정 전에는 숨겨져 있어서 드라이브로 갈 길이 아예 없었다)
  $("drive-again").style.display = "block";
  $("order-pinrow").style.display = has ? "flex" : "none";
  $("order-pin-set").style.display = "none";
  if (has) { $("order-pinname").textContent = f.name; }
}

/* --- ② 업체 양식 (여러 개 선택 가능) --- */
$("drive-tpl").onclick = () => openDrivePicker({
  key: "tpl", title: "드라이브에서 발주서 양식 가져오기", multiple: true,
  onPick: async files => {
    const list = [];
    for (const f of files) {
      const g = await GMAIL.driveFetchExcel(f.id);
      list.push({ fileName: g.name, buf: g.buf });
    }
    const named = await askVendorNames(list);   // 업체명 확인 후 저장
    if (!named) return;
    const r = await saveForms(named);
    await loadForms();
    msg("msg-o", "ok", `✔ 드라이브에서 발주서 양식 ${r.added + r.updated}개 저장 — ${r.names.join(", ")}`
      + (r.updated ? ` (${r.updated}개는 기존 양식 갱신)` : ""));
  },
});

/* --- ③ 송장취합양식 (하나) --- */
// 드라이브에서 송장취합양식을 받아 화면에 세팅 (되쓰기 위해 출처 기억)
async function loadSabFromDrive(fileId) {
  const r = await GMAIL.driveFetchExcel(fileId);
  S.sabBuf = r.buf; S.sabName = r.name;
  S.sabDrive = { id: fileId, name: r.name };   // 드라이브 출처 기억 → 결과를 이 파일에 되쓰기
  $("sab-name").textContent = "📁 " + r.name + " (드라이브)";
  $("drop-sab").classList.add("on"); $("sab-preview").style.display = "block";
  refreshI();
  return r;
}
// 폴더 탐색으로 골라 가져오기 (고른 파일은 '바로 가져오기' 대상으로도 자동 지정)
// ※ 별도 버튼은 없앴다 — '지정 파일 · 다른 파일로 변경' 으로 같은 일을 한다.
function pickSabFromDrive() {
  return openDrivePicker({
    key: "sab", title: "드라이브에서 송장취합양식 가져오기", multiple: false,
    onPick: async files => {
      const r = await loadSabFromDrive(files[0].id);
      await DB.set("driveSabFile", { id: files[0].id, name: r.name });
      drawSabRecent();
      msg("msg-i", "ok", `✔ 드라이브에서 송장취합양식을 가져왔어요: ${r.name}`);
    },
  });
}
if ($("drive-sab")) $("drive-sab").onclick = pickSabFromDrive;
// '바로 가져올 파일' 지정/변경 — 가져오지 않고 대상만 지정
function pickSabPin(alsoFetch) {
  openDrivePicker({
    key: "sab", title: alsoFetch ? "다른 송장취합양식으로 변경" : "바로 가져올 송장취합양식 파일 지정", multiple: false,
    onPick: async files => {
      const f = files[0];
      await DB.set("driveSabFile", { id: f.id, name: f.name });
      drawSabRecent();
      if (!alsoFetch) { msg("msg-i", "ok", `✔ 바로 가져올 송장취합양식으로 지정했어요: ${f.name}`); return; }
      // 바꿨으면 바로 불러온다 (지정만 바뀌고 화면은 이전 파일인 상태를 막는다)
      msg("msg-i", "", "");
      try {
        await ensureGmail();
        const r = await loadSabFromDrive(f.id);
        msg("msg-i", "ok", `✔ ${r.name} (으)로 바꿔서 가져왔어요`);
      } catch (e) { msg("msg-i", "err", "가져오기 실패: " + e.message); }
    },
  });
}
$("sab-pin-set").onclick = () => pickSabPin(false);
$("sab-pin-change").onclick = () => pickSabPin(true);
$("sab-pin-clear").onclick = async () => {
  await DB.set("driveSabFile", null);
  drawSabRecent();
  msg("msg-i", "ok", "지정을 해제했어요.");
};
// 지정한 송장취합양식을 한 번에 최신본으로 가져오기
$("sab-again").onclick = async function () {
  const f = await DB.get("driveSabFile", null);
  if (!f || !f.id) { pickSabPin(true); return; }      // 지정된 게 없으면 드라이브에서 고르게
  this.disabled = true; const orig = this.innerHTML; this.textContent = "가져오는 중…";
  msg("msg-i", "", "");
  try {
    await ensureGmail();
    const r = await loadSabFromDrive(f.id);
    msg("msg-i", "ok", `✔ 최신본을 가져왔어요: ${r.name}`);
  } catch (e) { msg("msg-i", "err", "가져오기 실패: " + e.message); }
  finally { this.disabled = false; this.innerHTML = orig; }
};
async function drawSabRecent() {
  const f = await DB.get("driveSabFile", null);
  const has = !!(f && f.id);
  $("sab-again").style.display = "block";      // 항상 보인다 (위 drawDriveRecent 주석 참고)
  $("sab-pinrow").style.display = has ? "flex" : "none";
  $("sab-pin-set").style.display = "none";
  if (has) { $("sab-pinname").textContent = f.name; }
}

/* --- ④ 업체 회신 송장 (여러 개 선택 가능) --- */
$("drive-rep").onclick = () => openDrivePicker({
  key: "rep", title: "드라이브에서 회신 송장 가져오기", multiple: true,
  onPick: async files => {
    let n = 0;
    for (const f of files) {
      const r = await GMAIL.driveFetchExcel(f.id);
      if (S.reps.some(x => x.name === r.name)) continue;
      S.reps.push({ name: r.name, data: r.buf }); n++;
    }
    drawReps(); refreshI();
    msg("msg-i", "ok", `✔ 드라이브에서 회신 송장 ${n}개를 가져왔어요.`);
  },
});

/* --- 날짜 --- */
$("dt-col").onchange = function () { loadDates(this.value).then(drawPreview); };
async function loadDates(header) {
  $("date-wrap").style.display = "block";
  $("dt-info").textContent = "· 읽는 중…";
  const wb = await QO.loadWorkbook(S.orderBuf.slice(0));
  const di = QO.orderDateInfo(wb, header);
  S.dateHeader = di.header;
  const sel = $("dt-col");
  if (di.candidates.length) {
    sel.style.display = "";
    sel.innerHTML = di.candidates.map(c => `<option value="${esc(c)}"${c === di.header ? " selected" : ""}>기준: ${esc(c)}</option>`).join("");
  } else sel.style.display = "none";

  const list = Object.entries(di.counts).sort((a, b) => b[0].localeCompare(a[0]))
    .map(([d, n]) => ({ date: d, label: QO.fmtDate(d), count: n }));
  S.dateAll = list;
  const box = $("dt-chips"); box.innerHTML = "";
  if (!list.length) {
    $("dt-info").textContent = "";
    $("dt-foot").textContent = "날짜 열이 없어 전체가 변환됩니다";
    S.dateSel = []; refreshO(); return;
  }
  $("dt-info").textContent = `· ${esc(di.header || "")} 기준 · 여러 날짜 선택 가능`;
  const valid = S.dateSel.filter(d => list.some(x => x.date === d));
  S.dateSel = valid.length ? valid : [list[0].date];

  const all = document.createElement("span");
  all.className = "brow"; all.style.borderStyle = "dashed";
  all.onclick = () => { S.dateSel = S.dateSel.length === list.length ? [] : list.map(d => d.date); drawDateChips(); };
  box.appendChild(all);
  list.forEach((d, i) => {
    const el = document.createElement("span");
    el.className = "brow" + (i >= DT_SHOW ? " dt-more" : "");   // 최근 3개만 기본 표시
    el.dataset.d = d.date;
    el.innerHTML = `<span class="box">${CHK}</span>${esc(String(d.label).slice(5))} (${d.count})`;
    el.onclick = () => {
      const j = S.dateSel.indexOf(d.date);
      if (j >= 0) S.dateSel.splice(j, 1); else S.dateSel.push(d.date);
      drawDateChips();
    };
    box.appendChild(el);
  });
  // 더보기/접기 (날짜가 많을 때 한 줄만 보이게)
  S.dtOpen = false;
  if (list.length > DT_SHOW) {
    const more = document.createElement("button");
    more.className = "minibtn"; more.id = "dt-more-btn";
    more.onclick = () => {
      S.dtOpen = !S.dtOpen;
      box.classList.toggle("open", S.dtOpen);        // 펼치면 여러 줄, 접으면 한 줄
      box.querySelectorAll(".dt-more").forEach(x => x.classList.toggle("show", S.dtOpen));
      more.textContent = S.dtOpen ? "접기" : `+ 더보기 (${list.length - DT_SHOW}개)`;
    };
    more.textContent = `+ 더보기 (${list.length - DT_SHOW}개)`;
    box.appendChild(more);
  }
  drawDateChips();
}
function drawDateChips() {
  const box = $("dt-chips");
  box.querySelectorAll(".brow[data-d]").forEach(el => el.classList.toggle("on", S.dateSel.includes(el.dataset.d)));
  const all = box.querySelector(".brow:not([data-d])");
  const totalAll = S.dateAll.reduce((s, d) => s + d.count, 0);
  if (all) all.textContent = (S.dateSel.length === S.dateAll.length && S.dateAll.length) ? "전체 해제" : `전체 ${totalAll}건`;
  const cnt = S.dateAll.filter(d => S.dateSel.includes(d.date)).reduce((s, d) => s + d.count, 0);
  // 건수 안내는 뺐다 (2026-08-05) — 날짜 칩에 이미 건수가 있다.
  // 날짜를 안 고른 경고만 남긴다 (이건 안 보이면 변환이 안 되는 이유를 모른다).
  $("dt-foot").textContent = S.dateSel.length ? "" : "⚠ 날짜를 하나 이상 선택하세요";
  renderPreview();     // 체크한 날짜에 맞춰 '내용 확인'도 즉시 갱신
  refreshO();
}

/* --- 내용 확인 --- */
$("pv-more").onclick = function () {
  const sc = $("pv-scroll"), open = sc.classList.toggle("collapsed") === false;
  this.textContent = open ? "접기" : "+ 더보기";
};
$("pv-toggle").onclick = function () { S.pvAll = !S.pvAll; this.textContent = S.pvAll ? "주요 열만 보기" : "전체 열 보기"; renderPreview(); };
async function drawPreview() {
  $("prev-wrap").style.display = "block";
  $("pv-cnt").textContent = "· 읽는 중…";
  const wb = await QO.loadWorkbook(S.orderBuf.slice(0));
  S.pv = QO.preview(wb, 5000, { dateHeader: S.dateHeader });   // 전체를 읽되, 화면에선 체크한 날짜만 표시
  S.pvAll = false;
  $("pv-toggle").textContent = "전체 열 보기";
  renderPreview();
}
function renderPreview() {
  const pv = S.pv; if (!pv) return;
  const idx = (S.pvAll || !pv.keyIdx.length) ? pv.columns.map((_, i) => i) : pv.keyIdx;
  // 체크한 수집일자에 해당하는 행만 표시 (선택이 없으면 전부)
  const selSet = S.dateSel && S.dateSel.length ? new Set(S.dateSel) : null;
  const hasDates = pv.rowDates && pv.rowDates.some(x => x);
  const view = (selSet && hasDates) ? pv.rows.filter((_, i) => selSet.has(pv.rowDates[i])) : pv.rows;
  $("pv-cnt").textContent = (selSet && hasDates)
    ? `· 선택한 ${S.dateSel.length}개 날짜 ${view.length}건 (파일 전체 ${pv.total}건) · 열 ${idx.length}/${pv.columns.length}`
    : `· 전체 ${view.length}건 · 열 ${idx.length}/${pv.columns.length}`;
  let h = "<tr>" + idx.map(i => `<th>${esc(pv.columns[i] || "열" + (i + 1))}</th>`).join("") + "</tr>";
  view.forEach(row => {
    h += "<tr>" + idx.map(i => {
      const v = row[i] == null ? "" : row[i];
      const num = /^[0-9,.\-]+$/.test(v) && v !== "";
      return `<td${num ? ' class="num"' : ""}>${esc(v)}</td>`;
    }).join("") + "</tr>";
  });
  $("pv-table").innerHTML = h;
  $("pv-foot").textContent = "";
  // 행이 몇 개 안 되면 '더보기' 자체를 숨김
  const many = view.length > 3;
  $("pv-more").style.display = many ? "" : "none";
  if (!many) $("pv-scroll").classList.remove("collapsed");
  $("pv-more").textContent = $("pv-scroll").classList.contains("collapsed") ? "+ 더보기" : "접기";
}

/* --- 업체 양식 --- */
$("f-tpl").addEventListener("change", function () { const fs = [...this.files]; this.value = ""; addForms(fs); });
bindDrop("drop-tpl", f => addForms(f));
/* 파일명에서 업체명 후보를 순서대로 뽑는다. 앞의 것이 이미 쓰이고 있으면 다음 것으로 넘어간다.
     "랩노마드 발주양식(디에스피).xlsx" → ①디에스피 ②랩노마드 ③디에스피 ④랩노마드 발주양식(디에스피)
     "디에스피_발주양식.xlsx"          → ①디에스피_발주양식 ②발주양식 …
   ※ 이미 저장돼 있는 양식의 이름은 바뀌지 않는다(같은 파일명이면 기존 이름을 그대로 씀). */
function candidateNames(fileName) {
  const stem = String(fileName).replace(/\.[^.]+$/, "").replace(/^.*[\\/]/, "").trim();
  const out = [];
  const push = v => {
    v = String(v || "").replace(/[\\/:*?"<>|()[\]【】]/g, "").trim();
    if (v && !out.includes(v)) out.push(v);
  };
  // ① 끝에 괄호가 있으면 그 안이 업체명 — "랩노마드 발주양식(디에스피)"
  const tail = stem.match(/[([【]([^)\]】]+)[)\]】]\s*$/);
  if (tail) push(tail[1]);

  // ② 군더더기를 걷어내고 남는 첫 단어 — "(랩노마드) DSP 발주서_2607" → "DSP"
  //    앞쪽 괄호(자사명)·문서 종류어(발주서/양식)·날짜숫자를 지운다.
  const core = stem
    .replace(/^\s*[([【][^)\]】]*[)\]】]\s*/, "")        // 맨 앞 괄호 묶음
    .replace(/발주\s*양식|발주서|발주\s*의뢰|주문서|발주|양식|form|order/gi, " ")
    .replace(/\d{4,8}/g, " ")                          // 2607 · 20260728 같은 날짜
    .replace(/[_\-.]+/g, " ")
    .replace(/\s+/g, " ").trim();
  if (core) { push(core.split(" ")[0]); push(core); }

  push(QO.nameFromFilename(fileName));                 // ③ 예전 규칙(첫 단어)
  push(stem);                                          // ④ 파일명 전체
  return out;
}

/* 업체명 입력 모달 — 파일마다 이름을 확인/수정한 뒤 저장한다.
   list = [{fileName, buf}] → 확정된 [{fileName, buf, name}] 로 resolve. 취소하면 null. */
function askVendorNames(list) {
  return new Promise(async resolve => {
    const existing = await DB.listForms();
    const box = $("vn-list");
    box.innerHTML = "";
    $("vn-msg").textContent = "";
    $("vn-title").textContent = list.length > 1 ? `업체명 정하기 (${list.length}개)` : "업체명 정하기";
    const used = new Set(existing.map(f => f.name));
    list.forEach(it => {
      // 같은 파일을 다시 올린 것이면 이전에 쓰던 이름을 그대로 보여준다
      const same = existing.find(f => f.file === it.fileName);
      let pre = same ? same.name : "";
      if (!pre) {
        const cands = candidateNames(it.fileName);
        pre = cands.find(c => !used.has(c)) || cands[0] || "";
      }
      used.add(pre);
      const row = document.createElement("div");
      row.className = "mitem";
      row.innerHTML = `<div class="vnfile">📄 ${esc(it.fileName)}</div>
        <input class="vnin" value="${esc(pre)}" placeholder="업체명 (예: 디에스피)"
               autocapitalize="off" autocorrect="off" spellcheck="false">`;
      box.appendChild(row);
    });
    $("vnmodal").classList.add("on");
    setTimeout(() => { const i = box.querySelector(".vnin"); if (i) { i.focus(); i.select(); } }, 80);

    const close = () => { $("vnmodal").classList.remove("on"); cleanup(); };
    const onCancel = () => { close(); resolve(null); };
    const onOk = () => {
      const ins = [...box.querySelectorAll(".vnin")];
      const names = ins.map(i => i.value.trim().replace(/[\\/:*?"<>|]/g, ""));
      if (names.some(n => !n)) { $("vn-msg").textContent = "⚠ 업체명을 모두 입력해주세요."; return; }
      const dup = names.find((n, i) => names.indexOf(n) !== i);
      if (dup) { $("vn-msg").textContent = `⚠ '${dup}' 이(가) 중복됩니다. 서로 다르게 적어주세요.`; return; }
      // 기존에 있는 이름이면 그 양식을 갱신하게 되므로 확인
      const clash = names.find((n, i) => {
        const ex = existing.find(f => f.name === n);
        return ex && ex.file !== list[i].fileName;
      });
      if (clash && !confirm(`'${clash}' 발주서 양식이 이미 있습니다.\n새 파일로 바꿀까요?`)) return;
      close();
      resolve(list.map((it, i) => Object.assign({}, it, { name: names[i] })));
    };
    const onBg = e => { if (e.target === $("vnmodal")) onCancel(); };
    function cleanup() {
      $("vn-cancel").onclick = null; $("vn-ok").onclick = null; $("vnmodal").onclick = null;
    }
    $("vn-cancel").onclick = onCancel;
    $("vn-ok").onclick = onOk;
    $("vnmodal").onclick = onBg;
  });
}

/* 업체 양식 저장. list = [{fileName, buf, name}]
   name 이 없으면 파일명에서 자동으로 짓되, 이름이 겹치면 덮어쓰지 않고 다른 후보로 피한다. */
async function saveForms(list) {
  const existing = await DB.listForms();
  const byName = new Map(existing.map(f => [f.name, f]));
  let added = 0, updated = 0;
  const names = [];
  for (const it of list) {
    let name = it.name;
    if (!name) {
      const cands = candidateNames(it.fileName);
      const same = existing.find(f => f.file === it.fileName);
      if (same) name = same.name;
      else {
        name = cands[0];
        let i = 1;
        while (byName.has(name) && i <= 50) {
          name = (i < cands.length) ? cands[i] : `${cands[0]} (${i - cands.length + 2})`;
          i++;
        }
      }
    }
    if (byName.has(name)) updated++; else added++;
    const rec = { name, file: it.fileName, data: it.buf, checked: true };
    await addForm(rec);          // 지웠던 이름이면 '지움' 표시도 같이 걷어낸다
    byName.set(name, rec);
    names.push(name);
  }
  return { added, updated, names };
}

async function addForms(files) {
  const list = [];
  for (const f of files) {
    if (!/\.xls[xm]$/i.test(f.name)) continue;
    list.push({ fileName: f.name, buf: await readFile(f) });
  }
  if (!list.length) return;
  const named = await askVendorNames(list);
  if (!named) return;                      // 사용자가 취소
  const r = await saveForms(named);
  await loadForms();
  msg("msg-o", "ok", `✔ 발주서 양식 ${r.added + r.updated}개 저장 — ${r.names.join(", ")}`
    + (r.updated ? ` (${r.updated}개는 기존 양식 갱신)` : "")
    + "\n다음부터는 체크만 하면 됩니다.");
}

/* 업체명 바꾸기 — 브랜드 학습·업체 메일 이력도 같이 옮긴다 */
async function renameForm(f) {
  const v = prompt("업체명을 바꿉니다.\n(발주서 파일명과 브랜드 배정에 쓰이는 이름)", f.name);
  if (v === null) return;
  const newName = v.trim().replace(/[\\/:*?"<>|]/g, "");
  if (!newName || newName === f.name) return;
  if (S.forms.some(x => x.name === newName)) { alert("같은 이름의 발주서 양식이 이미 있어요."); return; }
  await dropForm(f.name);
  await addForm({ name: newName, file: f.file, data: f.data, checked: f.checked !== false });
  for (const b in S.brandVendor) if (S.brandVendor[b] === f.name) S.brandVendor[b] = newName;
  await DB.set("brandVendor", S.brandVendor);
  const move = async (obj, key) => {
    if (obj && obj[f.name] !== undefined) { obj[newName] = obj[f.name]; delete obj[f.name]; await DB.set(key, obj); }
  };
  await move(S.vendorEmails, "vendorEmails");
  await move(S.vendorSent, "vendorSent");
  await move(S.vendorDomains, "vendorDomains");
  if (S.sel[f.name]) { S.sel[newName] = S.sel[f.name]; delete S.sel[f.name]; }
  await loadForms();
  msg("msg-o", "ok", `✔ '${f.name}' → '${newName}' 으로 바꿨어요.`);
}
/* 업체 양식 지우기 — 지운 이름을 남겨야 드라이브 백업에서 되살아나지 않는다.
   (2026-08-04: 지운 중복 양식이 앱을 새로 띄울 때마다 돌아와 업체 2곳에 양식이 4개가 됐다) */
async function dropForm(name) {
  await DB.delForm(name);
  const tomb = await DB.get("formsDeleted", {}) || {};
  tomb[name] = Date.now();
  await DB.set("formsDeleted", tomb);
  // 바로 백업에도 반영 — 다음에 켤 때 옛 백업이 다시 내려보내지 못하게
  try { if (window.SYNC && !GMAIL.needLogin()) await SYNC.syncUpNow(); } catch (e) {}
}
/* 업체 양식 넣기 — 같은 이름을 지운 적이 있으면 그 표시를 걷어낸다 */
async function addForm(f) {
  await DB.putForm(f);
  const tomb = await DB.get("formsDeleted", {}) || {};
  if (tomb[f.name] !== undefined) { delete tomb[f.name]; await DB.set("formsDeleted", tomb); }
}
async function loadForms() {
  S.forms = await DB.listForms();
  S.brandVendor = await DB.get("brandVendor", {});
  S.vendorEmails = await DB.get("vendorEmails", {});
  S.vendorSent = await DB.get("vendorSent", {});
  S.vendorDomains = await DB.get("vendorDomains", {});
  S.invEmails = await DB.get("invEmails", "");
  S.invSent = await DB.get("invSent", []);
  await MAILTPL.load();                     // 저장해둔 메일 문구
  drawForms(); buildVendorBrands(); refreshO();
}
/* 이름이 비슷한 것끼리 묶어 중복을 찾는다.
   '디에스피' 와 '디에스피_발주양식' 처럼 같은 업체를 두 번 올린 경우를 잡는다.
   자동으로 지우지는 않는다 — 파일 자체가 다를 수 있어서 판단은 사람이 한다.
   items: [{name, ...}] 형태면 되고, 이름은 nameOf 로 뽑는다. */
function dupNameGroups(items, nameOf) {
  const norm = s => String(s || "").replace(/[\s_\-().\[\]]/g, "").toLowerCase();
  const groups = [];
  (items || []).forEach(it => {
    const n = norm(nameOf ? nameOf(it) : it.name);
    if (!n) return;
    const g = groups.find(x => x.key.indexOf(n) >= 0 || n.indexOf(x.key) >= 0);
    if (g) { g.items.push(it); if (n.length < g.key.length) g.key = n; }
    else groups.push({ key: n, items: [it] });
  });
  return groups.filter(g => g.items.length > 1);
}
/* 중복 경고 상자를 목록 맨 위에 붙인다 */
function addDupWarning(box, groups, nameOf, tail) {
  if (!groups.length) return;
  const w = document.createElement("div");
  w.className = "msg show warn";
  w.innerHTML = "⚠ 같은 업체가 두 번 등록된 것 같아요.<br>"
    + groups.map(g => "· " + g.items.map(x => `<b>${esc(nameOf(x))}</b>`).join(" / ")).join("<br>")
    + `<br><span style="color:var(--muted)">${esc(tail)}</span>`;
  box.appendChild(w);
}
/* 중복 양식을 여기서 바로 지운다 — 목록에서 ✕ 를 찾아 누르는 것보다 확실하다.
   지우면 '지운 양식' 표시가 남아, 드라이브 백업에서 다시 내려오지 않는다. */
function addDupFixer(box, groups) {
  if (!groups.length) return;
  const w = document.createElement("div");
  w.style.cssText = "margin-top:6px;display:flex;flex-wrap:wrap;gap:6px;align-items:center";
  w.innerHTML = `<span style="font-size:11.5px;color:var(--muted);flex:1 1 100%">지울 쪽을 눌러주세요</span>`;
  groups.forEach(g => g.items.forEach(f => {
    const b = document.createElement("button");
    b.className = "minibtn";
    b.textContent = "✕ " + f.name;
    b.onclick = async ev => {
      ev.stopPropagation();
      if (!confirm(`'${f.name}' 양식을 지울까요?\n(다시 살아나지 않게 표시도 같이 남깁니다)`)) return;
      await dropForm(f.name);
      await loadForms();
    };
    w.appendChild(b);
  }));
  box.appendChild(w);
}
function drawForms() {
  const box = $("vlist");
  if (!S.forms.length) { box.innerHTML = '<div class="empty">저장된 발주서 양식이 없습니다</div>'; return; }
  box.innerHTML = "";
  const dups = dupNameGroups(S.forms, f => f.name);
  addDupWarning(box, dups, f => f.name,
    "브랜드 선택에도 두 번 나옵니다. 안 쓰는 쪽은 아래에서 지워주세요.");
  addDupFixer(box, dups);
  S.forms.forEach(f => {
    const el = document.createElement("div");
    el.className = "vrow" + (f.checked ? " on" : "");
    el.innerHTML = `<div class="vtop"><span class="box">${CHK}</span><b>${esc(f.name)}</b><button class="vdel">✕</button></div>
      <span class="vfile">${esc(f.file)}</span>
      <div class="vbtns"><button class="pv">미리보기</button><button class="rn">이름수정</button><button class="dl">엑셀 받기</button></div>`;
    el.onclick = async ev => {
      if (ev.target.classList.contains("pv")) {
        ev.stopPropagation();
        openPreview(f.data, f.name + " 양식"); return;
      }
      if (ev.target.classList.contains("rn")) {
        ev.stopPropagation();
        await renameForm(f); return;
      }
      if (ev.target.classList.contains("dl")) {
        ev.stopPropagation();
        download(f.data, f.file || (f.name + ".xlsx")); return;   // 실제 엑셀 파일 다운로드
      }
      if (ev.target.classList.contains("vdel")) {
        ev.stopPropagation();
        if (!confirm(`'${f.name}' 양식을 지울까요?`)) return;
        await dropForm(f.name); await loadForms(); return;
      }
      f.checked = !f.checked;
      el.classList.toggle("on", f.checked);
      await DB.putForm(f);
      buildVendorBrands(); refreshO();
    };
    box.appendChild(el);
  });
}

/* --- 업체별 브랜드 --- */
// 한 브랜드는 한 업체에만 배정 가능. 다른 업체가 이미 가진 브랜드는 여기서 못 고름.
function brandOwner(b, exceptName) {
  const checked = S.forms.filter(f => f.checked);
  for (const f of checked) {
    if (f.name === exceptName) continue;
    if ((S.sel[f.name] || []).includes(b)) return f.name;
  }
  return null;
}
function buildVendorBrands() {
  const checked = S.forms.filter(f => f.checked);
  const card = $("card3"), box = $("vbrands");
  if (!checked.length || !S.brands.length) { card.style.display = "none"; box.innerHTML = ""; return; }
  card.style.display = "block"; box.innerHTML = "";
  // 자동 배정 초기화 (학습된 brand_vendor) — 단, 다른 업체가 이미 쥔 건 제외
  checked.forEach(f => {
    if (!S.sel[f.name]) S.sel[f.name] = S.brands.filter(b => S.brandVendor[b] === f.name);
  });
  checked.forEach(f => {
    const wrap = document.createElement("div");
    wrap.className = "vendorbox";
    wrap.dataset.vendor = f.name;
    wrap.innerHTML = `<div class="vh">🏭 ${esc(f.name)}<span class="cnt"></span><button class="all">전체</button></div>
      <div class="brands"></div>`;
    box.appendChild(wrap);
    renderVendorChips(wrap, f);
    wrap.querySelector(".all").onclick = () => {
      const mine = S.sel[f.name] || [];
      // 자유롭거나 내 것인 브랜드만 대상 (남의 것은 건드리지 않음)
      const selectable = S.brands.filter(b => brandOwner(b, f.name) === null || mine.includes(b));
      S.sel[f.name] = (mine.length === selectable.length && mine.length > 0) ? [] : selectable;
      buildVendorBrands();                                    // 다른 업체 표시도 갱신
    };
  });
  refreshO();
}
function renderVendorChips(wrap, f) {
  const brandsBox = wrap.querySelector(".brands");
  brandsBox.innerHTML = S.brands.map(b => {
    const mine = (S.sel[f.name] || []).includes(b);
    const owner = brandOwner(b, f.name);
    if (mine) return `<span class="brow on" data-b="${esc(b)}"><span class="box">${CHK}</span>${esc(b)}</span>`;
    if (owner) return `<span class="brow taken" data-b="${esc(b)}" title="${esc(owner)}가 선택함">${esc(b)} <small>· ${esc(owner)}</small></span>`;
    return `<span class="brow" data-b="${esc(b)}"><span class="box">${CHK}</span>${esc(b)}</span>`;
  }).join("");
  brandsBox.querySelectorAll(".brow").forEach(chip => {
    chip.onclick = () => {
      const b = chip.dataset.b;
      const arr = S.sel[f.name] || (S.sel[f.name] = []);
      const i = arr.indexOf(b);
      if (i >= 0) { arr.splice(i, 1); }                       // 내 것 → 해제
      else {
        // 다른 업체가 쥐고 있으면 그쪽에서 떼어내 이리로 이동
        const owner = brandOwner(b, f.name);
        if (owner) { const o = S.sel[owner]; const k = o.indexOf(b); if (k >= 0) o.splice(k, 1); }
        arr.push(b);
      }
      buildVendorBrands();                                    // 전 업체 다시 그려서 중복 방지 반영
    };
  });
  updCnt(wrap, f);
}
function updCnt(wrap, f) {
  const n = S.sel[f.name].length;
  wrap.querySelector(".cnt").textContent = n ? `${n}개 선택` : "선택 없음 → 건너뜀";
}

/* 불러온 파일 '해제' — 드롭 영역에 파일이 올라와 있을 때만 버튼을 보여준다.
   잘못된 파일을 올렸을 때 새로고침하지 않고 비울 수 있게. */
function updateClearRows() {
  const a = $("order-clear-row"); if (a) a.style.display = S.orderBuf ? "flex" : "none";
  const b = $("sab-clear-row"); if (b) b.style.display = S.sabBuf ? "flex" : "none";
}
async function clearOrderFile() {
  S.orderWb = null; S.orderBuf = null; S.orderName = "";
  S.brands = []; S.dateSel = [];
  $("order-name").textContent = "";
  $("drop-order").classList.remove("on");
  const fi = $("f-order"); if (fi) fi.value = "";
  const dw = $("date-wrap"); if (dw) dw.style.display = "none";
  const ro = $("result-o"); if (ro) ro.style.display = "none";
  msg("msg-o", "", "");
  buildVendorBrands(); refreshO(); updateClearRows();
}
function clearSabFile() {
  S.sabBuf = null; S.sabName = ""; S.sabDrive = null;
  $("sab-name").textContent = "";
  $("drop-sab").classList.remove("on");
  const fi = $("f-sab"); if (fi) fi.value = "";
  $("sab-preview").style.display = "none";
  const ri = $("result-i"); if (ri) ri.style.display = "none";
  msg("msg-i", "", "");
  refreshI(); updateClearRows();
}
if ($("order-clear")) $("order-clear").onclick = clearOrderFile;
if ($("sab-clear")) $("sab-clear").onclick = clearSabFile;

function refreshO() {
  updateClearRows();
  let ok = !!S.orderBuf && S.forms.some(f => f.checked);
  if (ok && S.brands.length) ok = S.forms.some(f => f.checked && (S.sel[f.name] || []).length);
  if (ok && S.dateAll.length && !S.dateSel.length) ok = false;
  $("run-o").disabled = !ok;
}

/* --- 변환 실행 --- */
$("run-o").onclick = async function () {
  busy("run-o", "run-o-lbl", true, "변환 중…");
  msg("msg-o", "", "");
  try {
    const picked = S.forms.filter(f => f.checked);
    const results = [], skipped = [];
    for (const f of picked) {
      const sel = S.sel[f.name] || [];
      if (S.brands.length && !sel.length) { skipped.push(f.name + "(브랜드 미선택)"); continue; }
      const orderWb = await QO.loadWorkbook(S.orderBuf.slice(0));
      const tplWb = await QO.loadWorkbook(f.data.slice(0));
      const brandFilter = (S.brands.length && sel.length !== S.brands.length) ? sel : null;
      const r = QO.convert(orderWb, tplWb, {
        brands: brandFilter, dates: S.dateSel.length ? S.dateSel : null, dateHeader: S.dateHeader,
      });
      const out = await QO.saveWorkbook(tplWb);
      results.push({ supplier: f.name, count: r.count, buf: out,
        // 파일명 고정: 오늘날짜_랩노마드_업체명_발주서.xlsx
        // ※ 사용자가 별도로 요청하지 않는 한 이 형식을 바꾸지 말 것
        // 파일명 고정 형식: 오늘날짜_랩노마드_업체명_발주양식.xlsx
        // 가운데 이름은 CONFIG.orderTag (랩노마드는 "랩노마드" 고정). 비면 그 자리를 뺀다.
        filename: `${QO.todayStr()}_${CONFIG.orderTag ? CONFIG.orderTag + "_" : ""}${cleanVendor(f.name)}_발주양식.xlsx` });
      // 학습 저장
      if (S.brands.length && sel.length) sel.forEach(b => { S.brandVendor[b] = f.name; });
    }
    if (!results.length) throw new Error("변환된 업체가 없습니다. " + (skipped.length ? `(${skipped.join(", ")})` : ""));
    await DB.set("brandVendor", S.brandVendor);

    // --- 건수 검증: 원본 주문 수 == 업체별 변환 합계 (미배정 주문이 조용히 빠지는 것 방지) ---
    const srcWb = await QO.loadWorkbook(S.orderBuf.slice(0));
    const src = QO.countOrders(srcWb, { dates: S.dateSel.length ? S.dateSel : null, dateHeader: S.dateHeader });
    const converted = results.reduce((a, r) => a + r.count, 0);
    const hasBrands = S.brands.length > 0;
    // 브랜드가 있으면 주문이 업체별로 나뉨(합계=원본). 브랜드 열이 없으면 업체마다 전량 복사.
    const expected = hasBrands ? src.total : src.total * results.length;
    // 어느 업체에도 배정되지 않은 브랜드 찾기 (= 발주서에서 빠진 주문)
    const assigned = new Set();
    picked.forEach(f => (S.sel[f.name] || []).forEach(b => assigned.add(String(b).trim())));
    const unassigned = [];
    if (hasBrands) {
      for (const b in src.byBrand)
        if (!assigned.has(b)) unassigned.push({ brand: b || "(브랜드 없음)", count: src.byBrand[b] });
    }
    unassigned.sort((x, y) => y.count - x.count);
    const verify = { srcTotal: src.total, converted, expected, diff: expected - converted, unassigned, hasBrands };

    showResultO(results, skipped, verify);
    msg("msg-o", verify.diff === 0 ? "ok" : "warn",
      (verify.diff === 0 ? "✔ 변환 완료! 건수 일치 " : "⚠ 변환 완료 (건수 불일치 확인) ")
      + `주문 ${src.total}건 → ` + results.map(r => `${r.supplier}=${r.count}건`).join("; "));
  } catch (e) { msg("msg-o", "err", "변환 실패: " + e.message); }
  finally { busy("run-o", "run-o-lbl", false, "발주서 변환하기"); refreshO(); }
};

function showResultO(results, skipped, verify) {
  const box = $("rlist-o"); box.innerHTML = "";

  // --- 건수 검증: 쇼핑몰 주문 수 == 업체 발주서 합계 ---
  if (verify) {
    const d = document.createElement("div");
    const detail = results.map(r => `${r.supplier} ${r.count}`).join(" + ");
    if (verify.diff === 0) {
      d.className = "msg show ok";
      d.textContent = verify.hasBrands
        ? `✔ 건수 일치 — 주문 ${verify.srcTotal}건 = ${detail} (합계 ${verify.converted}건)`
        : `✔ 건수 일치 — 주문 ${verify.srcTotal}건이 업체별로 전량 반영됨 (${detail})`;
    } else if (verify.diff > 0) {
      let h = `⚠ 건수 불일치 — 주문 ${verify.srcTotal}건 중 발주서에 ${verify.converted}건만 들어갔습니다 (${verify.diff}건 누락)<br>${esc(detail)}`;
      if (verify.unassigned.length) {
        h += "<br><br>아래 브랜드가 어느 업체에도 배정되지 않아 발주서에서 빠졌습니다:";
        verify.unassigned.forEach(u => { h += `<br>· ${esc(u.brand)} — ${u.count}건`; });
        // 새로 들어온 브랜드는 이름을 그대로 크게 보여준다 (놓치면 발주 누락으로 이어짐)
        const names = verify.unassigned.map(u => u.brand).join(", ");
        h += `<div class="callout">🔔 신규 브랜드<br><b>${esc(names)}</b><br>의 업체를 선택해주세요</div>`;
      }
      d.className = "msg show err"; d.innerHTML = h;
    } else {
      d.className = "msg show err";
      d.textContent = `⚠ 발주서 합계(${verify.converted}건)가 주문 수(${verify.srcTotal}건)보다 많습니다. 같은 브랜드가 여러 업체에 중복 배정됐는지 확인하세요.\n${detail}`;
    }
    box.appendChild(d);
  }

  results.forEach(r => {
    const el = document.createElement("div");
    el.className = "rrow";
    el.innerHTML = `<div class="rtop"><div class="vinfo"><b>${esc(cleanVendor(r.supplier))}</b><span>${esc(r.filename)}</span></div>
      <span class="cnt">${r.count}건</span></div>
      <div class="cands"></div>
      <div class="rmail"><input type="text" placeholder="${esc(r.supplier)} 이메일 (여러 개는 쉼표로)"
        value="${esc(S.vendorEmails[r.supplier] || "")}" inputmode="email" autocapitalize="off" autocorrect="off" spellcheck="false">
        <button class="dlbtn send">메일 보내기</button></div>
      <div class="setrow" style="margin-top:6px"><span style="flex:1;font-size:11px;color:var(--faint)"></span>
        <button class="minibtn share">📤 카톡·공유</button><button class="minibtn pvbtn">미리보기</button><button class="minibtn dl">엑셀 받기</button></div>
      <div class="setrow" style="margin-top:4px"><span style="flex:1;font-size:11px;color:var(--faint)"></span>
        <button class="minibtn tpl">${mailtplLabel("order", r.supplier)}</button><button class="minibtn fn">✏️ 파일명 수정</button></div>`;
    const inp = el.querySelector("input");
    bindMailtplBtn(el.querySelector(".tpl"), "order", r.supplier, { 업체: r.supplier, 건수: r.count });
    fillRecipients(el.querySelector(".cands"), inp, {
      saved: S.vendorEmails[r.supplier], history: S.vendorSent[r.supplier],
      domains: parseDomains(S.vendorDomains[r.supplier]), query: r.supplier });
    // 파일명은 자동으로 지어지지만, 원하면 여기서 바꿀 수 있다 (저장·메일 첨부에 모두 반영)
    el.querySelector(".fn").onclick = () => {
      const cur = r.filename.replace(/\.xlsx$/i, "");
      const v = prompt("저장·발송될 파일명을 바꿉니다.\n(확장자 .xlsx 는 자동으로 붙습니다)", cur);
      if (v === null) return;
      const t = v.trim().replace(/[\\/:*?"<>|]/g, "").replace(/\.xlsx$/i, "");
      if (!t) return;
      r.filename = t + ".xlsx";
      el.querySelector(".vinfo span").textContent = r.filename;
    };
    el.querySelector(".pvbtn").onclick = () => openPreview(r.buf, r.supplier + " 발주서");
    el.querySelector(".dl").onclick = () => download(r.buf, r.filename);
    el.querySelector(".share").onclick = () => shareFile(r.buf, r.filename);
    inp.onchange = inp.onblur = async () => {
      const v = inp.value.trim(); if (v === (S.vendorEmails[r.supplier] || "")) return;
      S.vendorEmails[r.supplier] = v; await DB.set("vendorEmails", S.vendorEmails);
    };
    const sendBtn = el.querySelector(".send");
    sendBtn.onclick = async () => {
      const list = parseEmails(inp.value);
      if (!list.length) { inp.focus(); return; }
      const bad = invalidEmails(list);
      if (bad.length) { alert("이메일 형식이 이상해요:\n" + bad.join(", ")); inp.focus(); return; }
      sendBtn.disabled = true; sendBtn.textContent = "보내는 중…";
      try {
        await ensureGmail();
        const ymd = QO.todayStr().slice(2);
        const tpl = MAILTPL.render("order", { 회사: CONFIG.company, 업체: r.supplier, 날짜: ymd, 건수: r.count });
        await GMAIL.send({ to: list.join(", "), subject: tpl.subject, body: tpl.body,
          attachments: [{ filename: r.filename, data: r.buf }] });
        S.vendorEmails[r.supplier] = inp.value.trim(); await DB.set("vendorEmails", S.vendorEmails);
        await recordSent(r.supplier, list);      // 보낸 주소들을 이력에 기억
        sendBtn.textContent = list.length > 1 ? `✓ ${list.length}명 발송완료` : "✓ 발송완료";
        sendBtn.style.background = "var(--ok)";
      } catch (e) { sendBtn.disabled = false; sendBtn.textContent = "메일 보내기"; alert("발송 실패: " + e.message); }
    };
    box.appendChild(el);
  });
  if (skipped.length) {
    const w = document.createElement("div");
    w.className = "msg show warn"; w.textContent = "건너뜀: " + skipped.join(", ");
    box.appendChild(w);
  }
  $("result-o").style.display = "block";
  $("result-o").scrollIntoView({ behavior: "smooth", block: "start" });
}

/* 보낸 주소 목록을 최근 순·중복 제거로 병합(최대 12개) */
function mergeRecent(arr, emails, cap) {
  const cur = (arr || []).slice();
  for (const e of emails) {
    const i = cur.findIndex(x => x.toLowerCase() === e.toLowerCase());
    if (i >= 0) cur.splice(i, 1);
    cur.unshift(e);
  }
  return cur.slice(0, cap || 12);
}
/* 업체 발주서 발송 이력 저장 */
async function recordSent(supplier, emails) {
  S.vendorSent[supplier] = mergeRecent(S.vendorSent[supplier], emails);
  await DB.set("vendorSent", S.vendorSent);
}
/* 송장 취합본 발송 이력 저장 */
async function recordSentInv(emails) {
  S.invSent = mergeRecent(S.invSent, emails);
  await DB.set("invSent", S.invSent);
}

/* 받는사람 추천 칩 (발주/송장 공용)
   opts = { saved:"주소들", history:[주소...], domains:[도메인...], query:"메일검색어" }
   ⓐ저장(기본) ⓑ이전 발송 ⓒ메일에서 찾은 주소(query+도메인 있을 때, 도메인 필터) → 클릭 선택 */
/* 메일함에서 찾은 주소를 '업체 담당자일 확률' 순으로 정리한다.
   그냥 세어 올리면 우리 쪽 사람·다른 몰 담당자·본인 개인메일까지 다 올라온다.
   업체명을 로마자로 옮겨 도메인과 맞추는 건 어려우니 데이터에서 단서를 찾는다:
     · 한 도메인에 여러 명이 나오면 그 업체 회사다 (plasmer.co.kr 에 7명)
     · 회사 도메인(공용 메일 아님)이 개인 gmail 보다 먼저다
     · 보낸사람으로 나온 주소가 참조만 된 주소보다 먼저다 */
const FREEMAIL_RE = /^(gmail|googlemail|naver|daum|hanmail|nate|kakao|hanmir|empas|korea|outlook|hotmail|live|msn|yahoo|icloud|me|proton|protonmail|aol|qq|163)\./i;
function rankVendorAddrs(found, known) {
  const list = (found || []).filter(f => f && f.email);
  if (!list.length) return [];
  // 이 업체와 실제로 주고받은 적 있는 주소 (저장된 주소 + 발송 이력) — 개인메일이어도 살린다
  const knownSet = new Set((known || []).map(e => String(e).toLowerCase()));
  const byDom = {};
  list.forEach(f => {
    const d = (String(f.email).split("@")[1] || "").toLowerCase();
    (byDom[d] = byDom[d] || []).push(f);
  });
  const scored = list.map(f => {
    const lc = String(f.email).toLowerCase();
    const dom = (lc.split("@")[1] || "");
    const free = FREEMAIL_RE.test(dom + ".");
    const mates = byDom[dom].length;            // 같은 도메인에 몇 명이나 나왔나
    const isKnown = knownSet.has(lc);
    let s = 0;
    if (isKnown) s += 60;                        // 이미 이 업체와 주고받은 주소
    if (!free) s += 20;                          // 회사 도메인
    if (!free && mates >= 2) s += 25 + Math.min(mates, 6) * 3;   // 여럿 = 그 업체 회사
    s += Math.min(f.fromCount || 0, 3) * 6;      // 보낸사람으로 등장
    s += Math.min(f.count || 0, 4) * 2;          // 자주 등장
    return { email: f.email, score: s, free, mates, known: isKnown };
  }).sort((a, b) => b.score - a.score);

  // 회사 도메인이 하나라도 잡히면 그게 업체다.
  // 그때 개인메일(gmail·naver…)은 이 업체와 실제로 주고받은 적이 있는 것만 남긴다.
  // — 검색에 걸렸다는 이유만으로 우리 쪽 사람이나 남의 개인메일이 후보로 뜨면 안 된다.
  const hasCorp = scored.some(x => !x.free);
  if (!hasCorp) return scored;                   // 업체가 개인메일만 쓰는 경우
  return scored.filter(x => !x.free || x.known);
}

async function fillRecipients(container, inp, opts) {
  if (!container) return;
  opts = opts || {};
  container.innerHTML = "";
  const chosenSet = () => new Set(parseEmails(inp.value).map(e => e.toLowerCase()));
  const seen = new Set();
  const lbl = document.createElement("span");
  lbl.className = "candlbl"; lbl.textContent = "받는사람 추천:";
  container.appendChild(lbl);

  function refreshStates() {
    const cs = chosenSet();
    container.querySelectorAll(".cand").forEach(x => x.classList.toggle("on", cs.has((x.dataset.email || "").toLowerCase())));
  }
  function addChip(email) {
    const lc = email.toLowerCase();
    if (!lc || seen.has(lc)) return; seen.add(lc);
    const c = document.createElement("button");
    c.className = "cand" + (chosenSet().has(lc) ? " on" : "");
    c.dataset.email = email; c.textContent = email;
    c.onclick = () => {
      const cur = parseEmails(inp.value);
      const i = cur.findIndex(e => e.toLowerCase() === lc);
      if (i >= 0) cur.splice(i, 1); else cur.push(email);
      inp.value = cur.join(", ");
      inp.dispatchEvent(new Event("change"));
      refreshStates();
    };
    container.appendChild(c);
  }
  inp.addEventListener("input", refreshStates);

  // 이 업체와 이미 주고받은 주소 — 개인메일이어도 후보로 살려두는 근거가 된다
  const knownAddrs = () => parseEmails(opts.saved || "").concat(opts.history || []);
  const doms = opts.domains || [];
  const inDom = e => { const d = (String(e).split("@")[1] || "").toLowerCase(); return doms.some(x => d === x || d.endsWith("." + x)); };

  // ⓐ 저장(기본) + ⓑ 이전에 보냈던 주소들 (항상 후보)
  parseEmails(opts.saved || "").forEach(addChip);
  (opts.history || []).forEach(addChip);

  // ⓒ 메일에서 검색한 주소 — query가 있을 때만
  if (opts.query) {
    if (GMAIL.signedIn() && doms.length) {
      const hint = document.createElement("span");
      hint.className = "candhint"; hint.textContent = "메일에서 찾는 중…";
      container.appendChild(hint);
      try {
        const found = await GMAIL.searchAddresses({ query: opts.query, max: 20 });
        hint.remove();
        // 도메인을 설정해뒀으면 그것만, 아니면 연관성 순으로
        const inDomList = found.map(f => f.email).filter(inDom);
        if (inDomList.length) inDomList.slice(0, 8).forEach(addChip);
        else rankVendorAddrs(found, knownAddrs()).slice(0, 8).forEach(x => addChip(x.email));
      } catch (e) { hint.remove(); }
      if (seen.size === 0) {
        const s = document.createElement("span");
        s.className = "candhint"; s.textContent = "해당 도메인 주소를 못 찾음 — 직접 입력하세요";
        container.appendChild(s);
      }
    } else if (!doms.length) {
      // 도메인을 안 넣어뒀어도 메일함에서 찾을 수 있게 버튼을 준다.
      // (도메인이 있으면 위에서 자동으로 걸러 보여주고, 없으면 찾은 주소를 그대로 보여준다)
      const b = document.createElement("button");
      b.className = "cand ghost"; b.textContent = "📬 메일에서 주소 찾기";
      b.onclick = async () => {
        const old = b.textContent;
        b.textContent = "찾는 중…"; b.disabled = true;
        try { await ensureGmail(); } catch (e) { b.textContent = old; b.disabled = false; return; }
        try {
          const found = await GMAIL.searchAddresses({ query: opts.query, max: 30 });
          const list = rankVendorAddrs(found, knownAddrs()).map(x => x.email);
          b.remove();
          if (!list.length) {
            const s2 = document.createElement("span");
            s2.className = "candhint";
            s2.textContent = `'${opts.query}' 로 주고받은 메일에서 주소를 못 찾았어요 — 직접 입력하세요`;
            container.appendChild(s2);
          } else list.slice(0, 12).forEach(addChip);
        } catch (e) { b.textContent = "다시 시도"; b.disabled = false; }
      };
      container.appendChild(b);
      const s = document.createElement("span");
      s.className = "candhint";
      s.textContent = "";
      container.appendChild(s);
    } else {
      const b = document.createElement("button");
      b.className = "cand ghost"; b.textContent = "＋ 메일에서 받는사람 찾기";
      b.onclick = async () => {
        b.textContent = "로그인 중…";
        try { await ensureGmail(); } catch (e) { b.textContent = "＋ 메일에서 받는사람 찾기"; return; }
        fillRecipients(container, inp, opts);
      };
      container.appendChild(b);
    }
  } else if (seen.size === 0) {
    // 송장 취합본 등: 검색어 없음. 이력이 없으면 안내만.
    const s = document.createElement("span");
    s.className = "candhint"; s.textContent = "한 번 보내면 다음부터 여기서 골라 보낼 수 있어요.";
    container.appendChild(s);
  }
}

/* =================================================================
   ② 송장 취합
   ================================================================= */
$("f-sab").addEventListener("change", async function () {
  if (this.files[0]) { S.sabBuf = await readFile(this.files[0]); S.sabName = this.files[0].name; S.sabDrive = null;
    $("sab-name").textContent = "📄 " + this.files[0].name; $("drop-sab").classList.add("on"); $("sab-preview").style.display="block"; refreshI(); }
});
bindDrop("drop-sab", async f => {
  S.sabBuf = await readFile(f[0]); S.sabName = f[0].name; S.sabDrive = null;
  $("sab-name").textContent = "📄 " + f[0].name; $("drop-sab").classList.add("on"); $("sab-preview").style.display="block"; refreshI();
});
$("f-rep").addEventListener("change", function () { const fs = [...this.files]; this.value = ""; addReps(fs); });
bindDrop("drop-rep", f => addReps(f));

async function addReps(files) {
  for (const f of files) {
    if (!/\.xls[xm]$/i.test(f.name)) continue;
    if (S.reps.some(r => r.name === f.name)) continue;
    S.reps.push({ name: f.name, data: await readFile(f) });
  }
  drawReps();
}
function drawReps() {
  const box = $("replist"); box.innerHTML = "";
  // 같은 업체 회신이 두 번 들어오면 송장이 겹쳐 들어갈 수 있어 먼저 알려준다
  addDupWarning(box, dupNameGroups(S.reps, r => QO.nameFromFilename(r.name)),
    r => QO.nameFromFilename(r.name), "같은 업체 회신이 두 번 들어와 있습니다. 최신 것만 남기고 ✕ 로 지워주세요.");
  S.reps.forEach((r, i) => {
    const el = document.createElement("div");
    el.className = "vrow on";
    el.innerHTML = `<div class="vtop"><span class="box">${CHK}</span><b>${esc(QO.nameFromFilename(r.name))}</b><button class="vdel">✕</button></div>
      <span class="vfile">${esc(r.name)}</span>
      <div class="vbtns"><button class="pv">미리보기</button><button class="dl">엑셀 받기</button></div>`;
    el.querySelector(".pv").onclick = e => { e.stopPropagation(); openPreview(r.data, QO.nameFromFilename(r.name) + " 회신"); };
    el.querySelector(".dl").onclick = e => { e.stopPropagation(); download(r.data, r.name); };
    el.querySelector(".vdel").onclick = e => { e.stopPropagation(); S.reps.splice(i, 1); drawReps(); };
    box.appendChild(el);
  });
  if (S.reps.length) $("drop-rep").classList.add("on"); else $("drop-rep").classList.remove("on");
  refreshI();
}
function refreshI() { $("run-i").disabled = !(S.sabBuf && S.reps.length); updateClearRows(); }

$("run-i").onclick = async function () {
  busy("run-i", "run-i-lbl", true, "취합 중…");
  msg("msg-i", "", "");
  try {
    const sab = await QO.loadWorkbook(S.sabBuf.slice(0));
    const replies = [];
    for (const r of S.reps) replies.push({ name: r.name, wb: await QO.loadWorkbook(r.data.slice(0)) });
    // 업체별로 지난번에 쓴 택배사를 넘겨준다 — 회신에 택배사가 비어 와도 채울 수 있게
  const carriers = await DB.get("vendorCarriers", {}) || {};
  const out = QO.collectInvoices(sab, replies, { carriers });
  // 이번에 실제로 쓰인 택배사를 기억해 다음 회신에 쓴다
  if (out.carrierSeen && Object.keys(out.carrierSeen).length) {
    await DB.set("vendorCarriers", Object.assign(carriers, out.carrierSeen));
  }
    const buf = await QO.saveWorkbook(sab);
    const stem = S.sabName.replace(/\.[^.]+$/, "");
    showResultI(out, buf, `${QO.todayStr()}_${stem}_송장취합.xlsx`);
    msg("msg-i", "ok", `✔ 취합 완료! 총 ${out.total}건 기입`);
  } catch (e) { msg("msg-i", "err", "취합 실패: " + e.message); }
  finally { busy("run-i", "run-i-lbl", false, "송장 취합하기"); refreshI(); }
};

/* 취합본 빈칸(누락) 주문을 읽기 쉬운 표로 */
function missingTable(rows, total) {
  if (!rows.length) return "";
  const COLS = [["RECIPIENT", "수취인"], ["PRODUCT", "상품"], ["OPTION", "옵션"], ["QTY", "수량"], ["ORDERER", "주문자"], ["ADDR", "주소"]];
  let use = COLS.filter(([k]) => rows.some(r => r[k] != null && String(r[k]).trim() !== ""));
  if (!use.length) use = [["label", "주문"]];
  let h = `<div class="tblbox" style="margin-top:8px"><div class="tblscroll"><table class="pv"><tr><th>#</th>${use.map(c => `<th>${esc(c[1])}</th>`).join("")}</tr>`;
  rows.forEach((r, i) => {
    h += `<tr><td class="num">${i + 1}</td>` + use.map(c => {
      const v = r[c[0]] == null ? "" : String(r[c[0]]);
      const num = /^[0-9,.\-]+$/.test(v) && v !== "";
      return `<td${num ? ' class="num"' : ""}>${esc(v)}</td>`;
    }).join("") + "</tr>";
  });
  h += "</table></div>";
  if (total > rows.length) h += `<div class="pvfoot">앞 ${rows.length}건만 표시 · 전체 ${total}건</div>`;
  return h + "</div>";
}

function showResultI(out, buf, filename) {
  let h = `<div class="tblbox" style="margin-bottom:12px"><div class="tblscroll"><table class="pv">
    <tr><th>업체</th><th>기입</th><th>미매칭</th><th>상태</th></tr>`;
  out.per.forEach(p => { h += `<tr><td>${esc(p[0])}</td><td class="num">${p[1]}</td><td class="num">${p[2]}</td><td>${esc(p[3])}</td></tr>`; });
  h += "</table></div></div>";

  // (가) 회신 송장 처리 결과 대조
  const already = out.already || 0;
  const alreadyNote = already ? ` · 이미 취합된 송장 ${already}건은 건너뜀(덮어쓰기 안 함)` : "";
  if (out.gap === 0) {
    h += `<div class="msg show ok" style="margin-top:0">✔ 회신 송장 ${out.srcInvoice}건 모두 처리됨 — 신규 기입 <b>${out.writtenInvoice}건</b>${alreadyNote}</div>`;
  } else if (out.gap > 0) {
    h += `<div class="msg show err" style="margin-top:0">⚠ 회신 송장 <b>${out.gap}건</b>이 취합본의 주문과 매칭되지 않았습니다\n(신규 기입 ${out.writtenInvoice}건${already ? ` · 이미취합 ${already}건` : ""} / 회신 ${out.srcInvoice}건)\n\n회신본의 수취인·주소·상품명이 취합본과 다른지 확인하세요.</div>`;
  } else {
    h += `<div class="msg show err" style="margin-top:0">⚠ 취합본 기입(${out.writtenInvoice}건)이 회신 송장(${out.srcInvoice}건)보다 많습니다. 회신 파일 중복을 확인하세요.</div>`;
  }

  // (나) 취합본 빈칸(누락) 점검 — 주문행인데 송장이 안 채워진 행 → 표로 보여줌
  if (out.orderRows !== undefined) {
    if (out.missingCount === 0 && !out.oddCount) {
      h += `<div class="msg show ok" style="margin-top:8px">✔ 취합본 빈칸 없음 — 주문 ${out.orderRows}행 전부 송장 기입 완료</div>`;
    } else if (out.missingCount === 0) {
      h += `<div class="msg show ok" style="margin-top:8px">✔ 취합본 빈칸 없음 — 주문 ${out.orderRows}행</div>`;
    } else {
      h += `<div class="msg show err" style="margin-top:8px">⚠ 취합본 송장 빈칸 <b>${out.missingCount}건</b> / 주문 ${out.orderRows}행 — 아래 주문은 업체 회신에 송장이 없습니다</div>`;
      h += missingTable(out.missing || [], out.missingCount);
    }
  }

  // (나-1) 택배사가 비어 와서 대신 채운 내역 — 조용히 채우면 나중에 왜 그런지 모른다
  if ((out.carrierFills || []).length) {
    h += `<div class="msg show ok" style="margin-top:8px;text-align:left">🚚 택배사가 비어 있던 칸을 채웠습니다
      ${out.carrierFills.map(c => `<div style="padding:3px 0"><b>${esc(c.supplier)}</b> — ${esc(c.carrier)}
        <span style="color:var(--muted)">${c.n}건 · ${esc(c.from)} 기준</span></div>`).join("")}</div>`;
  }

  // (나-2) 송장 자리에 송장이 아닌 값(문장·메모)이 들어간 행
  //        빈칸이 아니라는 이유로 '기입 완료'로 넘어가면 안 된다. 두 줄만 보여주고 나머지는 펼치기.
  if (out.oddCount) {
    const odd = out.odd || [];
    const oddLine = o => {
      const bits = [];
      if (o.badCarrier && o.carrier) bits.push(`택배사: ${esc(o.carrier)}`);
      if (o.badInvoice && o.invoice) bits.push(`송장: ${esc(o.invoice)}`);
      return `<div style="padding:5px 0;border-top:1px solid var(--line)"><b>${o.row}행</b> ${esc(o.label || "")}<br>
        <span style="color:var(--err)">${bits.join(" · ")}</span></div>`;
    };
    const head = odd.slice(0, 2), rest = odd.slice(2);
    h += `<div class="msg show err" style="margin-top:8px;text-align:left">⚠ 송장 형태가 아닌 값 <b>${out.oddCount}건</b> — 택배사·송장 칸에 송장번호 대신 문구가 들어가 있습니다. 아직 출고되지 않은 주문일 수 있으니 확인하세요.
      ${head.map(oddLine).join("")}
      ${rest.length ? `<details style="margin-top:6px"><summary style="cursor:pointer;font-weight:600">나머지 ${rest.length}건 펼쳐보기</summary>${rest.map(oddLine).join("")}</details>` : ""}</div>`;
  }

  // (다) 모호 매칭 — 동일 정보 주문이 여러 개라 어느 행에 넣을지 자동 확정 못한 경우(확인 필요)
  const amb = out.ambiguous || [];
  if (amb.length) {
    h += `<div class="msg show warn" style="margin-top:8px">⚠ 확인 필요 <b>${amb.length}건</b> — 아래는 <b>똑같은 정보(수취인·상품·옵션·수량)의 주문이 2건 이상</b>이라, 어느 주문에 넣을지 자동으로 확정하지 못했습니다. 송장이 올바른 주문에 들어갔는지 확인하세요.</div>`;
    h += `<div class="tblbox" style="margin-top:8px"><div class="tblscroll"><table class="pv"><tr><th>업체</th><th>주문</th><th>옵션</th><th>송장</th><th>동일건</th></tr>`;
    amb.forEach(a => {
      h += `<tr><td>${esc(a.supplier || "")}</td><td>${esc(a.label || "")}</td><td>${esc(a.option || "")}</td><td>${esc(a.inv || "")}</td><td class="num">${a.count || ""}</td></tr>`;
    });
    h += "</table></div></div>";
  }
  h += `<div class="rrow" style="margin-top:12px"><div class="rtop"><div class="vinfo"><b>송장 취합본</b>
    <span>${esc(filename)}</span></div><span class="cnt">${out.total}건</span></div>
    <div class="cands" id="inv-cands"></div>
    <div class="rmail"><input type="text" id="inv-to" placeholder="받는 사람 이메일 (여러 개는 쉼표로)"
      value="${esc(S.invEmails || "")}" inputmode="email" autocapitalize="off" autocorrect="off" spellcheck="false">
      <button class="dlbtn" id="send-inv">메일 보내기</button></div>
    <div class="setrow" style="margin-top:6px"><span style="flex:1;font-size:11px;color:var(--faint)"></span><button class="minibtn" id="share-inv">📤 카톡·공유</button><button class="minibtn" id="pv-inv">미리보기</button><button class="minibtn" id="dl-inv">엑셀 받기</button></div>
    <div class="setrow" style="margin-top:4px"><span style="flex:1;font-size:11px;color:var(--faint)"></span><button class="minibtn" id="tpl-inv">✏️ 메일내용 수정</button></div>
    ${S.sabDrive ? `<button class="go" id="drv-writeback" style="margin-top:10px;font-size:14px;padding:12px;background:var(--ok);color:#fff">📥 드라이브 양식(${esc(S.sabDrive.name)})에 송장 기입</button>
      <div id="drv-wb-msg" style="font-size:11.5px;color:var(--muted);margin-top:6px;text-align:center"></div>` : ""}</div>`;
  $("rlist-i").innerHTML = h;
  $("tpl-inv").onclick = () => MAILTPL.open("invoice", MAILTPL.sampleVars("invoice"));
  $("pv-inv").onclick = () => openPreview(buf, "송장 취합본");
  $("dl-inv").onclick = () => download(buf, filename);
  $("share-inv").onclick = () => shareFile(buf, filename);
  // 드라이브에서 불러온 양식이면 → 그 원본 파일에 결과를 그대로 되쓰기
  if (S.sabDrive && $("drv-writeback")) {
    $("drv-writeback").onclick = async function () {
      this.disabled = true; const orig = this.textContent; this.textContent = "드라이브에 기입 중…";
      $("drv-wb-msg").textContent = "";
      try {
        await ensureGmail();
        const info = await GMAIL.driveUpdateFile(S.sabDrive.id, buf.slice ? buf.slice(0) : buf);
        this.textContent = "✔ 드라이브 양식에 기입 완료";
        const when = info.modifiedTime ? new Date(info.modifiedTime).toLocaleString("ko-KR") : "";
        $("drv-wb-msg").textContent = `${S.sabDrive.name} 파일이 갱신되었습니다${when ? " · " + when : ""}. (드라이브에서 바로 확인하세요)`;
      } catch (e) { this.disabled = false; this.textContent = orig; $("drv-wb-msg").textContent = "⚠ 기입 실패: " + e.message; }
    };
  }
  // 받는사람 추천: 마지막 발송(기본) + 이전 이력 + '발주서 보내는 곳'(주문 메일 발신자) 주소
  //  → 발주서 검색조건의 발신 도메인으로 메일을 찾아 그 발신자 주소를 후보로 띄움
  (async () => {
    let senders = [];
    try { senders = ((await getOrderFilter()).senders || []).filter(Boolean); } catch (e) {}
    const domains = senders.map(s => (s.includes("@") ? s.split("@")[1] : s).toLowerCase());
    const query = senders.length ? "from:(" + senders.join(" OR ") + ")" : null;
    fillRecipients($("inv-cands"), $("inv-to"), { saved: S.invEmails, history: S.invSent, domains, query });
  })();
  $("send-inv").onclick = async function () {
    const list = parseEmails($("inv-to").value);
    if (!list.length) { $("inv-to").focus(); return; }
    const bad = invalidEmails(list);
    if (bad.length) { alert("이메일 형식이 이상해요:\n" + bad.join(", ")); $("inv-to").focus(); return; }
    this.disabled = true; this.textContent = "보내는 중…";
    try {
      await ensureGmail();
      const ymd = QO.todayStr().slice(2);
      const tpl = MAILTPL.render("invoice", { 회사: CONFIG.company, 날짜: ymd, 업체: "" });
      await GMAIL.send({ to: list.join(", "), subject: tpl.subject, body: tpl.body,
        attachments: [{ filename, data: buf }] });
      S.invEmails = $("inv-to").value.trim(); await DB.set("invEmails", S.invEmails);
      await recordSentInv(list);      // 보낸 곳 이력에 기억
      this.textContent = list.length > 1 ? `✓ ${list.length}명 발송완료` : "✓ 발송완료"; this.style.background = "var(--ok)";
    } catch (e) { this.disabled = false; this.textContent = "메일 보내기"; alert("발송 실패: " + e.message); }
  };
  $("result-i").style.display = "block";
  $("result-i").scrollIntoView({ behavior: "smooth", block: "start" });
}

/* =================================================================
   설정 (업체 메일 · 저장 데이터)
   ================================================================= */
$("btn-settings").onclick = () => { drawSettings(); drawSyncStatus(); drawNotifyStatus(); drawLogin(); $("setmodal").classList.add("on"); };
async function drawLogin() {
  const el = $("set-company");
  if (el) {
    let who = ""; try { who = (LOCK.company && LOCK.company()) || ""; } catch (e) {}
    el.textContent = who ? who + " 로 로그인됨" : "로그인 정보";
  }
  /* 마스터 메뉴는 마스터로 들어왔을 때만 보인다.
     업체 담당자에게 관리자 메뉴가 보일 이유가 없다 — 마스터는 로그인 화면으로 들어온다. */
  const box = $("set-master-box");
  if (box) {
    let m = false; try { m = await LOCK.isMasterSession(); } catch (e) {}
    box.style.display = m ? "" : "none";
  }
}
$("set-close").onclick = () => $("setmodal").classList.remove("on");
/* 로그아웃 — 잠금만 풀린 상태를 지운다. 저장된 자료(업체 양식·공급가표)는 그대로 둔다.
   로그인 화면으로 돌아갈 길이 없으면 다른 업체로 바꿔 들어갈 수도, 마스터로 다시 들어갈 수도 없다. */
/* 확인창 없이 바로 나간다 — 자료가 지워지는 것도 아니고, 다시 들어오면 그만이다.
   되묻는 창이 오히려 성가시다는 지적을 받았다. */
function doLogout() {
  try { LOCK.signOut(); } catch (e) {}
  location.reload();
}
if ($("set-logout")) $("set-logout").onclick = doLogout;
if ($("btn-logout")) $("btn-logout").onclick = doLogout;
if ($("btn-sync")) $("btn-sync").onclick = () => { const b = $("sync-now"); if (b) b.click(); };
$("sync-now").onclick = async function () {
  this.disabled = true;
  try {
    await ensureGmail();                 // 로그인 보장(드라이브 권한 포함)
    const r = await SYNC.syncDown();      // 원격이 최신이면 내려받고
    if (r.changed) {
      await loadForms(); drawOrderFilter(); drawReplyFilter(); drawSettings();
      try { if (window.CS) await CS.reload(); } catch (e) {}
    }
    await SYNC.syncUpNow();               // 이 기기 상태도 올려서 최신 유지
  } catch (e) { S.syncState = "error"; S.syncDetail = e.message; drawSyncStatus(); }
  this.disabled = false;
};
$("setmodal").onclick = e => { if (e.target === $("setmodal")) $("setmodal").classList.remove("on"); };
$("set-add").onclick = async () => {
  const name = prompt("업체명"); if (!name) return;
  const mail = prompt(`${name} 이메일 주소`); if (!mail) return;
  S.vendorEmails[name.trim()] = mail.trim();
  await DB.set("vendorEmails", S.vendorEmails); drawSettings();
};
async function drawSettings() {
  const box = $("setlist"); box.innerHTML = "";

  // --- 구글 메일 연결 ---
  const cid = await clientId();          // 저장값 없으면 기본 내장 ID를 보여줌
  const gbox = document.createElement("div");
  gbox.className = "mitem"; gbox.style.marginBottom = "10px";
  const own = cid && cid !== DEFAULT_CLIENT_ID;
  gbox.innerHTML = `<div style="font-weight:700;font-size:13px">📧 구글 메일·드라이브 연결</div>
    <div style="font-size:11px;color:var(--muted);margin:4px 0 8px">메일에서 발주서·송장을 가져오고, 드라이브 파일을 불러오고, 결과를 메일로 보내려면 연결하세요.</div>
    <div style="font-size:11px;margin:0 0 8px;padding:8px 10px;border-radius:9px;background:var(--card2);border:1px solid var(--line);line-height:1.55">
      ${own ? "우리 회사 구글 프로젝트를 쓰는 중입니다." :
              "지금은 <b>제공된 기본 프로젝트</b>를 쓰고 있습니다.<br>" +
              "연결이 <b>403 (액세스 차단됨)</b> 으로 막히면, 그 프로젝트가 아직 '테스트' 상태라 그렇습니다. " +
              "관리자가 구글 콘솔에서 <b>[앱 게시]</b> 를 누르면 바로 풀립니다.<br>" +
              "회사 전용 구글 프로젝트를 따로 쓰려면 그 <b>클라이언트 ID</b> 를 아래에 넣고 [저장] 하세요."}
    </div>
    <div style="font-size:11px;color:var(--muted);margin:0 0 6px">
      연결에 쓰는 구글 프로젝트 번호 <b style="color:var(--ink);font-size:12px">${esc(GMAIL.projectNo(cid))}</b>
      &nbsp;— 막히면 이 번호를 관리자에게 알려주세요
    </div>
    <div class="fld"><label>클라이언트 ID</label>
      <input id="gmail-cid" value="${esc(cid)}" placeholder="xxxxx.apps.googleusercontent.com" spellcheck="false" autocapitalize="off"></div>
    <div style="display:flex;gap:7px">
      <button class="minibtn" id="gmail-save" style="padding:0 12px">저장</button>
      <button class="minibtn" id="gmail-connect" style="padding:0 12px;color:var(--brand)">구글 로그인</button>
      <span id="gmail-status" style="flex:1;font-size:11px;color:var(--muted);align-self:center"></span>
    </div>`;
  box.appendChild(gbox);
  $("gmail-status").textContent = !gmailReady ? "미연결" : (GMAIL.signedIn() ? "로그인됨 ✓" : "준비됨");
  $("gmail-save").onclick = async () => {
    const cid = $("gmail-cid").value.trim();
    const before = await clientId();
    await DB.set("gmailClientId", cid);
    /* 프로젝트가 바뀌면 앞 프로젝트에서 받은 토큰은 쓸 수 없다 — 붙잡고 있으면
       '로그인됨' 으로 보이면서 요청만 실패한다. 깨끗이 끊고 다시 로그인하게 한다. */
    if ((cid || DEFAULT_CLIENT_ID) !== before) { try { GMAIL.signOut(); } catch (e) {} }
    $("gmail-status").textContent = "저장됨 · 라이브러리 준비 중…";
    GMAIL.init(cid || DEFAULT_CLIENT_ID);
    gmailReady = await GMAIL.waitReady();
    updateGmailWho();
    $("gmail-status").textContent = gmailReady ? "✓ 준비됨 · 이제 [구글 로그인]" : "✕ 라이브러리 로드 실패 (새로고침/광고차단 확인)";
  };
  $("gmail-connect").onclick = async () => {
    $("gmail-status").textContent = "로그인 창 여는 중…";
    try { await ensureGmail(); const p = await GMAIL.profile();
      $("gmail-status").textContent = "✓ " + (p.emailAddress || "로그인됨"); updateGmailWho();
    } catch (e) { $("gmail-status").textContent = "✕ 연결 실패"; alert(e.message); }
  };

  const hr = document.createElement("div");
  hr.style.cssText = "border-top:1px solid var(--line);margin:6px 0 12px";
  box.appendChild(hr);

  const title = document.createElement("div");
  title.style.cssText = "font-size:12px;font-weight:700;margin-bottom:8px;color:var(--muted)";
  title.textContent = "업체 이메일";
  box.appendChild(title);

  const names = [...new Set([...Object.keys(S.vendorEmails), ...S.forms.map(f => f.name)])].sort();
  if (!names.length) { const e = document.createElement("div"); e.className = "empty"; e.textContent = "저장된 업체가 없습니다."; box.appendChild(e); return; }
  const istyle = "width:100%;box-sizing:border-box;border:1.5px solid var(--line);background:var(--card2);color:var(--ink);border-radius:9px;padding:9px 10px;font-family:inherit;font-size:13px;outline:none";
  names.forEach(name => {
    const el = document.createElement("div");
    el.className = "mitem";
    el.innerHTML = `<div style="font-weight:700;font-size:13px">🏭 ${esc(name)}</div>
      <div style="font-size:11px;color:var(--muted);margin:7px 0 3px">메일 도메인 <span style="color:var(--faint)">(여러 개는 쉼표)</span></div>
      <input class="vdom" value="${esc(S.vendorDomains[name] || "")}" placeholder="예: onekglobal.co.kr"
        inputmode="url" autocapitalize="off" spellcheck="false" style="${istyle}">
      <div style="font-size:11px;color:var(--muted);margin:8px 0 3px">메일 주소 <span style="color:var(--faint)">(여러 개는 쉼표)</span></div>
      <input class="vadr" type="text" value="${esc(S.vendorEmails[name] || "")}" placeholder="예: manager@onekglobal.co.kr"
        inputmode="email" autocapitalize="off" spellcheck="false" style="${istyle}">
      <div style="display:flex;gap:7px;margin-top:8px">
        <button class="minibtn save" style="padding:0 14px">저장</button>
        <button class="minibtn del" style="padding:0 14px;color:var(--danger)">삭제</button>
      </div>`;
    const dom = el.querySelector(".vdom"), adr = el.querySelector(".vadr");
    const saveBtn = el.querySelector(".save");
    saveBtn.onclick = async () => {
      S.vendorDomains[name] = dom.value.trim(); if (!S.vendorDomains[name]) delete S.vendorDomains[name];
      S.vendorEmails[name] = adr.value.trim();  if (!S.vendorEmails[name]) delete S.vendorEmails[name];
      await DB.set("vendorDomains", S.vendorDomains);
      await DB.set("vendorEmails", S.vendorEmails);
      saveBtn.textContent = "완료"; setTimeout(() => saveBtn.textContent = "저장", 1200);
    };
    el.querySelector(".del").onclick = async () => {
      if (!confirm(`${name} 의 저장된 도메인·주소를 지울까요?`)) return;
      delete S.vendorEmails[name]; delete S.vendorDomains[name];
      await DB.set("vendorEmails", S.vendorEmails); await DB.set("vendorDomains", S.vendorDomains);
      drawSettings();
    };
    box.appendChild(el);
  });
}

/* =================================================================
   Gmail 연동
   ================================================================= */
let gmailReady = false;
async function initGmail() {
  const cid = await clientId();          // 저장값 없으면 기본 내장 ID 사용
  GMAIL.init(cid);                       // 클라이언트 ID 등록(라이브러리 늦어도 됨)
  updateGmailWho();
  if (cid) {
    gmailReady = await GMAIL.waitReady();  // GSI 로드까지 기다렸다 준비
    updateGmailWho();

  }
}

function updateGmailWho() {
  // 어느 계정으로 들어와 있는지 보여준다 — 한 주소를 두 회사가 같이 써서, 이게 안 보이면
  // 남의 회사 자료를 보고 있어도 모른다
  const who = CONFIG.account ? ` (${CONFIG.account})` : "";
  /* 드라이브와 메일은 같은 구글 계정 하나로 연결된다 — '메일' 이라고만 쓰면
     드라이브는 따로 붙여야 하는 줄 안다. 문구를 '구글 계정' 으로 둔다. */
  const txt = !gmailReady ? "⚠ 구글 연결 준비 안 됨 (설정에서 연결하세요)"
    : GMAIL.signedIn() ? "✓ 구글 계정 연결됨 (드라이브·메일)" + who
    : "구글 계정 연결 필요 (버튼을 누르면 로그인)";
  document.querySelectorAll(".gwho").forEach(el => { el.textContent = txt; });
}
/* 구글 계정 바꾸기 — 탭마다 [계정 변경] 버튼에 걸린다.
   계정을 바꾸면 그 계정 저장소로 갈아타야 한다(useAccountStore). 안 그러면
   화면은 새 계정인데 자료는 앞 계정 것을 보게 된다. */
async function switchGoogleAccount(btn) {
  const old = btn ? btn.textContent : "";
  try {
    if (btn) { btn.disabled = true; btn.textContent = "여는 중…"; }
    if (!gmailReady) gmailReady = await GMAIL.waitReady();
    if (!gmailReady) throw new Error("구글 로그인 라이브러리를 불러오지 못했어요.\n새로고침 후 다시 시도하세요.");
    await GMAIL.switchAccount();
    let email = "";
    try { email = ((await GMAIL.profile()).emailAddress || "").toLowerCase(); } catch (e) {}
    /* ★ 순서가 중요하다. 로그인은 '앞 계정' 자리에서 끝나므로,
       ① 그 자리에 잘못 저장된 토큰을 지우고 ② 저장소를 새 계정 것으로 바꾼 뒤
       ③ 토큰을 새 자리에 다시 써넣는다. 안 그러면 다음에 열 때 앞 계정 드라이브가 보인다. */
    try { GMAIL.dropStored(); } catch (e) {}
    if (email) await useAccountStore(email, { quiet: true });
    try { GMAIL.persistToken(); } catch (e) {}
    updateGmailWho();
    msg("msg-o", "ok", email ? `👤 ${email} 계정으로 바꿨어요.` : "구글 계정을 바꿨어요.");
  } catch (e) {
    alert(e.message || String(e));
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = old || "계정 변경"; }
  }
}
function bindAccountButtons() {
  document.querySelectorAll(".gacct").forEach(b => { b.onclick = () => switchGoogleAccount(b); });
}
/* ※ 로그인 팝업은 '사용자 클릭' 안에서 열려야 브라우저가 막지 않는다.
   그래서 토큰 요청 전에 await(DB 읽기 등)를 하지 않도록, 클라이언트 ID는 시작할 때 미리 받아둔다. */
async function ensureGmail() {
  if (GMAIL.signedIn()) { updateGmailWho(); return; }   // 로그인 유효 → await 없이 바로 사용
  if (!gmailReady) {                                     // 준비 안 된 경우에만 기다림
    gmailReady = await GMAIL.waitReady(); updateGmailWho();
  }
  if (!gmailReady) throw new Error("구글 로그인 라이브러리를 불러오지 못했어요.\n인터넷/광고차단을 확인하고 새로고침 해보세요.");
  await GMAIL.token();    // 클릭 맥락에서 팝업 → 이미 승인했으면 잠깐 떴다 자동으로 닫힘
  GMAIL.profile().catch(() => {});   // 계정 이메일을 힌트로 저장(다음 재로그인 때 계정선택 생략)
  updateGmailWho();
  syncOnStart();          // 로그인 직후 다른 기기 데이터 내려받기
}

/* ★ 검색조건 기본값은 '이 배포본의 원래 회사' 기준으로 적혀 있다.
   한 주소를 여러 회사가 나눠 쓰므로, 다른 업체로 로그인하면 그 업체 이름으로 만들어 준다.
   (베타브릭스로 들어갔는데 '랩노마드 발주서' 가 기본 키워드로 보이던 문제)
   ※ 저장해 둔 값이 있으면 그게 우선이다 — 여기는 아직 아무것도 안 정한 경우만 쓴다. */
function defaultOrderFilter() {
  const co = String(CONFIG.company || "").trim();
  const home = String(CONFIG.homeCompany || "").trim().replace(/\s+/g, "");
  if (!co || co.replace(/\s+/g, "") === home) return CONFIG.order;
  return {
    senders: [],                                   // 발주서가 오는 곳은 회사마다 다르다
    keywords: [co + " 발주서", co + "발주서", "★" + co, co],
    exclude: ["송장", "회신", "운송장", "택배"],    // 업체명은 빼고 공통 단어만
  };
}
/* 발주서 검색조건 (PC 앱과 동일한 기본값) */
async function getOrderFilter() {
  const d = defaultOrderFilter();
  return {
    senders: await DB.get("orderSenders", d.senders),
    keywords: await DB.get("orderKeywords", d.keywords),
    exclude: await DB.get("orderExclude", d.exclude),
  };
}
async function drawOrderFilter() {
  const f = await getOrderFilter();
  const info = $("order-filter");
  if (info) info.textContent = "";
}

/* 회신 검색조건 */
async function getReplyFilter() {
  return {
    senders: await DB.get("replySenders", CONFIG.reply.senders),
    keywords: await DB.get("replyKeywords", CONFIG.reply.keywords),
    exclude: await DB.get("replyExclude", CONFIG.reply.exclude),
  };
}
async function drawReplyFilter() {
  const f = await getReplyFilter();
  const el = $("reply-filter");
  if (el) el.textContent = "";
}

/* ---------- 검색조건 관리 모달 (발주서/회신 공용) ---------- */
const filterModal = $("filtermodal");
let filterMode = "order";   // 'order' | 'reply'
// 각 목록의 저장 키
const FKEY = {
  order: { senders: "orderSenders", keywords: "orderKeywords", exclude: "orderExclude" },
  reply: { senders: "replySenders", keywords: "replyKeywords", exclude: "replyExclude" },
  cs: { senders: "csSenders", keywords: "csKeywords", exclude: "csExclude" },
  settle: { senders: "stSenders", keywords: "stKeywords", exclude: "stExclude" },
  pb: { senders: "pbSenders", keywords: "pbKeywords", exclude: "pbExclude" },
  pay: { senders: "paySenders", keywords: "payKeywords", exclude: "payExclude" },
};
/* 발주서와 송장취합양식은 같은 파일(사방넷 통합본)을 찾으므로 조건을 함께 쓴다 —
   제목에 둘 다 적어야 어느 쪽에서 열어도 헷갈리지 않는다. */
const FTITLE = { order: "발주서·송장취합양식 검색조건", reply: "회신 송장 검색조건", cs: "CS 검색조건",
                 settle: "정산 파일 검색조건", pb: "공급가표 검색조건", pay: "대금지급 내역 검색조건" };

async function openFilter(mode) {
  filterMode = mode;
  $("filter-title").textContent = FTITLE[mode] || "검색조건";
  await renderFilterLists();
  filterModal.classList.add("on");
}
async function renderFilterLists() {
  drawChipList("flt-senders", await getList("senders"), "senders");
  drawChipList("flt-keywords", await getList("keywords"), "keywords");
  drawChipList("flt-excludes", await getList("exclude"), "exclude");
}
function drawChipList(boxId, items, kind) {
  const box = $(boxId); box.innerHTML = "";
  if (!items.length) { const e = document.createElement("div"); e.className = "flt-none"; e.textContent = "(없음)"; box.appendChild(e); return; }
  items.forEach((val, i) => {
    const el = document.createElement("div");
    el.className = "flt-item";
    el.innerHTML = `<span>${esc(val)}</span><button class="ed">수정</button><button class="rm">삭제</button>`;
    el.querySelector(".ed").onclick = () => editFilterItem(kind, i);
    el.querySelector(".rm").onclick = () => removeFilterItem(kind, i);
    box.appendChild(el);
  });
}
/* 검색조건 기본값 — 저장해 둔 값이 없을 때만 쓴다 */
function filterDefaults() {
  return {
    order: defaultOrderFilter(),
    reply: CONFIG.reply,
    cs: { senders: [], keywords: ["문의", "교환", "반품", "취소", "환불", "누락", "파손", "불량", "CS"],
          exclude: ["발주", "정산"] },
    settle: { senders: [], keywords: ["정산", "정산내역", "지급"], exclude: ["발주", "송장"] },
    pb: { senders: [], keywords: ["공급가", "단가", "상품리스트", "정산참고"], exclude: [] },
    pay: { senders: [], keywords: ["대금", "지급", "매입", "정산"], exclude: [] },
  };
}
async function getList(kind) {
  const key = FKEY[filterMode][kind];
  const DEF = filterDefaults();
  const def = (DEF[filterMode] || DEF.order)[kind] || [];
  return await DB.get(key, def);
}
/* 화면 밖(정산 등)에서 그 종류의 검색조건을 통째로 읽어갈 때 쓴다 */
async function getFilter(mode) {
  const k = FKEY[mode] ? mode : "order";
  const def = filterDefaults()[k] || {};
  const out = {};
  for (const which of ["senders", "keywords", "exclude"]) {
    out[which] = await DB.get(FKEY[k][which], def[which] || []);
  }
  return out;
}
async function setList(kind, arr) {
  await DB.set(FKEY[filterMode][kind], arr);
  await renderFilterLists();
  drawOrderFilter(); drawReplyFilter();
  try { if (window.CS) CS.drawFilter(); } catch (e) {}
  try { if (window.ST) ST.drawFilter(); } catch (e) {}
}
async function addFilterItem(kind, inputId) {
  const inp = $(inputId), val = inp.value.trim();
  if (!val) return;
  const arr = await getList(kind);
  if (!arr.includes(val)) arr.push(val);
  inp.value = "";
  await setList(kind, arr);
}
async function editFilterItem(kind, i) {
  const arr = await getList(kind);
  const v = prompt("수정", arr[i]);
  if (v === null) return;
  const t = v.trim();
  if (!t) return;
  arr[i] = t;
  await setList(kind, arr);
}
async function removeFilterItem(kind, i) {
  const arr = await getList(kind);
  arr.splice(i, 1);
  await setList(kind, arr);
}
$("flt-sender-btn").onclick = () => addFilterItem("senders", "flt-sender-in");
$("flt-keyword-btn").onclick = () => addFilterItem("keywords", "flt-keyword-in");
$("flt-exclude-btn").onclick = () => addFilterItem("exclude", "flt-exclude-in");
$("flt-sender-in").onkeydown = e => { if (e.key === "Enter") { e.preventDefault(); addFilterItem("senders", "flt-sender-in"); } };
$("flt-keyword-in").onkeydown = e => { if (e.key === "Enter") { e.preventDefault(); addFilterItem("keywords", "flt-keyword-in"); } };
$("flt-exclude-in").onkeydown = e => { if (e.key === "Enter") { e.preventDefault(); addFilterItem("exclude", "flt-exclude-in"); } };
$("filter-close").onclick = () => filterModal.classList.remove("on");
filterModal.onclick = e => { if (e.target === filterModal) filterModal.classList.remove("on"); };
$("order-filter-btn") && ($("order-filter-btn").onclick = () => openFilter("order"));
$("reply-filter-btn") && ($("reply-filter-btn").onclick = () => openFilter("reply"));

/* 메일 선택 모달 */
const mailModal = $("mailmodal");
let mailItems = [], mailSel = [], mailMulti = false, mailTarget = null, mailHidden = 0;
$("mail-cancel").onclick = () => mailModal.classList.remove("on");
mailModal.onclick = e => { if (e.target === mailModal) mailModal.classList.remove("on"); };

let mailDays = 1;   // 기본: 오늘
// 기간 버튼
document.querySelectorAll("#mail-period button").forEach(b => {
  b.onclick = () => {
    mailDays = Number(b.dataset.d) || 1;
    document.querySelectorAll("#mail-period button").forEach(x => x.classList.toggle("on", x === b));
    if (mailModal.classList.contains("on")) loadMail();   // 열려 있으면 즉시 다시 검색
  };
});
async function openMail(target) {
  mailTarget = target;                       // 'order' | 'sab' | 'rep'
  mailMulti = (target === "rep");
  mailSel = [];
  try { await ensureGmail(); } catch (e) { return; }
  mailModal.classList.add("on");
  $("mail-ok").disabled = true;
  $("mail-ok").textContent = mailMulti ? "선택 항목 가져오기" : "이 파일 가져오기";
  const titles = { order: "메일에서 발주서 가져오기", sab: "메일에서 송장취합양식 가져오기", rep: "메일에서 회신 송장 가져오기" };
  $("mail-title").textContent = titles[target];
  document.querySelectorAll("#mail-period button")
    .forEach(x => x.classList.toggle("on", Number(x.dataset.d) === mailDays));
  await loadMail();
}
async function loadMail() {
  const target = mailTarget;
  const dayTxt = mailDays === 1 ? "오늘" : `최근 ${mailDays}일`;
  const list = $("mail-list");
  mailSel = []; $("mail-ok").disabled = true;
  list.innerHTML = `<div class="empty">${dayTxt} 메일함을 확인하고 있어요…<br><span id="mail-prog"></span></div>`;
  $("mail-sub").textContent = `${dayTxt} 메일 확인 중…`;
  try {
    let opt;
    if (target === "rep") {
      const f = await getReplyFilter();
      opt = { days: mailDays, senders: f.senders, keywords: f.keywords, exclude: f.exclude || [], union: true, scanText: true };
    } else {
      // 발주서/사방넷: 저장된 발신자·키워드·제외어로 선별 (PC 앱과 동일)
      const f = await getOrderFilter();
      opt = { days: mailDays, senders: f.senders, keywords: f.keywords, exclude: f.exclude, union: false, scanText: true };
    }
    opt.onProgress = (i, n) => {
      const p = $("mail-prog"); if (p) p.textContent = `${i} / ${n}`;
      BUSY.progress(i, n);                     // 로딩창 막대·퍼센트
    };
    mailItems = await GMAIL.listMails(opt);
    /* 회신 송장은 '보낸사람 또는 키워드' 로 찾는다(union). 그래서 본문에 '회신' 같은 말이
       들어간 남의 메일까지 딸려온다. 업체 주소를 알고 있으면 그것만 남긴다.
       — 등록된 업체 이메일 · 도메인 · 예전에 회신이 온 주소가 기준이다.
       하나도 안 남으면(주소를 아직 안 넣었거나 새 업체) 원래 목록을 그대로 보여준다. */
    if (target === "rep") {
      const addr = new Set(), dom = new Set();
      const add = t => String(t || "").split(/[,;\s]+/).forEach(x => {
        const e = x.trim().toLowerCase(); if (e.includes("@")) addr.add(e);
      });
      Object.values(S.vendorEmails || {}).forEach(add);
      Object.values(S.vendorSent || {}).forEach(v => (Array.isArray(v) ? v : [v]).forEach(add));
      Object.values(S.vendorDomains || {}).forEach(v => (Array.isArray(v) ? v : String(v || "").split(/[,;\s]+/))
        .forEach(d => { const t = String(d || "").trim().toLowerCase().replace(/^@/, ""); if (t) dom.add(t); }));
      if (addr.size || dom.size) {
        const known = m => {
          const f = (m.from.includes("<") ? m.from.split("<").pop().replace(">", "") : m.from).trim().toLowerCase();
          if (addr.has(f)) return true;
          const d = f.split("@")[1] || "";
          return d && [...dom].some(x => d === x || d.endsWith("." + x));
        };
        const hit = mailItems.filter(known);
        if (hit.length) { mailHidden = mailItems.length - hit.length; mailItems = hit; }
        else mailHidden = 0;
      } else mailHidden = 0;
    } else mailHidden = 0;
    if (!mailItems.length) {
      list.innerHTML = `<div class="empty">${dayTxt}간 해당 엑셀 첨부를 찾지 못했어요 — 위에서 기간을 늘려보세요</div>`;
      $("mail-sub").textContent = "결과 없음"; return;
    }
    $("mail-sub").textContent = (mailMulti ? "여러 개 선택 가능 · " : "하나 선택 · ") + mailItems.length + "건"
      + (mailHidden ? ` · 업체 주소가 아닌 메일 ${mailHidden}건은 숨김` : "");
    list.innerHTML = "";
    mailItems.forEach((m, i) => {
      const frm = m.from.includes("<") ? m.from.split("<").pop().replace(">", "") : m.from;
      const isToday = m.ts && new Date(m.ts).toDateString() === new Date().toDateString();
      const el = document.createElement("div");
      el.className = "mitem";
      el.innerHTML = `<div style="display:flex;align-items:center;gap:6px;margin-bottom:5px">
          <span style="font-weight:800;font-size:14px;color:var(--brand)">🕑 ${esc(m.date)}</span>
          ${isToday ? `<span style="font-size:11px;font-weight:800;color:#fff;background:#e8384f;padding:2px 8px;border-radius:999px">금일 수신</span>` : ""}
        </div>
        <div style="font-weight:700;font-size:13px;word-break:break-all">📄 ${esc(m.filename)}</div>
        <div style="font-size:11.5px;color:var(--muted);margin-top:3px">${esc(m.subject || "(제목 없음)")}</div>
        <div style="font-size:11px;color:var(--faint);margin-top:1px">${esc(frm)}</div>
        ${m.body && m.body.trim() ? `<div style="font-size:11.5px;color:var(--muted);margin-top:7px;padding-top:7px;border-top:1px dashed var(--line);white-space:pre-wrap;max-height:80px;overflow:auto">${esc(m.body.trim())}</div>` : ""}`;
      el.onclick = () => {
        if (mailMulti) {
          el.classList.toggle("on");
          el.style.borderColor = el.classList.contains("on") ? "var(--brand)" : "";
          el.style.background = el.classList.contains("on") ? "var(--brand-soft)" : "";
          const k = mailSel.indexOf(i); if (k >= 0) mailSel.splice(k, 1); else mailSel.push(i);
        } else {
          [...list.children].forEach(c => { c.style.borderColor = ""; c.style.background = ""; });
          el.style.borderColor = "var(--brand)"; el.style.background = "var(--brand-soft)";
          mailSel = [i];
        }
        $("mail-ok").disabled = !mailSel.length;
      };
      list.appendChild(el);
    });
  } catch (e) {
    list.innerHTML = `<div class="empty">⚠ ${esc(e.message)}</div>`;
    $("mail-sub").textContent = "오류";
  }
}
$("mail-ok").onclick = async function () {
  if (!mailSel.length) return;
  this.disabled = true; this.textContent = "가져오는 중…";
  try {
    const got = [];
    for (const i of mailSel) {
      const m = mailItems[i];
      const buf = await GMAIL.getAttachment(m.id, m.attachmentId);
      got.push({ name: m.filename, data: buf });
    }
    mailModal.classList.remove("on");
    if (mailTarget === "order") {
      await setOrderFromBuf(got[0].data, got[0].name);
      msg("msg-o", "ok", "✔ 메일에서 가져왔어요: " + got[0].name);
    } else if (mailTarget === "sab") {
      S.sabBuf = got[0].data; S.sabName = got[0].name;
      S.sabDrive = null;
      $("sab-name").textContent = "📧 " + got[0].name; $("drop-sab").classList.add("on"); $("sab-preview").style.display="block"; refreshI();
    } else {
      for (const g of got) if (!S.reps.some(r => r.name === g.name)) S.reps.push(g);
      drawReps();
      msg("msg-i", "ok", `✔ 메일에서 ${got.length}개 가져왔어요.`);
    }
  } catch (e) {
    alert("가져오기 실패: " + e.message);
  } finally { this.disabled = false; this.textContent = mailMulti ? "선택 항목 가져오기" : "이 파일 가져오기"; }
};
$("mail-order").onclick = () => openMail("order");
$("mail-sab").onclick = () => openMail("sab");
if ($("sab-filter-btn")) $("sab-filter-btn").onclick = () => openFilter("order");
$("mail-rep").onclick = () => openMail("rep");

// setOrder 를 버퍼 기반으로도 쓰도록 분리
async function setOrderFromBuf(buf, name) {
  msg("msg-o", "", "");
  S.orderBuf = buf; S.orderName = name;
  $("order-name").textContent = "📧 " + name;
  $("drop-order").classList.add("on");
  const wb = await QO.loadWorkbook(S.orderBuf.slice(0));
  S.brands = QO.listBrands(wb);
  S.dateSel = [];
  await loadDates(); await drawPreview();
  buildVendorBrands(); refreshO();
}

/* =================================================================
   새 발주·송장 알림 (앱이 열려 있을 때 주기적으로 확인 → 알림)
   ================================================================= */
const notifyEnabled = () => { try { return localStorage.getItem(CONFIG.ls("qo_notifyOn")) === "1"; } catch (e) { return false; } };
const setNotifyEnabled = v => { try { localStorage.setItem(CONFIG.ls("qo_notifyOn"), v ? "1" : "0"); } catch (e) {} };
const NOTIFY_MS = 3 * 60 * 1000;   // 3분마다
let notifyTimer = null;

function fireNotify(title, body) {
  try {
    if (window.Notification && Notification.permission === "granted")
      new Notification(title, { body, icon: "icon-192.png", tag: "qo-" + title });
  } catch (e) {}
}
function showNotifyBanner(items) {
  const el = $("notify-banner"); if (!el) return;
  const first = items[0];
  el.innerHTML = `${esc(first.title)}<small>${esc(first.body)}${items.length > 1 ? ` 외 ${items.length - 1}건` : ""} · 눌러서 보기</small>`;
  el.classList.add("show");
  el.onclick = () => { el.classList.remove("show"); if (first.tab) switchTab(first.tab); };
  clearTimeout(el._t); el._t = setTimeout(() => el.classList.remove("show"), 12000);
}
function drawNotifyStatus() {
  const on = notifyEnabled();
  $("notify-toggle").textContent = on ? "끄기" : "켜기";
  const perm = (window.Notification && Notification.permission) || "unsupported";
  $("notify-status").textContent = !on ? "꺼져 있음"
    : (perm === "granted" ? "✓ 켜짐 — 앱을 열어두면 새 발주·송장을 알려드려요"
    : perm === "denied" ? "⚠ 브라우저 알림이 차단됨 — 앱 안 배너로만 표시됩니다"
    : "켜짐 — 알림 권한을 허용하면 배너+알림 둘 다 떠요");
}
$("notify-toggle").onclick = async function () {
  if (this._busy) return; this._busy = true;          // 더블탭/중복실행 방지
  try {
    const on = !notifyEnabled();                        // 현재값 반대로
    setNotifyEnabled(on);                               // 먼저 저장(동기) → 화면 즉시 반영
    drawNotifyStatus();
    if (on) {
      if (window.Notification && Notification.permission === "default") {
        try { await Notification.requestPermission(); } catch (e) {}
        drawNotifyStatus();
      }
      startNotify(); notifyTick(false);
    } else { stopNotify(); }
  } finally { setTimeout(() => { this._busy = false; }, 400); }
};
function startNotify() { if (!notifyTimer) notifyTimer = setInterval(() => notifyTick(false), NOTIFY_MS); }
function stopNotify() { if (notifyTimer) { clearInterval(notifyTimer); notifyTimer = null; } }

async function notifyTick(manual) {
  if (!notifyEnabled()) return;
  if (!GMAIL.signedIn()) return;      // 로그인돼 있을 때만
  const hits = [];
  // ① 지정한 드라이브 발주 파일이 바뀌었나 (수정시각 비교)
  try {
    const df = await DB.get("driveOrderFile", null);
    if (df && df.id) {
      const info = await GMAIL.driveFileInfo(df.id);
      const last = await DB.get("notifyDriveMTime", "");
      if (info.modifiedTime) {
        if (last && info.modifiedTime !== last)
          hits.push({ title: "발주 내역 업데이트", body: `${df.name} 파일이 변경됐어요`, tab: "o", tag: "발주 내역 업데이트" });
        await DB.set("notifyDriveMTime", info.modifiedTime);
      }
    }
  } catch (e) {}
  // ② 지정 업체에서 송장 회신 메일이 새로 왔나
  try {
    const f = await getReplyFilter();
    const items = await GMAIL.listMails({ days: 2, senders: f.senders, keywords: f.keywords, exclude: f.exclude || [], union: true, scanText: true, max: 20 });
    const seen = new Set(await DB.get("notifySeenMails", []));
    const fresh = items.filter(m => !seen.has(m.id));
    if (fresh.length) {
      if (seen.size)   // 처음 켠 직후엔 기존 메일로 알림 폭탄 안 나게, 기준선만 잡음
        hits.push({ title: "송장 회신 메일 도착", body: `${fresh.length}건 — ${fresh[0].subject || fresh[0].filename}`, tab: "i", tag: "송장 회신 메일 도착" });
      const merged = [...new Set(items.map(m => m.id).concat([...seen]))].slice(0, 120);
      await DB.set("notifySeenMails", merged);
    }
  } catch (e) {}
  hits.forEach(h => fireNotify(h.title, h.body));
  if (hits.length) showNotifyBanner(hits);
  else if (manual) showNotifyBanner([{ title: "새 소식 없음", body: "지금은 변경·회신이 없어요", tab: "" }]);
}

/* ---------------- 동기화 (구글 드라이브) ---------------- */
// 데이터가 바뀔 때마다 자동 업로드(디바운스)
DB.onChange(() => SYNC.pushSoon());
// 동기화 상태를 설정화면 등에 반영
SYNC.onStatus((state, detail) => { S.syncState = state; S.syncDetail = detail || ""; drawSyncStatus(); });
function fmtAgo(ts) {
  if (!ts) return "아직 없음";
  const s = Math.floor((Date.now() - ts) / 1000);
  if (s < 60) return "방금";
  if (s < 3600) return Math.floor(s / 60) + "분 전";
  if (s < 86400) return Math.floor(s / 3600) + "시간 전";
  return Math.floor(s / 86400) + "일 전";
}
function drawSyncStatus() {
  const el = $("sync-status"); if (!el) return;
  const st = S.syncState;
  let t;
  if (st === "syncing") t = "🔄 " + (S.syncDetail || "동기화 중…");
  else if (st === "error") t = "⚠ 동기화 오류: " + (S.syncDetail || "");
  else if (st === "offline") t = "구글 로그인하면 자동 동기화됩니다";
  else t = "✓ 동기화됨 · 마지막 " + fmtAgo(SYNC.lastTime());
  el.textContent = t;
  /* 업체 양식 카드에도 같은 상태를 보여준다 — 모바일에서 '왜 안 넘어왔지' 를
     설정까지 열어보지 않고 바로 알 수 있어야 한다. */
  /* 헤더 동기화 버튼도 상태를 보여준다 — 눌러도 되는지, 도는 중인지 알 수 있게 */
  const b = $("btn-sync");
  if (b) {
    b.disabled = st === "syncing";
    b.style.opacity = st === "syncing" ? ".6" : "";
    b.title = st === "offline" ? "구글 로그인하면 다른 기기와 자동으로 맞춰집니다" : t;
  }
}
// 로그인돼 있으면 시작 시 내려받기 → 바뀌었으면 화면 갱신
/* 로그인한 구글 계정으로 저장소를 갈아탄다.
   한 주소를 여러 회사가 같이 쓰기 때문에, 계정이 바뀌면 저장소도 바뀌어야 한다.
   갈아탄 뒤에는 그 계정의 드라이브 백업에서 업체 양식·설정을 그대로 내려받는다. */
/* 마지막 로그인 계정 — 업체별로 따로 기억한다.
   한 브라우저를 두 회사가 나눠 쓰면, 앞 회사가 쓰던 구글 계정으로 저장소를 열어 버린다. */
const ACCT_KEY = () => CONFIG.lsCompany("qo_last_account");
/* 저장소 이름이 바뀌었을 때 화면까지 다시 그린다 (업체가 바뀌거나 계정이 바뀔 때) */
async function reopenStore(note) {
  await DB.reopen();
  await loadForms();
  try { if (window.CS) await CS.reload(); } catch (e) {}
  try { if (window.ST) await ST.reload(); } catch (e) {}
  drawOrderFilter(); drawReplyFilter(); drawDriveRecent(); drawSabRecent();
  if (note) msg("msg-o", "ok", note);
}
async function useAccountStore(email, opts) {
  if (!email) return false;
  try { localStorage.setItem(ACCT_KEY(), email); } catch (e) {}
  const changed = CONFIG.useAccount(email);
  if (!changed) return false;
  await reopenStore((opts && opts.quiet) ? "" :
    `👤 ${email} 계정으로 바꿨어요. 이 계정의 자료를 불러옵니다.`);
  return true;
}
/* 로그인이 확인되면 계정을 확인해 저장소를 맞춘다 */
async function syncAccount() {
  if (!GMAIL.signedIn()) return false;
  let email = "";
  try { email = ((await GMAIL.profile()).emailAddress || "").toLowerCase(); } catch (e) { return false; }
  return await useAccountStore(email);
}
async function syncOnStart() {
  try {
    await syncAccount();                      // 저장소를 먼저 맞추고
    const r = await SYNC.syncDown();
    if (r.changed) {
      await loadForms(); drawOrderFilter(); drawReplyFilter();
      if ($("setmodal").classList.contains("on")) drawSettings();
      try { if (window.CS) await CS.reload(); } catch (e) {}
    }
    // 클라우드에 백업이 아직 없고, 이 기기에 데이터가 있으면 최초 1회 올려서 씨딩
    // (데이터 없는 기기는 올리지 않음 → 빈 상태로 다른 기기를 덮어쓰지 않게)
    else if (r.hadRemote === false && S.forms.length) { await SYNC.syncUpNow(); }
  } catch (e) {}
  drawSyncStatus();
}

/* =====================================================================
   마스터 메뉴 — 승인 업체 관리
   마스터 아이디로 들어왔을 때만 보인다.
   승인코드는 업체명에서 계산된다(SALT + 업체명). 그래서 목록은 '누구를 승인했는지'
   적어두는 장부이고, 실제 승인은 그 코드를 업체에 전달하는 순간 이뤄진다.
   ※ 이미 나간 코드는 목록에서 지워도 계속 유효하다 — 완전히 막으려면 배포에 반영해야 한다.
   ===================================================================== */
const MST = (() => {
  let list = [];
  /* clipboard API 는 보안 컨텍스트가 아니거나 권한이 없으면 조용히 실패한다.
     복사가 됐는지 사용자는 알 수 없으니, 실패하면 옛 방식으로 한 번 더 시도하고
     그래도 안 되면 값을 화면에 띄워 직접 고르게 한다. */
  async function copyText(txt, okMsg) {
    const say = m => { const el = $("mst-msg"); if (el) el.textContent = m; };
    try { await navigator.clipboard.writeText(txt); return say(okMsg); } catch (e) {}
    try {
      const ta = document.createElement("textarea");
      ta.value = txt; ta.style.cssText = "position:fixed;top:0;left:0;opacity:0";
      document.body.appendChild(ta); ta.focus(); ta.select();
      const ok = document.execCommand("copy");
      document.body.removeChild(ta);
      if (ok) return say(okMsg);
    } catch (e) {}
    say("복사가 막혀 있어요. 아래 값을 직접 선택해 복사하세요:\n" + txt);
  }
  /* 예전엔 업체명 문자열만 담았다. 이제 사용여부·연락처까지 담아야 해서 객체로 바꾼다.
     옛 데이터를 읽을 때 그 자리에서 객체로 올려준다 (이미 쓰고 있는 사람이 있다). */
  const norm = x => (typeof x === "string" ? { name: x, on: true } : Object.assign({ on: true }, x));
  async function load() { list = ((await DB.get("masterCompanies", [])) || []).map(norm); }
  async function save() { await DB.set("masterCompanies", list); }
  /* 업체 목록을 공용 명단 형태로 — 나중에 드라이브·배포에 올려 로그인 차단에 쓴다 */
  function roster() {
    return { at: Date.now(), companies: list.map(c => ({ name: c.name, on: c.on !== false })) };
  }
  async function draw() {
    const box = $("mst-list"); if (!box) return;
    if (!list.length) { box.innerHTML = '<div class="empty">승인한 업체가 없습니다</div>'; return; }
    const codes = await Promise.all(list.map(c => LOCK.approvalCode(c.name)));
    box.innerHTML = list.map((c, i) => {
      const on = c.on !== false;
      return `<div style="display:flex;align-items:center;gap:8px;padding:10px 12px;
        border:1px solid var(--line);border-radius:10px;margin-bottom:6px;
        background:var(--card2);opacity:${on ? 1 : .55}">
        <div style="flex:1;min-width:0">
          <b style="font-size:13.5px;word-break:break-all">${esc(c.name)}</b>
          ${on ? "" : '<span style="font-size:11px;font-weight:800;color:var(--danger);margin-left:6px">사용 중지</span>'}
          <div style="font-size:17px;font-weight:800;letter-spacing:2px;color:var(--brand);margin-top:2px">${codes[i]}</div>
          ${c.email ? `<div style="font-size:11px;color:var(--muted);margin-top:2px">${esc(c.email)}${c.tel ? " · " + esc(c.tel) : ""}</div>` : ""}
        </div>
        <button class="minibtn mstonoff" data-i="${i}" style="flex:none;${on ? "" : "color:var(--danger);border-color:var(--danger)"}">${on ? "사용 중" : "중지됨"}</button>
        <button class="minibtn mstcopy" data-i="${i}" style="flex:none">번호 복사</button>
        <button class="minibtn mstmsg" data-i="${i}" style="flex:none">안내문</button>
        <button class="minibtn mstdel" data-i="${i}" style="flex:none">삭제</button></div>`;
    }).join("");
    box.querySelectorAll(".mstonoff").forEach(b => b.onclick = async () => {
      const i = Number(b.dataset.i);
      list[i].on = list[i].on === false;
      await save(); draw();
      $("mst-msg").textContent = list[i].on
        ? `${list[i].name} — 사용으로 바꿨습니다.`
        : `${list[i].name} — 사용 중지로 표시했습니다. ※ 아직 기록만 남습니다 — 실제 로그인 차단은 공용 명단 연결이 필요합니다.`;
    });
    /* ★ '복사' 는 승인번호만 복사한다.
       안내 문구까지 같이 복사하면 그대로 붙여넣었을 때 로그인이 안 된다 (실제로 그랬다). */
    box.querySelectorAll(".mstcopy").forEach(b => b.onclick = () => {
      const i = Number(b.dataset.i);
      copyText(codes[i], `승인번호 ${codes[i]} 복사했어요. 그대로 붙여넣으면 로그인됩니다.`);
    });
    box.querySelectorAll(".mstmsg").forEach(b => b.onclick = async () => {
      const i = Number(b.dataset.i);
      const c = list[i];
      /* 메일 주소를 알고 있으면 바로 보낸다 — 복사해서 옮겨 붙이는 단계를 없앤다 */
      if (c.email) { await sendGuide(c, codes[i]); return; }
      copyText(guideText(c.name, codes[i]), "안내문을 복사했어요. 업체에 그대로 보내세요.");
    });
    box.querySelectorAll(".mstdel").forEach(b => b.onclick = async () => {
      const i = Number(b.dataset.i);
      if (!confirm(`'${list[i].name}' 를 목록에서 지울까요?
(이미 전달한 승인번호는 계속 쓸 수 있습니다)`)) return;
      list.splice(i, 1); await save(); draw();
    });
  }
  /* 마스터인지 저장된 표시로 판단하면, 기기를 바꾸거나 표시가 지워졌을 때 들어갈 길이 막힌다.
     그래서 표시가 없으면 모달 안에서 아이디·비밀번호를 받아 그 자리에서 확인한다. */
  let authedNow = false;                       // 이번에 아이디·비번으로 확인한 경우
  async function isMasterNow() {
    if (authedNow) return true;
    try { return await LOCK.isMasterSession(); } catch (e) { return false; }
  }
  const APP_URL = location.origin + location.pathname;
  function guideText(name, code) {
    return `[퀵오더] ${name} 로그인 안내

주소: ${APP_URL}
업체명: ${name}
승인번호: ${code}

로그인 화면에서 업체명과 승인번호를 넣으시면 됩니다.
승인번호가 곧 비밀번호입니다. 다른 곳에 공유하지 마세요.`;
  }
  /* 안내문을 그 업체 메일로 바로 보낸다. 마스터는 구글에 로그인돼 있으므로 그 계정으로 나간다. */
  async function sendGuide(c, code) {
    const say = m => { const el = $("mst-msg"); if (el) el.textContent = m; };
    if (!c.email) { say("이 업체는 메일 주소가 없어요. 신청서로 들어온 업체만 바로 보낼 수 있습니다."); return; }
    if (!confirm(`${c.email} 로 승인 안내문을 보낼까요?`)) return;
    say("보내는 중…");
    try {
      await ensureGmail();
      await GMAIL.send({ to: c.email, subject: `[퀵오더] ${c.name} 로그인 안내`, body: guideText(c.name, code) });
      c.sentAt = Date.now(); await save(); draw();
      say(`✔ ${c.email} 로 보냈습니다.`);
    } catch (e) { say("⚠ 보내지 못했어요 — " + (e.message || e)); }
  }
  /* ── 가입 신청함 ─────────────────────────────────────────────
     서버가 없으니 신청은 메일로 온다. 마스터 지메일에서 신청 메일을 읽어 목록으로 보여준다.
     제목에 표식(SIGNUP_TAG)을 넣어 두고 그걸로 찾는다. */
  const SIGNUP_TAG = "[퀵오더 가입신청]";
  function parseReq(text) {
    const g = k => {
      const m = String(text || "").match(new RegExp("^\s*" + k + "\s*[:：]\s*(.+)$", "m"));
      return m ? m[1].trim() : "";
    };
    return { name: g("업체명"), email: g("이메일"), tel: g("연락처") };
  }
  async function loadReqs() {
    const box = $("mst-reqs"); if (!box) return;
    box.innerHTML = '<div class="empty">신청함을 확인하는 중…</div>';
    try {
      await ensureGmail();
      const items = await GMAIL.listTextMails({ days: 90, keywords: [SIGNUP_TAG], senders: [], max: 30 });
      const reqs = [];
      for (const m of items) {
        const r = parseReq(m.text || m.snippet || "");
        if (!r.name) continue;
        if (reqs.some(x => x.name === r.name)) continue;      // 같은 업체가 여러 번 보내면 하나만
        r.at = m.date || ""; reqs.push(r);
      }
      drawReqs(reqs);
    } catch (e) {
      box.innerHTML = `<div class="empty">신청함을 읽지 못했어요 — ${esc(e.message || String(e))}</div>`;
    }
  }
  function drawReqs(reqs) {
    const box = $("mst-reqs"); if (!box) return;
    const fresh = reqs.filter(r => !list.some(c => c.name === r.name));
    if (!fresh.length) { box.innerHTML = '<div class="empty">새 가입 신청이 없습니다</div>'; return; }
    box.innerHTML = fresh.map((r, i) => `<div style="display:flex;align-items:center;gap:8px;padding:10px 12px;
        border:1.5px solid var(--brand);border-radius:10px;margin-bottom:6px;background:var(--brand-soft)">
        <div style="flex:1;min-width:0">
          <b style="font-size:13.5px">${esc(r.name)}</b>
          <div style="font-size:11.5px;color:var(--muted);margin-top:2px">${esc(r.email || "메일 없음")}${r.tel ? " · " + esc(r.tel) : ""}</div>
        </div>
        <button class="minibtn reqok" data-i="${i}" style="flex:none;color:var(--brand);border-color:var(--brand);font-weight:800">승인</button>
        <button class="minibtn reqsend" data-i="${i}" style="flex:none">승인 + 안내문 발송</button></div>`).join("");
    const add = async r => {
      if (list.some(c => c.name === r.name)) return;
      list.push({ name: r.name, on: true, email: r.email || "", tel: r.tel || "", at: Date.now() });
      await save(); await draw();
    };
    box.querySelectorAll(".reqok").forEach(b => b.onclick = async () => {
      const r = fresh[Number(b.dataset.i)];
      await add(r); drawReqs(reqs);
      $("mst-msg").textContent = `${r.name} 승인했습니다. 안내문은 목록의 [안내문] 으로 보낼 수 있어요.`;
    });
    box.querySelectorAll(".reqsend").forEach(b => b.onclick = async () => {
      const r = fresh[Number(b.dataset.i)];
      await add(r); drawReqs(reqs);
      const c = list.find(x => x.name === r.name);
      if (c) await sendGuide(c, await LOCK.approvalCode(c.name));
    });
  }
  async function open() {
    wire();                                    // ★ 열 때마다 다시 연결한다 (아래 wire 주석 참고)
    $("mstmodal").classList.add("on");
    if (!await isMasterNow()) return authMode();
    $("mst-auth").style.display = "none"; $("mst-body").style.display = "";
    $("mst-sub").textContent = "승인할 업체명을 넣으면 승인번호가 나옵니다. 그 번호가 그 업체의 비밀번호입니다.";
    await load(); $("mst-msg").textContent = ""; await draw();
  }
  function authMode() {
    $("mst-auth").style.display = ""; $("mst-body").style.display = "none";
    $("mst-sub").textContent = "마스터 아이디와 비밀번호를 넣으세요.";
    $("mst-err").textContent = ""; $("mst-pw").value = "";
    /* 안 될 때 짐작하지 않도록, 지금 돌고 있는 파일이 어느 버전인지 보여준다 */
    const d = $("mst-diag");
    if (d) d.textContent = "v" + (typeof APP_VER !== "undefined" ? APP_VER : "?") +
      " · 확인 " + (typeof LOCK.signInMaster === "function" ? "새 방식" :
                   typeof LOCK.isMaster === "function" ? "옛 방식" : "없음");
    setTimeout(() => { try { ($("mst-id").value ? $("mst-pw") : $("mst-id")).focus(); } catch (e) {} }, 80);
  }
  /* 실패 이유를 뭉뚱그리지 않는다 — 비밀번호가 틀린 것과 파일이 옛것인 것은 대처가 다르다 */
  async function signIn() {
    const id = $("mst-id").value, pw = $("mst-pw").value;
    const err = $("mst-err"); err.textContent = "확인 중…";
    let ok = false, why = "";
    try {
      if (typeof LOCK.signInMaster === "function") ok = await LOCK.signInMaster(id, pw);
      else if (typeof LOCK.isMaster === "function") ok = await LOCK.isMaster(id, pw);   // 옛 qo-lock.js 라도 들어갈 수 있게
      else why = "앱 파일이 옛것입니다. [파일 새로 받기] 를 눌러 주세요.";
    } catch (e) { why = "확인 중 오류 — " + (e && e.message ? e.message : e); }
    if (!ok) {
      const m = why || "아이디 또는 비밀번호가 틀렸습니다.";
      err.textContent = m; $("mst-pw").value = "";
      try { alert(m); } catch (e) {}
      try { $("mst-pw").focus(); } catch (e) {} return;
    }
    authedNow = true; $("mst-pw").value = ""; err.textContent = ""; show(); await open();
  }
  /* 서비스워커와 캐시를 통째로 버리고 다시 받는다 — 파일이 섞여 있을 때의 마지막 수단 */
  async function hardReload() {
    try {
      if (navigator.serviceWorker) {
        const rs = await navigator.serviceWorker.getRegistrations();
        await Promise.all(rs.map(r => r.unregister()));
      }
      if (window.caches) { const ks = await caches.keys(); await Promise.all(ks.map(k => caches.delete(k))); }
    } catch (e) {}
    location.reload(true);
  }
  /* 헤더 배지 — 마스터면 👑 마스터(눌러서 관리), 업체면 업체명 */
  async function drawWho() {
    const b = $("who-badge"); if (!b) return;
    let master = false; try { master = await isMasterNow(); } catch (e) {}
    if (master) {
      b.textContent = "👑 마스터"; b.classList.add("master");
      b.style.display = "inline-block"; b.onclick = open; return;
    }
    let who = ""; try { who = (LOCK.company && LOCK.company()) || ""; } catch (e) {}
    b.classList.remove("master"); b.onclick = null;
    if (!who) { b.style.display = "none"; return; }
    b.textContent = who; b.style.display = "inline-block";
  }
  const show = drawWho;

  /* ★ 버튼 연결은 '한 번만' 하면 안 된다.
     하나라도 없는 요소를 만나면 거기서 예외가 나 뒤쪽 버튼이 통째로 죽는데,
     시작할 때 딱 한 번 도는 구조라 그러면 영영 안 먹는다 (확인 버튼 무반응이 그 경우).
     그래서 요소마다 따로 감싸고, 모달을 열 때마다 다시 연결한다. */
  const on = (id, ev, fn) => { try { const el = $(id); if (el) el[ev] = fn; } catch (e) {} };
  function wire() {
    on("set-master", "onclick", () => { $("setmodal").classList.remove("on"); open(); });
    on("mst-in", "onclick", signIn);
    on("mst-hard", "onclick", hardReload);
    on("mst-pw", "onkeydown", e => { if (e.key === "Enter") signIn(); });
    on("mst-id", "onkeydown", e => { if (e.key === "Enter") $("mst-pw").focus(); });
    on("mst-close", "onclick", () => $("mstmodal").classList.remove("on"));
    on("mstmodal", "onclick", e => { if (e.target === $("mstmodal")) $("mstmodal").classList.remove("on"); });
    on("mst-add", "onclick", async () => {
      const n = $("mst-name").value.trim().replace(/\s+/g, "");
      if (!n) { $("mst-name").focus(); return; }
      if (list.some(c => c.name === n)) { $("mst-msg").textContent = "이미 목록에 있어요."; return; }
      list.push({ name: n, on: true, at: Date.now() });
      await save(); $("mst-name").value = ""; $("mst-msg").textContent = ""; draw();
    });
    on("mst-name", "onkeydown", e => { if (e.key === "Enter") $("mst-add").onclick(); });
    on("mst-req-refresh", "onclick", loadReqs);
    on("mst-roster", "onclick", () => {
      const txt = JSON.stringify(roster(), null, 2);
      download(new TextEncoder().encode(txt).buffer, "roster.json", "application/json");
      $("mst-msg").textContent =
        "roster.json 을 내려받았습니다.\n이 파일을 앱과 같은 폴더(배포)에 올리면 '중지됨' 업체의 로그인이 실제로 막힙니다.";
    });
  }
  return { bind: wire, open, show };
})();
/* ★ 로그인 화면이 끝난 다음에 확인해야 한다.
   페이지가 뜨는 순간 확인하면, 그 자리에서 마스터로 로그인해도 이미 확인이 끝나 있어서
   버튼이 안 나타난다 (새로고침해야 보였다). LOCK.ready 를 기다린다. */
try { MST.bind(); } catch (e) {}            // 설정 안의 입구는 항상 살아 있어야 한다

/* ---------------- 시작 ---------------- */
/* 로그인할 때 넣은 업체명을 회사 이름으로 쓴다.
   한 주소를 여러 회사가 같이 쓰는데 회사 이름이 URL 에 박혀 있으면,
   베타브릭스가 들어와도 발주서 파일명·메일에 랩노마드가 나간다. */
/* ★ 업체 이름은 화면 표시용이 아니라 '저장소를 가르는 기준' 이다.
   이걸 저장소에 반영하지 않으면, 한 주소를 나눠 쓰는 두 회사가 같은 자료를 본다
   (랩노마드가 올려둔 업체 양식이 베타브릭스로 로그인해도 그대로 보였다).
   저장소 이름이 바뀌었으면 true 를 돌려준다. */
/* 화면에 회사 이름이 들어가는 자리는 <span class="co-name"> 로 표시해 두고 여기서 채운다.
   HTML 에 이름을 박아 두면, 다른 업체로 로그인해도 앞 회사 이름이 그대로 보인다. */
function drawCoName() {
  try {
    const n = CONFIG.company || "";
    if (!n) return;
    document.querySelectorAll(".co-name").forEach(el => { el.textContent = n; });
  } catch (e) {}
}
function applyLockCompany() {
  try {
    const c = (typeof LOCK !== "undefined" && LOCK.company && LOCK.company()) || "";
    if (c) { CONFIG.company = c; if (!CONFIG.orderTag) CONFIG.orderTag = c; }
    return CONFIG.useCompany(c);
  } catch (e) { return false; }
}
/* ★★ 로그인이 끝나기 전에는 아무것도 읽지 않는다.
   예전에는 페이지가 뜨자마자 loadForms() 와 구글 동기화를 시작했다. 그런데 그 시점에는
   아직 앞사람(랩노마드)의 저장소가 열려 있어서,
     · 화면에 앞 회사의 업체 양식이 뜨고,
     · 진행 중이던 동기화가 앞 회사 백업을 내려받아 새 회사 저장소에 그대로 써 넣고,
     · 구글 토큰이 메모리에 남아 새 회사가 앞 회사 계정으로 연결된 것처럼 보였다.
   그래서 시작 전체를 LOCK.ready 뒤로 미룬다. */
(async () => {
  applyLockCompany();                       // 이미 로그인돼 있으면 이 값이 곧 정답이다
  try { if (typeof LOCK !== "undefined") await LOCK.ready; } catch (e) {}
  applyLockCompany();                       // 방금 로그인했다면 여기서 확정된다
  try { const last = localStorage.getItem(ACCT_KEY()); if (last) CONFIG.useAccount(last); } catch (e) {}
  try { GMAIL.reloadToken(); } catch (e) {}   // 앞사람 토큰이 메모리에 남아 있지 않게
  /* ★ 저장소 연결도 다시 연다.
     qo-cs.js 같은 모듈이 로드되면서 이미 DB 를 열어 둔다. 그 연결은 로그인 전 이름으로
     열린 것이라, 이름만 바꿔서는 소용이 없다 — 읽고 쓰는 곳은 여전히 앞 회사 저장소다. */
  try { await DB.reopen(); } catch (e) {}
  try { if (window.CS && CS.reload) await CS.reload(); } catch (e) {}
  try { if (window.ST && ST.reload) await ST.reload(); } catch (e) {}
  try { await loadForms(); }
  catch (e) { $("vlist").innerHTML = '<div class="empty">저장소를 열지 못했어요</div>'; }
  initGmail();
  drawReplyFilter();
  drawOrderFilter();
  drawDriveRecent();
  drawSabRecent();
  drawCoName();                             // 화면에 박힌 회사 이름을 로그인한 업체로
  bindAccountButtons();                     // 탭마다 [계정 변경]
  try { await MST.show(); } catch (e) {}    // 헤더 배지 (👑 마스터 / 업체명)
  syncOnStart();                            // 저장소가 확정된 뒤에 동기화
})();
if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => {});
// 알림: 켜져 있으면 폴링 시작, 앱으로 돌아올 때마다 즉시 한 번 확인
if (notifyEnabled()) { startNotify(); setTimeout(() => notifyTick(false), 4000); }
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible" && notifyEnabled()) notifyTick(false);
});

/* ---- 당겨서 새로고침 ---- 화면 맨 위에서 아래로 당기면 새로고침한다.
   새로고침하면 메일·PC·드라이브로 불러온 파일(메모리 상태)이 전부 초기화된다.
   (드라이브 '바로 가져오기' 지정·업체 양식 등 저장된 설정은 유지) */
(function () {
  const ptr = $("ptr"), txt = $("ptr-txt");
  if (!ptr) return;
  const MAX = 80, TRIG = 60;
  let startY = 0, pulling = false, dist = 0;
  const atTop = () => (window.scrollY || document.documentElement.scrollTop || 0) <= 0;
  window.addEventListener("touchstart", e => {
    if (e.touches.length === 1 && atTop()) { startY = e.touches[0].clientY; pulling = true; dist = 0; }
    else pulling = false;
  }, { passive: true });
  window.addEventListener("touchmove", e => {
    if (!pulling) return;
    const dy = e.touches[0].clientY - startY;
    if (dy <= 0 || !atTop()) { dist = 0; ptr.style.transition = "none"; ptr.style.height = "0px"; if (dy <= 0) pulling = false; return; }
    dist = Math.min(dy * 0.5, MAX);
    ptr.style.transition = "none";
    ptr.style.height = dist + "px";
    txt.textContent = dist >= TRIG ? "↑ 놓으면 새로고침" : "↓ 당겨서 새로고침";
    if (dy > 6 && e.cancelable) e.preventDefault();   // 네이티브 바운스 억제
  }, { passive: false });
  function end() {
    if (!pulling) return;
    pulling = false;
    ptr.style.transition = "height .12s ease";
    if (dist >= TRIG) {
      ptr.classList.add("spin");
      ptr.style.height = MAX + "px";
      txt.textContent = "새로고침 중…";
      setTimeout(() => location.reload(), 150);
    } else { ptr.style.height = "0px"; }
  }
  window.addEventListener("touchend", end, { passive: true });
  window.addEventListener("touchcancel", end, { passive: true });
})();
