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
    const pv = QO.previewAny(wb, 2000);
    $("pv-modal-sub").textContent = `시트: ${esc(pv.sheet)} · 전체 ${pv.total}건` +
      (pv.sheets.length > 1 ? ` · (${pv.sheets.length}개 시트 중)` : "");
    if (!pv.columns.length) { $("pv-modal-foot").textContent = "표시할 내용이 없어요."; return; }
    if (pv.total === 0) {
      $("pv-modal-sub").textContent = `시트: ${esc(pv.sheet)} · 빈 양식(내용 없음)`;
      $("pv-modal-foot").innerHTML = "ℹ️ 이 파일은 <b>빈 양식(템플릿)</b>이라 채워진 내용이 없습니다. 아래는 열(항목) 목록입니다.";
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
    if (pv.total > 0) {   // 빈 양식일 때는 위의 안내문(ℹ️)을 덮어쓰지 않는다
      $("pv-modal-foot").textContent = pv.total > pv.rows.length
        ? `앞 ${pv.rows.length}건만 표시 · 전체 ${pv.total}건` : `전체 ${pv.total}건`;
    }
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
const readFile = f => new Promise((res, rej) => {
  const r = new FileReader();
  r.onload = () => res(r.result); r.onerror = () => rej(r.error);
  r.readAsArrayBuffer(f);
});

/* =====================================================================
   메일 문구 — 발주서·송장·정산·CS 가 같이 쓴다.
   기본 문구를 두되 사용자가 고치면 기억한다. {회사}·{업체} 처럼 중괄호로 값이 들어간다.
   ===================================================================== */
const MAILTPL = (() => {
  const DEF = {
    order:   { name: "발주서", subject: "[{회사}] {날짜}_발주서 송부",
               body: "안녕하세요 발주서 송부드립니다. 감사합니다!",
               vars: ["회사", "업체", "날짜", "건수"] },
    invoice: { name: "송장 취합본", subject: "[{회사}] {날짜}_송장 취합본 송부",
               body: "안녕하세요 송장 취합본 송부드립니다. 감사합니다!",
               vars: ["회사", "업체", "날짜"] },
    settle:  { name: "정산서", subject: "[{회사}] {정산월} 정산서 - {업체}",
               body: "안녕하세요, {회사}입니다.\n\n{정산월} 정산 내역을 보내드립니다. (작성일 {날짜})\n\n{요약}\n\n자세한 내역은 첨부 파일을 확인해주세요.\n감사합니다.",
               vars: ["회사", "업체", "정산월", "날짜", "건수", "수량", "지급액", "요약"] },
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
  /* vendor 를 주면 그 업체 전용 문구를 먼저 쓴다 (없으면 공통) */
  function render(kind, vars, vendor) {
    const t = get(kind, vendor !== undefined ? vendor : (vars && vars.업체));
    return { subject: fill(t.subject, vars), body: fill(t.body, vars) };
  }

  let cur = null;
  function open(kind, sample, vendor, after) {
    const v = (vendor || "").trim();
    cur = { kind, vendor: v, sample: sample || {}, after };
    const d = DEF[kind], t = get(kind, v);
    $("tpl-title").textContent = v ? `${v} — ${d.name} 메일 내용` : `${d.name} 메일 내용 (공통)`;
    $("tpl-sub").textContent = v
      ? `${v} 에게 보낼 때만 이 문구를 씁니다. (되돌리기를 누르면 공통 문구로 돌아갑니다)`
      : "여기서 고친 제목·본문을 앞으로 계속 씁니다.";
    $("tpl-subject").value = t.subject;
    $("tpl-body").value = t.body;
    $("tpl-vars").innerHTML = "쓸 수 있는 값: " +
      d.vars.map(v => `<b>{${esc(v)}}</b>`).join(" · ") + " — 보낼 때 실제 값으로 바뀝니다";
    preview();
    $("tplmodal").classList.add("on");
  }
  function preview() {
    if (!cur) return;
    const s = fill($("tpl-subject").value, cur.sample);
    const b = fill($("tpl-body").value, cur.sample);
    $("tpl-preview").textContent = "미리보기\n제목: " + s + "\n\n" + b;
  }
  function close() { $("tplmodal").classList.remove("on"); cur = null; }

  $("tpl-subject").oninput = $("tpl-body").oninput = preview;
  $("tpl-close").onclick = close;
  $("tplmodal").onclick = e => { if (e.target === $("tplmodal")) close(); };
  /* 되돌리기 — 업체별 창이면 공통 문구로, 공통 창이면 기본 문구로 */
  $("tpl-reset").onclick = () => {
    if (!cur) return;
    const base = cur.vendor ? get(cur.kind) : DEF[cur.kind];
    $("tpl-subject").value = base.subject;
    $("tpl-body").value = base.body;
    preview();
  };
  $("tpl-save").onclick = async () => {
    if (!cur) return;
    const val = { subject: $("tpl-subject").value, body: $("tpl-body").value };
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
    const sample = Object.assign(MAILTPL.sampleVars(kind), { 업체: vendor }, extraVars || {});
    MAILTPL.open(kind, sample, vendor, () => { btn.innerHTML = mailtplLabel(kind, vendor); });
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
const DRV = { multiple: false, onPick: null, path: [], sel: new Map(), home: null };
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
$("drv-done").onclick = async () => { if (DRV.sel.size) await drvPick([...DRV.sel.values()]); };

/* opts: { key, title, sub, multiple, onPick(files) } — files: [{id,name,mimeType}]
   key: 용도별 기본 폴더 저장용 (order/tpl/sab/rep) → 다음부터 그 폴더가 바로 열림 */
async function openDrivePicker(opts) {
  DRV.multiple = !!opts.multiple; DRV.onPick = opts.onPick; DRV.sel = new Map();
  DRV.key = opts.key || "";
  $("drv-title").textContent = opts.title || "구글 드라이브에서 가져오기";
  $("drv-sub").textContent = opts.sub || (opts.multiple
    ? "폴더 안에서 파일을 여러 개 고를 수 있어요." : "폴더 안에서 파일을 고르세요.");
  $("drv-done").style.display = opts.multiple ? "" : "none";
  $("drv-done").textContent = "선택 완료";
  $("drv-msg").textContent = ""; $("drv-q").value = ""; $("drv-link").value = "";
  $("drv-list").innerHTML = "";
  $("drvmodal").classList.add("on");
  drvFolderInfo();
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
/* 최상위(홈) 폴더 — 회사 작업 폴더가 드라이브 한참 안쪽에 있어서,
   매번 '내 드라이브'부터 파고들지 않게 여기를 바닥으로 삼는다.
   홈을 정해두면 '상위' 버튼도 홈 위로는 안 올라간다. 용도(key)와 무관하게 하나만 쓴다. */
const DRV_ROOT = { id: "root", name: "내 드라이브" };
const drvBase = () => (DRV.home && DRV.home.id ? { id: DRV.home.id, name: DRV.home.name } : Object.assign({}, DRV_ROOT));
/* 홈이 정해져 있으면 조상 사슬을 홈에서 잘라낸다 (홈 밖의 폴더면 그대로 둔다) */
function drvTrim(chain) {
  const path = [Object.assign({}, DRV_ROOT)].concat(chain || []);
  if (!(DRV.home && DRV.home.id)) return path;
  const i = path.findIndex(p => p.id === DRV.home.id);
  return i >= 0 ? path.slice(i) : path;
}
/* 시작 위치: ①고정한 기본 폴더 → ②마지막에 고른 폴더 → ③최상위(홈) 폴더 → ④내 드라이브 */
async function drvStart() {
  const all = await DB.get("driveFolders", {});
  DRV.home = all[":home"] || null;
  const pinned = DRV.key ? all[DRV.key] : null;
  const last = DRV.key ? all[DRV.key + ":last"] : null;
  const go = (pinned && pinned.id) ? pinned : (last && last.id ? last : (DRV.home && DRV.home.id ? DRV.home : null));
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
  drvFolderInfo();
}
async function drvFolderInfo() {
  const all = await DB.get("driveFolders", {});
  DRV.home = all[":home"] || null;
  const saved = DRV.key ? all[DRV.key] : null;
  const bits = [];
  bits.push(DRV.home && DRV.home.id ? `최상위: ${DRV.home.name}` : "최상위 미지정");
  if (saved && saved.id) bits.push(`기본 폴더: ${saved.name}`);
  $("drv-folder-info").textContent = bits.join(" · ");
}
/* 지금 보고 있는 폴더 — 최상위/기본 폴더로 지정할 수 있는지 확인해서 돌려준다 */
function drvCurFolder() {
  const cur = DRV.path[DRV.path.length - 1];
  if (!cur || cur.id === "root" || cur.id === "shared") {
    $("drv-msg").textContent = "⚠ 폴더를 하나 열고 눌러주세요 (내 드라이브 최상위는 지정 불가)";
    return null;
  }
  return cur;
}
$("drv-sethome").onclick = async () => {
  const all = await DB.get("driveFolders", {});
  const cur = DRV.path[DRV.path.length - 1];
  // 이미 그 폴더가 최상위인데 또 누르면 해제 (해제 버튼을 따로 두지 않으려고)
  if (DRV.home && cur && DRV.home.id === cur.id) {
    delete all[":home"];
    await DB.set("driveFolders", all);
    DRV.home = null; drvFolderInfo(); drvCrumb();
    $("drv-up").style.display = DRV.path.length > 1 ? "" : "none";
    $("drv-msg").textContent = "최상위 지정을 풀었어요. 이제 내 드라이브까지 올라갈 수 있습니다.";
    return;
  }
  const f = drvCurFolder(); if (!f) return;
  all[":home"] = { id: f.id, name: f.name };
  await DB.set("driveFolders", all);
  DRV.home = all[":home"];
  DRV.path = drvTrim(DRV.path.slice(1));      // 홈보다 위는 잘라낸다
  drvCrumb();
  $("drv-up").style.display = DRV.path.length > 1 ? "" : "none";
  drvFolderInfo();
  $("drv-msg").textContent = `✔ 최상위 폴더로 저장했어요: ${f.name}\n앞으로 모든 탭에서 여기부터 시작합니다.`;
};
$("drv-setfolder").onclick = async () => {
  if (!DRV.key) return;
  const f = drvCurFolder(); if (!f) return;
  const all = await DB.get("driveFolders", {});
  all[DRV.key] = { id: f.id, name: f.name };
  await DB.set("driveFolders", all);
  drvFolderInfo();
  $("drv-msg").textContent = `✔ 기본 폴더로 저장했어요: ${f.name}`;
};
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
        $("drv-done").textContent = DRV.sel.size ? `선택 완료 (${DRV.sel.size}개)` : "선택 완료";
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
   (둘러보기만 한 폴더까지 기억하면, 잠깐 다른 데를 봤을 뿐인데 시작 위치가 바뀐다) */
async function drvRememberFolder() {
  if (!DRV.key) return;
  const cur = DRV.path[DRV.path.length - 1];
  if (!cur || cur.id === "root" || cur.id === "shared") return;
  const all = await DB.get("driveFolders", {});
  all[DRV.key + ":last"] = { id: cur.id, name: cur.name };
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
  key: "tpl", title: "드라이브에서 업체 양식 가져오기", multiple: true,
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
    msg("msg-o", "ok", `✔ 드라이브에서 업체 양식 ${r.added + r.updated}개 저장 — ${r.names.join(", ")}`
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
  $("dt-foot").textContent = S.dateSel.length
    ? `선택한 ${S.dateSel.length}개 날짜 · 총 ${cnt}건만 변환됩니다`
    : "⚠ 날짜를 하나 이상 선택하세요";
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
  $("pv-foot").textContent = (selSet && hasDates)
    ? `체크한 날짜의 주문 ${view.length}건 — 이 내용이 그대로 변환됩니다`
    : `전체 ${view.length}건 표시`;
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
      if (clash && !confirm(`'${clash}' 업체 양식이 이미 있습니다.\n새 파일로 바꿀까요?`)) return;
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
    await DB.putForm(rec);
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
  msg("msg-o", "ok", `✔ 업체 양식 ${r.added + r.updated}개 저장 — ${r.names.join(", ")}`
    + (r.updated ? ` (${r.updated}개는 기존 양식 갱신)` : "")
    + "\n다음부터는 체크만 하면 됩니다.");
}

/* 업체명 바꾸기 — 브랜드 학습·업체 메일 이력도 같이 옮긴다 */
async function renameForm(f) {
  const v = prompt("업체명을 바꿉니다.\n(발주서 파일명과 브랜드 배정에 쓰이는 이름)", f.name);
  if (v === null) return;
  const newName = v.trim().replace(/[\\/:*?"<>|]/g, "");
  if (!newName || newName === f.name) return;
  if (S.forms.some(x => x.name === newName)) { alert("같은 이름의 업체 양식이 이미 있어요."); return; }
  await DB.delForm(f.name);
  await DB.putForm({ name: newName, file: f.file, data: f.data, checked: f.checked !== false });
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
function drawForms() {
  const box = $("vlist");
  if (!S.forms.length) { box.innerHTML = '<div class="empty">저장된 업체 양식이 없습니다.<br>아래에서 추가하세요.</div>'; return; }
  box.innerHTML = "";
  addDupWarning(box, dupNameGroups(S.forms, f => f.name), f => f.name,
    "브랜드 선택에도 두 번 나옵니다. 안 쓰는 쪽은 ✕ 로 지우거나 ‘이름수정’ 으로 구분해 주세요.");
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
        await DB.delForm(f.name); await loadForms(); return;
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
    const out = QO.collectInvoices(sab, replies);
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
$("btn-settings").onclick = () => { drawSettings(); drawSyncStatus(); drawNotifyStatus(); $("setmodal").classList.add("on"); };
$("set-close").onclick = () => $("setmodal").classList.remove("on");
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
  gbox.innerHTML = `<div style="font-weight:700;font-size:13px">📧 구글 메일 연결</div>
    <div style="font-size:11px;color:var(--muted);margin:4px 0 8px">메일에서 발주서·송장을 가져오고, 결과를 메일로 보내려면 연결하세요.</div>
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
    await DB.set("gmailClientId", cid);
    $("gmail-status").textContent = "저장됨 · 라이브러리 준비 중…";
    GMAIL.init(cid);
    gmailReady = await GMAIL.waitReady();
    updateGmailWho();
    $("gmail-status").textContent = gmailReady ? "✓ 준비됨 · 이제 [구글 로그인]" : "✕ 라이브러리 로드 실패 (새로고침/광고차단 확인)";
  };
  $("gmail-connect").onclick = async () => {
    $("gmail-status").textContent = "로그인 창 여는 중…";
    try { await ensureGmail(); const p = await GMAIL.profile();
      $("gmail-status").textContent = "✓ " + (p.emailAddress || "로그인됨"); updateGmailWho();
    } catch (e) { $("gmail-status").textContent = "✕ " + e.message; alert(e.message); }
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
  const txt = !gmailReady ? "⚠ 메일 연결 준비 안 됨 (설정에서 연결하세요)"
    : GMAIL.signedIn() ? "✓ 구글 메일 연결됨" : "구글 계정 연결 필요 (버튼을 누르면 로그인)";
  const a = $("gmail-who-o"); if (a) a.textContent = txt;
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

/* 발주서 검색조건 (PC 앱과 동일한 기본값) */
async function getOrderFilter() {
  return {
    senders: await DB.get("orderSenders", CONFIG.order.senders),
    keywords: await DB.get("orderKeywords", CONFIG.order.keywords),
    exclude: await DB.get("orderExclude", CONFIG.order.exclude),
  };
}
async function drawOrderFilter() {
  const f = await getOrderFilter();
  const info = $("order-filter");
  if (info) info.textContent = "발주서 검색: " + (f.senders.join(", ") || "(발신자 없음)") +
    " · 키워드 " + (f.keywords.join(", ") || "(없음)");
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
  if (el) el.textContent = "회신 검색: " + (f.senders.length ? f.senders.join(", ") + " · " : "") + (f.keywords.join(", ") || "(없음)");
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
};
const FTITLE = { order: "발주서 검색조건", reply: "회신 송장 검색조건", cs: "CS 검색조건", settle: "정산 파일 검색조건" };

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
async function getList(kind) {
  const key = FKEY[filterMode][kind];
  const DEF = {
    order: CONFIG.order,
    reply: CONFIG.reply,
    cs: { senders: [], keywords: ["문의", "교환", "반품", "취소", "환불", "누락", "파손", "불량", "CS"],
          exclude: ["발주", "정산"] },
    settle: { senders: [], keywords: ["정산", "정산내역", "지급"], exclude: ["발주", "송장"] },
  };
  const def = (DEF[filterMode] || DEF.order)[kind] || [];
  return await DB.get(key, def);
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
let mailItems = [], mailSel = [], mailMulti = false, mailTarget = null;
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
    opt.onProgress = (i, n) => { const p = $("mail-prog"); if (p) p.textContent = `${i} / ${n}`; };
    mailItems = await GMAIL.listMails(opt);
    if (!mailItems.length) {
      list.innerHTML = `<div class="empty">${dayTxt}간 해당 엑셀 첨부를 찾지 못했어요.<br>위에서 기간을 늘려보세요.</div>`;
      $("mail-sub").textContent = "결과 없음"; return;
    }
    $("mail-sub").textContent = (mailMulti ? "여러 개 선택 가능 · " : "하나 선택 · ") + mailItems.length + "건";
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
}
// 로그인돼 있으면 시작 시 내려받기 → 바뀌었으면 화면 갱신
async function syncOnStart() {
  try {
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

/* ---------------- 시작 ---------------- */
loadForms()
  .then(() => syncOnStart())
  .catch(e => { $("vlist").innerHTML = '<div class="empty">저장소를 열지 못했어요</div>'; });
initGmail();
drawReplyFilter();
drawOrderFilter();
drawDriveRecent();
drawSabRecent();
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
