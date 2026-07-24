/* ===================================================================
   퀵오더 앱 — 화면 동작 + 기기 저장(IndexedDB)
   =================================================================== */
"use strict";
// 구글 클라이언트 ID (기본 내장) — github 주소에서만 작동하게 묶여 있어 공개돼도 안전.
// 기기·주소가 바뀌어도 다시 입력할 필요가 없다.
const DEFAULT_CLIENT_ID = "598124965893-16qej37hhlah9ivtr9hdk76c50ms5aqs.apps.googleusercontent.com";
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
      const rq = indexedDB.open("quickorder", 1);
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

/* ---------------- 탭 ---------------- */
$("tab-o").onclick = () => switchTab("o");
$("tab-i").onclick = () => switchTab("i");
function switchTab(t) {
  $("pane-o").classList.toggle("on", t === "o");
  $("pane-i").classList.toggle("on", t === "i");
  $("tab-o").classList.toggle("on", t === "o");
  $("tab-i").classList.toggle("on", t === "i");
  window.scrollTo(0, 0);
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
const DRV = { multiple: false, onPick: null, path: [], sel: new Map() };
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
/* 시작 위치: ①고정한 기본 폴더 → ②마지막에 본 폴더 → ③내 드라이브 최상위 */
async function drvStart() {
  const all = DRV.key ? await DB.get("driveFolders", {}) : {};
  const pinned = DRV.key ? all[DRV.key] : null;
  const last = DRV.key ? all[DRV.key + ":last"] : null;
  const go = (pinned && pinned.id) ? pinned : (last && last.id ? last : null);
  if (go) {
    // 실제 드라이브 상위 폴더들을 따라 경로를 만든다 → '상위' 버튼이 제대로 동작
    DRV.path = [{ id: "root", name: "내 드라이브" }, { id: go.id, name: go.name }];
    drvOpen(go.id);
    GMAIL.driveAncestors(go.id).then(chain => {
      if (chain && chain.length) {
        DRV.path = [{ id: "root", name: "내 드라이브" }].concat(chain);
        drvCrumb();
        $("drv-up").style.display = DRV.path.length > 1 ? "" : "none";
      }
    }).catch(() => {});
  } else {
    DRV.path = [{ id: "root", name: "내 드라이브" }];
    drvOpen("root");
  }
  drvFolderInfo();
}
async function drvFolderInfo() {
  const saved = DRV.key ? (await DB.get("driveFolders", {}))[DRV.key] : null;
  $("drv-folder-info").textContent = saved && saved.id
    ? `기본 폴더: ${saved.name}`
    : "기본 폴더 없음 — 폴더를 연 뒤 [기본 폴더로]를 누르면 다음부터 바로 열립니다";
}
$("drv-setfolder").onclick = async () => {
  if (!DRV.key) return;
  const cur = DRV.path[DRV.path.length - 1];
  if (!cur || cur.id === "root") { $("drv-msg").textContent = "⚠ 폴더를 하나 열고 눌러주세요 (내 드라이브 최상위는 지정 불가)"; return; }
  const all = await DB.get("driveFolders", {});
  all[DRV.key] = { id: cur.id, name: cur.name };
  await DB.set("driveFolders", all);
  drvFolderInfo();
  $("drv-msg").textContent = `✔ 기본 폴더로 저장했어요: ${cur.name}`;
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
    // 마지막으로 본 폴더 기억 → 다음에 그 자리에서 시작 (매번 최상위부터 안 파고들게)
    if (DRV.key && folderId !== "root" && folderId !== "shared") {
      const cur = DRV.path[DRV.path.length - 1];
      const all = await DB.get("driveFolders", {});
      all[DRV.key + ":last"] = { id: folderId, name: cur ? cur.name : "" , path: DRV.path.slice() };
      await DB.set("driveFolders", all);
    }
  } catch (e) {
    if (/popup|로그인|권한|401|403|NEED_LOGIN/i.test(e.message || "")) { drvNeedLogin(); $("drv-msg").textContent = ""; }
    else { box.innerHTML = ""; $("drv-msg").textContent = "⚠ " + e.message; }
  }
}
async function drvSearch(q) {
  if (!q || !q.trim()) { DRV.path = [{ id: "root", name: "내 드라이브" }]; return drvOpen("root"); }
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
    $("drvmodal").classList.remove("on");
  } catch (e) { $("drv-msg").textContent = "⚠ " + e.message; }
}

/* --- ① 쇼핑몰 주문 파일 --- */
$("drive-order").onclick = () => openDrivePicker({
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
$("drive-again").onclick = async () => {
  const f = await DB.get("driveOrderFile", null); if (!f || !f.id) return;
  msg("msg-o", "", "가져오는 중…");
  try {
    const r = await GMAIL.driveFetchExcel(f.id);
    await setOrderFromBuf(r.buf, r.name);
    msg("msg-o", "ok", `✔ 최신본으로 다시 가져왔어요: ${r.name}`);
  } catch (e) { msg("msg-o", "err", "가져오기 실패: " + e.message); }
};
async function drawDriveRecent() {
  const f = await DB.get("driveOrderFile", null);
  const row = $("drive-recent-row");
  if (f && f.id) { $("drive-recent").textContent = "드라이브 최근 파일: " + f.name; row.style.display = "flex"; }
  else row.style.display = "none";
}

/* --- ② 업체 양식 (여러 개 선택 가능) --- */
$("drive-tpl").onclick = () => openDrivePicker({
  key: "tpl", title: "드라이브에서 업체 양식 가져오기", multiple: true,
  onPick: async files => {
    let n = 0;
    for (const f of files) {
      const r = await GMAIL.driveFetchExcel(f.id);
      await DB.putForm({ name: QO.nameFromFilename(r.name), file: r.name, data: r.buf, checked: true });
      n++;
    }
    await loadForms();
    msg("msg-o", "ok", `✔ 드라이브에서 업체 양식 ${n}개를 저장했어요.`);
  },
});

/* --- ③ 송장취합양식 (하나) --- */
$("drive-sab").onclick = () => openDrivePicker({
  key: "sab", title: "드라이브에서 송장취합양식 가져오기", multiple: false,
  onPick: async files => {
    const r = await GMAIL.driveFetchExcel(files[0].id);
    S.sabBuf = r.buf; S.sabName = r.name;
    S.sabDrive = { id: files[0].id, name: r.name };   // 드라이브 출처 기억 → 결과를 이 파일에 되쓰기
    $("sab-name").textContent = "📁 " + r.name + " (드라이브)";
    $("drop-sab").classList.add("on"); $("sab-preview").style.display = "block";
    refreshI(); msg("msg-i", "ok", `✔ 드라이브에서 송장취합양식을 가져왔어요: ${r.name}`);
  },
});

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
async function addForms(files) {
  let added = 0;
  for (const f of files) {
    if (!/\.xls[xm]$/i.test(f.name)) continue;
    const buf = await readFile(f);
    const name = QO.nameFromFilename(f.name);
    await DB.putForm({ name, file: f.name, data: buf, checked: true });
    added++;
  }
  if (added) { await loadForms(); msg("msg-o", "ok", `✔ 업체 양식 ${added}개를 이 기기에 저장했어요. 다음부터는 체크만 하면 됩니다.`); }
}
async function loadForms() {
  S.forms = await DB.listForms();
  S.brandVendor = await DB.get("brandVendor", {});
  S.vendorEmails = await DB.get("vendorEmails", {});
  S.vendorSent = await DB.get("vendorSent", {});
  S.vendorDomains = await DB.get("vendorDomains", {});
  S.invEmails = await DB.get("invEmails", "");
  S.invSent = await DB.get("invSent", []);
  drawForms(); buildVendorBrands(); refreshO();
}
function drawForms() {
  const box = $("vlist");
  if (!S.forms.length) { box.innerHTML = '<div class="empty">저장된 업체 양식이 없습니다.<br>아래에서 추가하세요.</div>'; return; }
  box.innerHTML = "";
  S.forms.forEach(f => {
    const el = document.createElement("div");
    el.className = "vrow" + (f.checked ? " on" : "");
    el.innerHTML = `<div class="vtop"><span class="box">${CHK}</span><b>${esc(f.name)}</b><button class="vdel">✕</button></div>
      <span class="vfile">${esc(f.file)}</span>
      <div class="vbtns"><button class="pv">미리보기</button><button class="dl">엑셀 받기</button></div>`;
    el.onclick = async ev => {
      if (ev.target.classList.contains("pv")) {
        ev.stopPropagation();
        openPreview(f.data, f.name + " 양식"); return;
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

function refreshO() {
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
        filename: `${QO.todayStr()}_랩노마드_${cleanVendor(f.name)}_발주양식.xlsx` });
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
      let t = `⚠ 건수 불일치 — 주문 ${verify.srcTotal}건 중 발주서에 ${verify.converted}건만 들어갔습니다 (${verify.diff}건 누락)\n${detail}`;
      if (verify.unassigned.length) {
        t += "\n\n아래 브랜드가 어느 업체에도 배정되지 않아 발주서에서 빠졌습니다:";
        verify.unassigned.forEach(u => { t += `\n· ${u.brand} — ${u.count}건`; });
        t += "\n\n해당 브랜드를 업체에 체크한 뒤 다시 변환하세요.";
      }
      d.className = "msg show err"; d.textContent = t;
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
      <div class="setrow" style="margin-top:6px"><span style="flex:1;font-size:11px;color:var(--faint)">여러 명에게 보내려면 쉼표로 구분 (담당자, 대표 등)</span>
        <button class="minibtn share">📤 카톡·공유</button><button class="minibtn pvbtn">미리보기</button><button class="minibtn dl">엑셀만 받기</button></div>`;
    const inp = el.querySelector("input");
    fillRecipients(el.querySelector(".cands"), inp, {
      saved: S.vendorEmails[r.supplier], history: S.vendorSent[r.supplier],
      domains: parseDomains(S.vendorDomains[r.supplier]), query: r.supplier });
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
        await GMAIL.send({ to: list.join(", "), subject: `[랩노마드] ${ymd}_발주서 송부`,
          body: "안녕하세요 발주서 송부드립니다. 감사합니다!",
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

/* 받는사람 후보 칩 (발주/송장 공용)
   opts = { saved:"주소들", history:[주소...], domains:[도메인...], query:"메일검색어" }
   ⓐ저장(기본) ⓑ이전 발송 ⓒ메일에서 찾은 주소(query+도메인 있을 때, 도메인 필터) → 클릭 선택 */
async function fillRecipients(container, inp, opts) {
  if (!container) return;
  opts = opts || {};
  container.innerHTML = "";
  const chosenSet = () => new Set(parseEmails(inp.value).map(e => e.toLowerCase()));
  const seen = new Set();
  const lbl = document.createElement("span");
  lbl.className = "candlbl"; lbl.textContent = "받는사람 후보:";
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
        found.map(f => f.email).filter(inDom).slice(0, 8).forEach(addChip);
      } catch (e) { hint.remove(); }
      if (seen.size === 0) {
        const s = document.createElement("span");
        s.className = "candhint"; s.textContent = "해당 도메인 주소를 못 찾음 — 직접 입력하세요";
        container.appendChild(s);
      }
    } else if (!doms.length) {
      const s = document.createElement("span");
      s.className = "candhint";
      s.textContent = "설정에서 이 업체의 메일 도메인을 넣으면, 메일함에서 받는사람을 자동으로 찾아줍니다.";
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
function refreshI() { $("run-i").disabled = !(S.sabBuf && S.reps.length); }

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
    if (out.missingCount === 0) {
      h += `<div class="msg show ok" style="margin-top:8px">✔ 취합본 빈칸 없음 — 주문 ${out.orderRows}행 전부 송장 기입 완료</div>`;
    } else {
      h += `<div class="msg show err" style="margin-top:8px">⚠ 취합본 송장 빈칸 <b>${out.missingCount}건</b> / 주문 ${out.orderRows}행 — 아래 주문은 업체 회신에 송장이 없습니다</div>`;
      h += missingTable(out.missing || [], out.missingCount);
    }
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
    <div class="setrow" style="margin-top:6px"><span style="flex:1;font-size:11px;color:var(--faint)">여러 명에게 보내려면 쉼표로 구분 (담당자, 대표 등)</span><button class="minibtn" id="share-inv">📤 카톡·공유</button><button class="minibtn" id="pv-inv">미리보기</button><button class="minibtn" id="dl-inv">엑셀만 받기</button></div>
    ${S.sabDrive ? `<button class="go" id="drv-writeback" style="margin-top:10px;font-size:14px;padding:12px;background:var(--ok);color:#fff">📥 드라이브 양식(${esc(S.sabDrive.name)})에 송장 기입</button>
      <div id="drv-wb-msg" style="font-size:11.5px;color:var(--muted);margin-top:6px;text-align:center"></div>` : ""}</div>`;
  $("rlist-i").innerHTML = h;
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
  // 받는사람 후보: 마지막 발송(기본) + 이전 이력 + '발주서 보내는 곳'(주문 메일 발신자) 주소
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
      await GMAIL.send({ to: list.join(", "), subject: `[랩노마드] ${ymd}_송장 취합본 송부`,
        body: "안녕하세요 송장 취합본 송부드립니다. 감사합니다!",
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
    if (r.changed) { await loadForms(); drawOrderFilter(); drawReplyFilter(); drawSettings(); }
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
  updateGmailWho();
  syncOnStart();          // 로그인 직후 다른 기기 데이터 내려받기
}

/* 발주서 검색조건 (PC 앱과 동일한 기본값) */
async function getOrderFilter() {
  return {
    senders: await DB.get("orderSenders", ["onekglobal.co.kr"]),
    keywords: await DB.get("orderKeywords", ["랩노마드 발주서", "랩노마드발주서", "★랩노마드", "랩노마드"]),
    exclude: await DB.get("orderExclude", ["플라스머", "디에스피", "송장", "회신", "운송장", "택배"]),
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
    senders: await DB.get("replySenders", []),
    keywords: await DB.get("replyKeywords", ["송장", "운송장", "회신"]),
    exclude: await DB.get("replyExclude", []),
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
};

async function openFilter(mode) {
  filterMode = mode;
  $("filter-title").textContent = mode === "order" ? "발주서 검색조건" : "회신 송장 검색조건";
  await renderFilterLists();
  filterModal.classList.add("on");
}
async function renderFilterLists() {
  const f = filterMode === "order" ? await getOrderFilter() : await getReplyFilter();
  drawChipList("flt-senders", f.senders, "senders");
  drawChipList("flt-keywords", f.keywords, "keywords");
  drawChipList("flt-excludes", f.exclude || [], "exclude");
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
  const def = kind === "senders" ? (filterMode === "order" ? ["onekglobal.co.kr"] : [])
    : kind === "keywords" ? (filterMode === "order" ? ["랩노마드 발주서", "랩노마드발주서", "★랩노마드", "랩노마드"] : ["송장", "운송장", "회신"])
    : (filterMode === "order" ? ["플라스머", "디에스피", "송장", "회신", "운송장", "택배"] : []);
  return await DB.get(key, def);
}
async function setList(kind, arr) {
  await DB.set(FKEY[filterMode][kind], arr);
  await renderFilterLists();
  drawOrderFilter(); drawReplyFilter();
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
      const el = document.createElement("div");
      el.className = "mitem";
      el.innerHTML = `<div style="font-weight:700;font-size:13px;word-break:break-all">📄 ${esc(m.filename)}</div>
        <div style="font-size:11.5px;color:var(--muted);margin-top:3px">${esc(m.date)} · ${esc(m.subject || "(제목 없음)")}</div>
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
async function drawNotifyStatus() {
  const on = await DB.get("notifyOn", false);
  $("notify-toggle").textContent = on ? "끄기" : "켜기";
  const perm = (window.Notification && Notification.permission) || "unsupported";
  $("notify-status").textContent = !on ? "꺼져 있음"
    : (perm === "granted" ? "✓ 켜짐 — 앱을 열어두면 새 발주·송장을 알려드려요"
    : perm === "denied" ? "⚠ 브라우저 알림이 차단됨 — 앱 안 배너로만 표시됩니다"
    : "켜짐 — 알림 권한을 허용하면 배너+알림 둘 다 떠요");
}
$("notify-toggle").onclick = async function () {
  let on = await DB.get("notifyOn", false);
  on = !on;
  if (on && window.Notification && Notification.permission === "default") {
    try { await Notification.requestPermission(); } catch (e) {}
  }
  await DB.set("notifyOn", on);
  drawNotifyStatus();
  if (on) { startNotify(); notifyTick(false); } else stopNotify();   // 켤 땐 조용히 기준선만
};
function startNotify() { if (!notifyTimer) notifyTimer = setInterval(() => notifyTick(false), NOTIFY_MS); }
function stopNotify() { if (notifyTimer) { clearInterval(notifyTimer); notifyTimer = null; } }

async function notifyTick(manual) {
  if (!(await DB.get("notifyOn", false))) return;
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
    if (r.changed) { await loadForms(); drawOrderFilter(); drawReplyFilter(); if ($("setmodal").classList.contains("on")) drawSettings(); }
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
if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => {});
// 알림: 켜져 있으면 폴링 시작, 앱으로 돌아올 때마다 즉시 한 번 확인
DB.get("notifyOn", false).then(on => { if (on) { startNotify(); setTimeout(() => notifyTick(false), 4000); } });
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible") DB.get("notifyOn", false).then(on => { if (on) notifyTick(false); });
});
