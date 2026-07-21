/* ===================================================================
   퀵오더 앱 — 화면 동작 + 기기 저장(IndexedDB)
   =================================================================== */
"use strict";
const CHK = '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"><path d="M5 13l4 4L19 7"/></svg>';
const $ = id => document.getElementById(id);
const esc = s => String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

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
  return {
    listForms: () => tx("forms", "readonly", s => s.getAll()),
    putForm: f => tx("forms", "readwrite", s => s.put(f)),
    delForm: n => tx("forms", "readwrite", s => s.delete(n)),
    get: async (k, dflt) => { const v = await tx("kv", "readonly", s => s.get(k)); return v && v.v !== undefined ? v.v : dflt; },
    set: (k, v) => tx("kv", "readwrite", s => s.put({ k, v })),
  };
})();

/* ---------------- 공통 상태 ---------------- */
const S = {
  orderWb: null, orderBuf: null, orderName: "",
  brands: [], dateAll: [], dateSel: [], dateHeader: null,
  pv: null, pvAll: false,
  forms: [], brandVendor: {}, vendorEmails: {}, sel: {},
  sabBuf: null, sabName: "", reps: [],
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

/* --- 날짜 --- */
$("dt-col").onchange = function () { loadDates(this.value); };
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
  list.forEach(d => {
    const el = document.createElement("span");
    el.className = "brow"; el.dataset.d = d.date;
    el.innerHTML = `<span class="box">${CHK}</span>${esc(d.label)} (${d.count}건)`;
    el.onclick = () => {
      const i = S.dateSel.indexOf(d.date);
      if (i >= 0) S.dateSel.splice(i, 1); else S.dateSel.push(d.date);
      drawDateChips();
    };
    box.appendChild(el);
  });
  drawDateChips();
}
function drawDateChips() {
  const box = $("dt-chips");
  box.querySelectorAll(".brow[data-d]").forEach(el => el.classList.toggle("on", S.dateSel.includes(el.dataset.d)));
  const all = box.querySelector(".brow:not([data-d])");
  const totalAll = S.dateAll.reduce((s, d) => s + d.count, 0);
  if (all) all.textContent = (S.dateSel.length === S.dateAll.length && S.dateAll.length) ? "전체 해제" : `전체 ${totalAll}건 선택`;
  const cnt = S.dateAll.filter(d => S.dateSel.includes(d.date)).reduce((s, d) => s + d.count, 0);
  $("dt-foot").textContent = S.dateSel.length
    ? `선택한 ${S.dateSel.length}개 날짜 · 총 ${cnt}건만 변환됩니다`
    : "⚠ 날짜를 하나 이상 선택하세요";
  refreshO();
}

/* --- 내용 확인 --- */
$("pv-toggle").onclick = function () { S.pvAll = !S.pvAll; this.textContent = S.pvAll ? "주요 열만 보기" : "전체 열 보기"; renderPreview(); };
async function drawPreview() {
  $("prev-wrap").style.display = "block";
  $("pv-cnt").textContent = "· 읽는 중…";
  const wb = await QO.loadWorkbook(S.orderBuf.slice(0));
  S.pv = QO.preview(wb); S.pvAll = false;
  $("pv-toggle").textContent = "전체 열 보기";
  renderPreview();
}
function renderPreview() {
  const pv = S.pv; if (!pv) return;
  const idx = (S.pvAll || !pv.keyIdx.length) ? pv.columns.map((_, i) => i) : pv.keyIdx;
  $("pv-cnt").textContent = `· 전체 ${pv.total}건 · 열 ${idx.length}/${pv.columns.length}`;
  let h = "<tr>" + idx.map(i => `<th>${esc(pv.columns[i] || "열" + (i + 1))}</th>`).join("") + "</tr>";
  pv.rows.forEach(row => {
    h += "<tr>" + idx.map(i => {
      const v = row[i] == null ? "" : row[i];
      const num = /^[0-9,.\-]+$/.test(v) && v !== "";
      return `<td${num ? ' class="num"' : ""}>${esc(v)}</td>`;
    }).join("") + "</tr>";
  });
  $("pv-table").innerHTML = h;
  $("pv-foot").textContent = pv.total > pv.rows.length
    ? `앞 ${pv.rows.length}건만 표시 · 전체 ${pv.total}건 모두 변환됩니다` : `전체 ${pv.total}건 표시`;
}

/* --- 업체 양식 --- */
$("f-tpl").addEventListener("change", function () { addForms(this.files); this.value = ""; });
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
  drawForms(); buildVendorBrands(); refreshO();
}
function drawForms() {
  const box = $("vlist");
  if (!S.forms.length) { box.innerHTML = '<div class="empty">저장된 업체 양식이 없습니다.<br>아래에서 추가하세요.</div>'; return; }
  box.innerHTML = "";
  S.forms.forEach(f => {
    const el = document.createElement("div");
    el.className = "vrow" + (f.checked ? " on" : "");
    el.innerHTML = `<span class="box">${CHK}</span><div class="vinfo"><b>${esc(f.name)}</b><span>${esc(f.file)}</span></div><button class="vdel">✕</button>`;
    el.onclick = async ev => {
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
function buildVendorBrands() {
  const checked = S.forms.filter(f => f.checked);
  const card = $("card3"), box = $("vbrands");
  if (!checked.length || !S.brands.length) { card.style.display = "none"; box.innerHTML = ""; return; }
  card.style.display = "block"; box.innerHTML = "";
  checked.forEach(f => {
    if (!S.sel[f.name]) S.sel[f.name] = S.brands.filter(b => S.brandVendor[b] === f.name);
    const wrap = document.createElement("div");
    wrap.className = "vendorbox";
    wrap.innerHTML = `<div class="vh">🏭 ${esc(f.name)}<span class="cnt"></span><button class="all">전체</button></div>
      <div class="brands">${S.brands.map(b => {
        const on = S.sel[f.name].includes(b);
        return `<span class="brow${on ? " on" : ""}" data-b="${esc(b)}"><span class="box">${CHK}</span>${esc(b)}</span>`;
      }).join("")}</div>`;
    wrap.querySelectorAll(".brow").forEach(chip => {
      chip.onclick = () => {
        const b = chip.dataset.b, arr = S.sel[f.name], i = arr.indexOf(b);
        if (i >= 0) arr.splice(i, 1); else arr.push(b);
        chip.classList.toggle("on", i < 0);
        updCnt(wrap, f); refreshO();
      };
    });
    wrap.querySelector(".all").onclick = () => {
      const all = S.sel[f.name].length === S.brands.length;
      S.sel[f.name] = all ? [] : S.brands.slice();
      wrap.querySelectorAll(".brow").forEach(c => c.classList.toggle("on", !all));
      updCnt(wrap, f); refreshO();
    };
    box.appendChild(wrap); updCnt(wrap, f);
  });
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
    const shop = QO.nameFromFilename(S.orderName);
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
        filename: `${QO.todayStr()}_${shop}_${f.name}_발주양식.xlsx` });
      // 학습 저장
      if (S.brands.length && sel.length) sel.forEach(b => { S.brandVendor[b] = f.name; });
    }
    if (!results.length) throw new Error("변환된 업체가 없습니다. " + (skipped.length ? `(${skipped.join(", ")})` : ""));
    await DB.set("brandVendor", S.brandVendor);
    showResultO(results, skipped);
    msg("msg-o", "ok", "✔ 변환 완료! " + results.map(r => `${r.supplier}=${r.count}건`).join("; "));
  } catch (e) { msg("msg-o", "err", "변환 실패: " + e.message); }
  finally { busy("run-o", "run-o-lbl", false, "발주서 변환하기"); refreshO(); }
};

function showResultO(results, skipped) {
  const box = $("rlist-o"); box.innerHTML = "";
  results.forEach(r => {
    const el = document.createElement("div");
    el.className = "rrow";
    el.innerHTML = `<div class="rtop"><div class="vinfo"><b>${esc(r.supplier)}</b><span>${esc(r.filename)}</span></div>
      <span class="cnt">${r.count}건</span></div>
      <div class="rmail"><input type="email" placeholder="${esc(r.supplier)} 이메일 (기억됩니다)"
        value="${esc(S.vendorEmails[r.supplier] || "")}" inputmode="email" autocapitalize="off" spellcheck="false">
        <button class="dlbtn send">메일 보내기</button></div>
      <div class="setrow" style="margin-top:6px"><span style="text-align:right;flex:1"></span>
        <button class="minibtn dl">엑셀만 받기</button></div>`;
    const inp = el.querySelector("input");
    el.querySelector(".dl").onclick = () => download(r.buf, r.filename);
    inp.onchange = inp.onblur = async () => {
      const v = inp.value.trim(); if (!v || v === S.vendorEmails[r.supplier]) return;
      S.vendorEmails[r.supplier] = v; await DB.set("vendorEmails", S.vendorEmails);
    };
    const sendBtn = el.querySelector(".send");
    sendBtn.onclick = async () => {
      const to = inp.value.trim();
      if (!to) { inp.focus(); return; }
      sendBtn.disabled = true; sendBtn.textContent = "보내는 중…";
      try {
        await ensureGmail();
        const ymd = QO.todayStr().slice(2);
        await GMAIL.send({ to, subject: `[랩노마드] ${ymd}_발주서 송부`,
          body: "안녕하세요 발주서 송부드립니다. 감사합니다!",
          attachments: [{ filename: r.filename, data: r.buf }] });
        S.vendorEmails[r.supplier] = to; await DB.set("vendorEmails", S.vendorEmails);
        sendBtn.textContent = "✓ 발송완료"; sendBtn.style.background = "var(--ok)";
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

/* =================================================================
   ② 송장 취합
   ================================================================= */
$("f-sab").addEventListener("change", async function () {
  if (this.files[0]) { S.sabBuf = await readFile(this.files[0]); S.sabName = this.files[0].name;
    $("sab-name").textContent = "📄 " + this.files[0].name; $("drop-sab").classList.add("on"); refreshI(); }
});
bindDrop("drop-sab", async f => {
  S.sabBuf = await readFile(f[0]); S.sabName = f[0].name;
  $("sab-name").textContent = "📄 " + f[0].name; $("drop-sab").classList.add("on"); refreshI();
});
$("f-rep").addEventListener("change", function () { addReps(this.files); this.value = ""; });
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
    el.innerHTML = `<span class="box">${CHK}</span><div class="vinfo"><b>${esc(QO.nameFromFilename(r.name))}</b>
      <span>${esc(r.name)}</span></div><button class="vdel">✕</button>`;
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

function showResultI(out, buf, filename) {
  let h = `<div class="tblbox" style="margin-bottom:12px"><div class="tblscroll"><table class="pv">
    <tr><th>업체</th><th>기입</th><th>미매칭</th><th>상태</th></tr>`;
  out.per.forEach(p => { h += `<tr><td>${esc(p[0])}</td><td class="num">${p[1]}</td><td class="num">${p[2]}</td><td>${esc(p[3])}</td></tr>`; });
  h += "</table></div></div>";
  if (out.gap === 0) {
    h += `<div class="msg show ok" style="margin-top:0">✔ 송장 갯수 일치 — 회신 ${out.srcInvoice}건 = 취합본 ${out.writtenInvoice}건, 누락 없음</div>`;
  } else if (out.gap > 0) {
    let d = "";
    for (const k in out.perSrc) d += `\n· ${k} 회신 ${out.perSrc[k]}건`;
    h += `<div class="msg show err" style="margin-top:0">⚠ 송장 누락 ${out.gap}건\n회신 양식 ${out.srcInvoice}건 중 취합본에 ${out.writtenInvoice}건만 기입되었습니다.${d}\n\n회신본의 수취인·주소·상품명이 원본과 달라졌는지 확인하세요.</div>`;
  } else {
    h += `<div class="msg show err" style="margin-top:0">⚠ 취합본 기입(${out.writtenInvoice}건)이 회신 송장(${out.srcInvoice}건)보다 많습니다. 회신 파일 중복을 확인하세요.</div>`;
  }
  h += `<div class="rrow" style="margin-top:12px"><div class="rtop"><div class="vinfo"><b>송장 취합본</b>
    <span>${esc(filename)}</span></div><span class="cnt">${out.total}건</span></div>
    <div class="rmail"><input type="email" id="inv-to" placeholder="받는 사람 이메일" inputmode="email" autocapitalize="off" spellcheck="false">
      <button class="dlbtn" id="send-inv">메일 보내기</button></div>
    <div class="setrow" style="margin-top:6px"><span style="flex:1"></span><button class="minibtn" id="dl-inv">엑셀만 받기</button></div></div>`;
  $("rlist-i").innerHTML = h;
  $("dl-inv").onclick = () => download(buf, filename);
  DB.get("vendorEmails", {}).then(v => { /* 기본값 없음 */ });
  $("send-inv").onclick = async function () {
    const to = $("inv-to").value.trim();
    if (!to) { $("inv-to").focus(); return; }
    this.disabled = true; this.textContent = "보내는 중…";
    try {
      await ensureGmail();
      const ymd = QO.todayStr().slice(2);
      await GMAIL.send({ to, subject: `[랩노마드] ${ymd}_송장 취합본 송부`,
        body: "안녕하세요 송장 취합본 송부드립니다. 감사합니다!",
        attachments: [{ filename, data: buf }] });
      this.textContent = "✓ 발송완료"; this.style.background = "var(--ok)";
    } catch (e) { this.disabled = false; this.textContent = "메일 보내기"; alert("발송 실패: " + e.message); }
  };
  $("result-i").style.display = "block";
  $("result-i").scrollIntoView({ behavior: "smooth", block: "start" });
}

/* =================================================================
   설정 (업체 메일 · 저장 데이터)
   ================================================================= */
$("btn-settings").onclick = () => { drawSettings(); $("setmodal").classList.add("on"); };
$("set-close").onclick = () => $("setmodal").classList.remove("on");
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
  const cid = await DB.get("gmailClientId", "");
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
  names.forEach(name => {
    const el = document.createElement("div");
    el.className = "mitem";
    el.innerHTML = `<div style="font-weight:700;font-size:13px">🏭 ${esc(name)}</div>
      <div style="display:flex;gap:7px;margin-top:8px">
        <input type="email" value="${esc(S.vendorEmails[name] || "")}" placeholder="이메일 주소"
          inputmode="email" autocapitalize="off" spellcheck="false"
          style="flex:1;min-width:0;border:1.5px solid var(--line);background:var(--card2);color:var(--ink);
          border-radius:9px;padding:10px;font-family:inherit;font-size:13px;outline:none">
        <button class="minibtn" style="padding:0 12px">저장</button>
        <button class="minibtn" style="padding:0 12px;color:var(--danger)">삭제</button>
      </div>`;
    const inp = el.querySelector("input"), btns = el.querySelectorAll("button");
    btns[0].onclick = async () => {
      S.vendorEmails[name] = inp.value.trim();
      if (!S.vendorEmails[name]) delete S.vendorEmails[name];
      await DB.set("vendorEmails", S.vendorEmails);
      btns[0].textContent = "완료"; setTimeout(() => btns[0].textContent = "저장", 1200);
    };
    btns[1].onclick = async () => {
      if (!confirm(`${name} 의 저장된 메일을 지울까요?`)) return;
      delete S.vendorEmails[name]; await DB.set("vendorEmails", S.vendorEmails); drawSettings();
    };
    box.appendChild(el);
  });
}

/* =================================================================
   Gmail 연동
   ================================================================= */
let gmailReady = false;
async function initGmail() {
  const cid = await DB.get("gmailClientId", "");
  GMAIL.init(cid);                      // 클라이언트 ID 등록(라이브러리 늦어도 됨)
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
async function ensureGmail() {
  const cid = await DB.get("gmailClientId", "");
  if (!cid) { $("btn-settings").click(); throw new Error("먼저 설정에서 클라이언트 ID를 저장하세요."); }
  if (!gmailReady) { GMAIL.init(cid); gmailReady = await GMAIL.waitReady(); updateGmailWho(); }
  if (!gmailReady) throw new Error("구글 로그인 라이브러리를 불러오지 못했어요.\n인터넷/광고차단을 확인하고 새로고침 해보세요.");
  if (GMAIL.signedIn()) { updateGmailWho(); return; }   // 저장된 로그인 유효 → 그대로 사용
  // 토큰이 아예 없으면 로그인창(동의), 만료됐으면 조용히 갱신
  await GMAIL.signIn(!GMAIL.hasToken());
  updateGmailWho();
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
  const list = $("mail-list");
  list.innerHTML = '<div class="empty">메일함을 확인하고 있어요…<br><span id="mail-prog"></span></div>';
  $("mail-sub").textContent = "최근 7일 메일 확인 중…";
  try {
    let opt;
    if (target === "rep") {
      const f = await getReplyFilter();
      opt = { days: 7, senders: f.senders, keywords: f.keywords, exclude: f.exclude || [], union: true, scanText: true };
    } else {
      // 발주서/사방넷: 저장된 발신자·키워드·제외어로 선별 (PC 앱과 동일)
      const f = await getOrderFilter();
      opt = { days: 7, senders: f.senders, keywords: f.keywords, exclude: f.exclude, union: false, scanText: true };
    }
    opt.onProgress = (i, n) => { const p = $("mail-prog"); if (p) p.textContent = `${i} / ${n}`; };
    mailItems = await GMAIL.listMails(opt);
    if (!mailItems.length) {
      list.innerHTML = '<div class="empty">최근 7일간 해당 엑셀 첨부를 찾지 못했어요.</div>';
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
      $("sab-name").textContent = "📧 " + got[0].name; $("drop-sab").classList.add("on"); refreshI();
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

/* ---------------- 시작 ---------------- */
loadForms().catch(e => { $("vlist").innerHTML = '<div class="empty">저장소를 열지 못했어요</div>'; });
initGmail();
drawReplyFilter();
drawOrderFilter();
if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => {});
