/* =====================================================================
   퀵오더 ④ 정산  (v6.0)
   쇼핑몰에서 확정된 정산금액 → 상품에서 브랜드/업체를 찾아 →
   우리 마진을 뺀 '업체 지급액'을 만든다.
     · 마진율(%) 방식  : 업체지급 = 정산금액 × (1 − 마진율/100)
     · 공급단가 방식   : 업체지급 = 수량 × 공급단가
   교환/반품 비용(③ CS에서 귀책 '업체' + 비용)은 업체별로 자동 차감한다.
   ===================================================================== */
"use strict";

const ST = (() => {
  const FIELDS = [
    { k: "date", n: "정산일", kw: ["정산일", "정산기준일", "결제일", "주문일", "구매확정일", "일자", "날짜"] },
    { k: "mall", n: "쇼핑몰", kw: ["쇼핑몰", "판매처", "마켓", "채널", "사이트", "몰"] },
    { k: "orderNo", n: "주문번호", kw: ["주문번호", "주문no", "오더번호", "결제번호", "order"] },
    { k: "product", n: "상품명", kw: ["상품명", "제품명", "상품", "품목", "노출상품명"], req: true },
    { k: "option", n: "옵션", kw: ["옵션", "옵션명", "선택옵션", "규격"] },
    { k: "qty", n: "수량", kw: ["수량", "판매수량", "개수"] },
    { k: "amount", n: "정산금액", kw: ["정산금액", "정산예정금액", "지급액", "정산대금", "실정산액", "정산", "금액"], req: true },
  ];

  let files = [];        // [{name, cols, rows, map, sig}]
  let rules = { def: 30, vendors: {} };
  let maps = {};
  let result = null;     // 계산 결과
  let useCs = true;
  let drawn = false;

  const s = v => (v === null || v === undefined) ? "" : String(v).trim();
  const num = v => {
    const n = Number(String(v == null ? "" : v).replace(/[^0-9.\-]/g, ""));
    return isFinite(n) ? n : 0;
  };
  const won = n => (Math.round(n || 0)).toLocaleString("ko-KR") + "원";

  async function load() {
    rules = await DB.get("settleRules", { def: 30, vendors: {} }) || { def: 30, vendors: {} };
    if (!rules.vendors) rules.vendors = {};
    maps = await DB.get("settleMaps", {}) || {};
    $("st-default").value = rules.def;
  }
  async function saveRules() { await DB.set("settleRules", rules); }

  /* =================================================================
     정산 파일 가져오기
     ================================================================= */
  async function addFile(buf, name) {
    const wb = await QO.loadWorkbook(buf.slice(0));
    const pv = QO.previewAny(wb, 20000);
    if (!pv.columns.length) throw new Error("표를 찾지 못했어요.");
    const sig = MAP.signature(pv.columns);
    const saved = maps[sig];
    const auto = Object.assign(MAP.autoMap(pv.columns, FIELDS), saved || {});
    const put = async m => {
      maps[sig] = m; await DB.set("settleMaps", maps);
      files = files.filter(f => f.name !== name);
      files.push({ name, cols: pv.columns, rows: pv.rows, map: m, sig });
      drawFiles(); drawRules(); refresh();
      msg("msg-s", "ok", `✔ ${name} — ${pv.rows.length}행 불러왔어요.`);
    };
    if (saved && MAP.ok(auto, FIELDS)) { await put(auto); return; }
    MAP.open({
      title: "정산 열 맞추기",
      sub: `${name} — 정산금액과 상품명이 어느 열인지 정해주세요. 같은 양식은 다음부터 자동입니다.`,
      columns: pv.columns, fields: FIELDS, saved: auto,
      onOk: m => put(m).catch(e => msg("msg-s", "err", "⚠ " + e.message)),
    });
  }
  function drawFiles() {
    const box = $("st-files");
    if (!files.length) { box.innerHTML = ""; return; }
    box.innerHTML = files.map((f, i) => `<div class="vrow on">
        <div class="vtop"><b>📄 ${esc(f.name)}</b><button class="vdel" data-i="${i}">✕</button></div>
        <span class="vfile">${f.rows.length}행</span></div>`).join("");
    box.querySelectorAll(".vdel").forEach(b => b.onclick = () => {
      files.splice(Number(b.dataset.i), 1);
      drawFiles(); drawRules(); refresh();
    });
  }

  /* =================================================================
     업체 목록 / 마진 규칙
     ================================================================= */
  function allRows() {
    const out = [];
    for (const f of files) {
      const g = (r, k) => (f.map[k] === undefined ? "" : r[f.map[k]]);
      for (const r of f.rows) {
        const product = s(g(r, "product"));
        const amount = num(g(r, "amount"));
        if (!product && !amount) continue;
        out.push({
          src: f.name,
          date: CS.toYmd(g(r, "date")),
          mall: s(g(r, "mall")) || guessMall(f.name),
          orderNo: s(g(r, "orderNo")),
          product, option: s(g(r, "option")),
          qty: num(g(r, "qty")) || 1,
          amount,
          vendor: CS.findVendor(product, g(r, "option")) || "",
          brand: CS.findBrand(product, g(r, "option")) || "",
        });
      }
    }
    return out;
  }
  function guessMall(filename) {
    const n = String(filename || "");
    const M = ["스마트스토어", "네이버", "쿠팡", "11번가", "지마켓", "옥션", "위메프", "티몬", "카카오", "SSG", "롯데"];
    for (const m of M) if (n.includes(m)) return m;
    return "";
  }
  function vendorList() {
    const set = new Set();
    allRows().forEach(r => set.add(r.vendor || "(업체 미지정)"));
    return [...set].sort();
  }
  function ruleOf(v) {
    const r = rules.vendors[v] || {};
    return { mode: r.mode || "rate", rate: r.rate === undefined ? rules.def : r.rate, unit: r.unit || 0 };
  }
  function drawRules() {
    const box = $("st-rules"), vs = vendorList();
    if (!vs.length) {
      box.innerHTML = "";
      $("st-rule-foot").textContent = "정산 파일을 올리면 업체 목록이 자동으로 나옵니다.";
      return;
    }
    $("st-rule-foot").textContent = "‘공급단가’를 고르면 수량 × 단가로 지급액을 계산합니다.";
    box.innerHTML = vs.map(v => {
      const r = ruleOf(v);
      return `<div class="rulerow" data-v="${esc(v)}">
        <b>🏭 ${esc(v)}</b>
        <select class="mode">
          <option value="rate"${r.mode === "rate" ? " selected" : ""}>마진율%</option>
          <option value="unit"${r.mode === "unit" ? " selected" : ""}>공급단가</option>
        </select>
        <input class="val" type="number" inputmode="decimal" step="${r.mode === "rate" ? "0.1" : "10"}"
               value="${r.mode === "rate" ? r.rate : r.unit}">
      </div>`;
    }).join("");
    box.querySelectorAll(".rulerow").forEach(row => {
      const v = row.dataset.v;
      const mode = row.querySelector(".mode"), val = row.querySelector(".val");
      const apply = async () => {
        const cur = rules.vendors[v] || {};
        cur.mode = mode.value;
        if (mode.value === "rate") cur.rate = num(val.value);
        else cur.unit = num(val.value);
        rules.vendors[v] = cur;
        await saveRules();
        if (result) calc();
      };
      mode.onchange = async () => { await apply(); drawRules(); };
      val.onchange = apply;
    });
  }

  /* =================================================================
     계산
     ================================================================= */
  function csDeduction() {
    // ③ CS 중 귀책 '업체' + 비용이 있는 건 → 업체별 차감액
    const out = {};
    if (!useCs || !window.CS) return out;
    for (const x of CS.items()) {
      if (x.settled) continue;
      if (x.fault !== "업체") continue;
      const c = Number(x.cost) || 0;
      if (!c) continue;
      const v = x.vendor || "(업체 미지정)";
      (out[v] = out[v] || { amount: 0, rows: [] });
      out[v].amount += c;
      out[v].rows.push(x);
    }
    return out;
  }
  function calc() {
    const rows = allRows();
    if (!rows.length) { result = null; return; }
    const ded = csDeduction();
    const byVendor = {};
    for (const r of rows) {
      const v = r.vendor || "(업체 미지정)";
      const rule = ruleOf(v);
      r.pay = rule.mode === "unit" ? r.qty * rule.unit : r.amount * (1 - rule.rate / 100);
      r.margin = r.amount - r.pay;
      const g = byVendor[v] = byVendor[v] || { vendor: v, rows: [], amount: 0, pay: 0, margin: 0, ded: 0, dedRows: [] };
      g.rows.push(r); g.amount += r.amount; g.pay += r.pay; g.margin += r.margin;
    }
    for (const v in byVendor) {
      const d = ded[v];
      if (d) { byVendor[v].ded = d.amount; byVendor[v].dedRows = d.rows; }
      byVendor[v].final = byVendor[v].pay - byVendor[v].ded;
    }
    const vendors = Object.values(byVendor).sort((a, b) => b.final - a.final);
    result = {
      vendors,
      total: {
        amount: vendors.reduce((s2, v) => s2 + v.amount, 0),
        pay: vendors.reduce((s2, v) => s2 + v.pay, 0),
        margin: vendors.reduce((s2, v) => s2 + v.margin, 0),
        ded: vendors.reduce((s2, v) => s2 + v.ded, 0),
        final: vendors.reduce((s2, v) => s2 + v.final, 0),
        count: rows.length,
      },
      unmatched: rows.filter(r => !r.vendor).length,
    };
    drawResult();
  }

  function drawResult() {
    if (!result) { $("result-s").style.display = "none"; return; }
    $("result-s").style.display = "block";
    const t = result.total;
    $("st-total").innerHTML =
      `<div class="totline"><b>몰 정산금액 합계</b><span>${won(t.amount)}</span></div>
       <div class="totline"><b>우리 마진</b><span style="color:var(--ok)">${won(t.margin)}</span></div>
       ${t.ded ? `<div class="totline"><b>CS 차감 (교환·반품)</b><span style="color:var(--danger)">− ${won(t.ded)}</span></div>` : ""}
       <div class="totline" style="border-top:1px solid var(--line);margin-top:6px;padding-top:8px">
         <b>업체 지급 합계</b><span style="color:var(--brand)">${won(t.final)}</span></div>
       <div class="synchint" style="margin-top:6px">${t.count}건${result.unmatched ? ` · <b style="color:var(--danger)">업체 미지정 ${result.unmatched}건</b> — 상품명에서 브랜드를 못 찾았어요. ① 발주 탭에서 브랜드-업체를 한 번 지정하면 다음부터 자동입니다.` : ""}</div>`;

    $("st-sumbox").innerHTML = result.vendors.map(v =>
      `<div class="sumrow"><b>🏭 ${esc(v.vendor)}</b>
        <span style="color:var(--muted);font-size:11.5px">${v.rows.length}건</span>
        <span class="pay">${won(v.final)}</span></div>`).join("");

    const box = $("st-vendors");
    box.innerHTML = "";
    result.vendors.forEach(v => {
      const row = document.createElement("div");
      row.className = "rrow";
      row.innerHTML = `<div class="rtop"><b style="flex:1">🏭 ${esc(v.vendor)}</b><span class="cnt">${v.rows.length}건</span></div>
        <div class="totline" style="font-size:12px"><span>정산금액</span><span>${won(v.amount)}</span></div>
        <div class="totline" style="font-size:12px"><span>우리 마진</span><span>${won(v.margin)}</span></div>
        ${v.ded ? `<div class="totline" style="font-size:12px;color:var(--danger)"><span>CS 차감 ${v.dedRows.length}건</span><span>− ${won(v.ded)}</span></div>` : ""}
        <div class="totline" style="font-size:13px"><b>지급액</b><span style="color:var(--brand)">${won(v.final)}</span></div>
        <div class="rmail"><input type="email" placeholder="업체 이메일 (쉼표로 여러 명)" autocapitalize="off" spellcheck="false">
          <button class="dlbtn">📧 정산서</button></div>
        <div class="cands"></div>
        <div style="display:flex;gap:6px;margin-top:8px">
          <button class="minibtn xls" style="flex:1;padding:9px">📥 엑셀 저장</button>
          <button class="minibtn pvv" style="flex:1;padding:9px">👁 내역 보기</button>
        </div>`;
      box.appendChild(row);
      const inp = row.querySelector("input");
      const known = S.vendorEmails[v.vendor] || "";
      inp.value = known;
      fillRecipients(row.querySelector(".cands"), inp, {
        saved: known, history: S.vendorSent[v.vendor] || [],
        domains: S.vendorDomains[v.vendor] || [], query: v.vendor !== "(업체 미지정)" ? v.vendor : "",
      });
      row.querySelector(".xls").onclick = async () => {
        const buf = await buildExcel([v]);
        download(buf, fileName(v.vendor));
      };
      row.querySelector(".pvv").onclick = () => showRows(v);
      row.querySelector(".dlbtn").onclick = () => sendOne(v, inp, row.querySelector(".dlbtn"));
    });
  }
  const fileName = v => `${QO.todayStr()}_${CONFIG.company}_${cleanVendor(v)}_정산서.xlsx`;

  function showRows(v) {
    const cols = ["정산일", "쇼핑몰", "주문번호", "상품명", "옵션", "수량", "정산금액", "우리마진", "지급액"];
    const rows = v.rows.map(r => [r.date, r.mall, r.orderNo, r.product, r.option, r.qty,
      Math.round(r.amount), Math.round(r.margin), Math.round(r.pay)]);
    $("pv-modal-title").textContent = v.vendor + " 정산 내역";
    $("pv-modal-sub").textContent = `${rows.length}건 · 지급액 ${won(v.final)}`;
    const tbl = $("pv-modal-table");
    tbl.innerHTML = `<thead><tr>${cols.map(c => `<th>${esc(c)}</th>`).join("")}</tr></thead>`
      + `<tbody>${rows.map(r => `<tr>${r.map((c, i) => `<td${i >= 5 ? ' class="num"' : ""}>${esc(c)}</td>`).join("")}</tr>`).join("")}</tbody>`;
    $("pv-modal-foot").textContent = "";
    $("pvmodal").classList.add("on");
  }

  /* =================================================================
     엑셀 / 메일
     ================================================================= */
  async function buildExcel(vendors) {
    const wb = new ExcelJS.Workbook();
    const head = ["정산일", "쇼핑몰", "주문번호", "상품명", "옵션", "수량", "정산금액", "우리마진", "지급액"];
    const safe = n => String(n).replace(/[\\\/\?\*\[\]:]/g, "_").slice(0, 28) || "정산";
    for (const v of vendors) {
      const ws = wb.addWorksheet(safe(v.vendor));
      ws.addRow([`${v.vendor} 정산서`]);
      ws.getRow(1).font = { bold: true, size: 14 };
      ws.addRow([`작성일 ${QO.fmtDate(QO.todayStr())}`]);
      ws.addRow([]);
      ws.addRow(head);
      const hr = ws.lastRow.number;
      ws.getRow(hr).font = { bold: true };
      ws.getRow(hr).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEFF3FB" } };
      v.rows.forEach(r => ws.addRow([r.date, r.mall, r.orderNo, r.product, r.option, r.qty,
        Math.round(r.amount), Math.round(r.margin), Math.round(r.pay)]));
      ws.addRow([]);
      const add = (label, val, color) => {
        const row = ws.addRow(["", "", "", "", "", "", "", label, Math.round(val)]);
        row.font = { bold: true, color: color ? { argb: color } : undefined };
      };
      add("정산금액 합계", v.amount);
      add("우리 마진", v.margin);
      if (v.ded) add("CS 차감", -v.ded, "FFCC0000");
      add("지급액", v.final, "FF1A56DB");
      if (v.dedRows && v.dedRows.length) {
        ws.addRow([]);
        ws.addRow(["[CS 차감 내역]"]).font = { bold: true };
        ws.addRow(["접수일", "유형", "주문번호", "상품명", "내용", "비용"]).font = { bold: true };
        v.dedRows.forEach(x => ws.addRow([x.date, x.type, x.orderNo, x.product,
          String(x.content || "").slice(0, 120), Math.round(Number(x.cost) || 0)]));
      }
      ws.columns.forEach((c, i) => { c.width = [12, 12, 20, 34, 18, 7, 13, 13, 13][i] || 14; });
      ws.getColumn(7).numFmt = ws.getColumn(8).numFmt = ws.getColumn(9).numFmt = "#,##0";
    }
    return await wb.xlsx.writeBuffer();
  }

  async function sendOne(v, inp, btn) {
    const to = parseEmails(inp.value);
    if (!to.length) { alert("받는사람 이메일을 입력해주세요."); return; }
    const old = btn.textContent;
    btn.disabled = true; btn.textContent = "보내는 중…";
    try {
      await ensureGmail();
      const buf = await buildExcel([v]);
      const body = `안녕하세요, ${CONFIG.company}입니다.\n\n${QO.fmtDate(QO.todayStr())} 기준 정산 내역을 보내드립니다.\n\n`
        + `· 건수: ${v.rows.length}건\n· 정산금액 합계: ${won(v.amount)}\n`
        + (v.ded ? `· CS 차감(교환·반품): -${won(v.ded)}\n` : "")
        + `· 지급액: ${won(v.final)}\n\n자세한 내역은 첨부 파일을 확인해주세요.\n감사합니다.`;
      await GMAIL.send({
        to: to.join(", "),
        subject: `[${CONFIG.company}] 정산서 ${QO.fmtDate(QO.todayStr())} - ${v.vendor}`,
        body,
        attachments: [{ filename: fileName(v.vendor), data: buf }],
      });
      S.vendorSent[v.vendor] = mergeRecent(S.vendorSent[v.vendor], to);
      await DB.set("vendorSent", S.vendorSent);
      btn.textContent = "✔ 보냈어요";
      setTimeout(() => { btn.textContent = old; btn.disabled = false; }, 2500);
    } catch (e) {
      alert("발송 실패: " + e.message);
      btn.textContent = old; btn.disabled = false;
    }
  }

  /* CS 차감건을 '정산 반영됨'으로 표시 → 다음 정산에서 또 빠지지 않게 */
  async function markSettled() {
    if (!result) return;
    const all = [];
    result.vendors.forEach(v => (v.dedRows || []).forEach(x => all.push(x)));
    if (!all.length) { alert("차감된 CS 건이 없습니다."); return; }
    if (!confirm(`CS ${all.length}건을 '정산 반영 완료'로 표시할까요?\n다음 정산부터는 차감되지 않습니다.`)) return;
    all.forEach(x => { x.settled = true; x.updatedAt = Date.now(); });
    await CS.save(); CS.draw();
    calc();
    msg("msg-s", "ok", `✔ CS ${all.length}건을 정산 반영 완료로 표시했습니다.`);
  }

  function refresh() { $("run-s").disabled = !files.length; }

  /* =================================================================
     이벤트
     ================================================================= */
  async function drawFilterLine() {
    const el = $("st-filter");
    if (!el) return;
    const kw = await DB.get("stKeywords", ["정산", "정산내역", "지급"]);
    el.textContent = "정산 검색: " + (kw.join(", ") || "(없음)");
  }

  function bind() {
    $("f-st").addEventListener("change", async function () {
      const fs = [...this.files]; this.value = "";
      for (const f of fs) {
        $("st-fname").textContent = "📄 " + f.name;
        try { await addFile(await readFile(f), f.name); }
        catch (e) { msg("msg-s", "err", "⚠ " + f.name + " — " + e.message); }
      }
    });
    bindDrop("drop-st", async fs => {
      for (const f of fs) {
        $("st-fname").textContent = "📄 " + f.name;
        try { await addFile(await readFile(f), f.name); }
        catch (e) { msg("msg-s", "err", "⚠ " + f.name + " — " + e.message); }
      }
    });
    $("st-drive").onclick = () => openDrivePicker({
      key: "settle", multiple: true, title: "드라이브에서 정산 파일 가져오기",
      onPick: async fs => {
        for (const f of fs) {
          try { const r = await GMAIL.driveFetchExcel(f.id); await addFile(r.buf, r.name || f.name); }
          catch (e) { msg("msg-s", "err", "⚠ " + f.name + " — " + e.message); }
        }
      },
    });
    $("st-mail").onclick = async () => {
      try { await ensureGmail(); } catch (e) { return; }
      msg("msg-s", "", "");
      try {
        const kw = await DB.get("stKeywords", ["정산", "정산내역", "지급"]);
        const sd = await DB.get("stSenders", []);
        const items = await GMAIL.listTextMails({ days: 30, keywords: kw, senders: sd, max: 30 });
        const withAtt = items.filter(m => m.atts.length);
        if (!withAtt.length) { msg("msg-s", "warn", "최근 30일 메일에서 정산 엑셀 첨부를 찾지 못했어요."); return; }
        let n = 0;
        for (const m of withAtt.slice(0, 5)) {
          for (const a of m.atts) {
            const buf = await GMAIL.getAttachment(m.id, a.attachmentId);
            await addFile(buf, a.filename); n++;
          }
        }
        msg("msg-s", "ok", `✔ 메일에서 ${n}개 파일을 불러왔어요.`);
      } catch (e) { msg("msg-s", "err", "⚠ " + e.message); }
    };
    $("st-default").onchange = async () => {
      rules.def = num($("st-default").value);
      await saveRules(); drawRules();
      if (result) calc();
    };
    $("run-s").onclick = async () => {
      const b = $("run-s");
      b.classList.add("loading"); b.disabled = true;
      try { calc(); msg("msg-s", "ok", "✔ 계산했습니다. 아래에서 업체별 지급액을 확인하세요."); }
      catch (e) { msg("msg-s", "err", "⚠ " + e.message); }
      finally { b.classList.remove("loading"); b.disabled = false; }
    };
    $("st-mark").onclick = () => markSettled();
    $("st-export-all").onclick = async () => {
      if (!result) return;
      const buf = await buildExcel(result.vendors);
      download(buf, `${QO.todayStr()}_${CONFIG.company}_정산표_전체.xlsx`);
    };
  }

  async function init() { bind(); await load(); drawFilterLine(); }
  async function onShow() {
    if (!drawn) { await load(); drawn = true; }
    drawFiles(); drawRules(); refresh();
    if (result) drawResult();
  }

  init();
  return { onShow, drawFilter: drawFilterLine, markSettled, calc, result: () => result };
})();
window.ST = ST;
