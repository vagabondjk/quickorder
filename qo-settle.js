/* =====================================================================
   퀵오더 ④ 정산  (v6.1)
   주문 파일에서 상품마다 업체를 찾아 '업체 지급액'을 만든다.
     · 공급가표 방식(기본) : 지급액 = 공급단가 × 수량 + 배송비   ← 업체별 공급가표에서 단가를 끌어옴
     · 마진율(%) 방식      : 지급액 = 매출 × (1 − 마진율/100)
     · 공급단가 방식       : 지급액 = 수량 × 업체별 단가 하나
   매출은 '우리가 쇼핑몰에 넘긴 개당 값 × 수량'이다.
   (랩노마드 통합 파일에서는 그 열 이름이 '원가' — 우리 원가가 아니라 우리 매출이다)
   교환/반품 비용(③ CS에서 귀책 '업체' + 비용)은 업체별로 자동 차감한다.
   ★ 업체로 나가는 정산서에는 매출·마진이 절대 들어가면 안 된다 (buildExcel 참고).
   ===================================================================== */
"use strict";

const ST = (() => {
  const FIELDS = [
    // 정산 기준일 = 주문일. 행사 공급가 적용기간도 이 날짜로 따진다.
    { k: "date", n: "정산일(주문일)", kw: ["주문일시", "주문일자", "주문일", "정산일", "정산기준일", "결제일", "구매확정일", "수집일자", "일자", "날짜"] },
    // ★ 송장이 있어야 정산에 포함한다 (출고된 건만). 이 열을 지정하면 빈 줄은 정산에서 빠진다.
    { k: "invoice", n: "송장번호", kw: ["운송장번호", "운송장", "송장번호", "송장"] },
    // '쇼핑몰명'을 앞에 둔다 — 뒤에 두면 '판매가(쇼핑몰)' 열이 먼저 걸려 쇼핑몰로 잡힌다
    { k: "mall", n: "쇼핑몰", kw: ["쇼핑몰명", "판매처", "마켓명", "채널명", "쇼핑몰", "마켓", "채널", "사이트", "몰"] },
    { k: "orderNo", n: "주문번호", kw: ["주문번호", "주문no", "오더번호", "결제번호", "order"] },
    { k: "brand", n: "브랜드", kw: ["브랜드", "제조사", "공급사", "벤더", "업체"] },
    { k: "product", n: "상품명", kw: ["상품명", "제품명", "상품", "품목", "노출상품명"], req: true },
    { k: "option", n: "옵션", kw: ["옵션", "옵션명", "선택옵션", "규격"] },
    { k: "qty", n: "수량", kw: ["수량", "판매수량", "개수"] },
    // ★ 우리 매출 = 우리가 쇼핑몰에 넘긴 개당 값. 매출 = 이 단가 × 수량.
    //   랩노마드 통합 파일에서는 이 열 이름이 '원가' 다. 이름과 달리 우리 원가가 아니라
    //   '쇼핑몰에 공급가로 기입하는 가격'(우리 매출)이다 — 그래서 맨 앞에 둔다.
    //   '판매가(쇼핑몰)'은 몰이 소비자에게 파는 값이라 우리 매출이 아니다. 헷갈리지 말 것.
    { k: "unitPrice", n: `${(typeof CONFIG !== "undefined" && CONFIG.company) || "우리"} 매출(개당)`,
      kw: ["원가", "공급가(쇼핑몰)", "쇼핑몰공급가", "공급단가", "공급가", "판매단가"] },
    // 몰이 소비자에게 판 값(결제금액). 위 '매출(개당)'이 없을 때만 매출 대신 쓴다.
    { k: "amount", n: "쇼핑몰 매출", kw: ["정산금액", "정산예정금액", "지급액", "정산대금", "실정산액", "정산", "금액"] },
  ];

  /* 업체별 공급가표 — 업체에 '지급할' 상품별 단가. 통합 파일엔 없어서 따로 올린다. */
  const PB_FIELDS = [
    { k: "vendor", n: "업체명", kw: ["업체명", "업체", "벤더", "거래처"] },
    { k: "brand", n: "브랜드", kw: ["브랜드명", "브랜드", "제조사", "공급사"] },
    { k: "product", n: "상품명", kw: ["상품명", "제품명", "모델명", "상품", "품목"], req: true },
    { k: "option", n: "옵션", kw: ["옵션", "옵션명", "규격", "색상"] },
    { k: "price", n: "공급단가", kw: ["공급단가", "공급가", "정산단가", "매입가", "원가", "단가", "금액"], req: true },
    { k: "ship", n: "배송비", kw: ["배송비"] },
    { k: "from", n: "적용시작일", kw: ["행사시작일", "적용시작일", "적용일", "시작일", "변경일"] },
    { k: "to", n: "적용종료일", kw: ["행사종료일", "적용종료일", "종료일", "마감일"] },
  ];
  /* 업체별 배송비 방식만 적힌 작은 시트 (업체명 + 배송비정산) */
  const PB_SHIP_FIELDS = [
    { k: "vendor", n: "업체명", kw: ["업체명", "업체", "벤더", "거래처"], req: true },
    { k: "shipMode", n: "배송비 정산", kw: ["배송비정산", "배송비 정산 방법", "배송비방식", "정산방법", "정산방식"], req: true },
  ];

  let files = [];        // [{name, cols, rows, map, sig}]
  let maps = {};
  let result = null;     // 계산 결과
  let useCs = true;
  let drawn = false;
  let pbRaw = null;      // {name, at, sheets:[{name,kind,map,rows}], off:[시트명]}
  let pbook = { items: [], errors: [] };
  let aliases = {};      // 연결표 — {주문키: 공급가표키}. 이름이 달라 자동으로 못 붙는 상품을 이어준다
  let brandFix = {};     // 브랜드 → 업체 수동 배정 (공급가표 자동 배정보다 우선). 한 번 정하면 기억한다
  let extraVendors = []; // 공급가표에 없어도 직접 만든 업체
  let carry = { at: 0, list: [] };   // 지난 정산에서 송장이 없어 뺀 주문 (이월 추적용)
  let names = {};        // 업체 → 저장·발송할 파일명 (수정하면 그대로 씀)

  const s = v => (v === null || v === undefined) ? "" : String(v).trim();
  const num = v => {
    const n = Number(String(v == null ? "" : v).replace(/[^0-9.\-]/g, ""));
    return isFinite(n) ? n : 0;
  };
  const won = n => (Math.round(n || 0)).toLocaleString("ko-KR") + "원";

  async function load() {
    maps = await DB.get("settleMaps", {}) || {};
    extraVendors = await DB.get("settleVendors", []) || [];
    pbRaw = await DB.get("priceBook", null);
    aliases = await DB.get("priceAliases", {}) || {};
    brandFix = await DB.get("settleBrandVendor", {}) || {};
    // 지난 정산에서 '송장 없어 뺀' 목록. 이번에 출고됐으면 이월 건으로 알려준다.
    carry = await DB.get("settleCarry", { at: 0, list: [] }) || { at: 0, list: [] };
    if (!carry.list) carry.list = [];
    rebuildBook();
  }
  /* 이번 정산의 미출고 목록을 남겨둔다 (다음 달에 이월 여부를 알려주기 위해).
     화면에 쓰는 기준값(carry)은 그대로 둬서, 같은 세션에서 다시 계산해도 비교 대상이 흔들리지 않는다. */
  async function saveCarry(unshipped) {
    const list = unshipped.filter(r => s(r.orderNo)).map(r => ({
      orderNo: s(r.orderNo), brand: r.brand || "", product: (r.product || "").slice(0, 60),
      qty: r.qty || 0, date: r.date || "",
    }));
    await DB.set("settleCarry", { at: Date.now(), list });
  }
  async function saveBrandFix() { await DB.set("settleBrandVendor", brandFix); }
  function rebuildBook() {
    if (!pbRaw || !pbRaw.sheets) { pbook = QO.buildPriceBook([]); return; }
    const off = pbRaw.off || [];
    const shipModes = {}, rows = [];
    pbRaw.sheets.forEach(sh => {
      if (off.indexOf(sh.name) >= 0) return;
      const g = (r, k) => (sh.map[k] === undefined ? "" : r[sh.map[k]]);
      if (sh.kind === "ship") {
        sh.rows.forEach(r => { const v = s(g(r, "vendor")); if (v) shipModes[v] = s(g(r, "shipMode")); });
      } else {
        sh.rows.forEach(r => rows.push({
          vendor: s(g(r, "vendor")), brand: s(g(r, "brand")), product: s(g(r, "product")),
          option: s(g(r, "option")), price: g(r, "price"), ship: g(r, "ship"),
          from: g(r, "from"), to: g(r, "to"), _sheet: sh.name,
        }));
      }
    });
    pbook = QO.buildPriceBook(rows, { shipModes });
    pbook.shipModes = shipModes;
  }
  const hasBook = () => pbook.items.length > 0;

  /* =================================================================
     업체별 공급가표
     ================================================================= */
  /* 공급가 파일은 시트가 여러 장이다 (운영 / 행사 / 배송비 방식 / 참고자료).
     시트마다 열 위치가 달라서 한 번의 매핑을 돌려 쓸 수 없다 — 시트별로 따로 맞춘다.
     참고용 시트가 섞여 들어갈 수 있으니 어떤 시트를 읽었는지 화면에 보여주고 끌 수 있게 한다. */
  async function addPriceBook(buf, name) {
    const wb = await QO.loadWorkbook(buf.slice(0));
    const sheets = QO.previewSheets(wb, 20000);
    if (!sheets.length) throw new Error("시트를 찾지 못했어요.");

    const parsed = [];
    for (const sh of sheets) {
      if (!sh.columns.length || !sh.rows.length) continue;
      const pm = MAP.autoMap(sh.columns, PB_FIELDS);
      const sm = MAP.autoMap(sh.columns, PB_SHIP_FIELDS);
      const isPrice = pm.product !== undefined && pm.price !== undefined;
      const isShip = !isPrice && sm.vendor !== undefined && sm.shipMode !== undefined;
      if (!isPrice && !isShip) continue;
      parsed.push({ name: sh.name, kind: isPrice ? "price" : "ship",
                    map: isPrice ? pm : sm, rows: sh.rows, count: sh.rows.length });
    }
    if (!parsed.some(p => p.kind === "price"))
      throw new Error("상품명과 공급단가가 있는 시트를 찾지 못했어요.");

    // 이전에 껐던 시트는 기억한다
    const off = (pbRaw && pbRaw.name === name && pbRaw.off) || [];
    // 업체명이 있는 가격 시트가 하나라도 있으면, 업체명 없는 가격 시트는 참고자료로 보고 꺼둔다.
    // (랩노마드 파일의 '최저가 기준 방식' 시트가 여기 걸린다 — 켜두면 엉뚱한 상품이 섞인다)
    const anyVendor = parsed.some(p => p.kind === "price" && p.map.vendor !== undefined);
    if (anyVendor) parsed.forEach(p => {
      if (p.kind === "price" && p.map.vendor === undefined && off.indexOf(p.name) < 0) off.push(p.name);
    });
    pbRaw = { name, at: Date.now(), sheets: parsed.map(p => ({
      name: p.name, kind: p.kind, count: p.count, map: p.map,
      rows: p.rows.map(r => r.slice()),
    })), off: off.slice() };
    await savePriceBook();
    drawPriceBook(); drawBrands(); refresh();
    if (result) calc();
    const bad = pbook.errors.length;
    msg("msg-pb", bad ? "warn" : "ok",
      `✔ ${name} — 시트 ${parsed.length}장에서 상품 ${pbook.items.length}개를 읽었어요.`
      + (bad ? ` (${bad}줄은 건너뛰었어요 — 아래 확인)` : ""));
  }

  async function savePriceBook() {
    await DB.set("priceBook", pbRaw);
    rebuildBook();
  }

  /* 정산 파일 ↔ 공급가표 매칭 결과 한 줄.
     다 맞으면 성공 표시만, 안 맞는 게 있으면 어떤 상품이 몇 건인지만 보여준다. */
  function matchBoxHtml() {
    if (!files.length) return `<div style="margin-top:6px;color:var(--muted)">정산 파일을 올리면 상품이 다 맞는지 확인해 드립니다.</div>`;
    const rows = allRows();
    if (!rows.length) return "";
    const miss = {};
    let ok = 0;
    rows.forEach(r => {
      const m = QO.matchPrice(pbook, r, aliases);
      if (m.ok) { ok++; return; }
      const k = (r.brand ? r.brand + " / " : "") + (r.product || "");
      (miss[k] = miss[k] || 0); miss[k]++;
    });
    const keys = Object.keys(miss).sort((a, b) => miss[b] - miss[a]);
    if (!keys.length)
      return `<div style="margin-top:8px;padding:9px;border:1px solid var(--ok);border-radius:8px;background:var(--ok-soft)">
        <b style="color:var(--ok)">✔ 정산 파일의 상품이 모두 공급가표에 있습니다</b>
        <span style="color:var(--muted);font-size:12px"> · ${ok}건</span></div>`;
    const total = keys.reduce((s2, k) => s2 + miss[k], 0);
    return `<div style="margin-top:8px;padding:9px;border:1.5px solid var(--danger);border-radius:8px">
      <b style="color:var(--danger)">⚠ 공급가표에 없는 상품 ${keys.length}종 · ${total}건</b>
      <div style="margin-top:5px;font-size:12.5px;line-height:1.7">${
        keys.slice(0, 10).map(k => `· ${esc(k.slice(0, 52))} <b>${miss[k]}건</b>`).join("<br>")}${
        keys.length > 10 ? `<br>· 외 ${keys.length - 10}종` : ""}</div>
      <div style="margin-top:5px;font-size:11.5px;color:var(--muted)">
        공급가표에 추가하거나, 아래 ‘정산내역 추출’ 후 화면에서 연결해주세요.</div></div>`;
  }

  function drawPriceBook() {
    const box = $("pb-state");
    if (!box) return;
    if (!hasBook()) {
      box.innerHTML = "공급가표를 올리면 상품별 지급 단가로 정산합니다.";
      return;
    }
    const off = pbRaw.off || [];
    const sheetList = (pbRaw.sheets || []).map(sh =>
      `<label style="display:flex;align-items:center;gap:6px;margin-top:3px;cursor:pointer">
         <input type="checkbox" class="pbsheet" data-n="${esc(sh.name)}"${off.indexOf(sh.name) < 0 ? " checked" : ""}>
         <span>${esc(sh.name)} <span style="color:var(--muted)">· ${sh.count}줄 · ${sh.kind === "ship" ? "배송비 방식" : "가격"}</span></span>
       </label>`).join("");
    box.innerHTML =
      `<b>📗 ${esc(pbRaw.name)}</b> — 상품 ${pbook.items.length}개` +
      matchBoxHtml() +
      `<details style="margin-top:8px"><summary style="cursor:pointer;color:var(--muted);font-size:12px">
         읽은 시트 ${(pbRaw.sheets || []).length}장 · 자세히</summary>${sheetList}</details>` +
      `<div style="margin-top:6px"><button class="minibtn" id="pb-clear" style="padding:6px 10px;font-size:12px">공급가표 지우기</button></div>`;
    box.querySelectorAll(".pbsheet").forEach(cb => cb.onchange = async () => {
      const nm = cb.dataset.n;
      const cur = (pbRaw.off || []).filter(x => x !== nm);
      if (!cb.checked) cur.push(nm);
      pbRaw.off = cur;
      await savePriceBook();
      drawPriceBook(); drawBrands();
      if (result) calc();
    });
    const cl = $("pb-clear");
    if (cl) cl.onclick = async () => {
      if (!confirm("공급가표를 지울까요?\n지우면 업체 지급액은 마진율 방식으로 계산됩니다.")) return;
      pbRaw = null; await DB.set("priceBook", null);
      rebuildBook(); drawPriceBook(); drawBrands(); if (result) calc();
      msg("msg-pb", "ok", "공급가표를 지웠어요.");
    };
  }

  /* 빈 양식 내려받기 — 뭘 적어야 하는지 바로 보이게 예시를 한 줄 넣는다 */
  async function priceBookTemplate() {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("공급가표");
    ws.addRow(["브랜드", "상품명", "옵션", "공급단가(부가세 포함)", "적용시작일"]);
    ws.getRow(1).font = { bold: true };
    ws.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEFF3FB" } };
    ws.addRow(["수피아토", "수피아토 쿨타월", "", 7370, ""]);
    ws.addRow(["수피아토", "수피아토 쿨타월", "", 8730, "2026-07-13"]);
    ws.addRow(["현우동", "카레 우동 밀키트 500g x 3팩", "", 21870, ""]);
    ws.addRow([]);
    ws.addRow(["※ 공급단가는 부가세가 포함된 금액입니다. 지급액 = 공급단가 × 수량 으로 그대로 계산합니다."]);
    ws.addRow(["※ 상품명은 주문 파일과 똑같지 않아도 됩니다. 모델명처럼 알아볼 수 있는 이름이면 찾아냅니다."]);
    ws.addRow(["※ 옵션이 비어 있으면 그 상품의 모든 옵션에 같은 단가를 적용합니다."]);
    ws.addRow(["※ 단가가 바뀐 적이 있으면 위 쿨타월처럼 줄을 나누고 적용시작일을 적어주세요."]);
    ws.addRow(["※ 적용시작일이 비어 있으면 '언제나' 적용됩니다."]);
    [14, 40, 12, 12, 14].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
    ws.getColumn(4).numFmt = "#,##0";
    return await QO.saveWorkbook(wb);
  }

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
      drawFiles(); drawPriceBook(); drawBrands(); refresh();
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
      drawFiles(); drawPriceBook(); drawBrands(); refresh();
    });
  }

  /* =================================================================
     업체 목록 / 마진 규칙
     ================================================================= */
  let skipped = [];       // 파일에는 있는데 정산에 안 담긴 줄 (검산용)
  function allRows() {
    const out = [];
    skipped = [];
    for (const f of files) {
      const g = (r, k) => (f.map[k] === undefined ? "" : r[f.map[k]]);
      for (const r of f.rows) {
        const product = s(g(r, "product"));
        const option = s(g(r, "option"));
        const qty = num(g(r, "qty")) || 1;
        // 개당 단가가 있으면 그걸 × 수량. (통합 파일의 '판매가(쇼핑몰)'이 개당 단가다)
        const unitPrice = f.map.unitPrice === undefined ? null : num(g(r, "unitPrice"));
        const amount = unitPrice ? unitPrice * qty : num(g(r, "amount"));
        if (!product && !amount) { skipped.push({ src: f.name, raw: r }); continue; }
        // 브랜드는 파일에 열이 있으면 그걸 그대로 쓴다(상품명에서 추측하는 것보다 정확).
        const brand = (f.map.brand === undefined ? "" : s(g(r, "brand")))
          || CS.findBrand(product, option) || "";
        const vendor = (brand && (S.brandVendor || {})[brand])
          || CS.findVendor(product, option) || "";
        out.push({
          src: f.name,
          srcCols: f.cols, raw: r,     // 업체용 정산서를 원본 양식 그대로 뽑기 위해 들고 다닌다
          date: CS.toYmd(g(r, "date")),
          mall: s(g(r, "mall")) || guessMall(f.name),
          orderNo: s(g(r, "orderNo")),
          product, option, qty,
          unitPrice: unitPrice || null,
          amount,
          vendor, brand,
          invoice: f.map.invoice === undefined ? null : s(g(r, "invoice")),
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
  /* 업체명 정하기.
     ① 공급가표에서 찾은 상품의 '업체명'  ② 브랜드로 되짚은 업체(단가를 못 찾은 줄도 같은 정산서에 남게)
     ③ 발주 탭의 브랜드-업체 지정  ④ 그래도 없으면 브랜드 이름 그대로 */
  function brandVendorMap() {
    const bv = {};
    pbook.items.forEach(it => { if (it.b && it.vendor && !bv[it.b]) bv[it.b] = it.vendor; });
    return bv;
  }
  const vOf = r => r.vendor || r.brand || "(업체 미지정)";
  /* 공급가표를 태워 업체를 확정한다. 업체가 정해져야 계산 방식을 고를 수 있어서 계산보다 먼저 한다. */
  function resolveVendors(rows) {
    const bv = hasBook() ? brandVendorMap() : {};
    const valid = vendorNames();          // 업체명은 공급가표 기준
    rows.forEach(r => {
      r.m = hasBook() ? QO.matchPrice(pbook, r, aliases) : null;
      // 수동 배정이 먼저다 — 단, 업체 목록에 있는 이름일 때만 (브랜드 이름이 업체로 새는 걸 막는다)
      const fixed = r.brand && brandFix[r.brand];
      const fromBook = r.m && r.m.ok && r.m.item && r.m.item.vendor;
      r.vendor = (fixed && valid.indexOf(fixed) >= 0 ? fixed : "")
        || fromBook || bv[QO.normPriceText(r.brand)] || "";
      // ※ 브랜드로 되돌리지 않는다. 업체를 모르면 '미지정'으로 남겨 화면에서 배정하게 한다.
    });
  }

  /* =================================================================
     업체별 브랜드 선택 — 발주 탭과 같은 방식. 바꾸면 기억한다.
     ================================================================= */
  /* 업체 목록 — 공급가표의 '업체명' 열이 기준이다 (+ 직접 추가한 업체).
     ※ 브랜드 배정값은 여기 넣지 않는다. 넣으면 잘못 배정된 브랜드 이름이
       그대로 업체로 올라와서 '마이푸드메이트' 같은 브랜드가 업체처럼 보인다. */
  function vendorNames() {
    const out = [];
    const add = v => { v = String(v || "").trim(); if (v && out.indexOf(v) < 0) out.push(v); };
    pbook.items.forEach(it => add(it.vendor));
    extraVendors.forEach(add);
    return out.sort();
  }
  /* 브랜드가 어느 업체 것인지 — 직접 지정 > 공급가표.
     직접 지정이 업체 목록에 없는 이름이면(표가 바뀌었거나 잘못 눌렀거나) 무시하고 자동 배정으로 돌아간다. */
  function ownerOf(brand) {
    const fixed = brandFix[brand];
    if (fixed && vendorNames().indexOf(fixed) >= 0) return fixed;
    const bv = brandVendorMap();
    return bv[QO.normPriceText(brand)] || "";
  }

  function drawBrands() {
    const card = $("st-card-brand"), box = $("st-vbrands");
    if (!card) return;
    const rows = allRows();
    if (!rows.length) { card.style.display = "none"; box.innerHTML = ""; return; }
    card.style.display = "block";

    const brands = [], cnt = {};
    rows.forEach(r => {
      const b = r.brand || "(브랜드 없음)";
      if (brands.indexOf(b) < 0) brands.push(b);
      cnt[b] = (cnt[b] || 0) + 1;
    });
    brands.sort();
    const vendors = vendorNames();
    const addBtn = `<button class="minibtn" id="st-add-vendor" style="margin-top:8px">＋ 업체 추가</button>`;

    if (!vendors.length) {
      box.innerHTML = `<div class="pvfoot">공급가표를 올리면 업체가 자동으로 잡힙니다.
        먼저 업체를 직접 만들어도 됩니다.</div>${addBtn}`;
      $("st-brand-foot").textContent = `브랜드 ${brands.length}개: ` + brands.join(", ");
      bindAddVendor();
      return;
    }

    box.innerHTML = vendors.map(v => {
      const chips = brands.map(b => {
        const owner = ownerOf(b);
        const mine = owner === v;
        const fixed = brandFix[b] === v;
        if (mine) return `<span class="brow on" data-b="${esc(b)}" data-v="${esc(v)}">` +
          `<span class="box">✓</span>${esc(b)} <small>· ${cnt[b]}건${fixed ? " · 직접" : ""}</small></span>`;
        if (owner) return `<span class="brow taken" data-b="${esc(b)}" data-v="${esc(v)}" title="${esc(owner)} 것">` +
          `${esc(b)} <small>· ${esc(owner)}</small></span>`;
        return `<span class="brow" data-b="${esc(b)}" data-v="${esc(v)}">` +
          `<span class="box"></span>${esc(b)} <small>· ${cnt[b]}건</small></span>`;
      }).join("");
      const mineCnt = brands.filter(b => ownerOf(b) === v).length;
      const del = extraVendors.indexOf(v) >= 0 ? `<button class="vdel" data-v="${esc(v)}">✕</button>` : "";
      return `<div class="vendorbox"><div class="vh">🏭 ${esc(v)}<span class="cnt">브랜드 ${mineCnt}개</span>${del}</div>
        <div class="brands">${chips}</div></div>`;
    }).join("") + addBtn;

    box.querySelectorAll(".brow").forEach(el => el.onclick = async () => {
      const b = el.dataset.b, v = el.dataset.v;
      if (brandFix[b] === v) delete brandFix[b];   // 다시 누르면 자동 배정으로
      else brandFix[b] = v;                        // 남의 것이어도 눌러서 가져올 수 있다
      await saveBrandFix();
      drawBrands(); if (result) calc();
    });
    box.querySelectorAll(".vdel").forEach(el => el.onclick = async e => {
      e.stopPropagation();
      const v = el.dataset.v;
      if (!confirm(`업체 '${v}' 를 목록에서 지울까요?\n(배정해둔 브랜드는 자동 배정으로 돌아갑니다)`)) return;
      extraVendors = extraVendors.filter(x => x !== v);
      Object.keys(brandFix).forEach(b => { if (brandFix[b] === v) delete brandFix[b]; });
      await DB.set("settleVendors", extraVendors); await saveBrandFix();
      drawBrands(); if (result) calc();
    });
    bindAddVendor();

    const none = brands.filter(b => !ownerOf(b));
    $("st-brand-foot").textContent = none.length
      ? `아직 업체가 정해지지 않은 브랜드: ${none.join(", ")} — 해당 업체 칸에서 눌러 배정하세요.`
      : "브랜드를 눌러 다른 업체로 옮길 수 있습니다. 옮기면 기억합니다.";
  }
  function bindAddVendor() {
    const b = $("st-add-vendor");
    if (!b) return;
    b.onclick = async () => {
      const v = prompt("업체명을 입력하세요.\n(예: 플라스머, 디에스피)");
      if (v === null) return;
      const t = v.trim();
      if (!t || extraVendors.indexOf(t) >= 0) return;
      extraVendors.push(t);
      await DB.set("settleVendors", extraVendors);
      drawBrands();
    };
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
    resolveVendors(rows);            // 업체 먼저 (공급가표의 업체명으로 묶기 위해)
    const ded = csDeduction();
    const byVendor = {};
    const unpriced = [];       // 공급가표에서 단가를 못 찾은 줄
    // ★ 송장이 없으면 아직 출고 전이라 이번 정산에서 뺀다 (송장 열을 지정한 경우에만).
    //   조용히 빼면 안 되므로 따로 세어 검산과 화면에 띄운다.
    const hasInvCol = files.some(f => f.map.invoice !== undefined);
    const unshipped = hasInvCol ? rows.filter(r => !String(r.invoice || "").trim()) : [];
    const shipped = hasInvCol ? rows.filter(r => String(r.invoice || "").trim()) : rows;
    // 지급액은 언제나 공급가표 기준이다 — 공급단가 × 수량 + 배송비.
    for (const r of shipped) {
      const v = vOf(r);
      r.mode = "book";
      r.unitCost = null; r.priced = true; r.why = ""; r.how = "";
      r.ship = 0; r.shipMode = "개당"; r.shipTotal = 0;
      const m = r.m || QO.matchPrice(pbook, r, aliases);
      if (m.ok) {
        const it = m.item || {};
        r.unitCost = Math.round(m.price); r.how = m.how; r.priceFrom = m.from || "";
        r.ship = Math.round(it.ship || 0);
        r.shipMode = it.shipMode || "개당";
        r.shipTotal = r.ship * (r.shipMode === "건당" ? (r.qty > 0 ? 1 : 0) : r.qty);
        // 단가는 줄마다 원 단위로 반올림한 뒤 수량을 곱한다
        r.pay = r.unitCost * r.qty + r.shipTotal;
      } else {
        // 못 찾은 줄은 0 원으로 두되 따로 모아 화면에 띄운다. 조용히 넘기지 않는다.
        r.priced = false; r.why = m.why; r.pay = 0;
        unpriced.push(r);
      }
      // 단가를 못 찾은 줄은 마진도 0. (지급액 0 원을 그대로 마진으로 잡으면 마진이 부풀려진다)
      r.margin = r.priced ? r.amount - r.pay : 0;
      const g = byVendor[v] = byVendor[v] ||
        { vendor: v, rows: [], amount: 0, pay: 0, margin: 0, ded: 0, dedRows: [],
          unpriced: 0, unpricedAmount: 0, loose: 0, ship: 0, brands: [] };
      g.rows.push(r); g.amount += r.amount; g.pay += r.pay; g.margin += r.margin;
      g.ship += r.shipTotal || 0;
      if (r.brand && g.brands.indexOf(r.brand) < 0) g.brands.push(r.brand);
      if (!r.priced) { g.unpriced++; g.unpricedAmount += r.amount; }
      if (r.how === "앞부분일치" || r.how === "핵심어일치") g.loose++;
    }
    for (const v in byVendor) {
      const d = ded[v];
      if (d) { byVendor[v].ded = d.amount; byVendor[v].dedRows = d.rows; }
      byVendor[v].final = byVendor[v].pay - byVendor[v].ded;
    }
    const vendors = Object.values(byVendor).sort((a, b) => b.final - a.final);
    const sum = f => vendors.reduce((s2, v) => s2 + v[f], 0);
    result = {
      vendors,
      total: {
        amount: sum("amount"), pay: sum("pay"), margin: sum("margin"),
        ded: sum("ded"), final: sum("final"), count: rows.length,
        unpricedAmount: sum("unpricedAmount"),
      },
      noVendor: rows.filter(r => !r.vendor && !r.brand).length,
      unpriced,
      loose: sum("loose"),
      usedBook: vendors.some(v => v.rows.some(r => r.mode === "book")),
    };
    result.check = reconcile(rows, shipped, unshipped, vendors, unpriced);
    result.unshipped = unshipped;
    result.period = periodOf(shipped.length ? shipped : rows);
    drawResult();
  }

  /* =================================================================
     검산 — 올린 주문이 하나도 빠짐없이 업체 정산에 담겼는지 확인한다.
     금액이 작다고 넘어가면 안 된다. 한 건이라도 어긋나면 빨간 상자로 띄운다.
     ================================================================= */
  function reconcile(rows, shipped, unshipped, vendors, unpriced) {
    const issues = [];
    const err = (why, detail) => issues.push({ level: "error", why, detail: detail || "" });
    const warn = (why, detail) => issues.push({ level: "warn", why, detail: detail || "" });

    // ① 파일에서 읽은 줄 수 (열 맞추기까지 끝난 미리보기 기준)
    const fileRows = files.reduce((s, f) => s + f.rows.length, 0);
    const inCount = rows.length;
    const inQty = rows.reduce((s, r) => s + (r.qty || 0), 0);
    const shipCount = shipped.length;
    const shipQty = shipped.reduce((s, r) => s + (r.qty || 0), 0);
    const outCount = vendors.reduce((s, v) => s + v.rows.length, 0);
    const outQty = vendors.reduce((s, v) => s + (v.rows.reduce((t, r) => t + (r.qty || 0), 0)), 0);
    const outPay = vendors.reduce((s, v) => s + v.pay, 0);

    if (skipped.length)
      err(`올린 파일 ${fileRows}줄 중 ${skipped.length}줄이 정산에서 빠졌어요`,
          "상품명과 금액이 모두 비어 있는 줄입니다. 열 맞추기가 잘못됐을 수 있어요");
    // 송장이 없는 건은 아직 출고 전이라 빼는 게 정상이다 — 오류가 아니라 안내로 남긴다
    if (unshipped.length)
      warn(`송장이 없어 이번 정산에서 뺀 주문이 ${unshipped.length}건 (수량 ${unshipped.reduce((s, r) => s + (r.qty || 0), 0)}) 있어요`,
           "출고되면 다음 정산에 포함됩니다");

    /* 이월 추적 — 지난 정산에서 미출고였던 건이 이번에 출고됐는지.
       (7월 주문인데 8월에 출고된 건은 8월 정산으로 넘어와야 한다) */
    if (carry.list.length) {
      const nowNos = {};
      rows.forEach(r => { const k = s(r.orderNo); if (k) nowNos[k] = r; });
      const shipNos = {};
      shipped.forEach(r => { const k = s(r.orderNo); if (k) shipNos[k] = r; });
      const 이월됨 = carry.list.filter(c => shipNos[c.orderNo]);
      const 아직 = carry.list.filter(c => !nowNos[c.orderNo]);
      const 여전히미출고 = carry.list.filter(c => nowNos[c.orderNo] && !shipNos[c.orderNo]);
      const when = carry.at ? new Date(carry.at).toISOString().slice(0, 10) : "";
      if (이월됨.length)
        issues.push({ level: "info",
          why: `지난 정산(${when})에서 미출고였던 ${이월됨.length}건이 이번에 출고돼 포함됐어요`,
          detail: "이월 처리된 건입니다 — 지난달에 중복 지급하지 않았는지 확인하세요" });
      if (아직.length)
        warn(`지난 정산의 미출고 ${아직.length}건이 이번 파일에 아예 없어요`,
             `주문번호 ${아직.slice(0, 5).map(c => c.orderNo).join(", ")}${아직.length > 5 ? " 외 " + (아직.length - 5) + "건" : ""} — 아직 출고 전인지 확인하세요`);
      if (여전히미출고.length)
        warn(`지난 정산의 미출고 ${여전히미출고.length}건이 이번에도 송장이 없어요`, "출고가 계속 밀리고 있습니다");
    }
    // 출고된 건은 하나도 빠짐없이 담겨야 한다
    if (shipCount !== outCount)
      err(`출고 ${shipCount}건 중 ${outCount}건만 업체 정산에 담겼어요`, `${shipCount - outCount}건이 사라졌습니다`);
    if (shipQty !== outQty)
      err(`출고 수량 ${shipQty}개와 업체 정산 수량 ${outQty}개가 달라요`, `${shipQty - outQty}개 차이`);

    // ② 줄별 지급액을 다시 계산해 업체 합계와 맞는지 (출고분만)
    let recomputed = 0, badRows = 0;
    shipped.forEach(r => {
      const want = r.priced === false ? 0 : (r.unitCost || 0) * r.qty + (r.shipTotal || 0);
      if (Math.round(want) !== Math.round(r.pay || 0)) badRows++;
      recomputed += want;
    });
    if (badRows) err(`지급액이 단가 × 수량과 맞지 않는 줄이 ${badRows}건 있어요`);
    if (Math.round(recomputed) !== Math.round(outPay))
      err(`줄별 지급액 합계 ${won(recomputed)} 와 업체별 합계 ${won(outPay)} 가 달라요`);

    // ③ 단가를 못 찾아 0 원으로 둔 건 — 정산이 덜 된 상태
    if (unpriced.length) {
      const q = unpriced.reduce((s, r) => s + r.qty, 0);
      const kinds = {};
      unpriced.forEach(r => { kinds[(r.brand || "") + "/" + r.product] = 1; });
      err(`단가를 못 찾아 0 원으로 둔 주문이 ${unpriced.length}건 (수량 ${q}) 있어요`,
          `상품 ${Object.keys(kinds).length}종 — 아래에서 연결해주면 해결됩니다`);
    }
    const zero = shipped.filter(r => !r.qty);
    if (zero.length) warn(`수량이 비어 있는 주문이 ${zero.length}건 있어요`, "그 줄은 지급액이 0 원입니다");
    const nov = shipped.filter(r => vOf(r) === "(업체 미지정)");
    if (nov.length) warn(`업체를 정하지 못한 주문이 ${nov.length}건 있어요`);

    return { ok: !issues.some(i => i.level === "error"), issues,
             fileRows, orderCount: inCount, orderQty: inQty,
             shipCount, shipQty, unshipped: unshipped.length,
             settledCount: outCount, settledQty: outQty, pay: outPay };
  }

  function checkBoxHtml(c) {
    if (!c) return "";
    const line = `주문 ${c.orderCount}건 · 수량 ${c.orderQty}개`
      + (c.unshipped ? `  →  출고 ${c.shipCount}건 · 수량 ${c.shipQty}개` : "")
      + `  →  업체 정산 ${c.settledCount}건 · 수량 ${c.settledQty}개`;
    if (c.ok && !c.issues.length)
      return `<div style="margin-top:10px;padding:10px;border:1px solid var(--ok);border-radius:8px;background:rgba(16,150,90,.06)">
        <b style="color:var(--ok)">✔ 검산 통과 — 주문이 하나도 빠지지 않았습니다</b>
        <div style="margin-top:4px;font-size:12px;color:var(--muted)">${esc(line)}</div></div>`;
    const bad = c.issues.some(i => i.level === "error");
    return `<div style="margin-top:10px;padding:10px;border:1.5px solid var(${bad ? "--danger" : "--warn"});border-radius:8px">
      <b style="color:var(${bad ? "--danger" : "--warn"})">${bad ? "⚠ 검산 실패 — 주문과 정산이 맞지 않습니다" : "⚠ 확인이 필요합니다"}</b>
      <div style="margin-top:4px;font-size:12px;color:var(--muted)">${esc(line)}</div>
      <div style="margin-top:6px;font-size:12.5px;line-height:1.8">${c.issues.map(i =>
        `<div style="color:var(${i.level === "error" ? "--danger" : "--warn"})">· ${esc(i.why)}${
          i.detail ? `<br><span style="color:var(--muted);font-size:11.5px">&nbsp;&nbsp;&nbsp;${esc(i.detail)}</span>` : ""}</div>`).join("")}</div>
    </div>`;
  }

  function drawResult() {
    if (!result) { $("result-s").style.display = "none"; return; }
    $("result-s").style.display = "block";
    const t = result.total;
    const up = result.unpriced || [];
    // 못 찾은 상품을 종류별로 묶어서 보여준다 (같은 상품이 수십 줄씩 나오므로)
    const upKinds = {};
    up.forEach(r => {
      const k = (r.brand ? r.brand + " / " : "") + r.product + (r.option ? " [" + r.option + "]" : "");
      const e = upKinds[k] = upKinds[k] || { n: 0, why: r.why, row: r };
      e.n++;
    });
    const kindKeys = Object.keys(upKinds);
    const per = result.period || { label: "" };
    $("st-total").innerHTML =
      (per.label ? `<div style="margin-bottom:6px;font-size:13px"><b>📅 정산월 ${esc(per.label)}</b>
         <span style="color:var(--muted);font-size:11.5px"> · 주문일 ${esc(QO.fmtDate(per.from))} ~ ${esc(QO.fmtDate(per.to))} · 작성일 ${esc(QO.fmtDate(QO.todayStr()))}</span></div>` : "") +
      `<div class="totline"><b>${result.usedBook ? "몰 매출 합계" : "몰 정산금액 합계"}</b><span>${won(t.amount)}</span></div>
       ${t.unpricedAmount ? `<div class="totline" style="font-size:12px;color:var(--danger)"><span>└ 단가 못 찾은 건 (지급·마진에서 빠짐)</span><span>${won(t.unpricedAmount)}</span></div>` : ""}
       <div class="totline"><b>${esc(CO())} 마진</b><span style="color:var(--ok)">${won(t.margin)}</span></div>
       ${t.ded ? `<div class="totline"><b>CS 차감 (교환·반품)</b><span style="color:var(--danger)">− ${won(t.ded)}</span></div>` : ""}
       <div class="totline" style="border-top:1px solid var(--line);margin-top:6px;padding-top:8px">
         <b>업체 지급 합계</b><span style="color:var(--brand)">${won(t.final)}</span></div>
       <div class="synchint" style="margin-top:6px">${t.count}건${
         result.noVendor ? ` · <b style="color:var(--danger)">업체 미지정 ${result.noVendor}건</b> — 상품명에서 브랜드를 못 찾았어요. ① 발주 탭에서 브랜드-업체를 한 번 지정하면 다음부터 자동입니다.` : ""}${
         result.loose ? ` · 이름이 조금 달라 비슷한 것으로 찾은 줄 ${result.loose}건` : ""}</div>` +
      checkBoxHtml(result.check) +
      (up.length ? unpricedBoxHtml(upKinds, kindKeys, up.length) : "");
    if (up.length) bindAliasPickers();

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
        <div class="totline" style="font-size:12px"><span>${result.usedBook ? "매출" : "정산금액"}</span><span>${won(v.amount)}</span></div>
        ${v.unpricedAmount ? `<div class="totline" style="font-size:12px;color:var(--danger)"><span>└ 단가 못 찾은 건</span><span>${won(v.unpricedAmount)}</span></div>` : ""}
        <div class="totline" style="font-size:12px"><span>${esc(CO())} 마진</span><span>${won(v.margin)}</span></div>
        ${v.unpriced ? `<div class="totline" style="font-size:12px;color:var(--danger)"><span>⚠ 단가 못 찾음</span><span>${v.unpriced}건 (0원 처리)</span></div>` : ""}
        ${v.loose ? `<div class="totline" style="font-size:12px;color:var(--muted)"><span>비슷한 이름으로 찾음</span><span>${v.loose}건</span></div>` : ""}
        ${v.ded ? `<div class="totline" style="font-size:12px;color:var(--danger)"><span>CS 차감 ${v.dedRows.length}건</span><span>− ${won(v.ded)}</span></div>` : ""}
        <div class="totline" style="font-size:13px"><b>지급액</b><span style="color:var(--brand)">${won(v.final)}</span></div>
        <div class="rmail"><input type="email" placeholder="업체 이메일 (쉼표로 여러 명)" autocapitalize="off" spellcheck="false">
          <button class="dlbtn">메일 보내기</button></div>
        <div class="cands"></div>
        <div class="setrow" style="margin-top:6px"><span style="flex:1;font-size:11px;color:var(--faint)">여러 명에게 보내려면 쉼표로 구분</span>
          <button class="minibtn share">📤 카톡·공유</button><button class="minibtn pvbtn">미리보기</button><button class="minibtn dl">엑셀만 받기</button></div>
        <div class="setrow" style="margin-top:4px"><span class="fnlbl" style="flex:1;font-size:11px;color:var(--faint)"></span>
          <button class="minibtn fn">✏️ 파일명 수정</button></div>
        <div class="setrow" style="margin-top:4px"><span style="flex:1;font-size:11px;color:var(--faint)">계산에 쓰인 줄을 표로 봅니다</span>
          <button class="minibtn pvv">👁 정산 내역</button></div>`;
      box.appendChild(row);
      const inp = row.querySelector("input");
      const known = S.vendorEmails[v.vendor] || "";
      inp.value = known;
      fillRecipients(row.querySelector(".cands"), inp, {
        saved: known, history: S.vendorSent[v.vendor] || [],
        domains: S.vendorDomains[v.vendor] || [], query: v.vendor !== "(업체 미지정)" ? v.vendor : "",
      });
      // 파일명은 자동으로 지어지되 여기서 바꿀 수 있다 (저장·공유·메일 첨부에 모두 반영)
      const lbl = row.querySelector(".fnlbl");
      const showName = () => { lbl.textContent = fileName(v.vendor); };
      showName();
      row.querySelector(".fn").onclick = () => {
        const cur = fileName(v.vendor).replace(/\.xlsx$/i, "");
        const t = prompt("저장·발송될 파일명을 바꿉니다.\n(확장자 .xlsx 는 자동으로 붙습니다)", cur);
        if (t === null) return;
        const clean = t.trim().replace(/[\\/:*?"<>|]/g, "").replace(/\.xlsx$/i, "");
        if (!clean) return;
        names[v.vendor] = clean + ".xlsx";
        showName();
      };
      const vendorBuf = async () => await buildExcel([v]);      // 업체용 — 매출·마진 없음
      const guard = async fn => {
        try { await fn(); }
        catch (e) { msg("msg-s", "err", "⚠ 정산서를 만들지 못했어요 — " + e.message); }
      };
      row.querySelector(".dl").onclick = () => guard(async () => download(await vendorBuf(), fileName(v.vendor)));
      row.querySelector(".share").onclick = () => guard(async () => shareFile(await vendorBuf(), fileName(v.vendor)));
      row.querySelector(".pvbtn").onclick = () => guard(async () => openPreview(await vendorBuf(), v.vendor + " 정산서"));
      row.querySelector(".pvv").onclick = () => showRows(v);
      row.querySelector(".dlbtn").onclick = () => sendOne(v, inp, row.querySelector(".dlbtn"));
    });
  }
  /* 저장·발송 파일명 — 앞에 정산월을 붙인다 (파일만 봐도 어느 달 정산인지 알게).
     사용자가 고쳤으면 그걸 쓴다. */
  const fileName = v => names[v] || `${periodTag()}_${CONFIG.company}_${cleanVendor(v)}_정산서.xlsx`;

  /* =================================================================
     연결표 — 이름이 달라 자동으로 못 붙는 상품을 사람이 한 번 이어준다
     ================================================================= */
  /* 공급가표에서 고를 수 있는 상품 목록 (같은 상품의 평시/행사 줄은 하나로 묶는다) */
  function bookChoices() {
    const seen = {}, out = [];
    pbook.items.forEach(it => {
      if (seen[it.key]) { seen[it.key].n++; return; }
      const e = { key: it.key, vendor: it.vendor, brand: it.brand, product: it.product,
                  option: it.option, price: it.price, n: 1 };
      seen[it.key] = e; out.push(e);
    });
    return out.sort((a, b) =>
      (a.brand || "").localeCompare(b.brand || "") || (a.product || "").localeCompare(b.product || ""));
  }
  /* 후보 목록 — 이름이 겹치는 것만 위에, 나머지는 접어서 뒤에.
     21개를 통째로 늘어놓으면 못 고른다 (닭갈비면 닭갈비만 보여야 한다) */
  function pickOptions(row) {
    const opt = c => `<option value="${esc(c.key)}">${esc((c.brand ? c.brand + " · " : "") + c.product)}` +
      `${c.option ? " [" + esc(c.option) + "]" : ""} — ${Math.round(c.price).toLocaleString("ko-KR")}원` +
      `${c.n > 1 ? " (기간별 " + c.n + "개)" : ""}</option>`;
    const all = bookChoices();
    const byKey = {}; all.forEach(c => { byKey[c.key] = c; });
    const ranked = QO.rankPriceCandidates(pbook, row);
    const top = ranked.map(x => byKey[x.item.key]).filter(Boolean);
    const topKeys = {}; top.forEach(c => { topKeys[c.key] = 1; });
    const rest = all.filter(c => !topKeys[c.key]);
    let h = '<option value="">— 공급가표에서 고르기 —</option>';
    if (top.length) h += `<optgroup label="이름이 비슷한 상품 ${top.length}개">` + top.map(opt).join("") + "</optgroup>";
    if (rest.length) h += `<optgroup label="${top.length ? "그 외 " : "전체 "}${rest.length}개">` + rest.map(opt).join("") + "</optgroup>";
    return h;
  }
  function unpricedBoxHtml(kinds, keys, total) {
    return `<div style="margin-top:10px;padding:10px;border:1px solid var(--danger);border-radius:8px">
      <b style="color:var(--danger)">⚠ 공급가표에서 단가를 못 찾은 ${total}건 — 지급액 0 원으로 뒀습니다</b>
      <div style="margin-top:4px;font-size:12px;color:var(--muted)">
        이름이 서로 달라서 그렇습니다. 아래에서 한 번만 이어주면 다음 달부터 자동입니다.</div>
      ${keys.map((k, i) => `
        <div style="margin-top:10px;padding-top:8px;border-top:1px solid var(--line)">
          <div style="font-size:12.5px"><b>${esc(kinds[k].row.product.slice(0, 60))}</b>
            ${kinds[k].row.option ? `<span style="color:var(--muted)"> [${esc(kinds[k].row.option)}]</span>` : ""}
            <span style="color:var(--danger)"> ${kinds[k].n}건</span></div>
          <div style="font-size:11.5px;color:var(--muted);margin-top:2px">${esc(kinds[k].row.brand || "")}</div>
          <select class="aliaspick" data-i="${i}" style="width:100%;margin-top:5px;padding:7px;font-size:12.5px">${pickOptions(kinds[k].row)}</select>
        </div>`).join("")}
    </div>`;
  }
  function bindAliasPickers() {
    const up = result.unpriced || [];
    const kinds = {};
    up.forEach(r => {
      const k = (r.brand ? r.brand + " / " : "") + r.product + (r.option ? " [" + r.option + "]" : "");
      (kinds[k] = kinds[k] || { row: r });
    });
    const keys = Object.keys(kinds);
    $("st-total").querySelectorAll(".aliaspick").forEach(sel => {
      sel.onchange = async () => {
        const r = kinds[keys[Number(sel.dataset.i)]].row;
        const okey = QO.priceRowKey(r.brand, r.product, r.option);
        if (sel.value) aliases[okey] = sel.value;
        else delete aliases[okey];
        await DB.set("priceAliases", aliases);
        calc();
        msg("msg-s", "ok", "✔ 연결했어요. 다시 계산했습니다.");
      };
    });
  }

  function showRows(v) {
    const cols = ["정산일", "쇼핑몰", "주문번호", "상품명", "옵션", "수량",
                  CO() + " 매출", "공급단가", CO() + " 마진", "지급액"];
    const rows = v.rows.map(r => [r.date, r.mall, r.orderNo, r.product, r.option, r.qty,
      Math.round(r.amount), r.priced === false ? "못 찾음" : (r.unitCost == null ? "" : Math.round(r.unitCost)),
      Math.round(r.margin), Math.round(r.pay)]);
    $("pv-modal-title").textContent = v.vendor + " 정산 내역";
    $("pv-modal-sub").textContent = `${rows.length}건 · 지급액 ${won(v.final)}`
      + (v.unpriced ? ` · ⚠ 단가 못 찾음 ${v.unpriced}건` : "");
    const tbl = $("pv-modal-table");
    tbl.innerHTML = `<thead><tr>${cols.map(c => `<th>${esc(c)}</th>`).join("")}</tr></thead>`
      + `<tbody>${rows.map(r => `<tr>${r.map((c, i) => `<td${i >= 5 ? ' class="num"' : ""}>${esc(c)}</td>`).join("")}</tr>`).join("")}</tbody>`;
    $("pv-modal-foot").textContent = "";
    $("pvmodal").classList.add("on");
  }

  /* =================================================================
     엑셀 / 메일
     ================================================================= */
  /* 정산서 엑셀.
     opts.internal 이 참일 때만 '매출·우리마진' 을 넣는다.
     ※ 업체로 나가는 파일에는 우리 매출과 마진이 절대 들어가면 안 된다.
       업체가 볼 것은 '무엇을 몇 개, 단가 얼마에, 얼마 받는지'까지다.
       업체별 저장/메일은 전부 업체용(internal 없음)으로만 부른다. */
  const CO = () => (typeof CONFIG !== "undefined" && CONFIG.company) || "우리";

  /* 정산월 — 주문일 범위에서 뽑는다. 한 달 안에 들어오면 '2026년 7월',
     달을 걸치면 '2026-07-28 ~ 2026-08-03' 으로 적는다. */
  function periodOf(rows) {
    const ds = [];
    (rows || []).forEach(r => {
      const d = String(r.date || "").replace(/\D/g, "").slice(0, 8);
      if (d.length === 8) ds.push(d);
    });
    if (!ds.length) return { from: "", to: "", label: "", tag: "" };
    ds.sort();
    const from = ds[0], to = ds[ds.length - 1];
    const f = d => d.slice(0, 4) + "-" + d.slice(4, 6) + "-" + d.slice(6, 8);
    if (from.slice(0, 6) === to.slice(0, 6))
      return { from, to, label: `${from.slice(0, 4)}년 ${Number(from.slice(4, 6))}월`,
               tag: from.slice(0, 4) + "-" + from.slice(4, 6) };
    return { from, to, label: `${f(from)} ~ ${f(to)}`,
             tag: from.slice(0, 4) + "-" + from.slice(4, 6) + "~" + to.slice(4, 6) };
  }
  const periodTag = () => (result && result.period && result.period.tag) || QO.todayStr().slice(0, 6);
  /* 파일 머리말에 공통으로 넣는 줄 */
  function stampLine(extra) {
    const p = (result && result.period) || { label: "" };
    return (p.label ? `정산월 ${p.label} · ` : "") + `작성일 ${QO.fmtDate(QO.todayStr())}` + (extra ? ` · ${extra}` : "");
  }

  /* 업체용 정산서 — 주문 통합 양식을 그대로 쓰되 가격 열만 걷어내고
     '업체→(회사) 공급가'와 정산금액을 붙인다. 업체가 자기 주문건과 바로 대조할 수 있게. */
  function vendorSheet(wb, v) {
    const safe = n => String(n).replace(/[\\\/\?\*\[\]:]/g, "_").slice(0, 28) || "정산";
    const ws = wb.addWorksheet(safe(v.vendor));
    const src = QO.vendorSheetColumns(
      [...new Set(v.rows.map(r => r.srcCols).filter(Boolean))]);
    const anyShip = v.rows.some(r => r.shipTotal);
    const head = src.concat([`업체→${CO()} 공급가`]).concat(anyShip ? ["배송비"] : []).concat(["정산금액"]);
    const nCol = head.length;

    const p = (result && result.period) || { label: "" };
    ws.addRow([`${v.vendor} 정산서${p.label ? " — " + p.label : ""}`]);
    ws.getRow(1).font = { bold: true, size: 14 };
    ws.addRow([stampLine(`${v.rows.length}건`)]);
    ws.addRow(["※ 아래 금액은 모두 부가세가 포함된 금액입니다."]);
    ws.addRow([]);
    ws.addRow(head);
    const hr = ws.lastRow.number;
    ws.getRow(hr).font = { bold: true };
    ws.getRow(hr).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEFF3FB" } };

    v.rows.forEach(r => {
      const idx = {};
      (r.srcCols || []).forEach((h, i) => { const t = String(h || "").trim(); if (t && idx[t] === undefined) idx[t] = i; });
      const line = src.map(h => (idx[h] === undefined ? "" : (r.raw ? r.raw[idx[h]] : "")));
      line.push(r.priced === false ? "미확정" : (r.unitCost == null ? "" : Math.round(r.unitCost)));
      if (anyShip) line.push(Math.round(r.shipTotal || 0));
      line.push(Math.round(r.pay || 0));
      ws.addRow(line);
    });

    ws.addRow([]);
    const add = (label, val, color) => {
      const cells = new Array(nCol).fill("");
      cells[nCol - 2] = label; cells[nCol - 1] = Math.round(val);
      const row = ws.addRow(cells);
      row.font = { bold: true, color: color ? { argb: color } : undefined };
    };
    add("공급가 합계", v.pay - (v.ship || 0));
    if (v.ship) add("배송비 합계", v.ship);
    if (v.ded) add("CS 차감", -v.ded, "FFCC0000");
    add("정산금액", v.final, "FF1A56DB");
    if (v.unpriced) {
      ws.addRow([]);
      const w = ws.addRow([`※ 단가가 아직 확정되지 않은 ${v.unpriced}건은 이번 정산금액에서 빠져 있습니다. 단가 확정 후 정산해 드리겠습니다.`]);
      w.font = { bold: true, color: { argb: "FF8A6D00" } };
    }
    if (v.dedRows && v.dedRows.length) {
      ws.addRow([]);
      ws.addRow(["[교환·반품 차감 내역]"]).font = { bold: true };
      ws.addRow(["접수일", "유형", "주문번호", "상품명", "내용", "비용"]).font = { bold: true };
      v.dedRows.forEach(x => ws.addRow([x.date, x.type, x.orderNo, x.product,
        String(x.content || "").slice(0, 120), Math.round(Number(x.cost) || 0)]));
    }
    src.forEach((h, i) => {
      const w = /상품명|주소|메시지|메세지/.test(h) ? 34 : /수취인|주문자|이름/.test(h) ? 12 : 14;
      ws.getColumn(i + 1).width = w;
    });
    for (let n = src.length + 1; n <= nCol; n++) { ws.getColumn(n).width = 14; ws.getColumn(n).numFmt = "#,##0"; }
    return ws;
  }

  /* 내부용 요약 장표 — 업체별 한 눈에 */
  function summarySheet(wb, vendors) {
    const ws = wb.addWorksheet("정산 요약");
    const co = CO();
    const p = (result && result.period) || { label: "" };
    ws.addRow([`${co} 정산 요약${p.label ? " — " + p.label : ""}`]);
    ws.getRow(1).font = { bold: true, size: 14 };
    ws.addRow([stampLine("부가세 포함")]);
    if (p.from) ws.addRow([`정산 대상 주문일 ${QO.fmtDate(p.from)} ~ ${QO.fmtDate(p.to)}`]);
    const w = ws.addRow(["※ 내부용입니다. 매출·마진이 들어 있으니 업체에 보내지 마세요."]);
    w.font = { bold: true, color: { argb: "FFCC0000" } };
    ws.addRow([]);

    const cols = vendors.map(v => v.vendor);
    ws.addRow([""].concat(cols).concat(["합계"]));
    const hr = ws.lastRow.number;
    ws.getRow(hr).font = { bold: true };
    ws.getRow(hr).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEFF3FB" } };

    const sum = f => vendors.reduce((s, v) => s + f(v), 0);
    const line = (label, f, opt) => {
      const row = ws.addRow([label].concat(vendors.map(v => Math.round(f(v)))).concat([Math.round(sum(f))]));
      if (opt && opt.bold) row.font = { bold: true, color: opt.color ? { argb: opt.color } : undefined };
      return row;
    };
    line("공급가", v => v.pay - (v.ship || 0));
    line("배송비", v => v.ship || 0);
    if (vendors.some(v => v.ded)) line("교환·반품 차감", v => -(v.ded || 0), { color: "FFCC0000" });
    line("업체 지급액", v => v.final, { bold: true, color: "FF1A56DB" });
    ws.addRow([]);
    line(`${co} 매출`, v => v.amount);
    line(`${co} 마진`, v => v.margin, { bold: true, color: "FF0A7A3D" });
    const rate = ws.addRow(["마진율"].concat(vendors.map(v =>
      v.amount ? Math.round(v.margin / v.amount * 1000) / 10 + "%" : "-"))
      .concat([sum(v => v.amount) ? Math.round(sum(v => v.margin) / sum(v => v.amount) * 1000) / 10 + "%" : "-"]));
    rate.font = { color: { argb: "FF666666" } };
    ws.addRow([]);
    line("건수", v => v.rows.length);
    line("수량", v => v.rows.reduce((s, r) => s + (r.qty || 0), 0));
    if (vendors.some(v => v.unpriced)) {
      const u = line("⚠ 단가 못 찾음(건)", v => v.unpriced || 0, { color: "FFCC0000" });
      u.font = { bold: true, color: { argb: "FFCC0000" } };
    }
    ws.addRow([]);
    const br = ws.addRow(["브랜드"].concat(vendors.map(v => (v.brands || []).join(", "))));
    br.font = { color: { argb: "FF666666" } };

    // 검산 결과 — 주문이 빠짐없이 담겼는지
    const c = result && result.check;
    if (c) {
      ws.addRow([]);
      const t = ws.addRow([c.ok && !c.issues.length ? "✔ 검산 통과" : "⚠ 검산 확인 필요"]);
      t.font = { bold: true, color: { argb: c.ok && !c.issues.length ? "FF0A7A3D" : "FFCC0000" } };
      ws.addRow([`주문 ${c.orderCount}건 · 수량 ${c.orderQty}개 → 업체 정산 ${c.settledCount}건 · 수량 ${c.settledQty}개`]);
      c.issues.forEach(i => {
        const row = ws.addRow([`· ${i.why}${i.detail ? " (" + i.detail + ")" : ""}`]);
        row.font = { color: { argb: i.level === "error" ? "FFCC0000" : "FF8A6D00" } };
      });
    }

    ws.getColumn(1).width = 18;
    for (let i = 2; i <= cols.length + 2; i++) { ws.getColumn(i).width = 16; ws.getColumn(i).numFmt = "#,##0"; }
    return ws;
  }

  async function buildExcel(vendors, opts) {
    const internal = !!(opts && opts.internal);
    const wb = new ExcelJS.Workbook();
    if (internal) {
      summarySheet(wb, vendors);
      const head = QO.settleSheetHead(true);
      const nCol = head.length;
      const safe = n => String(n).replace(/[\\\/\?\*\[\]:]/g, "_").slice(0, 26) || "정산";
      for (const v of vendors) {
        const ws = wb.addWorksheet(safe(v.vendor) + "_상세");
        const p = (result && result.period) || { label: "" };
        ws.addRow([`${v.vendor} 정산 상세 (내부용)${p.label ? " — " + p.label : ""}`]);
        ws.getRow(1).font = { bold: true, size: 14 };
        ws.addRow([stampLine("부가세 포함")]);
        const w = ws.addRow([`※ 내부용입니다. ${CO()} 매출·마진이 들어 있으니 업체에 그대로 보내지 마세요.`]);
        w.font = { bold: true, color: { argb: "FFCC0000" } };
        ws.addRow([]);
        ws.addRow(head);
        const hr = ws.lastRow.number;
        ws.getRow(hr).font = { bold: true };
        ws.getRow(hr).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEFF3FB" } };
        v.rows.forEach(r => ws.addRow(QO.settleSheetRow(r, true)));
        ws.addRow([]);
        const add = (label, val, color) => {
          const cells = new Array(nCol).fill("");
          cells[nCol - 2] = label; cells[nCol - 1] = Math.round(val);
          const row = ws.addRow(cells);
          row.font = { bold: true, color: color ? { argb: color } : undefined };
        };
        add(`${CO()} 매출`, v.amount);
        add(`${CO()} 마진`, v.margin);
        if (v.ded) add("교환·반품 차감", -v.ded, "FFCC0000");
        add("업체 지급액", v.final, "FF1A56DB");
        ws.columns.forEach((c, i) => { c.width = [12, 12, 20, 34, 18, 7, 13, 12, 13, 13][i] || 14; });
        for (let n = 7; n <= nCol; n++) ws.getColumn(n).numFmt = "#,##0";
      }
    } else {
      for (const v of vendors) vendorSheet(wb, v);
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
      const buf = await buildExcel([v]);          // 업체용 — 매출·마진 없음
      const p = (result && result.period) || { label: "" };
      const body = `안녕하세요, ${CONFIG.company}입니다.\n\n`
        + (p.label ? `${p.label} 정산 내역을 보내드립니다. (작성일 ${QO.fmtDate(QO.todayStr())})\n\n`
                   : `${QO.fmtDate(QO.todayStr())} 기준 정산 내역을 보내드립니다.\n\n`)
        + `· 건수: ${v.rows.length}건 (수량 ${v.rows.reduce((s2, r) => s2 + (r.qty || 0), 0)}개)\n`
        + `· 공급가 합계: ${won(v.pay)}\n`
        + (v.ded ? `· CS 차감(교환·반품): -${won(v.ded)}\n` : "")
        + `· 지급액: ${won(v.final)} (부가세 포함)\n`
        + (v.unpriced ? `\n※ 단가가 아직 확정되지 않은 ${v.unpriced}건은 이번 지급액에서 빠져 있습니다. 단가 확정 후 정산해 드리겠습니다.\n` : "")
        + `\n자세한 내역은 첨부 파일을 확인해주세요.\n감사합니다.`;
      await GMAIL.send({
        to: to.join(", "),
        subject: `[${CONFIG.company}] ${p.label ? p.label + " " : ""}정산서 - ${v.vendor} (작성 ${QO.fmtDate(QO.todayStr())})`,
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
    /* ---- 업체별 공급가표 ---- */
    const takePb = async f => {
      $("pb-fname").textContent = "📄 " + f.name;
      try { await addPriceBook(await readFile(f), f.name); }
      catch (e) { msg("msg-pb", "err", "⚠ " + f.name + " — " + e.message); }
    };
    $("f-pb").addEventListener("change", async function () {
      const fs = [...this.files]; this.value = "";
      for (const f of fs) await takePb(f);
    });
    bindDrop("drop-pb", async fs => { for (const f of fs) await takePb(f); });
    $("pb-drive").onclick = () => openDrivePicker({
      key: "pricebook", multiple: false, title: "드라이브에서 공급가표 가져오기",
      onPick: async fs => {
        for (const f of fs) {
          $("pb-fname").textContent = "📄 " + f.name;
          try { const r = await GMAIL.driveFetchExcel(f.id); await addPriceBook(r.buf, r.name || f.name); }
          catch (e) { msg("msg-pb", "err", "⚠ " + f.name + " — " + e.message); }
        }
      },
    });
    $("st-pb-tpl").onclick = async () => {
      try { download(await priceBookTemplate(), "공급가표_양식.xlsx"); }
      catch (e) { msg("msg-pb", "err", "⚠ " + e.message); }
    };

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
    $("run-s").onclick = async () => {
      const b = $("run-s");
      b.classList.add("loading"); b.disabled = true;
      try {
        calc();
        // 이번 미출고 목록을 남겨 다음 달에 이월 여부를 알려준다
        if (result) saveCarry(result.unshipped || []).catch(() => {});
        msg("msg-s", "ok", "✔ 뽑았습니다. 아래에서 업체별 정산내역을 확인하세요.");
      }
      catch (e) { msg("msg-s", "err", "⚠ " + e.message); }
      finally { b.classList.remove("loading"); b.disabled = false; }
    };
    $("st-mark").onclick = () => markSettled();
    $("st-export-all").onclick = async () => {
      if (!result) return;
      try {
        const buf = await buildExcel(result.vendors, { internal: true });   // 내부용 — 매출·마진 포함
        download(buf, `${periodTag()}_${CONFIG.company}_정산표_전체_내부용.xlsx`);
      } catch (e) { msg("msg-s", "err", "⚠ 정산표를 만들지 못했어요 — " + e.message); }
    };
  }

  async function init() { bind(); await load(); drawFilterLine(); }
  async function onShow() {
    if (!drawn) { await load(); drawn = true; }
    drawFiles(); drawPriceBook(); drawBrands(); refresh();
    if (result) drawResult();
  }

  init();
  return { onShow, drawFilter: drawFilterLine, markSettled, calc, result: () => result };
})();
window.ST = ST;
