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
    /* ★ 날짜 두 개를 따로 쓴다.
       · 정산일 = 주문수집일자 → 어느 달 정산에 넣을지를 정한다.
         주문일이 7/31 이어도 발주마감 뒤 주문이면 수집이 8/1 이고, 출고도 8월이라 8월 정산이다.
       · 주문일시 → 행사 공급가가 적용되는지(고객이 언제 주문했는지)를 따진다. */
    { k: "date", n: "정산일(수집일자)", kw: ["주문수집일자", "수집일자", "수집일", "정산일", "정산기준일", "구매확정일", "일자", "날짜"] },
    { k: "orderDate", n: "주문일시", kw: ["주문일시", "주문일자", "주문일", "결제일시", "결제일"] },
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
    // 고객이 실제로 결제한 금액. 이건 매출 계산에 쓰지 않고 화면에 '쇼핑몰 매출' 로만 보여준다.
    // 여기서 우리 매출을 빼면 몰이 가져가는 몫(쇼핑몰 수수료)이 나온다.
    { k: "mallAmount", n: "고객 결제금액", kw: ["결제금액", "판매가(쇼핑몰)", "총결제금액", "상품결제금액"] },
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
  let aliasInfo = {};    // 연결표에 쓴 원래 상품명 {주문키:{brand,product,option}} — 키만으론 사람이 못 읽는다
  let brandFix = {};     // 브랜드 → 업체 수동 배정 (공급가표 자동 배정보다 우선). 한 번 정하면 기억한다
  let extraVendors = []; // 공급가표에 없어도 직접 만든 업체
  let carry = { at: 0, list: [] };   // 지난 정산에서 송장이 없어 뺀 주문 (이월 추적용)
  let names = {};        // 업체 → 저장·발송할 파일명 (수정하면 그대로 씀)
  /* 파트너 MD 리워드 — 매출(또는 이익)의 일부를 MD에게 준다.
     [{ id, md, base:"매출"|"이익", rate: 3, picks: { 업체명: [브랜드…] } }]
     · picks 에 업체가 있고 배열이 비어 있으면 그 업체 '전체'가 대상이다.
     · MD 한 명이 여러 업체를 맡을 수 있어 업체를 여러 개 담는다.
     한 번 정해두면 다음 달 정산에서도 그대로 불러온다.
     ★ 리워드까지 빼야 진짜 우리 마진이다. 업체용 정산서에는 절대 나가지 않는다. */
  let rewards = [];

  const s = v => (v === null || v === undefined) ? "" : String(v).trim();
  const num = v => {
    const n = Number(String(v == null ? "" : v).replace(/[^0-9.\-]/g, ""));
    return isFinite(n) ? n : 0;
  };
  const won = n => (Math.round(n || 0)).toLocaleString("ko-KR") + "원";
  /* 마진율 = 마진 ÷ 매출. 내부용 정산표 엑셀과 같은 식이라 화면·파일 숫자가 어긋나지 않는다. */
  const rateOf = (margin, amount) => (amount ? Math.round((margin || 0) / amount * 1000) / 10 + "%" : "-");

  async function load() {
    maps = await DB.get("settleMaps", {}) || {};
    extraVendors = await DB.get("settleVendors", []) || [];
    pbRaw = await DB.get("priceBook", null);
    aliases = await DB.get("priceAliases", {}) || {};
    aliasInfo = await DB.get("priceAliasInfo", {}) || {};
    brandFix = await DB.get("settleBrandVendor", {}) || {};
    rewards = await DB.get("mdRewards", []) || [];      // 지난 정산에서 정해둔 MD 리워드 조건
    // v6.3.9 에 잠깐 있던 '리워드 없음' 은 없앴다. 요율 0 으로 바꿔 그대로 잠자게 둔다
    // (매출 대비로 되살리면 안 주던 리워드가 갑자기 붙는다)
    rewards.forEach(r => { if (r.base === "없음") { r.base = "매출"; r.rate = 0; } });
    // 업체 하나만 담던 옛 조건({vendor, brands})을 여러 업체({picks})로 옮긴다
    rewards.forEach(r => {
      if (!r.picks) {
        r.picks = {};
        if (r.vendor) r.picks[r.vendor] = (r.brands || []).slice();
        delete r.vendor; delete r.brands;
      }
      // 업체마다 기준·요율이 다를 수 있게 바꿨다. 배열이던 것을 객체로 옮긴다.
      Object.keys(r.picks).forEach(v => {
        const p = r.picks[v];
        if (Array.isArray(p)) r.picks[v] = { brands: p.slice(), base: r.base || "매출", rate: Number(r.rate) || 0 };
        else { p.brands = p.brands || []; p.base = p.base || "매출"; p.rate = Number(p.rate) || 0; }
      });
    });
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
  async function addPriceBook(buf, name, opts) {
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
      // 원본에서 숨겨둔 열은 기억해뒀다가 내려받을 때 그대로 숨긴다.
      // previewSheets 의 열 순서는 원본 1열부터 그대로라 번호가 어긋나지 않는다.
      const ws0 = wb.getWorksheet(sh.name);
      const hidden = [];
      for (let c = 1; c <= sh.columns.length; c++) {
        const col = ws0 && ws0.getColumn(c);
        hidden.push(!!(col && (col.hidden || col.width === 0)));
      }
      parsed.push({ name: sh.name, kind: isPrice ? "price" : "ship", cols: sh.columns, hidden,
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
      cols: (p.cols || []).slice(),          // 미리보기·내려받기로 되돌리려면 머리글도 있어야 한다
      hidden: (p.hidden || []).slice(),      // 원본에서 숨겨둔 열
      rows: p.rows.map(r => r.slice()),
    })), off: off.slice(),
      // 드라이브에서 가져온 파일이면 어디서 왔는지 기억한다 → 켤 때마다 최신본으로 맞춘다
      drive: (opts && opts.drive) || (pbRaw && pbRaw.name === name ? pbRaw.drive : null) || null };
    await savePriceBook();
    drawPriceBook(); drawBrands(); drawMd(); refresh();
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
    if (!files.length) return "";
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
        keys.length > 10 ? `<br>· 외 ${keys.length - 10}종` : ""}</div></div>`;
  }

  /* 공급가표를 드라이브 최신본으로 맞춘다.
     공급가표는 수시로 드라이브에서 고쳐지는데, 앱에는 올린 시점의 값이 박혀 있어
     모르는 사이 옛 단가로 정산할 수 있다. 그래서 켤 때마다 원본이 바뀌었는지 본다.
     · 드라이브에서 가져온 파일일 때만 동작한다 (PC 에서 끌어다 놓은 파일은 대상 아님)
     · 수정시각이 그대로면 내려받지 않는다 (매번 받아오면 느리다) */
  let pbChecked = false;
  async function refreshPriceBook(force) {
    const d = pbRaw && pbRaw.drive;
    if (!d || !d.id) return;
    if (pbChecked && !force) return;
    if (typeof GMAIL === "undefined" || GMAIL.needLogin()) return;
    pbChecked = true;
    try {
      const info = await GMAIL.driveFileInfo(d.id);
      const mt = (info && info.modifiedTime) || "";
      if (mt && d.mtime && mt === d.mtime) {                    // 안 바뀜 → 받지 않는다
        // 언제 확인했는지는 남긴다. 이게 없으면 '자동 확인이 도는지' 알 방법이 없다.
        pbRaw.checkedAt = Date.now();
        await savePriceBook(); drawPbLoaded();
        return;
      }
      const r = await GMAIL.driveFetchExcel(d.id);
      await addPriceBook(r.buf, r.name || d.name, { drive: { id: d.id, name: r.name || d.name, mtime: mt } });
      pbRaw.checkedAt = Date.now(); await savePriceBook(); drawPbLoaded();
      msg("msg-pb", "ok", `🔄 드라이브의 최신 공급가표로 맞췄어요 — ${r.name || d.name}`);
    } catch (e) {
      // 못 가져와도 지금 있는 공급가표로 그대로 쓴다. 다만 조용히 넘어가지 않는다.
      msg("msg-pb", "warn", "⚠ 드라이브 최신본을 확인하지 못했어요 — 지금 올려둔 공급가표로 계산합니다. (" + e.message + ")");
    }
  }

  /* 올려둔 공급가표를 엑셀로 되돌린다 — 미리보기·내려받기에 쓴다.
     원본 파일 자체를 들고 있지는 않아서(읽은 값만 저장) 서식은 빠지고 내용만 그대로다.
     정산에서 꺼둔 시트는 이름 뒤에 표시해, 어떤 시트가 계산에 안 쓰였는지 알 수 있게 한다. */
  async function priceBookExcel() {
    // 올린 내용 그대로를 되돌리는 것이므로, 단가가 하나도 안 잡힌 파일이어도 내보낸다
    if (!pbRaw || !(pbRaw.sheets || []).length) throw new Error("올려둔 공급가표가 없어요.");
    const off = pbRaw.off || [];
    const wb = new ExcelJS.Workbook();
    const used = {};
    (pbRaw.sheets || []).forEach(sh => {
      const skip = off.indexOf(sh.name) >= 0;
      let nm = String(sh.name || "시트").replace(/[\\\/\?\*\[\]:]/g, "_").slice(0, 24) + (skip ? " (미사용)" : "");
      while (used[nm]) nm = nm.slice(0, 27) + "_";      // 이름이 겹치면 시트가 사라진다
      used[nm] = 1;
      const ws = wb.addWorksheet(nm);
      const cols = (sh.cols || []).slice();
      if (cols.length) {
        ws.addRow(cols);
        ws.getRow(1).font = { bold: true };
        ws.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEFF3FB" } };
        ws.getRow(1).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      }
      (sh.rows || []).forEach(r => ws.addRow(r.slice()));
      const n = Math.max(cols.length, ...(sh.rows || []).map(r => r.length), 1);
      const hid = sh.hidden || [];
      for (let c = 1; c <= n; c++) {
        const h = String(cols[c - 1] || "");
        ws.getColumn(c).width = /상품명|품명|비고|방법/.test(h) ? 34 : /업체|브랜드/.test(h) ? 14 : 12;
        if (hid[c - 1]) ws.getColumn(c).hidden = true;      // 원본에서 숨겨둔 열은 그대로 숨긴다
      }
    });
    if (!wb.worksheets.length) throw new Error("되돌릴 시트가 없어요.");
    return await QO.saveWorkbook(wb);
  }

  /* 올려둔 공급가표가 있다는 걸 한눈에 — 초록 상자 + 작은 '해제' 버튼.
     공급가표는 한 번 올려두고 몇 달을 쓰는 파일이라, 지금 뭐가 걸려 있는지가 제일 중요하다. */
  /* 언제 읽어온 값인지 — 공급가표는 드라이브에서 수시로 바뀌어서,
     지금 화면의 단가가 언제 기준인지 안 보이면 옛 값으로 정산해도 모른다. */
  function fmtWhen(ms) {
    const d = new Date(Number(ms) || 0);
    if (!ms || isNaN(d.getTime())) return "";
    const p = n => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
  }
  function drawPbLoaded() {
    const row = $("pb-loaded"); if (!row) return;
    const on = hasBook();
    row.style.display = on ? "flex" : "none";
    if (!on) return;
    $("pb-fname").textContent = "📗 " + (pbRaw.name || "공급가표");
    // 드라이브를 확인한 시각이 있으면 그게 곧 '최신화된 시점'이다 (바뀐 게 없어도 최신 상태다).
    // 한 줄로 짧게만 적는다.
    const when = fmtWhen(pbRaw.checkedAt || pbRaw.at);
    $("pb-when").textContent = when ? `🔄 ${when} 최신화` : "";
  }
  function drawPriceBook() {
    drawPbLoaded();
    const box = $("pb-state");
    if (!box) return;
    // 파일명·상품수·시트 목록은 뺐다 (2026-08-04) — 위 초록 상자와 동어반복이다.
    // 시트 켜고 끄기도 같이 빠졌다. '최저가 기준 방식'처럼 업체명 열이 없는 참고 시트는
    // addPriceBook 이 알아서 제외하므로 평소엔 손댈 일이 없다.
    box.innerHTML = hasBook() ? matchBoxHtml() : "";
  }

  /* 빈 양식 내려받기 — 실제로 쓰는 '업체별 공급가 리스트_정산용.xlsx' 와 같은 모양으로 만든다.
     열 이름·위치가 그 파일과 같아야 받은 사람이 채워서 그대로 올릴 수 있다.
     ※ A열은 비우고 머리글은 2행 — 원본 파일이 그렇게 생겼다.
     ※ '2.최저가 기준 방식' 시트는 참고자료라 정산에 쓰지 않으므로 넣지 않는다.
     ※ 지급액 = (상품 공급가 + 배송비) × 수량. '브랜드→벤더 공급가' 는 그 둘을 더한 값이라
       사람이 눈으로 맞춰보는 용도다 — 프로그램은 앞의 두 열로 계산한다. */
  async function priceBookTemplate() {
    const wb = new ExcelJS.Workbook();
    const sheet = (name, head, rows, widths) => {
      const ws = wb.addWorksheet(name);
      ws.addRow([]);                                    // 1행 비움
      ws.addRow([""].concat(head));                     // 2행 머리글 (B열부터)
      const hr = ws.getRow(2);
      hr.font = { bold: true };
      hr.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      for (let c = 2; c <= head.length + 1; c++)
        hr.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEFF3FB" } };
      rows.forEach(r => ws.addRow([""].concat(r)));
      widths.forEach((w, i) => { ws.getColumn(i + 2).width = w; });
      ws.getRow(2).height = 30;
      return ws;
    };

    /* ① 평시 공급가 */
    const h1 = ["NO.", "업체명", "상품명", "모델명", "브랜드명", "사이트 검색어", "상품구분", "카테고리",
      "제조사", "원산지", "세금 구분", "배송비 구분", "배송비", "반품지", "상품 공급가", "옵션 제목",
      "옵션상세명칭", "대표이미지", "상세페이지", "판매가", "소비자가", "합배송 가능여부",
      "토요일 배송 가능여부", "재고관리여부", "재고관리수량", "영양성분 표시 대상 여부",
      "유전자 재조합 식품 여부", "브랜드→벤더 공급가 (배송비 포함)"];
    const r1 = [];
    const row1 = (no, vendor, product, brand, ship, price) => {
      const a = new Array(h1.length).fill("");
      a[0] = no; a[1] = vendor; a[2] = product; a[4] = brand; a[12] = ship; a[14] = price;
      a[27] = price + ship;      // 눈으로 맞춰보는 칸
      return a;
    };
    r1.push(row1(1, "플라스머", "[애플하우스] 일반떡볶기 3pk", "애플하우스", 4000, 15000));
    r1.push(row1(2, "플라스머", "[현우동] 카레우동 3pk", "현우동", 4000, 17870));
    r1.push(row1(3, "디에스피", "[수피아토] 쿨타월", "수피아토", 0, 7370));
    const w1 = [6, 12, 34, 12, 12, 14, 10, 10, 10, 10, 10, 11, 10, 10, 13, 12, 13, 12, 12, 11, 11, 13, 15, 12, 12, 16, 16, 20];
    const ws1 = sheet("1.운영 상품 리스트_정산용", h1, r1, w1);
    [14, 16, 29].forEach(c => { ws1.getColumn(c).numFmt = "#,##0"; });

    /* ② 행사 공급가 — 기간이 겹치면 이 값이 우선한다 */
    const h2 = ["NO.", "업체명", "행사 시작일", "행사 종료일", "상품명", "모델명", "브랜드명",
      "사이트 검색어", "상품구분", "카테고리", "제조사", "원산지", "세금 구분", "배송비 구분",
      "배송비", "반품지", "상품 공급가", "브랜드→벤더 행사공급가 (배송비 포함)"];
    const row2 = (no, vendor, from, to, product, brand, ship, price) => {
      const a = new Array(h2.length).fill("");
      a[0] = no; a[1] = vendor; a[2] = from; a[3] = to; a[4] = product; a[6] = brand;
      a[14] = ship; a[16] = price; a[17] = price + ship;
      return a;
    };
    const r2 = [row2(1, "플라스머", "2026-07-27", "2026-08-14", "[종로계림] 마늘삼계탕 2pk", "종로계림", 4000, 20210)];
    const w2 = [6, 12, 13, 13, 34, 12, 12, 14, 10, 10, 10, 10, 10, 11, 10, 10, 13, 22];
    const ws2 = sheet("2.행사 상품리스트_정산용", h2, r2, w2);
    [16, 18, 19].forEach(c => { ws2.getColumn(c).numFmt = "#,##0"; });
    [4, 5].forEach(c => { ws2.getColumn(c).numFmt = "yyyy-mm-dd"; });

    /* ③ 업체별 배송비 정산 방법 */
    sheet("3. 업체별 배송비 정산 방법", ["업체명", "배송비 정산 방법"], [
      ["디에스피", "주문당 배송비 1건으로 정산"],
      ["플라스머", "상품별 배송비 포함 공급가로 정산"],
    ], [16, 32]);

    /* ④ 최저가 기준 방식 — 참고자료.
       업체명 열이 없어서 정산에는 자동으로 쓰이지 않는다(addPriceBook 이 걸러낸다).
       원본 파일에 있는 시트라 모양을 맞춰 같이 넣는다. 머리글이 4행인 점이 다르다. */
    const ws4 = wb.addWorksheet("2.최저가 기준 방식");
    ws4.addRow([]);
    ws4.addRow(["", "카카오 선물하기 상품 LIST", "", "더벨로샵 수수료", "쇼핑몰 수수료", "배송비"]);
    ws4.addRow(["", "", "", 0.1, 0.15, 3500, "", "", "", "", "", "(단위 : 원)"]);
    const h4 = ["NO.", "상품명", "모델명", "브랜드→더벨로샵\n공급가\n(배송비 포함)", "더벨로샵\n수수료율",
      "더벨로샵→쇼핑몰\n공급가\n(배송비포함)", "쇼핑몰 \n수수료율", "쇼핑몰\n판매가",
      "최저가 대비\n할인율(%)", "온라인\n최저가\n(배송비포함)", "소비자가"];
    ws4.addRow([""].concat(h4));
    const hr4 = ws4.getRow(4);
    hr4.font = { bold: true };
    hr4.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    hr4.height = 46;
    for (let c = 2; c <= h4.length + 1; c++)
      hr4.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEFF3FB" } };
    // 예시 한 줄 — 수식은 원본과 같은 모양으로 넣는다 (열 위치가 바뀌면 여기도 같이 고쳐야 한다)
    const r5 = ws4.addRow(["", "예시", "미즈노 웨이브 라이더", "ascc1000", null, null, null, null, null, 0.1, 150000, 200000]);
    r5.getCell(5).value = { formula: "G5*(1-F5)" };
    r5.getCell(6).value = { formula: "$D$3" };
    r5.getCell(7).value = { formula: "I5*(1-H5)" };
    r5.getCell(8).value = { formula: "$E$3" };
    r5.getCell(9).value = { formula: "ROUND(K5*(1-J5),-2)" };
    [6, 8, 34, 14, 16, 12, 17, 12, 12, 13, 13, 12].forEach((w, i) => { ws4.getColumn(i + 1).width = w; });
    [5, 7, 9, 11, 12].forEach(c => { ws4.getColumn(c).numFmt = "#,##0"; });
    [6, 8, 10].forEach(c => { ws4.getColumn(c).numFmt = "0%"; });

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
      drawFiles(); drawPriceBook(); drawBrands(); drawMd(); refresh();
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
    // 한 줄에 하나씩, 폭을 다 써서 파일명이 잘리지 않게 (칩으로 나열하면 이름이 잘렸다)
    box.innerHTML = files.map((f, i) => `<div style="display:flex;align-items:center;gap:8px;
        padding:9px 10px;border:1px solid var(--line);border-radius:9px;background:var(--card2);margin-bottom:5px">
        <span style="flex:1;min-width:0;font-size:12.5px;line-height:1.45;word-break:break-all">📄 <b>${esc(f.name)}</b>
          <span style="color:var(--muted)"> · ${f.rows.length}행</span>
          <span style="display:flex;gap:6px;margin-top:6px">
            <button class="minibtn stpv" data-i="${i}">미리보기</button>
            <button class="minibtn stdl" data-i="${i}">엑셀 받기</button></span></span>
        <button class="minibtn vdel" data-i="${i}" style="flex:none">✕</button></div>`).join("")
      + `<div style="display:flex;justify-content:flex-end;margin-top:2px">
           <button class="minibtn" id="st-clear-all">전체 해제</button></div>`;
    /* 불러온 정산 파일도 원본을 들고 있지 않아(읽은 값만 저장) 읽어둔 표로 엑셀을 다시 만든다 */
    const fileExcel = async f => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet(String(f.name || "정산").replace(/[\\\/\?\*\[\]:]/g, "_").slice(0, 28) || "정산");
      ws.addRow((f.cols || []).slice());
      ws.getRow(1).font = { bold: true };
      ws.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEFF3FB" } };
      (f.rows || []).forEach(r => ws.addRow(r.slice()));
      (f.cols || []).forEach((h, i) => {
        ws.getColumn(i + 1).width = /상품명|주소|메시지|메세지/.test(String(h)) ? 34 : 13;
      });
      return await QO.saveWorkbook(wb);
    };
    box.querySelectorAll(".stpv").forEach(b => b.onclick = async () => {
      const f = files[Number(b.dataset.i)];
      try { openPreview(await fileExcel(f), f.name); }
      catch (e) { msg("msg-s", "err", "⚠ " + e.message); }
    });
    box.querySelectorAll(".stdl").forEach(b => b.onclick = async () => {
      const f = files[Number(b.dataset.i)];
      try { download(await fileExcel(f), String(f.name).replace(/\.xls[xm]$/i, "") + ".xlsx"); }
      catch (e) { msg("msg-s", "err", "⚠ " + e.message); }
    });
    box.querySelectorAll(".vdel").forEach(b => b.onclick = () => {
      files.splice(Number(b.dataset.i), 1);
      result = null; $("result-s").style.display = "none";
      drawFiles(); drawPriceBook(); drawBrands(); drawMd(); refresh();
    });
    const all = $("st-clear-all");
    if (all) all.onclick = () => {
      files = []; result = null;
      $("st-fname").textContent = "";
      const fi = $("f-st"); if (fi) fi.value = "";
      $("result-s").style.display = "none";
      msg("msg-s", "", "");
      drawFiles(); drawPriceBook(); drawBrands(); drawMd(); refresh();
    };
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
        // 고객 결제금액 (몰이 소비자에게 받은 값). 매출 계산엔 안 쓰고 화면 표시용이다.
        const mallAmount = f.map.mallAmount === undefined ? 0 : num(g(r, "mallAmount"));
        if (!product && !amount) { skipped.push({ src: f.name, raw: r }); continue; }
        // 브랜드는 파일에 열이 있으면 그걸 그대로 쓴다(상품명에서 추측하는 것보다 정확).
        const brand = (f.map.brand === undefined ? "" : s(g(r, "brand")))
          || CS.findBrand(product, option) || "";
        const vendor = (brand && (S.brandVendor || {})[brand])
          || CS.findVendor(product, option) || "";
        out.push({
          src: f.name,
          srcCols: f.cols, raw: r,     // 업체용 정산서를 원본 양식 그대로 뽑기 위해 들고 다닌다
          date: CS.toYmd(g(r, "date")),                    // 수집일자 — 정산 귀속
          orderDate: f.map.orderDate === undefined ? "" : CS.toYmd(g(r, "orderDate")),
          mall: s(g(r, "mall")) || guessMall(f.name),
          orderNo: s(g(r, "orderNo")),
          product, option, qty,
          unitPrice: unitPrice || null,
          amount, mallAmount,
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
      box.innerHTML = addBtn;
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
      drawBrands(); drawMd(); if (result) calc();
    });
    box.querySelectorAll(".vdel").forEach(el => el.onclick = async e => {
      e.stopPropagation();
      const v = el.dataset.v;
      if (!confirm(`업체 '${v}' 를 목록에서 지울까요?\n(배정해둔 브랜드는 자동 배정으로 돌아갑니다)`)) return;
      extraVendors = extraVendors.filter(x => x !== v);
      Object.keys(brandFix).forEach(b => { if (brandFix[b] === v) delete brandFix[b]; });
      await DB.set("settleVendors", extraVendors); await saveBrandFix();
      drawBrands(); drawMd(); if (result) calc();
    });
    bindAddVendor();

    const none = brands.filter(b => !ownerOf(b));
    $("st-brand-foot").textContent = none.length
      ? `아직 업체가 정해지지 않은 브랜드: ${none.join(", ")} — 해당 업체 칸에서 눌러 배정하세요.`
      : "";
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
      drawBrands(); drawMd();
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
        { vendor: v, rows: [], amount: 0, mallAmount: 0, pay: 0, margin: 0, ded: 0, dedRows: [],
          unpriced: 0, unpricedAmount: 0, loose: 0, ship: 0, brands: [] };
      g.rows.push(r); g.amount += r.amount; g.pay += r.pay; g.margin += r.margin;
      g.mallAmount += r.mallAmount || 0;
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
    calcMallFees(vendors);                            // 몰 수수료(삼성계열 카드 등)
    const mdList = calcRewards(vendors);              // 파트너 MD 리워드 (여기서 netMargin 까지 확정)
    const sum = f => vendors.reduce((s2, v) => s2 + (v[f] || 0), 0);
    result = {
      vendors,
      md: mdList,
      total: {
        amount: sum("amount"), mallAmount: sum("mallAmount"), pay: sum("pay"), margin: sum("margin"),
        ded: sum("ded"), final: sum("final"), count: rows.length,
        unpricedAmount: sum("unpricedAmount"),
        reward: sum("reward"), fee: sum("fee"), netMargin: sum("netMargin"),
      },
      // 몰 수수료 종류별 합계 (삼성계열 등)
      fees: (() => {
        const by = {};
        vendors.forEach(v => (v.feeRows || []).forEach(g => {
          const t = by[g.label] = by[g.label] || { label: g.label, rate: g.rate, amount: 0, fee: 0, count: 0 };
          t.amount += g.amount; t.fee += g.fee; t.count += g.count;
        }));
        return Object.values(by);
      })(),
      noVendor: rows.filter(r => !r.vendor && !r.brand).length,
      unpriced,
      loose: sum("loose"),
      usedBook: vendors.some(v => v.rows.some(r => r.mode === "book")),
    };
    result.check = reconcile(rows, shipped, unshipped, vendors, unpriced);
    result.unshipped = unshipped;
    result.period = periodOf(shipped.length ? shipped : rows);
    drawResult();
    drawMd();          // MD 카드에 계산된 리워드 금액을 채워 넣는다
  }

  /* 파트너 MD 리워드 — 조건 편집 화면.
     업체를 고르고, 그 업체의 브랜드까지 골라서 요율을 매긴다.
     브랜드를 하나도 안 고르면 그 업체 전체가 대상이다. */
  const brandsOfVendor = v => {
    const out = [];
    allRows().forEach(r => {
      if (ownerOf(r.brand || "(브랜드 없음)") !== v) return;
      const b = r.brand || "(브랜드 없음)";
      if (out.indexOf(b) < 0) out.push(b);
    });
    return out.sort();
  };
  async function saveRewards() { await DB.set("mdRewards", rewards); }
  /* MD 리워드 합계 — 이번 달에 우리가 MD 들에게 나갈 돈.
     (업체 지급액과는 별개다. 업체 지급액은 이 값에 영향받지 않는다) */
  function totalBoxHtml() {
    if (!result || !result.total || !result.total.reward) return "";
    const t = result.total;
    return `<div style="margin-top:12px;padding:12px 14px;border:1.5px solid var(--brand);border-radius:11px">
      ${(result.md || []).map(m => `<div class="totline" style="font-size:12.5px">
        <span>${esc(m.md)}</span><b>${won(m.reward)}</b></div>`).join("")}
      <div class="totline" style="border-top:1px solid var(--line);margin-top:6px;padding-top:8px">
        <b style="font-size:14px">MD 리워드 합계</b>
        <b style="font-size:19px;font-weight:800;color:var(--brand)">${won(t.reward)}</b></div>
      <div class="totline" style="font-size:12px;color:var(--muted);padding-top:4px">
        <span>${esc(CO())} 마진 ${won(t.margin)} − 리워드</span><b>최종 ${won(t.netMargin)}</b></div>
    </div>`;
  }
  function drawMd() {
    const card = $("st-card-md"), box = $("st-md-list");
    if (!card) return;
    const vendors = vendorNames();
    if (!allRows().length || !vendors.length) { card.style.display = "none"; box.innerHTML = ""; return; }
    card.style.display = "block";

    /* 계산이 끝나 있으면 그 조건이 실제로 얼마인지 여기서 바로 보여준다.
       금액을 보려고 아래 결과까지 내려갔다 올라올 일이 없게. */
    const lineOf = id => {
      for (const v of ((result && result.vendors) || []))
        for (const x of (v.rewardRows || [])) if (x.id === id) return x;
      return null;
    };
    box.innerHTML = rewards.map((r, i) => {
      const picks = r.picks || (r.picks = {});
      const hit = ruleTotals[r.id];
      /* 맡은 업체만 화면에 남긴다. 업체는 아래 드롭다운에서 골라 추가한다.
         업체가 늘어도 고른 것만 보이니 화면이 길어지지 않는다.
         '업체 전체' = 그 업체의 모든 브랜드 (나중에 브랜드가 늘어도 따라간다). */
      const mine = Object.keys(picks);
      const boxes = mine.map(v => {
        const p = picks[v] || {};
        const bs = brandsOfVendor(v);
        const all = !(p.brands || []).length;
        const chips = bs.map(b => {
          const sel = all || p.brands.indexOf(b) >= 0;
          return `<span class="brow${sel ? " on" : ""}" data-i="${i}" data-v="${esc(v)}" data-b="${esc(b)}">
            <span class="box">${sel ? "✓" : ""}</span>${esc(b)}</span>`;
        }).join("");
        // 업체마다 기준(매출/이익)과 요율을 따로 정한다
        const line = (v2 => { for (const x of ((result && result.vendors) || []))
          for (const y of (x.rewardRows || [])) if (y.id === r.id && y.vendor === v2) return y;
          return null; })(v);
        return `<div class="vendorbox" style="margin-top:6px;border-color:var(--brand)">
          <div class="vh" style="gap:6px">
            <span style="flex:1">🏭 ${esc(v)}</span>
            <span class="brow mdall${all ? " on" : ""}" data-i="${i}" data-v="${esc(v)}" style="flex:none">
              <span class="box">${all ? "✓" : ""}</span>업체 전체</span>
            <button class="minibtn mdvdel" data-i="${i}" data-v="${esc(v)}" style="flex:none">✕</button></div>
          <div style="display:flex;gap:6px;align-items:center;padding:6px 0 2px">
            <select class="mdvbase" data-i="${i}" data-v="${esc(v)}" style="flex:1;min-width:0;padding:6px;font-size:12px">
              <option value="매출"${rewardBase(p) === "매출" ? " selected" : ""}>매출 대비</option>
              <option value="이익"${rewardBase(p) === "이익" ? " selected" : ""}>이익 대비</option>
            </select>
            <input class="mdvrate" data-i="${i}" data-v="${esc(v)}" type="number" step="0.1" min="0"
              value="${Number(p.rate) || 0}" style="flex:none;width:64px;padding:6px;text-align:right;
              border:1px solid var(--line);border-radius:8px;background:var(--card);color:inherit;
              font-family:inherit;font-size:12px">
            <b style="flex:none;font-size:12px">%</b>
            <b style="flex:none;font-size:12.5px;color:${line && line.reward ? "var(--brand)" : "var(--faint)"}">${
              line ? won(line.reward) : "—"}</b></div>
          <div class="brands">${bs.length ? chips
            : `<span style="font-size:11.5px;color:var(--muted)">배정된 브랜드가 없습니다</span>`}</div></div>`;
      }).join("");
      const rest = vendors.filter(v => picks[v] === undefined);
      const adder = rest.length
        ? `<select class="mdadd" data-i="${i}" style="width:100%;margin-top:8px;padding:8px;font-size:12.5px">
             <option value="">＋ 담당업체 추가…</option>
             ${rest.map(v => `<option value="${esc(v)}">🏭 ${esc(v)}</option>`).join("")}
           </select>`
        : "";
      return `<div class="card" style="padding:12px;margin-bottom:12px;border-color:var(--line)">
        <div class="vh" style="gap:6px">
          <input class="mdname" data-i="${i}" value="${esc(r.md || "")}" placeholder="MD 이름"
            style="flex:1;min-width:0;padding:7px 9px;border:1px solid var(--line);border-radius:8px;
                   background:var(--card2);color:inherit;font-family:inherit;font-size:13px;font-weight:700">
          <button class="minibtn mddel" data-i="${i}" style="flex:none">✕</button></div>
        ${boxes}${adder}
        <div class="totline" style="margin-top:6px;padding-top:8px;border-top:1px solid var(--line)">
          <b style="font-size:13.5px">${esc(r.md || "MD")} 리워드</b>
          <b style="font-size:17px;font-weight:800;color:${hit && hit.reward ? "var(--brand)" : "var(--faint)"}">${
            hit ? won(hit.reward) : "— 정산내역 추출 후"}</b></div>
      </div>`;
    }).join("") + totalBoxHtml();

    const upd = async (i, fn) => { fn(rewards[i]); await saveRewards(); drawMd(); if (result) calc(); };
    // 드롭다운에서 고르면 그 업체가 '전체' 로 추가된다
    box.querySelectorAll(".mdadd").forEach(el => el.onchange = () => {
      if (!el.value) return;
      // 마지막에 쓴 기준·요율을 그대로 물려준다 (업체마다 다시 고르는 수고를 줄인다)
      upd(el.dataset.i, r => {
        const last = Object.keys(r.picks).map(k => r.picks[k]).pop() || {};
        r.picks[el.value] = { brands: [], base: last.base || "매출", rate: Number(last.rate) || 0 };
      });
    });
    box.querySelectorAll(".mdvdel").forEach(el => el.onclick = () =>
      upd(el.dataset.i, r => { delete r.picks[el.dataset.v]; }));
    box.querySelectorAll(".mdname").forEach(el => el.onchange = el.onblur = async () => {
      if (rewards[el.dataset.i].md === el.value.trim()) return;
      await upd(el.dataset.i, r => { r.md = el.value.trim(); });
    });
    box.querySelectorAll(".mdvbase").forEach(el => el.onchange = () =>
      upd(el.dataset.i, r => { r.picks[el.dataset.v].base = el.value; }));
    box.querySelectorAll(".mdvrate").forEach(el => el.onchange = () =>
      upd(el.dataset.i, r => { r.picks[el.dataset.v].rate = Number(el.value) || 0; }));
    box.querySelectorAll(".mdall").forEach(el => el.onclick = () =>
      upd(el.dataset.i, r => {
        const v = el.dataset.v, p = r.picks[el.dataset.v];
        if (p && !(p.brands || []).length) delete r.picks[v];   // 전체 → 해제
        else if (p) p.brands = [];                              // 업체 전체로
      }));
    // 브랜드 칩 — '업체 전체'(.mdall) 도 같은 class 라 빼고 잡는다
    box.querySelectorAll(".brow:not(.mdall)").forEach(el => el.onclick = () =>
      upd(el.dataset.i, r => {
        const v = el.dataset.v, b = el.dataset.b, p = r.picks[v];
        if (!p) return;
        const bs = p.brands || (p.brands = []);
        if (!bs.length) {                                       // '업체 전체' 에서 하나를 빼면
          p.brands = brandsOfVendor(v).filter(x => x !== b);     // 나머지 브랜드만 남긴다
          if (!p.brands.length) delete r.picks[v];
          return;
        }
        const j = bs.indexOf(b);
        if (j >= 0) { bs.splice(j, 1); if (!bs.length) delete r.picks[v]; }
        else {
          bs.push(b);
          if (bs.length === brandsOfVendor(v).length) p.brands = [];   // 다 고르면 '업체 전체'
        }
      }));
    box.querySelectorAll(".mddel").forEach(el => el.onclick = async () => {
      if (!confirm(`'${rewards[el.dataset.i].md || "이름 없음"}' 리워드 조건을 지울까요?`)) return;
      rewards.splice(el.dataset.i, 1);
      await saveRewards(); drawMd(); if (result) calc();
    });
  }

  /* =================================================================
     파트너 MD 리워드
     ─────────────────────────────────────────────────────────────────
     조건 하나 = { md, vendor, brands[], base, rate }.
     brands 가 비어 있으면 그 업체 전체가 대상이다.
     · 매출 대비 : 대상 브랜드의 매출 × 요율
     · 이익 대비 : (매출 − 업체 지급액) × 요율
       ★ '이익'은 리워드를 빼기 전 마진이다. 리워드를 뺀 값에 다시 요율을 곱하면
         자기 자신을 참조해 값이 정해지지 않는다.
     같은 브랜드에 조건을 여러 개 걸면 각각 계산해서 더한다 (MD 두 명이 나눠 갖는 경우).
     ================================================================= */
  /* 쇼핑몰별로 따로 떼가는 수수료 — 우리 매출에서 빠지므로 마진에서 공제한다.
     삼성블루베리몰·삼성카드쇼핑처럼 삼성계열은 카드 수수료 1.7% 가 별도로 나간다.
     요율이 바뀌거나 다른 몰이 생기면 여기만 고치면 된다. */
  /* ★ '블루베리' 도 삼성 계열이다 (삼성블루베리몰·블루베리몰).
       쇼핑몰명에 '삼성' 이 안 들어가 있어도 1.7% 대상이므로 반드시 같이 잡아야 한다.
       — 2026-08-05 사용자 확인. 여기서 빼면 그 건들이 조용히 공제에서 빠진다. */
  const MALL_FEES = [{ match: /삼성|블루베리/, rate: 1.7, label: "삼성계열 카드수수료" }];
  const mallFeeOf = mall => {
    const m = String(mall == null ? "" : mall);
    for (const f of MALL_FEES) if (f.match.test(m)) return f;
    return null;
  };
  /* 업체별 몰 수수료 — 대상 줄의 '우리 매출' 에 요율을 곱한다 */
  function calcMallFees(vendors) {
    vendors.forEach(v => {
      const by = {};
      v.rows.forEach(r => {
        const f = mallFeeOf(r.mall); if (!f) return;
        const g = by[f.label] = by[f.label] || { label: f.label, rate: f.rate, amount: 0, count: 0 };
        g.amount += r.amount || 0; g.count++;
      });
      v.feeRows = Object.values(by).map(g =>
        Object.assign(g, { fee: Math.round(g.amount * g.rate / 100) }));
      v.fee = v.feeRows.reduce((s2, g) => s2 + g.fee, 0);
    });
  }

  const rewardBase = r => (r.base === "이익" ? "이익" : "매출");
  let ruleTotals = {};        // 조건 id → 여러 업체를 합친 리워드 (카드에 바로 보여주려고)
  function calcRewards(vendors) {
    vendors.forEach(v => { v.reward = 0; v.rewardRows = []; });
    const byMd = {};
    ruleTotals = {};
    (rewards || []).forEach(rule => {
      const md = s(rule.md); if (!md) return;
      const picks = rule.picks || {};
      const tot = ruleTotals[rule.id] = { reward: 0, baseAmount: 0, count: 0, vendors: 0 };
      Object.keys(picks).forEach(vn => {
        const p = picks[vn] || {};
        const rate = Number(p.rate) || 0; if (!rate) return;      // 업체마다 요율이 따로다
        const v = vendors.find(x => x.vendor === vn);
        if (!v) return;                                // 이번 달에 그 업체 주문이 없으면 건너뜀
        const bl = p.brands || [];
        const pick = bl.length
          ? v.rows.filter(r => bl.indexOf(r.brand || "(브랜드 없음)") >= 0)
          : v.rows;                                    // 빈 배열 = 그 업체 전체
        if (!pick.length) return;
        const amount = pick.reduce((s2, r) => s2 + (r.amount || 0), 0);
        const pay = pick.reduce((s2, r) => s2 + (r.pay || 0), 0);
        const base = rewardBase(p) === "이익" ? amount - pay : amount;
        const won2 = Math.round(base * rate / 100);
        const line = { id: rule.id, md, vendor: v.vendor, brands: bl.slice(),
                       base: rewardBase(p), rate, baseAmount: base, reward: won2, count: pick.length };
        v.reward += won2; v.rewardRows.push(line);
        tot.reward += won2; tot.baseAmount += base; tot.count += pick.length; tot.vendors++;
        const g = byMd[md] = byMd[md] || { md, reward: 0, lines: [] };
        g.reward += won2; g.lines.push(line);
      });
    });
    // 리워드와 몰 수수료까지 빼야 진짜 우리 마진
    vendors.forEach(v => { v.netMargin = (v.margin || 0) - (v.reward || 0) - (v.fee || 0); });
    return Object.values(byMd).sort((a, b) => b.reward - a.reward);
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

  /* 돈 흐름 표시 — 전체 합계와 업체 카드가 같은 순서를 쓴다.
       쇼핑몰 매출 −(쇼핑몰 수수료) → 랩노마드 매출 −(업체 지급액) → 랩노마드 마진
       −(카드 수수료) −(MD 리워드) → 랩노마드 최종 마진
     쇼핑몰 수수료는 따로 적힌 값이 아니라 '고객 결제금액 − 우리 매출' 이다(몰이 가져가는 몫).
     결제금액 열이 없는 파일이면 위 두 줄은 그냥 나오지 않는다. */
  function moneyLines(d, o) {
    const big = o && o.big;
    const L = (label, val, opt) => {
      const s2 = opt || {};
      return `<div class="totline"${s2.small ? ' style="font-size:12px;color:var(--muted)"' : ""}>` +
        `${s2.small ? "<span>" : "<b>"}${label}${s2.small ? "</span>" : "</b>"}` +
        `<span${s2.color ? ` style="color:${s2.color}"` : ""}>${s2.minus ? "− " : ""}${won(val)}${
          s2.rate !== undefined ? ` <span style="color:var(--muted);font-weight:600">(${s2.rate})</span>` : ""}</span></div>`;
    };
    let h = "";
    const mallFee = (d.mallAmount || 0) - (d.amount || 0);
    if (d.mallAmount) {
      h += L("쇼핑몰 매출", d.mallAmount);
      // 요율은 언제나 오른쪽 금액 뒤 괄호에 (왼쪽은 항목 이름만)
      if (mallFee > 0) h += L("└ 쇼핑몰 수수료", mallFee,
        { small: true, minus: true, rate: rateOf(mallFee, d.mallAmount) });
    }
    h += L(`${esc(CO())} 매출`, d.amount);
    if (d.unpricedAmount)
      h += L("└ 단가 못 찾은 건 (지급·마진에서 빠짐)", d.unpricedAmount, { small: true, color: "var(--danger)" });
    if (d.ded) h += L("CS 차감 (교환·반품)", d.ded, { minus: true, color: "var(--danger)" });
    h += `<div class="totline"${big ? ' style="border-top:1px solid var(--line);margin-top:6px;padding-top:8px"' : ""}>` +
      `<b>${esc(o && o.payLabel || "업체 지급액")}</b><span style="color:var(--brand)">${won(d.pay)}</span></div>`;
    h += L(`${esc(CO())} 마진`, d.margin, { color: "var(--ok)", rate: rateOf(d.margin, d.amount) });
    // 대상 매출을 같이 적는다 — 이게 없으면 금액이 맞는지 눈으로 검산할 수 없다
    (d.feeRows || []).forEach(g =>
      h += L(`└ ${esc(g.label)} · ${g.count}건 · 대상 ${won(g.amount)}`, g.fee,
        { small: true, minus: true, rate: g.rate + "%" }));
    (d.rewardRows || []).forEach(m =>
      h += L(`└ MD 리워드 · ${esc(m.md)}`, m.reward,
        { small: true, minus: true, rate: m.rate ? m.rate + "%" : undefined }));
    if ((d.feeRows || []).length || (d.rewardRows || []).length)
      h += L(`${esc(CO())} 최종 마진`, d.netMargin, { color: "var(--ok)", rate: rateOf(d.netMargin, d.amount) });
    return h;
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
      moneyLines({
        mallAmount: t.mallAmount, amount: t.amount, pay: t.final, margin: t.margin,
        unpricedAmount: t.unpricedAmount, ded: t.ded, netMargin: t.netMargin,
        feeRows: result.fees || [],
        // 요율이 조건마다 다르면 합계 줄엔 안 적는다 (하나로 못 줄임)
        rewardRows: (result.md || []).map(m => {
          const rs = [...new Set((m.lines || []).map(x => x.rate))];
          return { md: m.md, reward: m.reward, rate: rs.length === 1 ? rs[0] : 0 };
        }),
      }, { big: true, payLabel: "업체 지급 합계" }) +
      `<div class="synchint" style="margin-top:6px">${t.count}건${
         result.noVendor ? ` · <b style="color:var(--danger)">업체 미지정 ${result.noVendor}건</b>` : ""}</div>` +
      checkBoxHtml(result.check) +
      (up.length ? unpricedBoxHtml(upKinds, kindKeys, up.length) : "");
    if (up.length) bindAliasPickers();

    // 업체별 요약 줄은 뺐다 (2026-08-04) — 바로 아래 업체 카드에 같은 숫자가 또 나온다
    $("st-sumbox").innerHTML = "";

    const box = $("st-vendors");
    box.innerHTML = "";
    result.vendors.forEach(v => {
      const row = document.createElement("div");
      row.className = "rrow";
      row.innerHTML = `<div class="rtop"><b style="flex:1">🏭 ${esc(v.vendor)}</b><span class="cnt">${v.rows.length}건</span></div>
        ${moneyLines({ mallAmount: v.mallAmount, amount: v.amount, pay: v.final, margin: v.margin,
                       unpricedAmount: v.unpricedAmount, ded: v.ded, netMargin: v.netMargin,
                       feeRows: v.feeRows, rewardRows: v.rewardRows }, { payLabel: "업체 지급액" })}
        ${v.unpriced ? `<div class="totline" style="font-size:12px;color:var(--danger)"><span>⚠ 단가 못 찾음</span><span>${v.unpriced}건 (0원 처리)</span></div>` : ""}
        <div class="rmail"><input type="email" placeholder="업체 이메일 (쉼표로 여러 명)" autocapitalize="off" spellcheck="false">
          <button class="dlbtn">메일 보내기</button></div>
        <div class="cands"></div>
        <div class="setrow" style="margin-top:6px"><span style="flex:1;font-size:11px;color:var(--faint)"></span>
          <button class="minibtn share">📤 카톡·공유</button><button class="minibtn pvbtn">미리보기</button><button class="minibtn dl">엑셀 받기</button></div>
        <div class="setrow" style="margin-top:4px"><span class="fnlbl" style="flex:1;font-size:11px;color:var(--faint)"></span>
          <button class="minibtn tpl">${mailtplLabel("settle", v.vendor)}</button><button class="minibtn fn">✏️ 파일명 수정</button></div>`;
      box.appendChild(row);
      bindMailtplBtn(row.querySelector(".tpl"), "settle", v.vendor, mailVars(v));
      const inp = row.querySelector("input");
      const ek = emailKey(v.vendor);            // ① 발주 탭에 저장된 업체 이메일을 같이 쓴다
      const known = (S.vendorEmails || {})[ek] || "";
      inp.value = known;
      fillRecipients(row.querySelector(".cands"), inp, {
        saved: known, history: (S.vendorSent || {})[ek] || [],
        domains: (S.vendorDomains || {})[ek] || [], query: v.vendor !== "(업체 미지정)" ? v.vendor : "",
      });
      // 여기서 고친 이메일은 ① 발주 탭과 같은 곳에 저장한다 (한 번만 적으면 양쪽에서 쓴다)
      inp.onchange = inp.onblur = async () => {
        const val = inp.value.trim();
        if (val === ((S.vendorEmails || {})[ek] || "")) return;
        S.vendorEmails = S.vendorEmails || {};
        S.vendorEmails[ek] = val;
        await DB.set("vendorEmails", S.vendorEmails);
      };
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
      row.querySelector(".dlbtn").onclick = () => sendOne(v, inp, row.querySelector(".dlbtn"));
    });
  }
  /* 정산 업체명 → ① 발주 탭에 저장된 업체 키.
     발주 들어가는 업체가 곧 정산하는 업체라, 이메일·발송이력·도메인을 같이 쓴다.
     이름이 정확히 같으면 그대로, 띄어쓰기·괄호만 다르면 그것도 찾아준다. */
  function emailKey(vendor) {
    const em = S.vendorEmails || {};
    if (em[vendor]) return vendor;
    const n = QO.normPriceText(vendor);
    if (!n) return vendor;
    // 후보: 이메일이 저장된 이름 + 발송 이력 + ① 발주 탭의 업체 양식 이름
    const cand = [];
    const push = k => { const t = String(k || "").trim(); if (t && cand.indexOf(t) < 0) cand.push(t); };
    Object.keys(em).forEach(push);
    Object.keys(S.vendorSent || {}).forEach(push);
    (S.forms || []).forEach(f => push(f.name));
    // ① 이름이 같다(띄어쓰기·괄호만 다른 경우 포함)
    for (const k of cand) if (QO.normPriceText(k) === n) return k;
    // ② 한쪽이 다른 쪽에 들어 있다 — '플라스머' ↔ '플라스머 발주양식' 같은 경우.
    //   여러 개가 걸리면 어느 것인지 알 수 없으니 고르지 않는다.
    const hit = cand.filter(k => {
      const kn = QO.normPriceText(k);
      return kn && (kn.indexOf(n) >= 0 || n.indexOf(kn) >= 0);
    });
    if (hit.length === 1) return hit[0];
    if (hit.length > 1) {
      // 이메일이 실제로 저장돼 있는 것을 우선한다
      const withMail = hit.filter(k => em[k]);
      if (withMail.length === 1) return withMail[0];
    }
    return vendor;
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
  function pickOptions(row, cur) {
    const opt = c => `<option value="${esc(c.key)}"${c.key === cur ? " selected" : ""}>${esc((c.brand ? c.brand + " · " : "") + c.product)}` +
      `${c.option ? " [" + esc(c.option) + "]" : ""} — ${Math.round(c.price).toLocaleString("ko-KR")}원` +
      `${c.n > 1 ? " (기간별 " + c.n + "개)" : ""}</option>`;
    const all = bookChoices();
    const byKey = {}; all.forEach(c => { byKey[c.key] = c; });
    const ranked = QO.rankPriceCandidates(pbook, row);
    const top = ranked.map(x => byKey[x.item.key]).filter(Boolean);
    const topKeys = {}; top.forEach(c => { topKeys[c.key] = 1; });
    const rest = all.filter(c => !topKeys[c.key]);
    let h = `<option value=""${cur ? "" : " selected"}>— 공급가표에서 고르기 —</option>`;
    if (top.length) h += `<optgroup label="이름이 비슷한 상품 ${top.length}개">` + top.map(opt).join("") + "</optgroup>";
    if (rest.length) h += `<optgroup label="${top.length ? "그 외 " : "전체 "}${rest.length}개">` + rest.map(opt).join("") + "</optgroup>";
    return h;
  }
  function unpricedBoxHtml(kinds, keys, total) {
    return `<div style="margin-top:10px;padding:10px;border:1px solid var(--danger);border-radius:8px">
      <b style="color:var(--danger)">⚠ 공급가표에서 단가를 못 찾은 ${total}건 — 지급액 0 원으로 뒀습니다</b>
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
        if (sel.value) {
          aliases[okey] = sel.value;
          // 나중에 '상품명 매칭관리'에서 사람이 읽을 수 있게 원래 이름을 같이 남긴다
          aliasInfo[okey] = { brand: r.brand || "", product: r.product || "", option: r.option || "" };
        } else { delete aliases[okey]; delete aliasInfo[okey]; }
        await DB.set("priceAliases", aliases);
        await DB.set("priceAliasInfo", aliasInfo);
        calc();
        msg("msg-s", "ok", "✔ 연결했어요. 다시 계산했습니다.");
      };
    });
  }

  /* =================================================================
     상품명 매칭관리 — 자동으로 못 붙어서 사람이 골라준 매칭을 다시 보고 고친다.
     여기가 틀리면 그 상품 지급액이 통째로 틀리는데, 한 번 저장하면 계속 쓰여서
     들여다볼 방법이 없으면 잘못된 연결이 몇 달을 간다.
     ================================================================= */
  /* 주문 쪽 이름 — 저장해둔 게 있으면 그걸, 없으면 이번 달 주문에서 찾아본다 */
  function aliasOrderLabel(okey) {
    const inf = aliasInfo[okey];
    if (inf && inf.product) return inf;
    const rows = (result && result.rows) || [];
    const hit = rows.find(r => QO.priceRowKey(r.brand, r.product, r.option) === okey);
    if (hit) return { brand: hit.brand || "", product: hit.product || "", option: hit.option || "" };
    return null;
  }
  function aliasBoxHtml() {
    const keys = Object.keys(aliases).filter(k => aliases[k]);
    if (!keys.length) return `<div class="empty">손으로 이어준 상품이 없습니다.</div>`;
    const byKey = {}; bookChoices().forEach(c => { byKey[c.key] = c; });
    return keys.map((k, i) => {
      const o = aliasOrderLabel(k);
      const t = byKey[aliases[k]];
      const left = o
        ? `<b>${esc(o.product.slice(0, 70))}</b>${o.option ? `<span style="color:var(--muted)"> [${esc(o.option)}]</span>` : ""}
           <div style="font-size:11.5px;color:var(--muted)">${esc(o.brand || "")}</div>`
        : `<b style="color:var(--muted)">(이름을 못 읽는 옛 연결)</b>
           <div style="font-size:11px;color:var(--faint);word-break:break-all">${esc(k.slice(0, 90))}</div>`;
      const right = t
        ? `<b style="color:var(--ok)">${esc((t.brand ? t.brand + " · " : "") + t.product)}</b>${t.option ? ` [${esc(t.option)}]` : ""}
           <span style="color:var(--muted)"> — ${Math.round(t.price).toLocaleString("ko-KR")}원</span>`
        : `<b style="color:var(--danger)">⚠ 지금 공급가표에 없는 항목입니다 — 다시 골라주세요</b>`;
      return `<div style="padding:10px 0;border-top:1px solid var(--line)">
        <div style="font-size:12.5px">${left}</div>
        <div style="font-size:12.5px;margin-top:4px">↳ ${right}</div>
        <div style="display:flex;gap:6px;margin-top:6px">
          <select class="aliasedit" data-k="${esc(k)}" style="flex:1;padding:7px;font-size:12.5px">${
            pickOptions(o || { brand: "", product: "", option: "" }, aliases[k])}</select>
          <button class="minibtn aliasdel" data-k="${esc(k)}" style="flex:none">연결 끊기</button>
        </div></div>`;
    }).join("");
  }
  async function saveAliases(why) {
    await DB.set("priceAliases", aliases);
    await DB.set("priceAliasInfo", aliasInfo);
    drawAliasBox();
    if (result) { calc(); }
    $("alias-msg").textContent = why;
  }
  function drawAliasBox() {
    const n = Object.keys(aliases).filter(k => aliases[k]).length;
    $("alias-sub").textContent = n ? `${n}개` : "";
    $("alias-list").innerHTML = aliasBoxHtml();
    $("alias-list").querySelectorAll(".aliasedit").forEach(sel => {
      sel.onchange = async () => {
        const k = sel.dataset.k;
        if (!sel.value) { delete aliases[k]; delete aliasInfo[k]; return saveAliases("연결을 끊었습니다."); }
        aliases[k] = sel.value;
        await saveAliases("✔ 연결을 바꿨습니다. 정산을 다시 계산했어요.");
      };
    });
    $("alias-list").querySelectorAll(".aliasdel").forEach(b => {
      b.onclick = async () => {
        const k = b.dataset.k;
        delete aliases[k]; delete aliasInfo[k];
        await saveAliases("연결을 끊었습니다. 이 상품은 다시 '단가 못 찾음'으로 나옵니다.");
      };
    });
  }

  /* '정산 내역' 화면 표는 뺐다 (2026-08-03) — 미리보기와 겹쳐서.
     내부 숫자(매출·마진)는 업체 카드 위쪽 요약과 '전체 정산표 (내부용)' 에 있다. */

  /* =================================================================
     엑셀 / 메일
     ================================================================= */
  /* 정산서 엑셀.
     opts.internal 이 참일 때만 '매출·우리마진' 을 넣는다.
     ※ 업체로 나가는 파일에는 우리 매출과 마진이 절대 들어가면 안 된다.
       업체가 볼 것은 '무엇을 몇 개, 단가 얼마에, 얼마 받는지'까지다.
       업체별 저장/메일은 전부 업체용(internal 없음)으로만 부른다. */
  const CO = () => (typeof CONFIG !== "undefined" && CONFIG.company) || "우리";

  /* 정산서 정렬 — 내용은 가운데, 금액 칸만 오른쪽.
     머리말(제목·작성일 줄)은 가운데로 몰면 어색해서 왼쪽에 그대로 둔다.
     moneyFrom = 이 열 번호부터 금액 (0 이면 금액 열 없음) */
  function alignSheet(ws, nCol, srcCols, headerRow) {
    const center = { horizontal: "center", vertical: "middle", wrapText: false };
    const right = { horizontal: "right", vertical: "middle" };
    for (let n = 1; n <= nCol; n++) ws.getColumn(n).alignment = n > srcCols ? right : center;
    // 표 머리(헤더)는 항상 가운데
    if (headerRow) for (let n = 1; n <= nCol; n++) ws.getRow(headerRow).getCell(n).alignment = center;
    // 제목·작성일 등 머리말 줄은 왼쪽
    for (let r = 1; r < (headerRow || 1); r++) ws.getRow(r).getCell(1).alignment = { horizontal: "left" };
  }

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
  /* 업체용 정산서에 적을 정산기간.
     한 달 안에 들어오면 그 달 1일~말일로 적는다 — 주문이 7/3~7/29 에만 있어도
     업체가 보는 정산 대상 기간은 '7월 한 달'이라서다. 달을 걸치면 실제 범위 그대로. */
  function periodRange() {
    const p = (result && result.period) || {};
    if (!p.from || !p.to) return "";
    const f = d => d.slice(0, 4) + "-" + d.slice(4, 6) + "-" + d.slice(6, 8);
    if (p.from.slice(0, 6) !== p.to.slice(0, 6)) return `${f(p.from)} ~ ${f(p.to)}`;
    const y = Number(p.from.slice(0, 4)), m = Number(p.from.slice(4, 6));
    const end = new Date(y, m, 0).getDate();          // m 월의 0일 = 그 달 말일
    const ym = p.from.slice(0, 4) + "-" + p.from.slice(4, 6);
    return `${ym}-01 ~ ${ym}-${String(end).padStart(2, "0")}`;
  }
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
    // 배송비를 개당 단가에 합칠 수 있는 업체(개당 정산)는 '공급가(배송비 포함)' 한 열로 끝낸다.
    // 주문당 1회로 붙는 업체만 배송비를 따로 적는다.
    const perOrderShip = v.rows.reduce((s2, r) => s2 + (r.shipMode === "건당" ? (r.shipTotal || 0) : 0), 0);
    // 열 이름에 실제 업체명을 넣는다 ('업체→랩노마드' 가 아니라 '플라스머→랩노마드').
    // 개당 업체는 배송비가 단가에 이미 합쳐져 있지만 '(배송비 포함)' 이라고 적지 않는다 — 업체 요청.
    // 주문한 달과 수집(출고)된 달이 다른 건이 있으면 '비고' 열을 붙인다
    const anyCarry = v.rows.some(r => QO.carryNote(r));
    const head = src.concat([`${v.vendor}→${CO()} 공급가`])
      .concat(perOrderShip ? ["배송비"] : []).concat(["정산금액"])
      .concat(anyCarry ? ["비고"] : []);
    const nCol = head.length;

    const p = (result && result.period) || { label: "" };
    const totalQty = v.rows.reduce((s2, r) => s2 + (r.qty || 0), 0);
    /* 파일을 열자마자 보이도록 최종 내역을 맨 위 D·E 열에 올린다.
       (아래 표 끝에도 같은 합계를 남겨 둔다 — 데이터 끝에서 바로 확인할 수 있게) */
    const tops = [["판매수량", totalQty, "0"]];
    tops.push(["공급가 합계", v.pay - perOrderShip, "#,##0"]);
    if (perOrderShip) tops.push(["배송비 합계", perOrderShip, "#,##0"]);
    if (v.ded) tops.push(["교환·반품 차감", -v.ded, "#,##0"]);
    tops.push(["정산금액", v.final, "#,##0"]);

    ws.addRow([`${v.vendor} 정산서${p.label ? " — " + p.label : ""}`]);
    ws.getRow(1).font = { bold: true, size: 14 };
    // 업체용에는 작성일을 넣지 않는다 — 업체가 볼 것은 '어느 기간을 정산했는지'다
    const pr = periodRange();
    ws.addRow([(pr ? `정산기간 ${pr} (출고완료된 주문건 기준) · ` : "") + `${v.rows.length}건`]);
    ws.addRow(["※ 아래 금액은 모두 부가세가 포함된 금액입니다."]);
    while (ws.rowCount < tops.length) ws.addRow([]);   // 요약이 길면 줄을 더 만든다
    ws.addRow([]);
    // D=4, E=5 에 항목/금액. 마지막 줄(정산금액)은 굵게·파랗게
    const topCells = {};                       // 나중에 실제 합계 수식으로 바꿔 넣는다
    tops.forEach((t, i) => {
      const row = ws.getRow(i + 1);
      const label = row.getCell(4), val = row.getCell(5);
      label.value = t[0]; val.value = Math.round(t[1]);
      topCells[t[0]] = val;
      const last = i === tops.length - 1;
      const font = { bold: true, size: last ? 13 : 11 };
      if (last) font.color = { argb: "FF1A56DB" };
      label.font = font; val.font = font;
      // 셀 서식은 새 객체로 만들어 넣는다 (공유 객체를 고치면 다른 칸까지 번진다)
      val.style = Object.assign({}, val.style, { numFmt: t[2] });
      label.alignment = { horizontal: "right", vertical: "middle" };
      val.alignment = { horizontal: "right", vertical: "middle" };
    });
    ws.addRow(head);
    const hr = ws.lastRow.number;
    ws.getRow(hr).font = { bold: true };
    ws.getRow(hr).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEFF3FB" } };

    /* 수량 열 — 원본 파일을 읽을 때 셀 값이 전부 문자열("1")로 넘어온다.
       그대로 적으면 엑셀이 '텍스트로 저장된 숫자'로 보고 SUM 이 0 이 된다.
       값이 온전히 숫자일 때만 진짜 숫자로 바꿔 적는다 (우편번호·전화번호처럼
       앞자리 0 이 뜻을 갖는 열은 건드리면 안 되므로 수량 열만).  */
    const qi = src.findIndex(h => /수량/.test(String(h || "")) && !/재고|누계|누적/.test(String(h || "")));
    const toNum = x => {
      if (typeof x === "number") return isFinite(x) ? x : null;
      const t = String(x == null ? "" : x).trim();
      if (!/^\d+(\.\d+)?$/.test(t) || /^0\d/.test(t)) return null;   // 앞자리 0 은 손대지 않는다
      const n = Number(t);
      return isFinite(n) ? n : null;
    };
    const qtyVals = [];

    /* 주문일이 이른 것부터 적는다 — 업체가 자기 출고 순서대로 훑을 수 있게.
       원본 v.rows 는 합계·검산이 함께 쓰므로 복사본을 정렬한다.
       주문일이 비어 있으면 수집일자로 대신하고, 그것도 없으면 원래 순서를 지킨다. */
    const dkey = r => String(r.orderDate || r.date || "").replace(/\D/g, "");
    const sorted = v.rows.map((r, i) => ({ r, i })).sort((a, b) => {
      const x = dkey(a.r), y = dkey(b.r);
      if (x && y && x !== y) return x < y ? -1 : 1;
      if (!!x !== !!y) return x ? -1 : 1;      // 날짜 없는 줄은 뒤로
      return a.i - b.i;                        // 같으면 원래 순서 유지
    }).map(o => o.r);

    sorted.forEach(r => {
      const idx = {};
      (r.srcCols || []).forEach((h, i) => { const t = String(h || "").trim(); if (t && idx[t] === undefined) idx[t] = i; });
      const line = src.map(h => (idx[h] === undefined ? "" : (r.raw ? r.raw[idx[h]] : "")));
      if (qi >= 0) {
        const n = toNum(line[qi]);
        if (n !== null) line[qi] = n;
        qtyVals.push(n);
      }
      // 개당 정산이면 배송비를 단가에 합쳐 적는다 (공급가 × 수량 = 정산금액 이 되게)
      const merged = r.shipMode === "건당" ? 0 : Math.round(r.ship || 0);
      line.push(r.priced === false ? "미확정" : (r.unitCost == null ? "" : Math.round(r.unitCost) + merged));
      if (perOrderShip) line.push(r.shipMode === "건당" ? Math.round(r.shipTotal || 0) : 0);
      line.push(Math.round(r.pay || 0));
      if (anyCarry) line.push(QO.carryNote(r));
      ws.addRow(line);
    });

    ws.addRow([]);
    // 합계는 금액 열 끝(정산금액)에 맞춘다 — 맨 끝에 '비고'가 붙어도 밀리지 않게
    const moneyEnd = nCol - (anyCarry ? 1 : 0);
    const add = (label, val, color) => {
      const cells = new Array(nCol).fill("");
      cells[moneyEnd - 2] = label; cells[moneyEnd - 1] = Math.round(val);
      const row = ws.addRow(cells);
      row.font = { bold: true, color: color ? { argb: color } : undefined };
      return row.getCell(moneyEnd);
    };
    const cSupply = add("공급가 합계", v.pay - perOrderShip);
    const cShip = perOrderShip ? add("배송비 합계", perOrderShip) : null;
    const cDed = v.ded ? add("CS 차감", -v.ded, "FFCC0000") : null;
    const cFinal = add("정산금액", v.final, "FF1A56DB");

    /* ---- 합계에 실제 엑셀 수식을 건다 ----
       숫자만 박아두면 업체가 줄을 지우거나 고쳐도 합계가 그대로라 서로 다른 값을 보게 된다.
       계산 결과(result)도 같이 넣어 두면 수식을 계산하지 않는 뷰어에서도 숫자가 보인다.
       ※ 수량 열은 원본 파일에서 온 값이라 숫자가 아닐 수 있다 → 합이 맞을 때만 수식으로 바꾼다. */
    if (v.rows.length) {
      const first = hr + 1, last = hr + v.rows.length;
      const L = n => ws.getColumn(n).letter;
      const payCol = L(moneyEnd);                                  // 정산금액
      const shipCol = perOrderShip ? L(src.length + 2) : null;     // 배송비(건당 업체만)
      const rng = c => `${c}${first}:${c}${last}`;
      const put = (cell, formula, result) => { if (cell) cell.value = { formula, result: Math.round(result) }; };

      // 공급가 합계 = 정산금액 합 − 배송비 합 (개당 업체는 배송비 열이 없어 그냥 합)
      put(cSupply, shipCol ? `SUM(${rng(payCol)})-SUM(${rng(shipCol)})` : `SUM(${rng(payCol)})`, v.pay - perOrderShip);
      put(cShip, `SUM(${rng(shipCol)})`, perOrderShip);
      // 정산금액 = 위 합계 칸들을 더한다 (같은 수를 두 번 세지 않게 SUM 을 다시 쓰지 않는다)
      const parts = [cSupply, cShip, cDed].filter(Boolean).map(c => c.address);
      put(cFinal, parts.join("+"), v.final);

      // 판매수량 = 원본 수량 열의 합.
      // 한 줄이라도 숫자로 못 바꿔 텍스트로 남았으면 SUM 이 그 줄을 빼먹으니 수식을 걸지 않는다.
      if (qi >= 0 && topCells["판매수량"] && qtyVals.length === v.rows.length
          && qtyVals.every(n => n !== null)
          && Math.round(qtyVals.reduce((s2, n) => s2 + n, 0)) === Math.round(totalQty))
        put(topCells["판매수량"], `SUM(${rng(L(qi + 1))})`, totalQty);
      // 위쪽 요약은 아래 합계 칸을 그대로 가리킨다 — 두 곳이 어긋날 수 없다
      const link = (name, cell) => { if (cell && topCells[name]) topCells[name].value = { formula: cell.address, result: Math.round(cell.value.result != null ? cell.value.result : cell.value) }; };
      link("공급가 합계", cSupply);
      link("배송비 합계", cShip);
      link("교환·반품 차감", cDed);
      link("정산금액", cFinal);
    }
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
    for (let n = src.length + 1; n <= moneyEnd; n++) { ws.getColumn(n).width = 14; ws.getColumn(n).numFmt = "#,##0"; }
    if (anyCarry) ws.getColumn(nCol).width = 20;      // 비고는 글자라 서식을 걸지 않는다
    alignSheet(ws, moneyEnd, src.length, hr);
    if (anyCarry) ws.getColumn(nCol).alignment = { horizontal: "center", vertical: "middle" };
    // alignSheet 는 열 단위로 정렬을 덮어쓴다 → 상단 요약은 그 뒤에 다시 오른쪽으로
    tops.forEach((t, i) => {
      const row = ws.getRow(i + 1);
      row.getCell(4).alignment = { horizontal: "right", vertical: "middle" };
      row.getCell(5).alignment = { horizontal: "right", vertical: "middle" };
    });
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
    // 몰 수수료(삼성계열 카드 등) — 우리 매출에서 떼가므로 마진에서 뺀다
    const feeLabels = [];
    vendors.forEach(v => (v.feeRows || []).forEach(g => {
      if (feeLabels.indexOf(g.label) < 0) feeLabels.push(g.label);
    }));
    feeLabels.forEach(lb => line(`└ ${lb}`, v =>
      -((v.feeRows || []).filter(g => g.label === lb).reduce((s2, g) => s2 + g.fee, 0)), { color: "FF8A6D00" }));
    // 파트너 MD 리워드 — 이걸 빼야 진짜 마진이다. 마진율도 최종 기준으로 적는다.
    const anyReward = vendors.some(v => v.reward) || feeLabels.length;
    if (anyReward) {
      (result && result.md ? result.md : []).forEach(m =>
        line(`└ MD 리워드 · ${m.md}`, v => -((v.rewardRows || [])
          .filter(x => x.md === m.md).reduce((s2, x) => s2 + x.reward, 0)), { color: "FF8A6D00" }));
      line(`${co} 최종 마진`, v => v.netMargin, { bold: true, color: "FF0A7A3D" });
    }
    const mg = v => (anyReward ? v.netMargin : v.margin);
    const rate = ws.addRow([anyReward ? "최종 마진율" : "마진율"].concat(vendors.map(v =>
      v.amount ? Math.round(mg(v) / v.amount * 1000) / 10 + "%" : "-"))
      .concat([sum(v => v.amount) ? Math.round(sum(mg) / sum(v => v.amount) * 1000) / 10 + "%" : "-"]));
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
    alignSheet(ws, cols.length + 2, 1, hr);   // 1열은 항목 이름, 2열부터 금액
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
        alignSheet(ws, nCol, 6, hr);          // 6열까지는 내용, 7열부터 금액
      }
    } else {
      for (const v of vendors) vendorSheet(wb, v);
    }
    return await wb.xlsx.writeBuffer();
  }

  /* 정산 메일에 채워 넣을 값 — 실제로 보낼 때와 '메일내용 수정' 미리보기가
     같은 함수를 쓴다. 미리보기에 견본 숫자가 나오면 진짜 금액인 줄 알고 헷갈린다. */
  function mailVars(v) {
    const p = (result && result.period) || { label: "" };
    const qty = v.rows.reduce((s2, r) => s2 + (r.qty || 0), 0);
    const summary = `· 건수: ${v.rows.length}건 (수량 ${qty}개)\n`
      + `· 공급가 합계: ${won(v.pay)}\n`
      + (v.ded ? `· CS 차감(교환·반품): -${won(v.ded)}\n` : "")
      + `· 지급액: ${won(v.final)} (부가세 포함)`
      + (v.unpriced ? `\n\n※ 단가가 아직 확정되지 않은 ${v.unpriced}건은 이번 지급액에서 빠져 있습니다. 단가 확정 후 정산해 드리겠습니다.` : "");
    return {
      회사: CONFIG.company, 업체: v.vendor, 정산월: p.label || QO.fmtDate(QO.todayStr()),
      정산기간: periodRange(),        // 그 달 1일~말일 (엑셀 머리말과 같은 값)
      날짜: QO.fmtDate(QO.todayStr()), 건수: v.rows.length, 수량: qty,
      지급액: won(v.final), 요약: summary,
    };
  }

  async function sendOne(v, inp, btn) {
    const to = parseEmails(inp.value);
    if (!to.length) { alert("받는사람 이메일을 입력해주세요."); return; }
    const old = btn.textContent;
    btn.disabled = true; btn.textContent = "보내는 중…";
    try {
      await ensureGmail();
      const buf = await buildExcel([v]);          // 업체용 — 매출·마진 없음
      const tpl = MAILTPL.render("settle", mailVars(v));
      await GMAIL.send({
        to: to.join(", "), subject: tpl.subject, body: tpl.body,
        attachments: [{ filename: fileName(v.vendor), data: buf }],
      });
      const ek = emailKey(v.vendor);              // ① 발주 탭과 같은 키로 이력을 쌓는다
      S.vendorSent = S.vendorSent || {};
      S.vendorSent[ek] = mergeRecent(S.vendorSent[ek], to);
      await DB.set("vendorSent", S.vendorSent);
      S.vendorEmails = S.vendorEmails || {};
      S.vendorEmails[ek] = inp.value.trim();
      await DB.set("vendorEmails", S.vendorEmails);
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
    el.textContent = "";      // 검색조건 안내 문구는 뺐다 (2026-08-05)
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
      try { await addPriceBook(await readFile(f), f.name); }
      catch (e) { msg("msg-pb", "err", "⚠ " + f.name + " — " + e.message); }
    };
    const pbFile = () => (pbRaw && pbRaw.name ? pbRaw.name.replace(/\.xls[xm]$/i, "") : "공급가표") + ".xlsx";
    $("pb-pv").onclick = async () => {
      try { openPreview(await priceBookExcel(), "업체별 공급가표"); }
      catch (e) { msg("msg-pb", "err", "⚠ " + e.message); }
    };
    $("pb-dl").onclick = async () => {
      try { download(await priceBookExcel(), pbFile()); }
      catch (e) { msg("msg-pb", "err", "⚠ " + e.message); }
    };
    $("pb-clear").onclick = async () => {
      if (!confirm("올려둔 공급가표를 해제할까요?\n해제하면 업체 지급액은 마진율 방식으로 계산됩니다.")) return;
      pbRaw = null; await DB.set("priceBook", null);
      rebuildBook(); drawPriceBook(); drawBrands(); drawMd(); if (result) calc();
      msg("msg-pb", "ok", "공급가표를 해제했어요.");
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
          try {
            const r = await GMAIL.driveFetchExcel(f.id);
            let mt = "";
            try { mt = (await GMAIL.driveFileInfo(f.id)).modifiedTime || ""; } catch (e2) {}
            // 어디서 가져왔는지 기억해두면, 다음에 켤 때 드라이브 최신본으로 자동으로 맞춘다
            await addPriceBook(r.buf, r.name || f.name, { drive: { id: f.id, name: r.name || f.name, mtime: mt } });
          } catch (e) { msg("msg-pb", "err", "⚠ " + f.name + " — " + e.message); }
        }
      },
    });
    $("st-md-add").onclick = async () => {
      rewards.push({ id: "md" + Date.now(), md: "", picks: {} });
      await saveRewards(); drawMd();
      const last = $("st-md-list").querySelectorAll(".mdname");
      if (last.length) last[last.length - 1].focus();
    };
    $("st-alias").onclick = () => {
      if (!hasBook()) { msg("msg-pb", "warn", "공급가표를 먼저 올려주세요 — 연결할 상대가 있어야 보여드릴 수 있어요."); return; }
      $("alias-msg").textContent = "";
      drawAliasBox();
      $("aliasmodal").classList.add("on");
    };
    $("alias-close").onclick = () => $("aliasmodal").classList.remove("on");
    $("aliasmodal").onclick = e => { if (e.target === $("aliasmodal")) $("aliasmodal").classList.remove("on"); };
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

  async function init() {
    bind(); await load(); drawFilterLine();
    // 켤 때 한 번 — 로그인이 늦게 붙는 경우가 있어 조금 뒤에 확인한다
    setTimeout(() => { refreshPriceBook().catch(() => {}); }, 3000);
  }
  async function onShow() {
    if (!drawn) { await load(); drawn = true; }
    drawFiles(); drawPriceBook(); drawBrands(); drawMd(); refresh();
    if (result) drawResult();
    refreshPriceBook().catch(() => {});     // 탭에 들어올 때도 (아직 확인 전이면)
  }

  init();
  return { onShow, drawFilter: drawFilterLine, markSettled, calc, result: () => result, periodRange,
           /* 검증용 — 업체용 정산서(합계 수식·머리말)를 화면 없이 만들어 볼 수 있게 열어둔다 */
           _make: (r, v) => { result = r; const wb = new ExcelJS.Workbook(); vendorSheet(wb, v); return wb; },
           // 실제 calc() 와 같은 순서로 돌린다 (수수료 → 리워드 → 최종 마진)
           _rewards: (rules, vendors) => { rewards = rules; calcMallFees(vendors); return calcRewards(vendors); },
           _money: moneyLines,
           _tpl: priceBookTemplate,
           _pbx: (raw) => { pbRaw = raw; return priceBookExcel(); } };
})();
window.ST = ST;
