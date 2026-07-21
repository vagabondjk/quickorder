/* ===================================================================
   퀵오더 — 엑셀 처리 로직 (발주변환.py 를 그대로 옮긴 것)
   브라우저와 Node 양쪽에서 동작. ExcelJS 필요.
   =================================================================== */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory(require("exceljs"));
  else root.QO = factory(root.ExcelJS);
})(typeof self !== "undefined" ? self : this, function (ExcelJS) {
"use strict";

/* ---------------- 표준 항목 정의 (ORDER_FIELDS) ---------------- */
const ORDER_FIELDS = [
  ["ORDER_NO",  ["주문번호"], ["회차","결제","배송","상품"]],
  ["PRODUCT",   ["상품명","제품명","상품"], ["코드","번호","회차","옵션","단품"]],
  ["OPTION",    ["옵션","단품명","단품"], ["코드"]],
  ["QTY",       ["수량"], []],
  ["AMOUNT",    ["결제금액","주문금액","판매금액","공급금액","금액"], ["할인"]],
  ["ORDER_DATE",["주문일시","주문일자","주문일"], ["지시","예정","예약","희망","완료","출하"]],
  ["PAY_DATE",  ["결제일시","결제일자","결제일"], []],
  ["RECIPIENT", ["수령인명","수령인","수취인","받는분","받는사람","이름"], ["전화","연락처","주소","코드","번호"]],
  ["RECIPIENT_PHONE", ["수령인연락처1","휴대폰번호","전화번호1","연락처1","수취인전화","휴대전화","휴대폰","전화번호","연락처"], ["가상","mail","2"]],
  ["RECIPIENT_PHONE2",["수령인연락처2","연락처2","전화번호2"], ["가상","mail"]],
  ["ADDRESS",   ["배송주소","전체받는사람주소","수령인주소","주소"], ["코드"]],
  ["MESSAGE",   ["배송메시지","배송메세지","고객배송요청사항","배송요청","요청사항","주문요청메시지","메시지","메세지"], []],
  ["ZIP",       ["우편번호"], []],
  ["ORDERER",   ["주문자명","주문자","구매자","보내는"], ["전화","연락처","mail","가상"]],
  ["CARRIER",   ["택배사"], []],
  ["INVOICE",   ["운송장","송장"], []],
];
const COPY_FIELDS = ORDER_FIELDS.map(f => f[0]).filter(n => n !== "CARRIER" && n !== "INVOICE");
const KEY_FIELDS = ["RECIPIENT","ADDR","PRODUCT","QTY","ORDERER","ZIP"];
const BRAND_HEADER = "브랜드";
const DATE_COL_KEYWORDS = ["수집일","주문일","일자","일시"];
const COLLECT_KEYWORDS = ["주문수집일","수집일자","수집일"];
const FIELD_KR = {ORDER_NO:"주문번호",PRODUCT:"상품",OPTION:"옵션",QTY:"수량",AMOUNT:"금액",
  ORDER_DATE:"주문일시",PAY_DATE:"결제일시",RECIPIENT:"수령인",RECIPIENT_PHONE:"연락처1",
  RECIPIENT_PHONE2:"연락처2",ADDRESS:"주소",MESSAGE:"배송메시지",ZIP:"우편번호",ORDERER:"주문자"};

/* ---------------- 셀 값 읽기 (ExcelJS 값 형태 정규화) ---------------- */
function cv(cell) {
  if (!cell) return null;
  let v = cell.value;
  if (v === null || v === undefined) return null;
  if (typeof v === "object") {
    if (v instanceof Date) return v;
    if (v.richText) return v.richText.map(t => t.text).join("");
    if (v.text !== undefined) return v.text;
    if (v.result !== undefined) return v.result;          // 수식 결과
    if (v.formula !== undefined) return null;
    if (v.hyperlink !== undefined) return v.text || v.hyperlink;
    return String(v);
  }
  return v;
}
function getV(ws, r, c) { try { return cv(ws.getRow(r).getCell(c)); } catch (e) { return null; } }
function isBlank(v) { return v === null || v === undefined || String(v).trim() === ""; }
function normHeader(v) {
  if (v === null || v === undefined) return "";
  return String(v).replace(/\s/g, "").replace(/\n/g, "");
}
function dims(ws) {
  let maxR = ws.rowCount || 0, maxC = ws.columnCount || 0;
  if (!maxC) { ws.eachRow({ includeEmpty: false }, row => { if (row.cellCount > maxC) maxC = row.cellCount; }); }
  return { rows: maxR, cols: maxC };
}

/* ---------------- 헤더 → 표준 항목 (_canon_field) ---------------- */
function canonField(h) {
  if (typeof h !== "string") { if (h === null || h === undefined) return null; h = String(h); }
  const s = normHeader(h);
  if (!s) return null;
  if (s.includes("택배사")) return "CARRIER";
  if (s.includes("운송장") || s.includes("송장")) return "INVOICE";
  if (s.includes("전화") || s.includes("연락처") || s.includes("휴대폰") || s.includes("핸드폰") || s.includes("번호2")) return null;
  if (s.includes("코드")) return null;
  if (s.includes("주문번호") || s.includes("결제번호") || s.includes("배송번호")) return null;
  if (s.includes("우편번호")) return "ZIP";
  if (s.includes("수량")) return "QTY";
  if (s.includes("옵션")) return null;
  if (s.includes("주소")) return "ADDR";
  if (s.includes("상품명")) return "PRODUCT";
  if (s.includes("상품") && !s.includes("번호")) return "PRODUCT";
  if (s.includes("수취인") || s.includes("수령인") || s.includes("받는사람") || s.includes("받는분") || s === "이름") return "RECIPIENT";
  if (s.includes("주문자") || s.includes("구매자") || s.includes("보내는")) return "ORDERER";
  return null;
}

/* ---------------- 헤더 행 찾기 ---------------- */
function findHeaderRow(ws, maxScan = 12) {
  const d = dims(ws);
  const lim = Math.min(d.rows, maxScan);
  for (let r = 1; r <= lim; r++) {
    let cnt = 0;
    const cmax = Math.min(d.cols, 60);
    for (let c = 1; c <= cmax; c++) if (canonField(getV(ws, r, c))) cnt++;
    if (cnt >= 3) return r;
  }
  return 1;
}

/* ---------------- 전화 열 ---------------- */
function phoneColumns(ws, headerRow) {
  const out = [], d = dims(ws);
  for (let c = 1; c <= d.cols; c++) {
    const v = getV(ws, headerRow, c);
    if (typeof v !== "string") continue;
    const s = normHeader(v);
    const hit = ["전화","연락처","휴대폰","핸드폰","휴대전화"].some(k => s.includes(k));
    const bad = ["가상","mail","이메일","메일"].some(k => s.includes(k));
    if (hit && !bad) out.push(c);
  }
  return out;
}

/* ---------------- 표준 항목 ↔ 열 매핑 ---------------- */
function buildOrderFieldMap(ws, headerRow, role = "target") {
  const d = dims(ws), headers = {};
  for (let c = 1; c <= d.cols; c++) {
    const v = getV(ws, headerRow, c);
    headers[c] = v ? normHeader(v) : "";
  }
  const result = {}, used = new Set();
  for (const [canon, patterns, excludes] of ORDER_FIELDS) {
    let bestCol = null, bestRank = patterns.length;
    for (let c = 1; c <= d.cols; c++) {
      const h = headers[c];
      if (used.has(c) || !h) continue;
      if (excludes.some(x => h.includes(x))) continue;
      for (let rank = 0; rank < patterns.length; rank++) {
        if (h.includes(patterns[rank])) { if (rank < bestRank) { bestRank = rank; bestCol = c; } break; }
      }
    }
    if (bestCol !== null) { result[canon] = bestCol; used.add(bestCol); }
  }
  // 연락처 재배정
  const phones = phoneColumns(ws, headerRow);
  delete result.RECIPIENT_PHONE; delete result.RECIPIENT_PHONE2;
  if (phones.length) {
    if (role === "target") {
      result.RECIPIENT_PHONE = phones[0];
      if (phones.length > 1) result.RECIPIENT_PHONE2 = phones[1];
    } else {
      let cand = phones.filter(c => ["수령인","수취인","받는"].some(k => headers[c].includes(k)));
      if (!cand.length && result.RECIPIENT) cand = phones.filter(c => c > result.RECIPIENT);
      if (!cand.length) cand = phones;
      result.RECIPIENT_PHONE = cand[0];
      const two = cand.slice(1).filter(c => headers[c].includes("2"));
      if (two.length) result.RECIPIENT_PHONE2 = two[0];
    }
  }
  return result;
}

/* ---------------- 시트 고르기 ---------------- */
function pickOrderSheet(wb) {
  let best = null, bestScore = -1;
  for (const ws of wb.worksheets) {
    let score;
    try {
      const hr = findHeaderRow(ws);
      score = Object.keys(buildOrderFieldMap(ws, hr, "source")).length;
    } catch (e) { continue; }
    const name = String(ws.name || "");
    if (["반품","교환","제품사양"].some(k => name.includes(k))) score -= 5;
    if ((ws.rowCount || 0) <= 2) score -= 3;
    if (score > bestScore) { bestScore = score; best = ws; }
  }
  return best || wb.worksheets[0];
}

/* ---------------- 브랜드 ---------------- */
function findBrandColumn(ws, headerRow) {
  const d = dims(ws);
  for (let c = 1; c <= d.cols; c++) {
    const v = getV(ws, headerRow, c);
    if (typeof v === "string" && v.trim() === BRAND_HEADER) return c;
  }
  return null;
}
function listBrands(wb) {
  const ws = pickOrderSheet(wb);
  const hr = findHeaderRow(ws);
  const bcol = findBrandColumn(ws, hr);
  if (!bcol) return [];
  const d = dims(ws), out = [], seen = new Set();
  for (let r = hr + 1; r <= d.rows; r++) {
    const v = getV(ws, r, bcol);
    if (v === null || v === undefined) continue;
    const s = String(v).trim();
    if (s && !seen.has(s)) { seen.add(s); out.push(s); }
  }
  return out;
}

/* ---------------- 날짜 ---------------- */
function extractDate(v) {
  if (v === null || v === undefined) return null;
  if (v instanceof Date) {
    const p = n => String(n).padStart(2, "0");
    return "" + v.getFullYear() + p(v.getMonth() + 1) + p(v.getDate());
  }
  const d = String(v).replace(/\D/g, "");
  return d.length >= 8 ? d.slice(0, 8) : null;
}
function isCollectHeader(h) {
  const s = String(h).replace(/\s/g, "");
  return COLLECT_KEYWORDS.some(k => s.includes(k));
}
function findDateColumns(ws, headerRow) {
  const out = [], d = dims(ws);
  for (let c = 1; c <= d.cols; c++) {
    const v = getV(ws, headerRow, c);
    if (typeof v !== "string") continue;
    const s = normHeader(v);
    if (DATE_COL_KEYWORDS.some(k => s.includes(k))) out.push([c, v.trim()]);
  }
  return out;
}
function defaultDateColumn(ws, headerRow) {
  const cols = findDateColumns(ws, headerRow);
  if (!cols.length) return [null, null];
  for (const [c, h] of cols) if (isCollectHeader(h)) return [c, h];
  return cols[0];
}
function orderDateInfo(wb, headerText) {
  const ws = pickOrderSheet(wb);
  const hr = findHeaderRow(ws);
  const cands = findDateColumns(ws, hr);
  if (!cands.length) return { counts: {}, header: null, candidates: [] };
  let dcol = null, dhdr = null;
  if (headerText) for (const [c, h] of cands) if (h === headerText) { dcol = c; dhdr = h; break; }
  if (dcol === null) { const r = defaultDateColumn(ws, hr); dcol = r[0]; dhdr = r[1]; }
  if (!dcol) return { counts: {}, header: null, candidates: cands.map(x => x[1]) };
  const counts = {}, d = dims(ws);
  for (let r = hr + 1; r <= d.rows; r++) {
    const dd = extractDate(getV(ws, r, dcol));
    if (dd) counts[dd] = (counts[dd] || 0) + 1;
  }
  return { counts, header: dhdr, candidates: cands.map(x => x[1]) };
}

/* ---------------- 값 후처리 ---------------- */
function formatPhone(v) {
  if (v === null || v === undefined) return v;
  const s = String(v).trim();
  let d = s.replace(/\D/g, "");
  if (!d) return v;
  if (!d.startsWith("0")) d = "0" + d;
  if (d.startsWith("02")) {
    if (d.length === 10) return `02-${d.slice(2,6)}-${d.slice(6)}`;
    if (d.length === 9)  return `02-${d.slice(2,5)}-${d.slice(5)}`;
  }
  if (d.startsWith("050") && d.length === 12) return `${d.slice(0,4)}-${d.slice(4,8)}-${d.slice(8)}`;
  if (d.length === 11) return `${d.slice(0,3)}-${d.slice(3,7)}-${d.slice(7)}`;
  if (d.length === 10) return `${d.slice(0,3)}-${d.slice(3,6)}-${d.slice(6)}`;
  if (d.length === 8)  return `${d.slice(0,4)}-${d.slice(4)}`;
  return s;
}
function stripHyphen(v) { return (v === null || v === undefined) ? v : String(v).replace(/[\s\-]/g, ""); }
function valueTransformForHeader(h) {
  if (typeof h !== "string") return null;
  const s = h.replace(/\s/g, "");
  if (s.includes("우편번호")) return stripHyphen;
  if (s.includes("가상")) return null;
  if (s.includes("연락처") || s.includes("전화") || s.includes("휴대폰") || s.includes("핸드폰")) return formatPhone;
  return null;
}

/* ---------------- 파일명 → 업체명 ---------------- */
function nameFromFilename(name) {
  let stem = String(name).replace(/\.[^.]+$/, "");
  stem = stem.replace(/^.*[\\/]/, "");
  let prev = null;
  while (prev !== stem) { prev = stem; stem = stem.replace(/^\s*(?:\d+|[A-Za-z]{1,2})\s*[.\-_)\]]+\s*/, ""); }
  stem = stem.trim();
  const parts = stem.split(/\s+/);
  return (parts[0] || stem).replace(/[\\/:*?"<>|]/g, "");
}

/* ---------------- 매칭 키 정규화 ---------------- */
function normKey(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "number" && Number.isInteger(v)) v = String(v);
  if (v instanceof Date) v = extractDate(v);
  let s = String(v).trim().replace(/[\s\-]/g, "");
  return s.toLowerCase();
}

/* ===================================================================
   발주서 변환 : 주문 → 업체 양식
   =================================================================== */
function convert(orderWb, tplWb, opts) {
  opts = opts || {};
  const brandFilter = opts.brands && opts.brands.length ? new Set(opts.brands.map(b => String(b).trim())) : null;
  const dateSet = opts.dates && opts.dates.length ? new Set(opts.dates.map(String)) : null;
  const log = opts.log || function () {};

  const sws = pickOrderSheet(orderWb);
  const srcHeaderRow = findHeaderRow(sws);
  const smap = buildOrderFieldMap(sws, srcHeaderRow, "source");

  // 업체 양식 시트: 표준항목이 가장 많이 잡히는 시트
  let tws = null, bestN = -1, tgtHeaderRow = 1, tmap = null;
  for (const w of tplWb.worksheets) {
    const hr = findHeaderRow(w);
    const m = buildOrderFieldMap(w, hr, "target");
    if (Object.keys(m).length > bestN) { bestN = Object.keys(m).length; tws = w; tgtHeaderRow = hr; tmap = m; }
  }
  if (!tws) throw new Error("업체 양식에서 시트를 찾지 못했습니다.");

  log(`[헤더 매칭] 주문 '${sws.name}'(헤더 ${srcHeaderRow}행) → 양식 '${tws.name}'(헤더 ${tgtHeaderRow}행)`);

  const pairs = [], matched = [];
  for (const canon of COPY_FIELDS) {
    if (tmap[canon] !== undefined && smap[canon] !== undefined) { pairs.push([smap[canon], [tmap[canon]]]); matched.push(canon); }
  }
  if (tmap.PAY_DATE !== undefined && smap.PAY_DATE === undefined && smap.ORDER_DATE !== undefined) {
    pairs.push([smap.ORDER_DATE, [tmap.PAY_DATE]]); matched.push("PAY_DATE*");
  }
  if (!pairs.length) throw new Error("헤더 이름으로 매칭되는 공통 항목을 찾지 못했습니다.\n주문/양식의 헤더(상품·수량·수령인·주소 등)를 확인하세요.");
  matched.forEach(c => log(`   ${FIELD_KR[c.replace("*","")] || c}`));

  // 브랜드 / 날짜 열
  let brandCol = null;
  if (brandFilter) {
    brandCol = findBrandColumn(sws, srcHeaderRow);
    if (!brandCol) throw new Error(`쇼핑몰 주문에서 '${BRAND_HEADER}' 열을 찾지 못했습니다.`);
  }
  let dateCol = null;
  if (dateSet) {
    const cands = findDateColumns(sws, srcHeaderRow);
    if (opts.dateHeader) { for (const [c, h] of cands) if (h === opts.dateHeader) { dateCol = c; break; } }
    if (!dateCol) dateCol = defaultDateColumn(sws, srcHeaderRow)[0];
    if (!dateCol) throw new Error("주문 파일에서 날짜 열을 찾지 못했습니다.");
  }

  // 대상 열 값 후처리
  const colTransform = {};
  const td = dims(tws);
  for (let c = 1; c <= td.cols; c++) {
    const tf = valueTransformForHeader(getV(tws, tgtHeaderRow, c));
    if (tf) colTransform[c] = tf;
  }

  const sd = dims(sws);
  let outRow = tgtHeaderRow + 1, count = 0;
  for (let r = srcHeaderRow + 1; r <= sd.rows; r++) {
    if (brandCol !== null) {
      const bv = getV(sws, r, brandCol);
      if (bv === null || !brandFilter.has(String(bv).trim())) continue;
    }
    if (dateCol !== null) {
      const dv = extractDate(getV(sws, r, dateCol));
      if (!dateSet.has(dv)) continue;
    }
    const vals = pairs.map(([scol, tcols]) => [tcols, getV(sws, r, scol)]);
    if (vals.every(([, v]) => isBlank(v))) continue;
    for (const [tcols, v] of vals) {
      for (const tcol of tcols) {
        let out = v;
        const tf = colTransform[tcol];
        if (tf && !isBlank(out)) out = tf(out);
        tws.getRow(outRow).getCell(tcol).value = (out === undefined ? null : out);
      }
    }
    outRow++; count++;
  }
  log(`총 ${count}건 기입`);
  return { count, sheet: tws.name };
}

/* ===================================================================
   송장 취합 : 업체 회신 → 송장취합양식
   =================================================================== */
function collectInvoices(sabangWb, replies, opts) {
  opts = opts || {};
  const log = opts.log || function () {};

  // 1) 회신들 읽기
  const loaded = [], errors = [];
  for (const rep of replies) {
    const supplier = rep.name ? nameFromFilename(rep.name) : "회신";
    let picked = null, bestScore = -1;
    for (const w of rep.wb.worksheets) {
      const hr = findHeaderRow(w);
      const d = dims(w); const fm = {};
      for (let c = 1; c <= d.cols; c++) { const cf = canonField(getV(w, hr, c)); if (cf && fm[cf] === undefined) fm[cf] = c; }
      const score = Object.keys(fm).length + (fm.INVOICE !== undefined ? 10 : 0);
      if (score > bestScore) { bestScore = score; picked = { ws: w, hr, fm }; }
    }
    if (!picked || picked.fm.INVOICE === undefined) { errors.push([supplier, 0, 0, "송장 열 없음"]); continue; }
    const { ws, hr, fm } = picked;
    const idf = KEY_FIELDS.filter(f => fm[f] !== undefined);
    const rows = [], d = dims(ws);
    for (let r = hr + 1; r <= d.rows; r++) {
      const inv = getV(ws, r, fm.INVOICE);
      const car = fm.CARRIER !== undefined ? getV(ws, r, fm.CARRIER) : null;
      if (isBlank(inv) && isBlank(car)) continue;
      const vals = {}; idf.forEach(f => { vals[f] = getV(ws, r, fm[f]); });
      rows.push({ car, inv, vals });
    }
    loaded.push({ supplier, fm, rows });
  }

  // 2) 대상 시트 후보
  const cands = [];
  for (const ws of sabangWb.worksheets) {
    const hr = findHeaderRow(ws);
    const d = dims(ws), fm = {};
    for (let c = 1; c <= d.cols; c++) { const cf = canonField(getV(ws, hr, c)); if (cf && fm[cf] === undefined) fm[cf] = c; }
    if (fm.CARRIER !== undefined && fm.INVOICE !== undefined &&
        KEY_FIELDS.filter(f => fm[f] !== undefined).length >= 2) cands.push({ ws, hr, fm });
  }
  if (!cands.length) throw new Error("송장취합양식에서 대상 시트(택배사/운송장 열)를 찾지 못했습니다.");

  // 3) 가장 많이 매칭되는 시트 선택
  function matchSheet(c) {
    const { ws, hr, fm } = c, d = dims(ws);
    const used = new Set(), fills = {}, per = [];
    let total = 0;
    for (const rp of loaded) {
      const common = KEY_FIELDS.filter(f => rp.fm[f] !== undefined && fm[f] !== undefined);
      if (!common.length) { per.push([rp.supplier, 0, 0, "공통 식별항목 없음"]); continue; }
      const index = new Map();
      for (let r = hr + 1; r <= d.rows; r++) {
        const key = common.map(f => normKey(getV(ws, r, fm[f]))).join("");
        if (common.every((f, i) => key.split("")[i] === "")) continue;
        if (!index.has(key)) index.set(key, []);
        index.get(key).push(r);
      }
      let vf = 0, vu = 0;
      for (const row of rp.rows) {
        const parts = common.map(f => normKey(row.vals[f]));
        if (parts.every(x => x === "")) continue;
        const key = parts.join("");
        let trow = null;
        for (const rr of (index.get(key) || [])) if (!used.has(rr)) { trow = rr; break; }
        if (trow === null) { vu++; continue; }
        used.add(trow);
        if (!fills[trow]) fills[trow] = {};
        if (!isBlank(row.car)) fills[trow][fm.CARRIER] = row.car;
        if (!isBlank(row.inv)) fills[trow][fm.INVOICE] = row.inv;
        vf++;
      }
      total += vf;
      per.push([rp.supplier, vf, vu, "OK"]);
    }
    return { total, per, fills, ...c };
  }
  let best = null;
  for (const c of cands) { const r = matchSheet(c); if (!best || r.total > best.total) best = r; }

  // 4) 실제 기입
  for (const [rowNo, cols] of Object.entries(best.fills))
    for (const [colNo, val] of Object.entries(cols))
      best.ws.getRow(Number(rowNo)).getCell(Number(colNo)).value = val;

  // 5) 송장 갯수 대조
  const perSrc = {};
  for (const rp of loaded) {
    const n = rp.rows.filter(x => !isBlank(x.inv)).length;
    perSrc[rp.supplier] = (perSrc[rp.supplier] || 0) + n;
  }
  const srcInvoice = Object.values(perSrc).reduce((a, b) => a + b, 0);
  const writtenInvoice = Object.values(best.fills).filter(c => c[best.fm.INVOICE] !== undefined).length;

  // 6) 취합본 빈칸(누락) 점검 — 주문행인데 송장이 안 채워진 행 찾기
  //    (업체 회신에서 해당 주문의 송장이 빠졌는지 확인)
  const bfm = best.fm, bhr = best.hr, bws = best.ws, bd = dims(bws);
  const keyPresent = KEY_FIELDS.filter(f => bfm[f] !== undefined);
  const labelFields = ["RECIPIENT", "PRODUCT", "ORDERER", "ADDR"].filter(f => bfm[f] !== undefined);
  const missing = [];
  let orderRows = 0;
  for (let r = bhr + 1; r <= bd.rows; r++) {
    // 주문행 판정: 핵심 식별항목 중 하나라도 값이 있으면 실제 주문행
    const hasKey = keyPresent.some(f => !isBlank(getV(bws, r, bfm[f])));
    if (!hasKey) continue;
    orderRows++;
    const invCell = getV(bws, r, bfm.INVOICE);   // 채운 뒤 값
    if (isBlank(invCell)) {
      const label = labelFields
        .map(f => { const v = getV(bws, r, bfm[f]); return v == null ? "" : String(v).trim(); })
        .filter(Boolean).join(" · ") || `${r}행`;
      missing.push({ row: r, label });
    }
  }
  const missingCount = missing.length;

  const per = best.per.concat(errors);
  log(`대상 시트 '${best.ws.name}' · 총 ${best.total}건 기입 · 회신 ${srcInvoice} / 취합본 ${writtenInvoice} · 주문행 ${orderRows} / 빈칸(누락) ${missingCount}`);
  return { total: best.total, per, srcInvoice, writtenInvoice, gap: srcInvoice - writtenInvoice,
    perSrc, sheet: best.ws.name, orderRows, missingCount, missing: missing.slice(0, 30) };
}

/* ---------------- 내용 미리보기 ---------------- */
function preview(wb, limit) {
  limit = limit || 2000;   // 최신 주문(아래쪽)까지 다 보이도록 넉넉히
  const ws = pickOrderSheet(wb);
  const hr = findHeaderRow(ws);
  const d = dims(ws);
  const cols = [];
  for (let c = 1; c <= d.cols; c++) { const v = getV(ws, hr, c); cols.push(v === null ? "" : String(v).trim()); }
  while (cols.length && !cols[cols.length - 1]) cols.pop();
  if (!cols.length) return { columns: [], rows: [], total: 0, keyIdx: [], sheet: ws.name };
  const KEY = new Set(["RECIPIENT","PRODUCT","QTY","ADDR","ORDERER"]);
  const keyIdx = [];
  cols.forEach((h, i) => { if (KEY.has(canonField(h))) keyIdx.push(i); });
  const bcol = findBrandColumn(ws, hr);
  if (bcol && bcol - 1 < cols.length && !keyIdx.includes(bcol - 1)) keyIdx.unshift(bcol - 1);
  const rows = []; let total = 0;
  for (let r = hr + 1; r <= d.rows; r++) {
    const vals = []; let empty = true;
    for (let c = 1; c <= cols.length; c++) { const v = getV(ws, r, c); vals.push(v); if (!isBlank(v)) empty = false; }
    if (empty) continue;
    total++;
    if (rows.length < limit) rows.push(vals.map(v => v === null || v === undefined ? "" : (v instanceof Date ? extractDate(v) : String(v))));
  }
  return { columns: cols, rows, total, keyIdx, sheet: ws.name };
}

/* ---------------- 범용 미리보기 (어떤 엑셀이든) ---------------- */
function previewAny(wb, limit) {
  limit = limit || 2000;   // 전체 내용이 다 보이도록 넉넉히
  // 시트 선택: '헤더 아래 실제 데이터 행이 가장 많은' 시트.
  //  (단순 rowCount는 서식만 있는 빈 시트가 크게 잡혀 잘못 골라짐)
  let ws = null, bestData = -1, wsFb = null, bestFb = -1;
  for (const w of wb.worksheets) {
    const d = dims(w);
    if (!d.rows || !d.cols) continue;
    const fb = (d.rows || 0) * 1000 + (d.cols || 0);
    if (fb > bestFb) { bestFb = fb; wsFb = w; }
    const hr = findHeaderRow(w);
    const cmax = Math.min(d.cols, 60);
    const rlim = Math.min(d.rows, hr + 5000);   // 과도한 스캔 방지
    let dataRows = 0;
    for (let r = hr + 1; r <= rlim; r++) {
      let any = false;
      for (let c = 1; c <= cmax; c++) { if (!isBlank(getV(w, r, c))) { any = true; break; } }
      if (any) dataRows++;
    }
    if (dataRows > bestData) { bestData = dataRows; ws = w; }
  }
  if (!ws || bestData <= 0) ws = wsFb || wb.worksheets[0];   // 전부 비었으면(빈 양식) 열만이라도
  if (!ws) return { columns: [], rows: [], total: 0, sheet: "", sheets: [] };
  const hr = findHeaderRow(ws);
  const d = dims(ws);
  const cols = [];
  for (let c = 1; c <= d.cols; c++) { const v = getV(ws, hr, c); cols.push(v === null ? "" : String(v).trim()); }
  while (cols.length && !cols[cols.length - 1]) cols.pop();
  const ncol = Math.max(cols.length, 1);
  const rows = []; let total = 0;
  for (let r = hr + 1; r <= d.rows; r++) {
    const vals = []; let empty = true;
    for (let c = 1; c <= ncol; c++) { const v = getV(ws, r, c); vals.push(v); if (!isBlank(v)) empty = false; }
    if (empty) continue;
    total++;
    if (rows.length < limit) rows.push(vals.map(v => v === null || v === undefined ? "" : (v instanceof Date ? extractDate(v) : String(v))));
  }
  return { columns: cols, rows, total, sheet: ws.name, sheets: wb.worksheets.map(w => w.name) };
}

/* ---------------- 워크북 헬퍼 ---------------- */
async function loadWorkbook(dataOrBuffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(dataOrBuffer);
  return wb;
}
async function saveWorkbook(wb) { return await wb.xlsx.writeBuffer(); }

function todayStr() {
  const d = new Date(), p = n => String(n).padStart(2, "0");
  return "" + d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate());
}
function fmtDate(ymd) { return ymd ? `${ymd.slice(0,4)}-${ymd.slice(4,6)}-${ymd.slice(6,8)}` : ""; }

return { ORDER_FIELDS, COPY_FIELDS, KEY_FIELDS, FIELD_KR, BRAND_HEADER,
  cv, getV, isBlank, dims, canonField, findHeaderRow, buildOrderFieldMap, phoneColumns,
  pickOrderSheet, findBrandColumn, listBrands, extractDate, isCollectHeader,
  findDateColumns, defaultDateColumn, orderDateInfo, formatPhone, stripHyphen,
  valueTransformForHeader, nameFromFilename, normKey,
  convert, collectInvoices, preview, previewAny, loadWorkbook, saveWorkbook, todayStr, fmtDate };
});
