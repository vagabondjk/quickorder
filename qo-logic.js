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
  /* ★ '상품판매유형'·'상품유형' 처럼 상품의 성격을 적은 열은 상품명이 아니다.
     LG 발주서에서 이게 상품명 자리를 차지해 발주서 상품칸이 전부 '기본' 으로 나갔다
     (진짜 상품명은 뒤쪽 '주문상품' 열이었다). 유형·구분 계열을 전부 막는다. */
  /* 업체 양식은 상품 칸 이름이 제각각이다 — 품목·품목명·품명·아이템명·발송제품.
     이게 안 잡히면 그 업체는 '상품 칸이 텅 빈' 발주서를 받는다. 건수는 맞아서
     눈으로는 안 보인다 (2026-08-18 실제 양식 14개 중 10개가 이랬다).
     ※ 앞에 있는 이름이 이긴다. '상품명' 을 '품명' 보다 먼저 둬야
       '상품명' 열이 '품명' 규칙에 걸려 엉뚱한 열로 가지 않는다. */
  ["PRODUCT",   ["상품명","제품명","주문상품","품목명","아이템명","발송제품","품명","품목","상품","아이템"],
                ["코드","번호","회차","옵션","단품","유형","구분","분류","상태","종류","카테고리","단가","금액"]],
  ["OPTION",    ["옵션","단품명","단품"], ["코드"]],
  ["QTY",       ["수량"], []],
  ["AMOUNT",    ["결제금액","주문금액","판매금액","공급금액","금액"], ["할인"]],
  // 수집일자 — 업체 양식에 이 칸이 있을 때만 채워진다(없으면 아무 일도 안 함).
  // 주문일시보다 먼저 둬서 '주문수집일자' 열을 주문일시가 채가지 않게 한다.
  ["COLLECT_DATE",["주문수집일자","주문수집일","수집일자","수집일"], []],
  ["ORDER_DATE",["주문일시","주문일자","주문일"], ["지시","예정","예약","희망","완료","출하"]],
  ["PAY_DATE",  ["결제일시","결제일자","결제일"], []],
  /* '수령자명'(메가존·이지웰·티딜·현대샵) 이 빠져 있어 수령인이 통째로 비었다 — 252줄.
     업체 양식 쪽은 '수화인명'(헤트라스)·'수신인'(프로퍼마켓)·'고객명'(신진) 도 쓴다. */
  ["RECIPIENT", ["수령인명","수령자명","수령인","수령자","수취인","수화인","수신인","받는분","받는사람","고객명","이름"],
                ["전화","연락처","휴대","주소","코드","번호","위치"]],
  ["RECIPIENT_PHONE", ["수령인연락처1","수령자휴대폰번호","휴대폰번호","전화번호1","연락처1","수취인전화","휴대전화","휴대폰","전화번호","연락처"], ["가상","mail","2"]],
  ["RECIPIENT_PHONE2",["수령인연락처2","연락처2","전화번호2"], ["가상","mail"]],
  // '수령지'(더그란)·'통합배송지'(모스스토리) 도 주소다
  ["ADDRESS",   ["배송주소","전체받는사람주소","수령인주소","주소","배송지","수령지"], ["코드","번호"]],
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
  COLLECT_DATE:"주문수집일자",ORDER_DATE:"주문일시",PAY_DATE:"결제일시",RECIPIENT:"수령인",RECIPIENT_PHONE:"연락처1",
  RECIPIENT_PHONE2:"연락처2",ADDRESS:"주소",MESSAGE:"배송메시지",ZIP:"우편번호",ORDERER:"주문자",
  // 이 둘이 빠져 있어 합친 표 헤더에 CARRIER/INVOICE 가 영문 그대로 찍혔다
  CARRIER:"택배사",INVOICE:"송장번호"};

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
  if (s.includes("옵션")) return "OPTION";
  /* '단품명' 은 옵션이다 (삼성계열·이제너두). 상품명보다 먼저 걸러야
     아래 '…명' 규칙에 상품으로 끌려가지 않는다. */
  if (s.includes("단품")) return "OPTION";
  if (s.includes("주소")) return "ADDR";
  /* ★ '상품판매유형' 처럼 상품의 '성격'을 적은 열이 상품명 자리를 차지하면
     발주서 상품칸이 통째로 '기본' 같은 값으로 나간다 (LG 파일에서 실제로 그랬다).
     이런 꼬리표가 붙은 열은 상품명이 아니다. */
  if (s.includes("상품") && /유형|구분|분류|상태|종류|카테고리/.test(s)) return null;
  if (s.includes("상품명")) return "PRODUCT";
  if (s.includes("상품") && !s.includes("번호")) return "PRODUCT";
  if (s.includes("수취인") || s.includes("수령인") || s.includes("수령자")
      || s.includes("받는사람") || s.includes("받는분") || s === "이름") return "RECIPIENT";
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
      let cand = phones.filter(c => ["수령인","수취인","수령자","받는"].some(k => headers[c].includes(k)));
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
/* 업체(브랜드) 열 찾기.
   쇼핑몰마다 이름이 다르다 — 삼성계열은 '브랜드', LG 는 '브랜드명'/'brandName'.
   ※ '브랜드코드'(brandNo) 는 숫자라 업체명이 아니다. 반드시 걸러낸다. */
const BRAND_HEADERS = ["브랜드", "브랜드명", "brandname"];
function findBrandColumn(ws, headerRow) {
  const d = dims(ws);
  let loose = null;
  for (let c = 1; c <= d.cols; c++) {
    const v = getV(ws, headerRow, c);
    if (typeof v !== "string") continue;
    const s = normHeader(v).toLowerCase();
    if (!s) continue;
    if (/코드|번호|no$|code$/.test(s)) continue;
    if (s === "브랜드") return c;                       // 정확히 맞으면 바로
    if (BRAND_HEADERS.includes(s) && loose === null) loose = c;
  }
  return loose;
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

/* 어떤 형태로 들어오든 진짜 Date 로 바꾼다 (못 바꾸면 null)
     · Date         → 그대로
     · 45000.604    → 엑셀 일련번호 (날짜 서식이 안 걸린 채 숫자로 들어온 경우)
     · "2026-07-28 14:30:00" / "2026.07.28" / "20260728" → 파싱
   ※ 일련번호는 시분초까지 직접 계산해 '로컬' Date 로 만든다.
      UTC 로 만들면 시간대 때문에 날짜가 하루 밀릴 수 있다. */
function toDateValue(v) {
  if (v === null || v === undefined || v === "") return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (typeof v === "number") {
    if (!isFinite(v) || v < 1 || v > 60000) return null;   // 엑셀에서 날짜로 볼 수 있는 범위 밖
    const days = Math.floor(v);
    const secs = Math.round((v - days) * 86400);
    const d = new Date(1899, 11, 30 + days);               // 엑셀 1900 체계(1900 윤년 버그 포함)
    d.setHours(Math.floor(secs / 3600), Math.floor(secs / 60) % 60, secs % 60, 0);
    return isNaN(d.getTime()) ? null : d;
  }
  const g = String(v).replace(/\D/g, "");
  if (g.length < 8) return null;
  const y = +g.slice(0, 4), mo = +g.slice(4, 6), da = +g.slice(6, 8);
  if (y < 1900 || y > 2200 || mo < 1 || mo > 12 || da < 1 || da > 31) return null;
  const d = new Date(y, mo - 1, da, +(g.slice(8, 10) || 0), +(g.slice(10, 12) || 0), +(g.slice(12, 14) || 0));
  return isNaN(d.getTime()) ? null : d;
}
/* 업체 양식의 이 열이 날짜 열인가 (헤더로 판단) */
function isDateHeader(h) {
  if (typeof h !== "string") return false;
  const s = normHeader(h);
  return !!s && DATE_COL_KEYWORDS.some(k => s.includes(k));
}
/* 이미 날짜 표시형식이 걸려 있는 셀인가 — 걸려 있으면 업체 양식 서식을 존중해 건드리지 않는다.
   "General"·"#,##0"·"@" 등에는 y/m/d/h/s 가 없으므로 걸러진다. */
function hasDateFormat(numFmt) {
  return typeof numFmt === "string" && /[ymdhs]/.test(numFmt.replace(/\[[^\]]*\]|"[^"]*"/g, ""));
}
/* 시간까지 보여주는 서식인가 — 업체 양식이 "yyyy-mm-dd hh:mm" 이면 00:00 이 붙어 보이므로 날짜만으로 바꿔준다 */
function hasTimeFormat(numFmt) {
  return typeof numFmt === "string" && /[hs]/.test(numFmt.replace(/\[[^\]]*\]|"[^"]*"/g, ""));
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
/* ---------------- 여러 쇼핑몰 주문 파일 합치기 ----------------
   쇼핑몰마다 열 이름·순서가 다르다. 그래서 통째로 이어붙이면 안 되고,
   각 파일을 '표준 항목' 으로 먼저 맞춘 뒤 그 결과를 합친다.
   합친 결과는 표준 이름(FIELD_KR)을 헤더로 쓴 워크북이라, convert() 가
   지금까지와 똑같이 읽는다 — 변환 로직은 손대지 않는다.

   · 값은 원래 형(날짜는 Date, 수량은 숫자)을 그대로 옮긴다. 문자열로 바꾸면
     날짜 서식·수량 계산이 깨진다.
   · 브랜드 열과 쇼핑몰명은 따로 챙긴다. 브랜드는 업체 선택에, 쇼핑몰명은
     어느 파일에서 왔는지 남기는 데 쓴다.
   · 표준 항목으로 잡히지 않은 열은 버린다. 파일마다 제각각이라 합칠 수 없다. */
/* ---------------- 업체(브랜드) 판정 ----------------
   쇼핑몰 14곳 중 브랜드 열이 있는 곳은 4곳뿐이다. 나머지는 상품명에서 알아내야 한다.
   상품명 앞에 [헤트라스] 처럼 업체가 붙어 있는 경우가 많지만, 그대로 믿으면 안 된다 —
   [베네특가]·[품절대란] 같은 행사 문구가 같은 자리에 온다. 그래서 '업체가 아닌 말'을
   먼저 걸러내고, 그래도 모르면 억지로 고르지 않고 미판정으로 남긴다.
   (현대샵 70건처럼 상품명에 업체 표시가 아예 없는 것들이 있다 — 사람만 안다) */
const NOT_BRAND = ["특가", "할인", "쿠폰", "단독", "한정", "품절", "대란", "인기", "베스트", "best",
  "행사", "이벤트", "무료배송", "당일발송", "신상", "리뉴얼", "세트", "증정", "사은품",
  "1+1", "2+1", "택1", "모음", "기획"];
/* 업체 이름이 아닌 대괄호인가.
   ★ 숫자가 들어간 대괄호는 규격이지 업체가 아니다 — [500mlx3개], [77mlx2개], [1013ml].
     이걸 안 막았더니 '500mlx3개' 라는 업체가 생겼다 (2026-08-18 실제로 나왔다).
     지금까지 본 업체 이름(헤트라스·내추럴이믹스·프로퍼마켓…)에는 숫자가 없다. */
const isPromoWord = b => {
  const s = String(b || "").trim(), low = s.toLowerCase();
  return !s || /\d/.test(s) || NOT_BRAND.some(w => low.includes(w));
};
/* 상품명 안의 대괄호를 앞에서부터 훑어 '행사 문구가 아닌' 첫 번째를 업체로 본다.
   [베네특가][헤트라스] … 처럼 행사 문구가 앞에 오거나
   디퓨져 BEST★[헤트라스]디퓨저… 처럼 중간에 오는 경우가 실제로 많다. */
function brandFromName(name) {
  const s = String(name == null ? "" : name);
  const all = s.match(/\[([^\]]{1,20})\]/g) || [];
  for (const raw of all) {
    const b = raw.slice(1, -1).trim();
    if (!isPromoWord(b)) return b;
  }
  return "";
}
/* 이름 안에 '아는 업체' 가 글자 그대로 들어 있는지 본다.
   [베네특가] 헤트라스 프리미엄… 처럼 대괄호 밖에 업체명이 있는 경우를 잡는다.
   ※ 아는 이름과 정확히 겹칠 때만 쓴다. 짐작이 아니라 대조다.
     두 글자 미만은 우연히 겹치기 쉬워 제외한다. */
function brandFromKnown(name, known) {
  if (!known || !known.length) return "";
  const s = String(name == null ? "" : name).replace(/\s+/g, "").toLowerCase();
  if (!s) return "";
  let best = "";
  for (const k of known) {
    const t = String(k || "").replace(/\s+/g, "").toLowerCase();
    if (t.length < 2 || !s.includes(t)) continue;
    if (t.length > best.replace(/\s+/g, "").length) best = k;   // 긴 이름이 더 확실하다
  }
  return best;
}
/* 한 줄의 업체를 정한다. 확실한 것부터 —
   ① 파일의 브랜드 열  ② 상품명 앞 [업체]  ③ 연결표(사람이 한 번 지정한 것)
   ※ 셋 다 아니면 "" 를 돌려준다. 찍지 않는다 — 잘못 찍으면 남의 업체로 주문이 나간다. */
function resolveBrand(rec, aliases) {
  const direct = String(rec.__brand == null ? "" : rec.__brand).trim();
  if (direct) return direct;
  const fromName = brandFromName(rec.PRODUCT);
  if (fromName) return fromName;
  if (aliases) {
    const key = aliasKey(rec.PRODUCT);
    if (key && aliases[key]) return aliases[key];
  }
  return "";
}
/* 연결표 열쇠 — 상품명을 느슨하게 다듬는다. 몰마다 앞뒤 문구가 조금씩 달라서
   완전 일치로 잡으면 몰이 하나 늘 때마다 다시 지정해야 한다. */
function aliasKey(name) {
  return String(name == null ? "" : name)
    .replace(/^\s*(?:★[^★]*★\s*)+/, "")
    .replace(/\[[^\]]*\]/g, " ")
    .replace(/[\s()（）,./·]+/g, "")
    .toLowerCase().slice(0, 40);
}

function mergeOrders(sources, opts) {
  opts = opts || {};
  const aliases = opts.aliases || null;
  const rows = [], fields = new Set();
  let brandSeen = false;
  (sources || []).forEach(src => {
    const wb = src.wb, mall = src.name || "";
    const ws = pickOrderSheet(wb);
    if (!ws) return;
    const hr = findHeaderRow(ws);
    const map = buildOrderFieldMap(ws, hr, "source");
    const brandCol = findBrandColumn(ws, hr);
    const maxRow = dims(ws).rows;
    for (let r = hr + 1; r <= maxRow; r++) {
      const rec = { __mall: mall, __corp: src.corp || "" };
      let any = false;
      Object.keys(map).forEach(canon => {
        const v = cv(ws.getRow(r).getCell(map[canon]));
        if (!isBlank(v)) { rec[canon] = v; fields.add(canon); any = true; }
      });
      if (brandCol) {
        const b = cv(ws.getRow(r).getCell(brandCol));
        if (!isBlank(b)) rec.__brand = b;
      }
      if (!any) continue;
      /* 브랜드 열이 없는 몰이라도 상품명·연결표로 업체를 정할 수 있다.
         그래서 '브랜드 열이 있었는가' 가 아니라 '실제로 정해졌는가' 로 판단한다 —
         예전엔 브랜드 열 있는 파일이 하나도 없으면 브랜드 칸 자체가 안 생겼다. */
      rec.__brand = resolveBrand(rec, aliases);
      rows.push(rec);
    }
  });
  /* 2차 — 이번에 확실히 알아낸 업체 이름들(+ 저장된 업체 양식 이름)을 사전 삼아
     아직 못 정한 줄의 상품명을 훑는다. 한 몰에서 브랜드 열로 알아낸 이름이
     다른 몰의 상품명 안에 글자 그대로 들어 있는 경우가 많다. */
  const known = [...new Set(rows.map(r => r.__brand).filter(Boolean).concat(opts.knownBrands || []))];
  rows.forEach(r => { if (!r.__brand) r.__brand = brandFromKnown(r.PRODUCT, known); });
  const brandSeen2 = rows.some(r => r.__brand);
  brandSeen = brandSeen || brandSeen2;
  /* 열 순서는 표준 항목 정의 순서를 따른다 — 사람이 열어봤을 때 늘 같은 자리에 있게 */
  const cols = ORDER_FIELDS.map(f => f[0]).filter(c => fields.has(c));
  const corpSeen = rows.some(r => r.__corp);
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("합친주문");
  const head = cols.map(c => FIELD_KR[c] || c);
  if (brandSeen) head.push(BRAND_HEADER);
  head.push("쇼핑몰명");
  if (corpSeen) head.push("법인");
  ws.addRow(head);
  rows.forEach(rec => {
    const line = cols.map(c => (rec[c] === undefined ? null : rec[c]));
    if (brandSeen) line.push(rec.__brand || null);
    line.push(rec.__mall);
    if (corpSeen) line.push(rec.__corp || null);
    const row = ws.addRow(line);
    /* 날짜는 셀 단위로 서식을 준다. 열 전체에 주면 수량·금액까지 날짜로 보인다
       (예전에 실제로 그런 사고가 있었다 — 상위 CLAUDE.md 참고) */
    cols.forEach((c, i) => {
      if (rec[c] instanceof Date) {
        const cell = row.getCell(i + 1);
        cell.style = Object.assign({}, cell.style, { numFmt: "yyyy-mm-dd" });
      }
    });
  });
  /* 업체를 못 정한 줄을 상품명별로 묶어 돌려준다 — 화면에서 한 번 지정하면
     연결표에 저장되고, 다음부터는 자동으로 붙는다. 조용히 빠뜨리지 않는다. */
  const unknown = new Map();
  rows.forEach(r => {
    if (r.__brand) return;
    const k = aliasKey(r.PRODUCT);
    if (!k) return;
    const cur = unknown.get(k);
    if (cur) { cur.count++; if (!cur.malls.includes(r.__mall)) cur.malls.push(r.__mall); }
    else unknown.set(k, { key: k, name: String(r.PRODUCT == null ? "" : r.PRODUCT), count: 1, malls: [r.__mall].filter(Boolean) });
  });
  return {
    wb, rows: rows.length, fields: cols,
    malls: [...new Set(rows.map(r => r.__mall).filter(Boolean))],
    corps: [...new Set(rows.map(r => r.__corp).filter(Boolean))],
    brands: [...new Set(rows.map(r => r.__brand).filter(Boolean))],
    unknown: [...unknown.values()].sort((a, b) => b.count - a.count),
    unknownRows: rows.filter(r => !r.__brand).length,
  };
}

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
  const dateCols = new Set();          // 날짜로 기입해야 하는 대상 열
  const td = dims(tws);
  for (let c = 1; c <= td.cols; c++) {
    const th = getV(tws, tgtHeaderRow, c);
    const tf = valueTransformForHeader(th);
    if (tf) colTransform[c] = tf;
    if (isDateHeader(th)) dateCols.add(c);
  }

  /* ★★ 양식에 남아 있던 지난번 주문을 먼저 지운다 (2026-08-18).
     업체 양식은 대개 '지난달 발주서 그대로'다 — 헤트라스 915줄, 아렌시아 392줄,
     코칸 수천 줄. 안 지우고 위에서부터 덮어쓰면 이번 주문 아래로 지난 주문이
     그대로 남아 업체가 같은 건을 다시 보낸다. 실제로 새 주문 줄에 지난 운송장번호가
     붙어 나가기도 했다 (이번 주문에는 송장이 없으니 옛 값이 그대로 남아서).
     값만 지우고 서식은 건드리지 않는다. 채워진 칸만 훑어서 큰 시트에서도 빠르다.
     ※ 남겨야 할 이유가 있으면 opts.keepExisting 으로 끌 수 있다. */
  if (!opts.keepExisting) {
    const rowsToClear = [];
    tws.eachRow({ includeEmpty: false }, (row, rn) => { if (rn > tgtHeaderRow) rowsToClear.push(rn); });
    rowsToClear.forEach(rn => {
      tws.getRow(rn).eachCell({ includeEmpty: false }, cell => { cell.value = null; });
    });
    if (rowsToClear.length) log(`   (양식에 있던 지난 주문 ${rowsToClear.length}줄을 비웠습니다)`);
  }

  const sd = dims(sws);
  /* 어느 쇼핑몰에서 몇 건인지 세어 둔다 — 결과 화면에 그대로 보여준다.
     합친 표에만 있는 열이라, 없으면 조용히 건너뛴다. */
  let mallCol = null;
  for (let c = 1; c <= sd.cols; c++) {
    if (normHeader(getV(sws, srcHeaderRow, c)) === "쇼핑몰명") { mallCol = c; break; }
  }
  const byMall = {};
  /* 날짜로 거를 때, 날짜 자체가 비어 있는 주문은 어느 날에도 안 걸린다.
     그냥 빠지면 그 주문은 어느 발주서에도 안 실린다 — 세어서 알려준다. */
  let noDate = 0;
  let outRow = tgtHeaderRow + 1, count = 0;
  for (let r = srcHeaderRow + 1; r <= sd.rows; r++) {
    if (brandCol !== null) {
      const bv = getV(sws, r, brandCol);
      if (bv === null || !brandFilter.has(String(bv).trim())) continue;
    }
    if (dateCol !== null) {
      const dv = extractDate(getV(sws, r, dateCol));
      if (!dv) { noDate++; continue; }
      if (!dateSet.has(dv)) continue;
    }
    const vals = pairs.map(([scol, tcols]) => [tcols, getV(sws, r, scol)]);
    if (vals.every(([, v]) => isBlank(v))) continue;
    for (const [tcols, v] of vals) {
      for (const tcol of tcols) {
        let out = v;
        const tf = colTransform[tcol];
        if (tf && !isBlank(out)) out = tf(out);
        const cell = tws.getRow(outRow).getCell(tcol);
        // 날짜 열은 '진짜 날짜'로 넣고 표시형식까지 걸어준다.
        // (안 걸면 엑셀·구글시트에서 46231.6 같은 일련번호로 보인다)
        const dv = dateCols.has(tcol) && !isBlank(out) ? toDateValue(out) : null;
        if (dv) {
          // 시간은 뺀다 — 업체에 나가는 발주서에는 날짜만 있으면 된다.
          // 값 자체를 자정으로 맞춰야 서식을 바꿔도 시간이 되살아나지 않는다.
          cell.value = new Date(dv.getFullYear(), dv.getMonth(), dv.getDate());
          // 날짜 서식이 없거나, 시간까지 보여주는 서식이면 날짜만 나오게 바꾼다.
          // (업체 양식이 "yyyy/mm/dd" 처럼 날짜만이면 그대로 존중)
          if (!hasDateFormat(cell.numFmt) || hasTimeFormat(cell.numFmt)) {
            // ★ cell.numFmt = ... 로 넣으면 안 된다.
            //   업체 양식은 표 전체가 같은 스타일을 '공유'하는 경우가 많아서,
            //   한 셀의 numFmt 를 바꾸면 공유 객체가 통째로 바뀌어 수량·금액까지 날짜로 보인다.
            //   반드시 새 스타일 객체를 만들어 이 셀에만 건다.
            cell.style = Object.assign({}, cell.style, { numFmt: "yyyy-mm-dd" });
          }
        } else {
          cell.value = (out === undefined ? null : out);
        }
      }
    }
    if (mallCol) {
      const mv = String(getV(sws, r, mallCol) || "").trim();
      if (mv) byMall[mv] = (byMall[mv] || 0) + 1;
    }
    outRow++; count++;
  }
  log(`총 ${count}건 기입`);
  return { count, sheet: tws.name, byMall, noDate };
}

/* 선택한 날짜 기준 '변환되어야 할' 주문 건수 (브랜드별) — 변환 결과 검증용 */
function countOrders(orderWb, opts) {
  opts = opts || {};
  const dateSet = opts.dates && opts.dates.length ? new Set(opts.dates.map(String)) : null;
  const ws = pickOrderSheet(orderWb);
  const hr = findHeaderRow(ws);
  const d = dims(ws);
  const bcol = findBrandColumn(ws, hr);
  let dcol = null;
  if (dateSet) {
    const cands = findDateColumns(ws, hr);
    if (opts.dateHeader) { for (const [c, h] of cands) if (h === opts.dateHeader) { dcol = c; break; } }
    if (!dcol) dcol = defaultDateColumn(ws, hr)[0];
  }
  const smap = buildOrderFieldMap(ws, hr, "source");
  const cols = Object.values(smap);
  let total = 0; const byBrand = {};
  for (let r = hr + 1; r <= d.rows; r++) {
    if (dcol && dateSet) { const dv = extractDate(getV(ws, r, dcol)); if (!dateSet.has(dv)) continue; }
    let any = false;
    for (const c of cols) { if (!isBlank(getV(ws, r, c))) { any = true; break; } }
    if (!any) continue;                       // 빈 행 제외
    total++;
    const bv = bcol ? getV(ws, r, bcol) : null;
    const b = bv == null ? "" : String(bv).trim();
    byBrand[b] = (byBrand[b] || 0) + 1;
  }
  return { total, byBrand };
}

/* ===================================================================
   송장 취합 : 업체 회신 → 송장취합양식
   =================================================================== */
/* 송장번호처럼 보이는 값인지.
   업체가 송장 칸에 '휴가라 8/6에 입력하겠습니다' 같은 문장을 적어 보내는 일이 있는데,
   빈칸이 아니라는 이유로 '기입 완료'로 세면 안 된다.
   송장번호는 숫자(가끔 영문 섞임) 8~20자리다. 한글이 들어가면 송장이 아니다. */
function looksLikeInvoice(v) {
  const s = String(v === null || v === undefined ? "" : v).trim();
  if (!s) return false;
  if (/[가-힣]/.test(s)) return false;                 // 한글이 있으면 문장
  const bare = s.replace(/[\s\-.]/g, "");
  if (!/^[A-Za-z0-9]+$/.test(bare)) return false;
  if (bare.length < 8 || bare.length > 20) return false;
  return /\d{6,}/.test(bare);                          // 숫자가 최소 6자리는 있어야
}
/* 택배사 칸도 이름이 아니라 문장이 들어오는 경우가 있다 */
function looksLikeCarrier(v) {
  const s = String(v === null || v === undefined ? "" : v).trim();
  if (!s) return true;                                  // 비어 있는 건 여기서 안 따진다
  if (s.length > 15) return false;
  if (/\d{5,}/.test(s)) return false;                   // 숫자 덩어리 = 송장이 잘못 들어감
  return true;
}

/* 택배사를 기억할 업체 키 — 파일명이 '디에스피_회신.xlsx', '디에스피 송장회신 8월.xlsx' 처럼
   달마다 조금씩 달라서, 앞의 날짜와 뒤의 '회신·양식·N월' 같은 꼬리를 떼고 업체명만 남긴다.
   이걸 안 하면 다음 달에 키가 달라져 기억해둔 택배사를 못 찾는다. */
function carrierKey(name) {
  let t = String(name == null ? "" : name).replace(/\.[^.]+$/, "").trim();
  t = t.replace(/^\d{6,8}[_\-\s]*/, "");
  for (let i = 0; i < 4; i++)
    t = t.replace(/[_\-\s]*(회신본|송장회신|회신|송장|발주양식|발주서|양식|사본|\d{1,2}월|\(\d+\))\s*$/, "").trim();
  return t || String(name == null ? "" : name);
}

function collectInvoices(sabangWb, replies, opts) {
  const carrierFills = [];      // 택배사를 대신 채워 넣은 내역 (화면에 알려준다)
  const carrierSeen = {};       // 업체별로 이번에 실제로 쓰인 택배사 (다음에 기억)
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
      const opt = fm.OPTION !== undefined ? getV(ws, r, fm.OPTION) : null;
      rows.push({ car, inv, vals, opt });
    }
    /* 택배사가 비어 있는 줄 채우기 —
       한 업체는 택배사가 거의 같은데 회신 파일에 몇 칸이 비어 오는 일이 있다.
       ① 같은 회신 파일에서 제일 많이 쓰인 택배사
       ② 없으면 지난번에 그 업체가 쓴 택배사 (opts.carriers 로 넘어온다)
       송장번호가 있는 줄만 채운다 — 송장도 없는 줄은 아직 출고 전이다. */
    const cnt = {};
    rows.forEach(x => { const c = String(x.car == null ? "" : x.car).trim();
      if (c && looksLikeCarrier(c)) cnt[c] = (cnt[c] || 0) + 1; });
    let top = ""; Object.keys(cnt).forEach(c => { if (!top || cnt[c] > cnt[top]) top = c; });
    const vkey = carrierKey(supplier);
    const remembered = (opts.carriers || {})[vkey] || (opts.carriers || {})[supplier] || "";
    const fill = top || remembered;
    let filled = 0;
    if (fill) rows.forEach(x => {
      if (isBlank(x.car) && !isBlank(x.inv)) { x.car = fill; x.carFilled = true; filled++; }
    });
    if (filled) carrierFills.push({ supplier: vkey, carrier: fill, n: filled, from: top ? "같은 파일" : "지난 회신" });
    if (top) carrierSeen[vkey] = top;          // 다음 달을 위해 기억할 값 (업체명만)
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
  const normInv = v => String(v == null ? "" : v).replace(/[^0-9a-zA-Z]/g, "").toLowerCase();
  // 취합본에 기입할 송장번호 — 하이픈·공백 등을 빼고 숫자만 남긴다. (1234-5678 → 12345678)
  //  · 영문이 섞인 운송장(해외 EMS 등)은 글자가 사라지면 안 되므로 구분기호만 뺀다.
  //    (EE123-456-789KR → EE123456789KR)
  //  · 뺄 게 없으면 원본 값을 그대로 반환 → 셀 형식(숫자/텍스트)이 바뀌지 않는다.
  //  · 실제로 뺀 경우에만 문자열로 넣는다 → 앞자리 0이 사라지지 않는다.
  const cleanInv = v => {
    const s = String(v == null ? "" : v).trim();
    if (!s) return v;
    // 송장이 아니라 메모·문장이면 손대지 않는다 ("휴가로 8/6에 입력" → "86" 이 되면 안 된다)
    if (/[가-힣]/.test(s)) return v;
    const out = /[A-Za-z]/.test(s) ? s.replace(/[\s\-.·]/g, "") : s.replace(/\D/g, "");
    if (!out) return v;          // 숫자가 하나도 없는 값 → 원본 유지
    return out === s ? v : out;
  };
  const SEP = "";   // 키 필드 구분자(연결 시 경계 뭉개짐 방지)
  function matchSheet(c) {
    const { ws, hr, fm } = c, d = dims(ws);
    const used = new Set(), fills = {}, per = [], ambiguous = [];
    let total = 0, already = 0;
    // 이미 채워진 송장값(중복 회신='이미 취합됨' 감지용)
    const existing = new Set();
    for (let r = hr + 1; r <= d.rows; r++) { const iv = getV(ws, r, fm.INVOICE); if (!isBlank(iv)) existing.add(normInv(iv)); }
    for (const rp of loaded) {
      const common = KEY_FIELDS.filter(f => rp.fm[f] !== undefined && fm[f] !== undefined);
      if (!common.length) { per.push([rp.supplier, 0, 0, "공통 식별항목 없음"]); continue; }
      const useOpt = rp.fm.OPTION !== undefined && fm.OPTION !== undefined;   // 옵션으로 세부 구분 가능?
      // 빈칸(송장 없는) 행만 후보로 인덱싱
      const index = new Map();
      for (let r = hr + 1; r <= d.rows; r++) {
        if (!isBlank(getV(ws, r, fm.INVOICE))) continue;   // 이미 송장 있는 행 제외 → 덮어쓰기 방지
        const parts = common.map(f => normKey(getV(ws, r, fm[f])));
        if (parts.every(x => x === "")) continue;
        const key = parts.join(SEP);
        if (!index.has(key)) index.set(key, []);
        index.get(key).push({ row: r, opt: useOpt ? normKey(getV(ws, r, fm.OPTION)) : "" });
      }
      let vf = 0, vu = 0, dup = 0;
      for (const row of rp.rows) {
        const parts = common.map(f => normKey(row.vals[f]));
        if (parts.every(x => x === "")) continue;
        const key = parts.join(SEP);
        const cands = (index.get(key) || []).filter(x => !used.has(x.row));
        if (!cands.length) {
          if (!isBlank(row.inv) && existing.has(normInv(row.inv))) dup++;   // 이미 취합됨(중복 회신)
          else vu++;                                                        // 진짜 미매칭
          continue;
        }
        let pick;
        if (cands.length === 1) { pick = cands[0]; }
        else {
          // 후보 여럿 → 옵션으로 유일하게 좁혀지면 그걸로
          const ro = useOpt ? normKey(row.opt) : "";
          const narrowed = (useOpt && ro) ? cands.filter(x => x.opt === ro) : [];
          if (narrowed.length === 1) { pick = narrowed[0]; }
          else {
            pick = (narrowed.length ? narrowed : cands)[0];
            // 유일하게 못 가림 → '확인 필요'로 기록(조용히 오배정 방지)
            ambiguous.push({
              supplier: rp.supplier, inv: row.inv == null ? "" : String(row.inv),
              label: common.map(f => row.vals[f]).filter(v => !isBlank(v)).map(String).join(" · "),
              option: (useOpt && row.opt != null) ? String(row.opt) : "",
              count: cands.length,
            });
          }
        }
        used.add(pick.row);
        if (!fills[pick.row]) fills[pick.row] = {};
        if (!isBlank(row.car)) fills[pick.row][fm.CARRIER] = row.car;
        if (!isBlank(row.inv)) fills[pick.row][fm.INVOICE] = cleanInv(row.inv);
        vf++;
      }
      total += vf; already += dup;
      per.push([rp.supplier, vf, vu, dup ? ("이미취합 " + dup + "건") : "OK"]);
    }
    return { total, per, fills, already, ambiguous, ...c };
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
  const cellStr = (r, f) => { const v = getV(bws, r, bfm[f]); return v == null ? "" : String(v).trim(); };
  const missing = [], odd = [];
  let orderRows = 0;
  for (let r = bhr + 1; r <= bd.rows; r++) {
    // 주문행 판정: 핵심 식별항목 중 하나라도 값이 있으면 실제 주문행
    const hasKey = keyPresent.some(f => !isBlank(getV(bws, r, bfm[f])));
    if (!hasKey) continue;
    orderRows++;
    const invCell = getV(bws, r, bfm.INVOICE);   // 채운 뒤 값
    // 송장 칸에 송장이 아닌 값(문장 등)이 들어온 행 — 빈칸이 아니라고 넘기면 안 된다
    if (!isBlank(invCell) || bfm.CARRIER !== undefined) {
      const carCell = bfm.CARRIER === undefined ? "" : getV(bws, r, bfm.CARRIER);
      const badInv = !isBlank(invCell) && !looksLikeInvoice(invCell);
      const badCar = !isBlank(carCell) && !looksLikeCarrier(carCell);
      if (badInv || badCar) {
        const it = { row: r, invoice: isBlank(invCell) ? "" : String(invCell).trim(),
                     carrier: isBlank(carCell) ? "" : String(carCell).trim(),
                     badInvoice: badInv, badCarrier: badCar };
        it.label = ["RECIPIENT", "PRODUCT", "ORDERER"]
          .map(f => (bfm[f] === undefined ? "" : cellStr(r, f))).filter(Boolean).join(" · ") || `${r}행`;
        odd.push(it);
      }
    }
    if (isBlank(invCell)) {
      const item = { row: r };
      // 항목별로 담아 UI에서 표로 보여줄 수 있게
      ["RECIPIENT", "PRODUCT", "OPTION", "QTY", "ORDERER", "ADDR"].forEach(f => {
        if (bfm[f] !== undefined) item[f] = cellStr(r, f);
      });
      item.label = ["RECIPIENT", "PRODUCT", "ORDERER", "ADDR"]
        .map(f => item[f]).filter(Boolean).join(" · ") || `${r}행`;
      missing.push(item);
    }
  }
  const missingCount = missing.length;
  const oddCount = odd.length;

  const already = best.already || 0;
  const per = best.per.concat(errors);
  log(`대상 시트 '${best.ws.name}' · 총 ${best.total}건 기입 · 회신 ${srcInvoice} / 취합본 ${writtenInvoice} · 주문행 ${orderRows} / 빈칸(누락) ${missingCount} / 송장 아닌 값 ${oddCount}`);
  return { total: best.total, per, srcInvoice, writtenInvoice, gap: srcInvoice - writtenInvoice - already, already,
    perSrc, sheet: best.ws.name, orderRows, missingCount, missing: missing.slice(0, 500),
    oddCount, odd: odd.slice(0, 500),
    carrierFills, carrierSeen,          // 택배사를 대신 채운 내역 · 업체별로 기억할 택배사
    ambiguous: best.ambiguous || [] };
}

/* ---------------- 내용 미리보기 ---------------- */
function preview(wb, limit, opts) {
  opts = opts || {};
  limit = limit || 5000;
  const ws = pickOrderSheet(wb);
  const hr = findHeaderRow(ws);
  const d = dims(ws);
  const cols = [];
  for (let c = 1; c <= d.cols; c++) { const v = getV(ws, hr, c); cols.push(v === null ? "" : String(v).trim()); }
  while (cols.length && !cols[cols.length - 1]) cols.pop();
  if (!cols.length) return { columns: [], rows: [], rowDates: [], total: 0, keyIdx: [], sheet: ws.name };
  const KEY = new Set(["RECIPIENT","PRODUCT","QTY","ADDR","ORDERER"]);
  const keyIdx = [];
  cols.forEach((h, i) => { if (KEY.has(canonField(h))) keyIdx.push(i); });
  const bcol = findBrandColumn(ws, hr);
  if (bcol && bcol - 1 < cols.length && !keyIdx.includes(bcol - 1)) keyIdx.unshift(bcol - 1);

  // 각 행의 날짜(수집일자 등)를 함께 돌려줘서, 화면에서 체크한 날짜만 즉시 걸러 보여줄 수 있게 함
  let dcol = null;
  const cands = findDateColumns(ws, hr);
  if (opts.dateHeader) { for (const [c, h] of cands) if (h === opts.dateHeader) { dcol = c; break; } }
  if (!dcol) dcol = defaultDateColumn(ws, hr)[0];

  const rows = [], rowDates = []; let total = 0;
  for (let r = hr + 1; r <= d.rows; r++) {
    const vals = []; let empty = true;
    for (let c = 1; c <= cols.length; c++) { const v = getV(ws, r, c); vals.push(v); if (!isBlank(v)) empty = false; }
    if (empty) continue;
    total++;
    if (rows.length < limit) {
      rows.push(vals.map(v => v === null || v === undefined ? "" : (v instanceof Date ? extractDate(v) : String(v))));
      rowDates.push(dcol ? extractDate(getV(ws, r, dcol)) : "");
    }
  }
  return { columns: cols, rows, rowDates, total, keyIdx, sheet: ws.name };
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

/* ==================================================================
   업체별 공급가표 · 정산 계산
   업체에 지급할 단가는 통합 파일에 없다. 따로 올린 '업체별 공급가표'에서 끌어온다.
   매칭 키는 브랜드 + 상품명 + 옵션.
   상품명이 마케팅 문구라('카레 우동 밀키트 500g x 3팩 / 일본 카레의 …')
   완전 일치만 보면 문구를 조금만 고쳐도 매칭이 끊긴다. 그래서 3단계로 좁혀 간다.
     ① 완전 일치  ② '/' 앞부분(상품 본체)만 일치  ③ 앞부분 접두 일치
   돈이 나가는 계산이라, 후보가 여러 개인데 단가가 서로 다르면 고르지 않고 '모호'로 남긴다.
   ================================================================== */
/* 접두 일치로 인정할 최소 길이(정규화 후).
   '수피아토 쿨타월'(7) · '아동비치모자'(6) 처럼 짧은 모델명이 실제로 쓰여서 6 으로 뒀다.
   짧게 잡아도 ① 한쪽이 다른 쪽의 '앞부분 전체'여야 하고 ② 후보가 둘 이상이면 거부하므로
   아무 상품에나 붙지는 않는다. */
const PRICE_PREFIX_MIN = 6;

/* 매칭용 정규화 — 공백·괄호·기호를 걷어내고 소문자로.
   'pk' 와 '팩' 은 같은 말로 본다 (공급가표는 '3pk', 주문 파일은 '3팩' 으로 적힌다). */
function normPriceText(v) {
  return String(v === null || v === undefined ? "" : v)
    .replace(/\s+/g, "")
    .replace(/[()\[\]{}<>·・,，.。/\\_+*~!?'"`|:;＋]/g, "")
    .toLowerCase()
    .replace(/(\d+)(pk|pack|개입세트)/g, "$1팩");
}
/* 앞머리에 붙은 [브랜드] 를 뗀다.
   공급가표는 '[현우동] 카레우동 3pk' 처럼 브랜드를 이름 앞에 달아두는데,
   주문 파일 상품명에는 그 말이 없어서 그대로 두면 낱말 비교가 늘 어긋난다.
   브랜드는 어차피 별도 열로 따로 맞추므로 떼도 정보가 사라지지 않는다. */
function stripLeadBracket(v) {
  return String(v === null || v === undefined ? "" : v)
    .replace(/^\s*[\[(【][^\]\)】]*[\]\)】]\s*/, "");
}
/* 금액 읽기 — 못 읽으면 null (0 과 구분해야 해서 0 을 돌려주지 않는다) */
function toPriceNumber(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === "number") return isFinite(v) ? v : null;
  const s = String(v).replace(/[^0-9.\-]/g, "");
  if (!s || s === "-" || s === "." || s === "-.") return null;
  const n = Number(s);
  return isFinite(n) ? n : null;
}
/* 한 행에서 매칭 키 3종을 뽑는다 */
function priceKeyParts(brand, product, option) {
  const raw = String(product === null || product === undefined ? "" : product);
  const bare = stripLeadBracket(raw);
  return {
    b: normPriceText(brand),
    pFull: normPriceText(raw),
    pBase: normPriceText(bare.split("/")[0]),   // 앞머리 [브랜드] 를 떼고, '/' 뒤 문구도 뗀다
    o: normPriceText(option),
  };
}
function prefixHit(a, b) {
  if (!a || !b) return false;
  const short = a.length <= b.length ? a : b;
  const long = a.length <= b.length ? b : a;
  return short.length >= PRICE_PREFIX_MIN && long.indexOf(short) === 0;
}

/* 상품명을 낱말로 쪼갠다 (2글자 미만은 버린다) */
function priceTokens(v) {
  return stripLeadBracket(v)
    .split(/[\s/·・,，.。+＋()\[\]{}<>~!?'"`|:;_\\-]+/)
    .map(normPriceText)
    .filter(t => t.length >= 2);
}
/* 표에 적은 낱말이 주문 상품명 안에 '전부' 들어 있는가.
   '수피아토 웻타월' 처럼 모델명 낱말이 상품명 중간에 흩어져 있는 경우를 잡는다.
     표: 수피아토 웻타월  ←→  주문: 수피아토 물없이 샤워하는 웻타월 화이트 민트 대형
   낱말이 하나뿐이거나 너무 짧으면(합쳐 6글자 미만) 아무 데나 붙을 수 있어 쓰지 않는다. */
function tokenHit(tokens, hay) {
  if (!tokens || tokens.length < 2 || !hay) return false;
  let len = 0;
  for (let i = 0; i < tokens.length; i++) {
    if (hay.indexOf(tokens[i]) < 0) return false;
    len += tokens[i].length;
  }
  return len >= PRICE_PREFIX_MIN;
}

/* 공급가표 만들기 — rows: [{brand, product, option, price, from}]
   from(적용시작일)은 선택이다. 없으면 그 단가가 '언제나' 적용된다(예전과 같은 동작).
   같은 상품을 적용일만 달리해 여러 줄 넣으면 주문일 기준으로 골라 쓴다.
   (2026-07 랩노마드 실제 사례: 쿨타월 공급가가 7/13 부터 7,370 → 8,730 으로 바뀌었다.
    단가 하나짜리 표로는 그 달 정산이 통째로 틀어진다) */
function buildPriceBook(rows, opts) {
  opts = opts || {};
  // 업체별 배송비 방식 — 공급가 파일의 '업체별 배송비 정산 방법' 시트에서 온다
  const vendorShip = opts.shipModes || {};
  const items = [], errors = [], seen = {};
  (rows || []).forEach((r, i) => {
    const line = i + 1;
    const brand = String(r.brand === null || r.brand === undefined ? "" : r.brand).trim();
    const product = String(r.product === null || r.product === undefined ? "" : r.product).trim();
    const option = String(r.option === null || r.option === undefined ? "" : r.option).trim();
    // 상품명이 없으면 상품 줄이 아니다 (빈 줄·소계 줄·머리말 등) — 조용히 넘어간다.
    // 예전엔 이걸 오류로 띄웠는데, 엑셀 아래쪽 빈 줄까지 잔뜩 잡혀서 시끄럽기만 했다.
    if (!product) return;
    // 공급가 칸이 아예 비어 있으면 상품 줄이 아니다 ('* 배송비 포함으로…' 같은 안내문 줄).
    // 오류로 볼 건 '가격 칸에 뭔가 적혀 있는데 숫자가 아닌' 경우뿐이다.
    if (r.price === undefined || r.price === null || String(r.price).trim() === "") return;
    const price = toPriceNumber(r.price);
    if (price === null) { errors.push({ line, why: "공급단가를 숫자로 읽을 수 없어요", brand, product, option, raw: r.price }); return; }
    if (price <= 0) { errors.push({ line, why: "공급단가가 0 이하예요", brand, product, option, price }); return; }
    const from = extractDate(r.from) || "";      // 적용시작일(선택) — yyyymmdd
    const to = extractDate(r.to) || "";          // 적용종료일(선택) — 행사가 여기 걸린다
    if (from && to && from > to) {
      errors.push({ line, why: "적용시작일이 종료일보다 늦어요", brand, product, option, from, to });
      return;
    }
    const ship = toPriceNumber(r.ship) || 0;     // 배송비(선택)
    const vendor = String(r.vendor === null || r.vendor === undefined ? "" : r.vendor).trim();
    // 배송비를 수량만큼 붙일지, 주문 한 건에 한 번만 붙일지.
    // 줄에 적혀 있으면 그걸, 없으면 업체별 설정을, 그것도 없으면 '개당'.
    //   "주문당 배송비 1회 정산" → 건당 / "배송비 포함 공급가 그대로 정산" → 개당
    const modeText = String(
      (r.shipMode === null || r.shipMode === undefined || r.shipMode === "" ? vendorShip[vendor] : r.shipMode) || ""
    );
    const shipMode = /건당|주문당|1회|1번|한번|한 번/.test(modeText) ? "건당" : "개당";
    const k = priceKeyParts(brand, product, option);
    const id = [k.b, k.pFull, k.o, from, to].join("|");
    if (Object.prototype.hasOwnProperty.call(seen, id)) {
      if (seen[id].price !== price)
        errors.push({ line, why: "같은 상품이 같은 적용기간에 단가가 다르게 두 번 있어요",
                      brand, product, option, price, before: seen[id].price });
      return;
    }
    const it = { brand, product, option, price, from, to, ship, shipMode, vendor,
                 key: [k.b, k.pFull, k.o].join("|"),
                 b: k.b, pFull: k.pFull, pBase: k.pBase, o: k.o, tk: priceTokens(product) };
    seen[id] = it;
    items.push(it);
  });
  return { items, errors };
}

/* 주문 줄을 가리키는 안정된 키 — 연결표(별칭) 저장에 쓴다 */
function priceRowKey(brand, product, option) {
  const k = priceKeyParts(brand, product, option);
  return [k.b, k.pFull, k.o].join("|");
}

/* 주문 한 줄에 맞는 공급단가 찾기
   aliases: {주문키: 공급가표키} — 이름이 달라 자동으로 못 붙는 상품을 사람이 한 번 이어준 것.
   별칭이 있으면 이름 비교를 건너뛰고 그 상품으로 바로 간다(적용기간은 그대로 따진다). */
function matchPrice(book, row, aliases) {
  const items = (book && book.items) || [];
  if (!items.length) return { ok: false, why: "공급가표가 비어 있어요" };
  const d0 = extractDate(row.date) || "";
  if (aliases) {
    const want = aliases[priceRowKey(row.brand, row.product, row.option)];
    if (want) {
      const hit = items.filter(it => it.key === want);
      if (!hit.length) return { ok: false, why: "연결해 둔 상품이 지금 공급가표에 없어요" };
      const p = pickByDate(hit, d0);
      if (!p.ok) return { ok: false, why: p.why, how: "연결표" };
      const ps = [];
      p.hit.forEach(h => { if (ps.indexOf(h.price) < 0) ps.push(h.price); });
      if (ps.length > 1)
        return { ok: false, why: "연결해 둔 상품의 단가가 여러 개예요", how: "연결표", candidates: p.hit.slice(0, 5) };
      return { ok: true, how: "연결표", price: ps[0], item: p.hit[0], from: p.hit[0].from || "" };
    }
  }
  const k = priceKeyParts(row.brand, row.product, row.option);
  const bookHasBrand = items.some(it => it.b);
  // 표에 브랜드 열이 아예 없으면 브랜드는 따지지 않는다
  const brandOk = it => !bookHasBrand || !k.b || it.b === k.b;
  // 표의 옵션이 비어 있으면 그 상품의 모든 옵션에 적용된다
  const optOk = it => !it.o || it.o === k.o;

  const stages = [
    ["완전일치", it => brandOk(it) && optOk(it) && it.pFull && it.pFull === k.pFull],
    ["본체일치", it => brandOk(it) && optOk(it) && it.pBase && it.pBase === k.pBase],
    ["앞부분일치", it => brandOk(it) && optOk(it) && prefixHit(it.pBase, k.pBase)],
    ["핵심어일치", it => brandOk(it) && optOk(it) && tokenHit(it.tk, k.pBase)],
  ];
  const d = extractDate(row.date) || "";
  for (let i = 0; i < stages.length; i++) {
    let hit = items.filter(stages[i][1]);
    if (!hit.length) continue;
    const picked = pickByDate(hit, d);
    if (!picked.ok) return { ok: false, why: picked.why, how: stages[i][0], candidates: hit.slice(0, 5) };
    hit = picked.hit;
    const prices = [];
    hit.forEach(h => { if (prices.indexOf(h.price) < 0) prices.push(h.price); });
    if (prices.length > 1)
      return { ok: false, why: "공급가표에 후보가 여러 개인데 단가가 서로 달라요", how: stages[i][0], candidates: hit.slice(0, 5) };
    return { ok: true, how: stages[i][0], price: prices[0], item: hit[0], from: hit[0].from || "" };
  }
  return { ok: false, why: "공급가표에서 못 찾았어요" };
}

/* 적용시작일로 후보 좁히기
   주문일 이하인 것들 중 '가장 늦게 시작한' 단가를 쓴다. 적용일이 없는 줄은 언제나 후보.
   표에 적용일이 하나도 없으면 예전과 똑같이 동작한다. */
function pickByDate(hit, d) {
  if (!hit.some(it => it.from || it.to)) return { ok: true, hit };
  if (!d) {
    const ps = [];
    hit.forEach(h => { if (ps.indexOf(h.price) < 0) ps.push(h.price); });
    if (ps.length > 1) return { ok: false, why: "적용시작일이 있는데 주문 날짜를 몰라 단가를 고를 수 없어요" };
    return { ok: true, hit };
  }
  const eff = hit.filter(it => (!it.from || it.from <= d) && (!it.to || d <= it.to));
  if (!eff.length) return { ok: false, why: "그 날짜에 적용되는 공급단가가 없어요 (적용기간을 확인해주세요)" };
  let maxFrom = "";
  eff.forEach(it => { if ((it.from || "") > maxFrom) maxFrom = it.from || ""; });
  return { ok: true, hit: eff.filter(it => (it.from || "") === maxFrom) };
}

/* '3팩'·'950g'·'3인분' 같은 크기·수량 낱말은 아무 상품에나 다 붙어서
   후보를 고르는 데 도움이 안 된다. 점수에서 뺀다. */
function isSizeToken(t) {
  return /^\d+(g|kg|ml|l|매|팩|개|개입|인분|입|세트|box|ea)?$/.test(t);
}
/* 글자가 순서대로 나타나는지 (연속일 필요 없음) — 줄임말 판별용 */
function subseqIn(tok, hay) {
  let i = 0;
  for (let c = 0; c < hay.length && i < tok.length; c++) if (hay[c] === tok[i]) i++;
  return i === tok.length;
}

/* 연결표에서 고를 후보 추리기.
   상품 21개를 통째로 늘어놓으면 사람이 못 고른다 — 이름이 겹치는 것만,
   많이 겹치는 순서로 준다. '닭갈비'면 닭갈비, '일떡'이면 일떡이 위로 온다.
   반환: [{item, score}] (점수 높은 순). 겹치는 낱말이 하나도 없으면 후보에서 뺀다. */
function rankPriceCandidates(book, row) {
  const items = (book && book.items) || [];
  const k = priceKeyParts(row.brand, row.product, row.option);
  const oTok = priceTokens(row.product);
  const seen = {}, out = [];
  items.forEach(it => {
    if (seen[it.key]) return;
    seen[it.key] = 1;
    let name = 0, miss = 0;
    (it.tk || []).forEach(t => {
      if (isSizeToken(t)) return;
      if (k.pBase.indexOf(t) >= 0) { name += 3; return; }                      // 표의 낱말이 주문명에 있나
      // '즉떡'(즉석떡볶기)·'일떡'(일반떡볶기) 같은 줄임말은 통째로는 안 들어 있다.
      // 글자가 순서대로 나타나면 약하게 가산해 둘을 구분한다. (추천용일 뿐 매칭 판정에는 안 쓴다)
      if (t.length <= 3 && subseqIn(t, k.pBase)) { name += 2; return; }
      miss++;   // 표 이름에만 있고 주문명엔 없는 낱말 — 많을수록 엉뚱한 상품이다
    });
    oTok.forEach(t => {                                                        // 주문 낱말이 표 이름에 있나
      if (isSizeToken(t)) return;
      if (it.pBase.indexOf(t) >= 0) name += 2;
    });
    if (!name) return;                        // 이름이 하나도 안 겹치면 후보 아님
    const sameBrand = k.b && it.b && k.b === it.b;
    out.push({ item: it, score: name - miss + (sameBrand ? 10 : 0), sameBrand: !!sameBrand });
  });
  out.sort((a, b) => b.score - a.score ||
    String(a.item.product).localeCompare(String(b.item.product)));
  return out;
}

/* 정산 계산
   rows: [{brand, vendor, product, option, qty, unitPrice, amount, ...}]
     · 매출  = unitPrice × 수량  (unitPrice 없으면 amount 를 그대로)
     · 지급액 = 공급단가 × 수량
   못 찾은 줄은 지급액을 0 으로 두고 unmatched 에 모아 화면에 띄운다 — 조용히 넘기지 않는다. */
function settle(rows, book, opts) {
  opts = opts || {};
  const aliases = opts.aliases || null;
  const out = [], unmatched = [];
  // 브랜드 → 업체. 단가를 못 찾은 줄도 업체는 알 수 있게 공급가표에서 미리 뽑아둔다.
  // (이게 없으면 미매칭 줄만 브랜드 이름으로 떨어져 나가 정산서가 둘로 갈린다)
  const brandVendor = {};
  ((book && book.items) || []).forEach(it => {
    if (it.b && it.vendor && !brandVendor[it.b]) brandVendor[it.b] = it.vendor;
  });
  (rows || []).forEach(r => {
    const qty = Number(r.qty) || 0;
    const unit = r.unitPrice === undefined || r.unitPrice === null || r.unitPrice === ""
      ? null : toPriceNumber(r.unitPrice);
    const revenue = unit === null ? (toPriceNumber(r.amount) || 0) : unit * qty;
    // 단가(행사 적용기간)는 '주문일' 기준으로 본다 — 고객이 언제 주문했는지가 기준.
    // 정산 귀속(어느 달 정산인지)은 r.date(수집일자)로 따로 잡는다.
    const priceRow = r.orderDate ? Object.assign({}, r, { date: r.orderDate }) : r;
    const m = matchPrice(book, priceRow, aliases);
    const it = m.ok ? (m.item || {}) : {};
    // 단가는 줄마다 원 단위로 반올림한 뒤 수량을 곱한다 (정산서 줄 금액이 눈으로 맞아떨어지게)
    const unitCost = m.ok ? Math.round(m.price) : null;
    const ship = m.ok ? Math.round(it.ship || 0) : 0;
    const shipMode = it.shipMode || "개당";
    const shipTotal = m.ok ? ship * (shipMode === "건당" ? (qty > 0 ? 1 : 0) : qty) : 0;
    const pay = m.ok ? unitCost * qty + shipTotal : 0;
    const o = {
      brand: r.brand || "",
      // 업체는 공급가표의 '업체명'을 우선한다 (브랜드 여러 개가 한 업체로 묶인다).
      // 단가를 못 찾았어도 브랜드로 업체를 되짚어 같은 정산서에 남긴다.
      vendor: (m.ok && it.vendor) || brandVendor[normPriceText(r.brand)]
              || r.vendor || r.brand || "(업체 미지정)",
      product: r.product || "", option: r.option || "",
      date: r.date || "", orderDate: r.orderDate || "", mall: r.mall || "", orderNo: r.orderNo || "",
      qty, unitPrice: unit, revenue,
      matched: m.ok, how: m.ok ? m.how : "", why: m.ok ? "" : m.why,
      unitCost,
      ship, shipMode, shipTotal,
      priceFrom: m.ok ? (m.from || "") : "",     // 어느 적용일의 단가를 썼는지
      pay,
      margin: m.ok ? revenue - pay : 0,
    };
    if (!m.ok) unmatched.push(o);
    out.push(o);
  });

  const byVendor = {};
  out.forEach(o => {
    const g = byVendor[o.vendor] = byVendor[o.vendor] ||
      { vendor: o.vendor, rows: [], qty: 0, revenue: 0, pay: 0, margin: 0, ship: 0, unmatched: 0, brands: [] };
    g.rows.push(o); g.qty += o.qty; g.revenue += o.revenue; g.pay += o.pay;
    g.margin += o.margin; g.ship += o.shipTotal;
    if (o.brand && g.brands.indexOf(o.brand) < 0) g.brands.push(o.brand);
    if (!o.matched) g.unmatched++;
  });
  const vendors = Object.keys(byVendor).map(k => byVendor[k]).sort((a, b) => b.pay - a.pay);
  const sum = f => vendors.reduce((s, v) => s + v[f], 0);
  const res = {
    rows: out, vendors, unmatched,
    total: { count: out.length, qty: sum("qty"), revenue: sum("revenue"),
             pay: sum("pay"), margin: sum("margin"), ship: sum("ship"),
             unmatched: unmatched.length },
  };
  res.check = settleCheck(rows || [], res);
  return res;
}

/* 검산 — 주문이 하나도 빠지지 않고 정산에 담겼는지 확인한다.
   금액이 작다고 넘어가면 안 된다. 한 건이라도 어긋나면 화면에 띄운다. */
function settleCheck(input, res) {
  const issues = [];
  const err = (why, detail) => issues.push({ level: "error", why, detail: detail || "" });
  const warn = (why, detail) => issues.push({ level: "warn", why, detail: detail || "" });

  const inCount = input.length;
  const inQty = input.reduce((s, r) => s + (Number(r.qty) || 0), 0);
  const outCount = res.vendors.reduce((s, v) => s + v.rows.length, 0);
  const outQty = res.vendors.reduce((s, v) => s + v.qty, 0);
  const outPay = res.vendors.reduce((s, v) => s + v.pay, 0);

  if (inCount !== outCount)
    err(`주문 ${inCount}건 중 ${outCount}건만 정산에 담겼어요`, `${inCount - outCount}건이 사라졌습니다`);
  if (inQty !== outQty)
    err(`주문 수량 ${inQty}개와 정산 수량 ${outQty}개가 달라요`, `${inQty - outQty}개 차이`);

  // 줄별 지급액을 다시 계산해 합계와 맞는지 본다
  let recomputed = 0, badRows = 0;
  res.rows.forEach(r => {
    const want = r.matched ? (r.unitCost * r.qty + r.shipTotal) : 0;
    if (want !== r.pay) badRows++;
    recomputed += want;
  });
  if (badRows) err(`지급액이 단가×수량과 맞지 않는 줄이 ${badRows}건 있어요`);
  if (Math.round(recomputed) !== Math.round(outPay))
    err(`줄별 지급액 합계(${Math.round(recomputed)})와 업체별 합계(${Math.round(outPay)})가 달라요`);

  // 단가를 못 찾아 0 원으로 둔 건 — 정산이 덜 된 상태다
  if (res.unmatched.length) {
    const q = res.unmatched.reduce((s, r) => s + r.qty, 0);
    const kinds = [];
    res.unmatched.forEach(r => { const k = (r.brand || "") + "/" + r.product; if (kinds.indexOf(k) < 0) kinds.push(k); });
    err(`단가를 못 찾아 지급액 0 원으로 둔 주문이 ${res.unmatched.length}건 (수량 ${q}) 있어요`,
        `상품 ${kinds.length}종 — 연결표로 이어주면 해결됩니다`);
  }
  // 수량이 비어 있으면 지급액이 0 이 된다
  const zero = res.rows.filter(r => !r.qty);
  if (zero.length) warn(`수량이 비어 있는 주문이 ${zero.length}건 있어요`, "그 줄은 지급액이 0 원입니다");
  // 업체를 못 정한 줄
  const noVendor = res.rows.filter(r => r.vendor === "(업체 미지정)");
  if (noVendor.length) warn(`업체를 정하지 못한 주문이 ${noVendor.length}건 있어요`);

  return {
    ok: !issues.some(i => i.level === "error"),
    issues,
    orderCount: inCount, orderQty: inQty,
    settledCount: outCount, settledQty: outQty,
    pricedCount: res.rows.filter(r => r.matched).length,
    pricedQty: res.rows.filter(r => r.matched).reduce((s, r) => s + r.qty, 0),
    pay: outPay,
  };
}

/* 정산서 시트의 열 구성.
   ★ 업체용에는 '매출'과 '우리마진'이 절대 들어가면 안 된다.
     업체가 볼 것은 '무엇을 몇 개, 단가 얼마에, 얼마 받는지'까지다.
     화면(qo-settle.js)에서 직접 배열을 만들지 말고 반드시 이 함수를 쓸 것 —
     여기 한 곳만 지키면 실수로 마진이 새어 나가지 않는다. */
const SETTLE_HEAD_VENDOR = ["정산일", "쇼핑몰", "주문번호", "상품명", "옵션", "수량", "공급단가", "지급액"];
const SETTLE_HEAD_INTERNAL = ["정산일", "쇼핑몰", "주문번호", "상품명", "옵션", "수량", "매출", "공급단가", "우리마진", "지급액"];

/* 업체로 나갈 표에서 빼야 할 열인지.
   ★ 통합 파일에는 TAG가·원가·판매가·결제금액이 들어 있다.
     '원가'는 이름과 달리 우리가 쇼핑몰에 넘기는 값(우리 매출)이라 절대 나가면 안 되고,
     판매가·결제금액도 업체가 알 필요가 없다. 애매하면 빼는 쪽으로 잡았다. */
function isPriceHeader(h) {
  const s = String(h === null || h === undefined ? "" : h).replace(/\s/g, "").toLowerCase();
  if (!s) return false;
  return /tag가|tag|원가|판매가|소비자가|공급가|공급단가|정산금액|결제금액|주문금액|매출|마진|수익|이익|단가|금액|수수료|할인|부가세|세액/.test(s);
}
/* 우리 내부에서만 쓰는 열 — 업체용 정산서에서는 뺀다 (2026-08-04 지정).
   수집일자·모델명은 우리 관리용이고, 주문번호 두 개는 우리 시스템/몰 쪽 번호다.
   ※ 여기서 뺀 열은 업체가 볼 수 없다. 늘리거나 줄일 때는 업체가 자기 주문과
     대조할 수 있는 정보(수취인·상품명·옵션·수량)가 남는지 확인할 것. */
function isInternalHeader(h) {
  const s = String(h === null || h === undefined ? "" : h).replace(/\s/g, "");
  if (!s) return false;
  if (/^(주문)?수집일자$/.test(s)) return true;
  if (/^모델명$/.test(s)) return true;
  if (/^배송메[세시]지$/.test(s)) return true;
  if (/사방넷/.test(s)) return true;                       // 주문번호(사방넷)
  if (/주문번호/.test(s) && /쇼핑몰|몰/.test(s)) return true; // 주문번호(쇼핑몰)
  return false;
}
/* 업체용 표에 남길 열 목록. 원래 파일의 열 순서를 지키고,
   가격 열·내부 전용 열·빈 헤더를 걷어낸다. */
function vendorSheetColumns(colLists) {
  const out = [];
  (colLists || []).forEach(cols => (cols || []).forEach(h => {
    const t = String(h === null || h === undefined ? "" : h).trim();
    if (!t) return;
    if (isPriceHeader(t) || isInternalHeader(t)) return;
    if (out.indexOf(t) < 0) out.push(t);
  }));
  return out;
}

/* 이월 건 표기 — 주문한 달과 수집(출고)된 달이 다르면 비고에 적는다.
   주문일이 7/31 이어도 발주마감 뒤 주문이면 수집이 8/1 이라 8월 정산으로 넘어온다.
   업체가 정산서를 볼 때 '왜 7월 주문이 8월 정산에 있지?' 하지 않도록 이유를 남긴다. */
function carryNote(row) {
  const o = extractDate(row && row.orderDate);
  const d = extractDate(row && row.date);
  if (!o || !d) return "";
  const om = o.slice(0, 6), dm = d.slice(0, 6);
  if (om === dm) return "";
  const mm = s => Number(s.slice(4, 6)) + "월";
  return `${mm(om)} 주문 · ${mm(dm)} 수집`;
}

function settleSheetHead(internal) {
  return (internal ? SETTLE_HEAD_INTERNAL : SETTLE_HEAD_VENDOR).slice();
}
function settleSheetRow(r, internal) {
  r = r || {};
  const n = v => Math.round(Number(v) || 0);
  // 단가를 못 찾은 줄을 업체 쪽에 '0원'으로만 보이면 오해를 사므로 '미확정'으로 적는다
  const notPriced = r.priced === false || r.matched === false;
  // '배송비 포함 공급가 그대로 정산'(개당)인 업체는 배송비를 단가에 합쳐 적는다.
  // 그래야 공급단가 × 수량 = 지급액 으로 눈에 맞아떨어진다.
  // '주문당 1회'(건당)는 개당 단가에 못 합치므로 그대로 둔다.
  const perUnitShip = r.shipMode === "건당" ? 0 : n(r.ship);
  const unit = notPriced ? "미확정"
    : (r.unitCost === null || r.unitCost === undefined ? "" : n(r.unitCost) + perUnitShip);
  const base = [r.date || "", r.mall || "", r.orderNo || "", r.product || "", r.option || "", Number(r.qty) || 0];
  if (!internal) return base.concat([unit, n(r.pay)]);
  const revenue = r.amount === undefined || r.amount === null ? r.revenue : r.amount;
  return base.concat([n(revenue), unit, n(r.margin), n(r.pay)]);
}

/* 시트를 전부 미리보기 — 공급가 파일처럼 여러 장에 나눠 담긴 경우에 쓴다.
   (previewAny 는 데이터가 제일 많은 시트 '하나'만 돌려줘서 행사 시트를 놓친다) */
function previewSheets(wb, limit) {
  limit = limit || 5000;
  const out = [];
  for (const ws of wb.worksheets) {
    const d = dims(ws);
    if (!d.rows || !d.cols) { out.push({ name: ws.name, columns: [], rows: [], headerRow: 1 }); continue; }
    const read = r => {
      const out = [];
      for (let c = 1; c <= d.cols; c++) {
        const v = getV(ws, r, c);
        out.push(v === null ? "" : String(v).trim().replace(/\s*\n\s*/g, " "));
      }
      while (out.length && !out[out.length - 1]) out.pop();
      return out;
    };
    // findHeaderRow 는 발주서 기준(수령인·상품명…)이라 '업체명/배송비정산' 같은
    // 작은 표에서는 못 찾고 1행을 준다. 그 줄이 비어 있으면 '채워진 칸이 가장 많은 줄'로 다시 찾는다.
    let hr = findHeaderRow(ws);
    let cols = read(hr);
    if (cols.filter(Boolean).length < 2) {
      let best = hr, bestN = cols.filter(Boolean).length;
      for (let r = 1; r <= Math.min(d.rows, 10); r++) {
        const n = read(r).filter(Boolean).length;
        if (n > bestN) { bestN = n; best = r; }
      }
      hr = best; cols = read(hr);
    }
    const ncol = Math.max(cols.length, 1);
    const rows = [];
    for (let r = hr + 1; r <= d.rows; r++) {
      const vals = []; let empty = true;
      for (let c = 1; c <= ncol; c++) { const v = getV(ws, r, c); vals.push(v); if (!isBlank(v)) empty = false; }
      if (empty) continue;
      if (rows.length < limit)
        rows.push(vals.map(v => v === null || v === undefined ? "" : (v instanceof Date ? extractDate(v) : v)));
    }
    out.push({ name: ws.name, columns: cols, rows, headerRow: hr });
  }
  return out;
}

/* ---------------- 워크북 헬퍼 ---------------- */
/* ---------------- 구버전 .xls (엑셀 97-2003) ----------------
   .xlsx 는 XML 을 묶은 zip 이고 .xls 는 그 자체가 바이너리 덩어리(OLE2/BIFF)라
   아예 다른 형식이다. ExcelJS 는 .xlsx 만 다루므로 .xls 는 SheetJS 로 읽어
   ExcelJS 워크북으로 옮겨 담는다. 그 뒤로는 모든 처리가 지금까지와 똑같다.

   ※ 업체가 양식을 .xlsx 로 바꿔주지 않는 경우가 많아 앱이 떠안는다 (2026-08-18).
   ※ 값·시트이름·열너비만 옮긴다. 글꼴·색 같은 꾸밈은 넘어오지 않는다 —
     실제 업체 양식(헤트라스·코칸)은 병합 없는 단순 표라 잃을 것이 없었다.
     결과물은 .xlsx 로 나간다. */
const OLE2 = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
function isOldXlsBuffer(data) {
  try {
    const u = new Uint8Array(data.buffer && data.byteLength !== undefined ? data.buffer : data,
                             data.byteOffset || 0, 8);
    return OLE2.every((b, i) => u[i] === b);
  } catch (e) { return false; }
}
function sheetJS() {
  if (typeof XLSX !== "undefined") return XLSX;                       // 브라우저 (CDN)
  if (typeof require === "function") { try { return require("xlsx"); } catch (e) {} }  // Node
  return null;
}
/* SheetJS 로 읽은 것을 ExcelJS 워크북으로 옮긴다.
   값은 형을 지켜서 넣는다 — 날짜는 Date, 수량은 숫자. 문자열로 바꾸면 뒤에서
   날짜 서식·수량 계산이 전부 깨진다. */
function xlsToWorkbook(data) {
  const X = sheetJS();
  if (!X) throw new Error("구버전 엑셀(.xls) 을 읽을 수 없습니다. .xlsx 로 저장해 주세요.");
  const src = X.read(data, { type: "array", cellDates: true, cellNF: false, cellStyles: false });
  const wb = new ExcelJS.Workbook();
  src.SheetNames.forEach(name => {
    const sws = src.Sheets[name];
    const ws = wb.addWorksheet(name || "Sheet1");
    if (!sws || !sws["!ref"]) return;
    const rows = X.utils.sheet_to_json(sws, { header: 1, raw: true, blankrows: true, defval: null });
    /* 뒤쪽 빈 줄은 버린다. .xls 는 쓰지도 않은 범위를 크게 잡아두는 일이 흔해서
       (코칸 양식은 65,523행으로 잡힌다) 그대로 옮기면 셀 수만 백만 단위로 불어난다. */
    let last = 0;
    rows.forEach((r, i) => {
      if (Array.isArray(r) && r.some(v => v !== null && v !== undefined && String(v).trim() !== "")) last = i + 1;
    });
    rows.slice(0, last).forEach(r => ws.addRow(Array.isArray(r) ? r : []));
    /* 열 너비는 사람이 열어봤을 때 표가 읽히느냐를 좌우한다 — 이건 옮겨준다.
       SheetJS 는 안 쓰는 열까지 257개를 채워 주므로 실제 쓰인 만큼만 본다. */
    const cols = sws["!cols"] || [];
    const used = ws.columnCount || 0;
    for (let i = 0; i < Math.min(cols.length, used); i++) {
      const w = cols[i] && (cols[i].wch || (cols[i].wpx ? cols[i].wpx / 7 : 0));
      if (w) ws.getColumn(i + 1).width = w;
    }
    (sws["!merges"] || []).forEach(m => {
      try { ws.mergeCells(m.s.r + 1, m.s.c + 1, m.e.r + 1, m.e.c + 1); } catch (e) {}
    });
  });
  if (!wb.worksheets.length) wb.addWorksheet("Sheet1");
  return wb;
}
async function loadWorkbook(dataOrBuffer) {
  if (isOldXlsBuffer(dataOrBuffer)) return xlsToWorkbook(dataOrBuffer);
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
  toDateValue, isDateHeader, hasDateFormat, hasTimeFormat,
  findDateColumns, defaultDateColumn, orderDateInfo, formatPhone, stripHyphen,
  valueTransformForHeader, nameFromFilename, normKey,
  mergeOrders, brandFromName, brandFromKnown, resolveBrand, aliasKey, convert, collectInvoices, looksLikeInvoice, looksLikeCarrier, carrierKey, countOrders, preview, previewAny, previewSheets, loadWorkbook, saveWorkbook, isOldXlsBuffer, todayStr, fmtDate,
  normPriceText, toPriceNumber, priceKeyParts, priceRowKey, buildPriceBook, matchPrice,
  rankPriceCandidates, settle, settleCheck,
  settleSheetHead, settleSheetRow, isPriceHeader, isInternalHeader, vendorSheetColumns, carryNote };
});
