/* 발주서 변환 — 주문 → 업체 양식
   2026-07-31 에 고친 날짜 처리(진짜 Date + 날짜만 + 셀 단위 서식)를 지킨다. */
"use strict";

const test = require("node:test");
const assert = require("node:assert");
const QO = require("../qo-logic.js");
const { makeWb, roundTrip, colByHeader, ymd } = require("./helpers.js");

/* 업체 양식(비어 있는 템플릿).
   수량·금액·날짜 열에 모두 같은 표시형식을 걸어 둔다 →
   저장하면 스타일이 '공유'되므로, 날짜 셀 하나를 고칠 때 수량·금액까지
   날짜로 바뀌어버리는 문제를 실제로 잡을 수 있다. */
function tplHeaders() {
  return ["주문수집일자", "주문일시", "상품명", "옵션", "수량", "금액", "수령인", "연락처", "주소", "우편번호"];
}
function makeTpl(opts) {
  opts = opts || {};
  // 전 열에 같은 서식(#,##0)을 걸어 스타일을 공유시킨다
  const numFmt = opts.numFmt || { 1: "#,##0", 2: "#,##0", 5: "#,##0", 6: "#,##0" };
  return makeWb([tplHeaders()], { numFmt, sheetName: "발주양식" });
}

/* 주문 파일(사방넷 형태) */
function makeOrder(rows) {
  const headers = ["브랜드", "주문번호", "주문수집일자", "주문일시", "상품명", "옵션",
                   "수량", "결제금액", "수령인명", "수령인연락처1", "배송주소", "우편번호"];
  return makeWb([headers, ...rows], { sheetName: "주문내역" });
}

const ROW1 = ["랩노마드", "20260728001", new Date(2026, 6, 28, 9, 15), new Date(2026, 6, 28, 14, 30),
              "비타민C 1000mg", "90정", 2, 25000, "홍길동", "01012345678", "서울시 강남구 테헤란로 1", "06234"];
const ROW2 = ["플라스머", "20260728002", new Date(2026, 6, 28, 10, 0), new Date(2026, 6, 28, 16, 45),
              "오메가3", "60캡슐", 1, 32000, "김철수", "01098765432", "부산시 해운대구 2", "48094"];

/* ---------------- 날짜 ---------------- */

test("날짜 열에는 진짜 Date 가 들어간다 (일련번호·문자열이 아니라)", async () => {
  const out = await QO.loadWorkbook(await QO.saveWorkbook(await run([ROW1])));
  const ws = out.worksheets[0];
  const c = ws.getRow(2).getCell(2);              // 주문일시
  assert.ok(c.value instanceof Date, `Date 여야 하는데 ${typeof c.value} 였다`);
});

test("발주서에 나가는 날짜는 시간이 없다 (자정으로 맞춘다)", async () => {
  const out = await run([ROW1]);
  const c = out.worksheets[0].getRow(2).getCell(2);
  assert.strictEqual(ymd(c.value), "2026-07-28");
  assert.strictEqual(c.value.getHours(), 0, "시가 남아 있다");
  assert.strictEqual(c.value.getMinutes(), 0, "분이 남아 있다");
  assert.strictEqual(c.value.getSeconds(), 0, "초가 남아 있다");
});

test("날짜 셀에 yyyy-mm-dd 서식이 걸린다 — 46231 같은 숫자로 보이면 안 된다", async () => {
  const out = await run([ROW1]);
  const c = out.worksheets[0].getRow(2).getCell(2);
  assert.strictEqual(QO.hasDateFormat(c.numFmt), true, `날짜 서식이 없다 (${c.numFmt})`);
  assert.strictEqual(QO.hasTimeFormat(c.numFmt), false, `시간 서식이 남아 있다 (${c.numFmt})`);
});

test("★ 날짜 서식이 수량·금액 칸으로 번지지 않는다 (스타일 공유 오염)", async () => {
  const out = await run([ROW1, ROW2]);
  const ws = out.worksheets[0];
  for (const row of [2, 3]) {
    const qty = ws.getRow(row).getCell(5);       // 수량
    const amt = ws.getRow(row).getCell(6);       // 금액
    assert.strictEqual(QO.hasDateFormat(qty.numFmt), false,
      `${row}행 수량이 날짜 서식으로 오염됐다 (${qty.numFmt})`);
    assert.strictEqual(QO.hasDateFormat(amt.numFmt), false,
      `${row}행 금액이 날짜 서식으로 오염됐다 (${amt.numFmt})`);
    assert.strictEqual(typeof qty.value, "number", "수량이 숫자가 아니다");
  }
});

test("업체 양식이 이미 '날짜만' 서식이면 그대로 존중한다", async () => {
  const tpl = await makeTpl({ numFmt: { 1: "yyyy/mm/dd", 2: "yyyy/mm/dd" } });
  const out = await run([ROW1], tpl);
  const c = out.worksheets[0].getRow(2).getCell(2);
  assert.strictEqual(c.numFmt, "yyyy/mm/dd", "업체 양식 서식을 바꿔버렸다");
});

test("업체 양식이 시간까지 보여주는 서식이면 날짜만으로 바꾼다", async () => {
  const tpl = await makeTpl({ numFmt: { 1: "yyyy-mm-dd hh:mm", 2: "yyyy-mm-dd hh:mm" } });
  const out = await run([ROW1], tpl);
  const c = out.worksheets[0].getRow(2).getCell(2);
  assert.strictEqual(QO.hasTimeFormat(c.numFmt), false, "00:00 이 붙어 보이게 된다");
});

test("주문 파일의 날짜가 일련번호(숫자)로 들어와도 날짜로 기입한다", async () => {
  // 날짜 서식이 안 걸린 채 숫자로 저장된 주문 파일 — 실제로 자주 있다
  const row = [...ROW1];
  row[3] = 46231;                                 // 주문일시 자리에 일련번호
  const out = await run([row]);
  const c = out.worksheets[0].getRow(2).getCell(2);
  assert.ok(c.value instanceof Date, "숫자가 그대로 남았다");
  assert.strictEqual(QO.hasDateFormat(c.numFmt), true);
});

test("주문 파일의 날짜가 문자열이어도 날짜로 기입한다", async () => {
  const row = [...ROW1];
  row[3] = "2026-07-28 14:30:00";
  const out = await run([row]);
  const c = out.worksheets[0].getRow(2).getCell(2);
  assert.ok(c.value instanceof Date);
  assert.strictEqual(ymd(c.value), "2026-07-28");
});

test("날짜 칸이 비어 있으면 엉뚱한 날짜를 만들지 않는다", async () => {
  const row = [...ROW1];
  row[3] = null;
  const out = await run([row]);
  const c = out.worksheets[0].getRow(2).getCell(2);
  assert.ok(!(c.value instanceof Date), "빈 칸에 날짜가 생겼다");
});

test("주문수집일자도 날짜로 기입된다", async () => {
  const out = await run([ROW1]);
  const c = out.worksheets[0].getRow(2).getCell(1);   // 주문수집일자
  assert.ok(c.value instanceof Date, "수집일자가 날짜가 아니다");
  assert.strictEqual(ymd(c.value), "2026-07-28");
});

/* ---------------- 값 후처리 ---------------- */

test("연락처는 하이픈을 넣고, 우편번호는 하이픈을 뗀다", async () => {
  const row = [...ROW1];
  row[11] = "062-34";
  const out = await run([row]);
  const ws = out.worksheets[0];
  assert.strictEqual(ws.getRow(2).getCell(8).value, "010-1234-5678");
  assert.strictEqual(ws.getRow(2).getCell(10).value, "06234");
});

/* ---------------- 행·필터 ---------------- */

test("주문 건수만큼 기입한다", async () => {
  const tpl = await makeTpl();
  const order = await makeOrder([ROW1, ROW2]);
  const r = QO.convert(order, tpl, {});
  assert.strictEqual(r.count, 2);
});

test("빈 행은 건너뛴다", async () => {
  const tpl = await makeTpl();
  const order = await makeOrder([ROW1, [null, null, null, null, null, null, null, null, null, null, null, null], ROW2]);
  const r = QO.convert(order, tpl, {});
  assert.strictEqual(r.count, 2, "빈 행까지 세었다");
});

test("브랜드로 거를 수 있다", async () => {
  const tpl = await makeTpl();
  const order = await makeOrder([ROW1, ROW2]);
  const r = QO.convert(order, tpl, { brands: ["랩노마드"] });
  assert.strictEqual(r.count, 1);
  assert.strictEqual(tpl.worksheets[0].getRow(2).getCell(7).value, "홍길동");
});

test("날짜로 거를 수 있다", async () => {
  const other = [...ROW2];
  other[2] = new Date(2026, 6, 29, 10, 0);
  other[3] = new Date(2026, 6, 29, 16, 45);
  const tpl = await makeTpl();
  const order = await makeOrder([ROW1, other]);
  const r = QO.convert(order, tpl, { dates: ["20260728"], dateHeader: "주문일시" });
  assert.strictEqual(r.count, 1);
});

test("공통 항목이 하나도 없으면 알아듣게 실패한다", async () => {
  const tpl = await makeWb([["가", "나", "다"]]);
  const order = await makeOrder([ROW1]);
  await assert.rejects(async () => QO.convert(order, tpl, {}), /매칭|찾지 못/);
});

/* ---------------- 결제일시 대체 ---------------- */

test("주문에 결제일시가 없고 양식에만 있으면 주문일시로 채운다", async () => {
  const tpl = await makeWb([["결제일시", "상품명", "수량", "수령인", "주소"]], { sheetName: "발주양식" });
  const order = await makeOrder([ROW1]);
  QO.convert(order, tpl, {});
  const c = tpl.worksheets[0].getRow(2).getCell(1);
  assert.ok(c.value instanceof Date, "결제일시가 채워지지 않았다");
  assert.strictEqual(ymd(c.value), "2026-07-28");
});

/* 주문 rows 로 변환을 돌리고 결과 워크북을 돌려준다 */
async function run(rows, tpl) {
  const t = tpl || await makeTpl();
  const order = await makeOrder(rows);
  QO.convert(order, t, {});
  return t;
}
