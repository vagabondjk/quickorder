/* 송장 취합 — 업체 회신 → 송장취합양식
   2026-07-31 에 넣은 송장번호 정리(하이픈 제거)가 값을 망가뜨리지 않는지 확인한다. */
"use strict";

const test = require("node:test");
const assert = require("node:assert");
const QO = require("../qo-logic.js");
const { makeWb } = require("./helpers.js");

const SABANG_H = ["주문번호", "상품명", "수량", "수령인명", "배송주소", "택배사", "운송장번호"];
const REPLY_H  = ["상품명", "수량", "수령인명", "배송주소", "택배사", "운송장번호"];

/* 취합양식: 송장 칸이 비어 있는 주문 목록 */
function sabang(rows) {
  return makeWb([SABANG_H, ...rows.map(r => [...r, null, null])], { sheetName: "주문목록" });
}
/* 업체 회신 */
function reply(rows, name) {
  return makeWb([REPLY_H, ...rows], { sheetName: "회신" })
    .then(wb => ({ name: name || "디에스피_회신.xlsx", wb }));
}

const O1 = ["20260728001", "비타민C 1000mg", 2, "홍길동", "서울시 강남구 테헤란로 1"];
const O2 = ["20260728002", "오메가3", 1, "김철수", "부산시 해운대구 2"];

/* 회신 한 건을 취합하고, 취합양식의 송장 칸 값을 돌려준다 */
async function collectOne(invValue, order) {
  const o = order || O1;
  const target = await sabang([o]);
  const rep = await reply([[o[1], o[2], o[3], o[4], "CJ대한통운", invValue]]);
  const r = QO.collectInvoices(target, [rep], {});
  const ws = target.worksheets[0];
  return { value: ws.getRow(2).getCell(7).value, carrier: ws.getRow(2).getCell(6).value, result: r };
}

/* ---------------- 송장번호 정리 ---------------- */

test("송장번호의 하이픈을 뗀다", async () => {
  const { value } = await collectOne("1234-5678");
  assert.strictEqual(value, "12345678");
});

test("송장번호의 공백·점도 뗀다", async () => {
  assert.strictEqual((await collectOne("1234 5678")).value, "12345678");
  assert.strictEqual((await collectOne("1234.5678")).value, "12345678");
});

test("★ 앞자리 0 이 사라지지 않는다", async () => {
  const { value } = await collectOne("0123-4567");
  assert.strictEqual(value, "01234567", "앞의 0 이 날아갔다");
  assert.strictEqual(typeof value, "string", "숫자로 바뀌면 앞의 0 이 사라진다");
});

test("영문이 섞인 송장(해외 EMS 등)은 글자를 살린다", async () => {
  const { value } = await collectOne("EE123-456-789KR");
  assert.strictEqual(value, "EE123456789KR");
});

test("뗄 게 없는 송장번호는 원래 값 그대로 둔다 (형식 안 바뀜)", async () => {
  const { value } = await collectOne(123456789012);
  assert.strictEqual(value, 123456789012);
  assert.strictEqual(typeof value, "number", "숫자였는데 문자로 바뀌었다");
});

test("숫자가 하나도 없는 값은 손대지 않는다", async () => {
  const { value } = await collectOne("미발송");
  assert.strictEqual(value, "미발송");
});

test("택배사도 같이 채워진다", async () => {
  const { carrier } = await collectOne("1234-5678");
  assert.strictEqual(carrier, "CJ대한통운");
});

/* ---------------- 매칭 ---------------- */

test("수령인·주소·상품으로 맞는 행에 채운다", async () => {
  const target = await sabang([O1, O2]);
  const rep = await reply([[O2[1], O2[2], O2[3], O2[4], "롯데택배", "999-888"]]);
  QO.collectInvoices(target, [rep], {});
  const ws = target.worksheets[0];
  assert.ok(QO.isBlank(ws.getRow(2).getCell(7).value), "엉뚱한 행(홍길동)에 채웠다");
  assert.strictEqual(ws.getRow(3).getCell(7).value, "999888");
});

test("취합 건수를 보고한다", async () => {
  const target = await sabang([O1, O2]);
  const rep = await reply([
    [O1[1], O1[2], O1[3], O1[4], "CJ대한통운", "111-111"],
    [O2[1], O2[2], O2[3], O2[4], "롯데택배", "222-222"],
  ]);
  const r = QO.collectInvoices(target, [rep], {});
  assert.strictEqual(r.total, 2);
});

test("회신에 없는 주문은 '누락'으로 잡아낸다", async () => {
  const target = await sabang([O1, O2]);
  const rep = await reply([[O1[1], O1[2], O1[3], O1[4], "CJ대한통운", "111-111"]]);
  const r = QO.collectInvoices(target, [rep], {});
  assert.ok(Array.isArray(r.missing), "missing 이 없다");
  assert.strictEqual(r.missing.length, 1, "누락 1건이 잡혀야 한다");
});

test("송장 열이 없는 회신은 오류로 남기고 나머지를 계속 처리한다", async () => {
  const target = await sabang([O1]);
  const bad = { name: "빈회신.xlsx", wb: await makeWb([["상품명", "수량", "수령인명"], ["비타민C 1000mg", 2, "홍길동"]]) };
  const good = await reply([[O1[1], O1[2], O1[3], O1[4], "CJ대한통운", "111-111"]]);
  const r = QO.collectInvoices(target, [bad, good], {});
  assert.strictEqual(r.total, 1, "정상 회신까지 막히면 안 된다");
});

test("대상 시트를 못 찾으면 알아듣게 실패한다", async () => {
  const target = await makeWb([["가", "나", "다"], [1, 2, 3]]);
  const rep = await reply([[O1[1], O1[2], O1[3], O1[4], "CJ대한통운", "111"]]);
  assert.throws(() => QO.collectInvoices(target, [rep], {}), /찾지 못/);
});
