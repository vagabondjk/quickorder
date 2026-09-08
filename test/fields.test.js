/* 헤더 → 표준 항목 매핑
   특히 '주문수집일자' 열을 '주문일시' 가 채가지 않는지(2026-07-31 수정) 확인한다. */
"use strict";

const test = require("node:test");
const assert = require("node:assert");
const QO = require("../qo-logic.js");
const { makeWb, colByHeader } = require("./helpers.js");

/* 헤더 한 줄짜리 시트를 만들어 표준 항목 매핑을 구한다 */
async function mapOf(headers, role) {
  const wb = await makeWb([headers, headers.map(() => "값")]);
  const ws = wb.worksheets[0];
  const hr = QO.findHeaderRow(ws);
  return { map: QO.buildOrderFieldMap(ws, hr, role || "target"), wb, hr };
}

test("주문수집일자와 주문일시가 같이 있으면 각각 제 열을 잡는다", async () => {
  const headers = ["주문번호", "주문수집일자", "주문일시", "상품명", "수량", "수령인명", "배송주소"];
  const { map } = await mapOf(headers);
  // 1-based 열 번호
  assert.strictEqual(map.COLLECT_DATE, 2, "주문수집일자가 2번 열이어야 한다");
  assert.strictEqual(map.ORDER_DATE, 3, "주문일시가 3번 열이어야 한다");
});

test("주문수집일자만 있으면 주문일시로 잡히지 않는다", async () => {
  // '주문수집일자' 에는 '주문일' 이 들어있어서, 순서가 틀리면 ORDER_DATE 가 채간다.
  const headers = ["주문번호", "주문수집일자", "상품명", "수량", "수령인명", "배송주소"];
  const { map } = await mapOf(headers);
  assert.strictEqual(map.COLLECT_DATE, 2);
  assert.strictEqual(map.ORDER_DATE, undefined, "주문일시가 수집일자 열을 채가면 안 된다");
});

test("주문일시만 있으면 수집일자는 비어 있다", async () => {
  const headers = ["주문번호", "주문일시", "상품명", "수량", "수령인명", "배송주소"];
  const { map } = await mapOf(headers);
  assert.strictEqual(map.ORDER_DATE, 2);
  assert.strictEqual(map.COLLECT_DATE, undefined);
});

test("한 열을 두 항목이 같이 쓰지 않는다", async () => {
  const headers = ["주문번호", "주문수집일자", "주문일시", "결제일시", "상품명", "옵션",
                   "수량", "결제금액", "수령인명", "수령인연락처1", "배송주소", "우편번호"];
  const { map } = await mapOf(headers);
  const cols = Object.values(map);
  assert.strictEqual(cols.length, new Set(cols).size, "같은 열이 두 번 쓰였다");
});

test("예약일·출하지시일 같은 열은 주문일시로 잡지 않는다", async () => {
  const headers = ["주문번호", "출하지시일자", "배송예정일", "상품명", "수량", "수령인명", "배송주소"];
  const { map } = await mapOf(headers);
  assert.strictEqual(map.ORDER_DATE, undefined, "지시·예정 이 붙은 열은 제외되어야 한다");
});

test("수령인 연락처 1·2 를 순서대로 잡는다", async () => {
  const headers = ["상품명", "수량", "수령인명", "수령인연락처1", "수령인연락처2", "배송주소"];
  const { map } = await mapOf(headers, "target");
  assert.ok(map.RECIPIENT_PHONE < map.RECIPIENT_PHONE2, "연락처1 이 연락처2 보다 앞이어야 한다");
});

test("가상번호·이메일 열은 연락처로 잡지 않는다", async () => {
  const headers = ["상품명", "수량", "수령인명", "가상전화번호", "이메일", "배송주소"];
  const { map } = await mapOf(headers, "target");
  assert.strictEqual(map.RECIPIENT_PHONE, undefined, "가상번호를 연락처로 쓰면 안 된다");
});

/* ---------------- 값 후처리 ---------------- */

test("우편번호는 하이픈을 뗀다", () => {
  const tf = QO.valueTransformForHeader("우편번호");
  assert.strictEqual(tf("123-456"), "123456");
});

test("연락처는 하이픈을 넣어 정리한다", () => {
  const tf = QO.valueTransformForHeader("수령인연락처1");
  assert.strictEqual(tf("01012345678"), "010-1234-5678");
  assert.strictEqual(tf("1012345678"), "010-1234-5678", "앞의 0 이 빠진 값도 복구해야 한다");
  assert.strictEqual(tf("0212345678"), "02-1234-5678");
});

test("가상번호 열은 손대지 않는다", () => {
  assert.strictEqual(QO.valueTransformForHeader("가상연락처"), null);
});

/* ---------------- 파일명 → 업체명 ---------------- */

test("파일명에서 업체명을 뽑는다", () => {
  assert.strictEqual(QO.nameFromFilename("디에스피_발주양식.xlsx"), "디에스피_발주양식");
  assert.strictEqual(QO.nameFromFilename("1. 디에스피 회신.xlsx"), "디에스피");
  assert.strictEqual(QO.nameFromFilename("/경로/있는/플라스머 송장.xlsx"), "플라스머");
});
