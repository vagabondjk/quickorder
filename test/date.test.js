/* 날짜 처리 — toDateValue / 헤더 판별 / 서식 판별
   업체 양식에 날짜가 46231.6 같은 일련번호로 나가던 문제를 막는 부분이다. */
"use strict";

const test = require("node:test");
const assert = require("node:assert");
const QO = require("../qo-logic.js");
const { ymd } = require("./helpers.js");

/* ---------------- toDateValue ---------------- */

test("toDateValue: Date 는 그대로 돌려준다", () => {
  const d = new Date(2026, 6, 28, 14, 30);
  assert.strictEqual(QO.toDateValue(d), d);
});

test("toDateValue: 잘못된 Date 는 null", () => {
  assert.strictEqual(QO.toDateValue(new Date("깨진값")), null);
});

test("toDateValue: 엑셀 일련번호 → 날짜 (45000 = 2023-03-15)", () => {
  // 엑셀 기준일(1899-12-30)에서 45000일. 1900 윤년 버그까지 포함한 값이다.
  const d = QO.toDateValue(45000);
  assert.strictEqual(ymd(d), "2023-03-15");
});

test("toDateValue: 일련번호는 로컬 시간으로 만든다 — 시간대 때문에 하루가 밀리면 안 된다", () => {
  // UTC 로 만들면 한국(UTC+9)에서 날짜가 하루 앞당겨져 보이는 문제를 막는다.
  const d = QO.toDateValue(45000);
  assert.strictEqual(d.getDate(), 15, "날짜가 밀렸다");
  assert.strictEqual(d.getHours(), 0);
  assert.strictEqual(d.getMinutes(), 0);
});

test("toDateValue: 소수부는 시:분:초로 푼다 (45000.5 = 정오)", () => {
  const d = QO.toDateValue(45000.5);
  assert.strictEqual(ymd(d), "2023-03-15");
  assert.strictEqual(d.getHours(), 12);
  assert.strictEqual(d.getMinutes(), 0);
});

test("toDateValue: 날짜로 볼 수 없는 숫자는 null", () => {
  assert.strictEqual(QO.toDateValue(0), null);       // 범위 밖
  assert.strictEqual(QO.toDateValue(-5), null);
  assert.strictEqual(QO.toDateValue(60001), null);   // 너무 큼 → 수량·금액을 날짜로 오인하지 않게
  assert.strictEqual(QO.toDateValue(Infinity), null);
});

test("toDateValue: 수량·금액 같은 작은 숫자를 날짜로 바꿔버리면 안 된다", () => {
  // 1~60000 범위는 통과하지만, 애초에 '날짜 열'로 판별된 열에서만 호출되므로
  // 여기서는 범위 경계만 확인한다.
  assert.notStrictEqual(QO.toDateValue(1), null);
  assert.strictEqual(QO.toDateValue(60001), null);
});

test("toDateValue: 문자열 날짜 여러 형태를 읽는다", () => {
  assert.strictEqual(ymd(QO.toDateValue("2026-07-28")), "2026-07-28");
  assert.strictEqual(ymd(QO.toDateValue("2026.07.28")), "2026-07-28");
  assert.strictEqual(ymd(QO.toDateValue("2026/07/28")), "2026-07-28");
  assert.strictEqual(ymd(QO.toDateValue("20260728")), "2026-07-28");
});

test("toDateValue: 문자열에 시간이 붙어 있으면 시간까지 읽는다", () => {
  const d = QO.toDateValue("2026-07-28 14:30:05");
  assert.strictEqual(ymd(d), "2026-07-28");
  assert.strictEqual(d.getHours(), 14);
  assert.strictEqual(d.getMinutes(), 30);
  assert.strictEqual(d.getSeconds(), 5);
});

test("toDateValue: 날짜가 아닌 값은 null", () => {
  for (const v of [null, undefined, "", "   ", "abc", "2026", "202607"]) {
    assert.strictEqual(QO.toDateValue(v), null, `${JSON.stringify(v)} → null 이어야 한다`);
  }
});

test("toDateValue: 말이 안 되는 날짜는 null", () => {
  assert.strictEqual(QO.toDateValue("20261328"), null);  // 13월
  assert.strictEqual(QO.toDateValue("20260732"), null);  // 32일
  assert.strictEqual(QO.toDateValue("18990728"), null);  // 범위 밖 연도
});

/* ---------------- 헤더가 날짜 열인가 ---------------- */

test("isDateHeader: 날짜 열 헤더를 알아본다", () => {
  for (const h of ["주문일시", "주문일자", "결제일시", "주문수집일자", "수집일", "발주일자"]) {
    assert.strictEqual(QO.isDateHeader(h), true, `${h} 는 날짜 열이어야 한다`);
  }
});

test("isDateHeader: 날짜가 아닌 열은 걸러낸다", () => {
  for (const h of ["상품명", "수량", "결제금액", "수령인명", "배송주소", "우편번호", "운송장번호"]) {
    assert.strictEqual(QO.isDateHeader(h), false, `${h} 를 날짜 열로 보면 안 된다`);
  }
  assert.strictEqual(QO.isDateHeader(null), false);
  assert.strictEqual(QO.isDateHeader(123), false);
});

/* ---------------- 표시형식 판별 ---------------- */

test("hasDateFormat: 날짜 서식만 참", () => {
  assert.strictEqual(QO.hasDateFormat("yyyy-mm-dd"), true);
  assert.strictEqual(QO.hasDateFormat("yyyy-mm-dd hh:mm"), true);
  assert.strictEqual(QO.hasDateFormat("m/d/yy"), true);
});

test("hasDateFormat: 일반·숫자·문자 서식은 거짓 — 수량·금액을 날짜로 오인하지 않게", () => {
  assert.strictEqual(QO.hasDateFormat("General"), false);
  assert.strictEqual(QO.hasDateFormat("#,##0"), false);
  assert.strictEqual(QO.hasDateFormat("@"), false);
  assert.strictEqual(QO.hasDateFormat("0.00"), false);
  assert.strictEqual(QO.hasDateFormat(undefined), false);
});

test("hasDateFormat: 따옴표 안의 글자·대괄호는 무시한다", () => {
  // "원"·"개" 같은 단위 글자가 서식에 있어도 날짜로 오인하면 안 된다
  assert.strictEqual(QO.hasDateFormat('#,##0"원"'), false);
  assert.strictEqual(QO.hasDateFormat('0"days"'), false);
  assert.strictEqual(QO.hasDateFormat("[Red]#,##0"), false);
});

test("hasTimeFormat: 시간이 들어간 서식만 참", () => {
  assert.strictEqual(QO.hasTimeFormat("yyyy-mm-dd hh:mm"), true);
  assert.strictEqual(QO.hasTimeFormat("yyyy-mm-dd"), false);
  assert.strictEqual(QO.hasTimeFormat("yyyy/mm/dd"), false);
  assert.strictEqual(QO.hasTimeFormat("General"), false);
});

/* ---------------- extractDate (날짜 칩 만들 때 쓰는 yyyymmdd) ---------------- */

test("extractDate: 여러 형태에서 yyyymmdd 를 뽑는다", () => {
  assert.strictEqual(QO.extractDate(new Date(2026, 6, 28)), "20260728");
  assert.strictEqual(QO.extractDate("2026-07-28 14:30"), "20260728");
  assert.strictEqual(QO.extractDate("2026.07.28"), "20260728");
  assert.strictEqual(QO.extractDate(null), null);
  assert.strictEqual(QO.extractDate("없음"), null);
});
