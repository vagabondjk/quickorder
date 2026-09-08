/* 정산·CS 의 날짜 정규화 (qo-cs.js 의 toYmd)
   엑셀 일련번호가 "46231-01-01" 같은 엉뚱한 날짜로 나오던 문제를 막는다.

   qo-cs.js 는 브라우저 전용(전역 변수를 그대로 쓴다)이라 통째로 require 할 수 없어서,
   toYmd 부분만 파일에서 떼어내 QO 를 넣고 돌린다.
   → 원본이 바뀌면 이 테스트도 같이 깨지므로 복사본이 낡을 걱정이 없다. */
"use strict";

const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");
const QO = require("../qo-logic.js");

/* qo-cs.js 에서 날짜 정규화 부분만 잘라내 실행한다 */
const toYmd = (() => {
  const src = fs.readFileSync(path.join(__dirname, "..", "qo-cs.js"), "utf8");
  const from = src.indexOf("const p2 =");
  const to = src.indexOf("function today()");
  if (from < 0 || to < 0) throw new Error("qo-cs.js 에서 날짜 정규화 부분을 찾지 못했습니다.");
  const body = src.slice(from, to);
  const s = v => (v === null || v === undefined) ? "" : String(v).trim();
  // eslint-disable-next-line no-new-func
  return new Function("QO", "s", body + "; return toYmd;")(QO, s);
})();

test("★ 엑셀 일련번호를 46231년으로 읽지 않는다", () => {
  // 이 버그의 정체: new Date("46231") 이 46231년으로 파싱된다
  assert.strictEqual(toYmd(46231), "2026-07-28");
  assert.notStrictEqual(toYmd(46231), "46231-01-01");
});

test("일련번호에 시간이 붙어 있어도 날짜만 뽑는다", () => {
  assert.strictEqual(toYmd(46231.6), "2026-07-28");
});

test("숫자가 글자로 들어와도(46231) 날짜로 읽는다", () => {
  assert.strictEqual(toYmd("46231"), "2026-07-28");
});

test("원래 되던 형태는 그대로 된다", () => {
  assert.strictEqual(toYmd(new Date(2026, 6, 28)), "2026-07-28");
  assert.strictEqual(toYmd("2026-07-28"), "2026-07-28");
  assert.strictEqual(toYmd("2026.7.28"), "2026-07-28");
  assert.strictEqual(toYmd("2026년 7월 28일"), "2026-07-28");
  assert.strictEqual(toYmd("20260728"), "2026-07-28");
  assert.strictEqual(toYmd("2026-07-28 14:30:00"), "2026-07-28");
});

test("빈 값은 빈 문자열", () => {
  for (const v of [null, undefined, "", "   "]) assert.strictEqual(toYmd(v), "");
});

test("말이 안 되는 연도로 날짜를 지어내지 않는다", () => {
  // new Date("46231-01-01") 은 46231년으로 파싱된다.
  // 그걸 날짜라고 받아들이면 안 된다 — 읽지 못한 값은 원문 그대로 둔다.
  const out = toYmd("46231-01-01");
  assert.ok(!/^\d{4}-\d{2}-\d{2}$/.test(out), `엉뚱한 연도가 날짜로 통과했다: ${out}`);
});

/* 지금 이렇게 동작한다는 기록.
   주문번호(20260728001)처럼 앞 8자리가 날짜인 긴 숫자는 날짜로 읽힌다.
   주문번호 앞자리가 실제 주문일자라 대개 맞는 값이 나오지만,
   정산 파일에서 '정산일' 열이 주문번호 열로 잘못 잡히면 그럴듯한 날짜가
   조용히 들어간다. 고칠 거면 의도적으로 고쳐야 해서 여기 남겨둔다. */
test("[현재 동작] 주문번호처럼 긴 숫자도 앞 8자리를 날짜로 읽는다", () => {
  assert.strictEqual(toYmd(20260728001), "2026-07-28");
});

/* ---------------- 엑셀에 날짜로 기입되는지 ---------------- */

test("정산서 엑셀의 정산일은 진짜 Date 로 들어간다", async () => {
  // qo-settle.js 의 dv() 와 같은 변환
  const dv = v => QO.toDateValue(String(v == null ? "" : v).trim()) || String(v == null ? "" : v).trim();
  const d = dv("2026-07-28");
  assert.ok(d instanceof Date, "글자 그대로 들어가면 엑셀에서 정렬·필터가 안 먹는다");
  assert.strictEqual(d.getFullYear(), 2026);
  assert.strictEqual(d.getMonth(), 6);
  assert.strictEqual(d.getDate(), 28);
});

test("날짜가 아닌 값은 원래대로 둔다 (빈칸·미정 등)", () => {
  const dv = v => QO.toDateValue(String(v == null ? "" : v).trim()) || String(v == null ? "" : v).trim();
  assert.strictEqual(dv(""), "");
  assert.strictEqual(dv("미정"), "미정");
});
