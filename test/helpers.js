/* 테스트용 엑셀 워크북 만들기 도우미
   실제 앱과 같은 경로를 타도록, 만든 워크북은 항상 buffer 로 저장했다가 다시 읽는다.
   (업체 양식의 '스타일 공유'는 파일로 저장될 때 생기므로, 이 왕복이 있어야
    수량·금액 칸이 날짜로 오염되는 문제를 실제로 잡을 수 있다) */
"use strict";

const QO = require("../qo-logic.js");

/* rows[0] 을 헤더로 보고 시트 하나짜리 워크북을 만든다.
   opts.numFmt = {열번호(1-based): "표시형식"} — 업체 양식의 기존 서식을 흉내낼 때 사용 */
async function makeWb(rows, opts) {
  opts = opts || {};
  const ExcelJS = require("exceljs");
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(opts.sheetName || "Sheet1");
  rows.forEach(r => ws.addRow(r));

  // 열 전체에 같은 표시형식을 건다 → 저장하면 스타일이 공유된 형태가 된다
  for (const [col, fmt] of Object.entries(opts.numFmt || {})) {
    const c = ws.getColumn(Number(col));
    c.numFmt = fmt;
    c.eachCell({ includeEmpty: true }, cell => { cell.numFmt = fmt; });
  }
  return roundTrip(wb);
}

/* 파일로 저장했다가 다시 읽기 — 앱이 업로드된 파일을 여는 것과 같은 상태로 만든다 */
async function roundTrip(wb) {
  const buf = await QO.saveWorkbook(wb);
  return QO.loadWorkbook(buf);
}

/* 시트의 셀 값 읽기 (1-based) */
function cellOf(wb, row, col, sheet) {
  const ws = sheet ? wb.getWorksheet(sheet) : wb.worksheets[0];
  return ws.getRow(row).getCell(col);
}

/* 헤더 이름으로 열 번호 찾기 (1-based, 없으면 null) */
function colByHeader(wb, header, headerRow, sheet) {
  const ws = sheet ? wb.getWorksheet(sheet) : wb.worksheets[0];
  const hr = headerRow || 1;
  const d = QO.dims(ws);
  for (let c = 1; c <= d.cols; c++) {
    const v = QO.getV(ws, hr, c);
    if (String(v == null ? "" : v).replace(/\s/g, "") === header.replace(/\s/g, "")) return c;
  }
  return null;
}

/* 엑셀 일련번호 → 사람이 읽는 값 확인용: 날짜를 yyyy-mm-dd 로 */
function ymd(d) {
  if (!(d instanceof Date)) return String(d);
  const p = n => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

module.exports = { makeWb, roundTrip, cellOf, colByHeader, ymd };
