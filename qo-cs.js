/* =====================================================================
   퀵오더 ③ CS · 교환/반품  (v6.0)
   - 사방넷·쇼핑몰 어드민 엑셀을 '열 맞추기' 한 번으로 가져온다(레이아웃별 기억).
   - 메일 문의(본문)·직접 등록도 같은 목록에 합친다.
   - 브랜드→업체 학습값(S.brandVendor)으로 담당 업체를 자동 배정.
   - 업체별로 묶어 메일 전달 + 상태 관리 + 엑셀 저장.
   qo-app.js 뒤에 로드되므로 $ / DB / S / esc / msg / download / readFile /
   bindDrop / QO / GMAIL / fillRecipients / parseEmails / cleanVendor 를 그대로 쓴다.
   ===================================================================== */
"use strict";

/* =====================================================================
   공용: 열 맞추기(매핑)  — CS·정산이 같이 쓴다
   ===================================================================== */
const MAP = (() => {
  const norm = s => String(s == null ? "" : s).replace(/[\s_\-().\[\]]/g, "").toLowerCase();
  let cur = null;

  /* 컬럼 목록에서 이 필드에 해당할 만한 열 번호(0-based) 추측 */
  function guess(columns, field) {
    const cols = columns.map(norm);
    for (const kw of field.kw) {                       // ① 완전히 같은 이름
      const i = cols.indexOf(norm(kw));
      if (i >= 0) return i;
    }
    for (const kw of field.kw) {                       // ② 포함
      const k = norm(kw);
      if (!k) continue;
      const i = cols.findIndex(c => c && c.includes(k));
      if (i >= 0) return i;
    }
    return -1;
  }
  /* 파일 레이아웃 서명 — 열 이름이 같으면 같은 양식으로 보고 매핑을 재사용 */
  function signature(columns) {
    return columns.map(c => String(c || "").trim()).filter(Boolean).slice(0, 40).join("|").slice(0, 400);
  }
  function autoMap(columns, fields) {
    const m = {};
    fields.forEach(f => { const i = guess(columns, f); if (i >= 0) m[f.k] = i; });
    return m;
  }
  function ok(m, fields) {
    const need = fields.filter(f => f.req).map(f => f.k);
    return need.every(k => m[k] !== undefined && m[k] >= 0);
  }

  /* 모달 열기. onOk(map) */
  function open({ title, sub, columns, fields, saved, onOk }) {
    cur = { columns, fields, onOk };
    $("map-title").textContent = title || "열 맞추기";
    $("map-sub").textContent = sub || "이 파일의 어느 열이 무슨 값인지 한 번만 정해주세요. 다음부터는 자동입니다.";
    $("map-msg").textContent = "";
    const base = Object.assign(autoMap(columns, fields), saved || {});
    const box = $("map-list");
    box.innerHTML = "";
    fields.forEach(f => {
      const row = document.createElement("div");
      row.className = "maprow";
      const opts = ['<option value="-1">— 없음 —</option>']
        .concat(columns.map((c, i) => `<option value="${i}">${esc(c || `(${i + 1}번째 열)`)}</option>`));
      row.innerHTML = `<b>${esc(f.n)}${f.req ? ' <span style="color:var(--danger)">*</span>' : ""}</b>
        <select data-k="${f.k}">${opts.join("")}</select>`;
      row.querySelector("select").value = String(base[f.k] === undefined ? -1 : base[f.k]);
      box.appendChild(row);
    });
    $("mapmodal").classList.add("on");
  }
  function close() { $("mapmodal").classList.remove("on"); cur = null; }

  $("map-cancel").onclick = close;
  $("mapmodal").onclick = e => { if (e.target === $("mapmodal")) close(); };
  $("map-ok").onclick = () => {
    if (!cur) return close();
    const m = {};
    $("map-list").querySelectorAll("select").forEach(s => {
      const v = Number(s.value);
      if (v >= 0) m[s.dataset.k] = v;
    });
    if (!ok(m, cur.fields)) {
      const miss = cur.fields.filter(f => f.req && m[f.k] === undefined).map(f => f.n);
      $("map-msg").textContent = "⚠ 꼭 필요한 항목을 지정해주세요: " + miss.join(", ");
      return;
    }
    const fn = cur.onOk;
    close();
    fn(m);
  };

  return { open, guess, signature, autoMap, ok };
})();

/* =====================================================================
   ③ CS 모듈
   ===================================================================== */
const CS = (() => {
  const TYPES = ["문의", "교환", "반품", "취소", "배송", "기타"];
  const STATUS = ["접수", "업체확인", "처리중", "완료", "보류"];
  const OPEN_ST = ["접수", "업체확인", "처리중"];          // '미처리'로 세는 상태
  const FAULTS = ["", "고객", "업체", "배송", "우리"];

  const FIELDS = [
    { k: "date", n: "접수일", kw: ["접수일", "등록일", "작성일", "문의일", "신청일", "일자", "날짜", "일시", "주문일"] },
    { k: "mall", n: "쇼핑몰", kw: ["쇼핑몰", "판매처", "마켓", "채널", "사이트", "매체", "몰"] },
    { k: "orderNo", n: "주문번호", kw: ["주문번호", "오더번호", "주문no", "order"], req: true },
    { k: "customer", n: "고객명", kw: ["고객명", "구매자", "주문자", "수취인", "수령인", "고객", "성명", "이름"] },
    { k: "phone", n: "연락처", kw: ["연락처", "휴대폰", "핸드폰", "전화번호", "전화", "수취인연락처"] },
    { k: "product", n: "상품명", kw: ["상품명", "제품명", "품목명", "상품", "제품", "모델"] },
    { k: "option", n: "옵션", kw: ["옵션", "옵션명", "선택옵션", "규격"] },
    { k: "qty", n: "수량", kw: ["수량", "개수", "주문수량"] },
    { k: "type", n: "유형", kw: ["유형", "구분", "분류", "처리구분", "클레임", "종류", "요청사항"] },
    { k: "title", n: "제목", kw: ["제목", "문의제목", "title"] },
    { k: "content", n: "내용", kw: ["내용", "문의내용", "사유", "요청내용", "상세", "비고", "메모"] },
    { k: "status", n: "상태", kw: ["상태", "처리상태", "진행상태", "진행", "처리여부"] },
  ];

  let items = [];             // CS 전체
  let maps = {};              // 레이아웃 서명 → 매핑
  let filter = { st: "미처리", type: "전체", q: "" };
  let checked = new Set();
  let editing = null;         // 상세 모달에서 편집 중인 항목
  let drawn = false;

  const uid = () => "cs" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  const s = v => (v === null || v === undefined) ? "" : String(v).trim();
  const num = v => { const n = Number(String(v == null ? "" : v).replace(/[^0-9.\-]/g, "")); return isFinite(n) ? n : 0; };

  /* ---------- 날짜 정규화 → YYYY-MM-DD ---------- */
  const p2 = n => String(n).padStart(2, "0");
  const fromDate = d => `${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())}`;
  function toYmd(v) {
    if (v instanceof Date) return isNaN(v.getTime()) ? "" : fromDate(v);
    const t = s(v);
    if (!t) return "";
    // 2026-07-26 / 2026.7.26 / 20260726 — 달·일 범위까지 확인해야
    // "Sun Jul 26 2026 14:07" 같은 문자열에서 '2026-14-07'로 잘못 읽지 않는다.
    const m = t.match(/(20\d{2})[.\-/년\s]*(\d{1,2})[.\-/월\s]*(\d{1,2})/);
    if (m && +m[2] >= 1 && +m[2] <= 12 && +m[3] >= 1 && +m[3] <= 31)
      return `${m[1]}-${p2(m[2])}-${p2(m[3])}`;
    const d = new Date(t);
    if (!isNaN(d.getTime())) return fromDate(d);
    return t.slice(0, 10);
  }
  function today() { return toYmd(new Date()); }

  /* ---------- 유형 추정 ---------- */
  function guessType(text) {
    const t = s(text);
    if (/교환|바꿔|다른\s*색|다른\s*사이즈/.test(t)) return "교환";
    if (/반품|환불|반송|회수/.test(t)) return "반품";
    if (/취소/.test(t)) return "취소";
    if (/배송|택배|송장|미도착|누락|파손|분실/.test(t)) return "배송";
    if (/문의|질문|문의드/.test(t)) return "문의";
    return "";
  }
  function normType(v, fallbackText) {
    const t = s(v);
    for (const x of TYPES) if (t.includes(x)) return x;
    return guessType(t + " " + s(fallbackText)) || "문의";
  }
  function normStatus(v) {
    const t = s(v);
    for (const x of STATUS) if (t.includes(x)) return x;
    if (/완료|처리됨|종결|답변완료/.test(t)) return "완료";
    if (/대기|미처리|접수/.test(t)) return "접수";
    return "접수";
  }

  /* ---------- 브랜드 → 업체 자동 배정 ---------- */
  function findVendor(product, option) {
    const hay = (s(product) + " " + s(option)).toLowerCase();
    const bv = S.brandVendor || {};
    let best = "", bestLen = 0;
    for (const b in bv) {
      const k = String(b).toLowerCase().trim();
      if (k && hay.includes(k) && k.length > bestLen) { best = bv[b]; bestLen = k.length; }
    }
    return best;
  }
  function findBrand(product, option) {
    const hay = (s(product) + " " + s(option)).toLowerCase();
    let best = "", bestLen = 0;
    for (const b in (S.brandVendor || {})) {
      const k = String(b).toLowerCase().trim();
      if (k && hay.includes(k) && k.length > bestLen) { best = b; bestLen = k.length; }
    }
    return best;
  }

  /* ---------- 저장/불러오기 ---------- */
  async function load() {
    items = await DB.get("csItems", []) || [];
    maps = await DB.get("csMaps", {}) || {};
    if (!Array.isArray(items)) items = [];
  }
  async function save() {
    items.sort((a, b) => (b.date || "").localeCompare(a.date || "") || (b.createdAt || 0) - (a.createdAt || 0));
    await DB.set("csItems", items);
    badge();
  }
  const openCount = () => items.filter(x => OPEN_ST.includes(x.status)).length;
  function badge() {
    const b = $("cs-badge");
    if (!b) return;
    const n = openCount();
    b.textContent = n;
    b.style.display = n ? "" : "none";
  }

  /* ---------- 중복 제거 키 ---------- */
  const dupKey = x => [s(x.mall), s(x.orderNo), s(x.type), s(x.product).slice(0, 20)].join("|");

  function addItems(list) {
    const have = new Set(items.map(dupKey));
    let added = 0, dup = 0;
    for (const it of list) {
      const k = dupKey(it);
      if (have.has(k)) { dup++; continue; }
      have.add(k); items.push(it); added++;
    }
    return { added, dup };
  }

  /* =================================================================
     가져오기: 엑셀
     ================================================================= */
  async function importExcel(buf, name) {
    const wb = await QO.loadWorkbook(buf.slice(0));
    const pv = QO.previewAny(wb, 5000);
    if (!pv.columns.length) throw new Error("표를 찾지 못했어요. 다른 시트인지 확인해주세요.");
    const sig = MAP.signature(pv.columns);
    const saved = maps[sig];
    const auto = Object.assign(MAP.autoMap(pv.columns, FIELDS), saved || {});
    const run = async m => {
      maps[sig] = m;
      await DB.set("csMaps", maps);
      const list = rowsToItems(pv.rows, m, name);
      const r = addItems(list);
      await save(); draw();
      msg("msg-c", "ok", `✔ ${esc(name)} — ${r.added}건 추가${r.dup ? ` (중복 ${r.dup}건 제외)` : ""}`);
    };
    if (MAP.ok(auto, FIELDS) && saved) { await run(auto); return; }
    MAP.open({
      title: "CS 열 맞추기",
      sub: `${name} — 어느 열이 무슨 값인지 정해주세요. 같은 양식은 다음부터 자동입니다.`,
      columns: pv.columns, fields: FIELDS, saved: auto,
      onOk: m => run(m).catch(e => msg("msg-c", "err", "⚠ " + e.message)),
    });
  }
  function rowsToItems(rows, m, src) {
    const g = (r, k) => (m[k] === undefined ? "" : s(r[m[k]]));
    const out = [];
    for (const r of rows) {
      const orderNo = g(r, "orderNo");
      const product = g(r, "product");
      const content = g(r, "content");
      if (!orderNo && !product && !content) continue;
      const type = normType(g(r, "type"), g(r, "title") + " " + content);
      const it = {
        id: uid(), src: src || "엑셀",
        date: toYmd(g(r, "date")) || today(),
        mall: g(r, "mall"), orderNo, customer: g(r, "customer"), phone: g(r, "phone"),
        product, option: g(r, "option"), qty: num(g(r, "qty")) || "",
        type, title: g(r, "title"), content,
        status: normStatus(g(r, "status")),
        brand: "", vendor: "", memo: "", retInv: "", reInv: "", fault: "", cost: 0, refund: 0,
        createdAt: Date.now(), updatedAt: Date.now(),
      };
      it.brand = findBrand(it.product, it.option);
      it.vendor = findVendor(it.product, it.option);
      out.push(it);
    }
    return out;
  }

  /* =================================================================
     가져오기: 메일 (본문 문의 + 첨부 엑셀)
     ================================================================= */
  const DEF_KW = ["문의", "교환", "반품", "취소", "환불", "누락", "파손", "불량", "CS"];
  let mailItems = [], mailSel = new Set(), mailDays = 3;

  async function csFilter() {
    return {
      keywords: await DB.get("csKeywords", DEF_KW),
      senders: await DB.get("csSenders", []),
      exclude: await DB.get("csExclude", ["발주", "정산"]),
    };
  }
  async function drawFilterLine() {
    const f = await csFilter();
    const el = $("cs-filter");
    if (el) el.textContent = "CS 검색: " + (f.keywords.join(", ") || "(키워드 없음)")
      + (f.senders.length ? ` · 보낸곳 ${f.senders.join(", ")}` : "");
  }

  async function openMailModal() {
    try { await ensureGmail(); } catch (e) { return; }
    mailSel = new Set();
    $("csmail-ok").disabled = true;
    $("csmailmodal").classList.add("on");
    document.querySelectorAll("#csmail-period button")
      .forEach(x => x.classList.toggle("on", Number(x.dataset.d) === mailDays));
    await loadMailList();
  }
  async function loadMailList() {
    const list = $("csmail-list");
    const dayTxt = mailDays === 1 ? "오늘" : `최근 ${mailDays}일`;
    list.innerHTML = `<div class="empty">${dayTxt} 메일을 확인하고 있어요…<br><span id="csmail-prog"></span></div>`;
    $("csmail-sub").textContent = `${dayTxt} 확인 중…`;
    mailSel = new Set(); $("csmail-ok").disabled = true;
    try {
      const f = await csFilter();
      mailItems = await GMAIL.listTextMails({
        days: mailDays, keywords: f.keywords, senders: f.senders, exclude: f.exclude, max: 40,
        onProgress: (i, n) => { const p = $("csmail-prog"); if (p) p.textContent = `${i} / ${n}`; },
      });
      if (!mailItems.length) {
        list.innerHTML = `<div class="empty">${dayTxt}간 해당하는 메일이 없어요 — 기간을 늘리거나 검색조건을 바꿔보세요</div>`;
        $("csmail-sub").textContent = "결과 없음";
        return;
      }
      $("csmail-sub").textContent = `여러 개 선택 가능 · ${mailItems.length}건`;
      list.innerHTML = "";
      mailItems.forEach((m, i) => {
        const frm = m.from.includes("<") ? m.from.split("<").pop().replace(">", "") : m.from;
        const el = document.createElement("div");
        el.className = "mitem";
        const t = guessType(m.subject + " " + m.body) || "문의";
        el.innerHTML = `<div style="display:flex;align-items:center;gap:6px;margin-bottom:4px">
            <span class="tag t-${t}">${t}</span>
            <span style="font-size:11.5px;color:var(--muted);font-weight:700">${esc(m.date)}</span>
            ${m.atts.length ? `<span style="font-size:11px;color:var(--brand);font-weight:700">📎 ${m.atts.length}</span>` : ""}
          </div>
          <div style="font-weight:700;font-size:13px;word-break:break-all">${esc(m.subject || "(제목 없음)")}</div>
          <div style="font-size:11px;color:var(--faint);margin-top:2px">${esc(frm)}</div>
          ${m.body.trim() ? `<div style="font-size:11.5px;color:var(--muted);margin-top:6px;padding-top:6px;border-top:1px dashed var(--line);white-space:pre-wrap;max-height:74px;overflow:auto">${esc(m.body.trim().slice(0, 300))}</div>` : ""}`;
        el.onclick = () => {
          const on = mailSel.has(i);
          if (on) mailSel.delete(i); else mailSel.add(i);
          el.style.borderColor = on ? "" : "var(--brand)";
          el.style.background = on ? "" : "var(--brand-soft)";
          $("csmail-ok").disabled = !mailSel.size;
        };
        list.appendChild(el);
      });
    } catch (e) {
      list.innerHTML = `<div class="empty">⚠ ${esc(e.message)}</div>`;
      $("csmail-sub").textContent = "오류";
    }
  }
  async function takeMails() {
    const btn = $("csmail-ok");
    btn.disabled = true; btn.textContent = "가져오는 중…";
    let added = 0, dup = 0, files = 0;
    try {
      const news = [];
      for (const i of mailSel) {
        const m = mailItems[i];
        // 첨부 엑셀이 있으면 CS 표로 취급 → 열 맞추기 흐름
        if (m.atts.length) {
          for (const a of m.atts) {
            const buf = await GMAIL.getAttachment(m.id, a.attachmentId);
            $("csmailmodal").classList.remove("on");
            await importExcel(buf, a.filename);
            files++;
          }
          continue;
        }
        const frm = m.from.includes("<") ? m.from.split("<").pop().replace(">", "") : m.from;
        const body = s(m.body);
        const it = {
          id: uid(), src: "메일",
          date: toYmd(m.date) || today(),
          mall: "", orderNo: (body.match(/\b\d{9,20}\b/) || [""])[0],
          customer: (m.from.split("<")[0] || "").replace(/["']/g, "").trim(), phone: "",
          product: "", option: "", qty: "",
          type: guessType(m.subject + " " + body) || "문의",
          title: s(m.subject), content: body.slice(0, 1200),
          status: "접수", brand: "", vendor: "", memo: "보낸사람: " + frm,
          retInv: "", reInv: "", fault: "", cost: 0, refund: 0,
          mailId: m.id, createdAt: Date.now(), updatedAt: Date.now(),
        };
        news.push(it);
      }
      if (news.length) {
        const r = addItems(news);
        added = r.added; dup = r.dup;
        await save(); draw();
      }
      $("csmailmodal").classList.remove("on");
      if (added || dup) msg("msg-c", "ok", `✔ 메일에서 ${added}건 등록${dup ? ` (중복 ${dup}건 제외)` : ""}`);
      else if (!files) msg("msg-c", "warn", "가져온 항목이 없습니다.");
    } catch (e) {
      msg("msg-c", "err", "⚠ " + e.message);
    } finally { btn.disabled = false; btn.textContent = "선택 항목 가져오기"; }
  }

  /* =================================================================
     목록 그리기
     ================================================================= */
  function match(x) {
    if (filter.st === "미처리" && !OPEN_ST.includes(x.status)) return false;
    if (filter.st !== "전체" && filter.st !== "미처리" && x.status !== filter.st) return false;
    if (filter.type !== "전체" && x.type !== filter.type) return false;
    const q = filter.q.trim().toLowerCase();
    if (q) {
      const hay = [x.orderNo, x.customer, x.phone, x.product, x.option, x.title, x.content, x.vendor, x.mall, x.memo]
        .map(v => s(v).toLowerCase()).join(" ");
      if (!hay.includes(q)) return false;
    }
    return true;
  }
  function drawSegs() {
    const stBar = $("cs-fstatus"), tyBar = $("cs-ftype");
    const stList = ["미처리", "전체"].concat(STATUS);
    stBar.innerHTML = stList.map(v => {
      const n = v === "전체" ? items.length
        : v === "미처리" ? openCount() : items.filter(x => x.status === v).length;
      return `<button class="seg${filter.st === v ? " on" : ""}" data-v="${esc(v)}">${esc(v)} ${n}</button>`;
    }).join("");
    stBar.querySelectorAll("button").forEach(b => b.onclick = () => { filter.st = b.dataset.v; draw(); });
    const tyList = ["전체"].concat(TYPES);
    tyBar.innerHTML = tyList.map(v => {
      const n = v === "전체" ? items.length : items.filter(x => x.type === v).length;
      if (v !== "전체" && !n) return "";
      return `<button class="seg${filter.type === v ? " on" : ""}" data-v="${esc(v)}">${esc(v)} ${n}</button>`;
    }).join("");
    tyBar.querySelectorAll("button").forEach(b => b.onclick = () => { filter.type = b.dataset.v; draw(); });
  }
  function draw() {
    drawSegs(); badge();
    const box = $("cs-list");
    const list = items.filter(match);
    $("cs-count").textContent = `${list.length} / ${items.length}건`;
    if (!items.length) {
      box.innerHTML = `<div class="empty">아직 등록된 CS가 없습니다</div>`;
      $("cs-foot").textContent = ""; drawActions(); return;
    }
    if (!list.length) {
      box.innerHTML = `<div class="empty">조건에 맞는 건이 없습니다.</div>`;
      $("cs-foot").textContent = ""; drawActions(); return;
    }
    const LIMIT = 200;
    box.innerHTML = list.slice(0, LIMIT).map(x => {
      const on = checked.has(x.id);
      const meta = [x.date, x.mall, x.orderNo && "주문 " + x.orderNo, x.customer,
                    x.vendor && "🏭 " + x.vendor].filter(Boolean).map(v => esc(v)).join("</span><span>");
      return `<div class="csrow${on ? " on" : ""}" data-id="${x.id}">
        <div class="cshead">
          <span class="box chk">${on ? CHK : ""}</span>
          <span class="tag t-${esc(x.type)}">${esc(x.type)}</span>
          <b>${esc(x.product || x.title || s(x.content).slice(0, 30) || "(내용 없음)")}</b>
          <span class="st ${esc(x.status)}">${esc(x.status)}</span>
        </div>
        <div class="csmeta"><span>${meta}</span></div>
        ${x.content ? `<div class="csbody">${esc(x.content)}</div>` : ""}
      </div>`;
    }).join("");
    $("cs-foot").textContent = list.length > LIMIT ? `최근 ${LIMIT}건만 표시 (전체 ${list.length}건)` : "";
    box.querySelectorAll(".csrow").forEach(el => {
      const id = el.dataset.id;
      el.querySelector(".chk").onclick = e => {
        e.stopPropagation();
        if (checked.has(id)) checked.delete(id); else checked.add(id);
        draw();
      };
      el.onclick = () => openDetail(id);
    });
    drawActions();
  }

  /* =================================================================
     선택 처리 (상태 일괄 변경 / 업체 전달)
     ================================================================= */
  function drawActions() {
    const card = $("cs-actions");
    const sel = items.filter(x => checked.has(x.id));
    if (!sel.length) { card.style.display = "none"; return; }
    card.style.display = "block";
    $("cs-bulk").innerHTML = `<span class="seg" style="border:none;background:none;color:var(--muted)">${sel.length}건 선택 →</span>`
      + STATUS.map(v => `<button class="seg" data-st="${esc(v)}">${esc(v)}로</button>`).join("")
      + `<button class="seg" data-clr="1">선택해제</button>`;
    $("cs-bulk").querySelectorAll("button").forEach(b => b.onclick = async () => {
      if (b.dataset.clr) { checked.clear(); draw(); return; }
      sel.forEach(x => { x.status = b.dataset.st; x.updatedAt = Date.now(); });
      await save(); draw();
      msg("msg-c", "ok", `✔ ${sel.length}건을 '${b.dataset.st}'(으)로 바꿨어요.`);
    });

    // 업체별 묶기
    const groups = {};
    sel.forEach(x => { const v = x.vendor || "(업체 미지정)"; (groups[v] = groups[v] || []).push(x); });
    const box = $("cs-send");
    box.innerHTML = "";
    Object.keys(groups).sort().forEach(v => {
      const rows = groups[v];
      const row = document.createElement("div");
      row.className = "rrow";
      row.innerHTML = `<div class="rtop"><b style="flex:1">🏭 ${esc(v)}</b><span class="cnt">${rows.length}건</span></div>
        <div class="rmail"><input type="email" placeholder="업체 이메일 (쉼표로 여러 명)" autocapitalize="off" spellcheck="false">
          <button class="dlbtn">📧 전달</button></div>
        <div class="cands"></div>
        <div style="display:flex;gap:6px;margin-top:8px">
          <button class="minibtn xls" style="flex:1;padding:9px">📥 이 업체 엑셀</button>
        </div>`;
      box.appendChild(row);
      const inp = row.querySelector("input");
      const known = v !== "(업체 미지정)" ? (S.vendorEmails[v] || "") : "";
      inp.value = known;
      fillRecipients(row.querySelector(".cands"), inp, {
        saved: known, history: S.vendorSent[v] || [],
        domains: S.vendorDomains[v] || [], query: v !== "(업체 미지정)" ? v : "",
      });
      row.querySelector(".xls").onclick = () => exportRows(rows, `${QO.todayStr()}_${CONFIG.company}_${cleanVendor(v)}_CS내역.xlsx`);
      row.querySelector(".dlbtn").onclick = () => sendToVendor(v, rows, inp, row.querySelector(".dlbtn"));
    });
  }

  function csText(rows) {
    return rows.map((x, i) => {
      const L = [`${i + 1}. [${x.type}] ${x.product || x.title || ""}`];
      if (x.orderNo) L.push(`   주문번호: ${x.orderNo}${x.mall ? " (" + x.mall + ")" : ""}`);
      if (x.customer || x.phone) L.push(`   고객: ${x.customer}${x.phone ? " / " + x.phone : ""}`);
      if (x.option) L.push(`   옵션: ${x.option}${x.qty ? " / 수량 " + x.qty : ""}`);
      if (x.content) L.push(`   내용: ${s(x.content).replace(/\n/g, " ").slice(0, 300)}`);
      if (x.retInv) L.push(`   회수송장: ${x.retInv}`);
      if (x.reInv) L.push(`   재발송송장: ${x.reInv}`);
      return L.join("\n");
    }).join("\n\n");
  }
  async function sendToVendor(vendor, rows, inp, btn) {
    const to = parseEmails(inp.value);
    if (!to.length) { alert("받는사람 이메일을 입력해주세요."); return; }
    const old = btn.textContent;
    btn.disabled = true; btn.textContent = "보내는 중…";
    try {
      await ensureGmail();
      const buf = await buildExcel(rows);
      const fname = `${QO.todayStr()}_${CONFIG.company}_${cleanVendor(vendor)}_CS내역.xlsx`;
      const tpl = MAILTPL.render("cs", { 회사: CONFIG.company, 업체: vendor,
        날짜: QO.fmtDate(QO.todayStr()), 건수: rows.length });
      const body = tpl.body + "\n" + csText(rows) + `\n\n자세한 내용은 첨부 파일을 참고해주세요.\n감사합니다.`;
      await GMAIL.send({
        to: to.join(", "), subject: tpl.subject, body,
        attachments: [{ filename: fname, data: buf }],
      });
      if (vendor !== "(업체 미지정)") {
        S.vendorSent[vendor] = mergeRecent(S.vendorSent[vendor], to);
        await DB.set("vendorSent", S.vendorSent);
      }
      rows.forEach(x => { if (x.status === "접수") { x.status = "업체확인"; x.updatedAt = Date.now(); } });
      await save(); draw();
      btn.textContent = "✔ 보냈어요";
      setTimeout(() => { btn.textContent = old; btn.disabled = false; }, 2500);
    } catch (e) {
      alert("발송 실패: " + e.message);
      btn.textContent = old; btn.disabled = false;
    }
  }

  /* =================================================================
     엑셀 저장
     ================================================================= */
  const XL_COLS = [
    ["접수일", "date"], ["유형", "type"], ["상태", "status"], ["쇼핑몰", "mall"], ["주문번호", "orderNo"],
    ["고객명", "customer"], ["연락처", "phone"], ["상품명", "product"], ["옵션", "option"], ["수량", "qty"],
    ["담당업체", "vendor"], ["제목", "title"], ["내용", "content"], ["회수송장", "retInv"], ["재발송송장", "reInv"],
    ["귀책", "fault"], ["비용", "cost"], ["환불액", "refund"], ["메모", "memo"], ["출처", "src"],
  ];
  async function buildExcel(rows) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("CS");
    ws.addRow(XL_COLS.map(c => c[0]));
    ws.getRow(1).font = { bold: true };
    ws.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEFF3FB" } };
    rows.forEach(x => ws.addRow(XL_COLS.map(c => {
      const v = x[c[1]];
      return (c[1] === "cost" || c[1] === "refund" || c[1] === "qty") ? (v || "") : s(v);
    })));
    ws.columns.forEach((c, i) => { c.width = [12, 8, 10, 12, 20, 10, 14, 30, 18, 7, 14, 24, 44, 16, 16, 8, 10, 10, 20, 12][i] || 14; });
    ws.views = [{ state: "frozen", ySplit: 1 }];
    return await wb.xlsx.writeBuffer();
  }
  async function exportRows(rows, filename) {
    if (!rows.length) { alert("저장할 내용이 없습니다."); return; }
    const buf = await buildExcel(rows);
    download(buf, filename);
  }

  /* =================================================================
     상세 / 직접 등록 모달
     ================================================================= */
  function fldText(label, key, val, type) {
    return `<div class="fld"><label>${esc(label)}</label>
      <input data-k="${key}" type="${type || "text"}" value="${esc(val == null ? "" : val)}" autocapitalize="off"></div>`;
  }
  function fldSel(label, key, val, opts) {
    return `<div class="fld"><label>${esc(label)}</label><select data-k="${key}">`
      + opts.map(o => `<option value="${esc(o)}"${s(val) === s(o) ? " selected" : ""}>${esc(o || "(없음)")}</option>`).join("")
      + `</select></div>`;
  }
  function openDetail(id) {
    const x = id ? items.find(i => i.id === id) : null;
    editing = x ? Object.assign({}, x) : {
      id: "", src: "직접등록", date: today(), mall: "", orderNo: "", customer: "", phone: "",
      product: "", option: "", qty: "", type: "문의", title: "", content: "", status: "접수",
      brand: "", vendor: "", memo: "", retInv: "", reInv: "", fault: "", cost: 0, refund: 0,
    };
    const vendors = [""].concat((S.forms || []).map(f => f.name));
    if (editing.vendor && !vendors.includes(editing.vendor)) vendors.push(editing.vendor);
    $("cs-title").textContent = x ? "CS 상세" : "CS 직접 등록";
    $("cs-sub").textContent = x ? `${x.src || ""} · 등록 ${toYmd(new Date(x.createdAt || Date.now()))}` : "전화·카톡 등으로 받은 문의를 등록합니다.";
    $("cs-del").style.display = x ? "" : "none";
    $("cs-body").innerHTML =
      `<div class="fld2">${fldSel("유형", "type", editing.type, TYPES)}${fldSel("상태", "status", editing.status, STATUS)}</div>
       <div class="fld2">${fldText("접수일", "date", editing.date)}${fldText("쇼핑몰", "mall", editing.mall)}</div>
       ${fldText("주문번호", "orderNo", editing.orderNo)}
       <div class="fld2">${fldText("고객명", "customer", editing.customer)}${fldText("연락처", "phone", editing.phone)}</div>
       ${fldText("상품명", "product", editing.product)}
       <div class="fld2">${fldText("옵션", "option", editing.option)}${fldText("수량", "qty", editing.qty)}</div>
       ${fldSel("담당 업체", "vendor", editing.vendor, vendors)}
       ${fldText("제목", "title", editing.title)}
       <div class="fld"><label>내용</label><textarea data-k="content">${esc(editing.content)}</textarea></div>
       <div class="fld2">${fldText("회수 송장", "retInv", editing.retInv)}${fldText("재발송 송장", "reInv", editing.reInv)}</div>
       <div class="fld2">${fldSel("귀책", "fault", editing.fault, FAULTS)}${fldText("비용(업체 차감)", "cost", editing.cost, "number")}</div>
       ${fldText("환불액", "refund", editing.refund, "number")}
       <div class="fld"><label>메모</label><textarea data-k="memo">${esc(editing.memo)}</textarea></div>`;
    $("csmodal").classList.add("on");
  }
  function readDetail() {
    const o = Object.assign({}, editing);
    $("cs-body").querySelectorAll("[data-k]").forEach(el => { o[el.dataset.k] = el.value; });
    o.cost = num(o.cost); o.refund = num(o.refund);
    o.date = toYmd(o.date) || today();
    return o;
  }

  /* =================================================================
     이벤트
     ================================================================= */
  function bind() {
    $("f-cs").addEventListener("change", async function () {
      const fs = [...this.files]; this.value = "";
      for (const f of fs) {
        $("cs-fname").textContent = "📄 " + f.name;
        try { await importExcel(await readFile(f), f.name); }
        catch (e) { msg("msg-c", "err", "⚠ " + f.name + " — " + e.message); }
      }
    });
    bindDrop("drop-cs", async fs => {
      for (const f of fs) {
        $("cs-fname").textContent = "📄 " + f.name;
        try { await importExcel(await readFile(f), f.name); }
        catch (e) { msg("msg-c", "err", "⚠ " + f.name + " — " + e.message); }
      }
    });
    $("cs-drive").onclick = () => openDrivePicker({
      key: "cs", multiple: true, title: "드라이브에서 CS 파일 가져오기",
      onPick: async files => {
        for (const f of files) {
          try {
            const r = await GMAIL.driveFetchExcel(f.id);
            await importExcel(r.buf, r.name || f.name);
          } catch (e) { msg("msg-c", "err", "⚠ " + f.name + " — " + e.message); }
        }
      },
    });
    $("cs-mail").onclick = () => openMailModal();
    $("csmail-cancel").onclick = () => $("csmailmodal").classList.remove("on");
    $("csmailmodal").onclick = e => { if (e.target === $("csmailmodal")) $("csmailmodal").classList.remove("on"); };
    $("csmail-ok").onclick = () => takeMails();
    document.querySelectorAll("#csmail-period button").forEach(b => {
      b.onclick = () => {
        mailDays = Number(b.dataset.d) || 1;
        document.querySelectorAll("#csmail-period button").forEach(x => x.classList.toggle("on", x === b));
        if ($("csmailmodal").classList.contains("on")) loadMailList();
      };
    });
    $("cs-filter-btn").onclick = () => openCsFilter();

    $("cs-new").onclick = () => openDetail(null);
    $("cs-close").onclick = () => $("csmodal").classList.remove("on");
    $("csmodal").onclick = e => { if (e.target === $("csmodal")) $("csmodal").classList.remove("on"); };
    $("cs-save").onclick = async () => {
      const o = readDetail();
      o.brand = o.brand || findBrand(o.product, o.option);
      if (!o.vendor) o.vendor = findVendor(o.product, o.option);
      o.updatedAt = Date.now();
      if (o.id) {
        const i = items.findIndex(x => x.id === o.id);
        if (i >= 0) items[i] = o;
      } else {
        o.id = uid(); o.createdAt = Date.now();
        items.push(o);
      }
      await save(); draw();
      $("csmodal").classList.remove("on");
      msg("msg-c", "ok", "✔ 저장했습니다.");
    };
    $("cs-del").onclick = async () => {
      if (!editing || !editing.id) return;
      if (!confirm("이 CS 건을 삭제할까요?")) return;
      items = items.filter(x => x.id !== editing.id);
      checked.delete(editing.id);
      await save(); draw();
      $("csmodal").classList.remove("on");
    };

    $("cs-q").addEventListener("input", () => { filter.q = $("cs-q").value; draw(); });
    $("cs-export").onclick = () => exportRows(items.filter(match), `${QO.todayStr()}_${CONFIG.company}_CS내역.xlsx`);
    $("cs-purge").onclick = async () => {
      const done = items.filter(x => x.status === "완료").length;
      if (!done) { alert("완료된 건이 없습니다."); return; }
      if (!confirm(`완료된 ${done}건을 목록에서 지웁니다.\n(먼저 엑셀로 저장해두세요)`)) return;
      items = items.filter(x => x.status !== "완료");
      checked.clear();
      await save(); draw();
      msg("msg-c", "ok", `✔ 완료 ${done}건을 정리했습니다.`);
    };
  }

  /* CS 검색조건 — 발주/회신과 같은 검색조건 모달을 재사용 (mode: cs) */
  function openCsFilter() { openFilter("cs"); }

  /* ---------- 시작 ---------- */
  async function init() {
    bind();
    await load();
    badge();
    drawFilterLine();
  }
  async function onShow() {
    if (!drawn) { await load(); drawn = true; }
    draw();
  }
  /* 동기화로 데이터가 바뀐 뒤 다시 읽기 */
  async function reload() { await load(); if (drawn) draw(); badge(); }

  init();
  return { onShow, reload, draw, save, items: () => items, TYPES, STATUS,
           findVendor, findBrand, toYmd, drawFilter: drawFilterLine };
})();
window.CS = CS;
