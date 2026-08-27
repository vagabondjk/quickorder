/* ===================================================================
   퀵오더 — 기기 간 동기화 (구글 드라이브 appDataFolder)
   업체 양식·설정을 본인 구글 계정의 숨김 폴더에 저장 →
   다른 기기에서 같은 계정으로 로그인하면 그대로 내려받음. 서버 불필요.
   =================================================================== */
"use strict";
const SYNC = (() => {
  /* ★★ 백업 파일명은 '부를 때' 읽어야 한다. 절대 여기서 값으로 받아두지 말 것.
     이 파일은 로그인 화면이 뜨기도 전에 로드된다. 그때 CONFIG.backupFile 은 아직
     이 배포본의 원래 회사(랩노마드) 것이다. 값으로 받아두면 나중에 베타브릭스로
     로그인해도 이름이 그대로라, 랩노마드 백업을 내려받아 베타브릭스 저장소에
     써 넣고( = 남의 업체 양식이 보인다), 반대로 베타브릭스 내용을 랩노마드
     백업 파일에 덮어썼다. 실제로 그런 일이 있었다 (2026-08-17, 모바일에서 발견). */
  const FILE = () => CONFIG.backupFile;   // 회사별 백업 파일명 (qo-config.js)
  const KV_KEYS = ["brandVendor", "vendorEmails", "vendorSent", "vendorDomains",
    "invEmails", "invSent", "driveOrderFile", "driveSabFile", "driveFolders",
    "orderSenders", "orderKeywords", "orderExclude",
    "replySenders", "replyKeywords", "replyExclude",
    // v6.0 — CS·정산
    "csItems", "csMaps", "csSenders", "csKeywords", "csExclude",
    "settleRules", "settleMaps", "stSenders", "stKeywords", "stExclude",
    // v6.1 — 업체별 공급가표. 이게 빠지면 기기를 바꿨을 때 단가와 연결표가 통째로 사라진다
    // 지운 업체 양식 표시 — 이게 빠지면 다른 기기의 백업이 지운 양식을 되살린다
    // 상품명 → 업체 연결표. 빠지면 기기를 바꿨을 때 지정해 둔 것이 통째로 사라진다
    "orderAliases", "orderDateCols",
    /* v2.4.1 — 승인 업체 목록. 이게 빠져 있어서 PC 에 3곳, 휴대폰에 1곳처럼
       기기마다 다른 목록을 들고 있었다 (2026-08-24 신고). 아래 병합 규칙과 한 쌍이다 */
    "masterCompanies", "masterDeleted",
    /* v2.5.0 — 월별 정산 확정 기록. 돈 기록이라 기기마다 달라지면 안 된다 */
    "settleConfirms",
    "formsDeleted", "masterRevoked", "signupSheetId", "rosterSheetId", "settlePins",
    "pbSenders", "pbKeywords", "pbExclude", "paySenders", "payKeywords", "payExclude",
    "priceBook", "priceAliases", "priceAliasInfo", "settleBrandVendor", "settleVendors", "settleCarry",
    "mdRewards", "payMaps", "vendorCarriers",        // 파트너 MD 리워드 조건 — 한 번 정하면 다음 달에도 그대로 쓴다

    // v6.1.6 — 메일 문구
    "mailTemplates"];
  const STAMP_KEY = () => CONFIG.ls("qo_sync_stamp");   // 이 기기가 마지막으로 반영/업로드한 시각
  const TIME_KEY = () => CONFIG.ls("qo_sync_time");     // 마지막 동기화 시각(표시용)

  /* 드라이브 파일 ID 를 기억해 두면 매번 찾지 않아도 된다. 다만 이건 '지금 저장소' 것이라
     업체나 계정이 바뀌면 반드시 버려야 한다 — 안 버리면 이름은 새 회사 것인데
     실제로 쓰는 파일은 앞 회사 것이 된다. reset() 이 그 일을 한다. */
  let fileId = null;
  let fileFor = "";        // 그 ID 가 어느 백업 파일 것인지
  /* 마지막으로 '내가 알고 있는' 백업의 수정시각.
     자동 동기화가 이 값과 드라이브의 값을 비교해, 다를 때만 실제로 내려받는다.
     ★ 내가 올린 뒤에도 반드시 갱신해야 한다. 안 그러면 내 업로드를 남의 변경으로
       착각해서 매번 백업 전체(업체 양식이 base64 로 들어 있어 수 MB)를 다시 받는다. */
  let remoteStamp = "";
  /* ★ 내려받기와 올리기가 겹치면 안 된다.
     syncDown 은 'updatedAt > 내 기록' 을 확인한 뒤에 실제 병합을 한다. 그 사이에
     자동 업로드가 끼어들어 기록을 바꾸면, 방금 올린 내 상태 위에 옛 백업을 덮어쓴다.
     둘 다 이 줄에 세워 한 번에 하나씩만 돌게 한다. */
  let chain = Promise.resolve();
  const serial = fn => (chain = chain.then(fn, fn));
  let pushTimer = null;
  let onStatus = () => {};
  /* 업체·계정이 바뀌면 앞 백업에 대해 알던 것을 전부 버린다.
     수정시각도 같이 버려야 한다 — 남겨두면 새 백업을 '안 바뀌었다' 고 넘겨버린다. */
  function reset() { fileId = null; fileFor = ""; remoteStamp = ""; clearTimeout(pushTimer); pushTimer = null; }
  /* 파일명이 달라졌으면(= 업체가 바뀌었으면) 캐시를 버린다 */
  function guard() { const f = FILE(); if (f !== fileFor) { fileId = null; fileFor = f; remoteStamp = ""; } return f; }

  const getStamp = () => { try { return Number(localStorage.getItem(STAMP_KEY())) || 0; } catch (e) { return 0; } };
  const setStamp = t => { try { localStorage.setItem(STAMP_KEY(), String(t)); } catch (e) {} };
  const markTime = () => { try { localStorage.setItem(TIME_KEY(), String(Date.now())); } catch (e) {} };
  const lastTime = () => { try { return Number(localStorage.getItem(TIME_KEY())) || 0; } catch (e) { return 0; } };

  function status(state, detail) { try { onStatus(state, detail); } catch (e) {} }

  /* base64 <-> ArrayBuffer */
  function b64FromBuf(buf) {
    const u = new Uint8Array(buf); let s = "";
    for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000));
    return btoa(s);
  }
  function bufFromB64(b64) {
    const bin = atob(b64); const u = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    return u.buffer;
  }

  async function buildBundle() {
    const forms = await DB.listForms();
    const kv = {};
    for (const k of KV_KEYS) { const v = await DB.get(k, undefined); if (v !== undefined) kv[k] = v; }
    return {
      version: 1,
      updatedAt: Date.now(),
      kv,
      /* 업체별 설정(상품명 매핑 · 공급가 넣기)도 같이 실어야 기기를 바꿔도 남는다 */
      forms: forms.map(f => ({ name: f.name, file: f.file, checked: f.checked !== false,
        nameMap: f.nameMap || null, nameMapList: f.nameMapList || null, nameMapFile: f.nameMapFile || "",
        withPrice: !!f.withPrice, extra: f.extra || null,
        renameTable: f.renameTable || null, renameFile: f.renameFile || "",
        extraTable: f.extraTable || null, extraFile: f.extraFile || "",
        data: b64FromBuf(f.data) })),
    };
  }

  /* 원격 내용을 이 기기에 반영 — '병합' 방식.
     한쪽 기기에서 실수로 지워도 다른 기기 것이 사라지지 않고, 오히려 복구된다.
     (예전엔 로컬을 통째로 지우고 원격으로 덮어써서, 한쪽 삭제가 양쪽 삭제가 됐음) */
  async function applyBundle(obj) {
    const remoteNames = new Set((obj.forms || []).map(f => f.name));
    const localBefore = await DB.listForms();
    const extraLocal = localBefore.some(f => !remoteNames.has(f.name));   // 이 기기에만 있는 양식
    /* ★ 지운 양식이 백업에서 되살아나지 않게 (2026-08-04)
       예전엔 '추가·갱신만' 했더니, 중복 양식을 지워도 앱이 새로 뜰 때마다
       백업에서 그대로 다시 내려와 업체 2곳에 양식이 4개가 됐다.
       지운 이름과 지운 시각을 남겨두고, 그보다 오래된 백업은 무시한다.
       (다른 기기에서 나중에 다시 올린 것이면 백업이 더 최신이라 살아난다) */
    const tomb = await DB.get("formsDeleted", {}) || {};
    const bundleAt = Number(obj.updatedAt) || 0;
    const revived = [];
    DB.suspend(true);
    try {
      // 업체 양식: 추가·갱신만 (원격에 없다고 로컬 것을 지우지 않음)
      for (const f of (obj.forms || [])) {
        const at = Number(tomb[f.name]) || 0;
        if (at && at >= bundleAt) continue;             // 지운 뒤로 바뀐 게 없으면 되살리지 않는다
        if (at) revived.push(f.name);                   // 백업이 더 최신 → 되살리고 표시는 지운다
        await DB.putForm({ name: f.name, file: f.file, checked: f.checked !== false,
          nameMap: f.nameMap || null, nameMapList: f.nameMapList || null, nameMapFile: f.nameMapFile || "",
          withPrice: !!f.withPrice, extra: f.extra || null,
          renameTable: f.renameTable || null, renameFile: f.renameFile || "",
          extraTable: f.extraTable || null, extraFile: f.extraFile || "",
          data: bufFromB64(f.data) });
      }
      if (revived.length) {
        revived.forEach(n => { delete tomb[n]; });
        await DB.set("formsDeleted", tomb);
      }
      // 설정: 객체형(업체메일·브랜드·도메인 등)은 병합, 배열/문자열은 교체
      for (const k in (obj.kv || {})) {
        const remote = obj.kv[k];
        // CS 목록은 '건 단위 병합' — 한쪽 기기의 등록/수정이 사라지지 않게.
        // 같은 id는 updatedAt이 더 최신인 쪽을 남긴다.
        if (k === "csItems" && Array.isArray(remote)) {
          const local = await DB.get("csItems", []) || [];
          const byId = new Map();
          for (const it of local) if (it && it.id) byId.set(it.id, it);
          for (const it of remote) {
            if (!it || !it.id) continue;
            const cur = byId.get(it.id);
            if (!cur || (it.updatedAt || 0) >= (cur.updatedAt || 0)) byId.set(it.id, it);
          }
          await DB.set("csItems", [...byId.values()]);
          continue;
        }
        /* ★★ 승인 업체 목록 — 이름 단위 병합 (2026-08-24).
           배열이라 아래 기본 규칙대로면 통째로 교체된다. 그러면 휴대폰에 1곳만
           남아 있던 상태가 PC 의 3곳을 그대로 덮어쓴다 — 고치려던 게 지우는 셈이다.
           그래서 이름을 키로 합치고, 같은 이름은 마지막에 손댄 쪽(at)을 남긴다.
           ★ 지운 업체가 다른 기기 백업에서 되살아나지 않게 삭제 시각(masterDeleted)을
             함께 본다. 업체 양식의 formsDeleted 와 같은 방식이다. */
        if (k === "masterCompanies" && Array.isArray(remote)) {
          const asObj = x => (typeof x === "string" ? { name: x, on: true } : Object.assign({ on: true }, x));
          const local = ((await DB.get("masterCompanies", [])) || []).map(asObj);
          // 삭제 표시는 이 기기 것과 백업 것을 합쳐서 본다 (어느 쪽에서 지웠든 존중)
          const tomb = Object.assign({}, (await DB.get("masterDeleted", {})) || {}, (obj.kv || {}).masterDeleted || {});
          const byName = new Map();
          const put = c => {
            if (!c || !c.name) return;
            const cur = byName.get(c.name);
            if (!cur || (Number(c.at) || 0) >= (Number(cur.at) || 0)) byName.set(c.name, c);
          };
          local.forEach(put);
          remote.map(asObj).forEach(put);
          // 지운 뒤로 손댄 적이 없으면 되살리지 않는다 (다시 승인했으면 at 이 더 크다)
          const out = [...byName.values()].filter(c => {
            const at = Number(tomb[c.name]) || 0;
            return !at || (Number(c.at) || 0) > at;
          });
          await DB.set("masterCompanies", out);
          continue;
        }
        /* ★★ 월별 정산 확정 기록 — 건 단위 병합 (2026-08-24).
           달·업체별로 한 건씩 들어 있는데, 통째로 교체하면 PC 에서 확정한 7월 기록이
           휴대폰이 올린 백업에 지워진다. 돈 기록이라 사라지면 안 된다.
           같은 달·같은 업체를 양쪽에서 확정했으면 나중에 확정한 쪽을 남긴다. */
        if (k === "settleConfirms" && remote && typeof remote === "object" && !Array.isArray(remote)) {
          const out = Object.assign({}, (await DB.get("settleConfirms", {})) || {});
          for (const key in remote) {
            const a = out[key], b = remote[key];
            if (!b) continue;
            if (!a || (Number(b.at) || 0) >= (Number(a.at) || 0)) out[key] = b;
          }
          await DB.set("settleConfirms", out);
          continue;
        }
        /* 지운 승인 업체 표시 — 더 늦게 지운 쪽을 남긴다 (formsDeleted 와 같은 규칙) */
        if (k === "masterDeleted" && remote && typeof remote === "object" && !Array.isArray(remote)) {
          const out = Object.assign({}, (await DB.get("masterDeleted", {})) || {});
          for (const n in remote) if ((Number(remote[n]) || 0) > (Number(out[n]) || 0)) out[n] = remote[n];
          await DB.set("masterDeleted", out);
          continue;
        }
        // 지운 양식 표시는 '더 늦게 지운 쪽'을 남긴다.
        // 이번에 되살린 이름은 표시를 지운다 (안 그러면 다음에 또 안 내려온다).
        if (k === "formsDeleted" && remote && typeof remote === "object" && !Array.isArray(remote)) {
          const out = Object.assign({}, await DB.get("formsDeleted", {}) || {});
          for (const n in remote) if ((Number(remote[n]) || 0) > (Number(out[n]) || 0)) out[n] = remote[n];
          revived.forEach(n => { delete out[n]; });
          await DB.set("formsDeleted", out);
          continue;
        }
        /* 공급가표는 '한 덩어리'다. 키별로 섞으면 옛 백업의 sheets 가 새 파일 위에 덮여
           내용은 옛것인데 파일명·시각만 새것인 괴물이 된다. 통째로 새것을 쓴다.
           (더 늦게 읽은 쪽이 이긴다 — at 이 그 기준) */
        if (k === "priceBook" && remote && typeof remote === "object" && !Array.isArray(remote)) {
          const local = await DB.get("priceBook", null);
          const newer = (Number(remote.at) || 0) >= (Number(local && local.at) || 0);
          await DB.set("priceBook", newer ? remote : local);
          continue;
        }
        if (remote && typeof remote === "object" && !Array.isArray(remote)) {
          const local = await DB.get(k, {});
          await DB.set(k, Object.assign({}, local, remote));
        } else {
          await DB.set(k, remote);
        }
      }
    } finally { DB.suspend(false); }
    setStamp(obj.updatedAt || Date.now());
    markTime();
    return { needPush: extraLocal };   // 이 기기에만 있던 게 있으면 클라우드에도 올려 합침
  }

  /* 내려받기: 원격이 더 최신이면 적용. 반환 {changed}
     ※ 바깥에서 부르는 syncDown 은 아래에서 serial() 로 감싼다 — 올리기와 겹치면
       방금 올린 상태 위에 옛 백업을 덮어쓴다. */
  async function downNow() {
    if (!GMAIL.signedIn()) { status("offline"); return { changed: false, skipped: true }; }
    status("syncing", "내려받는 중…");
    try {
      const name = guard();
      if (!fileId) fileId = await GMAIL.driveFind(name);
      if (!fileId) { status("ok"); return { changed: false, hadRemote: false }; }   // 아직 백업 없음
      const txt = await GMAIL.driveDownload(fileId);
      const obj = JSON.parse(txt);
      if ((obj.updatedAt || 0) > getStamp()) {
        const r = await applyBundle(obj);
        // 이 기기에만 있던 양식이 있으면 클라우드에도 올려 양쪽을 합집합으로 맞춘다
        if (r && r.needPush) { try { await upNow(); } catch (e) {} }
        else await noteStamp();          // 방금 받은 것이 최신임을 기록 (다음 확인에서 건너뛰게)
        status("ok");
        return { changed: true, hadRemote: true };
      }
      await noteStamp();
      markTime(); status("ok");
      return { changed: false, hadRemote: true };
    } catch (e) { status("error", e.message); return { changed: false, error: e.message }; }
  }

  /* 올리기 (즉시) */
  async function upNow() {
    if (!GMAIL.signedIn()) { status("offline"); return; }
    status("syncing", "올리는 중…");
    try {
      const name = guard();
      const bundle = await buildBundle();
      const txt = JSON.stringify(bundle);
      const up = await GMAIL.driveUpload(name, txt, fileId);
      /* 예전 버전은 id 문자열만 돌려줬다. 둘 다 받아들인다 — 배포가 엇갈려도 안 깨지게 */
      if (up && typeof up === "object") { fileId = up.id; remoteStamp = up.modifiedTime || ""; }
      else { fileId = up; remoteStamp = ""; }
      setStamp(bundle.updatedAt); markTime();
      status("ok");
    } catch (e) { status("error", e.message); }
  }

  /* 지금 드라이브에 있는 백업의 수정시각을 기억해 둔다 */
  async function noteStamp() {
    try {
      if (!fileId) return;
      const info = await GMAIL.driveFileInfo(fileId);
      remoteStamp = (info && info.modifiedTime) || "";
    } catch (e) {}
  }

  /* ★ 자동 동기화가 쓰는 '싼 확인'.
     백업 전체에는 업체 양식이 base64 로 다 들어 있어 수 MB 다. 그걸 30초마다
     내려받을 수는 없다. 그래서 수정시각(작은 JSON 한 번)만 먼저 보고,
     내가 아는 것과 다를 때에만 실제로 내려받는다.
     내가 올린 직후에는 remoteStamp 가 이미 그 값이라 그냥 지나간다. */
  async function pullIfChanged() {
    if (!GMAIL.signedIn()) return { changed: false, skipped: true };
    try {
      const name = guard();
      if (!fileId) fileId = await GMAIL.driveFind(name);
      if (!fileId) return { changed: false, hadRemote: false };
      const info = await GMAIL.driveFileInfo(fileId);
      const mt = (info && info.modifiedTime) || "";
      if (mt && remoteStamp && mt === remoteStamp) return { changed: false, hadRemote: true };
      return await downNow();
    } catch (e) { status("error", e.message); return { changed: false, error: e.message }; }
  }

  /* 데이터 변경 시 debounce 업로드 */
  function pushSoon() {
    if (!GMAIL.signedIn()) return;
    clearTimeout(pushTimer);
    pushTimer = setTimeout(() => { syncUpNow(); }, 2500);
  }

  /* 바깥으로 나가는 셋은 전부 같은 줄에 세운다 (위 chain 주석 참고) */
  const syncDown = () => serial(downNow);
  const syncUpNow = () => serial(upNow);
  const syncPull = () => serial(pullIfChanged);

  return {
    syncDown, syncUpNow, syncPull, pushSoon, reset,
    backupName: () => FILE(),   // 설정 화면에서 '지금 어느 백업을 쓰는지' 보여준다
    _apply: applyBundle,        // 검증용 — 드라이브 없이 백업 병합을 돌려볼 수 있게
    onStatus(fn) { onStatus = fn; },
    lastTime, enabled: () => GMAIL.signedIn(),
  };
})();
/* ★ const 로 선언한 값은 window 의 속성이 되지 않는다.
   그래서 다른 파일의 `if (window.SYNC)` 가 늘 거짓이었고, 거기 걸린 일이 통째로
   안 돌았다 — 업체 양식을 지웠을 때 '바로 백업에 반영' 하는 것(qo-app.js dropForm)이
   그랬다. qo-cs.js·qo-settle.js 는 원래부터 이렇게 공개하고 있었는데 여기만 빠져 있었다.
   ※ 테스트는 이 파일을 브라우저 밖(샌드박스)에서 돌린다 — window 가 없을 수 있다. */
if (typeof window !== "undefined") window.SYNC = SYNC;
