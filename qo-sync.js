/* ===================================================================
   퀵오더 — 기기 간 동기화 (구글 드라이브 appDataFolder)
   업체 양식·설정을 본인 구글 계정의 숨김 폴더에 저장 →
   다른 기기에서 같은 계정으로 로그인하면 그대로 내려받음. 서버 불필요.
   =================================================================== */
"use strict";
const SYNC = (() => {
  const FILE = CONFIG.backupFile;   // 회사별 백업 파일명 (qo-config.js)
  const KV_KEYS = ["brandVendor", "vendorEmails", "vendorSent", "vendorDomains",
    "invEmails", "invSent", "driveOrderFile", "driveSabFile", "driveFolders",
    "orderSenders", "orderKeywords", "orderExclude",
    "replySenders", "replyKeywords", "replyExclude",
    // v6.0 — CS·정산
    "csItems", "csMaps", "csSenders", "csKeywords", "csExclude",
    "settleRules", "settleMaps", "stSenders", "stKeywords", "stExclude",
    // v6.1 — 업체별 공급가표. 이게 빠지면 기기를 바꿨을 때 단가와 연결표가 통째로 사라진다
    // 지운 업체 양식 표시 — 이게 빠지면 다른 기기의 백업이 지운 양식을 되살린다
    "formsDeleted",
    "priceBook", "priceAliases", "priceAliasInfo", "settleBrandVendor", "settleVendors", "settleCarry",
    "mdRewards", "payMaps", "vendorCarriers",        // 파트너 MD 리워드 조건 — 한 번 정하면 다음 달에도 그대로 쓴다

    // v6.1.6 — 메일 문구
    "mailTemplates"];
  const STAMP_KEY = CONFIG.ls("qo_sync_stamp");   // 이 기기가 마지막으로 반영/업로드한 시각
  const TIME_KEY = CONFIG.ls("qo_sync_time");     // 마지막 동기화 시각(표시용)

  let fileId = null;
  let pushTimer = null;
  let onStatus = () => {};

  const getStamp = () => { try { return Number(localStorage.getItem(STAMP_KEY)) || 0; } catch (e) { return 0; } };
  const setStamp = t => { try { localStorage.setItem(STAMP_KEY, String(t)); } catch (e) {} };
  const markTime = () => { try { localStorage.setItem(TIME_KEY, String(Date.now())); } catch (e) {} };
  const lastTime = () => { try { return Number(localStorage.getItem(TIME_KEY)) || 0; } catch (e) { return 0; } };

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
      forms: forms.map(f => ({ name: f.name, file: f.file, checked: f.checked !== false, data: b64FromBuf(f.data) })),
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
        await DB.putForm({ name: f.name, file: f.file, checked: f.checked !== false, data: bufFromB64(f.data) });
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

  /* 내려받기: 원격이 더 최신이면 적용. 반환 {changed} */
  async function syncDown() {
    if (!GMAIL.signedIn()) { status("offline"); return { changed: false, skipped: true }; }
    status("syncing", "내려받는 중…");
    try {
      if (!fileId) fileId = await GMAIL.driveFind(FILE);
      if (!fileId) { status("ok"); return { changed: false, hadRemote: false }; }   // 아직 백업 없음
      const txt = await GMAIL.driveDownload(fileId);
      const obj = JSON.parse(txt);
      if ((obj.updatedAt || 0) > getStamp()) {
        const r = await applyBundle(obj);
        // 이 기기에만 있던 양식이 있으면 클라우드에도 올려 양쪽을 합집합으로 맞춘다
        if (r && r.needPush) { try { await syncUpNow(); } catch (e) {} }
        status("ok");
        return { changed: true, hadRemote: true };
      }
      markTime(); status("ok");
      return { changed: false, hadRemote: true };
    } catch (e) { status("error", e.message); return { changed: false, error: e.message }; }
  }

  /* 올리기 (즉시) */
  async function syncUpNow() {
    if (!GMAIL.signedIn()) { status("offline"); return; }
    status("syncing", "올리는 중…");
    try {
      const bundle = await buildBundle();
      const txt = JSON.stringify(bundle);
      fileId = await GMAIL.driveUpload(FILE, txt, fileId);
      setStamp(bundle.updatedAt); markTime();
      status("ok");
    } catch (e) { status("error", e.message); }
  }

  /* 데이터 변경 시 debounce 업로드 */
  function pushSoon() {
    if (!GMAIL.signedIn()) return;
    clearTimeout(pushTimer);
    pushTimer = setTimeout(() => { syncUpNow(); }, 2500);
  }

  return {
    syncDown, syncUpNow, pushSoon,
    _apply: applyBundle,        // 검증용 — 드라이브 없이 백업 병합을 돌려볼 수 있게
    onStatus(fn) { onStatus = fn; },
    lastTime, enabled: () => GMAIL.signedIn(),
  };
})();
