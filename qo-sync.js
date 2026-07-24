/* ===================================================================
   퀵오더 — 기기 간 동기화 (구글 드라이브 appDataFolder)
   업체 양식·설정을 본인 구글 계정의 숨김 폴더에 저장 →
   다른 기기에서 같은 계정으로 로그인하면 그대로 내려받음. 서버 불필요.
   =================================================================== */
"use strict";
const SYNC = (() => {
  const FILE = "qo-backup.json";
  const KV_KEYS = ["brandVendor", "vendorEmails", "vendorSent", "vendorDomains",
    "invEmails", "invSent", "driveOrderFile", "driveFolders",
    "orderSenders", "orderKeywords", "orderExclude",
    "replySenders", "replyKeywords", "replyExclude"];
  const STAMP_KEY = "qo_sync_stamp";   // 이 기기가 마지막으로 반영/업로드한 시각
  const TIME_KEY = "qo_sync_time";     // 마지막 동기화 시각(표시용)

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
    DB.suspend(true);
    try {
      // 업체 양식: 추가·갱신만 (원격에 없다고 로컬 것을 지우지 않음)
      for (const f of (obj.forms || []))
        await DB.putForm({ name: f.name, file: f.file, checked: f.checked !== false, data: bufFromB64(f.data) });
      // 설정: 객체형(업체메일·브랜드·도메인 등)은 병합, 배열/문자열은 교체
      for (const k in (obj.kv || {})) {
        const remote = obj.kv[k];
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
    onStatus(fn) { onStatus = fn; },
    lastTime, enabled: () => GMAIL.signedIn(),
  };
})();
