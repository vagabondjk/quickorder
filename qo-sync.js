/* ===================================================================
   퀵오더 — 기기 간 동기화 (구글 드라이브 appDataFolder)
   업체 양식·설정을 본인 구글 계정의 숨김 폴더에 저장 →
   다른 기기에서 같은 계정으로 로그인하면 그대로 내려받음. 서버 불필요.
   =================================================================== */
"use strict";
const SYNC = (() => {
  const FILE = "qo-backup.json";
  const KV_KEYS = ["brandVendor", "vendorEmails", "vendorSent",
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

  async function applyBundle(obj) {
    DB.suspend(true);
    try {
      const cur = await DB.listForms();
      for (const f of cur) await DB.delForm(f.name);
      for (const f of (obj.forms || []))
        await DB.putForm({ name: f.name, file: f.file, checked: f.checked !== false, data: bufFromB64(f.data) });
      for (const k in (obj.kv || {})) await DB.set(k, obj.kv[k]);
    } finally { DB.suspend(false); }
    setStamp(obj.updatedAt || Date.now());
    markTime();
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
        await applyBundle(obj);
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
