/* 퀵오더 서비스워커 — '항상 최신' 방식 (network-first)
   앱 파일은 매번 새로 받아오고, 인터넷이 안 될 때만 캐시를 사용한다.
   → 재배포하면 새로고침만으로 바로 최신 코드가 뜬다. */
const CACHE = "quickorder-v5.2";
const ASSETS = ["./", "./index.html", "./qo-lock.js", "./qo-logic.js", "./qo-gmail.js", "./qo-sync.js", "./qo-app.js",
  "./manifest.json", "./icon-192.png", "./icon-512.png", "./apple-touch-icon.png"];
// 절대 안 바뀌는 외부 라이브러리는 캐시 우선(빠름)
const CDN = "https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js";

self.addEventListener("install", e => {
  e.waitUntil(
    caches.open(CACHE)
      .then(c => c.addAll([...ASSETS, CDN].map(u => new Request(u, { cache: "reload" }))))
      .catch(() => {})
      .then(() => self.skipWaiting())   // 새 워커 즉시 활성화
  );
});

self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys()
      .then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())  // 열려있는 탭도 새 워커가 즉시 제어
  );
});

self.addEventListener("fetch", e => {
  if (e.request.method !== "GET") return;
  const url = e.request.url;

  // 외부 라이브러리(CDN): 캐시 우선 (변하지 않으므로)
  if (url === CDN) {
    e.respondWith(caches.match(e.request).then(hit => hit || fetch(e.request)));
    return;
  }

  // 나머지(앱 파일·페이지): 네트워크 우선 → 실패하면 캐시
  // cache:"no-cache" 로 서버에 항상 재검증(ETag) → GitHub Pages/브라우저 HTTP 캐시가
  // 옛 index.html·qo-app.js 등을 물고 있어 '버튼 무반응' 같은 짝안맞음이 생기는 걸 막는다.
  const revalReq = new Request(e.request, { cache: "no-cache" });
  e.respondWith(
    fetch(revalReq)
      .then(res => {
        // 성공하면 캐시도 최신으로 갱신 (오프라인 대비)
        const copy = res.clone();
        caches.open(CACHE).then(c => c.put(e.request, copy)).catch(() => {});
        return res;
      })
      .catch(() => caches.match(e.request).then(hit => hit || caches.match("./index.html")))
  );
});
