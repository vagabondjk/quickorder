/* =====================================================================
   퀵오더 — 회사별 설정 (이 파일 하나만 회사마다 다르게 만든다)

   ▣ 새 회사에 배포할 때
     1) 이 파일을 복사해서 아래 값만 그 회사 것으로 바꾼다.
     2) 나머지 파일(index.html, qo-*.js, sw.js, 아이콘)은 손대지 않고 그대로 복사한다.
     3) 회사별로 다른 주소에 올린다. 예) .../quickorder-a/ , .../quickorder-b/

   ▣ 왜 tenant 를 정해야 하나
     같은 도메인(github.io 등)에 여러 회사를 올리면 브라우저 저장소가 공유된다.
     tenant 를 서로 다르게 두면 저장소 이름이 갈라져 데이터가 절대 안 섞인다.
     ※ 랩노마드(원본)는 tenant 를 "" 로 둬야 기존에 저장된 업체 양식이 그대로 남는다.
   ===================================================================== */
"use strict";

const CONFIG = (() => {
  const C = {

    /* ── ① 회사 구분 ─────────────────────────────────────────────── */
    // 회사 코드(영문 소문자). 저장소 이름에 붙는다. 한 번 정하면 바꾸지 말 것.
    // 원본(랩노마드)은 반드시 "" 로 유지 — 바꾸면 저장된 업체 양식이 안 보인다.
    tenant: "",

    // 메일 제목에 들어가는 회사 이름.  예) [랩노마드] 260727_발주서 송부
    company: "랩노마드",

    // ★ 발주서 파일명 가운데에 들어가는 이름 — 업체가 그대로 받아 보는 이름이다.
    //   20260727_랩노마드_디에스피_발주양식.xlsx
    //           ~~~~~~~~
    //   업체 쪽 파일 정리 규칙과 직결되므로 회사 이름(company)과 따로 둔다.
    //   ※ 랩노마드는 "랩노마드" 고정. 임의로 바꾸지 말 것.
    //   ※ 비워두면 그 자리가 통째로 빠진다 → 20260727_디에스피_발주양식.xlsx
    orderTag: "랩노마드",

    // 잠금 화면에서 "누구에게 비밀번호를 받는지" 안내 문구
    adminLabel: "관리자(JK)",

    // 가입 신청서가 도착할 주소 (구글폼을 안 쓸 때의 연락처 표시용)
    adminEmail: "vagabondjk84@gmail.com",

    /* ── 가입 신청을 구글폼으로 받을 때 ──────────────────────────────
       ① signupFormUrl : 업체가 '사용 신청' 을 누르면 열릴 구글폼 주소
                         (폼 → 보내기 → 링크)
       ② signupSheetId : 그 폼의 '응답 시트' 파일 ID
                         (시트 주소 .../spreadsheets/d/★여기★/edit)
                         마스터가 퀵오더 안에서 신청서를 읽어올 때 쓴다.
       비워두면 구글폼 대신 '신청 코드' 방식으로 동작한다. */
    signupFormUrl: "https://docs.google.com/forms/d/e/1FAIpQLScKdt1Ni67r3Ur4t1Ck8fViR1Yz1WmsvzFrhi6MwGY9Vh3QSw/viewform",
    signupSheetId: "17vmzsm0w5T2w2T9DhX3XKREB7-raVgXulJKXIC30ie8",

    /* ── ② 구글 연동 ─────────────────────────────────────────────── */
    // 이 회사 구글 클라우드 프로젝트의 OAuth 클라이언트 ID (웹 애플리케이션).
    // 만드는 방법은 같이 드린 "구글 연동 설정 안내.txt" 참고.
    clientId: "598124965893-16qej37hhlah9ivtr9hdk76c50ms5aqs.apps.googleusercontent.com",

    // 메일·드라이브 기능을 아예 안 쓰는 회사면 false → 관련 버튼이 숨겨진다.
    useGoogle: true,

    /* ── ③ 메일 검색조건 기본값 ──────────────────────────────────── */
    // 앱을 처음 켰을 때 들어가 있는 값. 사용자가 화면에서 바꾸면 그게 우선한다.
    order: {
      senders: ["onekglobal.co.kr"],                                   // 발주서가 오는 곳
      keywords: ["랩노마드 발주서", "랩노마드발주서", "★랩노마드", "랩노마드"],  // 발주서 키워드
      exclude: ["플라스머", "디에스피", "송장", "회신", "운송장", "택배"],      // 제외할 단어
    },
    reply: {
      senders: [],
      keywords: ["송장", "운송장", "회신"],
      exclude: [],
    },

    /* ── ④ 사용 잠금(월별 비밀번호) ──────────────────────────────── */
    // null  → 원본(랩노마드) 비밀번호를 그대로 사용
    // false → 잠금 없이 바로 사용
    // {salt, master, months} → 이 회사 전용 비밀번호 (발급 요청하세요)
    lock: null,
  };

  /* ── 아래는 자동 계산 — 건드리지 않는다 ────────────────────────── */
  /* 저장소는 '로그인한 구글 계정' 으로 가른다.
     한 주소를 여러 회사가 같이 쓰기 때문에, tenant 만으로는 같은 브라우저에서 섞인다.
     · 계정을 아직 모르면(로그인 전) tenant 기준 이름을 그대로 쓴다 — 예전과 같다.
     · 계정을 알면 그 계정 전용 저장소로 갈아탄다. 드라이브 백업은 원래부터 계정별이라
       거기서 그대로 복구된다.
     ※ 이메일을 그대로 키에 쓰지 않고 짧은 해시로 줄인다 (저장소 이름에 주소가 남지 않게). */
  const base = C.tenant ? "_" + C.tenant : "";
  const HOME = String(C.company || "").trim().replace(/\s+/g, "");   // 이 배포본의 원래 회사
  function shortHash(v) {
    const t = String(v || "").trim().toLowerCase();
    let h = 5381;
    for (let i = 0; i < t.length; i++) h = ((h * 33) ^ t.charCodeAt(i)) >>> 0;
    return h.toString(36);
  }
  /* 저장소 꼬리표는 세 조각이다 — 배포본(tenant) + 로그인한 업체 + 구글 계정.
     ★ 업체 조각이 없으면, 한 주소를 나눠 쓰는 두 회사가 같은 저장소를 본다
       (랩노마드가 올린 업체 양식이 베타브릭스에게 그대로 보였다).
     ★ 이 배포본의 원래 회사(HOME)는 조각을 비워 둔다 — 붙이면 이미 저장해 둔
       업체 양식이 통째로 안 보이게 된다. */
  let coPart = "", acctPart = "", suffix = base;
  C.account = "";                    // 지금 로그인한 계정 (없으면 "")
  /* 잠금·승인·기기ID 처럼 '앱에 들어오기 전' 상태는 업체·계정과 무관해야 한다.
     꼬리표를 붙이면 저장할 때와 읽을 때 키가 달라져 로그인 상태가 사라진다. */
  C.lsBase = k => k + base;
  /* 이 배포본이 원래 어느 회사 것인지. 로그인한 업체(C.company)와 구분해야 한다 —
     아래 설정값들(메일 검색조건 등)이 전부 이 회사 기준으로 적혀 있기 때문이다. */
  C.homeCompany = HOME;
  /* '이 업체가 마지막에 쓴 구글 계정' 처럼 업체별이되 계정과는 무관해야 하는 값 */
  C.lsCompany = k => k + base + coPart;
  function apply() {
    suffix = base + coPart + acctPart;
    C.dbName = "quickorder" + suffix;                                    // IndexedDB 이름
    C.backupFile = "qo-backup" + (C.tenant ? "-" + C.tenant : "") + coPart + ".json";  // 드라이브 백업
    C.ls = k => k + suffix;                                              // localStorage 키
  }
  /* 로그인한 업체가 정해지면 그 업체 저장소로 바꾼다. 바뀌었으면 true. */
  C.useCompany = name => {
    const n = String(name || "").trim().replace(/\s+/g, "");
    const want = (!n || n === HOME) ? "" : "_c" + shortHash(n);
    if (want === coPart) return false;
    coPart = want; apply(); return true;
  };
  /* 계정이 정해지면 저장소를 그 계정 것으로 바꾼다. 바뀌었으면 true. */
  C.useAccount = email => {
    const e = String(email || "").trim().toLowerCase();
    const want = e ? "_u" + shortHash(e) : "";
    if (want === acctPart) { C.account = e; return false; }
    acctPart = want; C.account = e; apply();
    return true;
  };
  apply();
  // 파일명 앞부분: 20260727_랩노마드_업체명_발주양식.xlsx
  C.tag = () => C.company;
  return C;
})();
