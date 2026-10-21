/* MMS 쉐어포인트 자료 읽기 — 발주현황 · 구매요청 · 판매현황 대시보드 공통 (세 저장소에 같은 파일)
 *
 * 사내 로컬PC(MMS)가 매시간 이카운트 자료를 쉐어포인트 [TDC › 문서 › MMS] 에 올린다.
 *   MMS/태성정밀/재고현황/재고현황_최신.xlsx, MMS/플로우텍/판매현황_전체/판매현황_전체_최신.xlsx …
 * 대시보드는 회사 마이크로소프트 계정으로 로그인해 그 파일을 읽기만 한다(비밀번호는 보지 않음).
 *
 * 클라이언트 · 테넌트 ID 는 비밀번호가 아니라 앱 이름표라 코드에 둔다.
 * 로그인 후 돌아오는 주소는 이미 앱 등록에 있는 발주현황의 auth.html 을 세 대시보드가 같이 쓴다
 * (모두 같은 사이트 주소라 로그인 창 결과를 그대로 받을 수 있다).
 */
(function () {
  'use strict';

  const CFG = {
    clientId: '04714dbd-d7c3-4e27-b5e1-21ca15176c83',
    tenantId: 'affff796-2619-44f3-979f-6d2852623593',
    host: 'tsbath.sharepoint.com',
    site: '/sites/TDC',
    root: 'MMS',
    scopes: ['Sites.Read.All'],
    redirect: location.origin + '/po-dashboard/auth.html',
  };
  const COMPANIES = ['태성정밀', '플로우텍'];

  let appPromise = null;

  function loadMsal() {
    if (window.msal) return Promise.resolve();
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = window.MMS_SP_MSAL_SRC || 'vendor/msal-browser.min.js';
      s.onload = resolve;
      s.onerror = () => reject(new Error('로그인 도구(vendor/msal-browser.min.js)를 불러오지 못했습니다'));
      document.head.appendChild(s);
    });
  }

  function getApp() {
    if (!appPromise) {
      appPromise = (async () => {
        await loadMsal();
        const app = new msal.PublicClientApplication({
          auth: {
            clientId: CFG.clientId,
            authority: 'https://login.microsoftonline.com/' + CFG.tenantId,
            redirectUri: CFG.redirect,
          },
          cache: { cacheLocation: 'localStorage' },
        });
        await app.initialize();
        return app;
      })().catch(e => { appPromise = null; throw e; });
    }
    return appPromise;
  }

  function accountOf(app) {
    const accs = app.getAllAccounts();
    return accs.find(a => a.tenantId === CFG.tenantId) || null;
  }

  /** 로그인 토큰. interactive=false 면 창을 띄우지 않고, 로그인이 필요하면 needLogin 오류를 낸다 */
  async function token(interactive) {
    const app = await getApp();
    const acc = accountOf(app);
    if (acc) {
      try {
        return (await app.acquireTokenSilent({ scopes: CFG.scopes, account: acc })).accessToken;
      } catch (e) {
        if (!interactive) { e.needLogin = true; throw e; }
      }
    } else if (!interactive) {
      const e = new Error('쉐어포인트 로그인이 필요합니다');
      e.needLogin = true;
      throw e;
    }
    const r = await app.acquireTokenPopup({ scopes: CFG.scopes, prompt: acc ? undefined : 'select_account' });
    return r.accessToken;
  }

  /** 이 브라우저에 로그인 기록이 있는지 (창을 띄우지 않고 자동으로 가져와도 되는지) */
  async function signedIn() {
    try { return !!accountOf(await getApp()); } catch (e) { return false; }
  }

  // 사이트 주소(…:/sites/TDC:)와 파일 경로(root:/…)를 한 주소에 이어 쓰면 Graph 가 거부한다
  // ("Resource not found for the segment 'root:'"). 사이트 고유번호를 먼저 받아 두고 그걸로 파일을 찾는다.
  const SITE_KEY = 'mmsSpSiteId:' + CFG.host + CFG.site;
  let siteIdPromise = null;
  function siteId(t) {
    if (!siteIdPromise) {
      siteIdPromise = (async () => {
        try { const c = localStorage.getItem(SITE_KEY); if (c) return c; } catch (e) {}
        const res = await fetch(`https://graph.microsoft.com/v1.0/sites/${CFG.host}:${CFG.site}?$select=id`,
                                { headers: { Authorization: 'Bearer ' + t } });
        if (!res.ok) throw await httpError(res, '쉐어포인트 사이트 ' + CFG.site);
        const id = (await res.json()).id;
        try { localStorage.setItem(SITE_KEY, id); } catch (e) {}
        return id;
      })().catch(e => { siteIdPromise = null; throw e; });
    }
    return siteIdPromise;
  }

  async function itemUrl(path, t) {
    const parts = [CFG.root].concat(String(path).split('/')).filter(Boolean).map(encodeURIComponent);
    return `https://graph.microsoft.com/v1.0/sites/${await siteId(t)}/drive/root:/${parts.join('/')}`;
  }

  async function httpError(res, what) {
    let msg = '';
    try { const j = await res.json(); msg = j && j.error && (j.error.code + ': ' + j.error.message); } catch (e) {}
    const err = new Error(`${what} — 쉐어포인트 오류 ${res.status}${msg ? ' (' + msg + ')' : ''}`);
    err.status = res.status;
    return err;
  }

  /** 파일 정보(수정 시각, eTag, 다운로드 주소) */
  async function meta(path, interactive) {
    const t = await token(interactive);
    const res = await fetch(await itemUrl(path, t), { headers: { Authorization: 'Bearer ' + t }, cache: 'no-store' });
    if (!res.ok) throw await httpError(res, path);
    return res.json();
  }

  /** 파일 내용(ArrayBuffer)과 정보. 받은 적 있는 같은 eTag 면 bytes 없이 정보만 돌려준다 */
  async function download(path, interactive, knownETag) {
    const m = await meta(path, interactive);
    if (knownETag && m.eTag === knownETag) return { meta: m, bytes: null };
    const url = m['@microsoft.graph.downloadUrl'];
    if (!url) throw new Error(path + ' — 다운로드 주소를 받지 못했습니다');
    const res = await fetch(url, { cache: 'no-store' });   // 미리 인증된 주소라 토큰이 필요 없다
    if (!res.ok) throw await httpError(res, path);
    return { meta: m, bytes: await res.arrayBuffer() };
  }

  /** 보고서 최신 파일 경로. 예) reportPath('태성정밀', '재고현황') */
  function reportPath(company, report) {
    return `${company}/${report}/${report}_최신.xlsx`;
  }

  /** 마이크로소프트·쉐어포인트 오류를 쉬운 말로 (사용자가 창을 닫은 경우는 빈 문자열) */
  function errorText(e) {
    const s = String((e && (e.errorCode || '')) + ' ' + (e && e.message || e));
    if (/user_cancelled|user_canceled/.test(s)) return '';
    if (/popup_window_error|empty_window_error|popup.*block/i.test(s))
      return '로그인 창(팝업)이 차단됐어요. 주소창 오른쪽의 팝업 차단 아이콘을 눌러 "항상 허용"으로 바꾼 뒤 다시 눌러주세요.';
    if (/AADSTS90094|AADSTS90095|AADSTS65001|admin.*(consent|approval)|consent_required/i.test(s))
      return '회사 관리자 승인이 필요해요. 앱 등록 → API 권한에서 Microsoft Graph "Sites.Read.All"(위임)을 추가하고 "관리자 동의 허용"을 눌러달라고 요청해주세요.';
    if (/AADSTS50011|redirect_uri/i.test(s))
      return '앱 등록의 리디렉션 주소가 다릅니다. 다음 주소가 "단일 페이지 애플리케이션(SPA)" 으로 등록돼 있어야 합니다:\n' + CFG.redirect;
    if (e && e.status === 401) return '쉐어포인트 로그인이 만료됐어요. 다시 연결해주세요.';
    if (e && e.status === 403) return '이 계정에 쉐어포인트 MMS 폴더를 볼 권한이 없어요. 관리자에게 MMS 폴더 보기 권한을 요청해주세요.';
    if (e && e.status === 404) return '쉐어포인트에서 파일을 찾지 못했어요 (MMS 폴더 구성 확인).\n' + s.trim();
    return s.trim().slice(0, 400);
  }

  window.MmsSp = { CFG, COMPANIES, token, signedIn, meta, download, reportPath, errorText };
})();
