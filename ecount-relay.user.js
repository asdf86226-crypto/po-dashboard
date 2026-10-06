// ==UserScript==
// @name         발주현황 대시보드 — ECOUNT 자동입력 중계
// @namespace    https://asdf86226-crypto.github.io/po-dashboard/
// @version      1.0.4
// @description  이카운트 발주서·구매·판매 입력을 사용자 브라우저에서 직접 처리(사용자 공인 IP로 호출됨). 발주현황 대시보드에서만 작동.
// @author       Flowtech
// @match        *://asdf86226-crypto.github.io/po-dashboard*
// @match        https://asdf86226-crypto.github.io/po-dashboard/*
// @match        https://asdf86226-crypto.github.io/po-dashboard
// @grant        GM_xmlhttpRequest
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_deleteValue
// @grant        GM_registerMenuCommand
// @grant        unsafeWindow
// @connect      sboapica.ecount.com
// @connect      sboapicb.ecount.com
// @connect      sboapicc.ecount.com
// @connect      sboapicd.ecount.com
// @connect      oapica.ecount.com
// @connect      oapicb.ecount.com
// @connect      oapicc.ecount.com
// @connect      oapicd.ecount.com
// @connect      oapi.ecount.com
// @run-at       document-end
// @updateURL    https://asdf86226-crypto.github.io/po-dashboard/ecount-relay.user.js
// @downloadURL  https://asdf86226-crypto.github.io/po-dashboard/ecount-relay.user.js
// ==/UserScript==

/* eslint-disable */
(function () {
  'use strict';

  // 테스트키는 sandbox(sboapi), 실제키는 운영(oapi) — 자격 설정 시 함께 저장되는 값으로 판단
  function hostFor(cred) {
    const prefix = cred && cred.isTest ? 'sboapi' : 'oapi';
    return 'https://' + prefix + String(cred.zone || '').toLowerCase() + '.ecount.com';
  }

  // ─── 자격증명 저장소 ────────────────────────────────────────
  const CRED_KEY = 'ecount_credentials_v1';
  function getCreds() {
    const raw = GM_getValue(CRED_KEY, '');
    if (!raw) return null;
    try { return JSON.parse(raw); } catch (e) { return null; }
  }
  function setCreds(c) { GM_setValue(CRED_KEY, JSON.stringify(c)); sessionCache = {}; }

  function setupCreds() {
    const cur = getCreds() || {};
    function ask(label, prev) {
      const v = prompt(label + '  (취소하면 설정 중단)', prev || '');
      if (v === null) throw new Error('cancelled');
      return v.trim();
    }
    try {
      const isTest = confirm('이번에 저장할 API 키는 "테스트키" 인가요?\n\n확인 = 테스트키 (sboapi 서버)\n취소 = 실제 운영키 (oapi 서버)');
      const t = cur.taesung || {};
      const f = cur.flowtech || {};
      // 담당자(사원) 코드 — ECOUNT 기초등록 > 사원등록에서 확인. 비워두면 ERP에 담당자 없이 등록됨.
      const empT = ask('태성정밀 담당자 사원코드 EMP_CD (비워도 됨)', t.empCd || '');
      const empF = ask('플로우텍 담당자 사원코드 EMP_CD (비워도 됨)', f.empCd || '');
      const next = {
        taesung: {
          comCode: ask('태성정밀 COM_CODE', t.comCode || '76811'),
          userId:  ask('태성정밀 USER_ID',   t.userId  || 'JMOH'),
          apiKey:  ask('태성정밀 API_CERT_KEY', t.apiKey || ''),
          zone:    ask('태성정밀 ZONE (예: CC)',  t.zone   || 'CC'),
          empCd:   empT,
          isTest,
        },
        flowtech: {
          comCode: ask('플로우텍 COM_CODE', f.comCode || '611343'),
          userId:  ask('플로우텍 USER_ID',   f.userId  || 'JMOH'),
          apiKey:  ask('플로우텍 API_CERT_KEY', f.apiKey || ''),
          zone:    ask('플로우텍 ZONE (예: CA)',  f.zone   || 'CA'),
          empCd:   empF,
          isTest,
        },
      };
      setCreds(next);
      alert('✅ ECOUNT 자격 저장 완료\n\n• 태성 ' + next.taesung.comCode + ' / ZONE ' + next.taesung.zone + ' / 담당자 ' + (next.taesung.empCd || '(없음)') + '\n• 플로우텍 ' + next.flowtech.comCode + ' / ZONE ' + next.flowtech.zone + ' / 담당자 ' + (next.flowtech.empCd || '(없음)') + '\n• 서버: ' + (isTest ? 'sboapi(테스트)' : 'oapi(운영)'));
    } catch (e) {
      if (e.message !== 'cancelled') alert('저장 실패: ' + e.message);
    }
  }
  function clearCreds() {
    if (!confirm('ECOUNT 자격을 모두 삭제할까요?')) return;
    GM_deleteValue(CRED_KEY); sessionCache = {};
    alert('삭제 완료');
  }

  // ─── GM_xmlhttpRequest 를 Promise 로 ────────────────────────
  function gmFetch(url, bodyObj) {
    return new Promise(function (resolve, reject) {
      GM_xmlhttpRequest({
        method: 'POST', url: url,
        headers: { 'Content-Type': 'application/json' },
        data: JSON.stringify(bodyObj || {}),
        timeout: 30000,
        onload: function (res) {
          let json = null;
          try { json = JSON.parse(res.responseText); } catch (e) { /* raw */ }
          resolve({ status: res.status, body: res.responseText, json: json });
        },
        onerror: function (err) { reject(new Error('네트워크 오류: ' + JSON.stringify(err))); },
        ontimeout: function () { reject(new Error('요청 시간 초과(30초)')); },
      });
    });
  }

  // ─── 세션 캐시 (10분 유효) ──────────────────────────────────
  let sessionCache = {};

  async function login(company) {
    const now = Date.now();
    const cached = sessionCache[company];
    if (cached && cached.exp > now) return cached.sid;
    const creds = getCreds();
    if (!creds || !creds[company]) throw new Error('ECOUNT 자격이 없습니다 — Tampermonkey 메뉴 → "ECOUNT 자격 설정"');
    const c = creds[company];
    const url = hostFor(c) + '/OAPI/V2/OAPILogin';
    const res = await gmFetch(url, {
      COM_CODE: c.comCode, USER_ID: c.userId,
      API_CERT_KEY: c.apiKey, LAN_TYPE: 'ko-KR', ZONE: c.zone,
    });
    const sid = res.json && res.json.Data && res.json.Data.Datas && res.json.Data.Datas.SESSION_ID;
    if (!sid) throw new Error('[' + company + '] 로그인 실패: ' + (res.body || '').slice(0, 500));
    sessionCache[company] = { sid: sid, exp: now + 10 * 60 * 1000 };
    return sid;
  }

  async function apiCall(company, path, payload) {
    const sid = await login(company);
    const c = getCreds()[company];
    const url = hostFor(c) + path + (path.indexOf('?') >= 0 ? '&' : '?') + 'SESSION_ID=' + encodeURIComponent(sid);
    const res = await gmFetch(url, payload);
    return res.json || { raw: res.body };
  }

  // ─── 비즈니스 API 래퍼 ──────────────────────────────────────
  // 발주서입력 — bulkDatas: {UPLOAD_SER_NO, IO_DATE, CUST, WH_CD, PROD_CD, QTY, PRICE, REMARKS, ...}
  // 경로: /OAPI/V2/Purchases/SavePurchaseOrder (ECOUNT API 공식 명세)
  async function savePurchaseOrder(company, bulkDatas) {
    return apiCall(company, '/OAPI/V2/Purchases/SavePurchaseOrder', {
      PurchaseOrderList: [{ BulkDatas: bulkDatas }],
    });
  }
  // 구매입력 (입고)
  async function savePurchase(company, bulkDatas) {
    return apiCall(company, '/OAPI/V2/Purchases/SavePurchases', {
      PurchasesList: [{ BulkDatas: bulkDatas }],
    });
  }
  // 판매입력 (출고/매출)
  async function saveSale(company, bulkDatas) {
    return apiCall(company, '/OAPI/V2/Sale/SaveSale', {
      SaleList: [{ BulkDatas: bulkDatas }],
    });
  }
  // 디버그용: 임의의 엔드포인트 호출 (경로/필드 조정 필요할 때)
  async function rawCall(company, path, payload) {
    return apiCall(company, path, payload);
  }

  // ─── 대시보드 페이지에 노출 ─────────────────────────────────
  const relay = {
    __version: '1.0.1',
    __installed: true,
    __isTest: function () { const c = getCreds(); return !!(c && c.taesung && c.taesung.isTest); },
    setupCreds: setupCreds,
    hasCreds: function () { const c = getCreds(); return !!(c && c.taesung && c.taesung.apiKey && c.flowtech && c.flowtech.apiKey); },
    getEmpCd: function (company) { const c = getCreds(); return (c && c[company] && c[company].empCd) || ''; },
    login: login,
    savePurchaseOrder: savePurchaseOrder,
    savePurchase: savePurchase,
    saveSale: saveSale,
    rawCall: rawCall,
  };
  try { unsafeWindow.ecountRelay = relay; } catch (e) { window.ecountRelay = relay; }

  // ─── Tampermonkey 메뉴 ─────────────────────────────────────
  GM_registerMenuCommand('ECOUNT 자격 설정', setupCreds);
  GM_registerMenuCommand('ECOUNT 자격 삭제', clearCreds);

  console.log('%c[ECOUNT Relay] v' + relay.__version + ' 설치됨 · 사용 가능', 'color:#0d9488;font-weight:bold');
  console.log('→ window.ecountRelay 로 접근. 자격이 비어있으면 화면 우측상단 Tampermonkey 아이콘 → "ECOUNT 자격 설정" 메뉴 실행.');
})();
