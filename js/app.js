/* 고객컨설팅 마스터과정 운영관리 - 메인 앱 로직 */

(function () {
  const LS_KEY = "cmf.filter.v1";
  const LS_DEFAULTS_KEY = "cmf.masterTargetDefaults.v1";
  const LS_AWARD_PLANS_KEY = "cmf.awardPlans.v1";
  const LS_TARGET_GOALS_KEY = "cmf.targetGoals.v1";
  const DEFAULT_MASTER_TARGET = 200000; // 원 (= 200,000원)
  const DEFAULT_REGION = "호남지역단";

  const DEFAULT_AWARD_PLAN = {
    title: "2026년 고객컨설팅 마스터과정 활성화 시상안",
    // 1. 개인순증시상
    personalIncr: {
      enabled: true,
      items: [
        { critVal: 50, payType: "pct",   payVal: 150, payRate: "full" },
        { critVal: 30, payType: "pct",   payVal: 120, payRate: "full" },
        { critVal: 20, payType: "fixed", payVal: 20,  payRate: "full" },
        { critVal: 10, payType: "fixed", payVal: 10,  payRate: "full" },
        { critVal: 5,  payType: "fixed", payVal: 5,   payRate: "full" }
      ]
    },
    // 2. 신장X TopN (slot 1) — 기본: 신장률
    topAward1: {
      enabled: true, type: "rate", n: 10,
      payouts: [30, 20, 20, 5, 5, 5, 5, 5, 5, 5],
      minNetEnabled: true, minNet: 300000
    },
    // 3. 신장X TopN (slot 2) — 기본: 신장액
    topAward2: {
      enabled: true, type: "amt", n: 10,
      payouts: [50, 30, 30, 10, 10, 10, 10, 10, 10, 10],
      minNetEnabled: true, minNet: 300000
    },
    bothNodup: true,
    groupAward1: { enabled: false, threshold: 5, payout: 5 },
    groupAward2: { enabled: false, rateThreshold: 110, payout: 15 },
    // 4. 기준조건 (eligibility)
    eligibility: {
      enabled: true,
      operator: "and",
      conditions: [
        { field: "converted", threshold: 80 }
      ]
    },
    // 5. 기타사항
    notes: "※ 환산실적 80만원 미만 시상제외 | 합산 하이캡 배수 15 미만시 50% 지급"
  };

  function getAwardPlan(region) {
    try {
      const stored = JSON.parse(localStorage.getItem(LS_AWARD_PLANS_KEY) || "{}");
      const saved = stored[region];
      if (!saved) return JSON.parse(JSON.stringify(DEFAULT_AWARD_PLAN));
      // 새 구조 + 기본값 병합 (필드 누락 안전)
      // 기존 플랜에 bothNodup이 없으면: 신장률·신장액 둘 다 활성화된 경우 자동으로 true 마이그레이션
      const _autoNodup = saved.bothNodup == null
        ? !!(saved.topAward1?.enabled && saved.topAward2?.enabled)
        : saved.bothNodup;
      return {
        title:         saved.title         ?? DEFAULT_AWARD_PLAN.title,
        personalIncr:  saved.personalIncr  ?? DEFAULT_AWARD_PLAN.personalIncr,
        topAward1:     saved.topAward1     ?? DEFAULT_AWARD_PLAN.topAward1,
        topAward2:     saved.topAward2     ?? DEFAULT_AWARD_PLAN.topAward2,
        bothNodup:     _autoNodup,
        groupAward1:   saved.groupAward1   ?? DEFAULT_AWARD_PLAN.groupAward1,
        groupAward2:   saved.groupAward2   ?? DEFAULT_AWARD_PLAN.groupAward2,
        eligibility:   saved.eligibility   ?? DEFAULT_AWARD_PLAN.eligibility,
        notes:         saved.notes         ?? DEFAULT_AWARD_PLAN.notes
      };
    } catch { return JSON.parse(JSON.stringify(DEFAULT_AWARD_PLAN)); }
  }

  async function saveAwardPlan(region, plan) {
    const stored = JSON.parse(localStorage.getItem(LS_AWARD_PLANS_KEY) || "{}");
    stored[region] = { ...plan, _ts: Date.now() };
    // localStorage에 먼저 저장 — Firestore 실패해도 데이터 보호
    localStorage.setItem(LS_AWARD_PLANS_KEY, JSON.stringify(stored));
    if (window.DataAPI?.saveAwardPlans) {
      await window.DataAPI.saveAwardPlans(stored); // 실패 시 호출자가 처리
    }
  }

  // Firestore ↔ localStorage 시상안 동기화 (앱 초기화 시 1회 호출)
  // 타임스탬프(_ts) 기반 last-write-wins: 더 최근에 저장된 버전이 항상 우선
  async function syncAwardPlansFromFirestore() {
    if (!window.DataAPI?.loadAwardPlans) return;
    try {
      const fsPlans = await window.DataAPI.loadAwardPlans();
      const local   = JSON.parse(localStorage.getItem(LS_AWARD_PLANS_KEY) || "{}");

      // 모든 키를 합집합으로 처리
      const allKeys = new Set([...Object.keys(local), ...(fsPlans ? Object.keys(fsPlans) : [])]);
      if (!allKeys.size) return;

      let localUpdated = false;
      let fsNeedsUpdate = false;
      const merged = {};

      allKeys.forEach((k) => {
        const lPlan = local[k];
        const fPlan = fsPlans?.[k];
        if (!lPlan && fPlan) {
          // Firestore에만 있는 키 (다른 기기에서 저장) → 로컬로 가져옴
          merged[k] = fPlan;
          localUpdated = true;
        } else if (lPlan && !fPlan) {
          // 로컬에만 있는 키 (Firestore에 아직 없음) → Firestore로 올림
          merged[k] = lPlan;
          fsNeedsUpdate = true;
        } else {
          // 양쪽 모두 있음: _ts가 더 큰 쪽 우선, 같으면 로컬 유지
          const lTs = lPlan._ts || 0;
          const fTs = fPlan._ts || 0;
          if (fTs > lTs) {
            merged[k] = fPlan;
            localUpdated = true;
          } else {
            merged[k] = lPlan;
            if (lTs > fTs) fsNeedsUpdate = true;
          }
        }
      });

      // 로컬 캐시 업데이트
      if (localUpdated || fsNeedsUpdate) {
        localStorage.setItem(LS_AWARD_PLANS_KEY, JSON.stringify(merged));
      }
      // 로컬에 최신 데이터가 있으면 Firestore도 업데이트
      if (fsNeedsUpdate && window.DataAPI?.saveAwardPlans) {
        window.DataAPI.saveAwardPlans(merged)
          .catch((e) => console.warn("[AwardPlan] Firestore 역동기화 실패:", e));
      }

      console.info(`[AwardPlan] 동기화 완료: ${allKeys.size}개 시상안`);
      // 로컬이 갱신된 경우 화면 재렌더링
      if (localUpdated) renderDebounced();
    } catch (e) {
      console.warn("[AwardPlan] Firestore 동기화 실패 (로컬 데이터 유지):", e);
    }
  }

  // 시상 payout 정규화: 구형(숫자) → {type:"cash",val:N}, 신형 그대로
  function normPayout(p) {
    if (p && typeof p === "object") return p;
    return { type: "cash", val: Number(p) || 0 };
  }
  // 시상 payout 표시 라벨
  function payoutLabel(p) {
    const np = normPayout(p);
    if (np.type === "item") return String(np.val || "물품");
    return `${np.val}만원`;
  }

  // 기준조건 체크용: 학생 metric 추출 (만원 단위 기준)
  function getStudentMetric(s, field, sfx) {
    if (field === "converted") {
      // 환산실적 = 현재실적 (스텝별 pgCurrent2/3… 우선, 없으면 pgCurrent, 없으면 current)
      const key = sfx ? `pgCurrent${sfx}` : null;
      const cur = (key && s[key] !== undefined)
        ? Number(s[key])
        : s.pgCurrent !== undefined ? Number(s.pgCurrent) : Number(s.current || 0);
      return cur / 10000;
    }
    if (field === "hiCap")     return Number(s.hiCap || 0);
    if (field === "monthly")   return Number(s.insAvg || 0) / 10000;
    return 0;
  }

  // 시상 자격 확인
  function isEligibleForAward(student, plan, sfx) {
    if (!plan?.eligibility?.enabled) return true;
    const conds = plan.eligibility.conditions || [];
    if (!conds.length) return true;
    const checks = conds.map((c) => getStudentMetric(student, c.field, sfx) > Number(c.threshold || 0));
    return plan.eligibility.operator === "or" ? checks.some(Boolean) : checks.every(Boolean);
  }

  // 개인순증 시상금 계산 (원 단위)
  function calcPersonalAward(stat, plan) {
    if (!plan?.personalIncr?.enabled) return 0;
    const items = (plan.personalIncr.items || []).filter((it) => Number(it.critVal) > 0 || Number(it.payVal) > 0);
    if (!items.length) return 0;
    const sorted = [...items].sort((a, b) => Number(b.critVal) - Number(a.critVal));
    const netManwon = Number(stat.net || 0) / 10000;
    for (const item of sorted) {
      if (netManwon >= Number(item.critVal)) {
        if (item.payType === "pct") return Math.round(Number(stat.net) * (Number(item.payVal) / 100));
        if (item.payType === "item") return 0;
        return Math.round(Number(item.payVal) * 10000);
      }
    }
    return 0;
  }

  // 순위별 시상금 (원 단위) — rank: 1-based
  function calcRankAward(rank, topConfig) {
    if (!topConfig?.enabled) return 0;
    if (rank > Number(topConfig.n || 0)) return 0;
    const payouts = topConfig.payouts || [];
    const idx = rank - 1;
    const v = payouts[idx];
    if (v === undefined || v === null) return 0;
    const np = normPayout(v);
    if (np.type === "item") return 0;
    return Math.round(Number(np.val) * 10000);
  }

  // 정렬: type "rate" or "amt"
  function sortStatsForType(stats, type) {
    return [...stats].sort((a, b) => {
      if (type === "rate") {
        const ra = a.base > 0 ? a.net / a.base : 0;
        const rb = b.base > 0 ? b.net / b.base : 0;
        return rb - ra;
      }
      return b.net - a.net;
    });
  }
  // 앱 버전 — 코드 수정(커밋)마다 0.01 씩 증가
  const APP_VERSION = "2.03";

  // 실적진도현황 열 매핑 — 저장 필드 선택지
  const PG_FIELD_OPTIONS = [
    { value: "empNo",       label: "사원번호" },
    { value: "region",      label: "지역단" },
    { value: "center",      label: "비전센터" },
    { value: "branch",      label: "지점" },
    { value: "name",        label: "성명" },
    { value: "pgMonth",     label: "차월" },
    { value: "pgLeader",    label: "육성리더" },
    { value: "pgPreIns",    label: "직전6개월인보험" },
    { value: "pgPreConv",   label: "직전6개월환산" },
    { value: "pgPreIncome", label: "직전6개월육성소득" },
    { value: "pgBase",      label: "기준실적" },
    { value: "pgCurrent",   label: "현재실적" },
    { value: "pgIpumCount",  label: "인품건수" },
    { value: "pgIpumAmt",   label: "인품실적" },
    { value: "ignore",      label: "— 무시 (저장 안 함) —" },
  ];
  // 헤더 텍스트 → 필드 자동 매핑
  const PG_HEADER_AUTOMAP = {
    "지역단": "region", "비전센터": "center", "지점": "branch",
    "사원번호": "empNo", "사번": "empNo",
    "성명": "name", "차월": "pgMonth", "육성리더": "pgLeader",
    "직전6개월인보험": "pgPreIns", "직전6개월환산": "pgPreConv", "직전6개월육성소득": "pgPreIncome",
    "기준실적": "pgBase", "현재실적": "pgCurrent",
    "달성률": "ignore", "달성율": "ignore", "순증실적": "ignore", "순증": "ignore",
    "대리점명": "ignore",
    "인품건수": "pgIpumCount", "인붐건수": "pgIpumCount", "계약건수": "pgIpumCount",
    "인품실적": "pgIpumAmt",  "실적": "pgIpumAmt",
    "위촉차월": "pgMonth",
    "Step1현재실적": "pgCurrent",  "Step2현재실적": "pgCurrent2",
    "Step1인품건수": "pgIpumCount", "Step2인품건수": "pgIpumCount2",
    "Step1인품실적": "pgIpumAmt",  "Step2인품실적": "pgIpumAmt2",
    "Step1달성률": "ignore", "Step2달성률": "ignore",
    "Step1순증실적": "ignore", "Step2순증실적": "ignore",
    "시상금": "ignore",
  };
  // 교육생 등록 일괄입력 헤더 자동 매핑
  const BULK_HEADER_MAP = {
    "지역단": "region",
    "비전센터": "center", "비전센터명": "center",
    "지점": "branch", "지점명": "branch",
    "기수": "cohort", "과정기수": "cohort",
    "사번": "empNo", "사원번호": "empNo", "직원번호": "empNo",
    "성명": "name", "이름": "name",
    "연락처": "phone", "전화번호": "phone", "휴대폰": "phone", "핸드폰": "phone",
    "기준실적": "base", "기준실적(원)": "base",
    "마스터목표": "target", "마스터목표(원)": "target",
    "아너스목표": "honors", "아너스목표(원)": "honors",
    "차월": "tenureMonths", "위촉차월": "tenureMonths",
  };
  // 헤더 없을 때 기본 열 순서 (표준 16열 기준)
  const PG_DEFAULT_COLS = [
    "region","center","branch","empNo","name","pgMonth","pgLeader",
    "pgPreIns","pgPreConv","pgPreIncome","pgBase","pgCurrent","ignore","ignore","pgIpumCount","pgIpumAmt",
  ];
  // Step1 13열 단축 형식: 대리점명 포함 (헤더 없이 붙여넣기 시)
  // 지역단·비전센터·지점·사번·대리점명·성명·위촉차월·기준실적·현재실적·달성률·순증·인품건수·인품실적
  const PG_STEP1_SHORT_COLS = [
    "region","center","branch","empNo","ignore","name","pgMonth",
    "pgBase","pgCurrent","ignore","ignore","pgIpumCount","pgIpumAmt",
  ];
  // Step2 복합 형식: 기준실적(A) + Step2현재실적(C) · Step2인품 저장 / step1 데이터는 무시
  // 순서: 지역단·비전센터·지점·사번·대리점명·성명·위촉차월·기준실적
  //       step1현재실적~step1인품실적(5열 무시)·step2현재실적·step2달성률·step2순증(무시)
  //       step1대비달성률·step1대비순증(무시)·step2인품건수·step2인품실적·시상금(무시)
  const PG_STEP2_COMBINED_COLS = [
    "region","center","branch","empNo","ignore","name","pgMonth",
    "pgBase","ignore","ignore","ignore","ignore","ignore",
    "pgCurrent2","ignore","ignore","ignore","ignore","pgIpumCount2","pgIpumAmt2","ignore",
  ];
  // 스텝1·스텝2 공통 10열 형식 — 지역단·비전센터·지점·사번·성명·위촉차월·기준실적·현재실적·계약건수·실적
  // (구포맷 대리점명 제거, 현재실적 추가)
  const PG_STEP_UNIFIED_COLS = [
    "region","center","branch","empNo","name","pgMonth",
    "pgBase","pgCurrent","pgIpumCount","pgIpumAmt",
  ];

  // 상담고객 태그 선택지
  const CT = ["신규", "기존", "DB", "개척", "소개"];         // 고객유형 (단일)
  const CS = ["관계형성", "보장분석", "리모델링"];            // 상담단계 (단일)
  const MT = ["스마트제안서", "메디컬보장분석", "행복보장분석"]; // 활용자료 (복수)
  const AM = ["5만원↑", "10만원↑", "15만원↑", "20만원↑"];  // 제안금액 (단일)
  const BJ = ["건강체", "간편", "어린이", "운전자", "기타", "재물"]; // 보종 (단일)

  // 시상 계산기 상수 ('26년 2분기 매출아너스 기준)
  const HONORS = [
    { grade: "서밋 (Summit)",               criteria: "500만↑", critVal: 500, prize: 500 },
    { grade: "로얄 마스터 (Royal master)",  criteria: "400만↑", critVal: 400, prize: 400 },
    { grade: "엘리트 마스터 (Elite master)",criteria: "300만↑", critVal: 300, prize: 250 },
    { grade: "마스터 (Master)",             criteria: "200만↑", critVal: 200, prize: 150 },
    { grade: "프레스티지Ⅱ (Prestige Ⅱ)",  criteria: "150만↑", critVal: 150, prize: 100 },
    { grade: "프레스티지Ⅰ (Prestige Ⅰ)",  criteria: "100만↑", critVal: 100, prize: 70  },
    { grade: "프라임Ⅱ (Prime Ⅱ)",          criteria: "70만↑",  critVal: 70,  prize: 50  },
    { grade: "프라임Ⅰ (Prime Ⅰ)",          criteria: "50만↑",  critVal: 50,  prize: 30  },
    { grade: "프로 (Pro)",                  criteria: "30만↑",  critVal: 30,  prize: 20  }
  ];
  const INCR_CFG = { base: 5, rate: 50, mcap: 20, qcap: 50 };
  const MASTER_AWARD = [
    { criteria: "순증 50만원↑", critVal: 50, type: "pct",   val: 150, label: "순증×150%" },
    { criteria: "순증 30만원↑", critVal: 30, type: "pct",   val: 120, label: "순증×120%" },
    { criteria: "순증 20만원↑", critVal: 20, type: "fixed", val: 20,  label: "20만원 고정" },
    { criteria: "순증 10만원↑", critVal: 10, type: "fixed", val: 10,  label: "10만원 고정" },
    { criteria: "순증 5만원↑",  critVal: 5,  type: "fixed", val: 5,   label: "5만원 고정"  }
  ];
  const AWARD_POSITIVE_WORDS = [
    "희망","열정","도전","성공","빛남","에너지","가능성","미래","성장","기적",
    "행운","자신감","결실","최고","의지","기회","열매","승리","확신","번영",
    "영광","동력","추진력","원동력","변화","핵심","자랑","기대","보람","축복"
  ];

  function loadFilter() {
    const base = { region: DEFAULT_REGION, center: "", branch: "", cohort: "", step: "1", q: "" };
    try {
      const saved = JSON.parse(localStorage.getItem(LS_KEY) || "null");
      if (saved && typeof saved === "object") {
        return { ...base, ...saved, q: "" };
      }
    } catch (e) {}
    return base;
  }

  function persistFilter() {
    const { region, center, branch, cohort, step } = state.filter;
    try {
      localStorage.setItem(LS_KEY, JSON.stringify({ region, center, branch, cohort, step }));
    } catch (e) {}
  }

  const state = {
    students: [],
    filter: loadFilter(),
    form: { region: "", center: "", branch: "" },
    selectedEmpNo: null,
    consultations: [],
    consultUnsub: null,
    // Phase 1 interview form
    tgtAutoMode: true,
    // 교육생 등록 폼 마스터목표 팝업 선택값: null=수동입력, number=평균+N(천원)
    formTgtAddAmount: null,
    editingConsultId: null,  // Phase 4용 예약 (이번엔 사용 안 함)
    lastDetailEmpNo: null,   // 마지막으로 완전 렌더한 교육생 (폼 보존용)
    // Phase 2 clients
    crData: [],              // 현재 폼의 상담고객 배열 (최대 5)
    // Phase 3 시상 계산기
    calcOpen: true,          // 계산기 접힘/펼침 상태
    calcTgtUserEditing: false, // 희망목표 직접입력 중 플래그
    lastCalcResult: null,    // 마지막 계산 결과 (시상인쇄용)
    errorReports: [],
    errorReportUnsub: null,
    // 동기화 상태
    studentsLoaded: false,   // Firestore 첫 응답 여부
    syncMeta: { fromCache: false },
    // 교육생 패널 서브뷰 (form | history | print)
    studentSubView: "form",
    // 교육생 개별 선택 삭제 모달 상태
    sdSelected: new Set(),
    // 좌측 사이드바 — 펼쳐진 비전센터(기본 모두 펼침) / 지점(기본 모두 접힘)
    openCenters: new Set(),
    openBranches: new Set(),
    // 모바일 사이드바 열림 여부
    mobileSidebarOpen: false,
    // 실적진도 패널 상태
    progressRegion: "",
    progressSubTab: "home",
    progressYear: String(new Date().getFullYear()),
    progressCohort: "",
    progressStep: "",
    // 출력 서브뷰
    printMode: "personal",     // 'personal' | 'branch' | 'vc'
    printConsultCache: {}      // { empNo: consultations[] } — 다건 출력시 캐시
  };

  // ========== 유틸 ==========
  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => document.querySelectorAll(sel);

  function toast(msg, type = "") {
    const el = $("#toast");
    el.textContent = msg;
    el.className = "toast " + type;
    el.hidden = false;
    clearTimeout(toast._t);
    toast._t = setTimeout(() => (el.hidden = true), 2200);
  }

  function openModal(id) { $(id).hidden = false; }
  function closeModal(id) { $(id).hidden = true; }

  function syncOrgLabels() {
    syncFilterOrgSelects();
    syncFormOrgSelects();
  }

  // ========== 사이드바 필터 인라인 셀렉트 ==========
  function populateFilterRegionSelect() {
    const sel = $("#filter-region-select");
    if (!sel) return;
    const data = window.ORG_DATA;
    sel.innerHTML = data.regions
      .map((r) => `<option value="${escapeHtml(r.name)}">${escapeHtml(r.name)}</option>`)
      .join("");
  }

  function populateFilterCenterSelect() {
    const sel = $("#filter-center-select");
    if (!sel) return;
    const data = window.ORG_DATA;
    const reg = state.filter.region ? data.regions.find((r) => r.name === state.filter.region) : null;
    if (!reg) {
      sel.innerHTML = `<option value="">전체</option>`;
      sel.disabled = true;
      return;
    }
    sel.disabled = false;
    sel.innerHTML = `<option value="">전체</option>` +
      reg.centers.map((c) => `<option value="${escapeHtml(c.name)}">${escapeHtml(c.name)}</option>`).join("");
  }

  function populateFilterBranchSelect() {
    const sel = $("#filter-branch-select");
    if (!sel) return;
    const data = window.ORG_DATA;
    const reg = state.filter.region ? data.regions.find((r) => r.name === state.filter.region) : null;
    const ctr = reg && state.filter.center ? reg.centers.find((c) => c.name === state.filter.center) : null;
    if (!ctr) {
      sel.innerHTML = `<option value="">전체</option>`;
      sel.disabled = true;
      return;
    }
    sel.disabled = false;
    sel.innerHTML = `<option value="">전체</option>` +
      ctr.branches.map((b) => `<option value="${escapeHtml(b)}">${escapeHtml(b)}</option>`).join("");
  }

  function syncFilterOrgSelects() {
    populateFilterCenterSelect();
    populateFilterBranchSelect();
    const r = $("#filter-region-select"); if (r) r.value = state.filter.region || "";
    const c = $("#filter-center-select"); if (c) c.value = state.filter.center || "";
    const b = $("#filter-branch-select"); if (b) b.value = state.filter.branch || "";
  }

  // ========== 단건 입력 폼의 지역단/비전센터/지점 select ==========
  function populateRegionSelect() {
    const sel = $("#form-region-select");
    if (!sel) return;
    const data = window.ORG_DATA;
    sel.innerHTML = `<option value="">선택</option>` +
      data.regions.map((r) => `<option value="${escapeHtml(r.name)}">${escapeHtml(r.name)}</option>`).join("");
  }

  function populateCenterSelect() {
    const sel = $("#form-center-select");
    if (!sel) return;
    const data = window.ORG_DATA;
    const reg = state.form.region ? data.regions.find((r) => r.name === state.form.region) : null;
    if (!reg) {
      sel.innerHTML = `<option value="">선택</option>`;
      sel.disabled = true;
      return;
    }
    sel.disabled = false;
    sel.innerHTML = `<option value="">선택</option>` +
      reg.centers.map((c) => `<option value="${escapeHtml(c.name)}">${escapeHtml(c.name)}</option>`).join("");
  }

  function populateBranchSelect() {
    const sel = $("#form-branch-select");
    if (!sel) return;
    const data = window.ORG_DATA;
    const reg = state.form.region ? data.regions.find((r) => r.name === state.form.region) : null;
    const ctr = reg && state.form.center ? reg.centers.find((c) => c.name === state.form.center) : null;
    if (!ctr) {
      sel.innerHTML = `<option value="">선택</option>`;
      sel.disabled = true;
      return;
    }
    sel.disabled = false;
    sel.innerHTML = `<option value="">선택</option>` +
      ctr.branches.map((b) => `<option value="${escapeHtml(b)}">${escapeHtml(b)}</option>`).join("");
  }

  function syncFormOrgSelects() {
    populateCenterSelect();
    populateBranchSelect();
    const r = $("#form-region-select"); if (r) r.value = state.form.region || "";
    const c = $("#form-center-select"); if (c) c.value = state.form.center || "";
    const b = $("#form-branch-select"); if (b) b.value = state.form.branch || "";
  }

  // ========== 필터링 / 렌더링 ==========
  function filteredStudents() {
    const f = state.filter;
    const q = (f.q || "").trim().toLowerCase();
    return state.students.filter((s) => {
      if (f.region && s.region !== f.region) return false;
      if (f.center && s.center !== f.center) return false;
      if (f.branch && s.branch !== f.branch) return false;
      if (f.cohort && s.cohort && s.cohort !== f.cohort) return false;
      if (q) {
        const hay = [s.empNo, s.name, s.branch, s.center, s.region].join(" ").toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
  }

  function renderKPIs(list) {
    const sum = (k) => list.reduce((a, s) => a + (Number(s[k]) || 0), 0);
    $("#kpi-total").textContent = list.length.toLocaleString();
    $("#kpi-base").textContent = list.reduce((a, s) => a + Number(s.base || 0), 0).toLocaleString();
    $("#kpi-target").textContent = sum("target").toLocaleString();
    $("#kpi-honors").textContent = sum("honors").toLocaleString();
    $("#mini-total").textContent = list.length;
    const ptRegion = document.getElementById("page-title-region");
    if (ptRegion) {
      const r = state.filter.region;
      const c = state.filter.cohort;
      const st = state.filter.step || "1";
      const stepSuffix = ` (Step ${st})`;
      ptRegion.textContent = r ? ` — ${r} ${c || "전체"}${stepSuffix}` : "";
    }
  }

  // 사이드바 지점별 교육생 명단 — 클릭하면 면담관리 패널로 이동
  function renderSidebarStudentList(list) {
    const container = $("#sidebar-student-list");
    if (!container) return;
    if (list.length === 0) {
      let msg, showRetry = false;
      if (!state.studentsLoaded) {
        msg = "🔄 Firebase 연결 중...";
      } else if (state.students.length === 0) {
        if (state.syncMeta.fromCache) {
          msg = "❗ 서버에서 데이터를 받지 못했습니다.<br>WiFi/다른 브라우저로 시도하거나 새로고침하세요.";
          showRetry = true;
        } else {
          msg = "ℹ️ 등록된 교육생이 없습니다.";
        }
      } else {
        msg = "조건에 맞는 교육생 없음";
      }
      container.innerHTML = `<div class="empty-mini">${msg}${showRetry ? `
        <div style="margin-top:10px;display:flex;flex-direction:column;gap:6px">
          <button class="btn-outline small" id="btn-retry-sync">🔄 재시도</button>
          <button class="btn-outline small" id="btn-clear-cache">🗑 캐시 지우고 새로고침</button>
        </div>` : ""}</div>`;
      if (showRetry) {
        const r = $("#btn-retry-sync");
        if (r) r.addEventListener("click", retrySubscription);
        const c = $("#btn-clear-cache");
        if (c) c.addEventListener("click", clearCacheAndReload);
      }
      return;
    }
    // 2단계 그룹: 비전센터 → 지점 → 교육생
    const centerGroups = {};
    list.forEach((s) => {
      const c = s.center || "(비전센터 미지정)";
      const b = s.branch || "(지점 미지정)";
      if (!centerGroups[c]) centerGroups[c] = {};
      if (!centerGroups[c][b]) centerGroups[c][b] = [];
      centerGroups[c][b].push(s);
    });

    // 선택한 교육생이 속한 지점/센터는 자동 펼침
    const selectedStu = state.selectedEmpNo ? list.find((s) => s.empNo === state.selectedEmpNo) : null;
    const selectedBranch = selectedStu ? selectedStu.branch : null;
    const selectedCenter = selectedStu ? selectedStu.center : null;

    // 신규 비전센터는 기본 펼침 (Set 에 없으면 추가)
    Object.keys(centerGroups).forEach((c) => {
      if (!state.openCenters.has(c) && !state._centerSeen?.has(c)) {
        state.openCenters.add(c);
      }
    });
    state._centerSeen = new Set(Object.keys(centerGroups));

    container.innerHTML = Object.keys(centerGroups).sort().map((center) => {
      const branches = centerGroups[center];
      // 센터 합계 계산
      let centerTotal = 0, centerInterviewed = 0;
      Object.values(branches).forEach((arr) => {
        centerTotal += arr.length;
        centerInterviewed += arr.filter((s) => Number(s.consultCount || 0) > 0).length;
      });
      const centerOpen = state.openCenters.has(center) || center === selectedCenter;
      const branchHtml = Object.keys(branches).sort().map((branch) => {
        const rows = branches[branch].slice().sort((a, b) => (a.name || "").localeCompare(b.name || ""));
        const interviewed = rows.filter((s) => Number(s.consultCount || 0) > 0).length;
        const bOpen = state.openBranches.has(branch) || branch === selectedBranch;
        const isBranchUnassigned = branch === "(지점 미지정)";
        return `
          <details class="branch-mini" data-branch="${escapeHtml(branch)}"${bOpen ? " open" : ""}>
            <summary class="branch-mini-head">
              <span class="branch-name">${escapeHtml(branch)}</span>
              ${isBranchUnassigned ? `<button class="branch-assign-btn" data-center="${escapeHtml(center)}" title="지점 일괄 지정">📌 지정</button>` : ""}
              <span class="branch-cnt" title="교육생 / 면담된 교육생">${rows.length}/<em>${interviewed}</em></span>
              <span class="branch-chev">▾</span>
            </summary>
            <ul class="student-mini-list">
              ${rows.map((s) => {
                const nm = s.name || "(이름 미입력)";
                const initial = (s.name || "?").trim().charAt(0) || "?";
                const cc = Number(s.consultCount || 0);
                const ccBadge = cc > 0 ? `<span class="s-ivcnt" title="면담 기록 ${cc}회">${cc}회</span>` : "";
                return `
                <li class="${state.selectedEmpNo === s.empNo ? "selected" : ""}" data-emp="${escapeHtml(s.empNo)}" data-initial="${escapeHtml(initial)}">
                  <span class="s-name-wrap">
                    <span class="s-name">${escapeHtml(nm)}</span>
                    ${s.phone ? `<a href="tel:${escapeHtml(s.phone)}" class="s-phone s-phone-link" onclick="event.stopPropagation()">${escapeHtml(s.phone)}</a>` : `<span class="s-phone"></span>`}
                  </span>
                  ${ccBadge}
                </li>
              `;}).join("")}
            </ul>
          </details>
        `;
      }).join("");
      const isUnassigned = center === "(비전센터 미지정)";
      return `
        <details class="center-mini" data-center="${escapeHtml(center)}"${centerOpen ? " open" : ""}>
          <summary class="center-mini-head">
            <span class="center-name">🏢 ${escapeHtml(center)}</span>
            ${isUnassigned ? `<button class="center-assign-btn" title="비전센터 일괄 지정">📌 지정</button>` : ""}
            <span class="center-cnt" title="교육생 / 면담된 교육생">${centerTotal}/<em>${centerInterviewed}</em></span>
            <span class="center-chev">▾</span>
          </summary>
          <div class="center-branches">${branchHtml}</div>
        </details>
      `;
    }).join("");

    container.querySelectorAll("li[data-emp]").forEach((li) => {
      li.addEventListener("click", () => selectStudent(li.dataset.emp));
    });
    container.querySelectorAll("details.branch-mini").forEach((d) => {
      d.addEventListener("toggle", () => {
        const name = d.dataset.branch;
        if (!name) return;
        if (d.open) state.openBranches.add(name);
        else state.openBranches.delete(name);
      });
    });
    container.querySelectorAll("details.center-mini").forEach((d) => {
      d.addEventListener("toggle", () => {
        const name = d.dataset.center;
        if (!name) return;
        if (d.open) state.openCenters.add(name);
        else state.openCenters.delete(name);
      });
    });

    // "(지점 미지정)" 일괄 지정 버튼
    container.querySelectorAll(".branch-assign-btn").forEach((btn) => {
      btn.addEventListener("click", async (e) => {
        e.preventDefault();
        e.stopPropagation();
        const centerName = btn.dataset.center;
        const unassigned = list.filter((s) => s.center === centerName && !s.branch);
        if (!unassigned.length) { toast("미지정 교육생이 없습니다.", ""); return; }
        const region = state.filter.region;
        const orgRegion = window.ORG_DATA?.regions?.find((r) => r.name === region);
        const orgCenter = orgRegion?.centers?.find((c) => c.name === centerName);
        const branches = orgCenter ? orgCenter.branches : [];
        if (!branches.length) { toast("이 비전센터에 등록된 지점 정보가 없습니다.", "error"); return; }
        openPickerModal(
          `지점 선택 — ${centerName} 미지정 ${unassigned.length}명 일괄 지정`,
          branches,
          async (picked) => {
            const ok = await openConfirmModal(`"${picked}"으로\n${unassigned.length}명을 일괄 저장하시겠습니까?`);
            if (!ok) return;
            const records = unassigned.map((s) => ({ ...s, branch: picked }));
            try {
              if (typeof window.DataAPI.saveMany === "function") await window.DataAPI.saveMany(records);
              else for (const r of records) await window.DataAPI.save(r);
              toast(`✅ ${unassigned.length}명의 지점을 "${picked}"으로 저장했습니다.`, "");
            } catch (err) { toast("저장 오류: " + err.message, "error"); }
          }
        );
      });
    });

    // "(비전센터 미지정)" 일괄 지정 버튼
    container.querySelectorAll(".center-assign-btn").forEach((btn) => {
      btn.addEventListener("click", async (e) => {
        e.preventDefault();
        e.stopPropagation();
        const unassigned = list.filter((s) => !s.center);
        if (!unassigned.length) { toast("미지정 교육생이 없습니다.", ""); return; }
        const region = state.filter.region;
        const orgRegion = window.ORG_DATA?.regions?.find((r) => r.name === region);
        const centers = orgRegion ? orgRegion.centers.map((c) => c.name) : [];
        if (!centers.length) { toast("이 지역단에 등록된 비전센터 정보가 없습니다.", "error"); return; }

        // 지점 → 비전센터 역방향 맵
        const branchCenterMap = {};
        if (orgRegion) {
          for (const c of orgRegion.centers)
            for (const b of c.branches) branchCenterMap[b] = c.name;
        }
        const autoMatchable = unassigned.filter((s) => s.branch && branchCenterMap[s.branch]);
        const AUTO_OPT = autoMatchable.length
          ? `🔄 지점 기준 자동 매칭 (${autoMatchable.length}명)`
          : null;

        openPickerModal(
          `비전센터 선택 — ${unassigned.length}명 일괄 지정`,
          AUTO_OPT ? [AUTO_OPT, ...centers] : centers,
          async (picked) => {
            if (picked === AUTO_OPT) {
              const groups = {};
              for (const s of autoMatchable) {
                const c = branchCenterMap[s.branch];
                (groups[c] = groups[c] || []).push(s.name);
              }
              const summary = Object.entries(groups)
                .map(([c, names]) => `• ${c}: ${names.length}명`)
                .join("\n");
              const ok = await openConfirmModal(`지점 기준 자동 배정:\n${summary}\n\n총 ${autoMatchable.length}명을 배정하시겠습니까?`);
              if (!ok) return;
              const records = autoMatchable.map((s) => ({ ...s, center: branchCenterMap[s.branch] }));
              try {
                if (typeof window.DataAPI.saveMany === "function") await window.DataAPI.saveMany(records);
                else for (const r of records) await window.DataAPI.save(r);
                toast(`✅ ${autoMatchable.length}명을 지점 기준으로 자동 배정했습니다.`, "");
              } catch (err) { toast("저장 오류: " + err.message, "error"); }
              return;
            }
            const ok = await openConfirmModal(`"${picked}"으로\n${unassigned.length}명을 일괄 저장하시겠습니까?`);
            if (!ok) return;
            const records = unassigned.map((s) => ({ ...s, center: picked }));
            try {
              if (typeof window.DataAPI.saveMany === "function") await window.DataAPI.saveMany(records);
              else for (const r of records) await window.DataAPI.save(r);
              toast(`✅ ${unassigned.length}명의 비전센터를 "${picked}"으로 저장했습니다.`, "");
            } catch (err) { toast("저장 오류: " + err.message, "error"); }
          }
        );
      });
    });
  }

  // ========== 교육생 선택 → 면담 관리 ==========
  function selectStudent(empNo) {
    state.selectedEmpNo = empNo;
    // 기존 구독 해제
    if (state.consultUnsub) { state.consultUnsub(); state.consultUnsub = null; }
    state.consultations = [];
    renderSidebarStudentList(filteredStudents());
    renderStudentDetail();
    // 교육생 선택 시 교육생 관리 패널로 전환 (모바일은 사이드바도 닫기)
    if (isMobileViewport()) closeMobileSidebar();
    switchView("#students");
    // 면담 기록 실시간 구독
    if (window.DataAPI && typeof window.DataAPI.subscribeConsultations === "function") {
      state.consultUnsub = window.DataAPI.subscribeConsultations(empNo, (list) => {
        state.consultations = list;
        renderConsultations();
        // 차수/ins/tgt 자동채움 재실행 (빈 필드만 보정)
        const s = state.students.find((x) => x.empNo === empNo);
        if (s) {
          autoFillInterviewForm(s);
          // 자기치유: 저장된 consultCount 가 실제 목록 길이와 다르면 동기화
          const stored = Number(s.consultCount || 0);
          if (stored !== list.length && typeof window.DataAPI.syncConsultCount === "function") {
            window.DataAPI.syncConsultCount(empNo, list.length).catch(() => {});
          }
        }
      });
    }
  }

  function renderStudentDetail() {
    const body = $("#student-detail-body");
    const title = $("#detail-title");
    if (!body) return;
    const s = state.students.find((x) => x.empNo === state.selectedEmpNo);
    if (!s) {
      title.textContent = "교육생 면담 관리";
      body.innerHTML = `<div class="empty-state">좌측 [지점별 교육생] 목록에서 교육생을 선택하면 면담 관리를 시작할 수 있습니다.</div>`;
      state.lastDetailEmpNo = null;
      return;
    }
    title.textContent = `${s.name || "(이름 없음)"} — 면담 관리`;
    // 같은 학생 재렌더 요청이면 폼 입력 보존 (학생 정보 한 줄만 갱신)
    if (state.lastDetailEmpNo === s.empNo && document.getElementById("iv-coach")) {
      updateStudentInfoBar(s);
      renderConsultations();
      return;
    }
    state.lastDetailEmpNo = s.empNo;
    const sub = state.studentSubView;
    const validSub = ["form", "history", "print"].includes(sub) ? sub : "form";
    body.innerHTML = `
      <div class="detail-stack">
        ${renderStudentInfoBarHtml(s)}
        <div class="sub-view" data-sub="form" ${validSub !== "form" ? "hidden" : ""}>
          ${renderInterviewFormHtml(s)}
        </div>
        <div class="sub-view" data-sub="history" ${validSub !== "history" ? "hidden" : ""}>
          <div class="detail-card history-card">
            <h3>면담 이력 상세</h3>
            <div id="hist-list" class="hist-list">
              <div class="empty-mini">면담 기록 불러오는 중...</div>
            </div>
          </div>
        </div>
        <div class="sub-view" data-sub="print" ${validSub !== "print" ? "hidden" : ""}>
          ${renderPrintPanelHtml()}
        </div>
      </div>
    `;

    $("#btn-detail-edit").addEventListener("click", () => openEditForm(s.empNo));
    $("#btn-detail-del").addEventListener("click", () => removeStudent(s.empNo));
    // form 서브뷰일 때만 인풋 관련 바인딩
    if (state.studentSubView === "form") {
      bindInterviewFormEvents();
      autoFillInterviewForm(s);
    }
    // print 서브뷰일 때 출력 제어 바인딩
    if (state.studentSubView === "print") {
      bindPrintControls();
      renderPrintView();
    }
    renderConsultations();
    bindSubTabs();
  }

  // 서브탭 [면담관리]/[면담이력] 바인딩 + 갯수 배지
  function bindSubTabs() {
    document.querySelectorAll("#student-sub-tabs .sub-tab").forEach((btn) => {
      btn.classList.toggle("active", btn.dataset.sub === state.studentSubView);
      btn.onclick = () => setStudentSubView(btn.dataset.sub);
    });
    const cnt = document.getElementById("hist-cnt");
    if (cnt) cnt.textContent = state.consultations.length;
  }

  function setStudentSubView(sub) {
    if (state.studentSubView === sub) return;
    state.studentSubView = sub;
    state.lastDetailEmpNo = null; // 강제 재렌더
    renderStudentDetail();
  }

  // 교육생 한 줄 정보 바 (이름/사번/지점/연락처/기수/실적 + 액션)
  function renderStudentInfoBarHtml(s) {
    const initial = (s.name || "?").trim().charAt(0) || "?";
    const fmt = (n) => Number(n || 0).toLocaleString();
    return `
      <div class="student-info-bar" id="student-info-bar">
        <div class="sib-avatar">${escapeHtml(initial)}</div>
        <div class="sib-main">
          <div class="sib-name">${escapeHtml(s.name || "")}<span class="sib-emp">${escapeHtml(s.empNo)}</span></div>
          <div class="sib-meta">${[s.center, s.branch, s.cohort].filter(Boolean).map(escapeHtml).join(" · ")}${s.phone ? ` · <a href="tel:${escapeHtml(s.phone)}" class="sib-phone-link">${escapeHtml(s.phone)}</a>` : ""}${s.team ? ` <span class="sib-team">${escapeHtml(String(s.team))}팀</span>` : ""}</div>
        </div>
        <div class="sib-stats">
          <div><span>기준실적</span><strong>${fmt(Number(s.base || 0))}</strong></div>
          <div><span>마스터목표</span><strong>${Number(s.target) > 0 ? fmt(Number(s.target)) : "<span style='color:#aaa'>미설정</span>"}</strong></div>
          <div><span>아너스목표</span><strong>${fmt(Number(s.honors))}</strong></div>
        </div>
        <div class="sib-actions">
          <button class="btn-outline small" id="btn-detail-edit">정보 수정</button>
          <button class="btn-outline small danger" id="btn-detail-del">삭제</button>
        </div>
      </div>
    `;
  }

  function updateStudentInfoBar(s) {
    const bar = document.getElementById("student-info-bar");
    if (!bar) return;
    bar.outerHTML = renderStudentInfoBarHtml(s);
    // 새로 그려졌으니 액션 버튼 재바인딩
    const editBtn = document.getElementById("btn-detail-edit");
    const delBtn = document.getElementById("btn-detail-del");
    if (editBtn) editBtn.addEventListener("click", () => openEditForm(s.empNo));
    if (delBtn) delBtn.addEventListener("click", () => removeStudent(s.empNo));
  }

  // ========== 면담 입력 폼 (rich) ==========
  function renderInterviewFormHtml(s) {
    const today = new Date().toISOString().slice(0, 10);
    return `
      <div class="detail-card iv-form">
        <div class="iv-header">
          <div class="iv-title" id="iv-title">MASTER과정 면담일지 (&nbsp;차 / 활동점검)</div>
          <div class="iv-sub">주간 활동 점검 및 코칭 기록</div>
        </div>

        <div class="iv-grid-3">
          <div class="iv-field">
            <label>면담일시 <em>*</em></label>
            <input type="date" id="iv-date" value="${today}">
          </div>
          <div class="iv-field">
            <label>차수</label>
            <input type="text" id="iv-seq" placeholder="예: 1" inputmode="numeric">
          </div>
          <div class="iv-field">
            <label>교육생 성명</label>
            <input type="text" id="iv-name" value="${escapeHtml(s.name || "")}" readonly>
          </div>

          <div class="iv-field">
            <label>현재실적</label>
            <div class="iv-curAct-row">
              <input type="number" id="iv-curAct" placeholder="원" step="1000">
              <button type="button" class="btn-iv-cur-upd" id="btn-iv-curAct-update" title="현재실적만 실적진도현황에 저장">업데이트</button>
            </div>
          </div>
          <div class="iv-field">
            <label>진도 <span class="iv-hint" id="iv-pct-hint"></span></label>
            <input type="number" id="iv-pct" placeholder="%">
          </div>
          <div class="iv-field">
            <label>조(팀) 번호</label>
            <input type="number" id="iv-team" placeholder="예: 1" min="1" value="${escapeHtml(s.team ? String(s.team) : "")}">
          </div>

          <div class="iv-field">
            <label>가입설계</label>
            <input type="number" id="iv-plan" placeholder="건">
          </div>

          <div class="iv-field">
            <label>행복보장분석</label>
            <input type="number" id="iv-hap" placeholder="건">
          </div>
          <div class="iv-field">
            <label>주간예상실적</label>
            <input type="number" id="iv-exp" placeholder="원" step="1000">
          </div>
          <div class="iv-field">
            <label>1차 마감 실적 <span class="iv-hint">중간점검</span></label>
            <input type="number" id="iv-close1" placeholder="원" step="1000">
          </div>

          <div class="iv-field">
            <label>2차 마감 실적 <span class="iv-hint">최종</span></label>
            <input type="number" id="iv-close2" placeholder="원" step="1000">
          </div>
          <div class="iv-field"></div>
        </div>

        <div class="iv-clients">
          <div class="iv-clients-head">
            <span class="iv-clients-title">주간 활동 점검 — 상담고객</span>
            <button type="button" class="btn-outline small" id="btn-cr-add">+ 고객 추가 (최대 5)</button>
          </div>
          <div id="cr-rows"></div>
        </div>

        <div class="iv-field iv-coach">
          <label>핵심 코칭포인트 / 후속조치 / 다음주 계획</label>
          <textarea id="iv-coach" rows="5" placeholder="핵심 코칭포인트, 후속조치, 다음주 계획을 상세히 기록하세요"></textarea>
        </div>

        <div class="iv-actions iv-actions-top">
          <button class="btn-outline" id="btn-iv-cancel-edit-top" hidden>✕ 수정 취소</button>
          <button class="btn-primary" id="btn-iv-save-top">💾 저장</button>
        </div>

        <div class="iv-calc">
          <div class="iv-calc-head" id="btn-calc-toggle">
            <span class="iv-calc-title">📊 시상 계산기 — '26년 2분기 매출아너스</span>
            <span class="iv-calc-icon" id="calc-toggle-icon">▾</span>
          </div>
          <div class="iv-calc-body" id="calc-section" style="display:block">
            <div class="iv-grid-3">
              <div class="iv-field">
                <label>기준실적 <em>*</em> <span class="iv-hint">(원)</span></label>
                <input type="text" id="calc-avg" placeholder="예: 767,160" inputmode="numeric">
              </div>
              <div class="iv-field">
                <label>아너스 기본순증목표 <span class="iv-hint">(본사/원)</span></label>
                <input type="text" id="calc-base-tgt" placeholder="예: 620,000" inputmode="numeric">
              </div>
              <div class="iv-field">
                <label>희망목표금액 <em>*</em> <span class="iv-hint">(원) ▲▼ 등급 이동</span></label>
                <div class="calc-tgt-wrap">
                  <input type="text" id="calc-tgt" placeholder="직접 입력" inputmode="numeric">
                  <button type="button" class="calc-step" id="btn-tgt-down" title="이전 등급">▼</button>
                  <button type="button" class="calc-step" id="btn-tgt-up" title="다음 등급">▲</button>
                </div>
              </div>
            </div>
            <div id="calc-incr-preview" style="display:none">📊 순증 = <span id="calc-incr-val">—</span></div>
            <div id="calc-result"><div class="rc-placeholder">기본 입력 항목을 입력하면 시상금이 자동 계산됩니다</div></div>
          </div>
        </div>

        <div class="iv-actions">
          <button class="btn-outline" id="btn-iv-cancel-edit" hidden>✕ 수정 취소</button>
          <button class="btn-primary" id="btn-iv-save">💾 저장</button>
        </div>
      </div>
    `;
  }

  // 교육생 등록/수정 폼 — 마스터목표 팝업 버튼
  function bindFormTargetPopup() {
    const popupBtn = $("#form-tgt-popup-btn");
    const popup = $("#form-tgt-popup");
    if (!popupBtn || !popup) return;

    popupBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      popup.hidden = !popup.hidden;
    });

    popup.querySelectorAll("button[data-fadd]").forEach((opt) => {
      opt.addEventListener("click", (e) => {
        e.stopPropagation();
        const add = opt.dataset.fadd;
        const tgtEl = $("#form-target");
        if (!tgtEl) return;
        popup.hidden = true;

        if (add === "input") {
          tgtEl.removeAttribute("readonly");
          tgtEl.value = "";
          tgtEl.focus();
          state.formTgtAddAmount = null;
        } else {
          const addVal = parseInt(add, 10);
          const baseVal = parseFloat($("#form-base")?.value) || 0;
          tgtEl.value = baseVal > 0 ? baseVal + addVal : addVal;
          tgtEl.setAttribute("readonly", "readonly");
          state.formTgtAddAmount = addVal;
        }
      });
    });

    document.addEventListener("click", (e) => {
      if (!popup.hidden && !popup.contains(e.target) && e.target !== popupBtn) {
        popup.hidden = true;
      }
    }, { capture: false });

    // 평균실적 변경 시 선택된 가산액으로 마스터목표 자동 재계산
    $("#form-base").addEventListener("input", () => {
      if (state.formTgtAddAmount === null) return;
      const tgtEl = $("#form-target");
      const baseVal = parseFloat($("#form-base")?.value) || 0;
      if (tgtEl && baseVal > 0) tgtEl.value = baseVal + state.formTgtAddAmount;
    });
  }

  function bindInterviewFormEvents() {
    $("#iv-seq").addEventListener("input", updateIvTitle);
    $("#iv-curAct").addEventListener("input", calcIvPct);
    // 현재실적 단독 업데이트
    $("#btn-iv-curAct-update").addEventListener("click", async () => {
      const empNo = state.selectedEmpNo;
      if (!empNo) return;
      const newVal = parseFloat($("#iv-curAct")?.value || "0") || 0;
      const sfx = _pgStepSfx();
      const field = sfx ? `pgCurrent${sfx}` : "pgCurrent";
      const curSt = state.students.find((x) => x.empNo === empNo);
      if (!curSt) { toast("교육생 정보를 찾을 수 없습니다.", "error"); return; }
      const btn = $("#btn-iv-curAct-update");
      const orig = btn.textContent;
      btn.disabled = true; btn.textContent = "저장중...";
      try {
        await window.DataAPI.save({ ...curSt, [field]: newVal });
        toast(`현재실적 ${newVal.toLocaleString()}원 저장 완료`, "success");
      } catch (e) {
        toast("저장 실패: " + e.message, "error");
      } finally {
        btn.disabled = false; btn.textContent = orig;
      }
    });
    $("#btn-iv-save").addEventListener("click", saveInterview);
    $("#btn-iv-save-top").addEventListener("click", saveInterview);
    $("#btn-iv-cancel-edit").addEventListener("click", cancelEditInterview);
    $("#btn-iv-cancel-edit-top").addEventListener("click", cancelEditInterview);
    $("#btn-cr-add").addEventListener("click", addCR);
    renderCR(); // 초기 빈 상태

    // 시상 계산기 이벤트
    $("#btn-calc-toggle").addEventListener("click", toggleCalcSection);
    $("#calc-avg").addEventListener("input", (e) => { fmtInput(e.target); onCalcAvgInput(); });
    $("#calc-base-tgt").addEventListener("input", (e) => { fmtInput(e.target); calc(); });
    const tgtInput = $("#calc-tgt");
    tgtInput.addEventListener("focus", () => { state.calcTgtUserEditing = true; });
    tgtInput.addEventListener("blur", (e) => { state.calcTgtUserEditing = false; fmtInput(e.target); calc(); });
    tgtInput.addEventListener("input", () => calc());
    $("#btn-tgt-down").addEventListener("click", () => stepTgt(-1));
    $("#btn-tgt-up").addEventListener("click", () => stepTgt(1));
  }

  // ========== 상담고객 (최대 5) ==========
  function initCR(list) {
    state.crData = Array.isArray(list) ? JSON.parse(JSON.stringify(list)) : [];
    renderCR();
    calcExpFromAmounts();
  }

  function addCR() {
    if (state.crData.length >= 5) {
      toast("최대 5명까지 입력 가능합니다.", "error");
      return;
    }
    state.crData.push({
      name: "", types: [], consult: [], material: [],
      amount: [], amountDirect: "", bj: [], memo: ""
    });
    renderCR();
  }

  function removeCR(idx) {
    state.crData.splice(idx, 1);
    renderCR();
    calcExpFromAmounts();
  }

  function ckBtn(ri, field, val, sel, radio) {
    const on = (sel || []).includes(val);
    return `<span class="cr-ck${on ? ' on' : ''}" data-ri="${ri}" data-field="${field}" data-val="${escapeHtml(val)}" data-radio="${radio ? '1' : '0'}">` +
      `<span class="cr-ck-b">${on ? '✓' : ''}</span>${escapeHtml(val)}</span>`;
  }

  function renderCR() {
    const el = $("#cr-rows");
    if (!el) return;
    if (!state.crData.length) {
      el.innerHTML = `<div class="cr-empty">상담고객을 추가하세요 (최대 5명)</div>`;
      return;
    }
    el.innerHTML = state.crData.map((c, i) => `
      <div class="cr" data-ri="${i}">
        <div class="cr-top">
          <span class="cr-num">고객 ${i + 1}</span>
          <input class="cr-name" data-ri="${i}" placeholder="성명" value="${escapeHtml(c.name || "")}">
          <button type="button" class="cr-del" data-ri="${i}" title="삭제">✕</button>
        </div>
        <div class="cr-secs">
          <div class="cr-sec">
            <div class="cr-sl">고객유형 <span class="cr-hint">(단일)</span></div>
            ${CT.map((t) => ckBtn(i, "types", t, c.types, true)).join("")}
          </div>
          <div class="cr-sec">
            <div class="cr-sl">상담단계 <span class="cr-hint">(단일)</span></div>
            ${CS.map((t) => ckBtn(i, "consult", t, c.consult, true)).join("")}
          </div>
          <div class="cr-sec">
            <div class="cr-sl">활용자료 <span class="cr-hint">(복수)</span></div>
            ${MT.map((t) => ckBtn(i, "material", t, c.material, false)).join("")}
          </div>
          <div class="cr-sec">
            <div class="cr-sl">제안금액 <span class="cr-hint">(단일 또는 직접입력)</span></div>
            ${AM.map((t) => ckBtn(i, "amount", t, c.amount, true)).join("")}
            <input class="cr-direct" data-ri="${i}" placeholder="직접입력(만원)" value="${escapeHtml(c.amountDirect || "")}">
          </div>
          <div class="cr-sec">
            <div class="cr-sl">보종 <span class="cr-hint">(단일)</span></div>
            ${BJ.map((t) => ckBtn(i, "bj", t, c.bj, true)).join("")}
          </div>
        </div>
        <div class="cr-memo-wrap">
          <div class="cr-sl">면담 내용</div>
          <textarea class="cr-memo" data-ri="${i}" rows="2" placeholder="이 고객과의 면담 내용을 입력하세요">${escapeHtml(c.memo || "")}</textarea>
        </div>
      </div>
    `).join("");

    // Event delegation
    el.querySelectorAll(".cr-name").forEach((inp) => {
      inp.addEventListener("input", (e) => {
        const ri = Number(e.target.dataset.ri);
        state.crData[ri].name = e.target.value;
      });
    });
    el.querySelectorAll(".cr-direct").forEach((inp) => {
      inp.addEventListener("input", (e) => {
        const ri = Number(e.target.dataset.ri);
        state.crData[ri].amountDirect = e.target.value;
        calcExpFromAmounts();
      });
    });
    el.querySelectorAll(".cr-memo").forEach((inp) => {
      inp.addEventListener("input", (e) => {
        const ri = Number(e.target.dataset.ri);
        state.crData[ri].memo = e.target.value;
      });
    });
    el.querySelectorAll(".cr-del").forEach((btn) => {
      btn.addEventListener("click", () => removeCR(Number(btn.dataset.ri)));
    });
    el.querySelectorAll(".cr-ck").forEach((span) => {
      span.addEventListener("click", () => {
        const ri = Number(span.dataset.ri);
        const field = span.dataset.field;
        const val = span.dataset.val;
        const radio = span.dataset.radio === "1";
        togCk(ri, field, val, span, radio);
      });
    });
  }

  function togCk(ri, field, val, el, radio) {
    const rec = state.crData[ri];
    if (!rec[field]) rec[field] = [];
    const arr = rec[field];
    const idx = arr.indexOf(val);
    if (radio) {
      // 같은 섹션의 다른 칩 해제
      const sec = el.closest(".cr-sec");
      if (sec) sec.querySelectorAll(".cr-ck").forEach((c) => {
        c.classList.remove("on");
        const b = c.querySelector(".cr-ck-b"); if (b) b.textContent = "";
      });
      rec[field] = [];
      if (idx < 0) {
        rec[field] = [val];
        el.classList.add("on");
        el.querySelector(".cr-ck-b").textContent = "✓";
      }
    } else {
      if (idx >= 0) {
        arr.splice(idx, 1);
        el.classList.remove("on");
        el.querySelector(".cr-ck-b").textContent = "";
      } else {
        arr.push(val);
        el.classList.add("on");
        el.querySelector(".cr-ck-b").textContent = "✓";
      }
    }
    if (field === "amount") calcExpFromAmounts();
  }

  // 제안금액 합산 → 주간예상실적(천원) 자동 세팅
  function calcExpFromAmounts() {
    const parseAmt = (s) => {
      const m = String(s || "").match(/(\d+)/);
      return m ? parseInt(m[1], 10) * 10 : 0; // 만원 → 천원
    };
    let total = 0;
    state.crData.forEach((c) => {
      const sel = c.amount || [];
      if (sel.length) {
        total += parseAmt(sel[0]);
      } else if (c.amountDirect) {
        const v = parseFloat(String(c.amountDirect).replace(/[^0-9.]/g, ""));
        if (!isNaN(v)) total += v * 10; // 만원 → 천원
      }
    });
    const expEl = $("#iv-exp");
    if (expEl) expEl.value = total || "";
  }

  // ========== 시상 계산기 (Phase 3) ==========
  function fmtInput(el) {
    if (!el) return;
    const raw = (el.value || "").replace(/[^0-9]/g, "");
    const num = parseInt(raw || "0", 10);
    const cur = el.selectionStart;
    const prev = el.value.length;
    el.value = raw ? num.toLocaleString() : "";
    const diff = el.value.length - prev;
    try { el.setSelectionRange(cur + diff, cur + diff); } catch (e) {}
  }
  function getRawVal(id) {
    const v = ($("#" + id)?.value || "").replace(/,/g, "").trim();
    return v === "" ? NaN : parseFloat(v);
  }
  function fmtW(mw)  { return Math.round(mw * 10000).toLocaleString() + "원"; }
  function fmtWon(w) { return Math.round(w).toLocaleString() + "원"; }

  function toggleCalcSection() {
    state.calcOpen = !state.calcOpen;
    const sec = $("#calc-section");
    const icon = $("#calc-toggle-icon");
    if (sec) sec.style.display = state.calcOpen ? "block" : "none";
    if (icon) icon.style.transform = state.calcOpen ? "rotate(0deg)" : "rotate(-90deg)";
    if (state.calcOpen) {
      const s = state.students.find((x) => x.empNo === state.selectedEmpNo);
      const avgEl = $("#calc-avg");
      const baseTgtEl = $("#calc-base-tgt");
      const tgtEl = $("#calc-tgt");
      if (s?.base  && avgEl && !avgEl.value)       avgEl.value = Number(s.base).toLocaleString();
      if (s?.honors && baseTgtEl && !baseTgtEl.value) baseTgtEl.value = Number(s.honors).toLocaleString();
      if (tgtEl && !tgtEl.value) {
        if (s?.target) tgtEl.value = Number(s.target).toLocaleString();
        else if (s?.honors) tgtEl.value = Number(s.honors).toLocaleString();
      }
      calc();
    }
  }

  function stepTgt(dir) {
    const steps = HONORS.map((h) => h.critVal * 10000).slice().reverse();
    const cur = getRawVal("calc-tgt") || 0;
    let next;
    if (dir > 0) {
      next = steps.find((v) => v > cur);
      if (!next) next = steps[steps.length - 1];
    } else {
      const lower = steps.filter((v) => v < cur);
      next = lower.length ? lower[lower.length - 1] : steps[0];
    }
    const el = $("#calc-tgt");
    if (el) el.value = Number(next).toLocaleString();
    calc();
  }

  function onCalcAvgInput() {
    calc();
  }

  function calcMasterAward(incrMW) {
    for (const t of MASTER_AWARD) {
      if (incrMW >= t.critVal) {
        return t.type === "pct" ? Math.round(incrMW * t.val / 100 * 10) / 10 : t.val;
      }
    }
    return 0;
  }

  function buildAwardCertificateHtml(s, r, positiveWord) {
    const fw = (mw) => Math.round(mw * 10000).toLocaleString() + "원";
    const region = (s.region || "").replace(/지역단$|사업부$/, "");
    const branch = (s.branch || "").replace(/지점$/, "");
    const name = s.name || "";
    const empNo = s.empNo || "";
    const dateStr = new Date().toLocaleDateString("ko-KR");
    const award1Str = r.award1 ? fw(r.award1) : "해당없음";
    const award1GradeKo = (r.award1Grade || "해당없음").replace(/\s*\([^)]*\)\s*/g, "").trim();
    const award1CritStr = r.award1Idx >= 0 ? HONORS[r.award1Idx].criteria : "—";
    const totalStr = fw(r.total);
    return `
<div class="cert-a4">
  <div class="cert-hdr">
    <div class="cert-hdr-stars">★ ★ ★ ★ ★</div>
    <div class="cert-hdr-title">고객컨설팅마스터</div>
    <div class="cert-hdr-sub">시상 예상 답안지</div>
    <div class="cert-hdr-quarter">2026년 2분기 · ${escapeHtml(dateStr)} 기준</div>
  </div>

  <div class="cert-info-row">
    <div class="cert-info-card dark"><div class="cert-ic-lbl">지역단</div><div class="cert-ic-val">${escapeHtml(region)}</div></div>
    <div class="cert-info-card dark"><div class="cert-ic-lbl">지점</div><div class="cert-ic-val">${escapeHtml(branch)}</div></div>
    <div class="cert-info-card dark nm"><div class="cert-ic-lbl">성명</div><div class="cert-ic-val">${escapeHtml(name)}</div></div>
    <div class="cert-info-card"><div class="cert-ic-lbl">사번</div><div class="cert-ic-val">${escapeHtml(empNo)}</div></div>
  </div>

  <div class="cert-target-bar">
    <span class="cert-tgt-lbl">🎯 희망목표금액</span>
    <span class="cert-tgt-val">${Math.round(r.tgtRaw).toLocaleString()}원</span>
  </div>

  <div class="cert-awards-row">
    <div class="cert-aw-card purple">
      <div class="cert-aw-num">①</div>
      <div class="cert-aw-name">아너스클럽</div>
      <div class="cert-aw-grade">${escapeHtml(award1GradeKo)}</div>
      <div class="cert-aw-crit">${escapeHtml(award1CritStr)}</div>
      <div class="cert-aw-amt">${award1Str}</div>
    </div>
    <div class="cert-aw-card green">
      <div class="cert-aw-num">②</div>
      <div class="cert-aw-name">하이포인트</div>
      <div class="cert-aw-grade">3개월 합계</div>
      <div class="cert-aw-crit">기본 ${fw(5)} + 순증×50%</div>
      <div class="cert-aw-amt">${fw(r.award2M3)}</div>
    </div>
    <div class="cert-aw-card blue">
      <div class="cert-aw-num">③</div>
      <div class="cert-aw-name">마스터과정</div>
      <div class="cert-aw-grade">2개월 합계</div>
      <div class="cert-aw-crit">${r.award3Tier ? escapeHtml(r.award3Tier.criteria) : "순증 5만원 미만"}</div>
      <div class="cert-aw-amt">${r.award3 > 0 ? fw(r.award3 * 2) : "해당없음"}</div>
    </div>
  </div>

  <div class="cert-total-box">
    <div class="cert-total-lbl">🏆 최종 예상 시상금 합계 (①+②+③)</div>
    <div class="cert-total-formula">${award1Str} + ${fw(r.award2M3)} + ${r.award3 > 0 ? fw(r.award3 * 2) : "0원"}</div>
    <div class="cert-total-amt">${totalStr}</div>
  </div>

  <div class="cert-footer">
    <div class="cert-footer-msg">
      <span class="cert-footer-name">${escapeHtml(name)}</span>님은
      <span class="cert-footer-branch">${escapeHtml(branch)}</span>지점의
      <span class="cert-footer-word">'${escapeHtml(positiveWord)}'</span>
    </div>
    <div class="cert-footer-note">※ 아너스클럽: 3개월 연속달성 · 하이포인트: 월최대 20만·분기최대 50만 · 마스터: 2개월 지급</div>
  </div>
</div>`;
  }

  function getAwardCertificateCss() {
    return `
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'Noto Sans KR','Malgun Gothic','Apple SD Gothic Neo',sans-serif;background:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
.cert-a4{width:210mm;min-height:297mm;margin:0 auto;padding:14mm 16mm;background:#fff;display:flex;flex-direction:column;gap:10px;}
.cert-hdr{background:linear-gradient(135deg,#0D1B4B,#1A3A8F);border-radius:12px;padding:18px 20px;text-align:center;color:#fff;}
.cert-hdr-stars{font-size:18px;color:#FFD700;letter-spacing:8px;margin-bottom:6px;}
.cert-hdr-title{font-size:28px;font-weight:900;letter-spacing:2px;margin-bottom:4px;}
.cert-hdr-sub{font-size:18px;font-weight:700;color:#B0C4FF;margin-bottom:6px;}
.cert-hdr-quarter{font-size:13px;color:rgba(255,255,255,.7);}
.cert-info-row{display:grid;grid-template-columns:1fr 1fr 1.2fr 1fr;gap:6px;}
.cert-info-card{background:#F0F4FF;border-radius:8px;padding:8px 10px;text-align:center;border:1px solid #C5D0F0;}
.cert-info-card.dark{background:#1A3A8F;border-color:#1A3A8F;}
.cert-ic-lbl{font-size:10px;font-weight:700;color:#5C6BC0;margin-bottom:3px;}
.cert-info-card.dark .cert-ic-lbl{color:rgba(255,255,255,.65);}
.cert-ic-val{font-size:15px;font-weight:900;color:#0D1B4B;}
.cert-info-card.dark .cert-ic-val{color:#fff;font-size:16px;}
.cert-info-card.nm .cert-ic-val{font-size:18px;}
.cert-target-bar{background:linear-gradient(90deg,#FF6F00,#FFA000);border-radius:8px;padding:10px 16px;display:flex;align-items:center;justify-content:space-between;color:#fff;}
.cert-tgt-lbl{font-size:14px;font-weight:700;}
.cert-tgt-val{font-size:22px;font-weight:900;letter-spacing:-1px;}
.cert-awards-row{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;}
.cert-aw-card{border-radius:10px;padding:14px 10px;text-align:center;border:2px solid transparent;}
.cert-aw-card.purple{background:#F3E5F5;border-color:#7B1FA2;}
.cert-aw-card.green{background:#E8F5E9;border-color:#2E7D32;}
.cert-aw-card.blue{background:#E3F2FD;border-color:#1565C0;}
.cert-aw-num{font-size:22px;font-weight:900;margin-bottom:4px;}
.cert-aw-card.purple .cert-aw-num{color:#7B1FA2;}
.cert-aw-card.green .cert-aw-num{color:#2E7D32;}
.cert-aw-card.blue .cert-aw-num{color:#1565C0;}
.cert-aw-name{font-size:13px;font-weight:800;color:#1A1A1A;margin-bottom:3px;}
.cert-aw-grade{font-size:12px;font-weight:700;margin-bottom:2px;}
.cert-aw-card.purple .cert-aw-grade{color:#6A1B9A;}
.cert-aw-card.green .cert-aw-grade{color:#1B5E20;}
.cert-aw-card.blue .cert-aw-grade{color:#0D47A1;}
.cert-aw-crit{font-size:10px;color:#666;margin-bottom:6px;}
.cert-aw-amt{font-size:17px;font-weight:900;padding:6px 0;border-radius:6px;}
.cert-aw-card.purple .cert-aw-amt{color:#6A1B9A;background:rgba(123,31,162,.1);}
.cert-aw-card.green .cert-aw-amt{color:#1B5E20;background:rgba(46,125,50,.1);}
.cert-aw-card.blue .cert-aw-amt{color:#0D47A1;background:rgba(21,101,192,.1);}
.cert-total-box{background:linear-gradient(135deg,#0D1B4B,#1A3A8F);border-radius:12px;padding:16px 20px;text-align:center;color:#fff;}
.cert-total-lbl{font-size:13px;font-weight:700;color:rgba(255,255,255,.8);margin-bottom:4px;}
.cert-total-formula{font-size:11px;color:rgba(255,255,255,.55);margin-bottom:8px;}
.cert-total-amt{font-size:36px;font-weight:900;color:#FFD700;letter-spacing:-1px;}
.cert-footer{background:linear-gradient(135deg,#1B5E20,#2E7D32);border-radius:12px;padding:16px 20px;text-align:center;color:#fff;margin-top:auto;}
.cert-footer-msg{font-size:18px;font-weight:800;margin-bottom:8px;line-height:1.5;}
.cert-footer-name{color:#FFD700;font-size:20px;}
.cert-footer-branch{color:#A5D6A7;}
.cert-footer-word{color:#FFD700;font-size:22px;font-style:italic;}
.cert-footer-note{font-size:9px;color:rgba(255,255,255,.6);line-height:1.5;}
@media print{@page{size:A4 portrait;margin:0;}body{margin:0;}.cert-a4{padding:10mm 12mm;}}
`;
  }

  function openAwardPrintPreview() {
    const s = state.students.find((x) => x.empNo === state.selectedEmpNo);
    const r = state.lastCalcResult;
    if (!r || !s) { toast("시상 계산기를 먼저 실행하세요.", "warn"); return; }

    const fakeCalcItv = { calcAvg: String(r.avgRaw || ""), calcBaseTgt: String(r.baseTgtRaw || ""), calcTgt: String(r.tgtRaw || "") };
    const lastInsItv = state.consultations.slice().reverse().find((c) => c.ins);

    const wordKey = "awardPrintWordIdx";
    const wordIdx = ((parseInt(localStorage.getItem(wordKey) || "-1", 10) + 1) % AWARD_POSITIVE_WORDS.length);
    localStorage.setItem(wordKey, String(wordIdx));
    const positiveWord = AWARD_POSITIVE_WORDS[wordIdx];

    const sheetHtml = buildAwardSheetPageHtml(s, fakeCalcItv, lastInsItv, { positiveWord });
    if (!sheetHtml) { toast("시상 정보가 없습니다.", "warn"); return; }

    const fullHtml = `<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"><title>시상 예상 답안지</title><style>${AWARD_PRINT_CSS}</style></head><body>${sheetHtml}</body></html>`;

    const vc = s.center || s.region || "";
    const branchShort = (s.branch || "").replace(/지점$/, "");
    const sName = s.name || "교육생";
    const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, "");
    const filename = `${vc}_${branchShort}지점_${sName}_시상예상답안지_${dateStr}`;

    let modal = document.getElementById("modal-award-print");
    if (!modal) {
      modal = document.createElement("div");
      modal.id = "modal-award-print";
      modal.className = "modal";
      modal.hidden = true;
      modal.innerHTML = `
        <div class="modal-backdrop" data-close></div>
        <div class="modal-panel aprint-panel">
          <div class="modal-head">
            <h3>🖨️ 시상인쇄 미리보기</h3>
            <button class="modal-close" data-close aria-label="닫기">&times;</button>
          </div>
          <div class="modal-body aprint-body">
            <div class="aprint-toolbar">
              <button id="btn-aprint-printer" class="aprint-btn aprint-blue">🖨️ 인쇄</button>
              <button id="btn-aprint-pdf" class="aprint-btn aprint-red">📄 PDF저장</button>
              <button id="btn-aprint-png" class="aprint-btn aprint-green">🖼️ PNG저장</button>
              <button id="btn-aprint-kakao" class="aprint-btn aprint-kakao">💬 카카오톡으로 저장하기</button>
              <button id="btn-aprint-close" class="aprint-btn aprint-gray">✕ 닫기</button>
            </div>
            <div class="aprint-scroll">
              <iframe id="aprint-iframe" class="aprint-iframe"></iframe>
            </div>
          </div>
        </div>
      `;
      document.body.appendChild(modal);
      modal.querySelectorAll("[data-close]").forEach((el) => el.addEventListener("click", () => { modal.hidden = true; }));
    } else {
      // 버튼 3개 구조로 갱신 (이전 버전 modal 재사용 시)
      const tb = modal.querySelector(".aprint-toolbar");
      if (tb && !tb.querySelector("#btn-aprint-close")) {
        tb.innerHTML = `
          <button id="btn-aprint-printer" class="aprint-btn aprint-blue">🖨️ 인쇄</button>
          <button id="btn-aprint-pdf" class="aprint-btn aprint-red">📄 PDF저장</button>
          <button id="btn-aprint-png" class="aprint-btn aprint-green">🖼️ PNG저장</button>
          <button id="btn-aprint-kakao" class="aprint-btn aprint-kakao">💬 카카오톡으로 저장하기</button>
          <button id="btn-aprint-close" class="aprint-btn aprint-gray">✕ 닫기</button>
        `;
      }
    }

    const iframe = modal.querySelector("#aprint-iframe");
    iframe.srcdoc = fullHtml;

    // 스크립트 동적 로드 (이미 로드된 경우 globalOk()로 확인)
    const loadScript = (src, globalOk) => new Promise((resolve, reject) => {
      if (globalOk && globalOk()) { resolve(); return; }
      const prev = document.querySelector(`script[src="${src}"]`);
      if (prev) prev.remove(); // 이전 로드 실패 시 재시도
      const sc = document.createElement("script");
      sc.src = src;
      sc.onload = () => { resolve(); };
      sc.onerror = () => reject(new Error("라이브러리 로드 실패 (네트워크 확인 필요)"));
      document.head.appendChild(sc);
    });

    // html2canvas 로 시상 내용 캡처 — 모달 뒤에 z-index:1 로 렌더링
    const captureCanvas = async () => {
      await loadScript(
        "https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js",
        () => typeof html2canvas !== "undefined"
      );

      // AWARD_PRINT_CSS 를 #award-cap-tmp 스코프로 래핑 (메인 페이지 CSS 충돌 방지)
      const rawCss = AWARD_PRINT_CSS.split("@media print")[0].replace(/body\{[^}]*\}/, "");
      const scopedCss = rawCss.split("}").map((part) => {
        const bi = part.indexOf("{");
        if (bi === -1) return part;
        const sel = part.slice(0, bi).trim();
        if (!sel) return part;
        const scoped = sel.split(",").map((s) => `#award-cap-tmp ${s.trim()}`).join(",");
        return `${scoped}${part.slice(bi)}`;
      }).join("}");

      const styleEl = document.createElement("style");
      styleEl.textContent = scopedCss;
      document.head.appendChild(styleEl);

      // position:fixed + z-index:1 → 모달(z-index 높음) 뒤에 가려져 화면에 안 보이지만 정상 렌더링
      const wrap = document.createElement("div");
      wrap.id = "award-cap-tmp";
      wrap.style.cssText = "position:fixed;top:0;left:0;width:794px;z-index:1;pointer-events:none;font-family:'Noto Sans KR','Malgun Gothic',sans-serif;background:#fff;color:#1A1A1A;font-size:12px;line-height:1.4;";
      wrap.innerHTML = sheetHtml;
      document.body.appendChild(wrap);

      // 폰트·렌더링 대기
      await new Promise((res) => setTimeout(res, 400));

      let canvas;
      try {
        const pgEl = wrap.querySelector(".pg");
        if (!pgEl) throw new Error("인쇄 요소를 찾을 수 없습니다");
        canvas = await html2canvas(pgEl, {
          scale: 2, useCORS: true, backgroundColor: "#ffffff", logging: false,
          scrollX: 0, scrollY: 0
        });
      } finally {
        document.body.removeChild(wrap);
        document.head.removeChild(styleEl);
      }
      return canvas;
    };

    const shareOrDownload = async (blob, mimeType, ext) => {
      const file = new File([blob], `${filename}.${ext}`, { type: mimeType });
      let shared = false;
      if (typeof navigator.canShare === "function") {
        try {
          if (navigator.canShare({ files: [file] })) {
            await navigator.share({ files: [file] });
            shared = true;
          }
        } catch (e) {
          if (e.name !== "AbortError") throw e; // 사용자 취소는 무시
          shared = true; // 취소도 정상 종료
        }
      }
      if (!shared) {
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = `${filename}.${ext}`;
        link.click();
        setTimeout(() => URL.revokeObjectURL(url), 15000);
      }
    };

    const doPrint = () => {
      if (iframe && iframe.contentWindow) { iframe.contentWindow.focus(); iframe.contentWindow.print(); }
    };

    const doPdf = async () => {
      toast("PDF 생성 중...", "info");
      try {
        const canvas = await captureCanvas();
        await loadScript(
          "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js",
          () => !!(window.jspdf?.jsPDF || window.jsPDF)
        );
        const JsPDF = (window.jspdf && window.jspdf.jsPDF) || window.jsPDF;
        if (!JsPDF) throw new Error("jsPDF 라이브러리 로드 실패");
        const pdf = new JsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
        const pdfW = pdf.internal.pageSize.getWidth();
        const pdfH = pdf.internal.pageSize.getHeight();
        const imgH = (canvas.height * pdfW) / canvas.width;
        pdf.addImage(canvas.toDataURL("image/jpeg", 0.92), "JPEG", 0, 0, pdfW, Math.min(imgH, pdfH));
        const pdfBlob = pdf.output("blob");
        await shareOrDownload(pdfBlob, "application/pdf", "pdf");
        toast("PDF 저장 완료!", "success");
      } catch (e) { toast("PDF 실패: " + (e.message || String(e)), "error"); }
    };

    const doPng = async () => {
      toast("PNG 생성 중...", "info");
      try {
        const canvas = await captureCanvas();
        await new Promise((resolve, reject) => {
          canvas.toBlob(async (blob) => {
            if (!blob) { reject(new Error("canvas.toBlob 실패")); return; }
            try { await shareOrDownload(blob, "image/png", "png"); resolve(); }
            catch (e) { reject(e); }
          }, "image/png");
        });
        toast("PNG 저장 완료!", "success");
      } catch (e) { toast("PNG 실패: " + (e.message || String(e)), "error"); }
    };

    const doKakao = async () => {
      toast("이미지 생성 중...", "info");
      try {
        const canvas = await captureCanvas();
        const blob = await new Promise((resolve, reject) => {
          canvas.toBlob((b) => b ? resolve(b) : reject(new Error("이미지 생성 실패")), "image/png");
        });
        const file = new File([blob], `${filename}.png`, { type: "image/png" });

        // ① 파일 공유 API 지원 기기(Android Chrome, iOS Safari 최신) → 공유 시트
        if (typeof navigator.canShare === "function" && navigator.canShare({ files: [file] })) {
          try {
            await navigator.share({ files: [file], title: "시상 예상답안지" });
            toast("카카오톡 공유 완료!", "success");
            return;
          } catch (e) {
            if (e.name === "AbortError") return; // 사용자 취소
            // share 실패 시 ② 로 fallback
          }
        }

        // ② PNG 저장 후 카카오톡 앱 직접 열기
        const objUrl = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = objUrl;
        link.download = `${filename}.png`;
        link.click();
        setTimeout(() => URL.revokeObjectURL(objUrl), 15000);
        toast("이미지 저장됨 — 카카오톡이 열립니다.", "success");
        // 짧은 딜레이 후 카카오톡 URI 스킴으로 앱 열기
        setTimeout(() => {
          window.location.href = "kakaotalk://";
        }, 700);
      } catch (e) { toast("카카오톡 공유 실패: " + (e.message || String(e)), "error"); }
    };

    ["btn-aprint-printer", "btn-aprint-pdf", "btn-aprint-png", "btn-aprint-kakao", "btn-aprint-close"].forEach((id) => {
      const el = modal.querySelector("#" + id);
      if (!el) return;
      const clone = el.cloneNode(true);
      el.parentNode.replaceChild(clone, el);
    });
    modal.querySelector("#btn-aprint-printer").addEventListener("click", doPrint);
    modal.querySelector("#btn-aprint-pdf") .addEventListener("click", doPdf);
    modal.querySelector("#btn-aprint-png") .addEventListener("click", doPng);
    modal.querySelector("#btn-aprint-kakao").addEventListener("click", doKakao);
    modal.querySelector("#btn-aprint-close").addEventListener("click", () => { modal.hidden = true; });

    modal.hidden = false;
  }

  function calc() {
    const avgRaw = getRawVal("calc-avg");
    let tgtRaw = getRawVal("calc-tgt");
    const baseTgtRaw = getRawVal("calc-base-tgt") || 0;

    // calc-tgt 가 1000 미만이면 천원 단위 오입력 → 원으로 정규화
    if (!state.calcTgtUserEditing && !isNaN(tgtRaw) && tgtRaw > 0 && tgtRaw < 1000) {
      tgtRaw *= 1000;
      const tgtEl = $("#calc-tgt");
      if (tgtEl) tgtEl.value = Math.round(tgtRaw).toLocaleString();
    }

    const tgt = tgtRaw / 10000;
    const baseTgt = baseTgtRaw / 10000;
    const incr = Math.round(Math.max(0, tgt - baseTgt) * 10000) / 10000;

    // 순증 미리보기
    const prev = $("#calc-incr-preview");
    const incrVal = $("#calc-incr-val");
    if (prev && !isNaN(avgRaw) && !isNaN(tgtRaw)) {
      prev.style.display = "block";
      if (incrVal) incrVal.textContent = `${fmtWon(incr * 10000)} (희망목표 ${fmtWon(tgtRaw)} − 기본순증목표 ${fmtWon(baseTgtRaw)})`;
    } else if (prev) {
      prev.style.display = "none";
    }

    const res = $("#calc-result");
    if (!res) return;

    if (isNaN(avgRaw) || isNaN(tgtRaw)) {
      res.innerHTML = `<div class="rc-placeholder">기본 입력 항목을 입력하면 시상금이 자동 계산됩니다</div>`;
      return;
    }

    // ① 아너스클럽
    let award1 = 0, award1Grade = "해당없음", award1Idx = -1;
    for (let i = 0; i < HONORS.length; i++) {
      if (tgt >= HONORS[i].critVal) {
        award1 = HONORS[i].prize; award1Grade = HONORS[i].grade; award1Idx = i; break;
      }
    }

    // ② 개인 순증시상 (하이포인트)
    const baseTgtMet = (baseTgt <= 0) || (tgt >= baseTgt);
    const monthlyExtra = baseTgtMet ? Math.floor(incr * INCR_CFG.rate / 100 * 10) / 10 : 0;
    const monthlySub = baseTgtMet ? (INCR_CFG.base + monthlyExtra) : 0;
    const monthlyFinal = baseTgtMet ? Math.min(monthlySub, INCR_CFG.mcap) : 0;
    const award2M3 = baseTgtMet ? Math.min(monthlyFinal * 3, INCR_CFG.qcap) : 0;

    // ③ 마스터과정 개인시상 (교육생 평균실적 기준)
    const _calcS = state.students.find((x) => x.empNo === state.selectedEmpNo);
    const insRaw3 = _calcS ? Math.round(Number(_calcS.base)) : 0;
    const incrMaster = Math.max(0, tgtRaw - insRaw3) / 10000;
    const award3 = calcMasterAward(incrMaster);
    const award3Tier = MASTER_AWARD.find((t) => incrMaster >= t.critVal);
    const award3TierIdx = award3Tier ? MASTER_AWARD.indexOf(award3Tier) : MASTER_AWARD.length;
    const award3NextTier = award3TierIdx > 0 ? MASTER_AWARD[award3TierIdx - 1] : null;
    const award3NeedIncr = award3NextTier ? Math.max(0, award3NextTier.critVal - incrMaster) : 0;
    const award3NeedTgt = Math.ceil(award3NeedIncr * 10000);

    const total = award1 + award2M3 + award3 * 2;

    // ±2 tiers around current position only
    const honorVisibleIdxs = (() => {
      if (award1Idx >= 0) {
        const lo = Math.max(0, award1Idx - 2);
        const hi = Math.min(HONORS.length - 1, award1Idx + 2);
        const arr = [];
        for (let i = lo; i <= hi; i++) arr.push(i);
        return arr;
      }
      // No achievement: show 3 easiest tiers
      const arr = [];
      for (let i = Math.max(0, HONORS.length - 3); i < HONORS.length; i++) arr.push(i);
      return arr;
    })();
    const honorRows = honorVisibleIdxs.map((i) => {
      const h = HONORS[i];
      const isActive = i === award1Idx;
      const cls = isActive ? "rs-on" : "rs-miss";
      const icon = isActive ? "🏆" : "⬜";
      const dist = Math.abs(i - (award1Idx >= 0 ? award1Idx : HONORS.length));
      const gradeKo = h.grade.replace(/\s*\([^)]*\)\s*/g, "").trim();
      const prizeStr = (h.prize >= 100 && h.prize % 100 === 0) ? `${h.prize / 100}백만원` : `${h.prize}만원`;
      return `<tr class="${cls}" data-dist="${dist}">
        <td class="c">${icon}</td>
        <td class="nm">${escapeHtml(gradeKo)}</td>
        <td class="crit">${h.criteria}</td>
        <td class="prize">${prizeStr}</td>
      </tr>`;
    }).join("");

    res.innerHTML = `
      <div class="rc-header">
        <div class="rc-title">📈 계산 결과</div>
        <div class="rc-total-hdr">3개월 예상 ${fmtW(total)}</div>
      </div>
      <div class="rc-body">
        <div class="rc-metrics">
          <div class="rc-m"><div class="rc-m-l">순증</div><div class="rc-m-v blue">${fmtWon(incr * 10000)}</div><div class="rc-m-s">${fmtWon(tgtRaw)} − ${fmtWon(baseTgtRaw)}</div></div>
          <div class="rc-m"><div class="rc-m-l">희망목표</div><div class="rc-m-v red">${fmtWon(tgtRaw)}</div></div>
          <div class="rc-m"><div class="rc-m-l">합계(3개월)</div><div class="rc-m-v purple">${fmtW(total)}</div><div class="rc-m-s">①+②+③</div></div>
        </div>

        <div class="rc-sec-l">① 아너스클럽 <span>희망목표 ${fmtWon(tgtRaw)} 기준</span></div>
        <table class="rs-table"><thead><tr>
          <th class="c">달성</th><th>시상등급</th><th class="crit">기준</th><th class="prize">시상금</th>
        </tr></thead><tbody>${honorRows}</tbody></table>
        <div class="rs-applied">→ 현재 등급: <strong>${escapeHtml((award1Grade || "").replace(/\s*\([^)]*\)\s*/g, "").trim() || "해당없음")}</strong> · 시상금: <strong>${award1 ? fmtW(award1) : "해당없음"}</strong></div>

        <div class="rc-sec-l">② 개인 순증시상 (하이포인트) ${!baseTgtMet ? `<span class="warn-badge">⚠ 기본순증목표 미달</span>` : ""}</div>
        ${!baseTgtMet
          ? `<div class="warn-box">희망목표(${fmtWon(tgtRaw)})이 기본순증목표(${fmtWon(baseTgtRaw)})에 미달 → ② 0원</div>`
          : `<table class="calc-table"><tbody>
              <tr><td>기본 지급</td><td>기본 시상금</td><td class="r">${fmtW(INCR_CFG.base)}</td></tr>
              <tr><td>추가</td><td class="red">순증 ${fmtWon(incr * 10000)} × ${INCR_CFG.rate}%</td><td class="r">${fmtW(monthlyExtra)}</td></tr>
              <tr class="calc-sub"><td colspan="2">월 소계 (최대 ${fmtW(INCR_CFG.mcap)})</td><td class="r">${fmtW(monthlyFinal)}${monthlySub > INCR_CFG.mcap ? " <span class=\"cap\">⚠캡</span>" : ""}</td></tr>
              <tr class="calc-q"><td colspan="2">3개월 합계 (최대 ${fmtW(INCR_CFG.qcap)})</td><td class="r">${fmtW(award2M3)}${monthlyFinal * 3 > INCR_CFG.qcap ? " <span class=\"cap\">⚠캡</span>" : ""}</td></tr>
            </tbody></table>`}

        <div class="rc-sec-l blue">③ 고객컨설팅마스터 개인시상 <span>순증 기준</span></div>
        ${award3 > 0
          ? `<div class="master-box">
              <div class="mb-grade">🏆 ${escapeHtml(award3Tier.criteria)} 달성 (${escapeHtml(award3Tier.label)})</div>
              <div class="mb-incr">순증 <strong>${fmtWon(incrMaster * 10000)}</strong> = 희망목표 <strong>${fmtWon(tgtRaw)}</strong> − 인보험평균 <strong>${fmtWon(insRaw3)}</strong></div>
              <div class="mb-result">매월 ${fmtW(award3)} × 2개월 = <strong>${fmtW(award3 * 2)}</strong></div>
            </div>`
          : `<div class="master-none">순증 5만원 미만 — 해당없음 (순증 ${fmtWon(incrMaster * 10000)})</div>`}

        ${(() => {
          if (!_calcS) return "";
          const _region  = _calcS.region || state.filter.region || state.progressRegion || DEFAULT_REGION;
          const _cohort  = (_calcS.cohort || state.filter.cohort || "").replace(/기$/, "");
          const _step    = state.filter.step || state.progressStep || "1";
          const _pa      = getProgressAwardConfig(_region, _cohort, _step);
          const _plan    = _pa.plan;
          const _cohortLabel = _cohort ? `${_cohort}기` : "";
          const _stepLabel   = `Step ${_step}`;
          const _statMe  = getProgressStat(_calcS);
          const _empNo   = _calcS.empNo;

          const _peers = state.students.filter((s) => {
            const sc = (s.cohort || "").replace(/기$/, "");
            return s.region === _region && (!_cohort || sc === _cohort);
          });
          const _chkPeerTop = _pa.isTopEligible || _pa.isEligible;
          const _elig   = _peers.map((s) => getProgressStat(s)).filter((st) => _chkPeerTop(st.s));
          const _byRate = sortStatsForType(_elig, "rate");
          const _byAmt  = sortStatsForType(_elig, "amt");
          // dedup은 현금 비교용 — 아이템 시상은 effectiveAmt=0이므로 별도 처리
          const { rateAsgn, amtAsgn } = computeBothAwardAssignments(_byRate, _byAmt, _pa);

          // ── 신장률 순위 시상 (아이템 시상 포함) ──
          let _rateCard = null, _rateRankVal = 0;
          if (_empNo && _pa.rateConfig?.enabled && Number(_pa.rateConfig.n) > 0) {
            _rateRankVal = _byRate.findIndex(st => st.s.empNo === _empNo) + 1;
            const _n = Number(_pa.rateConfig.n);
            if (_rateRankVal > 0 && _rateRankVal <= _n) {
              const _p = _pa.rateConfig?.payouts?.[_rateRankVal - 1];
              if (_p != null) {
                const _np = normPayout(_p);
                const _rA = rateAsgn.get(_empNo);
                if (_np.type === "item" || _rA?.status === "mine") {
                  _rateCard = { icon: "📈", label: `신장률 ${_rateRankVal}위`, sub: `${_statMe.rate.toFixed(1)}%`, prize: payoutLabel(_p), isItem: _np.type === "item" };
                }
              }
            }
          }

          // ── 신장액 순위 시상 (아이템 시상 포함) ──
          let _amtCard = null, _amtRankVal = 0;
          if (_empNo && _pa.amtConfig?.enabled && Number(_pa.amtConfig.n) > 0) {
            _amtRankVal = _byAmt.findIndex(st => st.s.empNo === _empNo) + 1;
            const _n = Number(_pa.amtConfig.n);
            if (_amtRankVal > 0 && _amtRankVal <= _n) {
              const _p = _pa.amtConfig?.payouts?.[_amtRankVal - 1];
              if (_p != null) {
                const _np = normPayout(_p);
                const _aA = amtAsgn.get(_empNo);
                const _netMw = Math.round(_statMe.net / 10000);
                if (_np.type === "item" || _aA?.status === "mine") {
                  _amtCard = { icon: "💰", label: `신장액 ${_amtRankVal}위`, sub: `+${_netMw}만원`, prize: payoutLabel(_p), isItem: _np.type === "item" };
                }
              }
            }
          }

          // 신장률/신장액 중 더 좋은 순위 하나만 표시
          if (_rateCard && _amtCard) {
            if (_rateRankVal <= _amtRankVal) _amtCard = null;
            else _rateCard = null;
          }

          // ── 팀시상 (groupAward1: 인원, groupAward2: 달성률) ──
          let _ga1Card = null, _ga2Card = null;
          const _myTeam = String(_calcS.team || "").trim();
          if (_myTeam) {
            const _teamMates = _elig.filter(st => String(st.s.team || "").trim() === _myTeam);
            const _teamSize  = _teamMates.length;
            const ga1items = _ga1Items(_plan.groupAward1);
            for (const it of [...ga1items].sort((a, b) => Number(b.threshold) - Number(a.threshold))) {
              if (_teamSize >= Number(it.threshold)) {
                const np = normPayout(it.payout);
                _ga1Card = { icon: "👥", label: `팀시상 — ${_myTeam}조`, sub: `팀원 ${_teamSize}명`, prize: payoutLabel(it.payout), isItem: np.type === "item", isTeam: true };
                break;
              }
            }
            const ga2items = _ga2Items(_plan.groupAward2);
            if (_teamMates.length > 0 && ga2items.length) {
              const _teamAvg = _teamMates.reduce((s, st) => s + (st.rate || 0), 0) / _teamMates.length;
              for (const it of [...ga2items].sort((a, b) => Number(b.rateThreshold) - Number(a.rateThreshold))) {
                if (_teamAvg >= Number(it.rateThreshold)) {
                  const np = normPayout(it.payout);
                  _ga2Card = { icon: "🏅", label: `팀달성률 — ${_myTeam}조`, sub: `평균 ${_teamAvg.toFixed(1)}%`, prize: payoutLabel(it.payout), isItem: np.type === "item", isTeam: true };
                  break;
                }
              }
            }
          }

          const _allCards = [_rateCard, _amtCard, _ga1Card, _ga2Card].filter(Boolean);
          const hasAward  = _allCards.length > 0;
          const hasAnyConfig = _pa.rateConfig || _pa.amtConfig || _plan.groupAward1?.enabled || _plan.groupAward2?.enabled;

          // ── 개인 순증 시상 (과정 시상안 기준, 마스터목표 순증 기준) ──
          const _targetNetWon = Math.max(0, tgtRaw - insRaw3);
          const _targetNetMw  = Math.round(_targetNetWon / 10000);
          const _tierHit = _pa.tiers.find(t => _targetNetWon >= t.min);
          let _tierSectionHtml = "";
          if (_pa.tiers.length > 0) {
            const _tierBadge = `${_cohortLabel ? _cohortLabel + " " : ""}${_stepLabel}`;
            if (_tierHit) {
              const _tierPrize = _tierHit.type === "pct"
                ? `${_tierHit.payVal}% 지급 (${fmtWon(Math.round(_targetNetWon * _tierHit.val))})`
                : `${_tierHit.payVal}만원`;
              _tierSectionHtml = `
                <div class="aps-tier-card aps-tier-hit">
                  <div class="aps-tier-head">🎁 개인 순증 시상 <span class="aps-badge">${escapeHtml(_tierBadge)}</span></div>
                  <div class="aps-tier-body">마스터목표 순증 <strong>${_targetNetMw}만원</strong> → <span class="aps-tier-prize">${escapeHtml(_tierPrize)}</span> 획득 예정</div>
                </div>`;
            } else {
              const _lowestTier = _pa.tiers[_pa.tiers.length - 1];
              const _needMw = _lowestTier ? Math.max(0, Math.ceil((_lowestTier.min - _targetNetWon) / 10000)) : 0;
              _tierSectionHtml = `
                <div class="aps-tier-card aps-tier-miss">
                  <div class="aps-tier-head">🎁 개인 순증 시상 <span class="aps-badge">${escapeHtml(_tierBadge)}</span></div>
                  <div class="aps-tier-body">마스터목표 순증 ${_targetNetMw}만원 — 해당 없음${_needMw > 0 ? ` <span class="aps-tier-gap">(최저 기준까지 ${_needMw}만원 부족)</span>` : ""}</div>
                </div>`;
            }
          }

          // ── 지난 스텝 개인 순증 시상 결과 (Step 2에서만 표시) ──
          let _prevStepHtml = "";
          if (_step === "2") {
            const _prevPa = getProgressAwardConfig(_region, _cohort, "1");
            if (_prevPa.tiers.length > 0) {
              const _step1Current = Number(_calcS.pgCurrent !== undefined ? _calcS.pgCurrent : (_calcS.current || 0));
              const _step1Net = _step1Current - Number(_calcS.base || 0);
              const _prevTierHit = _prevPa.tiers.find(t => _step1Net >= t.min);
              if (_prevTierHit) {
                const _prevPrize = _prevTierHit.type === "pct"
                  ? `${Math.round(_step1Net * _prevTierHit.val / 10000)}만원`
                  : `${_prevTierHit.payVal}만원`;
                _prevStepHtml = `<div class="aps-prev-step aps-prev-hit">🎉 지난달 스텝1에서는 시상금 <strong>${escapeHtml(_prevPrize)}</strong>을 획득 하셨습니다.</div>`;
              } else {
                _prevStepHtml = `<div class="aps-prev-step aps-prev-miss">😅 지난달엔 아쉽지만 개인시상을 획득하지 못하셨습니다. 이달에는 더 화이팅!!</div>`;
              }
            }
          }

          const mkCard = (c) => `
            <div class="aps-award-card${c.isItem ? " aps-item-prize" : ""}${c.isTeam ? " aps-team-prize" : ""}">
              <div class="aps-awd-head">${c.icon} ${escapeHtml(c.label)}</div>
              <div class="aps-awd-body">
                <div class="aps-awd-stat">${escapeHtml(c.sub)}</div>
                <div class="aps-awd-result">${escapeHtml(c.prize)}<span class="aps-awd-near"> 근접</span></div>
              </div>
            </div>`;

          // 시상 미해당 → 최소 시상 진입까지 부족분 안내
          let noAwardHtml = "";
          if (!hasAward) {
            const gapCards = [];
            if (_pa.rateConfig?.enabled && Number(_pa.rateConfig.n) > 0) {
              const n = Number(_pa.rateConfig.n);
              if (_byRate.length >= n) {
                const rateGap = _byRate[n - 1].rate - _statMe.rate;
                if (rateGap > 0) {
                  const _p = _pa.rateConfig?.payouts?.[n - 1];
                  const prz = _p != null ? payoutLabel(_p) : `Top${n} 시상`;
                  const _np = _p ? normPayout(_p) : null;
                  gapCards.push({
                    icon: "📈", category: `신장률 Top${n}`,
                    gapLine: `현재 ${_statMe.rate.toFixed(1)}% · ${n}위 기준 ${_byRate[n-1].rate.toFixed(1)}%`,
                    gapShort: `${rateGap.toFixed(1)}%p 부족`, prize: prz, isItem: _np?.type === "item"
                  });
                }
              }
            }
            if (_pa.amtConfig?.enabled && Number(_pa.amtConfig.n) > 0) {
              const n = Number(_pa.amtConfig.n);
              if (_byAmt.length >= n) {
                const netGap = Math.ceil((_byAmt[n - 1].net - _statMe.net) / 10000);
                if (netGap > 0) {
                  const _p = _pa.amtConfig?.payouts?.[n - 1];
                  const prz = _p != null ? payoutLabel(_p) : `Top${n} 시상`;
                  const _np = _p ? normPayout(_p) : null;
                  gapCards.push({
                    icon: "💰", category: `신장액 Top${n}`,
                    gapLine: `현재 +${Math.round(_statMe.net / 10000)}만원 · ${n}위 기준 +${Math.round(_byAmt[n-1].net / 10000)}만원`,
                    gapShort: `${netGap}만원 부족`, prize: prz, isItem: _np?.type === "item"
                  });
                }
              }
            }
            const mkGapCard = (g) => `
              <div class="aps-gap-card${g.isItem ? " aps-item-prize" : ""}">
                <div class="aps-gap-head">${g.icon} ${escapeHtml(g.category)}</div>
                <div class="aps-gap-body">${g.gapLine}</div>
                <div class="aps-gap-result">${escapeHtml(g.gapShort)} 추가 시<strong>${escapeHtml(g.prize)} 근접</strong></div>
              </div>`;
            noAwardHtml = gapCards.length
              ? `<div class="aps-gap-list">${gapCards.map(mkGapCard).join("")}</div>`
              : `<div class="aps-no-award">현재 순위로는 시상 대상에 해당하지 않습니다</div>`;
          }

          if (!hasAnyConfig && !_pa.tiers.length) return "";

          return `
            <div class="aps-wrap">
              <div class="aps-header">④ 예상 시상 <span class="aps-badge">${escapeHtml(_cohortLabel)} ${escapeHtml(_stepLabel)}</span></div>
              ${_tierSectionHtml}
              ${_prevStepHtml}
              ${hasAnyConfig
                ? (hasAward
                    ? `<div class="aps-award-cards">${_allCards.map(mkCard).join("")}</div>`
                    : noAwardHtml)
                : ""}
              ${_plan.notes ? `<div class="aps-notes">${escapeHtml(_plan.notes)}</div>` : ""}
            </div>`;
        })()}

        <div class="rc-total-box">
          <div>
            <div class="tb-cap">고객마스터 2개월 최종 예상 시상금 합계</div>
            <div class="tb-det">아너스 ${fmtW(award1)} + 하이포 ${fmtW(award2M3)} + 마스터 ${fmtW(award3)}×2 = ${fmtW(award3 * 2)}</div>
          </div>
          <div class="tb-amt">${fmtW(total)}</div>
        </div>

        ${award3NextTier
          ? `<div class="next-tier">
              <div class="nt-head">🚀 다음 단계 달성 목표 — ${escapeHtml(award3NextTier.criteria)} (${escapeHtml(award3NextTier.label)})</div>
              <div class="nt-body">
                <div><div class="nt-lbl">희망목표 추가 필요</div><div class="nt-cur">현재 ${fmtWon(tgtRaw)} → ${fmtWon(tgtRaw + award3NeedTgt)}</div></div>
                <div class="nt-amt">+${fmtWon(award3NeedTgt)}</div>
              </div>
            </div>`
          : `<div class="nt-max">🏆 최상위 단계 달성!</div>`}

        <div class="award-print-row">
          <button id="btn-award-print" class="btn-award-print" type="button">🖨️ 시상인쇄</button>
        </div>
      </div>
    `;
    state.lastCalcResult = { tgtRaw, baseTgtRaw, avgRaw, award1, award1Grade, award1Idx, award2M3, monthlyFinal, award3, award3Tier, total, incrMaster, insRaw3, incr };
    const _printBtn = res.querySelector("#btn-award-print");
    if (_printBtn) _printBtn.addEventListener("click", openAwardPrintPreview);
  }

  function updateIvTitle() {
    const n = ($("#iv-seq").value || "").trim();
    $("#iv-title").innerHTML =
      `MASTER과정 면담일지 (${n ? escapeHtml(n) : "&nbsp;"}차 / 활동점검)`;
  }

  function calcIvPct() {
    const curAct = parseFloat($("#iv-curAct")?.value) || 0;
    const s = state.students.find((x) => x.empNo === state.selectedEmpNo);
    const { base } = getProgressStat(s || {});
    if (!curAct || !base) return;
    const pct = ((curAct / base) * 100).toFixed(1);
    $("#iv-pct").value = pct;
    $("#iv-pct-hint").textContent = `▲ 자동 (현재${Nf(curAct)}/기준${Nf(base)})`;
  }

  function autoFillInterviewForm(s) {
    if (!s) return;
    // 차수: 기존 최대 차수 + 1 (숫자 파싱 후 비교, 문자열 정렬 회피)
    const seqEl = $("#iv-seq");
    if (seqEl && !seqEl.value) {
      const maxSeq = state.consultations.reduce((m, c) => {
        const n = parseInt(c.seq, 10);
        return Number.isFinite(n) && n > m ? n : m;
      }, 0);
      seqEl.value = String(maxSeq + 1);
      updateIvTitle();
    }

    // 현재실적: 실적진도현황 pgCurrent 우선, 없으면 s.curAct 복원
    const curActEl = $("#iv-curAct");
    if (curActEl && !curActEl.value) {
      const { current } = getProgressStat(s);
      if (current > 0) {
        curActEl.value = current;
      } else if (Number(s.curAct) > 0) {
        curActEl.value = Math.round(Number(s.curAct));
      }
    }

    calcIvPct();

    // 시상 계산기: 직전 consultation 의 calc 값 → 없으면 student.base/honors prefill
    const avgEl = $("#calc-avg");
    const baseTgtEl = $("#calc-base-tgt");
    const tgtCalcEl = $("#calc-tgt");
    const lastCalc = state.consultations.find(
      (c) => c.calcAvg || c.calcBaseTgt || c.calcTgt
    );
    if (lastCalc) {
      if (avgEl && !avgEl.value && lastCalc.calcAvg) avgEl.value = lastCalc.calcAvg;
      if (baseTgtEl && !baseTgtEl.value && lastCalc.calcBaseTgt) baseTgtEl.value = lastCalc.calcBaseTgt;
      if (tgtCalcEl && !tgtCalcEl.value && lastCalc.calcTgt) {
        const raw = parseFloat(String(lastCalc.calcTgt).replace(/,/g, "")) || 0;
        const fixed = (raw > 0 && raw < 1000) ? raw * 1000 : raw;
        tgtCalcEl.value = fixed ? Math.round(fixed).toLocaleString() : lastCalc.calcTgt;
      }
    } else {
      if (avgEl && !avgEl.value && Number(s.base) > 0) avgEl.value = Number(s.base).toLocaleString();
      if (baseTgtEl && !baseTgtEl.value && Number(s.honors) > 0) baseTgtEl.value = Number(s.honors).toLocaleString();
    }
    // 희망목표금액 미입력 시 마스터목표로 자동 세팅
    if (tgtCalcEl && !tgtCalcEl.value && Number(s.target) > 0) {
      tgtCalcEl.value = Number(s.target).toLocaleString();
    }
    if (state.calcOpen) calc();
  }

  function clearInterviewForm() {
    ["iv-seq","iv-curAct","iv-pct","iv-plan","iv-hap","iv-exp","iv-close1","iv-close2","iv-coach"]
      .forEach((id) => { const el = $("#" + id); if (el) el.value = ""; });
    const today = new Date().toISOString().slice(0, 10);
    const d = $("#iv-date"); if (d) d.value = today;
    const phint = $("#iv-pct-hint"); if (phint) phint.textContent = "";
    initCR([]); // 상담고객 리셋
    // 시상 계산기 필드 리셋
    ["calc-avg","calc-base-tgt","calc-tgt"].forEach((id) => {
      const el = $("#" + id); if (el) el.value = "";
    });
    const preview = $("#calc-incr-preview"); if (preview) preview.style.display = "none";
    const res = $("#calc-result");
    if (res) res.innerHTML = `<div class="rc-placeholder">기본 입력 항목을 입력하면 시상금이 자동 계산됩니다</div>`;
    updateIvTitle();
  }

  function buildInterviewRecord() {
    const read = (id) => ($("#" + id)?.value || "").trim();
    const num = (id) => {
      const v = read(id);
      if (!v) return 0;
      const n = Number(v.replace(/[,\s]/g, ""));
      return Number.isFinite(n) ? n : 0;
    };
    return {
      date: read("iv-date"),
      seq: read("iv-seq"),
      pct: num("iv-pct"),
      curAct: Math.round(num("iv-curAct")),
      plan: num("iv-plan"),
      hap: num("iv-hap"),
      exp: Math.round(num("iv-exp")),
      close1: num("iv-close1"),
      close2: num("iv-close2"),
      coach: read("iv-coach"),
      clients: (state.crData || []).map((c) => ({
        name: (c.name || "").trim(),
        types: c.types || [],
        consult: c.consult || [],
        material: c.material || [],
        amount: c.amount || [],
        amountDirect: c.amountDirect || "",
        bj: c.bj || [],
        memo: c.memo || ""
      })),
      calcAvg: read("calc-avg"),
      calcBaseTgt: read("calc-base-tgt"),
      calcTgt: read("calc-tgt")
    };
  }

  async function saveInterview() {
    const empNo = state.selectedEmpNo;
    if (!empNo) return;
    const rec = buildInterviewRecord();
    if (!rec.date) { toast("면담일시를 입력하세요.", "error"); return; }
    // 상담고객 없이 저장 시 확인
    const hasClients = rec.clients.some((c) => c.name || c.memo || (c.amount && c.amount.length));
    if (!hasClients && !state.editingConsultId) {
      const ok = await confirmSaveWithoutClients();
      if (!ok) return;
    }
    // 담당자 선택 팝업
    const managerName = await openManagerPickerModal();
    if (managerName === null) return; // 취소
    rec.manager = managerName;
    await doSaveInterview(rec);
  }

  function openManagerPickerModal() {
    return new Promise((resolve) => {
      let modal = document.getElementById("modal-manager-pick");
      if (!modal) {
        modal = document.createElement("div");
        modal.id = "modal-manager-pick";
        modal.className = "modal";
        modal.hidden = true;
        modal.innerHTML = `
          <div class="modal-backdrop" id="mgr-backdrop"></div>
          <div class="modal-panel" style="max-width:340px">
            <div class="modal-head">
              <h3 style="font-size:15px">담당자 입력</h3>
              <button class="modal-close" id="mgr-close">&times;</button>
            </div>
            <div class="modal-body" style="padding:16px 20px">
              <p style="font-size:13px;color:#666;margin:0 0 12px">결재란에 표시할 담당자 이름을 입력하세요.</p>
              <input type="text" id="mgr-name-input" class="side-input" style="width:100%;font-size:14px;padding:8px 10px;box-sizing:border-box" placeholder="예) 홍길동">
            </div>
            <div style="padding:10px 16px 14px;display:flex;gap:8px;justify-content:flex-end">
              <button class="btn-outline small" id="mgr-skip">건너뛰기</button>
              <button class="btn-primary small" id="mgr-confirm">확인</button>
            </div>
          </div>
        `;
        document.body.appendChild(modal);
      }
      const input = modal.querySelector("#mgr-name-input");
      input.value = localStorage.getItem("iv_manager_name") || "";
      modal.hidden = false;
      setTimeout(() => { input.focus(); input.select(); }, 80);

      const done = (name) => {
        modal.hidden = true;
        if (name !== null && name !== "") localStorage.setItem("iv_manager_name", name);
        resolve(name);
      };
      const confirm = () => done(input.value.trim());
      const skip    = () => { modal.hidden = true; resolve(null); };
      const cancel  = () => { modal.hidden = true; resolve(null); };

      modal.querySelector("#mgr-confirm").onclick = confirm;
      modal.querySelector("#mgr-skip").onclick    = skip;
      modal.querySelector("#mgr-close").onclick   = cancel;
      modal.querySelector("#mgr-backdrop").onclick = cancel;
      input.onkeydown = (e) => {
        if (e.key === "Enter") confirm();
        if (e.key === "Escape") cancel();
      };
    });
  }

  async function doSaveInterview(rec) {
    const empNo = state.selectedEmpNo;
    const btn = $("#btn-iv-save");
    const origLabel = btn ? btn.textContent : "";
    if (btn) { btn.disabled = true; btn.textContent = "저장중..."; }
    try {
      if (state.editingConsultId) {
        await window.DataAPI.updateConsultation(empNo, state.editingConsultId, rec);
        toast("면담 기록이 수정되었습니다.", "success");
      } else {
        await window.DataAPI.addConsultation(empNo, rec);
        toast("면담 기록이 저장되었습니다.", "success");
      }
      if (rec.ins > 0 && typeof window.DataAPI.updateStudentInsAvg === "function") {
        window.DataAPI.updateStudentInsAvg(empNo, rec.ins).catch((e) => {
          console.warn("[insAvg sync]", e);
        });
      }
      // 팀(조) 번호 변경 시 학생 레코드에 저장
      // 팀(조) 번호: 값이 있을 때만 저장 — 빈 값으로는 기존 팀 데이터를 절대 덮어쓰지 않음
      const teamNum = parseInt(($("#iv-team")?.value || "").trim(), 10);
      if (teamNum > 0) {
        const teamToSave = String(teamNum);
        const curStudent = state.students.find((x) => x.empNo === empNo);
        if (curStudent && teamToSave !== (curStudent.team || "")) {
          window.DataAPI.save({ ...curStudent, team: teamToSave }).catch((e) => console.warn("[team sync]", e));
        }
      }
      state.editingConsultId = null;
      clearInterviewForm();
      const s = state.students.find((x) => x.empNo === empNo);
      if (s) autoFillInterviewForm(s);
      updateSaveButtonLabel();
    } catch (err) {
      console.error(err);
      toast("저장 실패: " + err.message, "error");
      if (btn) { btn.disabled = false; btn.textContent = origLabel || "💾 저장"; }
      return;
    }
    if (btn) { btn.disabled = false; btn.textContent = "💾 저장"; }
  }

  function confirmSaveWithoutClients() {
    return new Promise((resolve) => {
      const msgEl = $("#dup-msg");
      const oldTbl = $("#dup-old");
      const newTbl = $("#dup-new");
      const titleEl = document.querySelector("#modal-duplicate .modal-head h3");
      const okBtn = $("#btn-dup-overwrite");
      if (!msgEl || !okBtn) { resolve(true); return; }
      // 재사용: modal-duplicate 를 확인 대화상자로 용도 변경
      if (titleEl) titleEl.textContent = "상담고객 정보 없음";
      msgEl.textContent = "상담고객 정보가 없습니다. 코칭포인트만 저장하시겠습니까?";
      if (oldTbl) oldTbl.innerHTML = "";
      if (newTbl) newTbl.innerHTML = "";
      // 비교표 숨김
      const compare = document.querySelector("#modal-duplicate .dup-compare");
      if (compare) compare.style.display = "none";
      okBtn.textContent = "예, 저장";
      openModal("#modal-duplicate");

      const cleanup = () => {
        okBtn.removeEventListener("click", onYes);
        $$("#modal-duplicate [data-close]").forEach((el) => el.removeEventListener("click", onNo));
        // 원복
        if (compare) compare.style.display = "";
        okBtn.textContent = "덮어쓰기";
        if (titleEl) titleEl.textContent = "이미 등록된 사번입니다";
      };
      const onYes = () => { cleanup(); closeModal("#modal-duplicate"); resolve(true); };
      const onNo = () => { cleanup(); closeModal("#modal-duplicate"); resolve(false); };
      okBtn.addEventListener("click", onYes);
      $$("#modal-duplicate [data-close]").forEach((el) => el.addEventListener("click", onNo));
    });
  }

  // 이력 렌더: form 서브뷰에서는 #consult-history (없으면 생략),
  // history 서브뷰에서는 #hist-list 에 풍부한 카드로 렌더
  function renderConsultations() {
    // 서브탭 배지 갱신
    const cnt = document.getElementById("hist-cnt");
    if (cnt) cnt.textContent = state.consultations.length;

    // history 서브뷰: 풍부한 카드
    if (state.studentSubView === "history") {
      renderHistoryCards();
      return;
    }
    // 아래는 (현재 제거된) form 뷰의 컴팩트 리스트 경로 — noop
    const container = $("#consult-history");
    if (!container) return;
    if (state.consultations.length === 0) {
      container.innerHTML = `<div class="empty-mini">등록된 면담 기록이 없습니다.</div>`;
      return;
    }
    const fmt = (n) => Number(n || 0).toLocaleString();
    container.innerHTML = state.consultations.map((c) => {
      const seqLabel = c.seq ? `${escapeHtml(c.seq)}차` : "";
      const badges = [];
      if (c.ins)    badges.push(`<span class="cs-badge">인보험 ${fmt(Number(c.ins)*1000)}</span>`);
      if (c.tgt)    badges.push(`<span class="cs-badge">당월 ${fmt(Number(c.tgt)*1000)}</span>`);
      if (c.curAct) badges.push(`<span class="cs-badge">현재 ${fmt(Number(c.curAct))}</span>`);
      if (c.pct)    badges.push(`<span class="cs-badge blue">진도 ${fmt(c.pct)}%</span>`);
      if (c.plan)   badges.push(`<span class="cs-badge">가입 ${fmt(c.plan)}</span>`);
      if (c.hap)    badges.push(`<span class="cs-badge">행복 ${fmt(c.hap)}</span>`);
      if (c.exp)    badges.push(`<span class="cs-badge">예상 ${fmt(Number(c.exp))}</span>`);
      const body = c.coach || c.content || "";
      const clients = Array.isArray(c.clients) ? c.clients.filter((cl) => cl.name || cl.memo || (cl.amount && cl.amount.length)) : [];
      const clientHtml = clients.length ? `
        <div class="consult-clients">
          <div class="cc-title">상담고객 ${clients.length}명</div>
          ${clients.map((cl) => {
            const tags = [
              ...(cl.types || []), ...(cl.consult || []),
              ...(cl.material || []), ...(cl.amount || []),
              ...(cl.bj || [])
            ];
            const amt = cl.amountDirect ? `${cl.amountDirect}만원(직접)` : "";
            return `
              <div class="cc-row">
                <span class="cc-name">${escapeHtml(cl.name || "(무기명)")}</span>
                ${tags.length ? `<span class="cc-tags">${tags.map((t) => `<span class="cc-tag">${escapeHtml(t)}</span>`).join("")}</span>` : ""}
                ${amt ? `<span class="cc-amt">${escapeHtml(amt)}</span>` : ""}
                ${cl.memo ? `<div class="cc-memo">${escapeHtml(cl.memo)}</div>` : ""}
              </div>
            `;
          }).join("")}
        </div>
      ` : "";
      const isEditing = state.editingConsultId === c.id;
      return `
        <div class="consult-entry${isEditing ? " editing" : ""}" data-id="${escapeHtml(c.id)}">
          <div class="consult-head">
            <span class="consult-date">${escapeHtml(c.date || "")}${seqLabel ? " · " + seqLabel : ""}${isEditing ? ' <span class="edit-badge">수정 중</span>' : ""}</span>
            <div class="consult-actions">
              <button class="consult-print" data-id="${escapeHtml(c.id)}" title="인쇄">인쇄</button>
              <button class="consult-edit" data-id="${escapeHtml(c.id)}" title="수정">수정</button>
              <button class="consult-del" data-id="${escapeHtml(c.id)}" title="삭제">×</button>
            </div>
          </div>
          ${badges.length ? `<div class="consult-summary">${badges.join("")}</div>` : ""}
          ${clientHtml}
          ${body ? `<div class="consult-body">${escapeHtml(body)}</div>` : ""}
        </div>
      `;
    }).join("");
    container.querySelectorAll(".consult-del").forEach((btn) => {
      btn.addEventListener("click", () => removeConsultation(btn.dataset.id));
    });
    container.querySelectorAll(".consult-edit").forEach((btn) => {
      btn.addEventListener("click", () => editInterview(btn.dataset.id));
    });
    container.querySelectorAll(".consult-print").forEach((btn) => {
      btn.addEventListener("click", () => printConsultation(btn.dataset.id));
    });
  }

  // 면담 이력 — 풍부한 카드 레이아웃 (면담이력 서브탭)
  // 댓글 역할 선택지 (향후 결제 연동 대상)
  const COMMENT_ROLES = ["비전센터장", "전임강사", "조직파트장", "지역단장"];

  // 면담 이력 카드 하단의 댓글 섹션 렌더
  function renderConsultComments(e) {
    const comments = Array.isArray(e.comments) ? e.comments : [];
    const nl = (s) => escapeHtml(s || "").replace(/\n/g, "<br>");
    const fmtDate = (iso) => {
      if (!iso) return "";
      try {
        const d = new Date(iso);
        return d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0") + "-" + String(d.getDate()).padStart(2,"0") +
               " " + String(d.getHours()).padStart(2,"0") + ":" + String(d.getMinutes()).padStart(2,"0");
      } catch { return iso; }
    };
    return `
      <div class="cm-sec">
        <div class="cm-hd">
          <strong>💬 댓글 <span class="cm-cnt">${comments.length}</span></strong>
          <button class="btn-outline small cm-add-btn" data-id="${escapeHtml(e.id)}">✏️ 댓글 달기</button>
        </div>
        ${comments.length ? `<ul class="cm-list">
          ${comments.map((c) => `
            <li class="cm-item cm-role-${escapeHtml(c.role || "")}">
              <div class="cm-meta">
                <span class="cm-role">${escapeHtml(c.role || "")}</span>
                <span class="cm-by">${escapeHtml(c.author || "")}</span>
                <span class="cm-dt">${escapeHtml(fmtDate(c.createdAt))}</span>
                <button class="cm-del-btn" data-consult-id="${escapeHtml(e.id)}" data-comment-id="${escapeHtml(c.id)}" title="삭제">🗑</button>
              </div>
              <div class="cm-txt">${nl(c.text)}</div>
            </li>
          `).join("")}
        </ul>` : `<div class="cm-empty">아직 댓글이 없습니다. 첫 댓글을 남겨보세요.</div>`}
      </div>
    `;
  }

  // 댓글 작성 모달
  function openCommentModal(consultationId) {
    let modal = document.getElementById("modal-comment");
    if (!modal) {
      modal = document.createElement("div");
      modal.id = "modal-comment";
      modal.className = "modal";
      modal.hidden = true;
      modal.innerHTML = `
        <div class="modal-backdrop" data-close></div>
        <div class="modal-panel">
          <div class="modal-head">
            <h3>💬 댓글 달기</h3>
            <button class="modal-close" data-close aria-label="닫기">&times;</button>
          </div>
          <div class="modal-body">
            <div class="cm-form">
              <label class="cm-label">역할 (작성자 직책) <em class="req">*</em></label>
              <div class="cm-role-chips">
                ${COMMENT_ROLES.map((r) => `<button type="button" class="cm-role-chip" data-role="${escapeHtml(r)}">${escapeHtml(r)}</button>`).join("")}
              </div>
              <label class="cm-label" style="margin-top:14px;">작성자 이름 <span class="cm-hint">(선택)</span></label>
              <input type="text" id="cm-author" class="side-input" placeholder="예) 박부장">
              <label class="cm-label" style="margin-top:14px;">댓글 내용 <em class="req">*</em></label>
              <textarea id="cm-text" rows="5" class="cm-textarea" placeholder="의견을 남겨주세요"></textarea>
              <p class="cm-note">※ 결제 창 연동 시 역할·이름·작성일시가 승인 이력으로 기록됩니다.</p>
            </div>
          </div>
          <div class="modal-foot">
            <button class="btn-outline" data-close>취소</button>
            <button class="btn-primary" id="btn-cm-save">💾 저장</button>
          </div>
        </div>
      `;
      document.body.appendChild(modal);
      modal.querySelectorAll("[data-close]").forEach((el) => el.addEventListener("click", () => { modal.hidden = true; }));
      // 역할 칩 토글
      modal.querySelectorAll(".cm-role-chip").forEach((chip) => {
        chip.addEventListener("click", () => {
          modal.querySelectorAll(".cm-role-chip").forEach((c) => c.classList.remove("on"));
          chip.classList.add("on");
          modal.dataset.selectedRole = chip.dataset.role;
        });
      });
      // 저장
      modal.querySelector("#btn-cm-save").addEventListener("click", async () => {
        const role = modal.dataset.selectedRole;
        const text = modal.querySelector("#cm-text").value.trim();
        const author = modal.querySelector("#cm-author").value.trim();
        if (!role) { toast("역할을 선택하세요.", "error"); return; }
        if (!text) { toast("댓글 내용을 입력하세요.", "error"); modal.querySelector("#cm-text").focus(); return; }
        const consultId = modal.dataset.consultId;
        const empNo = state.selectedEmpNo;
        if (!consultId || !empNo) { toast("대상 면담을 찾을 수 없습니다.", "error"); return; }
        const saveBtn = modal.querySelector("#btn-cm-save");
        saveBtn.disabled = true;
        try {
          await window.DataAPI.addConsultationComment(empNo, consultId, {
            id: "cm_" + Date.now() + "_" + Math.random().toString(36).slice(2, 6),
            role, author, text,
            createdAt: new Date().toISOString()
          });
          toast("댓글이 등록되었습니다.", "success");
          modal.hidden = true;
        } catch (err) {
          toast("댓글 저장 실패: " + (err.message || err), "error");
        }
        saveBtn.disabled = false;
      });
    }
    // 모달 초기화
    modal.dataset.consultId = consultationId;
    modal.dataset.selectedRole = "";
    modal.querySelectorAll(".cm-role-chip").forEach((c) => c.classList.remove("on"));
    modal.querySelector("#cm-text").value = "";
    modal.querySelector("#cm-author").value = "";
    modal.hidden = false;
  }

  function renderHistoryCards() {
    const el = document.getElementById("hist-list");
    if (!el) return;
    const list = state.consultations.slice();
    if (!list.length) {
      el.innerHTML = `<div class="empty-state"><div class="empty-ico">📂</div>면담 기록이 없습니다</div>`;
      return;
    }
    const fn = (v) => {
      const n = parseInt(v, 10);
      return Number.isFinite(n) && n > 0 ? n.toLocaleString() : "-";
    };
    const nl = (s) => escapeHtml(s || "").replace(/\n/g, "<br>");

    el.innerHTML = list.map((e) => {
      const clients = Array.isArray(e.clients) ? e.clients : [];
      const clientsHtml = clients.length ? `
        <table class="hi-ct">
          <thead><tr>
            <th>#</th><th>성명</th><th>고객유형</th><th>상담단계</th>
            <th>활용자료</th><th>제안금액</th><th>보종</th><th>면담내용</th>
          </tr></thead>
          <tbody>
            ${clients.map((c, ci) => `
              <tr>
                <td data-label="#">${ci + 1}</td>
                <td data-label="성명" class="cc-name-td">${escapeHtml(c.name || "")}</td>
                <td data-label="고객유형">${escapeHtml((c.types || []).join(", ")) || "-"}</td>
                <td data-label="상담단계">${escapeHtml((c.consult || []).join(", ")) || "-"}</td>
                <td data-label="활용자료">${escapeHtml((c.material || []).join(", ")) || "-"}</td>
                <td data-label="제안금액" class="hi-ct-amt">${escapeHtml((c.amount || []).join(", ")) || "-"}${c.amountDirect ? ` <span class="amt-direct">+${escapeHtml(c.amountDirect)}만</span>` : ""}</td>
                <td data-label="보종" class="hi-ct-bj">${escapeHtml((c.bj || []).join(", ")) || "-"}</td>
                <td data-label="면담내용" class="cc-memo-td">${escapeHtml(c.memo || "")}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>` : "";

      const calcHtml = (e.calcAvg || e.calcBaseTgt || e.calcTgt) ? `
        <div class="hi-calc">
          ${e.calcAvg     ? `<span>평균실적: <strong>${escapeHtml(e.calcAvg)}</strong>원</span>`       : ""}
          ${e.calcBaseTgt ? `<span>기본순증목표: <strong>${escapeHtml(e.calcBaseTgt)}</strong>원</span>` : ""}
          ${e.calcTgt     ? `<span>희망목표: <strong>${escapeHtml(e.calcTgt)}</strong>원</span>`       : ""}
        </div>` : "";

      const isEditing = state.editingConsultId === e.id;
      return `
        <div class="hi${isEditing ? " editing" : ""}" data-id="${escapeHtml(e.id)}">
          <div class="hi-hd">
            <div class="hi-meta">
              <span class="hi-dt">${escapeHtml(e.date || "")}</span>
              ${e.seq ? `<span class="hi-bg">${escapeHtml(e.seq)}차</span>` : ""}
              ${isEditing ? `<span class="edit-badge">수정 중</span>` : ""}
            </div>
            <div class="hi-btns">
              <button class="ib" data-act="print" data-id="${escapeHtml(e.id)}" title="인쇄">🖨️</button>
              <button class="ib" data-act="edit"  data-id="${escapeHtml(e.id)}" title="수정">✏️</button>
              <button class="ib d" data-act="del" data-id="${escapeHtml(e.id)}" title="삭제">🗑️</button>
            </div>
          </div>
          <div class="hi-bd">
            <div class="hi-fg">
              <div class="hf"><div class="hfl">평균실적(6개월)</div><div class="hfv">${fn(Number(e.ins||0)*1000)} <small>원</small></div></div>
              <div class="hf"><div class="hfl">마스터 목표</div><div class="hfv">${fn(Number(e.tgt||0)*1000)} <small>원</small></div></div>
              <div class="hf"><div class="hfl">현재실적</div><div class="hfv">${fn(Number(e.curAct||0))} <small>원</small></div></div>
              <div class="hf"><div class="hfl">진도</div><div class="hfv">${fn(e.pct)} <small>%</small></div></div>
              <div class="hf"><div class="hfl">가입설계</div><div class="hfv">${fn(e.plan)} <small>건</small></div></div>
              <div class="hf"><div class="hfl">행복보장</div><div class="hfv">${fn(e.hap)} <small>건</small></div></div>
              <div class="hf"><div class="hfl">주간예상</div><div class="hfv">${fn(Number(e.exp||0))} <small>원</small></div></div>
              <div class="hf"><div class="hfl">1차 마감</div><div class="hfv">${fn(e.close1)} <small>원</small></div></div>
              <div class="hf"><div class="hfl">2차 마감</div><div class="hfv">${fn(e.close2)} <small>원</small></div></div>
            </div>
            ${clientsHtml}
            ${e.coach ? `<div class="note nv"><strong>📌 코칭포인트</strong><p>${nl(e.coach)}</p></div>` : ""}
            ${e.calcComment ? `<div class="note blue"><strong>✍️ 면담자 의견</strong><p>${nl(e.calcComment)}</p></div>` : ""}
            ${calcHtml}
            ${renderConsultComments(e)}
          </div>
        </div>
      `;
    }).join("");

    // 댓글 추가 버튼
    el.querySelectorAll(".cm-add-btn").forEach((btn) => {
      btn.addEventListener("click", () => openCommentModal(btn.dataset.id));
    });
    // 댓글 삭제 버튼
    el.querySelectorAll(".cm-del-btn").forEach((btn) => {
      btn.addEventListener("click", async () => {
        if (!confirm("댓글을 삭제하시겠습니까?")) return;
        try {
          await window.DataAPI.removeConsultationComment(
            state.selectedEmpNo, btn.dataset.consultId, btn.dataset.commentId
          );
          toast("댓글 삭제 완료", "success");
        } catch (err) {
          toast("댓글 삭제 실패: " + err.message, "error");
        }
      });
    });

    el.querySelectorAll(".ib").forEach((btn) => {
      const id = btn.dataset.id;
      const act = btn.dataset.act;
      btn.addEventListener("click", () => {
        if (act === "print") printConsultation(id);
        else if (act === "edit") {
          setStudentSubView("form"); // 수정하러 폼으로 이동
          editInterview(id);
        } else if (act === "del") removeConsultation(id);
      });
    });
  }

  // ========== 면담 인쇄 ==========
  function printConsultation(consultId) {
    const c = state.consultations.find((x) => x.id === consultId);
    const s = state.students.find((x) => x.empNo === state.selectedEmpNo);
    if (!c || !s) { toast("인쇄할 면담을 찾을 수 없습니다.", "error"); return; }
    const fmt = (n) => Number(n || 0).toLocaleString();
    const clients = Array.isArray(c.clients) ? c.clients.filter((cl) => cl.name || cl.memo || (cl.amount && cl.amount.length)) : [];

    // 시상 계산 (c 자체 + 이력 전체 활용)
    const allItvs = state.consultations.slice();
    const lastInsItv = allItvs.slice().sort((a, b) => (b.date || "").localeCompare(a.date || "")).find((x) => x.ins);
    const b = computeAwardBreakdown(s, c, lastInsItv);

    const clientsHtml = clients.length ? `
      <section class="pr-section">
        <h3>상담고객 (${clients.length}명)</h3>
        <table class="pr-table pr-cc-table">
          <colgroup>
            <col style="width:3%;">
            <col style="width:7%;">
            <col style="width:6%;">
            <col style="width:7%;">
            <col style="width:9%;">
            <col style="width:6%;">
            <col style="width:5%;">
            <col style="width:57%;">
          </colgroup>
          <thead><tr>
            <th>#</th><th>성명</th><th>유형</th><th>단계</th><th>자료</th><th>금액</th><th>보종</th><th>면담 내용</th>
          </tr></thead>
          <tbody>
            ${clients.map((cl, i) => `
              <tr>
                <td>${i + 1}</td>
                <td>${escapeHtml(cl.name || "")}</td>
                <td>${escapeHtml((cl.types || []).join(", "))}</td>
                <td>${escapeHtml((cl.consult || []).join(", "))}</td>
                <td>${escapeHtml((cl.material || []).join(", "))}</td>
                <td>${escapeHtml((cl.amount || []).join(", "))}${cl.amountDirect ? ` / ${escapeHtml(cl.amountDirect)}만` : ""}</td>
                <td>${escapeHtml((cl.bj || []).join(", "))}</td>
                <td class="memo">${escapeHtml(cl.memo || "")}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </section>
    ` : "";

    const award1Html = b.award1Idx >= 0
      ? `<div class="aw-row"><span class="aw-ic">🏆</span><div class="aw-body"><div class="aw-t">① 아너스클럽 · ${escapeHtml(b.award1Grade)}</div><div class="aw-sub">${escapeHtml(HONORS[b.award1Idx].criteria)}</div></div><div class="aw-v">${fmtW(b.award1)}</div></div>`
      : `<div class="aw-row none">① 아너스클럽 — 해당없음</div>`;
    const award2Html = b.baseTgtMet
      ? `<div class="aw-row"><span class="aw-ic">💧</span><div class="aw-body"><div class="aw-t">② 하이포인트 (3개월)</div><div class="aw-sub">기본 5만 + 순증 ${fmtW(b.incr)}×50% = 월 ${fmtWon(b.monthlyFinal*10000)}</div></div><div class="aw-v">${fmtW(b.award2M3)}</div></div>`
      : `<div class="aw-row warn"><span class="aw-ic">⚠️</span><div class="aw-body"><div class="aw-t">② 하이포인트 — 미해당</div><div class="aw-sub">희망목표 ${fmtWon(b.tgtRaw)} < 기본순증 ${fmtWon(b.baseTgtRaw)}</div></div><div class="aw-v rd">0원</div></div>`;
    const award3Html = b.award3 > 0
      ? `<div class="aw-row blue"><span class="aw-ic">🎯</span><div class="aw-body"><div class="aw-t">③ 마스터 · ${escapeHtml(b.award3Tier.criteria)}</div><div class="aw-sub">순증 ${fmtWon(b.incrMaster*10000)} / ${escapeHtml(b.award3Tier.label)} · ×2개월</div></div><div class="aw-v bl">${fmtW(b.award3 * 2)}</div></div>`
      : `<div class="aw-row none">③ 마스터 — 순증 5만원 미만 (${fmtWon(b.incrMaster*10000)})</div>`;

    const awardHtml = (b.tgtRaw > 0) ? `
      <section class="pr-section">
        <h3>시상 예상답안지 (희망목표 ${fmtWon(b.tgtRaw)})</h3>
        ${award1Html}${award2Html}${award3Html}
        <div class="aw-total"><div><strong>🏆 최종 예상 합계</strong> &nbsp;<span class="sub">아너스 ${fmtW(b.award1)} + 하이포 ${fmtW(b.award2M3)} + 마스터 ${fmtW(b.award3*2)}</span></div><div class="aw-total-v">${fmtW(b.total)}</div></div>
      </section>
    ` : "";

    const cmtHtml = c.calcComment ? `
      <section class="pr-section tight">
        <div class="pr-comment"><strong>✍️ 면담자 의견</strong> ${escapeHtml(c.calcComment)}</div>
      </section>
    ` : "";

    const html = `<!DOCTYPE html>
<html lang="ko"><head><meta charset="UTF-8">
<title>면담일지 - ${escapeHtml(s.name || "")} ${escapeHtml(c.seq || "")}차</title>
<style>
  @page { size: A4 portrait; margin: 6mm 7mm; }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: "Noto Sans KR", "Malgun Gothic", sans-serif; font-size: 10.5px; color: #1A1A1A; line-height: 1.35; }
  header { border-bottom: 2px solid #E8651A; padding-bottom: 5px; margin-bottom: 7px; display: flex; justify-content: space-between; align-items: baseline; }
  header h1 { font-size: 14px; font-weight: 900; color: #1A2744; letter-spacing: -0.3px; }
  header .sub { font-size: 10px; color: #666; }
  .pr-section { margin-bottom: 7px; page-break-inside: avoid; }
  .pr-section.tight { margin-bottom: 4px; }
  .pr-section h3 { font-size: 11px; font-weight: 800; color: #E8651A; padding: 3px 8px; background: #FFF3EC; border-left: 3px solid #E8651A; margin-bottom: 4px; }
  .pr-info { width: 100%; border-collapse: collapse; font-size: 10px; }
  .pr-info th { background: #F5F5F5; color: #444; font-weight: 700; padding: 2px 6px; text-align: left; border: 1px solid #D5D5D5; white-space: nowrap; }
  .pr-info td { padding: 2px 6px; border: 1px solid #D5D5D5; }
  .pr-table { width: 100%; border-collapse: collapse; font-size: 10px; table-layout: fixed; }
  .pr-table th, .pr-table td { border: 1px solid #D5D5D5; padding: 2px 5px; text-align: left; vertical-align: top; word-break: keep-all; overflow-wrap: break-word; }
  .pr-table th { background: #F5F5F5; font-weight: 700; font-size: 9.5px; }
  .pr-cc-table { table-layout: fixed; }
  .pr-cc-table th, .pr-cc-table td { padding: 3px 4px; font-size: 10px; }
  .pr-cc-table td.memo { white-space: pre-wrap; line-height: 1.35; font-size: 10.5px; word-break: keep-all; overflow-wrap: break-word; }
  .pr-coach { padding: 6px 10px; background: #FFF8F5; border-left: 3px solid #E8651A; border-radius: 0 5px 5px 0; white-space: pre-wrap; font-size: 10.5px; line-height: 1.5; }
  .pr-comment { padding: 4px 8px; background: #F5F7FF; border-left: 3px solid #1A2744; border-radius: 0 5px 5px 0; font-size: 10px; line-height: 1.4; }
  .pr-comment strong { color: #1A2744; margin-right: 4px; }

  /* 시상안 블록 — 압축형 */
  .aw-row { display: flex; align-items: center; gap: 7px; padding: 4px 8px; margin-bottom: 3px; background: #E8F5E9; border-left: 3px solid #2E7D32; border-radius: 0 5px 5px 0; }
  .aw-row.warn { background: #FFEBEE; border-left-color: #C62828; }
  .aw-row.blue { background: #E3F2FD; border-left-color: #1565C0; }
  .aw-row.none { padding: 3px 8px; background: #F5F5F5; color: #999; font-size: 10px; border-left-color: #BBB; }
  .aw-ic { font-size: 14px; flex-shrink: 0; }
  .aw-body { flex: 1; min-width: 0; }
  .aw-t { font-size: 11px; font-weight: 800; color: #1B5E20; }
  .aw-row.warn .aw-t { color: #C62828; }
  .aw-row.blue .aw-t { color: #1565C0; }
  .aw-sub { font-size: 10px; color: #388E3C; margin-top: 1px; }
  .aw-row.warn .aw-sub { color: #E57373; }
  .aw-row.blue .aw-sub { color: #1976D2; }
  .aw-v { font-size: 15px; font-weight: 900; color: #2E7D32; white-space: nowrap; }
  .aw-v.rd { color: #C62828; }
  .aw-v.bl { color: #1565C0; }
  .aw-total { margin-top: 4px; padding: 5px 10px; background: linear-gradient(135deg, #1A2744, #2C3F6E); color: #fff; border-radius: 5px; display: flex; justify-content: space-between; align-items: center; }
  .aw-total .sub { color: rgba(255,255,255,.7); font-size: 9px; }
  .aw-total-v { color: #FFE082; font-size: 17px; font-weight: 900; }

  footer { margin-top: 6px; padding-top: 4px; border-top: 1px solid #D5D5D5; font-size: 9px; color: #999; text-align: center; }

  /* 미리보기용 상단 바 — 화면에서만 노출, 인쇄 시에는 감춤 */
  .print-ovl {
    position: fixed; top: 0; left: 0; right: 0; z-index: 9999;
    background: #1A2744; color: #fff;
    display: flex; align-items: center; justify-content: space-between;
    padding: 10px 14px; box-shadow: 0 2px 8px rgba(0,0,0,0.2);
    font-family: inherit;
  }
  .print-ovl .po-ttl { font-size: 13px; font-weight: 700; color: #fff; letter-spacing: -0.2px; }
  .print-ovl button {
    border: none; border-radius: 6px; padding: 8px 14px;
    font-size: 13px; font-weight: 800; cursor: pointer;
    font-family: inherit; line-height: 1;
  }
  .po-back { background: #E8651A; color: #fff; }
  .po-back:hover { background: #d65910; }
  .po-print { background: #fff; color: #1A2744; }
  .po-print:hover { background: #f0f2f8; }
  body { padding-top: 50px; }
  @media print {
    .print-ovl { display: none !important; }
    body { padding-top: 0 !important; }
  }
</style>
</head><body>
  <header>
    <h1>고객컨설팅 MASTER과정 면담일지 (${escapeHtml(c.seq || "")}차 / 활동점검)</h1>
    <div class="sub">${escapeHtml(c.date || "")} · ${escapeHtml(s.region || "")} · ${escapeHtml(s.center || "")} · ${escapeHtml(s.branch || "")}</div>
  </header>

  <section class="pr-section">
    <h3>기본 정보</h3>
    <table class="pr-info">
      <tr><th>성명</th><td>${escapeHtml(s.name || "")}</td>
          <th>사번</th><td>${escapeHtml(s.empNo)}</td>
          <th>기수</th><td>${escapeHtml(s.cohort || "")}</td>
          <th>연락처</th><td>${escapeHtml(s.phone || "")}</td></tr>
      <tr><th>평균실적(6개월)</th><td>${fmt(Number(c.ins||0)*1000)} 원</td>
          <th>마스터 목표</th><td>${fmt(Number(c.tgt||0)*1000)} 원</td>
          <th>현재실적</th><td>${fmt(Number(c.curAct||0))} 원</td>
          <th>진도</th><td>${fmt(c.pct)} %</td></tr>
      <tr><th>가입설계</th><td>${fmt(c.plan)} 건</td>
          <th>행복보장</th><td>${fmt(c.hap)} 건</td>
          <th>주간예상</th><td>${fmt(Number(c.exp||0))} 원</td>
          <th>차수</th><td>${escapeHtml(c.seq || "")}차</td></tr>
    </table>
  </section>

  ${clientsHtml}

  <section class="pr-section">
    <h3>핵심 코칭포인트 / 후속조치 / 다음주 계획</h3>
    <div class="pr-coach">${escapeHtml(c.coach || c.content || "-")}</div>
  </section>

  ${awardHtml}
  ${cmtHtml}

  <footer>출력일시: ${new Date().toLocaleString("ko-KR")}</footer>

  <!-- 모바일/데스크톱 공통 오버레이 — 인쇄 시에는 @media print 로 숨김 -->
  <div class="print-ovl" id="print-ovl">
    <button type="button" class="po-back" onclick="window.close();return false;">← 돌아가기</button>
    <span class="po-ttl">📋 면담일지 미리보기</span>
    <button type="button" class="po-print" onclick="window.print();return false;">🖨️ 인쇄하기</button>
  </div>
</body></html>`;

    const win = window.open("", "_blank", "width=900,height=1200");
    if (!win) { toast("팝업이 차단되었습니다. 팝업 허용 후 다시 시도하세요.", "error"); return; }
    win.document.write(html);
    win.document.close();
  }

  // ========== 출력 서브뷰 ==========
  function renderPrintPanelHtml() {
    const sel = state.students.find((x) => x.empNo === state.selectedEmpNo);
    const curBranch = state.filter.branch || sel?.branch || "";
    const curCenter = state.filter.center || sel?.center || "";
    const mode = state.printMode || "personal";
    return `
      <div class="detail-card print-card">
        <div class="print-controls no-print">
          <span class="pc-title">🖨️ 면담일지 출력</span>
          <label class="pc-field">출력 단위
            <select id="print-mode-sel" class="side-input">
              <option value="personal" ${mode === "personal" ? "selected" : ""}>개인별</option>
              <option value="branch"   ${mode === "branch"   ? "selected" : ""}>지점별</option>
              <option value="vc"       ${mode === "vc"       ? "selected" : ""}>비전센터별</option>
            </select>
          </label>
          <span class="pc-scope" id="print-scope-hint">
            ${mode === "personal" ? escapeHtml(sel?.name || "미선택")
              : mode === "branch" ? `지점: <strong>${escapeHtml(curBranch || "미지정")}</strong>`
              : `비전센터: <strong>${escapeHtml(curCenter || "미지정")}</strong>`}
          </span>
          <label class="pc-field">이력 건수
            <select id="print-cnt-sel" class="side-input">
              <option value="all">전체 이력</option>
              <option value="5">최근 5건</option>
              <option value="3">최근 3건</option>
              <option value="1">최근 1건</option>
            </select>
          </label>
          <button class="btn-primary" id="btn-do-print">🖨️ 프린터 출력</button>
          <button class="btn-outline" id="btn-save-png">🖼️ PNG 저장</button>
        </div>
        <div id="print-area" class="print-area"></div>
      </div>
    `;
  }

  function bindPrintControls() {
    const mSel = $("#print-mode-sel");
    if (mSel) mSel.addEventListener("change", (e) => {
      state.printMode = e.target.value;
      renderPrintView();
    });
    const cSel = $("#print-cnt-sel");
    if (cSel) cSel.addEventListener("change", renderPrintView);
    const p = $("#btn-do-print");
    if (p) p.addEventListener("click", () => window.print());
    const png = $("#btn-save-png");
    if (png) png.addEventListener("click", savePrintPNG);
  }

  // 현재 모드에 따른 출력 대상 학생 목록 + 스코프 이름 반환
  function getPrintScope() {
    const mode = state.printMode;
    const sel = state.students.find((x) => x.empNo === state.selectedEmpNo);
    if (mode === "personal") {
      return { mode, students: sel ? [sel] : [], scopeName: sel?.name || "", vc: sel?.center || "", branch: sel?.branch || "" };
    }
    if (mode === "branch") {
      const branch = state.filter.branch || sel?.branch || "";
      const students = state.students.filter((x) => x.branch === branch);
      const vc = students[0]?.center || sel?.center || "";
      return { mode, students, scopeName: branch, vc, branch };
    }
    if (mode === "vc") {
      const center = state.filter.center || sel?.center || "";
      const students = state.students.filter((x) => x.center === center);
      return { mode, students, scopeName: center, vc: center, branch: "" };
    }
    return { mode, students: [], scopeName: "" };
  }

  async function renderPrintView() {
    const area = $("#print-area");
    if (!area) return;

    // 제어바 scope 힌트 갱신
    const hint = $("#print-scope-hint");
    const scope = getPrintScope();
    if (hint) {
      if (scope.mode === "personal") hint.innerHTML = escapeHtml(scope.scopeName || "미선택");
      else if (scope.mode === "branch") hint.innerHTML = `지점: <strong>${escapeHtml(scope.scopeName || "미지정")}</strong> · ${scope.students.length}명`;
      else hint.innerHTML = `비전센터: <strong>${escapeHtml(scope.scopeName || "미지정")}</strong> · ${scope.students.length}명`;
    }

    if (!scope.students.length) {
      area.innerHTML = `<div class="empty-state"><div class="empty-ico">📄</div>출력 대상 교육생이 없습니다. 좌측 필터를 확인하세요.</div>`;
      return;
    }

    // 다건 모드면 consultations fetch 필요
    area.innerHTML = `<div class="empty-state">면담 기록 불러오는 중... (${scope.students.length}명)</div>`;
    await ensureConsultationsFetched(scope.students);

    // 빈 교육생(이력 없음) 제거
    const withItvs = scope.students
      .map((s) => ({ s, itvs: getStudentConsultations(s.empNo) }))
      .filter((x) => x.itvs.length);

    if (!withItvs.length) {
      area.innerHTML = `<div class="empty-state"><div class="empty-ico">📄</div>출력할 면담 기록이 없습니다.</div>`;
      return;
    }

    // cnt 적용
    const cntSel = $("#print-cnt-sel")?.value || "all";
    const maxCnt = cntSel === "all" ? Infinity : parseInt(cntSel, 10);
    withItvs.forEach((x) => {
      const sorted = x.itvs.slice().sort((a, b) => (a.date || "").localeCompare(b.date || ""));
      x.itvs = cntSel === "all" ? sorted : sorted.slice(-maxCnt);
    });

    // 그룹화
    //   personal → 1그룹 (학생 하나)
    //   branch   → 1그룹 (지점 전체)
    //   vc       → 지점별 여러 그룹
    const groups = [];
    if (scope.mode === "vc") {
      const byBranch = {};
      withItvs.forEach((x) => {
        const k = x.s.branch || "(지점 미지정)";
        if (!byBranch[k]) byBranch[k] = [];
        byBranch[k].push(x);
      });
      Object.keys(byBranch).sort().forEach((k) => {
        groups.push({ title: `${scope.vc} · ${k}`, branch: k, vc: scope.vc, rows: byBranch[k] });
      });
    } else if (scope.mode === "branch") {
      groups.push({ title: scope.scopeName, branch: scope.branch, vc: scope.vc, rows: withItvs });
    } else {
      groups.push({ title: withItvs[0].s.name, branch: withItvs[0].s.branch, vc: withItvs[0].s.center, rows: withItvs });
    }

    area.innerHTML = groups.map((g, gi) => buildGroupPagesHtml(g, gi === 0)).join("");
  }

  // 다건 모드에서 아직 fetch 안 한 학생만 Firestore 에서 가져와 캐시
  async function ensureConsultationsFetched(students) {
    const todo = [];
    for (const s of students) {
      if (s.empNo === state.selectedEmpNo) {
        // 현재 선택 학생은 이미 subscribe 중
        state.printConsultCache[s.empNo] = state.consultations.slice();
      } else if (!state.printConsultCache[s.empNo]) {
        todo.push(s.empNo);
      }
    }
    // 동시 8개씩 fetch
    const chunk = 8;
    for (let i = 0; i < todo.length; i += chunk) {
      const part = todo.slice(i, i + chunk);
      const results = await Promise.allSettled(part.map((emp) => window.DataAPI.getConsultationsOnce(emp)));
      results.forEach((r, j) => {
        state.printConsultCache[part[j]] = r.status === "fulfilled" ? r.value : [];
      });
    }
  }

  function getStudentConsultations(empNo) {
    if (empNo === state.selectedEmpNo) return state.consultations.slice();
    return (state.printConsultCache[empNo] || []).slice();
  }

  // 그룹(지점 또는 개인) 단위로 페이지 HTML 생성
  // 모바일 화면 전용: 출력 미리보기용 카드 렌더 (print 시에는 @media print 로 숨김)
  function renderPrintMobileView(studentBlocks, fn, nl, validClients, showSel) {
    return `
      <div class="print-mobile-view">
        ${studentBlocks.map((blk) => `
          <div class="pm-stu-card">
            <div class="pm-stu-hd">
              <span class="pm-stu-nm">${escapeHtml(blk.s.name || "")}</span>
              <span class="pm-stu-meta">평균실적 ${blk.stuIns} · 마스터목표 ${blk.stuTgt}</span>
            </div>
            ${blk.itvs.map((e) => {
              const clients = validClients(e.clients);
              return `
                <div class="pm-itv-card">
                  <div class="pm-itv-hd">
                    <span class="pm-itv-seq">${escapeHtml(e.seq || "-")}차</span>
                    <span class="pm-itv-dt">${escapeHtml(e.date || "")}</span>
                    <span class="pm-itv-pct">진도 ${fn(e.pct)}%</span>
                  </div>
                  <div class="pm-itv-stat">
                    <span><em>가입설계</em> ${fn(e.plan)}</span>
                    <span><em>행복보장</em> ${fn(e.hap)}</span>
                    <span><em>예상실적</em> ${fn(e.exp)}</span>
                  </div>
                  ${clients.length ? `
                    <div class="pm-clients">
                      ${clients.map((c, ci) => `
                        <div class="pm-cli-card">
                          <div class="pm-cli-num">${ci + 1} <small>번 고객</small></div>
                          <dl class="pm-cli-dl">
                            <dt>성명</dt><dd class="pm-strong">${escapeHtml(c.name || "")}</dd>
                            <dt>고객유형</dt><dd>${showSel(c.types) || "-"}</dd>
                            <dt>상담단계</dt><dd>${showSel(c.consult) || "-"}</dd>
                            <dt>활용자료</dt><dd>${showSel(c.material) || "-"}</dd>
                            <dt>제안금액</dt><dd>${showSel(c.amount) || "-"}${c.amountDirect ? ` <span class="pm-amt-d">+${escapeHtml(c.amountDirect)}만</span>` : ""}</dd>
                            <dt>보종</dt><dd>${showSel(c.bj) || "-"}</dd>
                            <dt>면담내용</dt><dd class="pm-memo">${nl(c.memo || "-")}</dd>
                          </dl>
                        </div>
                      `).join("")}
                    </div>
                  ` : ""}
                  ${e.coach ? `<div class="pm-coach"><strong>📌 코칭포인트</strong><p>${nl(e.coach)}</p></div>` : ""}
                </div>
              `;
            }).join("")}
          </div>
        `).join("")}
      </div>
    `;
  }

  function buildGroupPagesHtml(group, isFirstGroup) {
    // flatten 모든 학생의 행 + 학생 메타 유지
    const fn = (v) => {
      const n = parseInt(v, 10);
      return Number.isFinite(n) && n > 0 ? n.toLocaleString() : "-";
    };
    const nl = (x) => escapeHtml(x || "").replace(/\n/g, "<br>");
    const showSel = (arr) => escapeHtml((arr || []).join(", "));
    const validClients = (clients) => (clients || []).filter((c) => c && (c.name || (c.types && c.types.length) || (c.consult && c.consult.length) || (c.material && c.material.length) || (c.amount && c.amount.length) || (c.bj && c.bj.length) || (c.memo && c.memo.trim())));

    // 학생 단위 flat + 학생 블록 경계
    // 한 학생이 끝나면 페이지 분할 기준점 삽입 가능
    const studentBlocks = group.rows.map(({ s, itvs }) => {
      const firstItv = itvs.find((e) => e.seq === "1") || itvs[0];
      const stuIns = fn(Number(firstItv?.ins||0) * 1000); // 천원 → 원 표시
      const stuTgt = fn(Number(firstItv?.tgt||0) * 1000);
      // 코칭포인트: 모든 면담의 seq·date·coach 를 합쳐 한 줄로
      const coachText = itvs
        .filter((e) => (e.coach || "").trim())
        .map((e) => `${e.seq ? `[${e.seq}차] ` : ""}${e.coach.trim()}`)
        .join("\n");
      const rows = [];
      itvs.forEach((e) => {
        const clients = validClients(e.clients);
        const count = Math.max(clients.length, 1);
        for (let ci = 0; ci < count; ci++) {
          const c = clients[ci] || null;
          rows.push({
            itvFirst: ci === 0, itvRows: count,
            seq: e.seq || "", date: e.date || "", pct: e.pct || "",
            plan: e.plan || "", hap: e.hap || "", exp: fn(e.exp),
            cName: c ? (c.name || "") : "",
            cTypes: c ? showSel(c.types) : "",
            cConsult: c ? showSel(c.consult) : "",
            cMaterial: c ? showSel(c.material) : "",
            cAmount: c ? (showSel(c.amount) || (c.amountDirect ? escapeHtml(c.amountDirect) + "만원" : "")) : "",
            cBj: c ? showSel(c.bj) : "",
            cMemo: c ? (c.memo || "") : ""
          });
        }
      });
      return { s, itvs, stuIns, stuTgt, rows, coachText };
    });

    // 그룹 합계 계산
    const uIns = new Map(), uTgt = new Map();
    studentBlocks.forEach(({ s, itvs }) => {
      const f = itvs.find((e) => e.seq === "1") || itvs[0];
      if (f) { uIns.set(s.empNo, parseFloat(f.ins) || 0); uTgt.set(s.empNo, parseFloat(f.tgt) || 0); }
    });
    const sumIns = [...uIns.values()].reduce((a, v) => a + v, 0);
    const sumTgt = [...uTgt.values()].reduce((a, v) => a + v, 0);
    const sumExp = studentBlocks.flatMap((b) => b.itvs).reduce((a, e) => a + (parseFloat(e.exp) || 0), 0);
    const dRate = sumTgt ? ((sumIns / sumTgt) * 100).toFixed(1) + "%" : "-";
    const lRate = sumTgt ? ((sumExp / sumTgt) * 100).toFixed(1) + "%" : "-";
    const remain = sumTgt - sumExp;

    // 페이지 분할: 학생 블록을 통째로 유지하면서 13행 근처에서 자름
    const MAX = 12;
    const pages = [];
    let curPage = [];
    let curRows = 0;
    studentBlocks.forEach((blk) => {
      let start = 0;
      while (start < blk.rows.length) {
        const remaining = MAX - curRows;
        if (remaining <= 0) {
          pages.push(curPage);
          curPage = [];
          curRows = 0;
          continue;
        }
        const take = Math.min(remaining, blk.rows.length - start);
        curPage.push({ blk, startRow: start, endRow: start + take, isFirst: start === 0 });
        curRows += take;
        start += take;
      }
    });
    if (curPage.length) pages.push(curPage);
    if (!pages.length) pages.push([]);

    // 시상계산기 / 의견 (그룹 마지막 페이지에만, 개인별 모드면 해당 학생 기준)
    const allItvs = studentBlocks.flatMap((b) => b.itvs);
    const calcItv = allItvs.slice().reverse().find((e) => e.calcTgt || e.calcAvg);
    const cmtItv = allItvs.slice().reverse().find((e) => e.calcComment);
    const _printStep = state.filter.step || "1";
    const _printStudent = studentBlocks[0]?.s;
    let awardHtml = "";
    if (state.printMode === "personal") {
      if (_printStep === "2" && _printStudent?.pgCurrent) {
        awardHtml = buildPrintAwardHtmlStep2(_printStudent, allItvs);
      } else if (calcItv) {
        awardHtml = buildPrintAwardHtml(calcItv, allItvs);
      }
    }
    const cmtHtml = cmtItv?.calcComment ? `
      <div class="print-comment">
        <strong>✍️ 면담자 의견</strong>
        <p>${nl(cmtItv.calcComment)}</p>
      </div>` : "";

    // 날짜/차수 종합
    const dateAll = allItvs.map((e) => e.date).filter(Boolean).sort();
    const dateStr = dateAll.length > 1 ? `${dateAll[0]} ~ ${dateAll[dateAll.length - 1]}` : (dateAll[0] || "");
    const seqs = [...new Set(allItvs.map((e) => e.seq).filter(Boolean))];
    const seqTxt = seqs.length === 1 ? seqs[0] : (seqs.join(", ") || "");

    const managerName = allItvs.find((e) => e.manager)?.manager || "";
    const apvHtml = `
      <div class="apv-wrap">
        <div class="apv-box"><div class="apv-title">담당</div><div class="apv-sign apv-sign-name">${escapeHtml(managerName)}</div></div>
        <div class="apv-box"><div class="apv-title">파트장</div><div class="apv-sign"></div></div>
        <div class="apv-box"><div class="apv-title">지역단장</div><div class="apv-sign"></div></div>
      </div>`;

    const tableHdr = `
      <colgroup>
        <col style="width:72px"><col style="width:52px"><col style="width:52px">
        <col style="width:34px"><col style="width:32px"><col style="width:34px">
        <col style="width:58px"><col style="width:46px"><col style="width:64px">
        <col style="width:96px"><col style="width:62px"><col style="width:46px">
        <col style="width:225px"><col style="width:48px">
      </colgroup>
      <thead>
        <tr>
          <th rowspan="3">교육생<br>성명</th>
          <th colspan="2">고객마스터 인보험(천원)</th>
          <th colspan="3">설계진도</th>
          <th colspan="8">주간활동 점검내용</th>
        </tr>
        <tr>
          <th rowspan="2">월평균<br>(6개월)</th>
          <th rowspan="2">당월<br>목표</th>
          <th rowspan="2">진도</th>
          <th rowspan="2">가입<br>설계</th>
          <th rowspan="2">행복<br>보장</th>
          <th colspan="6">상담고객</th>
          <th rowspan="2">면담내용</th>
          <th rowspan="2">예상<br>실적<br>(천원)</th>
        </tr>
        <tr>
          <th>성명</th><th>구분</th><th>상담단계</th><th>활용자료</th><th>제안금액</th><th>보종</th>
        </tr>
      </thead>`;

    return pages.map((pgSlices, pi) => {
      const isLast = pi === pages.length - 1;
      const isVeryFirst = isFirstGroup && pi === 0;

      // 학생별 연속 행 그룹 정리 (같은 페이지 안에서 같은 학생 블록)
      let tbody = "";
      // pgSlices: [{blk, startRow, endRow, isFirst}]
      // 같은 학생 연속 슬라이스를 합쳐 stuName rowspan 적용
      // 한 페이지 내 같은 학생이 여러 조각으로 등장할 수 있으므로, 각 조각별 학생 헤더는 첫 조각에만.
      const grouped = [];
      pgSlices.forEach((slice) => {
        const last = grouped[grouped.length - 1];
        if (last && last.blk === slice.blk) {
          last.endRow = slice.endRow;
          last.count += (slice.endRow - slice.startRow);
        } else {
          grouped.push({ blk: slice.blk, startRow: slice.startRow, endRow: slice.endRow, count: slice.endRow - slice.startRow, isFirst: slice.isFirst });
        }
      });

      grouped.forEach(({ blk, startRow, endRow, count, isFirst }) => {
        const rows = blk.rows.slice(startRow, endRow);
        tbody += rows.map((r, ri) => {
          const isStuFirst = ri === 0;
          const stuTds = isStuFirst ? `
            <td rowspan="${count}" class="stu-first">${escapeHtml(blk.s.name || "")}</td>
            <td rowspan="${count}" class="stu-first-c">${blk.stuIns}</td>
            <td rowspan="${count}" class="stu-first-c">${blk.stuTgt}</td>` : "";
          const itvTds = r.itvFirst ? `
            <td rowspan="${r.itvRows}" class="itv-c">${escapeHtml(r.pct)}${r.seq ? `<div class="seq-sub">${escapeHtml(r.seq)}차</div>` : ""}</td>
            <td rowspan="${r.itvRows}" class="itv-c">${escapeHtml(r.plan)}</td>
            <td rowspan="${r.itvRows}" class="itv-c">${escapeHtml(r.hap)}</td>` : "";
          const expTd = r.itvFirst ? `<td rowspan="${r.itvRows}" class="exp-c">${r.exp}</td>` : "";
          return `<tr>
            ${stuTds}${itvTds}
            <td class="c-name">${escapeHtml(r.cName)}</td>
            <td class="c-c">${r.cTypes}</td>
            <td class="c-c">${r.cConsult}</td>
            <td class="c-c">${r.cMaterial}</td>
            <td class="c-amt">${r.cAmount}</td>
            <td class="c-bj">${r.cBj}</td>
            <td class="c-memo">${nl(r.cMemo)}</td>
            ${expTd}
          </tr>`;
        }).join("");
        // 학생 블록의 마지막 슬라이스라면 코칭포인트 행 추가 (전폭 14열)
        const isLastSliceOfStudent = endRow === blk.rows.length;
        if (isLastSliceOfStudent && blk.coachText) {
          tbody += `<tr class="coach-row"><td colspan="14" class="coach-cell">📌 <strong>코칭포인트</strong> — ${nl(blk.coachText)}</td></tr>`;
        }
      });

      const summaryHtml = isLast ? `<tr class="sr">
        <td class="c">합계</td>
        <td class="c strong">${fn(sumIns * 1000)}</td>
        <td class="c strong">${fn(sumTgt * 1000)}</td>
        <td colspan="9" class="sr-text">
          달성률 D(월평균/목표): <strong>${dRate}</strong> &nbsp;
          L(예상/목표): <strong>${lRate}</strong> &nbsp;&nbsp;
          잔여목표: <strong>${fn(remain * 1000)}원</strong>
        </td>
        <td></td>
        <td class="c strong">${fn(sumExp * 1000)}</td>
      </tr>` : "";

      const brkClass = (pi === 0 && !isFirstGroup) ? " pg-break-before" : "";
      return `
        ${pi > 0 ? `<hr class="page-dashed-sep">` : ""}
        <div class="print-page${brkClass}">
          <div class="doc-title-row">
            <div class="doc-title">고객컨설팅 MASTER과정 지점별 면담일지 ( ${seqTxt || "&nbsp;&nbsp;"} 차 / 활동점검 )</div>
            ${isVeryFirst ? apvHtml : ""}
          </div>
          <div class="doc-hdr">
            <div><strong>비젼센터:</strong> ${escapeHtml(group.vc || "")} &nbsp;&nbsp; <strong>지점명:</strong> ${escapeHtml(group.branch || "")}</div>
            <div><strong>면담일시:</strong> ${escapeHtml(dateStr)} &nbsp;<span class="pg-cnt">(${pi + 1}/${pages.length})</span></div>
          </div>
          <table class="dt">
            ${tableHdr}
            <tbody>
              ${tbody}
              ${summaryHtml}
            </tbody>
          </table>
          ${isVeryFirst ? renderPrintMobileView(studentBlocks, fn, nl, validClients, showSel) : ""}
          ${isLast ? cmtHtml : ""}
          ${isLast ? awardHtml : ""}
        </div>
      `;
    }).join("");
  }

  // ========== 시상안 렌더 공통 헬퍼 ==========
  // 입력 우선순위: 면담이력 calc값 → student.base/honors/insAvg 로 폴백
  function computeAwardBreakdown(student, calcItv, lastInsItv) {
    const parseR = (v) => parseFloat(String(v || "").replace(/,/g, "")) || 0;
    let avgRaw = calcItv?.calcAvg ? parseR(calcItv.calcAvg) : (Number(student?.base) || 0);
    let baseTgtRaw = calcItv?.calcBaseTgt ? parseR(calcItv.calcBaseTgt) : (Number(student?.honors) || 0);
    let tgtRaw = 0;
    if (calcItv?.calcTgt) {
      let v = parseR(calcItv.calcTgt);
      if (v > 0 && v < 1000) v = v * 1000; // 천원 오저장 보정
      tgtRaw = v;
    }
    if (!tgtRaw && Number(student?.insAvg) > 0) tgtRaw = (Number(student.insAvg) + 200) * 1000;
    if (!tgtRaw && Number(student?.honors) > 0) tgtRaw = Number(student.honors);
    // 인보험 평균(원) 우선순위: insAvg(천원→원) → 최근 면담 ins(천원→원) → student.base(원)
    let insRaw = 0;
    if (Number(student?.insAvg) > 0) insRaw = Number(student.insAvg) * 1000;
    else if (lastInsItv?.ins) insRaw = (parseFloat(lastInsItv.ins) || 0) * 1000;
    else insRaw = Number(student?.base) || 0;

    const tgt = tgtRaw / 10000;
    const baseTgt = baseTgtRaw / 10000;
    const incr = Math.max(0, tgt - baseTgt);

    // ① 아너스
    let award1 = 0, award1Idx = -1, award1Grade = "해당없음";
    for (let i = 0; i < HONORS.length; i++) {
      if (tgt >= HONORS[i].critVal) { award1 = HONORS[i].prize; award1Idx = i; award1Grade = HONORS[i].grade; break; }
    }
    // ② 하이포인트
    const baseTgtMet = (baseTgt <= 0) || (tgt >= baseTgt);
    const mExtra = baseTgtMet ? Math.floor(incr * INCR_CFG.rate / 100 * 10) / 10 : 0;
    const mSub = baseTgtMet ? INCR_CFG.base + mExtra : 0;
    const mFinal = baseTgtMet ? Math.min(mSub, INCR_CFG.mcap) : 0;
    const award2M3 = baseTgtMet ? Math.min(mFinal * 3, INCR_CFG.qcap) : 0;
    // ③ 마스터
    const incrMaster = Math.max(0, tgtRaw - insRaw) / 10000;
    const award3 = calcMasterAward(incrMaster);
    const award3Tier = MASTER_AWARD.find((t) => incrMaster >= t.critVal);
    const total = award1 + award2M3 + award3 * 2;

    // 상위 등급 비교 (최대 3단계)
    const upper = [];
    const startIdx = award1Idx >= 0 ? award1Idx - 1 : HONORS.length - 1;
    for (let i = startIdx; i >= 0 && upper.length < 3; i--) {
      const h = HONORS[i];
      const ui = Math.max(0, h.critVal - baseTgt);
      const uMet = h.critVal >= baseTgt;
      const ue = uMet ? Math.floor(ui * INCR_CFG.rate / 100 * 10) / 10 : 0;
      const uSub = uMet ? (INCR_CFG.base + ue) : 0;
      const uFinal = uMet ? Math.min(uSub, INCR_CFG.mcap) : 0;
      const uM3 = uMet ? Math.min(uFinal * 3, INCR_CFG.qcap) : 0;
      const uAw3 = calcMasterAward(Math.max(0, h.critVal * 10000 - insRaw) / 10000);
      upper.push({ grade: h.grade, criteria: h.criteria, prize: h.prize, needIncr: ui, monthly: uFinal, mCapped: uSub > INCR_CFG.mcap, award2: uM3, qCapped: uFinal * 3 > INCR_CFG.qcap, award3m: uAw3, total3m: h.prize + uM3 + uAw3 * 2 });
    }

    return {
      avgRaw, baseTgtRaw, tgtRaw, insRaw, tgt, baseTgt, incr,
      award1, award1Idx, award1Grade, baseTgtMet, mExtra, monthlyFinal: mFinal, award2M3, incrMaster, award3, award3Tier, total, upper
    };
  }

  function fmtW(mw) { return Math.round(mw * 10000).toLocaleString() + "원"; }
  function fmtWon(w) { return Math.round(w).toLocaleString() + "원"; }

  // 한 교육생의 시상안 A4 포트레이트 페이지 HTML
  // opts.positiveWord 가 있으면: 제목 → "이름사장님의 희망답안지", 하단 메시지 추가
  function buildAwardSheetPageHtml(student, calcItv, lastInsItv, opts = {}) {
    const b = computeAwardBreakdown(student, calcItv, lastInsItv);
    if (!b.avgRaw && !b.tgtRaw) return "";

    const upperHtml = b.upper.length ? `
      <div class="up-title">📈 상위 등급 달성 시 시상금 비교</div>
      <table class="up-table">
        <thead><tr>
          <th>등급</th><th>기준</th><th>초과달성</th>
          <th>①아너스</th><th>월 하이포</th><th>②3개월</th><th>③마스터×2</th><th>총합계</th>
        </tr></thead>
        <tbody>${b.upper.map((g, i) => `<tr class="${i === 0 ? "up-next" : ""}">
          <td><strong>${escapeHtml(g.grade.split("(")[0].trim())}</strong></td>
          <td>${escapeHtml(g.criteria)}</td>
          <td class="rd">${fmtW(g.needIncr)}</td>
          <td>${fmtW(g.prize)}</td>
          <td>${fmtW(g.monthly)}${g.mCapped ? "<br><span class=\"warn\">⚠최대</span>" : ""}</td>
          <td>${fmtW(g.award2)}${g.qCapped ? "<br><span class=\"warn\">⚠최대</span>" : ""}</td>
          <td class="bl">${g.award3m ? fmtW(g.award3m * 2) : "—"}</td>
          <td class="grn"><strong>${fmtW(g.total3m)}</strong></td>
        </tr>`).join("")}</tbody>
      </table>` : "";

    const award1Html = b.award1Idx >= 0 ? `
      <div class="hl-row">
        <span class="hl-icon">🏆</span>
        <div class="hl-info">
          <div class="hl-grade">아너스클럽 · ${escapeHtml(b.award1Grade)}</div>
          <div class="hl-crit">${escapeHtml(HONORS[b.award1Idx].criteria)} 달성</div>
        </div>
        <div class="hl-amt">${fmtW(b.award1)}</div>
      </div>` : `<div class="hl-none">아너스클럽 — 해당없음</div>`;

    const award2Html = b.baseTgtMet ? `
      <div class="hl-row">
        <span class="hl-icon">💧</span>
        <div class="hl-info">
          <div class="hl-grade">하이포인트 지급 (개인순증시상)</div>
          <div class="hl-crit">기본 5만원 + 순증 ${fmtW(b.incr)} × 50% = 월 <strong>${fmtWon(b.monthlyFinal * 10000)}</strong> × 3개월 = <strong>${fmtW(b.award2M3)}</strong></div>
        </div>
        <div class="hl-amt">${fmtW(b.award2M3)}</div>
      </div>` : `<div class="hl-row warn">
        <span class="hl-icon">⚠️</span>
        <div class="hl-info">
          <div class="hl-grade">하이포인트 — 미해당</div>
          <div class="hl-crit">희망목표(${fmtWon(b.tgtRaw)})가 기본순증목표(${fmtWon(b.baseTgtRaw)})에 미달</div>
        </div>
        <div class="hl-amt rd">0원</div>
      </div>`;

    const award3Html = b.award3 > 0 ? `
      <div class="hl-row blue">
        <span class="hl-icon">🎯</span>
        <div class="hl-info">
          <div class="hl-grade">마스터과정 · ${escapeHtml(b.award3Tier.criteria)}</div>
          <div class="hl-crit">순증 ${fmtWon(b.incrMaster * 10000)} = 희망 ${fmtWon(b.tgtRaw)} − 인보험평균 ${fmtWon(b.insRaw)} &nbsp;<span class="lbl">${escapeHtml(b.award3Tier.label)}</span></div>
          <div class="hl-sub">매월 ${fmtW(b.award3)} × 2개월</div>
        </div>
        <div class="hl-amt bl">${fmtW(b.award3 * 2)}</div>
      </div>` : `<div class="hl-none">마스터과정 — 순증 5만원 미만 (${fmtWon(b.incrMaster * 10000)})</div>`;

    const region = student.region || "";
    const vc = student.center || "";
    const branch = student.branch || "";
    const sName = student.name || "";

    // ④ 기수·스텝 개인 순증 시상
    const _p4Region = region;
    const _p4Cohort = (student.cohort || state.filter.cohort || "").replace(/기$/, "");
    const _p4Step   = state.filter.step || "1";
    const _p4Pa     = getProgressAwardConfig(_p4Region, _p4Cohort, _p4Step);
    const _p4CoLabel = _p4Cohort ? `${_p4Cohort}기` : "";
    const _p4StLabel = `Step ${_p4Step}`;
    const _p4NetWon  = Math.max(0, b.tgtRaw - b.insRaw);
    const _p4TierHit = _p4Pa.tiers.find(t => _p4NetWon >= t.min);
    const _p4Badge   = `${_p4CoLabel ? _p4CoLabel + " " : ""}${_p4StLabel}`;
    let award4Html = "";
    if (_p4Pa.tiers.length > 0) {
      if (_p4TierHit) {
        const _p4Prize = _p4TierHit.type === "pct"
          ? `${_p4TierHit.payVal}% 지급 (${fmtWon(Math.round(_p4NetWon * _p4TierHit.val))})`
          : `${_p4TierHit.payVal}만원`;
        award4Html = `
          <div class="hl-row green4">
            <span class="hl-icon">🎁</span>
            <div class="hl-info">
              <div class="hl-grade">${escapeHtml(_p4Badge)} 개인 순증 시상</div>
              <div class="hl-crit">마스터목표 순증 ${fmtWon(_p4NetWon)} → <strong>${escapeHtml(_p4Prize)}</strong> 획득 예정</div>
            </div>
            <div class="hl-amt grn4">${escapeHtml(_p4Prize)}</div>
          </div>`;
      } else {
        award4Html = `<div class="hl-none">${escapeHtml(_p4Badge)} 개인 순증 시상 — 해당없음 (마스터목표 순증 ${fmtWon(_p4NetWon)})</div>`;
      }
      if (_p4Step === "2") {
        const _prevPa4 = getProgressAwardConfig(_p4Region, _p4Cohort, "1");
        if (_prevPa4.tiers.length > 0) {
          const _s1Net = Number(student.pgCurrent || 0) - Number(student.base || 0);
          const _prevHit4 = _prevPa4.tiers.find(t => _s1Net >= t.min);
          if (_prevHit4) {
            const _prevPrize4 = _prevHit4.type === "pct"
              ? `${Math.round(_s1Net * _prevHit4.val / 10000)}만원`
              : `${_prevHit4.payVal}만원`;
            award4Html += `<div class="hl-prev-hit">🎉 지난달 스텝1에서는 시상금 <strong>${escapeHtml(_prevPrize4)}</strong>을 획득 하셨습니다.</div>`;
          } else {
            award4Html += `<div class="hl-prev-miss">😅 지난달엔 아쉽지만 개인시상을 획득하지 못하셨습니다. 이달에는 더 화이팅!!</div>`;
          }
        }
      }
    }
    const branchShort = branch.replace(/지점$/, "");
    const hdrTitle = opts.positiveWord
      ? `🏆 ${escapeHtml(sName)}사장님의 희망답안지`
      : `🏆 ${escapeHtml(vc || region)} 시상 예상답안지`;
    const footerMsgHtml = opts.positiveWord
      ? `<div class="print-footer-msg"><span class="print-footer-name">${escapeHtml(sName)}</span>님은 <span class="print-footer-branch">${escapeHtml(branchShort)}</span>지점의 <span class="print-footer-word">'${escapeHtml(opts.positiveWord)}'</span></div>`
      : "";

    return `
      <div class="pg">
        <div class="hdr">
          <div class="hdr-title">${hdrTitle}</div>
          <div class="hdr-date">${new Date().toLocaleDateString("ko-KR")} 기준</div>
        </div>
        <div class="info-row1">
          <div class="info-card key"><div class="info-lbl">지역단</div><div class="info-val">${escapeHtml(region.replace(/지역단$|사업부$/, ""))}</div></div>
          <div class="info-card key"><div class="info-lbl">지점</div><div class="info-val">${escapeHtml(branch.replace(/지점$/, ""))}</div></div>
          <div class="info-card key"><div class="info-lbl">성명</div><div class="info-val">${escapeHtml(student.name || "")}</div></div>
          <div class="info-card"><div class="info-lbl">사번</div><div class="info-val">${escapeHtml(student.empNo || "")}</div></div>
        </div>
        <div class="info-row2">
          <div class="info-card stat">
            <div class="stat-lbl">📊 평균실적</div>
            <div class="stat-row"><span class="stat-key">매출 아너스:</span><span class="stat-val">${fmtWon(b.avgRaw)}</span></div>
            <div class="stat-row"><span class="stat-key bl">마스터목표:</span><span class="stat-val bl">${fmtWon(b.tgtRaw)}</span></div>
          </div>
          <div class="info-card stat">
            <div class="stat-lbl">🎯 기본순증목표</div>
            <div class="stat-row"><span class="stat-key">아너스기본목표:</span><span class="stat-val">${b.baseTgtRaw ? fmtWon(b.baseTgtRaw) : "—"}</span></div>
            <div class="stat-row"><span class="stat-key bl">고객마스터 희망:</span><span class="stat-val bl">${fmtWon(b.tgtRaw)}</span></div>
          </div>
        </div>
        <div class="sec-title bl">🎯 고객컨설팅마스터 과정 개인시상</div>
        ${award3Html}
        ${award4Html ? `<div class="sec-title grn4">🎁 ${escapeHtml(_p4Badge)} 개인 순증 시상</div>${award4Html}` : ""}
        <div class="sec-title">📌 아너스 희망목표금액 기준 시상 (${fmtWon(b.tgtRaw)})</div>
        ${award1Html}${award2Html}
        <div class="total-bar">
          <div>
            <div class="total-lbl">🏆 최종 예상 시상금 합계 (①+②+③)</div>
            <div class="total-sub">아너스 ${fmtW(b.award1)} + 하이포 ${fmtW(b.award2M3)} + 마스터 ${fmtW(b.award3 * 2)}</div>
          </div>
          <div class="total-val">${fmtW(b.total)}</div>
        </div>
        ${upperHtml}
        <div class="note">
          ※ 아너스클럽: 3개월 연속달성 기준 · 하이포인트: 월 기본 5만원 + 순증×50% (월 최대 20만, 분기 최대 50만)<br>
          ※ 마스터과정: 순증 5/10/20만원 고정지급, 30/50만원↑ 순증 120%/150%
        </div>
        ${footerMsgHtml}
      </div>
    `;
  }

  // 시상안 출력용 CSS (스탠드얼론 인쇄창에 삽입)
  const AWARD_PRINT_CSS = `
    *{box-sizing:border-box;margin:0;padding:0;}
    body{font-family:'Noto Sans KR','Malgun Gothic',sans-serif;background:#fff;color:#1A1A1A;font-size:13px;line-height:1.45;}
    .pg{padding:8mm 10mm;page-break-after:always;}
    .pg:last-child{page-break-after:auto;}
    .hdr{background:linear-gradient(135deg,#1A2744,#2C3F6E);border-radius:8px;padding:10px 14px;margin-bottom:9px;display:flex;align-items:center;justify-content:space-between;}
    .hdr-title{color:#fff;font-size:20px;font-weight:900;}
    .hdr-date{color:rgba(255,255,255,.7);font-size:13px;white-space:nowrap;}
    .info-row1{display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:5px;margin-bottom:5px;}
    .info-row2{display:grid;grid-template-columns:1fr 1fr;gap:5px;margin-bottom:9px;}
    .info-card{background:#F5F7FF;border-radius:6px;padding:7px 9px;text-align:center;border:1px solid #D5DCF5;}
    .info-lbl{font-size:12px;color:#5C6BC0;font-weight:700;margin-bottom:3px;}
    .info-val{font-size:14px;font-weight:900;color:#1A2744;}
    .info-card.key{background:#1A2744;border-color:#1A2744;}
    .info-card.key .info-lbl{color:rgba(255,255,255,.65);}
    .info-card.key .info-val{color:#fff;font-size:20px;white-space:nowrap;}
    .info-card.stat{text-align:left;padding:8px 11px;background:#F8F9FF;}
    .stat-lbl{font-size:13px;color:#5C6BC0;font-weight:800;margin-bottom:5px;padding-bottom:4px;border-bottom:1px solid #D5DCF5;}
    .stat-row{display:flex;justify-content:space-between;align-items:center;margin-bottom:3px;}
    .stat-key{font-size:12px;font-weight:700;color:#4A148C;white-space:nowrap;}
    .stat-key.bl{color:#1565C0;}
    .stat-val{font-size:17px;font-weight:900;color:#1A2744;white-space:nowrap;}
    .stat-val.bl{color:#1565C0;}
    .sec-title{font-size:13px;font-weight:800;color:#4A148C;margin:7px 0 4px;}
    .sec-title.bl{color:#1565C0;}
    .hl-row{display:flex;align-items:center;gap:8px;background:#E8F5E9;border-radius:7px;padding:7px 11px;margin-bottom:4px;border-left:4px solid #2E7D32;}
    .hl-row.warn{background:#FFEBEE;border-left-color:#C62828;}
    .hl-row.blue{background:#E3F2FD;border-left-color:#1565C0;}
    .hl-icon{font-size:18px;flex-shrink:0;}
    .hl-info{flex:1;min-width:0;}
    .hl-grade{font-size:14px;font-weight:700;color:#1B5E20;}
    .hl-row.warn .hl-grade{color:#C62828;}
    .hl-row.blue .hl-grade{color:#1565C0;}
    .hl-crit{font-size:12px;color:#388E3C;margin-top:2px;line-height:1.5;}
    .hl-row.warn .hl-crit{color:#E57373;}
    .hl-row.blue .hl-crit{color:#1976D2;}
    .hl-crit .lbl{background:#1565C0;color:#fff;font-size:11px;padding:1px 6px;border-radius:8px;font-weight:700;}
    .hl-sub{font-size:13px;font-weight:800;color:#0D47A1;margin-top:2px;}
    .hl-amt{font-size:20px;font-weight:900;color:#2E7D32;white-space:nowrap;}
    .hl-amt.rd{color:#C62828;}
    .hl-amt.bl{color:#1565C0;font-size:22px;}
    .hl-none{padding:6px 10px;color:#999;font-size:12px;background:#F5F5F5;border-radius:6px;margin-bottom:4px;}
    .total-bar{background:linear-gradient(135deg,#1A2744,#2C3F6E);border-radius:7px;padding:9px 15px;display:flex;align-items:center;justify-content:space-between;margin:7px 0 7px;}
    .total-lbl{color:rgba(255,255,255,.85);font-size:13px;font-weight:700;}
    .total-sub{color:rgba(255,255,255,.6);font-size:11px;margin-top:2px;}
    .total-val{color:#FFE082;font-size:26px;font-weight:900;}
    .up-title{font-size:12px;font-weight:800;color:#1565C0;margin:5px 0 3px;}
    .up-table{width:100%;border-collapse:collapse;font-size:11px;}
    .up-table th{background:#E3F2FD;padding:3px 4px;border:1px solid #BBDEFB;color:#1565C0;font-weight:700;font-size:11px;}
    .up-table td{padding:3px 4px;border:1px solid #E3E3E3;text-align:center;font-size:11px;}
    .up-table td.rd{color:#C62828;font-weight:700;}
    .up-table td.bl{color:#1565C0;}
    .up-table td.grn{color:#1B5E20;font-size:12px;}
    .up-table tr.up-next td{background:#FFF9C4;font-weight:700;}
    .warn{color:#C62828;font-size:10px;font-weight:800;}
    .note{font-size:10px;color:#888;margin-top:5px;line-height:1.4;}
    .print-footer-msg{margin-top:8px;text-align:center;font-size:14px;font-weight:700;color:#1A2744;padding:10px 14px;background:linear-gradient(135deg,#F0F4FF,#E8F0FE);border-radius:7px;border:1.5px solid #C5D0F0;}
    .print-footer-name{color:#7B1FA2;font-weight:900;}
    .print-footer-branch{color:#1565C0;font-weight:900;}
    .print-footer-word{color:#E65100;font-size:17px;font-weight:900;font-style:italic;}
    .sec-title.grn4{color:#1B5E20;}
    .hl-row.green4{background:#E8F5E9;border-left-color:#2E7D32;}
    .hl-row.green4 .hl-grade{color:#1B5E20;}
    .hl-row.green4 .hl-crit{color:#388E3C;}
    .hl-amt.grn4{color:#2E7D32;font-size:20px;}
    .hl-prev-hit{background:#E8F5E9;border-left:3px solid #4CAF50;border-radius:6px;padding:6px 10px;margin-bottom:4px;font-size:13px;color:#1B5E20;font-weight:700;}
    .hl-prev-miss{background:#FFF8F0;border-left:3px solid #f59e0b;border-radius:6px;padding:6px 10px;margin-bottom:4px;font-size:13px;color:#b45309;font-weight:700;}
    @media print{
      @page{size:A4 portrait;margin:5mm 7mm;}
      body{font-size:11px;}
      .pg{padding:4mm 6mm;}
      .hdr{padding:6px 11px;margin-bottom:6px;}
      .hdr-title{font-size:16px;}
      .info-card.key .info-val{font-size:16px;}
      .stat-val{font-size:14px;}
      .hl-amt{font-size:16px;}
      .hl-amt.bl{font-size:18px;}
      .total-val{font-size:22px;}
      .info-row1,.info-row2{margin-bottom:3px;}
      .hl-row{padding:5px 9px;margin-bottom:3px;}
      .up-table th,.up-table td{padding:2px 3px;}
      .sec-title{margin:4px 0 2px;}
      .total-bar{padding:6px 12px;margin:4px 0;}
      .note{margin-top:3px;}
      .print-footer-msg{margin-top:5px;padding:6px 10px;}
    }
  `;

  // (출력 서브탭 personal 모드용) 직전 calcItv 기반 시상 결과 박스
  function buildPrintAwardHtml(calcItv, allItvs) {
    if (!calcItv.calcTgt) return "";
    const fw = (mw) => Math.round(mw * 10000).toLocaleString() + "원";
    const fwo = (w) => Math.round(w).toLocaleString() + "원";
    const parseR = (v) => parseFloat(String(v || "").replace(/,/g, "")) || 0;
    const tgtRaw = parseR(calcItv.calcTgt);
    const baseTgtRaw = parseR(calcItv.calcBaseTgt);
    const tgt = tgtRaw / 10000;
    const baseTgt = baseTgtRaw / 10000;
    const incr = Math.max(0, tgt - baseTgt);

    // ① 아너스
    let aw1 = 0, aw1Grade = "해당없음", aw1Idx = -1;
    for (let i = 0; i < HONORS.length; i++) {
      if (tgt >= HONORS[i].critVal) { aw1 = HONORS[i].prize; aw1Grade = HONORS[i].grade; aw1Idx = i; break; }
    }
    const baseTgtMet = (baseTgt <= 0) || (tgt >= baseTgt);
    const mExtra = baseTgtMet ? Math.floor(incr * INCR_CFG.rate / 100 * 10) / 10 : 0;
    const mSub = baseTgtMet ? INCR_CFG.base + mExtra : 0;
    const mFinal = baseTgtMet ? Math.min(mSub, INCR_CFG.mcap) : 0;
    const aw2M3 = baseTgtMet ? Math.min(mFinal * 3, INCR_CFG.qcap) : 0;
    const lastIns = allItvs.slice().reverse().find((e) => e.ins);
    const insRaw = (parseFloat(lastIns?.ins || "0") || 0) * 1000;
    const incrMW3 = Math.max(0, tgtRaw - insRaw) / 10000;
    const aw3 = calcMasterAward(incrMW3);
    const aw3Tier = MASTER_AWARD.find((t) => incrMW3 >= t.critVal);
    const aw3Next = aw3Tier ? MASTER_AWARD[MASTER_AWARD.indexOf(aw3Tier) - 1] : (incrMW3 < MASTER_AWARD[MASTER_AWARD.length - 1].critVal ? MASTER_AWARD[MASTER_AWARD.length - 1] : null);
    const total = aw1 + aw2M3 + aw3 * 2;

    return `
      <div class="print-award">
        <div class="pa-hdr">
          <div class="pa-t">📊 시상 계산기 결과 (${escapeHtml(calcItv.seq || "직전")}차 기준)</div>
          <div class="pa-sub">
            <span class="pa-sub-k">2분기 아너스</span> 희망목표 <strong>${fwo(tgtRaw)}</strong><br>
            <small>인보험평균 ${fwo(insRaw)} &nbsp;|&nbsp; 기본순증목표 ${baseTgtRaw ? fwo(baseTgtRaw) : "미입력"}</small>
          </div>
        </div>
        <div class="pa-total">
          <div>최종 예상 시상금 합계 (3개월) &nbsp; ① + ② + ③</div>
          <div class="pa-total-val">${fw(total)}</div>
        </div>
        <div class="pa-grid">
          <div class="pa-cell purple">
            <div class="pa-lbl">① 아너스클럽</div>
            <div class="pa-val">${aw1 ? fw(aw1) : "해당없음"}</div>
            <div class="pa-sub2">${escapeHtml(aw1Grade)}</div>
          </div>
          <div class="pa-cell green">
            <div class="pa-lbl">② 하이포인트 (3개월)</div>
            <div class="pa-val">${fw(aw2M3)}</div>
            <div class="pa-sub2">순증 ${fwo(incr * 10000)} × 50%</div>
          </div>
          <div class="pa-cell blue">
            <div class="pa-lbl">③ 마스터과정 (×2)</div>
            <div class="pa-val">${aw3 ? fw(aw3 * 2) : "해당없음"}</div>
            <div class="pa-sub2">${aw3Tier ? escapeHtml(aw3Tier.label) : "순증 5만원 미만"}</div>
          </div>
        </div>
        ${aw3Next ? `<div class="pa-next">
          🚀 다음 단계: <strong>${escapeHtml(aw3Next.criteria)}</strong> — 희망목표 +${fwo(Math.ceil(Math.max(0, aw3Next.critVal - incrMW3) * 10000))} 추가 시 달성
        </div>` : ""}
      </div>
    `;
  }

  // Step 2 면담일지 출력용 — Step 1 실제 실적(pgCurrent/pgBase) 기반 시상 결과
  function buildPrintAwardHtmlStep2(student, allItvs) {
    const fw = (mw) => Math.round(mw * 10000).toLocaleString() + "원";
    const fwo = (w) => Math.round(w).toLocaleString() + "원";
    const pgCurrent = Number(student?.pgCurrent || 0);
    const pgBase = Number(student?.base || 0);
    const curMW = pgCurrent / 10000;
    const baseMW = pgBase / 10000;
    const incrMW = Math.max(0, curMW - baseMW);

    // ① 아너스클럽
    let aw1 = 0, aw1Grade = "미달성";
    for (let i = 0; i < HONORS.length; i++) {
      if (curMW >= HONORS[i].critVal) { aw1 = HONORS[i].prize; aw1Grade = HONORS[i].grade; break; }
    }

    // ② 하이포인트
    const baseTgtMet = baseMW <= 0 || curMW >= baseMW;
    const mExtra = baseTgtMet ? Math.floor(incrMW * INCR_CFG.rate / 100 * 10) / 10 : 0;
    const mFinal = baseTgtMet ? Math.min(INCR_CFG.base + mExtra, INCR_CFG.mcap) : 0;
    const aw2M3 = baseTgtMet ? Math.min(mFinal * 3, INCR_CFG.qcap) : 0;

    // ③ 마스터과정
    const lastIns = allItvs.slice().reverse().find((e) => e.ins);
    const insRaw = (parseFloat(lastIns?.ins || "0") || 0) * 1000;
    const masterIncrMW = Math.max(0, pgCurrent - insRaw) / 10000;
    const aw3 = calcMasterAward(masterIncrMW);
    const aw3Tier = MASTER_AWARD.find((t) => masterIncrMW >= t.critVal);
    const total = aw1 + aw2M3 + aw3 * 2;

    return `
      <div class="print-award">
        <div class="pa-hdr">
          <div class="pa-t">📊 Step 1 실적 기준 시상 결과</div>
          <div class="pa-sub">
            <span class="pa-sub-k">Step 1 현재실적</span> <strong>${fwo(pgCurrent)}</strong><br>
            <small>기준실적 ${fwo(pgBase)} &nbsp;|&nbsp; 순증 ${fwo(Math.max(0, pgCurrent - pgBase))}</small>
          </div>
        </div>
        <div class="pa-total">
          <div>Step 1 시상금 합계 (3개월) &nbsp; ① + ② + ③</div>
          <div class="pa-total-val">${fw(total)}</div>
        </div>
        <div class="pa-grid">
          <div class="pa-cell purple">
            <div class="pa-lbl">① 아너스클럽</div>
            <div class="pa-val">${aw1 ? fw(aw1) : "미달성"}</div>
            <div class="pa-sub2">${escapeHtml(aw1Grade)}</div>
          </div>
          <div class="pa-cell green">
            <div class="pa-lbl">② 하이포인트 (3개월)</div>
            <div class="pa-val">${fw(aw2M3)}</div>
            <div class="pa-sub2">순증 ${fwo(Math.max(0, pgCurrent - pgBase))} × 50%</div>
          </div>
          <div class="pa-cell blue">
            <div class="pa-lbl">③ 마스터과정 (×2)</div>
            <div class="pa-val">${aw3 ? fw(aw3 * 2) : "해당없음"}</div>
            <div class="pa-sub2">${aw3Tier ? escapeHtml(aw3Tier.label) : "순증 5만원 미만"}</div>
          </div>
        </div>
      </div>
    `;
  }

  async function savePrintPNG() {
    if (typeof window.html2canvas !== "function") {
      toast("PNG 라이브러리 로딩 중입니다. 잠시 후 다시 시도하세요.", "error");
      return;
    }
    const area = $("#print-area");
    if (!area || !area.children.length) { toast("출력할 내용이 없습니다.", "error"); return; }
    const btn = $("#btn-save-png");
    const orig = btn ? btn.textContent : "";
    if (btn) { btn.disabled = true; btn.textContent = "저장중..."; }
    try {
      const pages = area.querySelectorAll(".print-page");
      const targets = pages.length ? [...pages] : [area];
      for (let i = 0; i < targets.length; i++) {
        const target = targets[i];
        const w = target.scrollWidth;
        const clone = target.cloneNode(true);
        clone.style.cssText = `position:fixed;left:-9999px;top:0;width:${w}px;overflow:visible;visibility:visible;pointer-events:none;z-index:-1;background:#fff;padding:10px;`;
        document.body.appendChild(clone);
        await new Promise((r) => setTimeout(r, 80));
        const canvas = await window.html2canvas(clone, {
          scale: 2, useCORS: true, backgroundColor: "#ffffff", logging: false,
          width: w, height: clone.scrollHeight,
          windowWidth: w, windowHeight: clone.scrollHeight
        });
        document.body.removeChild(clone);
        const link = document.createElement("a");
        link.download = makePngFilename(targets.length > 1 ? `_${i + 1}페이지` : "");
        link.href = canvas.toDataURL("image/png");
        link.click();
        if (i < targets.length - 1) await new Promise((r) => setTimeout(r, 300));
      }
      toast(`PNG ${targets.length}장 저장 완료`, "success");
    } catch (e) {
      console.error(e);
      toast("저장 중 오류: " + e.message, "error");
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = orig; }
    }
  }

  function makePngFilename(suffix) {
    const s = state.students.find((x) => x.empNo === state.selectedEmpNo);
    const now = new Date();
    const pad = (n) => String(n).padStart(2, "0");
    const ts = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}`;
    return `면담일지_${s ? s.name || s.empNo : ""}_${ts}${suffix || ""}.png`;
  }

  // 통계 탭에서 호출 — 현재 필터(지역단/비전센터) 기준 학생 전체 시상안 출력
  async function printAwardSheets() {
    const f = state.filter;
    if (!f.region) { toast("먼저 지역단을 선택하세요.", "error"); return; }
    let students = state.students.filter((s) => s.region === f.region);
    if (f.center) students = students.filter((s) => s.center === f.center);
    if (f.branch) students = students.filter((s) => s.branch === f.branch);
    if (f.cohort) students = students.filter((s) => !s.cohort || s.cohort === f.cohort);
    if (!students.length) { toast("현재 필터 범위에 교육생이 없습니다.", "error"); return; }
    if (students.length > 50 && !confirm(`${students.length}명의 시상안을 일괄 출력합니다. 진행할까요?`)) return;

    const win = window.open("", "_blank", "width=900,height=1000");
    if (!win) { toast("팝업이 차단되었습니다. 허용 후 다시 시도하세요.", "error"); return; }

    toast(`${students.length}명의 면담이력 수집중...`, "");
    await ensureConsultationsFetched(students);

    const pages = [];
    for (const s of students) {
      const itvs = getStudentConsultations(s.empNo);
      const sorted = itvs.slice().sort((a, b) => (b.date || "").localeCompare(a.date || ""));
      const calcItv = sorted.find((i) => i.calcAvg || i.calcTgt || i.calcBaseTgt) || null;
      const lastInsItv = sorted.find((i) => i.ins) || null;
      const html = buildAwardSheetPageHtml(s, calcItv, lastInsItv);
      if (html) pages.push(html);
    }

    if (!pages.length) { win.close(); toast("출력할 데이터가 없습니다.", "error"); return; }

    const scopeText = [f.region, f.center, f.branch].filter(Boolean).join(" · ");
    win.document.write(`<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">
<title>시상 예상답안지 — ${escapeHtml(scopeText)}</title>
<style>${AWARD_PRINT_CSS}
  /* 미리보기용 상단 바 — 화면에서만 노출, 인쇄 시 숨김 */
  .print-ovl {
    position: fixed; top: 0; left: 0; right: 0; z-index: 9999;
    background: #1A2744; color: #fff;
    display: flex; align-items: center; justify-content: space-between;
    padding: 10px 14px; box-shadow: 0 2px 8px rgba(0,0,0,0.2);
    font-family: inherit;
  }
  .print-ovl .po-ttl { font-size: 13px; font-weight: 700; color: #fff; letter-spacing: -0.2px; }
  .print-ovl button {
    border: none; border-radius: 6px; padding: 8px 14px;
    font-size: 13px; font-weight: 800; cursor: pointer;
    font-family: inherit; line-height: 1;
  }
  .po-back { background: #E8651A; color: #fff; }
  .po-print { background: #fff; color: #1A2744; }
  body { padding-top: 50px; }
  @media print {
    .print-ovl { display: none !important; }
    body { padding-top: 0 !important; }
  }
</style>
</head><body>
  <div class="print-ovl">
    <button type="button" class="po-back" onclick="window.close();return false;">← 돌아가기</button>
    <span class="po-ttl">🏆 시상 예상답안지 미리보기 (${pages.length}명)</span>
    <button type="button" class="po-print" onclick="window.print();return false;">🖨️ 인쇄하기</button>
  </div>
  ${pages.join("")}
</body></html>`);
    win.document.close();
  }

  // ========== 실적진도 (Progress) ==========
  // 교육생 doc 의 base(기준실적) / current(현재실적, 신규) / ipumCount / ipumAmt 사용
  // 현재실적·인품실적이 없으면 0으로 처리
  // 시상 규칙은 호남지역단 기준 하드코딩 상수 (추후 per-region override 가능)

  // PROGRESS_AWARDS 는 v1.22 부터 지역단별 시상안(localStorage)에서 동적 로드
  // → getProgressAwardConfig(region) 사용

  function openProgressRegionPicker() {
    // 학생이 존재하는 지역단만 추출
    const regions = [...new Set(state.students.map((s) => s.region).filter((r) => r && r.endsWith("지역단")))].sort();
    if (!regions.length) { toast("등록된 교육생이 없습니다.", "error"); return; }
    // 간단 모달 대신 prompt + select — 모달 재사용 대신 빠른 구현
    openPickerModal("지역단 선택", regions, (picked) => {
      state.progressRegion = picked;
      const pgRegSel = document.getElementById("pg-region-sel");
      if (pgRegSel) pgRegSel.value = picked;
      renderProgressPanel();
    });
  }

  // 간단 선택 모달 (목록 + 검색)
  function openPickerModal(title, options, onPick) {
    let modal = document.getElementById("modal-simple-picker");
    if (!modal) {
      modal = document.createElement("div");
      modal.id = "modal-simple-picker";
      modal.className = "modal";
      modal.hidden = true;
      modal.innerHTML = `
        <div class="modal-backdrop" data-close></div>
        <div class="modal-panel">
          <div class="modal-head">
            <h3 id="sp-title">선택</h3>
            <button class="modal-close" data-close>&times;</button>
          </div>
          <div class="modal-body">
            <input type="search" id="sp-search" class="modal-search" placeholder="검색...">
            <ul class="org-list" id="sp-list"></ul>
          </div>
        </div>
      `;
      document.body.appendChild(modal);
      modal.querySelectorAll("[data-close]").forEach((el) => el.addEventListener("click", () => { modal.hidden = true; }));
      modal.querySelector("#sp-search").addEventListener("input", (e) => renderSpList(e.target.value));
    }
    modal._opts = options;
    modal._onPick = onPick;
    modal.querySelector("#sp-title").textContent = title;
    modal.querySelector("#sp-search").value = "";
    function renderSpList(q) {
      const list = modal.querySelector("#sp-list");
      const needle = (q || "").toLowerCase();
      const filtered = modal._opts.filter((o) => o.toLowerCase().includes(needle));
      list.innerHTML = filtered.map((name) => `<li data-v="${escapeHtml(name)}">${escapeHtml(name)}</li>`).join("") || `<li class="disabled">검색 결과 없음</li>`;
      list.querySelectorAll("li[data-v]").forEach((li) => {
        li.addEventListener("click", () => {
          modal.hidden = true;
          modal._onPick(li.dataset.v);
        });
      });
    }
    renderSpList("");
    modal.hidden = false;
    setTimeout(() => modal.querySelector("#sp-search").focus(), 50);
  }

  // 미지정 교육생 알림 버튼 표시/숨김 업데이트
  function updateUnassignedAlert() {
    const wrap = document.getElementById("unassigned-alert-wrap");
    if (!wrap) return;
    const count = state.students.filter((s) => !s.region || !s.center).length;
    wrap.hidden = count === 0;
    const btn = wrap.querySelector("#btn-unassigned-alert");
    if (btn) btn.textContent = `⚠️ 미지정 교육생 확인 (${count}명)`;
  }

  // 미지정 교육생(지역단·비전센터 없음) 지정 모달
  function openUnassignedModal() {
    const students = state.students.filter((s) => !s.region || !s.center);
    if (!students.length) { toast("미지정 교육생이 없습니다.", ""); return; }

    const regions = [...new Set(state.students.map((s) => s.region).filter(Boolean))].sort();
    const cohorts = [...new Set(state.students.map((s) => s.cohort).filter(Boolean))].sort();

    const orgCentersByRegion = {};
    if (window.ORG_DATA?.regions) {
      for (const r of window.ORG_DATA.regions) orgCentersByRegion[r.name] = r.centers.map((c) => c.name);
    }
    const mkCenterOpts = (reg, cur) => {
      const cs = reg ? (orgCentersByRegion[reg] || []) : [];
      return `<option value="">-- 선택 --</option>` + cs.map((c) => `<option value="${escapeHtml(c)}"${c === cur ? " selected" : ""}>${escapeHtml(c)}</option>`).join("");
    };
    const mkCohortOpts = (cur) =>
      `<option value="">--</option>` + cohorts.map((c) => `<option value="${escapeHtml(c)}"${c === cur ? " selected" : ""}>${escapeHtml(c)}</option>`).join("");

    let modal = document.getElementById("modal-unassigned");
    if (!modal) {
      modal = document.createElement("div");
      modal.id = "modal-unassigned";
      modal.className = "modal";
      modal.style.cssText = "z-index:10003";
      document.body.appendChild(modal);
    }

    const renderRows = () => students.map((s, i) => `
      <tr data-idx="${i}" style="border-bottom:1px solid #eee">
        <td style="padding:5px 8px">${escapeHtml(s.name || "(이름없음)")}</td>
        <td style="padding:5px 8px;font-size:11px;color:#888">${escapeHtml(s.empNo)}</td>
        <td style="padding:5px 8px;font-size:11px">${escapeHtml(s.branch || "")}</td>
        <td style="padding:4px">
          <select class="ua-region side-input" data-idx="${i}" style="width:110px;font-size:12px">
            <option value="">-- 선택 --</option>
            ${regions.map((r) => `<option value="${escapeHtml(r)}"${r === (s.region || "") ? " selected" : ""}>${escapeHtml(r)}</option>`).join("")}
          </select>
        </td>
        <td style="padding:4px">
          <select class="ua-center side-input" data-idx="${i}" style="width:130px;font-size:12px">
            ${mkCenterOpts(s.region || "", s.center || "")}
          </select>
        </td>
        <td style="padding:4px;white-space:nowrap">
          <select class="ua-cohort side-input" data-idx="${i}" style="width:60px;font-size:12px">
            ${mkCohortOpts(s.cohort || "")}
          </select>
          <label style="margin-left:6px;font-size:11px;color:#e53935;cursor:pointer;white-space:nowrap">
            <input type="checkbox" class="ua-del-chk" data-idx="${i}" style="accent-color:#e53935;vertical-align:middle"> 삭제
          </label>
        </td>
      </tr>`).join("");

    modal.innerHTML = `
      <div class="modal-backdrop"></div>
      <div class="modal-panel xwide" style="max-height:85vh;display:flex;flex-direction:column">
        <div class="modal-head">
          <h3 id="ua-title" style="font-size:15px">미지정 교육생 — ${students.length}명</h3>
          <button class="modal-close" id="ua-close">&times;</button>
        </div>
        <div style="flex:1;overflow-y:auto;padding:12px 16px">
          <p style="font-size:12px;color:#888;margin:0 0 10px">지역단·비전센터가 없는 교육생을 지정한 뒤 저장하세요. 삭제할 행은 <span style="color:#e53935;font-weight:600">삭제</span> 체크 후 저장하세요.</p>
          <table style="width:100%;border-collapse:collapse;font-size:13px">
            <thead>
              <tr style="background:#f0f2f5;font-size:12px">
                <th style="padding:6px 8px;text-align:left;font-weight:600">이름</th>
                <th style="padding:6px 8px;text-align:left;font-weight:600">사번</th>
                <th style="padding:6px 8px;text-align:left;font-weight:600">지점</th>
                <th style="padding:6px 8px;font-weight:600">지역단</th>
                <th style="padding:6px 8px;font-weight:600">비전센터</th>
                <th style="padding:6px 8px;font-weight:600;white-space:nowrap">기수 / <label style="color:#e53935;cursor:pointer;font-weight:600" title="전체 선택/해제"><input type="checkbox" id="ua-del-all" style="accent-color:#e53935;vertical-align:middle;margin-right:3px">전체삭제</label></th>
              </tr>
            </thead>
            <tbody id="ua-tbody">${renderRows()}</tbody>
          </table>
        </div>
        <div style="padding:10px 16px;border-top:1px solid #e0e0e0;display:flex;gap:10px;justify-content:flex-end;align-items:center">
          <span id="ua-msg" style="font-size:13px;flex:1"></span>
          <button class="btn-outline small" id="ua-cancel">취소</button>
          <button class="btn-primary small" id="ua-save">💾 저장</button>
        </div>
      </div>
    `;

    const hide = () => { modal.hidden = true; };
    modal.querySelector("#ua-close").onclick  = hide;
    modal.querySelector("#ua-cancel").onclick = hide;
    modal.querySelector(".modal-backdrop").onclick = hide;

    // 지역단 변경 → 비전센터 옵션 갱신
    modal.querySelectorAll(".ua-region").forEach((sel) => {
      sel.addEventListener("change", () => {
        const i = sel.dataset.idx;
        modal.querySelector(`.ua-center[data-idx="${i}"]`).innerHTML = mkCenterOpts(sel.value, "");
      });
    });

    // 전체 삭제 체크박스
    const delAllChk = modal.querySelector("#ua-del-all");
    if (delAllChk) {
      delAllChk.addEventListener("change", () => {
        modal.querySelectorAll(".ua-del-chk").forEach((chk) => {
          if (chk.checked === delAllChk.checked) return;
          chk.checked = delAllChk.checked;
          const row = chk.closest("tr");
          if (chk.checked) {
            row.style.background = "#fff0f0";
            row.style.opacity = "0.7";
            row.querySelectorAll("select").forEach((s) => s.disabled = true);
          } else {
            row.style.background = "";
            row.style.opacity = "";
            row.querySelectorAll("select").forEach((s) => s.disabled = false);
          }
        });
      });
    }

    // 삭제 체크박스 → 행 배경색 토글
    modal.querySelector("#ua-tbody").addEventListener("change", (e) => {
      const chk = e.target.closest(".ua-del-chk");
      if (!chk) return;
      const row = chk.closest("tr");
      if (chk.checked) {
        row.style.background = "#fff0f0";
        row.style.opacity = "0.7";
        row.querySelectorAll("select").forEach((s) => s.disabled = true);
      } else {
        row.style.background = "";
        row.style.opacity = "";
        row.querySelectorAll("select").forEach((s) => s.disabled = false);
        if (delAllChk) delAllChk.checked = false;
      }
    });

    // 저장 (삭제 체크된 행은 영구 삭제, 나머지는 저장)
    modal.querySelector("#ua-save").addEventListener("click", async () => {
      const btn = modal.querySelector("#ua-save");
      const msg = modal.querySelector("#ua-msg");
      const toDelete = [];
      const toSave = [];
      let incomplete = 0;

      modal.querySelectorAll("tbody tr[data-idx]").forEach((row) => {
        const i  = parseInt(row.dataset.idx, 10);
        const del = row.querySelector(".ua-del-chk")?.checked;
        if (del) { toDelete.push(students[i]); return; }
        const r  = row.querySelector(".ua-region").value;
        const c  = row.querySelector(".ua-center").value;
        const co = row.querySelector(".ua-cohort").value;
        if (!r || !c) { incomplete++; return; }
        toSave.push({ ...students[i], region: r, center: c, ...(co ? { cohort: co } : {}) });
      });

      if (incomplete) { msg.textContent = `⚠️ ${incomplete}명의 지역단·비전센터를 선택하세요.`; return; }
      if (!toDelete.length && !toSave.length) { hide(); return; }

      btn.disabled = true;
      msg.textContent = "처리중...";
      try {
        if (toDelete.length) {
          for (const s of toDelete) {
            const docId = s._docId || s.empNo;
            if (docId) await window.DataAPI.removeByDocId(docId);
          }
        }
        if (toSave.length) {
          if (typeof window.DataAPI.saveMany === "function") await window.DataAPI.saveMany(toSave);
          else for (const rec of toSave) await window.DataAPI.save(rec);
        }
        const parts = [];
        if (toSave.length)   parts.push(`${toSave.length}명 저장`);
        if (toDelete.length) parts.push(`${toDelete.length}명 삭제`);
        msg.textContent = `✅ ${parts.join(", ")} 완료`;
        toast(parts.join(", ") + " 완료", "success");
        setTimeout(hide, 1500);
      } catch (e) {
        msg.textContent = "❌ 처리 실패: " + e.message;
        btn.disabled = false;
      }
    });

    modal.hidden = false;
  }

  // 재사용 가능한 Yes/No 확인 모달 (Promise 반환)
  function openConfirmModal(msg) {
    return new Promise((resolve) => {
      let modal = document.getElementById("modal-confirm");
      if (!modal) {
        modal = document.createElement("div");
        modal.id = "modal-confirm";
        modal.className = "modal";
        modal.hidden = true;
        modal.innerHTML = `
          <div class="modal-backdrop"></div>
          <div class="modal-panel" style="max-width:360px">
            <div class="modal-body" style="padding:28px 24px 20px">
              <p id="confirm-msg" style="margin:0 0 22px;font-size:14px;line-height:1.6;text-align:center;white-space:pre-wrap"></p>
              <div style="display:flex;gap:10px;justify-content:center">
                <button class="btn-primary small" id="confirm-yes">✅ 예, 저장</button>
                <button class="btn-outline small" id="confirm-no">❌ 아니오</button>
              </div>
            </div>
          </div>
        `;
        document.body.appendChild(modal);
      }
      modal.querySelector("#confirm-msg").textContent = msg;
      const yesEl = modal.querySelector("#confirm-yes");
      const noEl  = modal.querySelector("#confirm-no");
      const yes = yesEl.cloneNode(true); yesEl.replaceWith(yes);
      const no  = noEl.cloneNode(true);  noEl.replaceWith(no);
      const hide = (val) => { modal.hidden = true; modal.querySelector(".modal-backdrop").onclick = null; resolve(val); };
      yes.addEventListener("click", () => hide(true));
      no.addEventListener("click",  () => hide(false));
      modal.querySelector(".modal-backdrop").onclick = () => hide(false);
      modal.hidden = false;
    });
  }

  function openPasteSaveConfirmModal(region, cohort, step, updateRecords, newRecords) {
    return new Promise((resolve) => {
      let modal = document.getElementById("modal-paste-confirm");
      if (!modal) {
        modal = document.createElement("div");
        modal.id = "modal-paste-confirm";
        modal.className = "modal";
        document.body.appendChild(modal);
      }

      const mkList = (records) => records.length
        ? records.map((r) => `<li>${r.name || "(이름없음)"} &nbsp;<span style="color:#888;font-size:11px">${r.empNo}</span></li>`).join("")
        : `<li style="color:#aaa">없음</li>`;

      const scrollStyle = (records) => records.length > 10 ? "max-height:200px;overflow-y:auto;" : "";

      modal.innerHTML = `
        <div class="modal-backdrop"></div>
        <div class="modal-panel" style="max-width:480px">
          <div class="modal-body" style="padding:24px 24px 20px">
            <h3 style="margin:0 0 16px;font-size:15px;text-align:center">${region} ${cohort} Step ${step} 저장 확인</h3>
            <p style="font-size:13px;font-weight:600;margin:0 0 5px">업데이트 (${updateRecords.length}명)</p>
            <ul class="paste-confirm-list" style="${scrollStyle(updateRecords)}">${mkList(updateRecords)}</ul>
            <p style="font-size:13px;font-weight:600;margin:14px 0 5px">신규 등록 (${newRecords.length}명)</p>
            <ul class="paste-confirm-list" style="${scrollStyle(newRecords)}">${mkList(newRecords)}</ul>
            <div style="display:flex;gap:10px;justify-content:center;margin-top:20px">
              <button class="btn-primary small" id="paste-confirm-yes">✅ 저장</button>
              <button class="btn-outline small" id="paste-confirm-no">❌ 취소</button>
            </div>
          </div>
        </div>
      `;

      const hide = (val) => { modal.hidden = true; resolve(val); };
      modal.querySelector("#paste-confirm-yes").addEventListener("click", () => hide(true));
      modal.querySelector("#paste-confirm-no").addEventListener("click",  () => hide(false));
      modal.querySelector(".modal-backdrop").onclick = () => hide(false);
      modal.hidden = false;
    });
  }

  function openPgColMapModal(colDefs, sampleRows) {
    return new Promise((resolve) => {
      const modal = document.getElementById("modal-pg-col-map");
      if (!modal) { resolve(null); return; }
      const wrap       = document.getElementById("pg-col-map-wrap");
      const confirmBtn = modal.querySelector("#btn-pg-col-map-confirm");
      const cancelBtn  = modal.querySelector("#btn-pg-col-map-cancel");
      const backdrop   = modal.querySelector(".modal-backdrop");

      const colStates = colDefs.map((c) => ({ label: c.label, field: c.field, deleted: false }));

      function buildTable() {
        const sc = Math.min(sampleRows.length, 3);
        let html = `<table class="pg-col-map-tbl"><thead><tr>
          <th>#</th><th>원본 제목</th><th>저장 필드</th>`;
        for (let r = 0; r < sc; r++) html += `<th>예시 ${r + 1}행</th>`;
        html += `<th></th></tr></thead><tbody>`;

        colStates.forEach((col, i) => {
          const del = col.deleted;
          html += `<tr class="${del ? "pg-col-del" : ""}" data-ci="${i}">
            <td class="pg-col-num">${i + 1}</td>
            <td class="pg-col-lbl">${col.label}</td>
            <td class="pg-col-sel"><select class="col-field-sel" data-ci="${i}" ${del ? "disabled" : ""}>`;
          PG_FIELD_OPTIONS.forEach((o) => {
            html += `<option value="${o.value}"${(del ? "ignore" : col.field) === o.value ? " selected" : ""}>${o.label}</option>`;
          });
          html += `</select></td>`;
          for (let r = 0; r < sc; r++) {
            html += `<td class="pg-col-sample">${(sampleRows[r] || [])[i] ?? ""}</td>`;
          }
          html += `<td class="pg-col-act"><button class="btn-outline small col-del-btn" data-ci="${i}">${del ? "↩ 복원" : "✕ 삭제"}</button></td>
          </tr>`;
        });
        html += `</tbody></table>`;
        wrap.innerHTML = html;

        wrap.querySelectorAll(".col-field-sel").forEach((sel, i) => {
          sel.addEventListener("change", (e) => { colStates[i].field = e.target.value; });
        });
        wrap.querySelectorAll(".col-del-btn").forEach((btn) => {
          btn.addEventListener("click", (e) => {
            colStates[+e.currentTarget.dataset.ci].deleted = !colStates[+e.currentTarget.dataset.ci].deleted;
            buildTable();
          });
        });
      }

      buildTable();

      const done = (result) => { modal.hidden = true; backdrop.onclick = null; resolve(result); };
      confirmBtn.onclick = () => {
        const mapping = colStates.map((c) => (c.deleted ? "ignore" : c.field));
        if (!mapping.includes("empNo")) {
          toast("사원번호 열이 지정되지 않아 저장할 수 없습니다.", "error");
          return;
        }
        done(mapping);
      };
      cancelBtn.onclick = () => done(null);
      backdrop.onclick  = () => done(null);
      modal.hidden = false;
    });
  }

  function openProgressAdminOverlay() {
    const region = state.filter.region || "";
    if (!region) { toast("좌측 필터에서 지역단을 먼저 선택하세요.", "error"); return; }
    const list = state.students.filter((s) => s.region === region);
    const overlay = document.getElementById("pg-admin-overlay");
    const body = document.getElementById("pg-admin-overlay-body");
    if (!overlay || !body) return;
    const titleEl = overlay.querySelector(".pg-admin-overlay-title");
    if (titleEl) titleEl.textContent = `⚙️ 실적관리 — ${region}`;
    body.innerHTML = renderProgressAdmin(list);
    overlay.hidden = false;
    bindProgressAdminEvents(list, "pg-admin-overlay-body");
  }

  function renderProgressPanel() {
    const body = $("#progress-body");
    if (!body) return;
    // 지역단 선택 셀렉트 옵션 갱신 (교육생 데이터 기준)
    const pgRegSel = document.getElementById("pg-region-sel");
    if (pgRegSel) {
      const allRegions = [...new Set(state.students.map((s) => s.region).filter(Boolean))].sort();
      const prev = pgRegSel.value;
      pgRegSel.innerHTML = '<option value="">지역단 미선택</option>' +
        allRegions.map((r) => `<option value="${escapeHtml(r)}">${escapeHtml(r)}</option>`).join("");
      const want = prev || state.progressRegion || state.filter.region || "";
      if (want) pgRegSel.value = want;
    }
    const region = pgRegSel?.value || state.progressRegion || state.filter.region || "";
    state.progressRegion = region;

    if (!region) {
      body.innerHTML = `<div class="empty-state">위 지역단 선택에서 <strong>지역단</strong>을 선택하면 해당 지역단의 실적진도가 표시됩니다.</div>`;
      return;
    }
    // 좌측 사이드바 기수(state.filter.cohort)를 우선 사용; 없으면 진도탭 자체 선택값
    const pgCohort = (state.filter.cohort || state.progressCohort || "").replace(/기$/, "");
    if (pgCohort) state.progressCohort = pgCohort; // 두 선택자 동기화
    // pg-cohort-sel / pg-step-sel 과 사이드바 필터를 동기화
    const _pgCohortSel = document.getElementById("pg-cohort-sel");
    if (_pgCohortSel && _pgCohortSel.value !== pgCohort) _pgCohortSel.value = pgCohort;
    const _pgStepVal = state.filter.step || state.progressStep || "1";
    state.progressStep = _pgStepVal; // filter.step 와 progressStep 항상 동기화
    const _pgStepSel = document.getElementById("pg-step-sel");
    if (_pgStepSel && _pgStepSel.value !== _pgStepVal) _pgStepSel.value = _pgStepVal;
    const list = state.students.filter((s) => {
      if (s.region !== region) return false;
      if (pgCohort && s.cohort && String(s.cohort).replace("기", "") !== String(pgCohort)) return false;
      return true;
    });
    if (!list.length) {
      body.innerHTML = `<div class="empty-state">${escapeHtml(region)}${pgCohort ? ` ${pgCohort}기` : ""} 에 등록된 교육생이 없습니다.</div>`;
      return;
    }

    if (state.progressSubTab === "admin") {
      body.innerHTML = renderProgressAdmin(list);
      bindProgressAdminEvents(list);
    } else {
      body.innerHTML = renderProgressHome(list);
      bindProgressHomeEvents(list);
    }
  }

  function _pgStepSfx() {
    const s = String(state.filter.step || state.progressStep || "1");
    return s === "1" ? "" : s;
  }

  // 학생 데이터에서 계산된 지표 얻기
  function getProgressStat(s) {
    const sfx = _pgStepSfx();
    const base    = Number(s.base || 0);
    const current = sfx
      ? Number(s[`pgCurrent${sfx}`] || 0)
      : Number(s.pgCurrent || s.current || 0);
    const hiCap     = Number(s[`hiCap${sfx}`]        || 0);
    const ipumCount = Number(s[`pgIpumCount${sfx}`]  || 0);
    const ipumAmt   = Number(s[`pgIpumAmt${sfx}`]    || 0);
    const net  = current - base;
    const rate = base > 0 ? (current / base) * 100 : 0;
    return { s, base, current, hiCap, net, rate, ipumCount, ipumAmt };
  }

  // 지역단별 시상안 → PROGRESS_AWARDS 호환 객체 생성
  // cohortOverride/stepOverride 를 전달하면 state.progressCohort/Step 대신 사용
  function getProgressAwardConfig(region, cohortOverride, stepOverride) {
    // state.filter.cohort 는 "1기" 형식, 시상안 키는 "1" 형식 — "기" 제거 후 비교
    const _cohort = (cohortOverride || state.progressCohort || "").replace(/기$/, "");
    const _step   = stepOverride   || state.filter.step || state.progressStep;
    let _planKey = region;
    if (_cohort && _step) {
      const _y = state.progressYear || String(new Date().getFullYear());
      const _ck = makeAwardPlanKey(_y, region, _cohort, _step);
      try {
        const _st = JSON.parse(localStorage.getItem(LS_AWARD_PLANS_KEY) || "{}");
        if (_st[_ck]) _planKey = _ck;
      } catch { /**/ }
    }
    const plan = getAwardPlan(_planKey);
    const personalItems = (plan.personalIncr?.enabled ? (plan.personalIncr.items || []) : []);
    const tiers = personalItems
      .filter((it) => Number(it.critVal) > 0)
      .slice()
      .sort((a, b) => Number(b.critVal) - Number(a.critVal))
      .map((it) => ({
        min: Number(it.critVal) * 10000,
        type: it.payType,
        val: it.payType === "pct" ? Number(it.payVal) / 100 : it.payType === "item" ? 0 : Number(it.payVal) * 10000,
        payVal: it.payType === "item" ? 0 : Number(it.payVal),
        itemName: it.payType === "item" ? String(it.payVal || "물품") : "",
        payType: it.payType
      }));
    const top1 = plan.topAward1?.enabled ? plan.topAward1 : null;
    const top2 = plan.topAward2?.enabled ? plan.topAward2 : null;
    const rateConfig = (top1 && top1.type === "rate") ? top1 : (top2 && top2.type === "rate") ? top2 : null;
    const amtConfig  = (top1 && top1.type === "amt")  ? top1 : (top2 && top2.type === "amt")  ? top2 : null;
    const rateTop = rateConfig ? Array.from({ length: Number(rateConfig.n) }, (_, i) => calcRankAward(i + 1, rateConfig)) : [];
    const amtTop  = amtConfig  ? Array.from({ length: Number(amtConfig.n)  }, (_, i) => calcRankAward(i + 1, amtConfig))  : [];
    // Step 2 이상이면 sfx="2"/"3"…, Step 1은 sfx=""
    const _sfx = (_step && _step !== "1") ? String(_step) : "";
    return {
      plan,
      tiers,
      rateTop10: rateTop,
      amtTop10:  amtTop,
      rateConfig,
      amtConfig,
      bothEnabled: !!(rateConfig && amtConfig) && !!plan.bothNodup,
      isEligible: (student) => isEligibleForAward(student, plan, _sfx),
      // Step 2+: TOP10 순위시상은 환산실적 조건 미적용 — 순증 기준으로만 판단
      isTopEligible: _sfx ? () => true : (student) => isEligibleForAward(student, plan, _sfx)
    };
  }

  function tierAward(net, region, pa) {
    const _pa = pa || getProgressAwardConfig(region || state.progressRegion || DEFAULT_REGION);
    for (const t of _pa.tiers) {
      if (net >= t.min) {
        return t.type === "pct" ? Math.round(net * t.val) : t.val;
      }
    }
    return 0;
  }
  function tierLabel(net, region, pa) {
    const _pa = pa || getProgressAwardConfig(region || state.progressRegion || DEFAULT_REGION);
    for (const t of _pa.tiers) {
      if (net >= t.min) {
        if (t.type === "pct") return `${t.payVal}% 지급`;
        if (t.type === "item") return t.itemName || "물품";
        return `${t.payVal}만원`;
      }
    }
    return "-";
  }
  const Nf = (v) => Math.round(Number(v) || 0).toLocaleString();
  const RB = (r) => {
    if (!r) return `<span style="color:#ccc;font-size:10px;">-</span>`;
    const cls = r === 1 ? "r1" : r === 2 ? "r2" : r === 3 ? "r3" : "rt";
    return `<span class="pg-rb ${cls}">${r}</span>`;
  };

  // Normalize groupAward2 to always return items array (backward compat with old single-item shape)
  function _ga2Items(ga2) {
    if (!ga2?.enabled) return [];
    if (ga2.items?.length) return ga2.items;
    return [{ rateThreshold: ga2.rateThreshold ?? 110, payout: ga2.payout ?? 15 }];
  }
  function _ga1Items(ga1) {
    if (!ga1?.enabled) return [];
    if (ga1?.items?.length) return ga1.items;
    return [{ threshold: ga1?.threshold ?? 5, payout: normPayout(ga1?.payout ?? 5) }];
  }

  function renderGroupTable(groupRanking, plan, hasAnyTeam) {
    const ga1 = plan?.groupAward1;
    const ga2 = plan?.groupAward2;
    const ga1En = !!ga1?.enabled;
    const ga2En = !!ga2?.enabled;
    const ga1ItemList = _ga1Items(ga1);
    const ga1PrimaryThr = ga1En && ga1ItemList.length ? Number(ga1ItemList[0].threshold || 5) : 5;

    const thCells = [
      `<th>#</th>`,
      `<th>${hasAnyTeam ? "팀" : "지점"}</th>`,
      `<th>구성원</th>`,
      `<th>목표합계<br><small>(천원)</small></th>`,
      `<th>현재합계<br><small>(천원)</small></th>`,
      `<th>달성률</th>`,
      ...(ga1En ? [`<th>전원${ga1PrimaryThr}만↑</th>`, `<th>그룹시상1</th>`] : []),
      ...(ga2En ? [`<th>그룹시상2</th>`] : [])
    ].join("");

    const rows = groupRanking.map((g, i) => {
      const short = g.members.slice(0, 5).join("·") + (g.members.length > 5 ? "…" : "");
      let metGa1Items = []; // ga2 연계 판단을 위해 상위 스코프로 선언
      let ga1Cell = "", ga1AwardCell = "";
      if (ga1En) {
        const primaryIt = ga1ItemList[0] || { threshold: 5, payout: { type: "cash", val: 5 } };
        const thr = Number(primaryIt.threshold || 5) * 10000;
        const total = (g.memberStats || []).length || g.members.length;
        const achieved = (g.memberStats || []).filter(st => st.net >= thr).length;
        const allAchieved = achieved === total && total > 0;
        const icon = allAchieved ? "✅" : (achieved === 0 ? "✕" : "△");
        const bdg = allAchieved ? "pg-grp-ok" : (achieved > 0 ? "pg-grp-half" : "pg-grp-no");
        metGa1Items = ga1ItemList.filter(it => {
          const itThr = Number(it.threshold || 5) * 10000;
          return (g.memberStats || []).every(st => (st.net || 0) >= itThr) && total > 0;
        });
        const awardLabels = metGa1Items.map(it => escapeHtml(payoutLabel(it.payout))).join("+");
        ga1Cell = `<td><span class="pg-grp-badge ${bdg}">${icon}${achieved}/${total}명</span></td>`;
        ga1AwardCell = `<td>${metGa1Items.length ? `<span class="pg-grp-award">${awardLabels}</span>` : `<span class="pg-grp-miss">미달</span>`}</td>`;
      }
      let ga2Cell = "";
      if (ga2En) {
        const g2items = _ga2Items(ga2);
        // 그룹시상1이 활성화된 경우 항상 연계 — ga1 미달이면 ga2도 미달
        if (ga1En && !metGa1Items.length) {
          ga2Cell = `<td><span class="pg-grp-miss">미달(그룹1↑필요)</span></td>`;
        } else {
          const metGa2Items = g2items.filter(it => g.rate >= Number(it.rateThreshold || 110));
          if (metGa2Items.length > 0) {
            const bestPay = metGa2Items.reduce((best, it) => {
              const np = normPayout(it.payout ?? 15);
              if (np.type === "item") return best || np;
              return (!best || Number(np.val) > Number(normPayout(best).val)) ? np : best;
            }, null);
            ga2Cell = `<td><span class="pg-grp-award">${escapeHtml(payoutLabel(bestPay))}</span></td>`;
          } else {
            const firstThr = g2items[0]?.rateThreshold || 110;
            ga2Cell = `<td><span class="pg-grp-miss">미달(${firstThr}%↑)</span></td>`;
          }
        }
      }
      return `<tr><td>${RB(i + 1)}</td><td><strong>${escapeHtml(g.name)}</strong></td><td><small>${escapeHtml(short)}</small></td><td class="r">${Nf(Math.round(g.base/1000))}</td><td class="r">${Nf(Math.round(g.current/1000))}</td><td>${g.rate.toFixed(1)}%</td>${ga1Cell}${ga1AwardCell}${ga2Cell}</tr>`;
    }).join("");

    return `<table class="pg-tbl pg-tbl-wide"><thead><tr>${thCells}</tr></thead><tbody>${rows}</tbody></table>`;
  }

  function renderTeamCards(groupRanking, plan) {
    const ga1 = plan?.groupAward1;
    const ga2 = plan?.groupAward2;
    const ga1En = !!ga1?.enabled;
    const ga2En = !!ga2?.enabled;
    const ga1PrimaryItem = _ga1Items(ga1)[0];
    const ga1Thr = ga1En && ga1PrimaryItem ? Number(ga1PrimaryItem.threshold || 5) * 10000 : 0;

    return groupRanking.map((g) => {
      const mStats = g.memberStats || [];
      const total = mStats.length || g.members.length;
      const ga1Ach = ga1En ? mStats.filter(st => (st.net || 0) >= ga1Thr).length : 0;
      const ga1All = ga1En && ga1Ach === total && total > 0;
      // ga1 활성화 시 항상 연계 — ga1 미달이면 ga2도 미달
      const ga2Met = ga2En && (!ga1En || ga1All) && _ga2Items(ga2).some(it => g.rate >= Number(it.rateThreshold || 110));
      let cardCls = "tac-gray";
      if (ga1En || ga2En) {
        if ((!ga1En || ga1All) && (!ga2En || ga2Met)) cardCls = "tac-green";
        else if (ga1Ach > 0 || ga2Met) cardCls = "tac-orange";
      }
      const memberSpans = mStats.map(st => {
        const mNet = st.net || 0;
        const missed = mNet < (ga1Thr > 0 ? ga1Thr : 50000);
        const mName = st.s?.name || "";
        const mEmp = st.s?.empNo || "";
        return `<span class="tac-member${missed ? " tac-member-miss" : ""}${mEmp ? " hr-member-click" : ""}" data-emp="${escapeHtml(mEmp)}">${escapeHtml(mName)}</span>`;
      }).join("");
      const badges = [];
      if (ga1En) {
        const bc = ga1All ? "tac-badge-ok" : (ga1Ach > 0 ? "tac-badge-half" : "tac-badge-no");
        badges.push(`<span class="tac-badge ${bc}">${ga1All ? "✅" : (ga1Ach > 0 ? "△" : "✕")} ${ga1Ach}/${total}명</span>`);
      }
      if (ga2En) {
        badges.push(`<span class="tac-badge ${ga2Met ? "tac-badge-ok" : "tac-badge-no"}">그룹2 ${ga2Met ? "✅" : "✕"}</span>`);
      }
      return `<div class="tac-card ${cardCls}" data-teamcard="${escapeHtml(g.name)}">
        <div class="tac-team-name">${escapeHtml(g.name)}팀</div>
        <div class="tac-rate">${g.rate.toFixed(1)}%</div>
        <div class="tac-member-wrap">${memberSpans || "<span class='tac-member'>구성원 없음</span>"}</div>
        ${badges.length ? `<div class="tac-badges">${badges.join("")}</div>` : ""}
      </div>`;
    }).join("");
  }

  function renderGroupAwardInfo(plan) {
    const ga1 = plan?.groupAward1;
    const ga2 = plan?.groupAward2;
    const boxes = [];
    if (ga1?.enabled) {
      _ga1Items(ga1).forEach(it => {
        const thr = Number(it.threshold || 5);
        boxes.push(`<div class="gai-box">
          <div class="gai-title">🏅 그룹시상1</div>
          <div class="gai-desc">팀원 전원 순증 ${thr}만원↑</div>
          <div class="gai-payout">달성 시 ${escapeHtml(payoutLabel(it.payout ?? 5))}/인</div>
        </div>`);
      });
    }
    if (ga2?.enabled) {
      const g2items = _ga2Items(ga2);
      g2items.forEach(it => {
        const rate = Number(it.rateThreshold || 110);
        boxes.push(`<div class="gai-box">
          <div class="gai-title">🏅 그룹시상2</div>
          <div class="gai-desc">팀 달성률 ${rate}%↑</div>
          <div class="gai-payout">달성 시 ${escapeHtml(payoutLabel(it.payout ?? 15))}/팀</div>
        </div>`);
      });
    }
    return boxes.length ? `<div class="gai-boxes">${boxes.join("")}</div>` : "";
  }

  function renderTeamAwardFull(groupRanking, plan, hasAnyTeam) {
    if (!hasAnyTeam) {
      return `<div class="taf-no-team-notice">팀구분이 없습니다<span class="taf-notice-sub">지점별 달성률 현황</span></div>
        <div class="pg-tbl-wrap">${renderGroupTable(groupRanking, plan, hasAnyTeam)}</div>`;
    }
    const cards = renderTeamCards(groupRanking, plan);
    const awardInfo = renderGroupAwardInfo(plan);
    return `<div class="tac-grid">${cards}</div>
      ${awardInfo}
      <div class="pg-tbl-wrap taf-tbl-section"><h5 class="taf-tbl-title">📊 전체 순위표</h5>${renderGroupTable(groupRanking, plan, hasAnyTeam)}</div>`;
  }

  function buildTeamDetailHTML(group, plan) {
    const ga1 = plan?.groupAward1;
    const ga2 = plan?.groupAward2;
    const ga1En = !!ga1?.enabled;
    const ga2En = !!ga2?.enabled;
    const ga1ItemsArr = _ga1Items(ga1);
    const ga1Thr = ga1En && ga1ItemsArr.length ? Number(ga1ItemsArr[0].threshold || 5) * 10000 : 0;
    const mStats = group.memberStats || [];
    const total = mStats.length || group.members.length;
    const memberRows = mStats.map(st => {
      const net = st.net || 0;
      const rate = st.rate != null ? st.rate.toFixed(1) : "—";
      const ga1Hit = ga1Thr > 0 ? (net >= ga1Thr ? "✅" : `<span class="taf-red">✕</span>`) : "—";
      const netFmt = net >= 0 ? `+${Nf(net)}` : Nf(net);
      const netMiss = net < (ga1Thr > 0 ? ga1Thr : 50000);
      return `<tr class="pg-tr-click" data-emp="${escapeHtml(st.s?.empNo || "")}">
        <td><strong>${escapeHtml(st.s?.name || "")}</strong></td>
        <td class="r">${Nf(st.base || 0)}</td>
        <td class="r">${Nf(st.current || 0)}</td>
        <td class="r${netMiss ? " taf-red" : ""}">${netFmt}</td>
        <td class="r">${rate}%</td>
        <td class="r">${ga1Hit}</td>
      </tr>`;
    }).join("");
    const ga1Ach = ga1En ? mStats.filter(st => (st.net || 0) >= ga1Thr).length : 0;
    const ga1All = ga1En && ga1Ach === total && total > 0;
    let awardBoxes = "";
    if (ga1En) {
      ga1ItemsArr.forEach(it => {
        const itThr = Number(it.threshold || 5) * 10000;
        const itPayObj = normPayout(it.payout ?? 5);
        const missed = mStats.filter(st => (st.net || 0) < itThr);
        const allMet = missed.length === 0 && total > 0;
        const missedNames = missed.map(st => escapeHtml(st.s?.name || "")).join(", ");
        const perPerson = payoutLabel(itPayObj);
        const totalLabel = itPayObj.type === "item" ? `전원 ${perPerson}` : `팀 합계 ${Number(itPayObj.val || 5) * total}만원 지급`;
        awardBoxes += `<div class="td-award-box ${allMet ? "td-award-ok" : "td-award-miss"}">
          <div class="td-award-title">그룹시상1 — 전원 순증 ${Number(it.threshold || 5)}만원↑</div>
          <div class="td-award-detail">${allMet ? `✅ 전원 달성! 1인당 ${escapeHtml(perPerson)}` : `미달 ${missed.length}명: ${missedNames || "-"}`}</div>
          <div class="td-award-payout">${allMet ? escapeHtml(totalLabel) : "시상 미달"}</div>
        </div>`;
      });
    }
    if (ga2En) {
      _ga2Items(ga2).forEach((it) => {
        const rateThr = Number(it.rateThreshold || 110);
        const met = group.rate >= rateThr;
        awardBoxes += `<div class="td-award-box ${met ? "td-award-ok" : "td-award-miss"}">
          <div class="td-award-title">그룹시상2 — 달성률 ${rateThr}%↑</div>
          <div class="td-award-detail">현재 달성률 ${group.rate.toFixed(1)}% ${met ? "✅ 달성" : "✕ 미달"}</div>
          <div class="td-award-payout">${met ? `${escapeHtml(payoutLabel(it.payout ?? 15))}/팀 지급` : "시상 미달"}</div>
        </div>`;
      });
    }
    const cheers = ["함께하면 반드시 이긴다! 화이팅! 🔥", "오늘의 노력이 내일의 결과다! 💪", "우리 팀은 최고다! 전진! 🚀", "포기하지 않는 팀이 승리한다! ✊", "한 명도 포기 없이, 모두 함께 달성하자! 🏆"];
    const cheer = cheers[Math.floor(Math.random() * cheers.length)];
    return `<div class="td-stats-tbl-wrap"><table class="pg-tbl">
      <thead><tr><th>성명</th><th>기준(원)</th><th>현재(원)</th><th>순증(원)</th><th>달성률</th><th>${ga1Thr > 0 ? Math.round(ga1Thr / 10000) + "만↑" : "-"}</th></tr></thead>
      <tbody>${memberRows || "<tr><td colspan='6' class='pg-empty'>구성원 없음</td></tr>"}</tbody>
    </table></div>
    ${awardBoxes ? `<div class="td-award-boxes">${awardBoxes}</div>` : ""}
    <div class="td-cheer"><div class="td-cheer-label">🎯 오늘의 팀 구호</div><div class="td-cheer-text">${cheer}</div></div>`;
  }

  function bindTacCardClicks() {
    const modal = document.getElementById("modal-pg-full");
    const body = modal?.querySelector("#pg-full-modal-body");
    if (!body) return;
    const rgn = state.progressRegion || state.filter.region || "";
    body.querySelectorAll(".tac-card[data-teamcard]").forEach(tc => {
      tc.addEventListener("click", e => {
        e.stopPropagation();
        const teamName = tc.dataset.teamcard;
        const group = (state._hrankGroupData || {})[teamName];
        if (!group) return;
        const plan = getProgressAwardConfig(rgn).plan;
        openPgFullModal({
          title: `🏅 ${escapeHtml(teamName)}팀 상세`,
          subtitle: `달성률 ${group.rate.toFixed(1)}%`,
          bodyHTML: buildTeamDetailHTML(group, plan),
          pushStack: true
        });
      });
    });
  }

  function renderProgressHome(list) {
    const stats = list.map(getProgressStat);
    const total = stats.length;
    const avgR = stats.reduce((a, s) => a + s.rate, 0) / total;
    const over5 = stats.filter((s) => s.net >= 50000).length;
    const _pgCohort = (state.filter.cohort || state.progressCohort || "").replace(/기$/, "");
    const _pgStep   = state.filter.step || state.progressStep || "1";
    const _pa = getProgressAwardConfig(state.progressRegion, _pgCohort, _pgStep);
    const elig = stats.filter((s) => s.net >= 300000); // 순증 30만원 이상
    const byRate = [...stats].sort((a, b) => (b.net / (b.base || 1)) - (a.net / (a.base || 1)));
    const byAmt  = [...stats].sort((a, b) => b.net - a.net);
    const byIpum = [...stats].filter((s) => s.ipumAmt > 0).sort((a, b) => b.ipumAmt - a.ipumAmt || b.ipumCount - a.ipumCount);

    // 중복시상: 신장률·신장액 양쪽 대상이 되는 경우 더 큰 시상만 지급 (표시는 양쪽 그대로)
    const rateFinalList = byRate;

    const a5 = stats.filter((s) => s.net >= 500000).length;
    const a4 = stats.filter((s) => s.net >= 300000 && s.net < 500000).length;

    // 시상안 체크박스 활성화 여부 — 비활성 섹션은 숨김
    const _rateEnabled  = !!_pa.rateConfig;
    const _amtEnabled   = !!_pa.amtConfig;

    // 그룹 시상 — team 필드가 설정된 학생이 있으면 team 기준, 아니면 branch(지점)
    const hasAnyTeam = stats.some((s) => (s.s.team || "").toString().trim());
    const groupKeyFn = hasAnyTeam
      ? ((s) => (s.s.team || "").toString().trim() || "(팀 미배정)")
      : ((s) => s.s.branch || "(미지정)");
    const groupMap = {};
    stats.forEach((st) => {
      const k = groupKeyFn(st);
      if (!groupMap[k]) groupMap[k] = { base: 0, current: 0, members: [], memberStats: [] };
      groupMap[k].base += st.base;
      groupMap[k].current += st.current;
      groupMap[k].members.push(st.s.name || "");
      groupMap[k].memberStats.push(st);
    });
    const groupRanking = Object.entries(groupMap).map(([name, g]) => ({
      name,
      rate: g.base > 0 ? (g.current / g.base) * 100 : 0,
      members: g.members,
      memberStats: g.memberStats,
      base: g.base,
      current: g.current
    })).sort((a, b) => b.rate - a.rate);
    const groupLabel = hasAnyTeam ? "팀별 인보험 순증" : "지점별 인보험 순증 (팀 미배정)";
    const _groupEnabled = groupRanking.length >= 2;

    // Compute two-pass duplicate-award assignments once; reuse for preview cards + full modals
    const _bothAsgn = computeBothAwardAssignments(byRate, byAmt, _pa);
    // 중복시상 없음 적용: "other" 대상자를 각 TOP 테이블에서 제외 (슬라이딩카드와 동일한 방식)
    const _rateFinalDedup = _pa.bothEnabled
      ? rateFinalList.filter(st => (_bothAsgn.rateAsgn.get(st.s.empNo)?.status ?? "none") !== "other")
      : rateFinalList;
    const _byAmtDedup = _pa.bothEnabled
      ? byAmt.filter(st => (_bothAsgn.amtAsgn.get(st.s.empNo)?.status ?? "none") !== "other")
      : byAmt;
    const _mkRatePrize = (st) => {
      const a = _bothAsgn.rateAsgn.get(st.s.empNo);
      if (!a || a.status === "ineligible") {
        const _nTxt = a?.needsNet > 0 ? `+${Nf(a.needsNet)}원 필요` : "기준미달";
        return { txt: _nTxt, cls: "pg-b-no" };
      }
      if (a.status === "other") return { txt: "💰 신장액 시상", cls: "pg-rank-swap-text" };
      if (a.status === "mine") {
        if (a.effectiveAmt > 0) return { txt: `시상 ${Math.round(a.effectiveAmt / 10000)}만원`, cls: "" };
        const v = (_pa.rateConfig?.payouts || [])[a.effectiveRank - 1];
        return { txt: v != null ? payoutLabel(v) : "물품", cls: "" };
      }
      return { txt: "-", cls: "" };
    };
    const _mkAmtPrize = (st) => {
      const a = _bothAsgn.amtAsgn.get(st.s.empNo);
      if (!a || a.status === "ineligible") {
        const _nTxt = a?.needsNet > 0 ? `+${Nf(a.needsNet)}원 필요` : "기준미달";
        return { txt: _nTxt, cls: "pg-b-no" };
      }
      if (a.status === "other") return { txt: "📈 신장률 시상", cls: "pg-rank-swap-text" };
      if (a.status === "mine") {
        if (a.effectiveAmt > 0) return { txt: `시상 ${Math.round(a.effectiveAmt / 10000)}만원`, cls: "" };
        const v = (_pa.amtConfig?.payouts || [])[a.effectiveRank - 1];
        return { txt: v != null ? payoutLabel(v) : "물품", cls: "" };
      }
      return { txt: "-", cls: "" };
    };
    const pcardRateTop3 = rateFinalList.slice(0, 3).map((st, i) => {
      const rate = st.rate || 0;
      const p = _mkRatePrize(st);
      return `<li class="pg-pcard-row" data-emp="${escapeHtml(st.s.empNo)}">
        <span class="pg-rb ${i===0?"r1":i===1?"r2":"r3"}">${i+1}</span>
        <div class="pg-pcard-content">
          <div class="pg-pcard-nm"><strong>${escapeHtml(st.s.name||"")}</strong> <span class="pg-pcard-val">${rate.toFixed(1)}%</span></div>
          <span class="pg-pcard-prize ${p.cls}">${p.txt}</span>
        </div>
      </li>`;
    }).join("");

    const pcardAmtTop3 = byAmt.slice(0, 3).map((st, i) => {
      const p = _mkAmtPrize(st);
      return `<li class="pg-pcard-row" data-emp="${escapeHtml(st.s.empNo)}">
        <span class="pg-rb ${i===0?"r1":i===1?"r2":"r3"}">${i+1}</span>
        <div class="pg-pcard-content">
          <div class="pg-pcard-nm"><strong>${escapeHtml(st.s.name||"")}</strong> <span class="pg-pcard-val">${st.net>=0?"+":""}${Nf(st.net)}원</span></div>
          <span class="pg-pcard-prize ${p.cls}">${p.txt}</span>
        </div>
      </li>`;
    }).join("");

    const pcardIpumTop3 = byIpum.slice(0, 3).map((st, i) => {
      const grade = ["인품의 황제", "인품의 제왕", "인품의 왕"][i] || "";
      return `<li class="pg-pcard-row" data-emp="${escapeHtml(st.s.empNo)}">
        <span class="pg-rb ${i===0?"r1":i===1?"r2":"r3"}">${i+1}</span>
        <div class="pg-pcard-content">
          <div class="pg-pcard-nm"><strong>${escapeHtml(st.s.name||"")}</strong> <span class="pg-pcard-val">${st.ipumCount ? st.ipumCount + "건 " : ""}${Nf(st.ipumAmt)}원</span></div>
          <span class="pg-pcard-prize pg-b-p">${grade}</span>
        </div>
      </li>`;
    }).join("") || `<li class="pg-pcard-empty">실적관리 탭에서 인품 데이터를 입력하세요</li>`;

    const pcardGroupTop3 = groupRanking.slice(0, 3).map((g, i) => {
      const short = g.members.slice(0, 3).join("·") + (g.members.length > 3 ? "…" : "");
      return `<li class="pg-pcard-row pg-pcard-group-row">
        <span class="pg-rb ${i===0?"r1":i===1?"r2":"r3"}">${i+1}</span>
        <div class="pg-pcard-content">
          <div class="pg-pcard-nm"><strong>${escapeHtml(g.name)}</strong><br><small>${escapeHtml(short)}</small></div>
          <span class="pg-pcard-prize pg-b-g">${g.rate.toFixed(1)}%</span>
        </div>
      </li>`;
    }).join("");

    // 풀스크린 모달에 띄울 데이터 캐시 — 전체 순위 노출, 중복시상 반영
    state._pgCardFullData = {
      rate: {
        title: `📈 신장률 전체 순위 (${rateFinalList.length}명)`,
        subtitle: "달성률 (현재실적 ÷ 기준실적)",
        bodyHTML: renderProgressRankFullBothAware(byRate, byAmt, _pa, "rate")
      },
      amt: {
        title: `💰 신장액 전체 순위 (${byAmt.length}명)`,
        subtitle: "순증 금액 절대값",
        bodyHTML: renderProgressRankFullBothAware(byRate, byAmt, _pa, "amt")
      },
      ipum: {
        title: `✨ 인품왕 전체 순위 (${byIpum.length}명)`,
        subtitle: "신상품 판매액 기준",
        bodyHTML: byIpum.length ? renderProgressTop10(byIpum, "ipum", Infinity, _pa) : `<div class="pg-empty">실적관리 탭에서 인품 데이터를 입력하세요.</div>`
      },
      group: {
        title: `🏅 ${hasAnyTeam ? "팀시상" : "그룹 순증"} (${groupRanking.length}${hasAnyTeam ? "팀" : "개 지점"})`,
        subtitle: groupLabel,
        bodyHTML: renderTeamAwardFull(groupRanking, _pa.plan, hasAnyTeam)
      }
    };
    state._hrankGroupData = {};
    groupRanking.forEach(g => { state._hrankGroupData[g.name] = g; });

    // 공유 모드 배너 — 공유 링크로 열었을 때 상단에 표시
    const _shareCenter = (() => {
      const cs = [...new Set(list.map(s => s.center).filter(Boolean))];
      return cs.length === 1 ? cs[0] : (cs[0] || "");
    })();
    const _shareBannerHtml = state.pgShareMode ? `
      <div class="pg-share-banner">
        <h2>고객컨설팅 마스터 ${escapeHtml(state.progressRegion)}${_pgCohort ? ` ${_pgCohort}기` : ""}${_shareCenter ? ` ${escapeHtml(_shareCenter)}` : ""} Step${escapeHtml(_pgStep)}과정 실적진도</h2>
        <span class="pg-share-badge">공유 보기</span>
      </div>` : "";

    // 시상안 박스 + KPI 영역
    const _plan = _pa.plan;
    let _tiersHtml = "";
    if (_plan.personalIncr?.enabled) {
      const _items = (_plan.personalIncr.items || []).slice().sort((a, b) => Number(a.critVal) - Number(b.critVal));
      _tiersHtml = _items.map((it) => {
        const cond = `순증 ${escapeHtml(String(it.critVal))}만원↑`;
        const pay = it.payType === "pct" ? `${escapeHtml(String(it.payVal))}%` : `${escapeHtml(String(it.payVal))}만원`;
        return `<div class="pg-tier"><div class="pg-tc">${cond}</div><div class="pg-tn">${pay}</div></div>`;
      }).join("");
    }
    const _topToHtml = (top) => {
      if (!top?.enabled) return "";
      const typeLabel = top.type === "rate" ? "신장률" : "신장액";
      const icon = top.type === "rate" ? "📈" : "💰";
      const payouts = (top.payouts || []).slice(0, Number(top.n));
      const items = payouts.map((p, i) => `${i + 1}위 ${escapeHtml(payoutLabel(p))}`).join(" / ");
      return `<div class="pg-an"><strong>${icon} ${typeLabel} TOP${escapeHtml(String(top.n))}:</strong> ${items}</div>`;
    };
    const _eligText = (() => {
      if (!_plan.eligibility?.enabled) return "";
      const conds = _plan.eligibility.conditions || [];
      if (!conds.length) return "";
      const fLabel = (f) => ({ converted: "환산실적", hiCap: "하이캡", monthly: "월납보험료" }[f] || f);
      const fUnit  = (f) => f === "hiCap" ? "" : "만원";
      const op = _plan.eligibility.operator === "or" ? " 또는 " : " 그리고 ";
      const txt = conds.map((c) => `${escapeHtml(fLabel(c.field))} ${escapeHtml(String(c.threshold))}${fUnit(c.field)} 이하`).join(op);
      return `<div class="pg-an-crit">⚠️ 시상 제외 조건: ${txt} → 시상 제외</div>`;
    })();
    // 인쇄용 데이터 캐시 (print 버튼 클릭 시 openProgressPrintWindow 가 사용)
    state.pgPrintData = { byRate, byAmt, _pa, groupRanking, plan: _plan, stats,
      region: state.progressRegion, cohort: _pgCohort, step: _pgStep,
      hasAnyTeam, _rateFinalDedup, _byAmtDedup };
    return `
      <div class="pg-wrap">
        ${_shareBannerHtml}
        <div class="pg-award-box">
          <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:6px;">
            <h3 style="margin:0">📋 ${escapeHtml(state.progressRegion)}${_pgCohort ? ` ${_pgCohort}기` : ""} Step ${_pgStep} 활성화 시상안</h3>
            ${!state.pgShareMode ? `<button id="btn-pg-share" class="btn-outline pg-share-btn">🔗 링크 공유</button><button id="btn-pg-print" class="btn-outline pg-share-btn" style="background:#fff8e1;border-color:#f9a825;color:#f57f17;">🖨️ PDF</button>` : ""}
          </div>
          ${_plan.title ? `<div class="pg-award-subtitle">${escapeHtml(_plan.title)}</div>` : ""}
          ${_tiersHtml ? `<div class="pg-tier-row">${_tiersHtml}</div>` : ""}
          ${_topToHtml(_plan.topAward1)}
          ${_topToHtml(_plan.topAward2)}
          ${_eligText}
          ${_pa.bothEnabled ? `<div class="pg-an-warn">※ 신장률·신장액 TOP 중복시상 없음 — 두 항목 모두 해당 시 더 큰 시상 1회만 지급</div>` : ""}
          ${_plan.notes ? `<div class="pg-an-warn">${escapeHtml(_plan.notes)}</div>` : ""}
        </div>

        <div class="pg-kpi-card">
          <div class="pg-kpi"><div class="pgl">총 인원</div><div class="pgv">${total}명</div></div>
          <div class="pg-kpi"><div class="pgl">평균 달성률</div><div class="pgv">${avgR.toFixed(1)}%</div></div>
          <div class="pg-kpi g"><div class="pgl">5만원↑ 순증</div><div class="pgv">${over5}명</div></div>
          <div class="pg-kpi gd"><div class="pgl">순증 30만↑</div><div class="pgv">${elig.length}명</div></div>
          <div class="pg-kpi g"><div class="pgl">150% 지급</div><div class="pgv">${a5}명</div></div>
          <div class="pg-kpi or"><div class="pgl">120% 지급</div><div class="pgv">${a4}명</div></div>
        </div>

        <!-- 모바일 전용 TOP3 미리보기 카드 (2x2) — ≤640px 에서만 노출 -->
        <div class="pg-mobile-grid">
          ${_rateEnabled ? `<div class="pg-pcard pg-pcard-rate" data-pcard="rate" role="button" tabindex="0">
            <div class="pg-pcard-head">
              <div class="pg-pcard-icon">📈</div>
              <div class="pg-pcard-titles">
                <h5>최고 신장률 TOP10</h5>
                <p>달성률 (현재실적 ÷ 기준실적)</p>
              </div>
              <span class="pg-pcard-chev">›</span>
            </div>
            <ol class="pg-top3-list">${pcardRateTop3}</ol>
          </div>` : ""}

          ${_amtEnabled ? `<div class="pg-pcard pg-pcard-amt" data-pcard="amt" role="button" tabindex="0">
            <div class="pg-pcard-head">
              <div class="pg-pcard-icon">💰</div>
              <div class="pg-pcard-titles">
                <h5>최고 신장액 TOP10</h5>
                <p>순증 금액 절대값</p>
              </div>
              <span class="pg-pcard-chev">›</span>
            </div>
            <ol class="pg-top3-list">${pcardAmtTop3}</ol>
          </div>` : ""}

          <div class="pg-pcard pg-pcard-ipum" data-pcard="ipum" role="button" tabindex="0">
            <div class="pg-pcard-head">
              <div class="pg-pcard-icon">✨</div>
              <div class="pg-pcard-titles">
                <h5>인품왕 TOP10</h5>
                <p>신상품 판매액</p>
              </div>
              <span class="pg-pcard-chev">›</span>
            </div>
            <ol class="pg-top3-list">${pcardIpumTop3}</ol>
          </div>

          ${_groupEnabled ? `<div class="pg-pcard pg-pcard-group" data-pcard="group" role="button" tabindex="0">
            <div class="pg-pcard-head">
              <div class="pg-pcard-icon">🏅</div>
              <div class="pg-pcard-titles">
                <h5>그룹 순증 시상</h5>
                <p>${escapeHtml(groupLabel)}</p>
              </div>
              <span class="pg-pcard-chev">›</span>
            </div>
            <ol class="pg-top3-list">${pcardGroupTop3}</ol>
          </div>` : ""}
        </div>

        ${(_rateEnabled || _amtEnabled) ? `<div class="pg-grid2 pg-desktop-only" style="${!_rateEnabled || !_amtEnabled ? "grid-template-columns:1fr" : ""}">
          ${_rateEnabled ? `<div class="pg-card">
            <h4>📈 신장률 TOP${_pa.rateConfig?.n || ""}${_pa.bothEnabled ? ` <small>(중복시상 시 더 큰 시상만 지급)</small>` : ""}${_pa.rateConfig?.minNetEnabled ? ` <small class="pg-top-crit-badge">순증 ${Nf(Number(_pa.rateConfig.minNet||0))}원↑ 기준</small>` : ""}</h4>
            ${renderProgressTop10(_rateFinalDedup, "rate", undefined, _pa)}
          </div>` : ""}
          ${_amtEnabled ? `<div class="pg-card">
            <h4>💰 신장액 TOP${_pa.amtConfig?.n || ""}${_pa.amtConfig?.minNetEnabled ? ` <small class="pg-top-crit-badge">순증 ${Nf(Number(_pa.amtConfig.minNet||0))}원↑ 기준</small>` : ""}</h4>
            ${renderProgressTop10(_byAmtDedup, "amt", undefined, _pa)}
          </div>` : ""}
        </div>` : ""}

        <div class="pg-grid-ipum-group pg-desktop-only" style="${!_groupEnabled ? "grid-template-columns:1fr" : ""}">
          <div class="pg-card">
            <h4>✨ 인품왕 TOP10 <small>(신상품 판매액 기준)</small></h4>
            ${byIpum.length ? renderProgressTop10(byIpum, "ipum", undefined, _pa) : `<div class="pg-empty">실적관리 탭에서 인품 데이터를 입력하세요.</div>`}
          </div>
          ${_groupEnabled ? `<div class="pg-card">
            <h4>🏅 ${hasAnyTeam ? "조별" : "지점별"} 순증 시상</h4>
            <div class="pg-tbl-wrap">${renderGroupTable(groupRanking, _pa.plan, hasAnyTeam)}</div>
          </div>` : ""}
        </div>

        <div class="pg-card pg-full-tbl-card" id="pg-full-tbl-card">
          <h4 style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">📊 전체 교육생 실적표 <small>(신장액 내림차순, 클릭 시 상세)</small>
            <button type="button" class="pg-full-tbl-toggle btn-outline small" id="btn-pg-full-toggle">펼쳐보기</button>
            <button type="button" class="btn-primary small" id="btn-pg-tbl-save" hidden style="margin-left:auto;font-size:12px;padding:4px 10px;">💾 변경 저장</button>
          </h4>
          <div class="pg-tbl-wrap"><table class="pg-tbl pg-full-rank-tbl">
            <thead><tr><th style="width:44px">순위</th><th>성명</th><th style="width:76px">사번</th><th>지점</th><th class="r" style="width:68px">기준실적</th><th class="r" style="width:68px">현재실적</th><th class="r" style="width:96px">마스터목표</th><th style="width:52px">마스터<br>달성률</th><th class="r" style="width:74px">마스터<br>순증</th><th style="width:52px">기준<br>달성률</th><th class="r" style="width:74px">기준<br>순증</th><th>전화번호</th><th>시상</th></tr></thead>
            <tbody>${byAmt.map((st, i) => {
              const masterGoal = Number(st.s.target) > 0 ? Number(st.s.target) : st.base;
              const masterRate = masterGoal > 0 ? (st.current / masterGoal * 100) : 0;
              const masterNet  = st.current - masterGoal;
              const baseRate   = st.base > 0 ? (st.current / st.base * 100) : 0;
              const baseNet    = st.current - st.base;
              const nc  = masterRate >= 120 ? "pg-c-over" : masterRate >= 100 ? "pg-c-good" : masterRate >= 80 ? "pg-c-mid" : "pg-c-low";
              const bc  = baseRate   >= 120 ? "pg-c-over" : baseRate   >= 100 ? "pg-c-good" : baseRate   >= 80 ? "pg-c-mid" : "pg-c-low";
              const netC  = masterNet > 0 ? "pg-net-p" : masterNet < 0 ? "pg-net-m" : "";
              const bnetC = baseNet   > 0 ? "pg-net-p" : baseNet   < 0 ? "pg-net-m" : "";
              const aw = tierLabel(st.net, undefined, _pa);
              const baseDisp     = st.base > 0    ? Nf(st.base)                    : "—";
              const goalDisp     = masterGoal > 0 ? Nf(masterGoal)                 : "—";
              const rateDisp     = masterGoal > 0 ? masterRate.toFixed(1) + "%"    : "—";
              const baseRateDisp = st.base > 0    ? baseRate.toFixed(1) + "%"      : "—";
              const phone = st.s.phone ? `<a href="tel:${escapeHtml(st.s.phone)}" class="pg-tel-link" onclick="event.stopPropagation()">${escapeHtml(st.s.phone)}</a>` : "—";
              return `<tr data-emp="${escapeHtml(st.s.empNo)}" class="pg-tr-click"><td>${RB(i + 1)}</td><td><strong>${escapeHtml(st.s.name || "")}</strong></td><td class="pg-empno-cell">${escapeHtml(st.s.empNo || "")}</td><td>${escapeHtml(st.s.branch || "")}</td><td class="r">${baseDisp}</td><td class="r">${Nf(st.current)}</td><td class="r pg-goal-cell" data-emp="${escapeHtml(st.s.empNo)}" data-goal="${masterGoal}" data-base="${st.base}"><span class="pg-goal-val">${goalDisp}</span><button type="button" class="pg-goal-adj-btn" data-emp="${escapeHtml(st.s.empNo)}" data-goal="${masterGoal}" data-base="${st.base}" onclick="event.stopPropagation()">±</button></td><td class="${nc}">${rateDisp}</td><td class="r ${netC}">${masterNet >= 0 ? "+" : ""}${Nf(masterNet)}</td><td class="${bc}">${baseRateDisp}</td><td class="r ${bnetC}">${baseNet >= 0 ? "+" : ""}${Nf(baseNet)}</td><td class="pg-tel-cell">${phone}</td><td>${aw}</td></tr>`;
            }).join("")}</tbody>
          </table></div>
        </div>
      </div>
    `;
  }

  function renderHomeRanks() {
    const section = document.getElementById("home-ranks");
    if (!section) return;

    const region = state.filter.region;
    const mainKpi = document.getElementById("main-kpi-grid");
    if (!region) {
      section.hidden = true;
      if (mainKpi) mainKpi.style.display = "";
      if (state._hrankTimer) { clearTimeout(state._hrankTimer); state._hrankTimer = null; }
      const kpiInline = document.getElementById("kpi-inline-stats");
      if (kpiInline) { kpiInline.innerHTML = ""; kpiInline.setAttribute("hidden", ""); }
      return;
    }
    section.hidden = false;
    if (mainKpi) mainKpi.style.display = "none";

    const cohort = state.filter.cohort || "";
    const step = state.filter.step || state.progressStep || "1";

    // 같은 지역단+기수+스텝으로 캐러셀이 이미 동작 중이면 재렌더 생략 (타이머 리셋 방지)
    if (state._hrankTimer && section.dataset.hrRegion === region && section.dataset.hrCohort === cohort && section.dataset.hrStep === step && section.querySelector(".hr-slide")) return;
    if (state._hrankTimer) { clearTimeout(state._hrankTimer); state._hrankTimer = null; }
    section.dataset.hrRegion = region;
    section.dataset.hrCohort = cohort;
    section.dataset.hrStep = step;

    const _cohortNum = cohort.replace(/기$/, "");
    const list = state.students.filter((s) => {
      if (s.region !== region) return false;
      if (_cohortNum && s.cohort && String(s.cohort).replace(/기$/, "") !== _cohortNum) return false;
      return true;
    });
    if (!list.length) { section.innerHTML = ""; return; }

    const stats = list.map(getProgressStat);
    const byAmt = [...stats].sort((a, b) => b.net - a.net);
    const byRateRaw = [...stats].sort((a, b) => (b.net / (b.base || 1)) - (a.net / (a.base || 1)));
    const byIpum = [...stats].filter((s) => s.ipumAmt > 0).sort((a, b) => b.ipumAmt - a.ipumAmt || b.ipumCount - a.ipumCount);
    // 홈 필터의 cohort/step 으로 시상안 조회 (state.progressCohort/Step 과 다를 수 있음)
    const _hrankPa = getProgressAwardConfig(region, cohort, step);
    // 중복시상 정책: 양쪽 대상 중 더 큰 시상만 지급 — 중복 대상자는 해당 카테고리 순위에서 제외
    const { rateAsgn: _hrRateAsgn, amtAsgn: _hrAmtAsgn } = computeBothAwardAssignments(byRateRaw, byAmt, _hrankPa);
    const byRate    = byRateRaw.filter(st => (_hrRateAsgn.get(st.s.empNo)?.status ?? "none") !== "other");
    const byAmtDedup = byAmt.filter(st   => (_hrAmtAsgn.get(st.s.empNo)?.status  ?? "none") !== "other");
    const _hrAwardPlan = _hrankPa.plan; // 정규화된 플랜 재사용

    const hasAnyTeam = stats.some((s) => (s.s.team || "").toString().trim());
    const groupKeyFn = hasAnyTeam
      ? ((s) => (s.s.team || "").toString().trim() || "(팀 미배정)")
      : ((s) => s.s.branch || "(미지정)");
    const groupMap = {};
    stats.forEach((st) => {
      const k = groupKeyFn(st);
      if (!groupMap[k]) groupMap[k] = { base: 0, current: 0, members: [], memberStats: [] };
      groupMap[k].base += st.base;
      groupMap[k].current += st.current;
      groupMap[k].members.push(st.s.name || "");
      groupMap[k].memberStats.push(st);
    });
    const groupRanking = Object.entries(groupMap).map(([k, g]) => ({
      name: k, rate: g.base > 0 ? (g.current / g.base) * 100 : 0, members: g.members, memberStats: g.memberStats, base: g.base, current: g.current
    })).sort((a, b) => b.rate - a.rate);

    const RNKS = ["r1", "r2", "r3"];

    function hrListRows(arr, isGroup, valFn, limit) {
      const top = arr.slice(0, limit);
      if (!top.length) return `<li class="hr-empty">데이터 없음</li>`;
      return top.map((item, i) => {
        const empNo = isGroup ? "" : item.s.empNo;
        const name = isGroup ? item.name : (item.s.name || "");
        return `<li class="hr-row${empNo ? " hr-clickable" : ""}" data-emp="${escapeHtml(empNo || "")}">
          <span class="pg-rb ${i < 3 ? RNKS[i] : "rt"}">${i + 1}</span>
          <span class="hr-name">${escapeHtml(name)}</span>
          <span class="hr-val">${escapeHtml(valFn(item))}</span>
        </li>`;
      }).join("");
    }

    const _hrRateEnabled  = !!_hrankPa.rateConfig;
    const _hrAmtEnabled   = !!_hrankPa.amtConfig;
    const _hrGroupEnabled = groupRanking.length >= 2;

    const hrCats = [
      { key: "rate", icon: "📈", title: "최고 신장률", sub: "달성률 (현재실적 ÷ 기준실적)", cls: "hr-rate",
        arr: byRate, worstArr: byRateRaw, isGroup: false,
        valFn: (st) => `${st.rate != null ? st.rate.toFixed(1) : "0.0"}%` },
      { key: "amt", icon: "💰", title: "최고 신장액", sub: "순증 금액 절대값", cls: "hr-amt",
        arr: byAmtDedup, worstArr: byAmt, isGroup: false,
        valFn: (st) => `${st.net >= 0 ? "+" : ""}${Nf(st.net)}원` },
      { key: "ipum", icon: "✨", title: "인품왕", sub: "신상품 판매액", cls: "hr-ipum",
        arr: byIpum, isGroup: false,
        valFn: (st) => `${st.ipumCount ? st.ipumCount + "건 " : ""}${Nf(st.ipumAmt)}원` },
      { key: "group", icon: "🏅", title: hasAnyTeam ? "팀시상" : "그룹 순증", sub: hasAnyTeam ? "팀별 달성률" : "팀구분이 없습니다", cls: "hr-group",
        arr: groupRanking, isGroup: true,
        valFn: (g) => `${g.rate.toFixed(1)}%` }
    ].filter(cat => {
      if (cat.key === "rate")  return _hrRateEnabled;
      if (cat.key === "amt")   return _hrAmtEnabled;
      if (cat.key === "group") return _hrGroupEnabled;
      return true;
    });

    state._hrankGroupData = {};
    groupRanking.forEach(g => { state._hrankGroupData[g.name] = g; });

    const _closeBtnBar = `<div class="pg-body-close-bar"><button class="btn-outline small" data-pg-close>✕ 닫기</button></div>`;
    state._hrankFullData = {};
    hrCats.forEach(cat => {
      if (cat.key === "group") {
        state._hrankFullData[cat.key] = {
          title: `🏅 ${hasAnyTeam ? "팀시상" : "그룹 순증"} (${groupRanking.length}${hasAnyTeam ? "팀" : "개 지점"})`,
          subtitle: cat.sub,
          bodyHTML: _closeBtnBar + renderTeamAwardFull(groupRanking, _hrAwardPlan, hasAnyTeam)
        };
      } else {
        const ttls = { rate: `📈 신장률 전체 순위 (${byRateRaw.length}명)`, amt: `💰 신장액 전체 순위 (${byAmt.length}명)`, ipum: `✨ 인품왕 전체 순위 (${byIpum.length}명)` };
        const subs = { rate: "달성률 (현재실적 ÷ 기준실적)", amt: "순증 금액 절대값", ipum: "신상품 판매액" };
        let body;
        if (cat.key === "rate") {
          body = byRateRaw.length ? renderProgressRankFullBothAware(byRateRaw, byAmt, _hrankPa, "rate") : `<div class="pg-empty">데이터 없음</div>`;
        } else if (cat.key === "amt") {
          body = byAmt.length ? renderProgressRankFullBothAware(byRateRaw, byAmt, _hrankPa, "amt") : `<div class="pg-empty">데이터 없음</div>`;
        } else {
          body = cat.arr.length ? renderProgressTop10(cat.arr, cat.key, Infinity, _hrankPa) : `<div class="pg-empty">데이터 없음</div>`;
        }
        state._hrankFullData[cat.key] = {
          title: ttls[cat.key], subtitle: subs[cat.key],
          bodyHTML: _closeBtnBar + body
        };
      }
    });

    // 교육생 통계 슬라이드용 KPI 계산 — stats는 getProgressStat() 통과, 선택된 스텝 기준
    const kpiTotal   = stats.length;
    const kpiBase    = stats.reduce((a, st) => a + st.base, 0);
    const kpiCurrent = stats.reduce((a, st) => a + st.current, 0);
    const kpiRate    = kpiBase > 0 ? (kpiCurrent / kpiBase * 100).toFixed(1) : "0.0";
    // KPI를 슬라이드 대신 헤더 desc 줄에 인라인 표시
    const kpiInlineEl = document.getElementById("kpi-inline-stats");
    if (kpiInlineEl) {
      kpiInlineEl.removeAttribute("hidden");
      kpiInlineEl.innerHTML = `
        <div class="kpi-card">
          <div class="kpi-label">전체 교육생</div>
          <div class="kpi-value">${kpiTotal}<span class="kpi-unit">명</span></div>
        </div>
        <div class="kpi-card">
          <div class="kpi-label">기준실적(A) 합계</div>
          <div class="kpi-value">${kpiBase.toLocaleString()}<span class="kpi-unit">원</span></div>
        </div>
        <div class="kpi-card">
          <div class="kpi-label">현재실적(B) 합계</div>
          <div class="kpi-value">${kpiCurrent.toLocaleString()}<span class="kpi-unit">원</span></div>
        </div>
        <div class="kpi-card highlight">
          <div class="kpi-label">달성률</div>
          <div class="kpi-value">${kpiRate}<span class="kpi-unit">%</span></div>
        </div>
      `;
    }

    const worstTitles = { rate: "최저 신장률", amt: "최저 신장액", ipum: "인품 하위", group: hasAnyTeam ? "팀시상 하위" : "그룹 하위" };

    const cmpFmt = v => { const n = Number(v) || 0; return Math.abs(n) >= 10000 ? (n / 10000).toFixed(1) + '만' : String(n); };

    const _hrPlan = _hrAwardPlan;
    const _hrGa1Thr = _hrPlan?.groupAward1?.enabled
      ? Number(_hrPlan.groupAward1.threshold || 5) * 10000 : 0;
    // 카테고리별 시상 payouts 매핑
    const _hrAwardMap = {};
    ["topAward1", "topAward2"].forEach(k => {
      const t = _hrPlan?.[k];
      if (t?.enabled && t.payouts?.length) _hrAwardMap[t.type] = t.payouts;
    });

    function hrMakeSlide(isWorst) {
      const hidden = isWorst ? " hr-slide-hidden" : "";
      return `<div class="hr-slide${hidden}"><div class="hr-grid">${
        hrCats.map(cat => {
          const dispArr = isWorst ? [...(cat.worstArr || cat.arr)].reverse().slice(0, 3) : cat.arr.slice(0, 3);
          const catAwards = (!isWorst && _hrAwardMap[cat.key]) ? _hrAwardMap[cat.key] : null;
          const title = isWorst ? (worstTitles[cat.key] || cat.title) : cat.title;
          const rows = dispArr.length
            ? dispArr.map((item, i) => {
                const empNo = cat.isGroup ? "" : item.s.empNo;
                const name = cat.isGroup ? item.name : (item.s.name || "");
                const rbCls = isWorst ? "hr-rb-worst" : (i < 3 ? RNKS[i] : "rt");
                let nameBlock;
                if (cat.isGroup) {
                  const base_k = Math.round(item.base / 1000);
                  const cur_k  = Math.round(item.current / 1000);
                  const statsLine = `목표 ${Nf(base_k)}천 · 현재 ${Nf(cur_k)}천 · ${item.rate.toFixed(1)}%`;
                  const teamSuffix = hasAnyTeam ? "팀" : "";
                  const memberSpans = (item.memberStats || []).map(st => {
                    const mNet = st.net || 0;
                    const missed = _hrGa1Thr > 0 ? mNet < _hrGa1Thr : mNet < 0;
                    const mEmp = st.s?.empNo || "";
                    const mName = st.s?.name || "";
                    return `<span class="hr-member${missed ? " hr-member-miss" : ""}${mEmp ? " hr-member-click" : ""}" data-emp="${escapeHtml(mEmp)}">${escapeHtml(mName)}</span>`;
                  }).join("");
                  nameBlock = `<div class="hr-name-wrap">
                    <span class="hr-name">${escapeHtml(name)}${teamSuffix}</span>
                    <span class="hr-row-stats">${statsLine}</span>
                    ${memberSpans ? `<span class="hr-members">${memberSpans}</span>` : ""}
                  </div>`;
                } else {
                  const pct = item.rate != null ? item.rate.toFixed(1) : '—';
                  const statsLine = `기준 ${cmpFmt(item.base)} · 현재 ${cmpFmt(item.current)} · 순증 ${item.net >= 0 ? '+' : ''}${cmpFmt(item.net)} (${pct}%)`;
                  nameBlock = `<div class="hr-name-wrap"><span class="hr-name">${escapeHtml(name)}</span><span class="hr-row-stats">${statsLine}</span></div>`;
                }
                // 순증 기준 미달 시 "필요" 배지 우선 표시
                let awardBadge = "";
                if (!isWorst && !cat.isGroup && (cat.key === "rate" || cat.key === "amt")) {
                  const _hrAsgnMap = cat.key === "rate" ? _hrRateAsgn : _hrAmtAsgn;
                  const _hrAsgn = empNo ? _hrAsgnMap.get(empNo) : null;
                  if (_hrAsgn?.status === "ineligible" && _hrAsgn.needsNet > 0) {
                    awardBadge = `<span class="hr-award-badge hr-award-need">+${Nf(_hrAsgn.needsNet)}원 필요</span>`;
                  } else if (catAwards?.[i] !== undefined) {
                    awardBadge = `<span class="hr-award-badge">${escapeHtml(payoutLabel(catAwards[i]))}</span>`;
                  }
                } else if (catAwards?.[i] !== undefined) {
                  awardBadge = `<span class="hr-award-badge">${escapeHtml(payoutLabel(catAwards[i]))}</span>`;
                }
                return `<li class="hr-row${empNo ? " hr-clickable" : ""}" data-emp="${escapeHtml(empNo || "")}">
                  <span class="pg-rb ${rbCls}">${i + 1}</span>
                  ${nameBlock}
                  <span class="hr-val">${cat.isGroup ? item.rate.toFixed(1) + "%" : escapeHtml(cat.valFn(item))}</span>
                  ${awardBadge}
                </li>`;
              }).join("")
            : `<li class="hr-empty">데이터 없음</li>`;
          return `
            <div class="hr-card ${cat.cls}${isWorst ? " hr-card-worst" : ""}" data-hrpcard="${cat.key}">
              <div class="hr-head">
                <span class="hr-icon" aria-hidden="true">${cat.icon}</span>
                <div class="hr-titles">
                  <strong>${title}</strong>
                  <span class="hr-sub">${cat.sub}</span>
                </div>
                <span class="hr-chev">›</span>
              </div>
              <ol class="hr-list">${rows}</ol>
            </div>`;
        }).join("")
      }</div></div>`;
    }

    section.innerHTML = `
      <div class="hr-slides">
        ${hrMakeSlide(false)}
        ${hrMakeSlide(true)}
      </div>
      <div class="hr-dots">
        <span class="hr-dot hr-dot-active"></span>
        <span class="hr-dot"></span>
      </div>
    `;

    function hrBindClicks() {
      section.querySelectorAll(".hr-clickable[data-emp]").forEach(el => {
        el.addEventListener("click", e => {
          e.stopPropagation();
          if (state.pgShareMode) return;
          if (el.dataset.emp) openProgressStudentPopup(el.dataset.emp);
        });
      });
      section.querySelectorAll(".hr-member-click[data-emp]").forEach(el => {
        el.addEventListener("click", e => {
          e.stopPropagation();
          if (state.pgShareMode) return;
          if (el.dataset.emp) openProgressStudentPopup(el.dataset.emp);
        });
      });
      section.querySelectorAll(".hr-card[data-hrpcard]").forEach(card => {
        card.addEventListener("click", e => {
          if (e.target.closest(".hr-clickable, .hr-member-click")) return;
          const data = state._hrankFullData && state._hrankFullData[card.dataset.hrpcard];
          if (data) {
            openPgFullModal(data);
            if (card.dataset.hrpcard === "group") bindTacCardClicks();
          }
        });
      });
    }
    hrBindClicks();

    function hrGoTo(idx) {
      const slides = section.querySelectorAll(".hr-slide");
      slides.forEach((s, i) => {
        if (i === idx) {
          s.classList.remove("hr-slide-hidden", "hr-slide-enter");
          void s.offsetWidth;
          s.classList.add("hr-slide-enter");
        } else {
          s.classList.add("hr-slide-hidden");
          s.classList.remove("hr-slide-enter");
        }
      });
      section.querySelectorAll(".hr-dot").forEach((d, i) => d.classList.toggle("hr-dot-active", i === idx));
    }

    const hrDurations = [5000, 10000];
    let hrSlide = 0;
    function hrStep() {
      if (!document.getElementById("home-ranks")) { state._hrankTimer = null; return; }
      hrSlide = (hrSlide + 1) % 2;
      hrGoTo(hrSlide);
      state._hrankTimer = setTimeout(hrStep, hrDurations[hrSlide]);
    }
    state._hrankTimer = setTimeout(hrStep, hrDurations[0]);

    // 터치 스와이프 지원
    const slidesEl = section.querySelector(".hr-slides");
    if (slidesEl) {
      let _tx = 0;
      slidesEl.addEventListener("touchstart", (e) => { _tx = e.touches[0].clientX; }, { passive: true });
      slidesEl.addEventListener("touchend", (e) => {
        const dx = e.changedTouches[0].clientX - _tx;
        if (Math.abs(dx) < 40) return;
        const total = section.querySelectorAll(".hr-slide").length;
        hrSlide = dx < 0 ? (hrSlide + 1) % total : (hrSlide - 1 + total) % total;
        hrGoTo(hrSlide);
        if (state._hrankTimer) clearTimeout(state._hrankTimer);
        state._hrankTimer = setTimeout(hrStep, hrDurations[hrSlide]);
      }, { passive: true });
    }
  }

  function renderProgressTop10(list, kind, limit, pa) {
    const region = state.progressRegion;
    if (!pa) pa = getProgressAwardConfig(region);
    let max;
    if (limit === undefined || limit === null) {
      if (kind === "rate") max = pa.rateConfig?.n || 10;
      else if (kind === "amt") max = pa.amtConfig?.n || 10;
      else max = 10;
    } else {
      max = limit;
    }
    const top = (max === Infinity || max <= 0) ? list.slice() : list.slice(0, max);
    if (!top.length) return `<div class="pg-empty">데이터 없음</div>`;
    // 지급 기준 문구 (신장률/신장액에만 표시)
    const _rateMinNet = pa.rateConfig?.minNetEnabled ? Number(pa.rateConfig.minNet || 0) : 0;
    const _amtMinNet  = pa.amtConfig?.minNetEnabled  ? Number(pa.amtConfig.minNet  || 0) : 0;
    const _criteriaHtml = (() => {
      if (kind === "rate" && _rateMinNet > 0) {
        return `<div class="pg-top-criteria">지급 기준: 순증 <strong>${Nf(_rateMinNet)}원</strong> 이상</div>`;
      }
      if (kind === "amt" && _amtMinNet > 0) {
        return `<div class="pg-top-criteria">지급 기준: 순증 <strong>${Nf(_amtMinNet)}원</strong> 이상</div>`;
      }
      return "";
    })();
    return `${_criteriaHtml}
      <table class="pg-tbl">
        <thead><tr><th>#</th><th>성명</th><th>지점</th>${kind === "ipum" ? "<th>인품실적</th>" : "<th>달성률</th><th>순증</th>"}<th>시상</th></tr></thead>
        <tbody>${top.map((st, i) => {
          let value, value2, prize;
          if (kind === "ipum") {
            value = (st.ipumCount ? st.ipumCount + "건 " : "") + Nf(st.ipumAmt) + "원";
            value2 = null;
            const grade = ["인품의 황제", "인품의 제왕", "인품의 왕"][i];
            prize = grade ? `<span class="pg-bdg pg-b-p">${grade}</span>` : "-";
          } else if (kind === "rate") {
            value = (st.rate || 0).toFixed(1) + "%";
            value2 = (st.net >= 0 ? "+" : "") + Nf(st.net) + "원";
            const _rNetMiss = _rateMinNet > 0 && st.net < _rateMinNet;
            const _chkTop10 = pa.isTopEligible || pa.isEligible;
            if (!_chkTop10(st.s)) {
              prize = `<span class="pg-bdg pg-b-no">기준미달</span>`;
            } else if (_rNetMiss) {
              prize = `<span class="pg-bdg pg-b-no">+${Nf(_rateMinNet - st.net)}원 필요</span>`;
            } else {
              const v = (pa.rateConfig?.payouts || [])[i];
              const lbl = v != null ? payoutLabel(v) : (pa.rateTop10[i] > 0 ? `${Math.round(pa.rateTop10[i]/10000)}만원` : null);
              prize = lbl ? `<span class="pg-bdg pg-b-g">${escapeHtml(lbl)}</span>` : "-";
            }
          } else {
            value = (st.net >= 0 ? "+" : "") + Nf(st.net) + "원";
            value2 = (st.rate || 0).toFixed(1) + "%";
            const _aNetMiss = _amtMinNet > 0 && st.net < _amtMinNet;
            const _chkTop10 = pa.isTopEligible || pa.isEligible;
            if (!_chkTop10(st.s)) {
              prize = `<span class="pg-bdg pg-b-no">기준미달</span>`;
            } else if (_aNetMiss) {
              prize = `<span class="pg-bdg pg-b-no">+${Nf(_amtMinNet - st.net)}원 필요</span>`;
            } else {
              const v = (pa.amtConfig?.payouts || [])[i];
              const lbl = v != null ? payoutLabel(v) : (pa.amtTop10[i] > 0 ? `${Math.round(pa.amtTop10[i]/10000)}만원` : null);
              prize = lbl ? `<span class="pg-bdg pg-b-g">${escapeHtml(lbl)}</span>` : "-";
            }
          }
          const valueCells = value2 != null
            ? `<td class="r">${value}</td><td class="r">${value2}</td>`
            : `<td class="r">${value}</td>`;
          return `<tr class="pg-tr-click" data-emp="${escapeHtml(st.s.empNo)}"><td>${RB(i + 1)}</td><td><strong>${escapeHtml(st.s.name || "")}</strong></td><td>${escapeHtml(st.s.branch || "")}</td>${valueCells}<td>${prize}</td></tr>`;
        }).join("")}</tbody>
      </table>
    `;
  }

  // 중복시상 반영 전체 순위표 — 신장률·신장액 중 더 큰 시상만 표시
  // Two-pass duplicate-award assignment:
  // Pass 1 assigns rate prizes using raw amt positions.
  // Pass 2 assigns amt prizes using EFFECTIVE rate prizes from pass 1,
  // so a person bumped up in rate (due to exclusions above them) is correctly
  // excluded from amt when their effective rate prize >= their potential amt prize.
  function computeBothAwardAssignments(byRate, byAmt, pa) {
    const rateTop = pa.rateTop10 || [];
    const amtTop  = pa.amtTop10  || [];
    const rateN   = Number(pa.rateConfig?.n) || rateTop.length;
    const amtN    = Number(pa.amtConfig?.n)  || amtTop.length;
    const bothEnabled = pa.bothEnabled && !!(pa.rateConfig && pa.amtConfig) && rateN > 0 && amtN > 0;
    // 순증 최소 기준 (minNet 조건이 활성화된 경우에만 적용)
    const rateMinNet = pa.rateConfig?.minNetEnabled ? Number(pa.rateConfig.minNet || 0) : 0;
    const amtMinNet  = pa.amtConfig?.minNetEnabled  ? Number(pa.amtConfig.minNet  || 0) : 0;

    const amtRankMap = {};
    byAmt.forEach((st, i) => { amtRankMap[st.s.empNo] = i; });

    // Pass 1: rate prizes
    // 물품(item) 시상은 calcRankAward=0 → 현금 비교 불가 → 순위(rank)로 비교
    const _chkTopElig = pa.isTopEligible || pa.isEligible;
    let rSlot = 0;
    const rateAsgn = new Map();
    for (const st of byRate) {
      const empNo = st.s.empNo;
      const rateNetMiss = rateMinNet > 0 && st.net < rateMinNet;
      if (!_chkTopElig(st.s) || rateNetMiss) {
        rateAsgn.set(empNo, { status: "ineligible", effectiveRank: 0, effectiveAmt: 0, otherAmt: 0,
          needsNet: rateNetMiss ? rateMinNet - st.net : 0 });
        continue;
      }
      const inRateRange = rSlot < rateN;
      if (!inRateRange) { rateAsgn.set(empNo, { status: "none", effectiveRank: 0, effectiveAmt: 0, otherAmt: 0 }); continue; }

      const rateCash   = rateTop[rSlot] || 0;
      const amtRankIdx = amtRankMap[empNo] !== undefined ? amtRankMap[empNo] : Infinity;
      const inAmtRange = bothEnabled && amtRankIdx < amtN;
      const amtCash    = inAmtRange ? (amtTop[amtRankIdx] || 0) : 0;

      if (!bothEnabled || !inAmtRange) {
        // 충돌 없음 → rate 획득
        rateAsgn.set(empNo, { status: "mine", effectiveRank: ++rSlot, effectiveAmt: rateCash, otherAmt: amtCash });
      } else if (rateCash > 0 || amtCash > 0) {
        // 현금 비교: 더 큰 시상 쪽으로 (동점은 rate 우선)
        if (amtCash > rateCash) {
          rateAsgn.set(empNo, { status: "other", effectiveRank: 0, effectiveAmt: 0, otherAmt: amtCash });
        } else {
          rateAsgn.set(empNo, { status: "mine", effectiveRank: ++rSlot, effectiveAmt: rateCash, otherAmt: amtCash });
        }
      } else {
        // 둘 다 물품(0원): 순위 비교 — 더 좋은 순위 쪽으로 (동점은 rate 우선)
        if (amtRankIdx < rSlot) {
          // amt 순위가 더 좋음 → amt 쪽으로
          rateAsgn.set(empNo, { status: "other", effectiveRank: 0, effectiveAmt: 0, otherAmt: 0 });
        } else {
          // rate 순위가 같거나 더 좋음 → rate 획득
          rateAsgn.set(empNo, { status: "mine", effectiveRank: ++rSlot, effectiveAmt: 0, otherAmt: 0 });
        }
      }
    }

    // Pass 2: amt prizes
    // rate에서 이미 "mine"을 받은 사람은 amt에서 "other"로 처리 (물품 포함)
    let aSlot = 0;
    const amtAsgn = new Map();
    for (const st of byAmt) {
      const empNo = st.s.empNo;
      const amtNetMiss = amtMinNet > 0 && st.net < amtMinNet;
      if (!_chkTopElig(st.s) || amtNetMiss) {
        amtAsgn.set(empNo, { status: "ineligible", effectiveRank: 0, effectiveAmt: 0, otherAmt: 0,
          needsNet: amtNetMiss ? amtMinNet - st.net : 0 });
        continue;
      }
      const inAmtRange = aSlot < amtN;
      if (!inAmtRange) { amtAsgn.set(empNo, { status: "none", effectiveRank: 0, effectiveAmt: 0, otherAmt: 0 }); continue; }

      const amtCash         = amtTop[aSlot] || 0;
      const rateResult      = rateAsgn.get(empNo);
      const effectiveRateCash = bothEnabled ? (rateResult?.effectiveAmt || 0) : 0;
      const rateAlreadyMine   = bothEnabled && rateResult?.status === "mine";

      if (!bothEnabled) {
        amtAsgn.set(empNo, { status: "mine", effectiveRank: ++aSlot, effectiveAmt: amtCash, otherAmt: 0 });
      } else if (amtCash > 0 || effectiveRateCash > 0) {
        // 현금 비교: rate 시상이 amt보다 크거나 같으면 → amt = other
        if (effectiveRateCash >= amtCash && rateAlreadyMine) {
          amtAsgn.set(empNo, { status: "other", effectiveRank: 0, effectiveAmt: 0, otherAmt: effectiveRateCash });
        } else {
          amtAsgn.set(empNo, { status: "mine", effectiveRank: ++aSlot, effectiveAmt: amtCash, otherAmt: effectiveRateCash });
        }
      } else {
        // 둘 다 물품: rate에서 이미 획득했으면 → amt = other
        if (rateAlreadyMine) {
          amtAsgn.set(empNo, { status: "other", effectiveRank: 0, effectiveAmt: 0, otherAmt: 0 });
        } else {
          amtAsgn.set(empNo, { status: "mine", effectiveRank: ++aSlot, effectiveAmt: 0, otherAmt: 0 });
        }
      }
    }

    return { rateAsgn, amtAsgn };
  }

  function renderProgressRankFullBothAware(byRate, byAmt, pa, kind) {
    const list  = kind === "rate" ? byRate : byAmt;
    const topN  = kind === "rate"
      ? (Number(pa.rateConfig?.n) || (pa.rateTop10 || []).length)
      : (Number(pa.amtConfig?.n)  || (pa.amtTop10  || []).length);
    const bothEnabled = pa.bothEnabled && !!(pa.rateConfig && pa.amtConfig) &&
      (pa.rateTop10 || []).length > 0 && (pa.amtTop10 || []).length > 0;

    const { rateAsgn, amtAsgn } = computeBothAwardAssignments(byRate, byAmt, pa);
    const asgnMap      = kind === "rate" ? rateAsgn : amtAsgn;
    const otherAsgnMap = kind === "rate" ? amtAsgn  : rateAsgn;

    const otherIcon = kind === "rate" ? "💰" : "📈";
    const otherName = kind === "rate" ? "신장액" : "신장률";

    // 이 종류의 시상 지급 내역 (payouts 배열)
    const myPayouts    = kind === "rate" ? (pa.rateConfig?.payouts || []) : (pa.amtConfig?.payouts  || []);
    const otherPayouts = kind === "rate" ? (pa.amtConfig?.payouts  || []) : (pa.rateConfig?.payouts || []);

    // 순위 표시 헬퍼: 불투명도 조절 버전 (미사용 유지)
    const RBMuted = (r) => `<span class="pg-rb rt" style="opacity:.38">${r}</span>`;
    // 시상 라벨: payouts 배열에서 해당 순위 아이템 찾아 표시 (물품명 포함)
    const prizeLabel = (rank, payoutsArr, fallbackAmt) => {
      const p = payoutsArr[rank - 1];
      return p != null ? payoutLabel(p) : (fallbackAmt > 0 ? `${Math.round(fallbackAmt / 10000)}만원` : "물품");
    };

    const rows = list.map((st, listIdx) => {
      const empNo      = st.s.empNo;
      const displayRank = listIdx + 1;            // 절대 위치(리스트 순서)
      const inTopN     = displayRank <= topN;
      const a = asgnMap.get(empNo) || { status: "none" };
      const rateVal = (st.rate != null ? st.rate.toFixed(1) : "0.0") + "%";
      const netVal  = (st.net >= 0 ? "+" : "") + Nf(st.net) + "원";
      let rankCell, prizeCell, rowCls = "";

      if (!inTopN || a.status === "none") {
        // TOP N 밖 — 순위·시상 표시 없음
        rankCell  = `<span style="color:#ccc;font-size:10px">-</span>`;
        prizeCell = `<span style="color:#bbb">-</span>`;
      } else if (a.status === "mine") {
        const rank = a.effectiveRank || displayRank;
        rankCell  = RB(rank);
        prizeCell = `<span class="pg-bdg pg-b-g">${escapeHtml(prizeLabel(rank, myPayouts, a.effectiveAmt))}</span>`;
      } else if (a.status === "other") {
        // 이 카테고리는 제외 → 순위 표시 없음, 상대 카테고리 시상명 표시
        const oa      = otherAsgnMap.get(empNo);
        const oRank   = oa?.effectiveRank || 0;
        const rankSuf = oRank > 0 ? ` ${oRank}위` : "";
        const oPrize  = oRank > 0 ? escapeHtml(prizeLabel(oRank, otherPayouts, oa?.effectiveAmt || 0)) : "";
        rankCell  = `<span style="color:#ccc;font-size:10px">-</span>`;
        prizeCell = `<span class="pg-bdg pg-rank-swap">${otherIcon} ${otherName}${rankSuf}${oPrize ? ` ${oPrize}` : ""} 시상</span>`;
        rowCls    = " pg-rank-deferred";
      } else if (a.status === "ineligible") {
        // 순증기준 미달 — 순위 숨김
        rankCell  = `<span style="color:#ccc;font-size:10px">-</span>`;
        const _needTxt = a.needsNet > 0
          ? `+${Nf(a.needsNet)}원 필요`
          : "기준미달";
        prizeCell = `<span class="pg-bdg pg-b-no">${_needTxt}</span>`;
      }

      return `<tr class="pg-tr-click${rowCls}" data-emp="${escapeHtml(empNo)}">
        <td>${rankCell}</td><td><strong>${escapeHtml(st.s.name || "")}</strong></td><td>${escapeHtml(st.s.branch || "")}</td><td class="r">${rateVal}</td><td class="r">${netVal}</td><td>${prizeCell}</td>
      </tr>`;
    }).join("");

    const notice = bothEnabled
      ? `<div class="pg-rank-notice">⚠️ 신장률·신장액 TOP 중복시상 없음 — 순위 높은 시상 1회만 지급 (<span style="background:#dbeafe;color:#1d4ed8;padding:0 4px;border-radius:3px">${otherIcon} ${otherName} 시상</span> 표시된 인원은 ${otherName} 항목에서 수상)</div>`
      : "";
    return notice + `<table class="pg-tbl">
      <thead><tr><th>#</th><th>성명</th><th>지점</th><th>달성률</th><th>순증(원)</th><th>시상</th></tr></thead>
      <tbody>${rows}</tbody>
    </table>`;
  }

  function openProgressPrintWindow() {
    const d = state.pgPrintData;
    if (!d || !d.stats?.length) { toast("표시할 데이터가 없습니다.", "error"); return; }
    const { byRate, byAmt, _pa, groupRanking, plan, stats, region, cohort, step, hasAnyTeam, _rateFinalDedup, _byAmtDedup } = d;
    const title = `${region}${cohort ? ` ${cohort}기` : ""} Step${step} 실적진도`;
    const today = new Date().toLocaleDateString("ko-KR");
    const shareUrl = `${location.origin}${location.pathname}#share?r=${encodeURIComponent(region)}&c=${encodeURIComponent(cohort || "")}&s=${encodeURIComponent(step)}`;

    // ─── 페이지 1: 신장 시상 + 그룹 시상 ───
    const rateHtml = _pa.rateConfig ? renderProgressRankFullBothAware(byRate, byAmt, _pa, "rate") : "";
    const amtHtml  = _pa.amtConfig  ? renderProgressRankFullBothAware(byRate, byAmt, _pa, "amt")  : "";
    const grpHtml  = groupRanking.length >= 2 ? renderGroupTable(groupRanking, plan, hasAnyTeam) : "";

    // ─── 페이지 2: 개인 실적 ───
    const piRows = byAmt.map((st, i) => {
      const rateStr = st.rate.toFixed(1);
      const netStr  = (st.net >= 0 ? "+" : "") + Nf(st.net);
      const aw = tierLabel(st.net, region, _pa);
      return `<tr>
        <td class="c">${i + 1}</td>
        <td>${escapeHtml(st.s.name || "")}</td>
        <td>${escapeHtml(st.s.branch || "")}</td>
        <td class="r">${Nf(st.base)}</td>
        <td class="r">${Nf(st.current)}</td>
        <td class="r">${rateStr}%</td>
        <td class="r">${netStr}</td>
        <td>${escapeHtml(aw)}</td>
      </tr>`;
    }).join("");

    const html = `<!DOCTYPE html>
<html lang="ko"><head><meta charset="UTF-8">
<title>${escapeHtml(title)}</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'Apple SD Gothic Neo','Malgun Gothic',sans-serif;font-size:11px;color:#111;background:#fff;}
.ctrl{position:fixed;top:0;left:0;right:0;z-index:999;background:#1e293b;color:#fff;display:flex;align-items:center;gap:10px;padding:8px 16px;flex-wrap:wrap;}
.ctrl span{flex:1;font-weight:700;font-size:13px;}
.cbtn{background:#3b82f6;color:#fff;border:none;padding:6px 14px;border-radius:6px;cursor:pointer;font-size:12px;font-weight:700;}
.cbtn.kk{background:#ffd700;color:#381f00;}
.page{padding:12mm 10mm;min-height:100vh;}
.page+.page{border-top:3px dashed #e2e8f0;}
.ph{display:flex;justify-content:space-between;align-items:flex-end;border-bottom:2px solid #1e293b;padding-bottom:6px;margin-bottom:10px;}
.ph h1{font-size:15px;font-weight:800;}
.ph small{font-size:10px;color:#64748b;}
.pg2{grid-template-columns:1fr 1fr;display:grid;gap:12px;}
.sect h3{font-size:12px;font-weight:700;margin-bottom:6px;padding:4px 8px;background:#f1f5f9;border-radius:4px;}
.pg-tbl{width:100%;border-collapse:collapse;font-size:10px;}
.pg-tbl th{background:#334155;color:#fff;padding:4px 5px;text-align:center;font-size:9px;}
.pg-tbl td{padding:3px 5px;border-bottom:1px solid #e2e8f0;}
.pg-tbl tr:nth-child(even) td{background:#f8fafc;}
.pg-rb{display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;border-radius:50%;font-size:9px;font-weight:800;color:#fff;}
.pg-rb.r1{background:#f59e0b;}.pg-rb.r2{background:#94a3b8;}.pg-rb.r3{background:#b45309;}
.pg-rb.rt{background:#cbd5e1;color:#334155;}
.pg-bdg{font-size:9px;padding:2px 5px;border-radius:3px;font-weight:600;}
.pg-b-g{background:#dcfce7;color:#166534;}.pg-b-no{background:#fee2e2;color:#991b1b;}
.pg-rank-swap{background:#dbeafe;color:#1e40af;}
.pg-rank-notice{font-size:9px;color:#d97706;margin-bottom:4px;padding:3px 6px;background:#fffbeb;border-radius:3px;}
.pg-tbl-wrap,.pg-tbl-scroll{overflow:visible!important;}
.pg-grp-tbl{width:100%;border-collapse:collapse;font-size:10px;}
.pg-grp-tbl th{background:#334155;color:#fff;padding:4px 5px;text-align:center;font-size:9px;}
.pg-grp-tbl td{padding:3px 5px;border-bottom:1px solid #e2e8f0;}
.pg-grp-badge{font-size:9px;padding:1px 4px;border-radius:3px;font-weight:600;display:inline-block;}
.pg-grp-ok{background:#dcfce7;color:#166534;}.pg-grp-half{background:#fef3c7;color:#92400e;}.pg-grp-no{background:#fee2e2;color:#991b1b;}
.pg-grp-award{background:#dcfce7;color:#166534;font-size:9px;padding:1px 5px;border-radius:3px;font-weight:700;}
.pg-grp-miss{color:#9ca3af;font-size:9px;}
.r{text-align:right!important;}.c{text-align:center!important;}
.note{font-size:9px;color:#64748b;margin-top:4px;}
@media print{
  @page{margin:8mm;}
  .ctrl{display:none!important;}
  .page{padding:0;min-height:unset;}
  .page+.page{border-top:none;page-break-before:always;}
  body{-webkit-print-color-adjust:exact;print-color-adjust:exact;}
}
</style></head>
<body>
<div class="ctrl">
  <span>🖨️ ${escapeHtml(title)} — ${today}</span>
  <button class="cbtn" onclick="window.print()">🖨️ 인쇄 / PDF 저장</button>
  <button class="cbtn kk" onclick="doKakaoShare()" id="btn-kk" style="display:none;">💬 카카오톡 공유</button>
  <button class="cbtn" style="background:#475569;" onclick="window.close()">✕ 닫기</button>
</div>

<div class="page" style="padding-top:48px;">
  <div class="ph"><h1>🏆 ${escapeHtml(title)} — 신장 시상 현황</h1><small>기준일: ${today}</small></div>
  <div class="pg2">
    ${_pa.rateConfig ? `<div class="sect"><h3>📈 신장률 TOP${_pa.rateConfig.n}${_pa.rateConfig.minNetEnabled ? ` · 순증 ${Nf(_pa.rateConfig.minNet)}원↑` : ""}</h3>${rateHtml}</div>` : ""}
    ${_pa.amtConfig  ? `<div class="sect"><h3>💰 신장액 TOP${_pa.amtConfig.n}${_pa.amtConfig.minNetEnabled  ? ` · 순증 ${Nf(_pa.amtConfig.minNet)}원↑`  : ""}</h3>${amtHtml}</div>` : ""}
  </div>
  ${grpHtml ? `<div class="sect" style="margin-top:10px;"><h3>🏅 ${hasAnyTeam ? "팀별" : "지점별"} 순증 시상</h3>${grpHtml}</div>` : ""}
  ${plan.notes ? `<div class="note">※ ${escapeHtml(plan.notes)}</div>` : ""}
</div>

<div class="page" style="padding-top:12mm;">
  <div class="ph"><h1>📊 ${escapeHtml(title)} — 개인 실적</h1><small>기준일: ${today} · 총 ${stats.length}명</small></div>
  <table class="pg-tbl">
    <thead><tr><th style="width:30px">#</th><th>성명</th><th>지점</th><th class="r">기준실적</th><th class="r">현재실적</th><th class="r">달성률</th><th class="r">순증</th><th>개인시상</th></tr></thead>
    <tbody>${piRows}</tbody>
  </table>
</div>

<script>
(function(){
  if (navigator.share || navigator.canShare) {
    document.getElementById("btn-kk").style.display = "";
  }
  window.doKakaoShare = function() {
    if (navigator.share) {
      navigator.share({ title: "${escapeHtml(title).replace(/"/g, '\\"')} 실적진도", text: "${escapeHtml(today)} 기준 실적진도 공유", url: "${shareUrl}" })
        .catch(function(){});
    }
  };
})();
<\/script>
</body></html>`;

    const win = window.open("", "_blank", "noopener,noreferrer");
    if (!win) { toast("팝업이 차단되었습니다. 팝업 허용 후 다시 시도하세요.", "error"); return; }
    win.document.write(html);
    win.document.close();
  }

  function bindProgressHomeEvents(list) {
    document.querySelectorAll("#progress-body .pg-tr-click").forEach((tr) => {
      tr.addEventListener("click", () => {
        if (state.pgShareMode) return;
        openProgressStudentPopup(tr.dataset.emp);
      });
    });
    // 모바일 TOP3 미리보기 카드의 이름 행 클릭 → 교육생 팝업 (버블링 방지)
    document.querySelectorAll("#progress-body .pg-pcard-row[data-emp]").forEach((row) => {
      row.addEventListener("click", (e) => {
        e.stopPropagation();
        e.preventDefault();
        if (state.pgShareMode) return;
        openProgressStudentPopup(row.dataset.emp);
      });
    });
    // 공유 링크 생성 버튼
    document.getElementById("btn-pg-share")?.addEventListener("click", () => {
      const region = state.progressRegion || "";
      const cohort = (state.progressCohort || "").replace(/기$/, "");
      const step   = state.filter.step || state.progressStep || "1";
      const base   = location.origin + location.pathname;
      const shareUrl = `${base}#share?r=${encodeURIComponent(region)}&c=${encodeURIComponent(cohort)}&s=${encodeURIComponent(step)}`;
      if (navigator.clipboard) {
        navigator.clipboard.writeText(shareUrl)
          .then(() => toast("링크가 클립보드에 복사되었습니다!", "success"))
          .catch(() => prompt("아래 링크를 복사하세요:", shareUrl));
      } else {
        prompt("아래 링크를 복사하세요:", shareUrl);
      }
    });
    // PDF 인쇄 버튼
    document.getElementById("btn-pg-print")?.addEventListener("click", openProgressPrintWindow);
    // 모바일 카드 자체 클릭 → 풀스크린 TOP10 모달
    document.querySelectorAll("#progress-body .pg-pcard[data-pcard]").forEach((card) => {
      card.addEventListener("click", () => {
        const key = card.dataset.pcard;
        const data = state._pgCardFullData && state._pgCardFullData[key];
        if (!data) return;
        openPgFullModal(data);
        if (key === "group") bindTacCardClicks();
      });
    });
    // 모바일: 전체 실적표 접기/펼치기
    const fullCard = document.getElementById("pg-full-tbl-card");
    const toggleBtn = document.getElementById("btn-pg-full-toggle");
    if (fullCard && toggleBtn) {
      toggleBtn.addEventListener("click", () => {
        fullCard.classList.toggle("show");
        toggleBtn.textContent = fullCard.classList.contains("show") ? "접기" : "펼쳐보기";
      });
    }

    // 마스터목표 개별조정 팝업 (한 번만 생성)
    if (!state._pgTblPendingGoals) state._pgTblPendingGoals = new Map();
    let adjPopup = document.getElementById("pg-goal-adj-popup");
    if (!adjPopup) {
      adjPopup = document.createElement("div");
      adjPopup.id = "pg-goal-adj-popup";
      adjPopup.className = "pg-goal-adj-popup";
      adjPopup.hidden = true;
      adjPopup.innerHTML = `<div class="pg-goal-adj-label">기준실적 기준 마스터목표 설정<span class="pg-goal-adj-base-lbl" id="pg-goal-adj-base-lbl"></span></div><div class="pg-goal-adj-btns"><button type="button" data-amt="-500000">-50만</button><button type="button" data-amt="-400000">-40만</button><button type="button" data-amt="-300000">-30만</button><button type="button" data-amt="-200000">-20만</button><button type="button" data-amt="-100000">-10만</button><button type="button" data-amt="-50000">-5만</button><button type="button" data-amt="50000">+5만</button><button type="button" data-amt="100000">+10만</button><button type="button" data-amt="200000">+20만</button><button type="button" data-amt="300000">+30만</button><button type="button" data-amt="400000">+40만</button><button type="button" data-amt="500000">+50만</button></div>`;
      document.body.appendChild(adjPopup);
      document.addEventListener("click", (e) => {
        if (!adjPopup.hidden && !adjPopup.contains(e.target) && !e.target.classList.contains("pg-goal-adj-btn")) {
          adjPopup.hidden = true;
        }
      });
      adjPopup.addEventListener("click", (e) => {
        const amtBtn = e.target.closest("button[data-amt]");
        if (!amtBtn || !adjPopup._emp) return;
        const amt = Number(amtBtn.dataset.amt);
        const newGoal = Math.max(0, (adjPopup._base || 0) + amt);
        const emp = adjPopup._emp;
        if (!state._pgTblPendingGoals) state._pgTblPendingGoals = new Map();
        state._pgTblPendingGoals.set(emp, { newGoal });
        const td = document.querySelector(`#pg-full-tbl-card td.pg-goal-cell[data-emp="${emp}"]`);
        if (td) {
          const valSpan = td.querySelector(".pg-goal-val");
          if (valSpan) valSpan.textContent = Nf(newGoal);
          const adjBtn = td.querySelector(".pg-goal-adj-btn");
          if (adjBtn) adjBtn.dataset.goal = String(newGoal);
          td.dataset.goal = String(newGoal);
          td.classList.add("pg-goal-pending");
        }
        adjPopup._currentGoal = newGoal;
        adjPopup.hidden = true;
        const saveBtn = document.getElementById("btn-pg-tbl-save");
        if (saveBtn) saveBtn.hidden = false;
      });
    }

    // 기존 pending 재적용 (re-render 후 DOM에 반영)
    if (state._pgTblPendingGoals && state._pgTblPendingGoals.size > 0) {
      state._pgTblPendingGoals.forEach(({ newGoal }, emp) => {
        const td = document.querySelector(`#pg-full-tbl-card td.pg-goal-cell[data-emp="${emp}"]`);
        if (td) {
          const valSpan = td.querySelector(".pg-goal-val");
          if (valSpan) valSpan.textContent = Nf(newGoal);
          td.classList.add("pg-goal-pending");
        }
      });
      const saveBtn2 = document.getElementById("btn-pg-tbl-save");
      if (saveBtn2) saveBtn2.hidden = false;
    }

    // ± 버튼 이벤트 (매 렌더마다 바인딩)
    document.querySelectorAll("#pg-full-tbl-card .pg-goal-adj-btn").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        const emp = btn.dataset.emp;
        const base = Number(btn.dataset.base) || 0;
        adjPopup._base = base;
        adjPopup._emp = emp;
        const baseLbl = adjPopup.querySelector(".pg-goal-adj-base-lbl");
        if (baseLbl) baseLbl.textContent = `  — 기준실적 ${Nf(base)}원`;
        const rect = btn.getBoundingClientRect();
        adjPopup.hidden = false;
        const popupW = 330;
        let left = rect.left;
        if (left + popupW > window.innerWidth - 8) left = window.innerWidth - popupW - 8;
        if (left < 8) left = 8;
        adjPopup.style.left = left + "px";
        adjPopup.style.top = (rect.bottom + 4) + "px";
      });
    });

    // 저장 버튼
    const saveBtn = document.getElementById("btn-pg-tbl-save");
    if (saveBtn) {
      saveBtn.addEventListener("click", async () => {
        if (!state._pgTblPendingGoals || state._pgTblPendingGoals.size === 0) return;
        saveBtn.disabled = true;
        saveBtn.textContent = "저장 중...";
        try {
          const updates = [];
          state._pgTblPendingGoals.forEach(({ newGoal }, empNo) => {
            const s = state.students.find((x) => x.empNo === empNo);
            if (s) updates.push({ ...s, target: newGoal });
          });
          if (updates.length > 0) {
            if (typeof window.DataAPI.saveMany === "function") await window.DataAPI.saveMany(updates);
            else for (const r of updates) await window.DataAPI.save(r);
          }
          state._pgTblPendingGoals.clear();
          saveBtn.hidden = true;
          document.querySelectorAll("#pg-full-tbl-card td.pg-goal-pending").forEach((td) => td.classList.remove("pg-goal-pending"));
        } catch (err) {
          alert("저장 실패: " + (err.message || err));
        } finally {
          saveBtn.disabled = false;
          saveBtn.textContent = "💾 변경 저장";
        }
      });
    }
  }

  function openProgressStudentPopup(empNo, pushStack) {
    const s = state.students.find((x) => x.empNo === empNo);
    if (!s) return;
    const st = getProgressStat(s);
    const _popupPa = getProgressAwardConfig(state.progressRegion,
      (state.filter.cohort || state.progressCohort || ""),
      (state.filter.step   || state.progressStep   || "1"));
    const netAwd = tierAward(st.net, undefined, _popupPa);
    const awdText = netAwd > 0 ? `${Nf(netAwd)}원 (${tierLabel(st.net, undefined, _popupPa)})` : "해당없음";
    const rateC = st.rate >= 120 ? "#0040b0" : st.rate >= 100 ? "#006030" : st.rate >= 80 ? "#884400" : "#880000";
    const netC  = st.net > 0 ? "#0040b0" : st.net < 0 ? "#880000" : "#333";
    const initial = (s.name || "?").trim().charAt(0) || "?";
    const targetDisp = Number(s.target) > 0 ? Number(s.target) : 0;
    function row(label, val) {
      return `<div class="pg-si-row"><span class="pg-si-label">${label}</span><span class="pg-si-val">${val || "—"}</span></div>`;
    }
    openPgFullModal({
      title: `👤 ${escapeHtml(s.name || "")} 교육생 정보`,
      subtitle: `${escapeHtml(s.region || "")} · ${escapeHtml(s.center || "")} · ${escapeHtml(s.branch || "")}`,
      bodyHTML: `
        <div class="pg-dm-id">
          <div class="pg-dm-id-avatar">${escapeHtml(initial)}</div>
          <div class="pg-dm-id-main">
            <div class="pg-dm-id-name">${escapeHtml(s.name || "(이름 없음)")}</div>
            <div class="pg-dm-id-meta">
              <span class="pg-dm-id-branch">🏢 ${escapeHtml(s.branch || "(미지정)")}</span>
              <span class="pg-dm-id-emp">사번 ${escapeHtml(s.empNo || "")}</span>
              ${s.cohort ? `<span class="pg-dm-id-cohort">${escapeHtml(s.cohort)}</span>` : ""}
              ${s.team ? `<span class="pg-dm-id-team">${escapeHtml(String(s.team))}조</span>` : ""}
            </div>
          </div>
        </div>

        <div class="pg-si-section">
          <div class="pg-si-head">📋 기본 정보</div>
          <div class="pg-si-grid">
            ${row("사번", escapeHtml(s.empNo || ""))}
            ${row("성명", escapeHtml(s.name || ""))}
            ${s.phone ? row("연락처", `<a href="tel:${escapeHtml(s.phone)}" class="pg-si-phone-link">${escapeHtml(s.phone)}</a>`) : row("연락처", "")}
            ${row("기수", escapeHtml(s.cohort || ""))}
            ${row("지역단", escapeHtml(s.region || ""))}
            ${row("비전센터", escapeHtml(s.center || ""))}
            ${row("지점", escapeHtml(s.branch || ""))}
            ${row("조편성", s.team ? `${escapeHtml(String(s.team))}조` : "")}
          </div>
        </div>

        <div class="pg-si-section">
          <div class="pg-si-head">🎯 실적 목표</div>
          <div class="pg-si-grid">
            ${row("기준실적", `${Nf(st.base)}원`)}
            ${row("마스터목표", targetDisp ? `${Nf(targetDisp)}원` : "")}
            ${row("아너스목표", Number(s.honors) ? `${Nf(Number(s.honors))}원` : "")}
            ${st.hiCap ? row("장기하이캡", `${Nf(st.hiCap)}원`) : ""}
          </div>
        </div>

        <div class="pg-si-section">
          <div class="pg-si-head">📊 현재 실적</div>
          <div class="pg-si-grid">
            ${row("현재실적", `${Nf(st.current)}원`)}
            ${row("달성률", `<span style="color:${rateC};font-weight:700">${st.base > 0 ? st.rate.toFixed(1) + "%" : "—"}</span>`)}
            ${row("순증", `<span style="color:${netC};font-weight:700">${st.net >= 0 ? "+" : ""}${Nf(st.net)}원</span>`)}
            ${st.ipumAmt ? row("인품(신상품)", `${st.ipumCount}건 · ${Nf(st.ipumAmt)}원`) : ""}
          </div>
        </div>

        <div class="pg-si-section pg-si-award">
          <div class="pg-si-head">🏆 예상 시상</div>
          <div class="pg-si-award-val">${awdText}</div>
        </div>
      `,
      closeLabel: pushStack ? "← 목록으로" : "← 돌아가기",
      pushStack: !!pushStack
    });
    // 면담관리 버튼 설정
    const pgModal = document.getElementById("modal-pg-full");
    const extraBtn = pgModal && pgModal.querySelector("#pg-full-modal-extra");
    if (extraBtn) {
      extraBtn.textContent = "📝 면담관리";
      extraBtn.hidden = false;
      extraBtn.onclick = () => {
        pgModal.hidden = true;
        selectStudent(s.empNo);
        switchView("#students");
      };
    }
  }

  // 재사용 가능한 풀스크린 실적진도 모달
  // pushStack=true 면 현재 모달 내용을 스택에 저장 후 새 내용 노출 → 닫을 때 이전 내용 복원
  function openPgFullModal({ title, subtitle, bodyHTML, closeLabel, pushStack }) {
    let modal = document.getElementById("modal-pg-full");
    if (!modal) {
      modal = document.createElement("div");
      modal.id = "modal-pg-full";
      modal.className = "modal pg-full-modal";
      modal.hidden = true;
      state._pgModalStack = [];
      modal.innerHTML = `
        <div class="modal-backdrop"></div>
        <div class="modal-panel pg-full-modal-panel">
          <div class="modal-head pg-full-modal-head">
            <div class="pg-full-modal-titles">
              <h3 id="pg-full-modal-title">제목</h3>
              <p id="pg-full-modal-sub"></p>
            </div>
            <button class="modal-close" data-pg-close aria-label="닫기">&times;</button>
          </div>
          <div class="modal-body pg-full-modal-body" id="pg-full-modal-body"></div>
          <div class="modal-foot">
            <button class="btn-outline" id="pg-full-modal-extra" hidden></button>
            <button class="btn-primary" id="pg-full-modal-close" data-pg-close>닫기</button>
          </div>
        </div>
      `;
      document.body.appendChild(modal);

      // 닫기 요소(×, 하단 버튼) — click + touchend for mobile (백드롭 클릭 닫기 제거)
      modal.querySelectorAll("[data-pg-close]").forEach((el) => {
        el.addEventListener("click", (e) => { e.stopPropagation(); closePgFullModal(); });
        el.addEventListener("touchend", (e) => { e.stopPropagation(); e.preventDefault(); closePgFullModal(); }, { passive: false });
      });
    }
    // 현재 모달이 이미 열려 있고 pushStack 옵션이면 기존 상태를 스택에 저장
    if (pushStack && !modal.hidden) {
      state._pgModalStack = state._pgModalStack || [];
      state._pgModalStack.push({
        title: modal.querySelector("#pg-full-modal-title").innerHTML,
        subtitle: modal.querySelector("#pg-full-modal-sub").innerHTML,
        bodyHTML: modal.querySelector("#pg-full-modal-body").innerHTML,
        closeLabel: modal.querySelector("#pg-full-modal-close").textContent
      });
    } else if (!pushStack) {
      // 최상위 진입 — 스택 초기화
      state._pgModalStack = [];
    }
    modal.querySelector("#pg-full-modal-title").innerHTML = title || "";
    modal.querySelector("#pg-full-modal-sub").innerHTML = subtitle || "";
    modal.querySelector("#pg-full-modal-body").innerHTML = bodyHTML || "";
    const closeBtn = modal.querySelector("#pg-full-modal-close");
    if (closeBtn) closeBtn.textContent = closeLabel || "닫기";
    // extra 버튼 항상 리셋 (각 호출처에서 직접 설정)
    const extraBtn = modal.querySelector("#pg-full-modal-extra");
    if (extraBtn) { extraBtn.hidden = true; extraBtn.onclick = null; extraBtn.textContent = ""; }
    modal.hidden = false;
    // 랭킹 테이블 포함 시 이름 검색바 삽입 (내 순위 찾기)
    const _modalBody = modal.querySelector("#pg-full-modal-body");
    if (_modalBody && _modalBody.querySelector(".pg-tbl")) {
      const _srchWrap = document.createElement("div");
      _srchWrap.className = "pg-modal-search";
      _srchWrap.innerHTML = `<input type="text" id="pg-modal-name-search" placeholder="이름으로 내 순위 찾기..." class="pg-modal-search-input" autocomplete="off">`;
      _modalBody.insertBefore(_srchWrap, _modalBody.firstChild);
      const _inp = _srchWrap.querySelector("input");
      _inp.addEventListener("input", () => {
        const q = _inp.value.trim();
        _modalBody.querySelectorAll(".pg-tbl tbody tr").forEach(tr => {
          const nm = tr.querySelector("td:nth-child(2)")?.textContent?.trim() || "";
          const hit = !q || nm.includes(q);
          tr.classList.toggle("pg-modal-search-match", hit && !!q);
          tr.classList.toggle("pg-modal-search-nomatch", !hit && !!q);
        });
        if (q) {
          const first = _modalBody.querySelector(".pg-modal-search-match");
          if (first) first.scrollIntoView({ behavior: "smooth", block: "center" });
        }
      });
    }
    // 모달 바디의 이름 클릭 → 교육생 상세 팝업(스택 push)
    modal.querySelectorAll(".pg-tr-click, .pg-pcard-row[data-emp]").forEach((el) => {
      el.addEventListener("click", (e) => {
        e.stopPropagation();
        const emp = el.dataset.emp;
        if (state.pgShareMode) return;
        if (emp) openProgressStudentPopup(emp, true /* pushStack */);
      });
    });
    // 바디 내 동적 닫기 버튼 바인딩
    modal.querySelector("#pg-full-modal-body").querySelectorAll("[data-pg-close]").forEach((el) => {
      el.addEventListener("click", (e) => { e.stopPropagation(); closePgFullModal(); });
      el.addEventListener("touchend", (e) => { e.stopPropagation(); e.preventDefault(); closePgFullModal(); }, { passive: false });
    });
  }

  // 모달 닫기 — 스택에 이전 내용이 있으면 복원, 없으면 완전 닫기
  function closePgFullModal() {
    const modal = document.getElementById("modal-pg-full");
    if (!modal) return;
    if (state._pgModalStack && state._pgModalStack.length > 0) {
      const prev = state._pgModalStack.pop();
      modal.querySelector("#pg-full-modal-title").innerHTML = prev.title || "";
      modal.querySelector("#pg-full-modal-sub").innerHTML = prev.subtitle || "";
      modal.querySelector("#pg-full-modal-body").innerHTML = prev.bodyHTML || "";
      const closeBtn = modal.querySelector("#pg-full-modal-close");
      if (closeBtn) closeBtn.textContent = prev.closeLabel || "닫기";
      // 스택 복원 시 extra 버튼 숨김
      const extraBtn2 = modal.querySelector("#pg-full-modal-extra");
      if (extraBtn2) { extraBtn2.hidden = true; extraBtn2.onclick = null; extraBtn2.textContent = ""; }
      // 복원된 바디의 이름 클릭 재바인딩
      modal.querySelectorAll(".pg-tr-click, .pg-pcard-row[data-emp]").forEach((el) => {
        el.addEventListener("click", (ev) => {
          ev.stopPropagation();
          const emp = el.dataset.emp;
          if (emp) openProgressStudentPopup(emp, true);
        });
      });
      // 복원된 바디의 동적 닫기 버튼 재바인딩
      modal.querySelector("#pg-full-modal-body").querySelectorAll("[data-pg-close]").forEach((el) => {
        el.addEventListener("click", (e) => { e.stopPropagation(); closePgFullModal(); });
        el.addEventListener("touchend", (e) => { e.stopPropagation(); e.preventDefault(); closePgFullModal(); }, { passive: false });
      });
      // 팀 카드 클릭 재바인딩 (팀 상세 닫고 팀 목록으로 복원 시)
      if (modal.querySelector("#pg-full-modal-body .tac-card[data-teamcard]")) bindTacCardClicks();
      return;
    }
    modal.hidden = true;
  }

  function openPgNewStudentModal(newRecords) {
    let modal = document.getElementById("modal-pg-new");
    if (!modal) {
      modal = document.createElement("div");
      modal.id = "modal-pg-new";
      modal.className = "modal";
      modal.hidden = true;
      modal.innerHTML = `
        <div class="modal-backdrop" data-pgn-close></div>
        <div class="modal-panel wide">
          <div class="modal-head">
            <h3>📥 교육생이 존재하지 않아 입력합니다</h3>
            <button class="modal-close" data-pgn-close>&times;</button>
          </div>
          <div class="modal-body" id="pg-new-modal-body"></div>
          <div class="modal-foot">
            <button class="btn-outline" data-pgn-close>취소</button>
            <button class="btn-primary" id="btn-pgn-save">💾 저장</button>
          </div>
        </div>
      `;
      document.body.appendChild(modal);
      modal.querySelectorAll("[data-pgn-close]").forEach((el) => {
        el.addEventListener("click", () => { modal.hidden = true; });
      });
    }

    const body = modal.querySelector("#pg-new-modal-body");
    body.innerHTML = `
      <p class="form-hint">아래 교육생은 시스템에 등록되어 있지 않습니다. 저장하면 신규 교육생으로 추가됩니다.</p>
      <div class="pg-tbl-wrap"><table class="pg-tbl">
        <thead><tr>
          <th>사원번호</th><th>성명</th><th>지점</th>
          <th>기준실적(원)</th><th>현재실적(원)</th>
          <th>기수</th><th>연락처</th>
        </tr></thead>
        <tbody>${newRecords.map((r, i) => `<tr>
          <td>${escapeHtml(r.empNo)}</td>
          <td><input class="pg-input pgn-name" data-idx="${i}" value="${escapeHtml(r.name)}" placeholder="성명"></td>
          <td>${escapeHtml(r.branch)}</td>
          <td class="r">${Nf(r.base)}</td>
          <td class="r">${Nf(r.pgCurrent)}</td>
          <td><input class="pg-input pgn-cohort" data-idx="${i}" value="" placeholder="기수"></td>
          <td><input class="pg-input pgn-phone" data-idx="${i}" value="" placeholder="연락처"></td>
        </tr>`).join("")}</tbody>
      </table></div>
    `;
    modal.hidden = false;

    // 저장 버튼 — 이전 리스너 제거 후 재등록
    const oldBtn = modal.querySelector("#btn-pgn-save");
    const saveBtn = oldBtn.cloneNode(true);
    oldBtn.parentNode.replaceChild(saveBtn, oldBtn);
    saveBtn.addEventListener("click", async () => {
      saveBtn.disabled = true;
      saveBtn.textContent = "저장중...";
      try {
        // region·center 누락 시 현재 진도 탭 지역단 + ORG_DATA branch→center 맵으로 자동 보완
        const _dfltRegion = state.progressRegion || state.filter.region || "";
        const _orgR = window.ORG_DATA?.regions?.find((r) => r.name === _dfltRegion);
        const _bcMap = {};
        if (_orgR) for (const c of _orgR.centers) for (const b of c.branches) _bcMap[b] = c.name;

        const records = newRecords.map((r, i) => {
          const rgn = r.region || _dfltRegion;
          const ctr = r.center || _bcMap[r.branch] || "";
          return {
            ...r,
            region: rgn,
            center: ctr,
            name:   (modal.querySelector(`.pgn-name[data-idx="${i}"]`)?.value.trim()  || r.name),
            cohort: (modal.querySelector(`.pgn-cohort[data-idx="${i}"]`)?.value.trim() || ""),
            phone:  (modal.querySelector(`.pgn-phone[data-idx="${i}"]`)?.value.trim()  || ""),
            target: 0, honors: 0
          };
        });
        if (typeof window.DataAPI.saveMany === "function") {
          await window.DataAPI.saveMany(records);
        } else {
          for (const rec of records) await window.DataAPI.save(rec);
        }
        toast(`${records.length}명 신규 교육생 등록 완료`, "success");
        modal.hidden = true;
      } catch (e) {
        console.error(e);
        toast("저장 실패: " + e.message, "error");
        saveBtn.disabled = false;
        saveBtn.textContent = "💾 저장";
      }
    });
  }

  function renderProgressAdmin(list) {
    const rows = list.slice().sort((a, b) => (a.branch || "").localeCompare(b.branch || "") || (a.name || "").localeCompare(b.name || ""));
    const stats = rows.map(getProgressStat);
    const total = stats.length;
    const over5 = stats.filter((s) => s.net >= 50000).length;
    const mid80 = stats.filter((s) => s.rate >= 80 && s.rate < 100).length;
    const mid50 = stats.filter((s) => s.rate >= 50 && s.rate < 80).length;
    const low50 = stats.filter((s) => s.rate < 50).length;
    const avgR = total > 0 ? (stats.reduce((a, s) => a + s.rate, 0) / total) : 0;
    const _adminPa = getProgressAwardConfig(state.progressRegion);
    const elig = stats.filter((s) => _adminPa.isEligible(s.s)).length;
    const a5 = stats.filter((s) => s.net >= 500000).length;
    const a4 = stats.filter((s) => s.net >= 300000 && s.net < 500000).length;
    const cash = stats.filter((s) => s.net >= 50000 && s.net < 300000).length;
    const exc = total - elig - over5 + a5 + a4; // 제외: 순증 5만 미만
    const exclude = stats.filter((s) => s.net < 50000).length;
    const today = new Date();
    const baseDate = today.getFullYear() + "-" + String(today.getMonth() + 1).padStart(2, "0") + "-" + String(today.getDate()).padStart(2, "0");
    const firstCohort = rows[0]?.cohort || "";
    const cohortTitle = firstCohort ? `${firstCohort} ` : "";


    // ── 카드1: 기수별 인원 (현재 지역단 내) ──────────────────────────────────
    const _cohortMap = {};
    rows.forEach((s) => {
      const key = s.cohort ? `${s.cohort}기` : "(기수 미지정)";
      _cohortMap[key] = (_cohortMap[key] || 0) + 1;
    });
    const _cohortEntries = Object.entries(_cohortMap).sort(([a], [b]) => {
      const na = parseInt(a, 10); const nb = parseInt(b, 10);
      return isNaN(na) || isNaN(nb) ? a.localeCompare(b) : na - nb;
    });

    // ── 카드2: 지역단별 인원 그리드 (state.students 전체, 단순 카운트) ───────────
    const _rgnMap = {};
    state.students.forEach((s) => {
      const r = s.region || "(미지정)";
      if (!_rgnMap[r]) _rgnMap[r] = 0;
      _rgnMap[r]++;
    });
    const _rgnEntries = Object.entries(_rgnMap)
      .filter(([r]) => r !== "(미지정)" && !r.includes("|"))
      .sort(([a], [b]) => a.localeCompare(b));
    const _rgnGridHTML = _rgnEntries.map(([reg, cnt]) => {
      const isCur = reg === state.progressRegion;
      return `<div class="pg-rgn-cell${isCur ? " pg-rgn-cur" : ""}">
        <div class="pg-rgn-name">${escapeHtml(reg)}</div>
        <div class="pg-rgn-cnt">${cnt}명</div>
      </div>`;
    }).join("");

    // ── 카드3: 전체 교육생 단순 통계 ─────────────────────────────────────────
    const _allTotal = state.students.length;
    const _rgnCount = _rgnEntries.length;
    // 기수별 카운트
    const _cohortAll = {};
    state.students.forEach((s) => {
      const c = s.cohort ? `${s.cohort}기` : "(미지정)";
      _cohortAll[c] = (_cohortAll[c] || 0) + 1;
    });
    const _cohortAllEntries = Object.entries(_cohortAll).sort(([a], [b]) => {
      const na = parseInt(a, 10); const nb = parseInt(b, 10);
      return isNaN(na) || isNaN(nb) ? a.localeCompare(b) : na - nb;
    });
    return `
      <div class="pg-wrap">
        <!-- [1] 기수별 인원 카드 -->
        <div class="pg-info-grid">
          <div class="pg-info-card">
            <h5>📘 ${escapeHtml(state.progressRegion)} 기수별 인원</h5>
            <dl>
              ${_cohortEntries.map(([c, n]) => {
                const isSpecial = c === "(기수 미지정)";
                const origCohort = isSpecial ? "" : c.replace(/기$/, "");
                const delBtn = isSpecial ? "" : `<button type="button" class="btn-cohort-del" data-cohort="${escapeHtml(origCohort)}" data-scope="region" title="${escapeHtml(c)} ${n}명 삭제">🗑️</button>`;
                return `<dt>${escapeHtml(c)}</dt><dd><strong>${n}명</strong>${delBtn}</dd>`;
              }).join("")}
              <dt class="pg-dt-total">합계</dt><dd class="pg-dt-total"><strong>${total}명</strong></dd>
              <dt>기준일</dt><dd>${baseDate}</dd>
            </dl>
          </div>
          <!-- [2] 지역단 인원 그리드 -->
          <div class="pg-info-card pg-info-card--wide">
            <h5>🗺️ 지역단별 인원 현황</h5>
            <div class="pg-rgn-grid">${_rgnGridHTML}</div>
          </div>
          <!-- [3] 전체 교육생 현황 -->
          <div class="pg-info-card">
            <h5>📊 전체 현황 (${_allTotal}명)</h5>
            <dl>
              <dt>지역단 수</dt><dd><strong>${_rgnCount}개</strong></dd>
              ${_cohortAllEntries.map(([c, n]) => {
                const isSpecial = c === "(미지정)";
                const origCohort = isSpecial ? "" : c.replace(/기$/, "");
                const delBtn = isSpecial ? "" : `<button type="button" class="btn-cohort-del" data-cohort="${escapeHtml(origCohort)}" data-scope="all" title="전체: ${escapeHtml(c)} ${n}명 삭제">🗑️</button>`;
                return `<dt>${escapeHtml(c)}</dt><dd>${n}명${delBtn}</dd>`;
              }).join("")}
            </dl>
          </div>
        </div>
        <!-- [4] 안내 -->
        <div class="pg-admin-note">
          <strong>🔧 ${escapeHtml(state.progressRegion)} 실적 입력</strong>
          <p>아래 각 섹션을 클릭해 펼친 뒤, 값을 입력하고 [💾 저장] 을 누르세요. 기준실적은 교육생 등록화면에서 수정할 수 있어요.</p>
        </div>

        <!-- [5] 일괄 붙여넣기 (총괄월별실적 / 실적진도현황) -->
        <details class="pg-accordion">
          <summary>
            <span class="pg-ac-title">📋 현재실적 일괄 붙여넣기</span>
            <span class="pg-ac-sub">— 붙여넣기 방식 선택</span>
            <span class="pg-ac-chev">▾</span>
          </summary>
          <div class="pg-ac-body">
            <div class="pg-paste-mode-btns">
              <button class="btn-outline pg-paste-mode-btn active" data-mode="monthly">📊 총괄월별실적 붙여넣기</button>
              <button class="btn-outline pg-paste-mode-btn" data-mode="progress">📈 실적진도현황 붙여넣기</button>
              <button class="btn-outline pg-paste-mode-btn" data-mode="honors">🏆 아너스목표 붙여넣기</button>
              <button class="btn-outline pg-paste-mode-btn" data-mode="ipum">✨ 인품실적 붙여넣기</button>
            </div>

            <!-- 총괄월별실적 -->
            <div id="pg-paste-mode-monthly" class="pg-admin-paste">
              <div class="pg-paste-desc">"사번 장기하이캡 실적" 탭/공백으로 구분 (단위: 천원)</div>
              <textarea id="pg-paste" rows="6" placeholder="예)
1B1312	15,613	710
986037	61,050	2,782
9A1520	45,624	1,955
9A1766	37,956	1,699"></textarea>
              <div class="pg-actions">
                <button class="btn-primary" id="btn-pg-paste-apply">📥 총괄월별실적 저장</button>
                <button class="btn-outline small" id="btn-pg-paste-clear">🗑 초기화</button>
                <span id="pg-paste-msg" class="pg-msg"></span>
              </div>
            </div>

            <!-- 실적진도현황 -->
            <div id="pg-paste-mode-progress" class="pg-admin-paste" style="display:none">
              <div class="pg-progress-paste-step-row">
                <span class="pg-paste-save-target">
                  📌 <strong>저장 대상:</strong>
                  <span id="pg-paste-region-disp" class="pg-paste-region-badge">─</span>
                  <select id="pg-paste-cohort-sel" class="pg-paste-cohort-sel">
                    <option value="">기수 선택 ▾</option>
                    <option value="1기">1기</option>
                    <option value="2기">2기</option>
                    <option value="3기">3기</option>
                    <option value="4기">4기</option>
                    <option value="5기">5기</option>
                  </select>
                </span>
                <span class="pg-paste-step-sep">│</span>
                <strong>저장할 스텝:</strong>
                <label><input type="radio" name="pg-progress-paste-step" value="1" checked> Step 1</label>
                <label><input type="radio" name="pg-progress-paste-step" value="2"> Step 2</label>
              </div>
              <div class="pg-paste-desc"><strong>[Step 1 / Step 2 공통 형식]</strong> 지역단·비전센터·지점·사원번호·성명·위촉차월·기준실적·현재실적·계약건수·실적 (탭 구분, 금액단위: 원) — 스텝 선택에 따라 해당 스텝 실적에 저장</div>
              <textarea id="pg-progress-paste" rows="7" placeholder="지역단	비전센터	지점	사원번호	성명	위촉차월	기준실적	현재실적	계약건수	실적
강북지역단	성동비전센터	강북수유지점	069563	권명숙	340	274273	199270	0	0
강북지역단	성동비전센터	강북수유지점	070041	김미영	339	209138	104130	0	0
강북지역단	성동비전센터	강북수유지점	111435	최연화	251	506915	350200	2	144330"></textarea>
              <div id="pg-progress-paste-confirm" class="pg-paste-confirm" hidden>
                <div class="pg-paste-confirm-msg" id="pg-progress-paste-confirm-msg"></div>
                <div class="pg-paste-confirm-btns">
                  <button class="btn-primary small" id="btn-pg-progress-paste-yes">✅ 예, 저장</button>
                  <button class="btn-outline small" id="btn-pg-progress-paste-no">❌ 아니오</button>
                </div>
              </div>
              <div class="pg-actions">
                <button class="btn-primary" id="btn-pg-progress-paste-apply">📥 실적진도현황 저장</button>
                <span id="pg-progress-paste-msg" class="pg-msg"></span>
              </div>
            </div>

            <!-- 아너스목표 붙여넣기 -->
            <div id="pg-paste-mode-honors" class="pg-admin-paste" style="display:none">
              <div class="pg-paste-desc">사번·아너스목표 (탭 구분, 금액단위: 원) — 사번 미매칭 데이터는 자동으로 제외됩니다</div>
              <textarea id="pg-honors-paste" rows="6" placeholder="예)
959167	1500000
1B1312	2000000"></textarea>
              <div class="pg-actions">
                <button class="btn-primary" id="btn-pg-honors-paste-apply">📥 아너스목표 저장</button>
                <span id="pg-honors-paste-msg" class="pg-msg"></span>
              </div>
            </div>
            <!-- 인품실적 붙여넣기 (전체 사번 기준) -->
            <div id="pg-paste-mode-ipum" class="pg-admin-paste" style="display:none">
              <div class="pg-paste-desc">사번·인품건수·인품실적 (탭/공백 구분, 금액단위: 원) — <strong>전체 지역단</strong> 대상, 사번 미매칭 건 자동 제외</div>
              <textarea id="pg-ipum-global-paste" rows="6" placeholder="예)
1001234	3	450000
1005678	2	300000"></textarea>
              <div class="pg-actions">
                <button class="btn-primary" id="btn-pg-ipum-global-apply">📥 인품실적 저장 (전체)</button>
                <button class="btn-outline small" id="btn-pg-ipum-global-clear">🗑 초기화</button>
                <span id="pg-ipum-global-msg" class="pg-msg"></span>
              </div>
            </div>
          </div>
        </details>

        <!-- [6] 메인 편집 테이블 (접기/펼치기) -->
        <details class="pg-accordion">
          <summary>
            <span class="pg-ac-title">✏️ 메인 편집 테이블</span>
            <span class="pg-ac-sub">— 현재실적 / 인품건 / 인품실적 수정 (${rows.length}명)</span>
            <span class="pg-ac-chev">▾</span>
          </summary>
          <div class="pg-ac-body">
            <div class="pg-tbl-wrap"><table class="pg-tbl pg-admin-tbl">
              <thead><tr>
                <th>#</th><th>지점</th><th>성명</th>
                <th>기준실적(원)</th><th>장기하이캡(원)</th><th>현재실적(원)</th><th>달성률</th><th>순증</th>
                <th>인품건</th><th>인품실적(원)</th>
              </tr></thead>
              <tbody>${rows.map((s, i) => {
                const sfx   = _pgStepSfx();
                const base  = Number(s.base || 0);
                const hiCap = Number(s[`hiCap${sfx}`] || 0);
                const cur   = sfx ? Number(s[`pgCurrent${sfx}`] || 0)
                                  : (s.pgCurrent !== undefined ? Number(s.pgCurrent) : Number(s.current || 0));
                const iCnt  = Number(s[`pgIpumCount${sfx}`] || 0);
                const iAmt  = Number(s[`pgIpumAmt${sfx}`]   || 0);
                const net   = cur - base;
                const rate  = base > 0 ? (cur / base) * 100 : 0;
                return `<tr>
                  <td>${i + 1}</td>
                  <td>${escapeHtml(s.branch || "")}</td>
                  <td><strong>${escapeHtml(s.name || "")}</strong></td>
                  <td class="r">${Nf(base)}</td>
                  <td><input type="number" class="pg-input" data-emp="${escapeHtml(s.empNo)}" data-f="${sfx ? `hiCap${sfx}` : 'hiCap'}" value="${hiCap}" min="0" step="1"></td>
                  <td><input type="number" class="pg-input" data-emp="${escapeHtml(s.empNo)}" data-f="${sfx ? `pgCurrent${sfx}` : 'current'}" data-pg-role="current" value="${cur}" min="0" step="1"></td>
                  <td class="r" data-calc="rate-${escapeHtml(s.empNo)}">${rate.toFixed(1)}%</td>
                  <td class="r" data-calc="net-${escapeHtml(s.empNo)}">${net >= 0 ? "+" : ""}${Nf(net)}</td>
                  <td><input type="number" class="pg-input" data-emp="${escapeHtml(s.empNo)}" data-f="${sfx ? `pgIpumCount${sfx}` : 'pgIpumCount'}" value="${iCnt}" min="0"></td>
                  <td><input type="number" class="pg-input" data-emp="${escapeHtml(s.empNo)}" data-f="${sfx ? `pgIpumAmt${sfx}` : 'pgIpumAmt'}" value="${iAmt}" min="0"></td>
                </tr>`;
              }).join("")}</tbody>
            </table></div>
            <div class="pg-actions" style="margin-top:12px;">
              <button class="btn-primary" id="btn-pg-save">💾 저장 (Firestore)</button>
              <span id="pg-save-msg" class="pg-msg"></span>
            </div>
          </div>
        </details>

        <!-- [7] 인품 데이터 편집 섹션 (접기/펼치기) -->
        <details class="pg-accordion pg-ipum-accordion">
          <summary>
            <span class="pg-ac-title">✨ 인품 데이터 편집</span>
            <span class="pg-ac-sub">— 신상품 계약건 / 신상품 실적</span>
            <span class="pg-ac-chev">▾</span>
          </summary>
          <div class="pg-ac-body">
            <div class="pg-admin-paste">
              <strong>📋 인품 붙여넣기 — "사번 인품건 인품실적" (탭/공백 구분)</strong>
              <textarea id="pg-ipum-paste" rows="4" placeholder="예)
1001234  3  450000
1005678  2  300000"></textarea>
              <div class="pg-actions">
                <button class="btn-outline" id="btn-pg-ipum-paste-apply">📥 인품 붙여넣기 반영</button>
                <button class="btn-outline small" id="btn-pg-ipum-paste-clear">🗑 초기화</button>
                <span id="pg-ipum-paste-msg" class="pg-msg"></span>
              </div>
            </div>
            <div class="pg-tbl-wrap"><table class="pg-tbl pg-admin-tbl">
              <thead><tr>
                <th>#</th><th>지점</th><th>성명</th>
                <th>신상품 계약건</th><th>신상품 실적(원)</th>
              </tr></thead>
              <tbody>${rows.map((s, i) => `<tr>
                <td>${i + 1}</td>
                <td>${escapeHtml(s.branch || "")}</td>
                <td><strong>${escapeHtml(s.name || "")}</strong></td>
                <td><input type="number" class="pg-input" data-emp="${escapeHtml(s.empNo)}" data-f2="ipumCount" value="${Number(s.ipumCount || 0)}" min="0"></td>
                <td><input type="number" class="pg-input" data-emp="${escapeHtml(s.empNo)}" data-f2="ipumAmt" value="${Number(s.ipumAmt || 0)}" min="0"></td>
              </tr>`).join("")}</tbody>
            </table></div>
            <div class="pg-actions" style="margin-top:12px;">
              <button class="btn-primary" id="btn-pg-ipum-save">💾 인품 저장 (Firestore)</button>
              <span id="pg-ipum-save-msg" class="pg-msg"></span>
            </div>
          </div>
        </details>

        <!-- [8] 팀 배정 섹션 -->
        <details class="pg-accordion pg-team-accordion">
          <summary>
            <span class="pg-ac-title">🏅 팀 배정</span>
            <span class="pg-ac-sub">— 그룹 순증 시상용 팀 구성</span>
            <span class="pg-ac-chev">▾</span>
          </summary>
          <div class="pg-ac-body">
            <div class="pg-team-controls">
              <label class="pg-team-count-label">팀 수
                <input type="number" id="pg-team-count" min="2" max="20" value="8" style="width:60px;margin-left:6px;">
              </label>
              <button class="btn-outline" id="btn-pg-team-auto">🎲 자동 배정 (랜덤 고르게)</button>
              <button class="btn-outline small" id="btn-pg-team-clear">🗑 전체 초기화</button>
              <span id="pg-team-msg" class="pg-msg"></span>
            </div>
            <div class="pg-team-summary" id="pg-team-summary">
              <!-- 팀별 합산 실적 박스 (실시간) -->
            </div>
            <div class="pg-team-filter-bar">
              <label class="pg-team-filter-lbl">기수</label>
              <select id="pg-team-cohort-filter" class="pg-input pg-team-cohort-sel">
                <option value="">전체 기수</option>
                ${[...new Set(rows.map(s => s.cohort || "").filter(Boolean))].sort().map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join("")}
              </select>
              <span class="pg-team-unassigned-info">미배정: <strong id="pg-team-unassigned-cnt">${rows.filter(s => !(parseInt(s.team, 10) > 0)).length}</strong>명</span>
            </div>
            <div class="pg-tbl-wrap"><table class="pg-tbl pg-admin-tbl">
              <thead><tr>
                <th>#</th><th>기수</th><th>지점</th><th>성명</th>
                <th>기준실적</th><th>현재실적</th><th>팀</th>
              </tr></thead>
              <tbody>${rows.slice().sort((a, b) => {
                const aT = parseInt(a.team, 10), bT = parseInt(b.team, 10);
                const aHas = aT > 0, bHas = bT > 0;
                if (!aHas && bHas) return -1;
                if (aHas && !bHas) return 1;
                if (aHas && bHas && aT !== bT) return aT - bT;
                return (a.cohort || "").localeCompare(b.cohort || "") || (a.branch || "").localeCompare(b.branch || "") || (a.name || "").localeCompare(b.name || "");
              }).map((s, i) => {
                const base = Number(s.base || 0);
                const cur = Number(s.current || 0);
                const teamVal = parseInt(s.team, 10);
                const hasTeam = teamVal > 0;
                return `<tr class="pg-team-tbl-row${!hasTeam ? " pg-team-unassigned-row" : ""}" data-cohort="${escapeHtml(s.cohort || "")}">
                  <td>${i + 1}</td>
                  <td><span class="pg-cohort-tag">${escapeHtml(s.cohort || "—")}</span></td>
                  <td>${escapeHtml(s.branch || "")}</td>
                  <td><strong>${escapeHtml(s.name || "")}</strong></td>
                  <td class="r">${Nf(base)}</td>
                  <td class="r">${Nf(cur)}</td>
                  <td><input type="number" class="pg-input pg-team-input" data-emp="${escapeHtml(s.empNo)}" value="${hasTeam ? teamVal : ""}" placeholder="조번호" min="1" style="width:70px"></td>
                </tr>`;
              }).join("")}</tbody>
            </table></div>
            <div class="pg-actions" style="margin-top:12px;">
              <button class="btn-primary" id="btn-pg-team-save">💾 팀 배정 저장 (Firestore)</button>
              <span id="pg-team-save-msg" class="pg-msg"></span>
            </div>
          </div>
        </details>
      </div>
    `;
  }

  // 팀별 합산 실적 계산 — 입력 중인 현재값(DOM) 반영
  function computeTeamSummary(list, root) {
    root = root || document.getElementById("progress-body");
    const byTeam = {};
    list.forEach((s) => {
      const inp = root ? root.querySelector(`.pg-team-input[data-emp="${s.empNo}"]`) : null;
      // 기수 필터로 숨겨진 행은 집계에서 제외
      if (inp && inp.closest("tr")?.hidden) return;
      const team = (inp ? inp.value : s.team || "").toString().trim();
      if (!team) return;
      if (!byTeam[team]) byTeam[team] = { base: 0, current: 0, members: [] };
      byTeam[team].base += Number(s.base || 0);
      byTeam[team].current += Number(s.current || 0);
      byTeam[team].members.push(s.name || "");
    });
    return byTeam;
  }

  function renderTeamSummary(list, root) {
    const box = document.getElementById("pg-team-summary");
    if (!box) return;
    const byTeam = computeTeamSummary(list, root);
    const keys = Object.keys(byTeam).sort((a, b) => {
      // "1팀", "2팀"... 숫자 우선 정렬
      const na = parseInt(a, 10), nb = parseInt(b, 10);
      if (!isNaN(na) && !isNaN(nb)) return na - nb;
      return a.localeCompare(b);
    });
    if (!keys.length) {
      box.innerHTML = `<div class="pg-empty">팀 배정된 교육생이 없습니다. 아래 표의 "팀" 칸에 값을 입력하면 자동 집계됩니다.</div>`;
      return;
    }
    const ranked = keys.map((k) => {
      const g = byTeam[k];
      const rate = g.base > 0 ? (g.current / g.base) * 100 : 0;
      const net = g.current - g.base;
      return { name: k, ...g, rate, net };
    }).sort((a, b) => b.rate - a.rate);
    box.innerHTML = `
      <div class="pg-team-cards">
        ${ranked.map((g, i) => `
          <div class="pg-team-card">
            <div class="pg-team-rank ${i===0?"r1":i===1?"r2":i===2?"r3":"rt"}">${i+1}</div>
            <div class="pg-team-info">
              <div class="pg-team-name">${escapeHtml(g.name)} <small>(${g.members.length}명)</small></div>
              <div class="pg-team-mem" title="${escapeHtml(g.members.join(", "))}">${escapeHtml(g.members.slice(0, 4).join("·"))}${g.members.length > 4 ? "…" : ""}</div>
            </div>
            <div class="pg-team-stat">
              <div class="pg-team-rate">${g.rate.toFixed(1)}%</div>
              <div class="pg-team-net ${g.net >= 0 ? "pg-net-p" : "pg-net-m"}">${g.net >= 0 ? "+" : ""}${Nf(g.net)}</div>
            </div>
          </div>
        `).join("")}
      </div>
    `;
  }

  function bindProgressAdminEvents(list, rootId = "progress-body") {
    const root = document.getElementById(rootId);
    if (!root) return;
    // 붙여넣기 저장 대상 초기화 (지역단·기수)
    const _rdDisp = root.querySelector("#pg-paste-region-disp");
    const _cSel   = root.querySelector("#pg-paste-cohort-sel");
    if (_rdDisp) _rdDisp.textContent = state.filter.region || "─";
    if (_cSel && state.filter.cohort) _cSel.value = state.filter.cohort;
    // 실시간 계산 (현재실적 변경 시 달성률/순증 재계산)
    root.querySelectorAll(".pg-input[data-pg-role='current']").forEach((inp) => {
      inp.addEventListener("input", (e) => {
        const emp = e.target.dataset.emp;
        const s = state.students.find((x) => x.empNo === emp);
        const sfx = _pgStepSfx();
        const base = Number(s?.base || 0);
        const cur  = parseFloat(e.target.value) || 0;
        const net  = cur - base;
        const rate = base > 0 ? (cur / base) * 100 : 0;
        const rateEl = document.querySelector(`[data-calc="rate-${emp}"]`);
        const netEl  = document.querySelector(`[data-calc="net-${emp}"]`);
        if (rateEl) rateEl.textContent = rate.toFixed(1) + "%";
        if (netEl)  netEl.textContent  = (net >= 0 ? "+" : "") + Nf(net);
      });
    });

    // 저장
    const saveBtn = $("#btn-pg-save");
    if (saveBtn) saveBtn.addEventListener("click", async () => {
      const updates = {};
      root.querySelectorAll(".pg-input").forEach((inp) => {
        const emp = inp.dataset.emp;
        const f = inp.dataset.f;
        if (!updates[emp]) updates[emp] = {};
        updates[emp][f] = parseFloat(inp.value) || 0;
      });
      const msg = $("#pg-save-msg");
      if (msg) msg.textContent = "저장중...";
      saveBtn.disabled = true;
      try {
        const records = Object.keys(updates).map((emp) => {
          const s = state.students.find((x) => x.empNo === emp);
          return { ...s, ...updates[emp] };
        });
        if (typeof window.DataAPI.saveMany === "function") {
          await window.DataAPI.saveMany(records);
        } else {
          for (const r of records) await window.DataAPI.save(r);
        }
        if (msg) { msg.textContent = `✅ ${records.length}건 저장 완료`; msg.className = "pg-msg ok"; }
        toast(`${records.length}건 저장 완료`, "success");
      } catch (e) {
        console.error(e);
        if (msg) { msg.textContent = "❌ 저장 실패: " + e.message; msg.className = "pg-msg err"; }
        toast("저장 실패: " + e.message, "error");
      }
      saveBtn.disabled = false;
    });

    // 붙여넣기 모드 전환 버튼
    root.querySelectorAll(".pg-paste-mode-btn").forEach((btn) => {
      btn.addEventListener("click", () => {
        root.querySelectorAll(".pg-paste-mode-btn").forEach((b) => b.classList.remove("active"));
        btn.classList.add("active");
        const mode = btn.dataset.mode;
        const monthlyDiv = document.getElementById("pg-paste-mode-monthly");
        const progressDiv = document.getElementById("pg-paste-mode-progress");
        const honorsDiv  = document.getElementById("pg-paste-mode-honors");
        const ipumGlDiv  = document.getElementById("pg-paste-mode-ipum");
        if (monthlyDiv)  monthlyDiv.style.display  = mode === "monthly"  ? "" : "none";
        if (progressDiv) progressDiv.style.display = mode === "progress" ? "" : "none";
        if (honorsDiv)   honorsDiv.style.display   = mode === "honors"   ? "" : "none";
        if (ipumGlDiv)   ipumGlDiv.style.display   = mode === "ipum"     ? "" : "none";
      });
    });

    // 스텝 라디오 변경 → 설명 텍스트 업데이트 (Step1/2 공통 포맷이므로 스텝 이름만 바꿔 표시)
    const _pgDescEl = document.querySelector("#pg-paste-mode-progress .pg-paste-desc");
    const _pgPasteTA = document.querySelector("#pg-progress-paste");
    const _buildPasteDesc = (step) => {
      const sfxLabel = step === "2" ? "Step 2" : "Step 1";
      return `<strong>[${sfxLabel} 저장]</strong> 지역단·비전센터·지점·사원번호·성명·위촉차월·기준실적·현재실적·계약건수·실적 (탭 구분, 금액단위: 원) — 헤더 포함 입력 권장`;
    };
    if (_pgDescEl) {
      root.querySelectorAll('input[name="pg-progress-paste-step"]').forEach((radio) => {
        radio.addEventListener("change", () => {
          _pgDescEl.innerHTML = _buildPasteDesc(radio.value);
        });
      });
    }

    // 총괄월별실적 붙여넣기 반영 (사번 | 장기하이캡(천원) | 실적(천원) → 즉시 Firestore 저장)
    const pasteApply = $("#btn-pg-paste-apply");
    if (pasteApply) pasteApply.addEventListener("click", async () => {
      const txt = $("#pg-paste").value;
      const m = $("#pg-paste-msg");
      const records = [];
      const unmatched = [];
      txt.split(/\r?\n/).forEach((line) => {
        const parts = line.trim().split(/[\t]+/).map((p) => p.replace(/,/g, "").trim()).filter(Boolean);
        if (parts.length < 3) return;
        const empNo   = parts[0];
        const hiCapVal = parseInt(parts[1], 10);
        const curVal   = parseInt(parts[2], 10);
        if (isNaN(hiCapVal) || isNaN(curVal)) return;
        const s = list.find((x) => x.empNo === empNo);
        if (!s) { unmatched.push(empNo); return; }
        records.push({ ...s, hiCap: hiCapVal * 1000, current: curVal * 1000 }); // 천원 입력 → 원으로 변환 저장
        // Admin 테이블 input 즉시 반영
        const hiCapInp = document.querySelector(`.pg-input[data-emp="${escapeHtml(empNo)}"][data-f="hiCap"]`);
        const curInp   = document.querySelector(`.pg-input[data-emp="${escapeHtml(empNo)}"][data-f="current"]`);
        if (hiCapInp) { hiCapInp.value = hiCapVal * 1000; }
        if (curInp)   { curInp.value   = curVal * 1000;   curInp.dispatchEvent(new Event("input")); }
      });
      if (records.length === 0) {
        if (m) { m.textContent = "❌ 매칭된 사번 없음. 탭 구분 및 사번을 확인하세요."; m.className = "pg-msg err"; }
        return;
      }
      if (m) m.textContent = "저장중...";
      pasteApply.disabled = true;
      try {
        if (typeof window.DataAPI.saveMany === "function") {
          await window.DataAPI.saveMany(records);
        } else {
          for (const r of records) await window.DataAPI.save(r);
        }
        let msg = `✅ ${records.length}명 저장 완료`;
        if (unmatched.length) msg += ` (미매칭 사번: ${unmatched.join(", ")})`;
        if (m) { m.textContent = msg; m.className = "pg-msg ok"; setTimeout(() => { m.textContent = ""; }, 6000); }
        toast(`${records.length}명 총괄월별실적 저장`, "success");
      } catch (e) {
        console.error(e);
        if (m) { m.textContent = "❌ 저장 실패: " + e.message; m.className = "pg-msg err"; }
        toast("저장 실패: " + e.message, "error");
      }
      pasteApply.disabled = false;
    });
    const pasteClear = $("#btn-pg-paste-clear");
    if (pasteClear) pasteClear.addEventListener("click", () => { $("#pg-paste").value = ""; });

    // ── 아너스목표 붙여넣기 핸들러 ──────────────────────────────
    const honorsPasteApply = $("#btn-pg-honors-paste-apply");
    if (honorsPasteApply) honorsPasteApply.addEventListener("click", async () => {
      const txt = $("#pg-honors-paste").value.trim();
      const m = $("#pg-honors-paste-msg");
      if (!txt) { if (m) { m.textContent = "❌ 붙여넣을 내용이 없습니다."; m.className = "pg-msg err"; } return; }

      const records = [];
      const unmatched = [];
      txt.split(/\r?\n/).forEach((line) => {
        const parts = line.split(/\t/).map((c) => c.replace(/,/g, "").replace(/[ ​﻿]/g, "").trim()).filter(Boolean);
        if (parts.length < 2) return;
        const empNo = parts[0];
        const honorsAmt = parseInt(parts[1], 10);
        if (!empNo || isNaN(honorsAmt)) return;
        const s = state.students.find((x) => x.empNo === empNo);
        if (!s) { unmatched.push(empNo); return; }
        records.push({ ...s, honors: honorsAmt });
      });

      if (records.length === 0) {
        if (m) { m.textContent = "❌ 매칭된 사번 없음. 탭 구분 및 사번을 확인하세요."; m.className = "pg-msg err"; }
        return;
      }
      honorsPasteApply.disabled = true;
      if (m) { m.textContent = "저장중..."; m.className = "pg-msg"; }
      try {
        if (typeof window.DataAPI.saveMany === "function") {
          await window.DataAPI.saveMany(records);
        } else {
          for (const r of records) await window.DataAPI.save(r);
        }
        let msg = `✅ ${records.length}명 아너스목표 저장 완료`;
        if (unmatched.length) msg += ` (미매칭 사번 ${unmatched.length}건 제외)`;
        if (m) { m.textContent = msg; m.className = "pg-msg ok"; setTimeout(() => { m.textContent = ""; }, 6000); }
        toast(`${records.length}명 아너스목표 저장`, "success");
      } catch (e) {
        console.error(e);
        if (m) { m.textContent = "❌ 저장 실패: " + e.message; m.className = "pg-msg err"; }
        toast("저장 실패: " + e.message, "error");
      }
      honorsPasteApply.disabled = false;
    });

    // ── 인품실적 전체 붙여넣기 핸들러 (전체 사번 기준) ────────────
    const ipumGlobalClear = $("#btn-pg-ipum-global-clear");
    if (ipumGlobalClear) ipumGlobalClear.addEventListener("click", () => { const t = $("#pg-ipum-global-paste"); if (t) t.value = ""; });

    const ipumGlobalApply = $("#btn-pg-ipum-global-apply");
    if (ipumGlobalApply) ipumGlobalApply.addEventListener("click", async () => {
      const txt = $("#pg-ipum-global-paste")?.value?.trim() || "";
      const m = $("#pg-ipum-global-msg");
      if (!txt) { if (m) { m.textContent = "❌ 붙여넣을 내용이 없습니다."; m.className = "pg-msg err"; } return; }

      const empMap = new Map();
      state.students.forEach((s) => { if (s.empNo) empMap.set(String(s.empNo).trim(), s); });

      const records = [];
      let skipped = 0;
      txt.split(/\r?\n/).forEach((line) => {
        const parts = line.trim().split(/[\t\s]+/).map((p) => p.replace(/,/g, "").trim()).filter(Boolean);
        if (parts.length < 3) return;
        const rawEmp = parts[0];
        const count  = parseInt(parts[1], 10);
        const amt    = parseInt(parts[parts.length - 1], 10);
        if (isNaN(count) || isNaN(amt)) return;
        const s = empMap.get(rawEmp);
        if (s) {
          records.push({ ...s, ipumCount: count, ipumAmt: amt });
        } else {
          skipped++;
        }
      });

      if (records.length === 0) {
        if (m) { m.textContent = `❌ 매칭된 사번 없음 (미매칭 ${skipped}건)。 사번을 확인하세요.`; m.className = "pg-msg err"; }
        toast("매칭된 사번이 없습니다.", "error");
        return;
      }

      ipumGlobalApply.disabled = true;
      if (m) { m.textContent = "저장중..."; m.className = "pg-msg"; }
      try {
        if (typeof window.DataAPI.saveMany === "function") {
          await window.DataAPI.saveMany(records);
        } else {
          for (const r of records) await window.DataAPI.save(r);
        }
        const resultMsg = `✅ ${records.length}명 저장 완료 / ${skipped}건 미매칭(버림)`;
        if (m) { m.textContent = resultMsg; m.className = "pg-msg ok"; setTimeout(() => { if (m) m.textContent = ""; }, 10000); }
        toast(`인품실적 저장 완료: ${records.length}명 성공, ${skipped}건 버림`, "success");
      } catch (e) {
        console.error(e);
        if (m) { m.textContent = "❌ 저장 실패: " + e.message; m.className = "pg-msg err"; }
        toast("저장 실패: " + e.message, "error");
      }
      ipumGlobalApply.disabled = false;
    });

    // ── 실적진도현황 붙여넣기 핸들러 (열 매핑 확인 팝업 포함) ────────
    const progressPasteApply = $("#btn-pg-progress-paste-apply");
    if (progressPasteApply) progressPasteApply.addEventListener("click", async () => {
      const txt = $("#pg-progress-paste").value.trim();
      const m = $("#pg-progress-paste-msg");
      // 기수 선택을 루프 전에 먼저 검증 — 기수 미선택 시 데이터가 엉뚱한 기수에 저장되는 것을 방지
      const _cohort = $("#pg-paste-cohort-sel")?.value || "";
      if (!_cohort) {
        if (m) { m.textContent = "❌ 기수를 먼저 선택하세요."; m.className = "pg-msg err"; }
        toast("기수를 선택한 뒤 저장하세요.", "error");
        return;
      }
      if (!txt) { if (m) { m.textContent = "❌ 붙여넣을 내용이 없습니다."; m.className = "pg-msg err"; } return; }

      const pasteStepVal = (root.querySelector('input[name="pg-progress-paste-step"]:checked') || {}).value || "1";
      const sfxOverride  = pasteStepVal === "1" ? "" : pasteStepVal;

      const lines = txt.split(/\r?\n/).filter((l) => l.trim());
      if (!lines.length) return;

      // 첫 행 헤더 자동 감지 — 행 앞뒤 탭 제거 후 분리 (Excel 복사 시 선행 탭 대응)
      const firstRow = lines[0].trim().split(/\t/).map((c) => c.trim());
      const isHeader = firstRow.some((h) => Object.prototype.hasOwnProperty.call(PG_HEADER_AUTOMAP, h));

      let dataLines, colDefs;
      if (isHeader) {
        dataLines = lines.slice(1);
        colDefs = firstRow.map((h) => ({ label: h, field: PG_HEADER_AUTOMAP[h] ?? "ignore" }));
      } else {
        dataLines = lines;
        // 열 수·선택 스텝에 따라 기본 매핑 결정
        // · step2 선택 + 18열↑: step1·step2 복합 형식 (기준실적+step1현재+step2현재)
        // · 13열: 대리점명 포함 단축 형식 (지역단·비전센터·지점·사번·대리점명·성명·위촉차월·기준실적·현재실적·달성률·순증·인품건수·인품실적)
        // · 12열↓: 직전6개월 없는 단축 형식 (대리점명 없음)
        // · 나머지: 표준 16열 형식
        const PG_SHORT_COLS = [
          "region","center","branch","empNo","name","pgMonth","pgLeader",
          "pgBase","pgCurrent","ignore","ignore","pgIpumCount","pgIpumAmt",
        ];
        let effectiveCols;
        if (pasteStepVal === "2" && firstRow.length >= 18) {
          effectiveCols = PG_STEP2_COMBINED_COLS;
        } else if (firstRow.length === 10) {
          effectiveCols = PG_STEP_UNIFIED_COLS;   // 공통 10열 (지역단~기준실적·인품건수·인품실적)
        } else if (firstRow.length === 13) {
          effectiveCols = PG_STEP1_SHORT_COLS;    // 대리점명 포함 13열
        } else if (firstRow.length <= 12) {
          effectiveCols = PG_SHORT_COLS;           // 단축 12열 이하
        } else {
          effectiveCols = PG_DEFAULT_COLS;
        }
        colDefs = firstRow.map((_, i) => ({
          label: PG_FIELD_OPTIONS.find((o) => o.value === (effectiveCols[i] || "ignore"))?.label ?? `열 ${i + 1}`,
          field: effectiveCols[i] || "ignore",
        }));
      }

      // 랜덤 3개 샘플 행 (데이터가 3개 이하면 전체, 이상이면 무작위 선택)
      const _sampleIdxs = (() => {
        const n = dataLines.length;
        if (n <= 3) return Array.from({ length: n }, (_, i) => i);
        const s = new Set();
        while (s.size < 3) s.add(Math.floor(Math.random() * n));
        return [...s].sort((a, b) => a - b);
      })();
      const sampleRows = _sampleIdxs.map((i) => dataLines[i].trim().split(/\t/).map((c) => c.trim()));

      // 복합 형식(21열) 여부: 팝업 전에 colDefs 기반으로 감지 (Step2 드롭다운 항목 제거 후에도 동작)
      const isCombinedFormat = colDefs.some((c) => c.field === "pgCurrent2");
      // 원본 열 순서 — 복합 형식에서 Step2 필드를 위치로 추출할 때 사용
      const origFM  = colDefs.map((c) => c.field);
      const parseAmt  = (v) => parseInt((v || "").replace(/,/g, "").trim(), 10) || 0;
      const getOrigAmt = (p, f) => { const i = origFM.indexOf(f); return i >= 0 ? parseAmt(p[i]) : 0; };

      // ── 열 매핑 확인 팝업 ─────────────────────────────────────────
      const fieldMapping = await openPgColMapModal(colDefs, sampleRows);
      if (!fieldMapping) return;

      // 데이터 파싱 (열 이름 기준)
      const getCol = (p, f) => { const i = fieldMapping.indexOf(f); return i >= 0 ? (p[i] || "") : ""; };
      const getAmt = (p, f) => { const i = fieldMapping.indexOf(f); return i >= 0 ? parseAmt(p[i]) : 0; };

      const updateRecords   = [];
      const newRecords      = [];
      const cohortMismatch  = []; // 다른 기수에 등록된 사번 — 저장 대상에서 제외

      dataLines.forEach((line) => {
        const p = line.trim().split(/\t/).map((c) => c.replace(/,/g, "").trim());
        const empNo = getCol(p, "empNo").replace(/[/\\\s]/g, "");
        if (!empNo) return;

        const region      = getCol(p, "region");
        const center      = getCol(p, "center");
        const branch      = getCol(p, "branch");
        const name        = getCol(p, "name");
        const pgMonth     = getCol(p, "pgMonth");
        const pgLeader    = getCol(p, "pgLeader");
        const pgPreIns    = getAmt(p, "pgPreIns");
        const pgPreConv   = getAmt(p, "pgPreConv");
        const pgPreIncome = getAmt(p, "pgPreIncome");
        const pgBase      = getAmt(p, "pgBase");
        const pgCurrent   = getAmt(p, "pgCurrent");
        const pgIpumCount = getAmt(p, "pgIpumCount");
        const pgIpumAmt   = getAmt(p, "pgIpumAmt");

        // 기수 필터: 선택한 기수와 일치하는 학생만 매칭
        const existing = state.students.find((x) => x.empNo === empNo && (!_cohort || !x.cohort || x.cohort === _cohort));
        // 동일 사번이 다른 기수에 등록되어 있으면 건너뜀 (데이터 혼재 방지)
        if (!existing && state.students.some((x) => x.empNo === empNo)) {
          cohortMismatch.push(empNo);
          return;
        }

        let pgFields, baseUpdate;
        if (isCombinedFormat) {
          // Step2 복합 형식: 원본 열 위치로 step2 필드 추출 (드롭다운에 Step2 항목 없으므로)
          pgFields = {
            pgCurrent2:   getOrigAmt(p, "pgCurrent2"),
            pgIpumCount2: getOrigAmt(p, "pgIpumCount2"),
            pgIpumAmt2:   getOrigAmt(p, "pgIpumAmt2"),
          };
          if (fieldMapping.includes("pgMonth"))  pgFields.pgMonth  = pgMonth;
          if (fieldMapping.includes("pgLeader")) pgFields.pgLeader = pgLeader;
          baseUpdate = pgBase > 0 ? { base: pgBase } : {};
        } else {
          // 단일 스텝 형식 — 매핑에 있는 열만 저장 (없는 열은 기존 값 유지)
          const sfx = sfxOverride;
          pgFields = {};
          if (fieldMapping.includes("pgBase"))       baseUpdate = pgBase > 0 ? { base: pgBase } : {};
          if (fieldMapping.includes("pgCurrent"))    pgFields[`pgCurrent${sfx}`]   = pgCurrent;
          if (fieldMapping.includes("pgIpumCount"))  pgFields[`pgIpumCount${sfx}`] = pgIpumCount;
          if (fieldMapping.includes("pgIpumAmt"))    pgFields[`pgIpumAmt${sfx}`]   = pgIpumAmt;
          if (fieldMapping.includes("pgPreIns"))     pgFields.pgPreIns    = pgPreIns;
          if (fieldMapping.includes("pgPreConv"))    pgFields.pgPreConv   = pgPreConv;
          if (fieldMapping.includes("pgPreIncome"))  pgFields.pgPreIncome = pgPreIncome;
          if (fieldMapping.includes("pgMonth"))      pgFields.pgMonth     = pgMonth;
          if (fieldMapping.includes("pgLeader"))     pgFields.pgLeader    = pgLeader;
          if (!baseUpdate) baseUpdate = {};
          if (sfx === "" && pgBase > 0 && fieldMapping.includes("pgCurrent")) {
            baseUpdate = { ...baseUpdate, current: pgCurrent };
          }
        }

        if (existing) {
          const targetUpdate = (region !== "호남지역단" && pgBase > 0) ? { target: pgBase + 50000 } : {};
          const nameUpdate   = name   ? { name }   : {};
          const regionUpdate = region ? { region } : {};
          const centerUpdate = center ? { center } : {};
          const branchUpdate = branch ? { branch } : {};
          updateRecords.push({ ...existing, ...regionUpdate, ...centerUpdate, ...branchUpdate, ...nameUpdate, ...pgFields, ...baseUpdate, ...targetUpdate });
        } else {
          const newTarget = region !== "호남지역단" && pgBase > 0 ? pgBase + 50000 : 0;
          const baseFields = isCombinedFormat
            ? { base: pgBase }
            : { base: pgBase, current: pgCurrent, ipumCount: pgIpumCount, ipumAmt: pgIpumAmt };
          newRecords.push({ region, center, branch, cohort: _cohort, empNo, name, ...pgFields, ...baseFields, target: newTarget });
        }
      });

      if (updateRecords.length === 0 && newRecords.length === 0) {
        const mismatchHint = cohortMismatch.length ? ` (${cohortMismatch.length}건은 다른 기수 학생으로 건너뜀)` : "";
        if (m) { m.textContent = `❌ 파싱된 행이 없습니다. 사원번호 열 매핑을 확인하세요.${mismatchHint}`; m.className = "pg-msg err"; }
        return;
      }

      // ── 최종 저장 확인 ────────────────────────────────────────────
      const _region = state.filter.region || "?";
      if (!await openPasteSaveConfirmModal(_region, _cohort, pasteStepVal, updateRecords, newRecords)) return;

      // 저장 — 진행 상황 실시간 표시
      if (updateRecords.length > 0) {
        progressPasteApply.disabled = true;
        const total = updateRecords.length;
        const setMsg = (txt, cls = "pg-msg") => { if (m) { m.textContent = txt; m.className = cls; } };
        setMsg(`저장중... 0 / ${total}건`);
        try {
          let saveResult;
          if (typeof window.DataAPI.saveMany === "function") {
            // 배치 저장: 청크 완료마다 진행 수 갱신
            saveResult = await window.DataAPI.saveMany(updateRecords, (done, tot) => {
              setMsg(`저장중... ${done} / ${tot}건`);
            });
          } else {
            // 폴백: 1건씩 저장하며 진행 수 표시
            let saved = 0; const errs = [];
            for (const r of updateRecords) {
              try { await window.DataAPI.save(r); saved++; }
              catch (e) { errs.push({ empNo: r.empNo, message: e.message }); }
              setMsg(`저장중... ${saved} / ${total}건`);
            }
            saveResult = { committed: saved, errors: errs };
          }

          const committed  = saveResult?.committed ?? 0;
          const saveErrors = saveResult?.errors   ?? [];
          if (saveErrors.length) {
            console.warn("[saveMany 오류 목록]", saveErrors);
            // 첫 번째 오류 메시지를 토스트로 표시
            const firstErr = saveErrors[0]?.message || "알 수 없는 오류";
            console.error("첫 번째 오류:", firstErr);
          }

          let msgTxt = committed > 0 ? `✅ ${committed}명 저장 완료` : `❌ 저장된 건 없음`;
          if (saveErrors.length) msgTxt += ` (오류 ${saveErrors.length}건: ${(saveErrors[0]?.message || "").slice(0, 40)})`;
          if (newRecords.length) msgTxt += ` · 신규 ${newRecords.length}명 팝업 확인 필요`;
          if (cohortMismatch.length) msgTxt += ` · 기수 불일치 ${cohortMismatch.length}건 제외`;
          setMsg(msgTxt, (committed === 0 || saveErrors.length) ? "pg-msg err" : "pg-msg ok");
          setTimeout(() => { if (m) m.textContent = ""; }, 15000);
          toast(`${committed}명 실적진도현황 저장`, committed > 0 ? "success" : "error");
          if (committed > 0) { const ta = $("#pg-progress-paste"); if (ta) ta.value = ""; }
          // 저장 완료 즉시 화면 갱신 (onSnapshot 보다 먼저 확정 렌더)
          if (committed > 0) renderDebounced();
        } catch (e) {
          console.error(e);
          setMsg("❌ 저장 실패: " + e.message, "pg-msg err");
          toast("저장 실패: " + e.message, "error");
          progressPasteApply.disabled = false;
          return;
        }
        progressPasteApply.disabled = false;
      }

      if (newRecords.length > 0) openPgNewStudentModal(newRecords);
    });

    // 인품 테이블 ↔ 메인 테이블 동기화 (같은 empNo 의 두 입력을 동시 반영)
    root.querySelectorAll(".pg-input[data-f2]").forEach((inp) => {
      inp.addEventListener("input", (e) => {
        const emp = e.target.dataset.emp;
        const f = e.target.dataset.f2; // "ipumCount" or "ipumAmt"
        const twin = root.querySelector(`.pg-input[data-emp="${emp}"][data-f="${f}"]`);
        if (twin) twin.value = e.target.value;
      });
    });
    root.querySelectorAll(".pg-input[data-f='ipumCount'], .pg-input[data-f='ipumAmt']").forEach((inp) => {
      inp.addEventListener("input", (e) => {
        const emp = e.target.dataset.emp;
        const f = e.target.dataset.f;
        const twin = root.querySelector(`.pg-input[data-emp="${emp}"][data-f2="${f}"]`);
        if (twin) twin.value = e.target.value;
      });
    });

    // 인품 붙여넣기 반영: "이름 건수 실적"
    const ipumPasteApply = $("#btn-pg-ipum-paste-apply");
    if (ipumPasteApply) ipumPasteApply.addEventListener("click", () => {
      const txt = $("#pg-ipum-paste").value;
      const empMap = new Map();
      list.forEach((s) => { if (s.empNo) empMap.set(String(s.empNo).trim(), s); });
      let cnt = 0, skipped = 0;
      txt.split(/\r?\n/).forEach((line) => {
        const parts = line.trim().split(/\s+/);
        if (parts.length < 3) return;
        const rawEmp = parts[0].replace(/,/g, "").trim();
        const count = parseInt(parts[1].replace(/,/g, ""), 10);
        const amt = parseInt(parts[parts.length - 1].replace(/,/g, ""), 10);
        if (isNaN(count) || isNaN(amt)) return;
        const s = empMap.get(rawEmp);
        if (s) {
          ["f2", "f"].forEach((attr) => {
            const cEl = document.querySelector(`.pg-input[data-emp="${s.empNo}"][data-${attr}="ipumCount"]`);
            const aEl = document.querySelector(`.pg-input[data-emp="${s.empNo}"][data-${attr}="ipumAmt"]`);
            if (cEl) cEl.value = count;
            if (aEl) aEl.value = amt;
          });
          cnt++;
        } else {
          skipped++;
        }
      });
      const m = $("#pg-ipum-paste-msg");
      if (m) {
        const skipPart = skipped ? ` · 사번 미일치 ${skipped}건 버림` : "";
        m.innerHTML = `✅ ${cnt}명 반영 (아래 [💾 인품 저장] 눌러야 확정)${skipPart}`;
        m.className = skipped ? "pg-msg warn" : "pg-msg ok";
        setTimeout(() => { m.textContent = ""; }, 8000);
      }
      if (cnt === 0) {
        toast("매칭되는 사번이 없습니다. 사번을 확인하세요.", "error");
      } else {
        toast(`${cnt}명 인품 데이터 반영 — [💾 인품 저장] 으로 확정`, "success");
      }
    });
    const ipumPasteClear = $("#btn-pg-ipum-paste-clear");
    if (ipumPasteClear) ipumPasteClear.addEventListener("click", () => { $("#pg-ipum-paste").value = ""; });

    // 인품 전용 저장 (ipumCount/ipumAmt 만 업데이트)
    const ipumSaveBtn = $("#btn-pg-ipum-save");
    if (ipumSaveBtn) ipumSaveBtn.addEventListener("click", async () => {
      const updates = {};
      root.querySelectorAll(".pg-input[data-f2]").forEach((inp) => {
        const emp = inp.dataset.emp;
        const f = inp.dataset.f2;
        if (!emp || !f) return;
        if (!updates[emp]) updates[emp] = {};
        updates[emp][f] = parseFloat(inp.value) || 0;
      });
      const msg = $("#pg-ipum-save-msg");
      // 변경된 것만 저장 — 0→0 은 제외해 배치 크기 축소
      const changedRecords = [];
      Object.keys(updates).forEach((emp) => {
        const s = state.students.find((x) => x.empNo === emp);
        if (!s) return;
        const newCount = Number(updates[emp].ipumCount || 0);
        const newAmt = Number(updates[emp].ipumAmt || 0);
        const oldCount = Number(s.ipumCount || 0);
        const oldAmt = Number(s.ipumAmt || 0);
        if (newCount !== oldCount || newAmt !== oldAmt) {
          changedRecords.push({ ...s, ipumCount: newCount, ipumAmt: newAmt });
        }
      });
      if (!changedRecords.length) {
        if (msg) { msg.textContent = "변경된 인품 데이터가 없습니다. (붙여넣기 반영 먼저 필요)"; msg.className = "pg-msg warn"; }
        toast("변경된 값이 없습니다.", "error");
        return;
      }
      if (msg) msg.textContent = `저장중... (${changedRecords.length}명)`;
      ipumSaveBtn.disabled = true;
      try {
        if (typeof window.DataAPI.saveMany === "function") {
          const result = await window.DataAPI.saveMany(changedRecords);
          if (result && result.errors && result.errors.length) {
            console.warn("[saveMany] 일부 실패:", result.errors);
            if (msg) { msg.textContent = `⚠️ ${result.committed || 0}건 성공 / ${result.errors.length}건 실패 — 콘솔 확인`; msg.className = "pg-msg err"; }
            toast(`일부 실패: ${result.errors.length}건`, "error");
          } else {
            if (msg) { msg.textContent = `✅ ${changedRecords.length}명 인품 저장 완료`; msg.className = "pg-msg ok"; }
            toast(`인품 ${changedRecords.length}명 저장 완료`, "success");
          }
        } else {
          for (const r of changedRecords) await window.DataAPI.save(r);
          if (msg) { msg.textContent = `✅ ${changedRecords.length}명 저장 완료`; msg.className = "pg-msg ok"; }
          toast(`인품 ${changedRecords.length}명 저장 완료`, "success");
        }
      } catch (e) {
        console.error("[ipumSave] 예외:", e);
        if (msg) { msg.textContent = "❌ 저장 실패: " + (e.message || e); msg.className = "pg-msg err"; }
        toast("저장 실패: " + (e.message || e), "error");
      }
      ipumSaveBtn.disabled = false;
    });

    // ========== 팀 배정 ==========
    // 초기 요약
    renderTeamSummary(list, root);

    // 기수 필터
    const cohortFilter = root.querySelector("#pg-team-cohort-filter");
    const updateUnassignedCnt = () => {
      const selCohort = cohortFilter ? cohortFilter.value : "";
      const cnt = [...root.querySelectorAll(".pg-team-tbl-row")].filter(tr => {
        if (selCohort && tr.dataset.cohort !== selCohort) return false;
        return tr.classList.contains("pg-team-unassigned-row") || !tr.querySelector(".pg-team-input")?.value;
      }).length;
      const el = root.querySelector("#pg-team-unassigned-cnt");
      if (el) el.textContent = cnt;
    };
    if (cohortFilter) {
      cohortFilter.addEventListener("change", () => {
        const val = cohortFilter.value;
        root.querySelectorAll(".pg-team-tbl-row").forEach(tr => {
          tr.hidden = !!(val && tr.dataset.cohort !== val);
        });
        updateUnassignedCnt();
        renderTeamSummary(list, root);
      });
    }

    // 팀 입력 실시간 반영
    root.querySelectorAll(".pg-team-input").forEach((inp) => {
      inp.addEventListener("input", () => { renderTeamSummary(list, root); updateUnassignedCnt(); });
    });

    // 자동 배정 (N 팀으로 랜덤 균등 분배)
    const autoBtn = $("#btn-pg-team-auto");
    if (autoBtn) autoBtn.addEventListener("click", () => {
      const nInp = $("#pg-team-count");
      const n = Math.max(2, Math.min(20, parseInt(nInp ? nInp.value : "8", 10) || 8));
      if (!confirm(`현재 ${list.length}명을 ${n}개 팀으로 랜덤하게 고르게 배정합니다. 기존 배정은 덮어쓰여요. 계속할까요?`)) return;
      // Fisher-Yates shuffle
      const shuffled = list.slice();
      for (let i = shuffled.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
      }
      shuffled.forEach((s, idx) => {
        const teamNo = (idx % n) + 1;
        const teamName = String(teamNo);
        const inp = root.querySelector(`.pg-team-input[data-emp="${s.empNo}"]`);
        if (inp) inp.value = teamName;
      });
      renderTeamSummary(list, root);
      const msg = $("#pg-team-msg");
      if (msg) { msg.textContent = `✅ ${n}개 팀으로 고르게 배정 완료 (저장 버튼을 눌러 확정)`; msg.className = "pg-msg ok"; setTimeout(() => { msg.textContent = ""; }, 4000); }
    });

    // 전체 초기화
    const clrBtn = $("#btn-pg-team-clear");
    if (clrBtn) clrBtn.addEventListener("click", () => {
      if (!confirm("모든 교육생의 팀 배정을 비웁니다. 저장 버튼을 눌러야 반영됩니다. 계속할까요?")) return;
      root.querySelectorAll(".pg-team-input").forEach((inp) => { inp.value = ""; });
      renderTeamSummary(list, root);
    });

    // 팀 저장
    const teamSaveBtn = $("#btn-pg-team-save");
    if (teamSaveBtn) teamSaveBtn.addEventListener("click", async () => {
      const updates = [];
      root.querySelectorAll(".pg-team-input").forEach((inp) => {
        const emp = inp.dataset.emp;
        const team = inp.value.trim();
        const s = state.students.find((x) => x.empNo === emp);
        if (!s) return;
        // 변경된 것만 저장 (네트워크 최적화)
        if ((s.team || "") !== team) updates.push({ ...s, team });
      });
      const msg = $("#pg-team-save-msg");
      if (!updates.length) { if (msg) { msg.textContent = "변경된 팀 배정이 없습니다."; msg.className = "pg-msg"; } return; }
      if (msg) msg.textContent = "저장중...";
      teamSaveBtn.disabled = true;
      try {
        if (typeof window.DataAPI.saveMany === "function") await window.DataAPI.saveMany(updates);
        else for (const r of updates) await window.DataAPI.save(r);
        if (msg) { msg.textContent = `✅ ${updates.length}명 팀 배정 저장 완료`; msg.className = "pg-msg ok"; }
        toast(`팀 배정 ${updates.length}건 저장 완료`, "success");
        renderDebounced();
      } catch (e) {
        console.error(e);
        if (msg) { msg.textContent = "❌ 저장 실패: " + e.message; msg.className = "pg-msg err"; }
        toast("저장 실패: " + e.message, "error");
      }
      teamSaveBtn.disabled = false;
    });

    // ── 기수 삭제 버튼 ────────────────────────────────────────────
    root.querySelectorAll(".btn-cohort-del").forEach((btn) => {
      btn.addEventListener("click", async (e) => {
        e.stopPropagation();
        const cohort = btn.dataset.cohort;
        const scope  = btn.dataset.scope;
        const toDelete = scope === "region"
          ? state.students.filter((s) => s.region === state.progressRegion && (s.cohort || "") === cohort)
          : state.students.filter((s) => (s.cohort || "") === cohort);
        if (!toDelete.length) { toast("해당 기수 교육생이 없습니다.", "error"); return; }
        const scopeLabel = scope === "region" ? `${state.progressRegion} ` : "전체 ";
        const deleted = await confirmAndDeleteStudents(toDelete, { label: `기수 "${cohort}기" ${scopeLabel}` });
        if (deleted) renderProgressPanel();
      });
    });
  }

  // 학생 doc 저장 시 current/ipumCount/ipumAmt 도 포함되도록 확장
  // (DataAPI.save 는 이미 record 통째를 merge 하므로 추가 필드만 있으면 그대로 저장됨)

  // 이력 항목을 폼으로 불러와 수정 모드 진입
  function editInterview(consultId) {
    const c = state.consultations.find((x) => x.id === consultId);
    if (!c) return;
    state.editingConsultId = consultId;

    const setVal = (id, v) => { const el = $("#" + id); if (el) el.value = v ?? ""; };
    setVal("iv-date", c.date || "");
    setVal("iv-seq", c.seq || "");
    setVal("iv-pct", c.pct || "");
    setVal("iv-curAct", c.curAct || "");
    setVal("iv-plan", c.plan || "");
    setVal("iv-hap", c.hap || "");
    setVal("iv-exp",    c.exp    || "");
    setVal("iv-close1", c.close1 || "");
    setVal("iv-close2", c.close2 || "");
    setVal("iv-coach", c.coach || c.content || "");

    // 상담고객 복원
    initCR(Array.isArray(c.clients) ? c.clients : []);

    // 시상 계산기 값 복원
    setVal("calc-avg", c.calcAvg || "");
    setVal("calc-base-tgt", c.calcBaseTgt || "");
    setVal("calc-tgt", c.calcTgt || "");
    if (c.calcAvg || c.calcTgt) calc();

    updateIvTitle();
    renderConsultations(); // editing 표시 갱신
    updateSaveButtonLabel();

    // 폼으로 스크롤
    const form = document.querySelector(".iv-form");
    if (form) form.scrollIntoView({ behavior: "smooth", block: "start" });
    toast(`${c.seq || ""}차 면담을 불러왔습니다. 수정 후 [저장]을 누르세요.`, "");
  }

  function cancelEditInterview() {
    state.editingConsultId = null;
    const s = state.students.find((x) => x.empNo === state.selectedEmpNo);
    clearInterviewForm();
    if (s) autoFillInterviewForm(s);
    renderConsultations();
    updateSaveButtonLabel();
    toast("수정이 취소되었습니다.", "");
  }

  function updateSaveButtonLabel() {
    const btn = $("#btn-iv-save");
    const btnTop = $("#btn-iv-save-top");
    const cancelBtn = $("#btn-iv-cancel-edit");
    const cancelBtnTop = $("#btn-iv-cancel-edit-top");
    if (!btn) return;
    if (state.editingConsultId) {
      btn.textContent = "✏️ 수정 저장";
      btn.classList.add("editing");
      if (btnTop) { btnTop.textContent = "✏️ 수정 저장"; btnTop.classList.add("editing"); }
      if (cancelBtn) cancelBtn.hidden = false;
      if (cancelBtnTop) cancelBtnTop.hidden = false;
    } else {
      btn.textContent = "💾 저장";
      btn.classList.remove("editing");
      if (btnTop) { btnTop.textContent = "💾 저장"; btnTop.classList.remove("editing"); }
      if (cancelBtn) cancelBtn.hidden = true;
      if (cancelBtnTop) cancelBtnTop.hidden = true;
    }
  }

  async function removeConsultation(id) {
    if (!confirm("이 면담 기록을 삭제하시겠습니까?")) return;
    try {
      await window.DataAPI.removeConsultation(state.selectedEmpNo, id);
      toast("삭제되었습니다.", "success");
    } catch (err) {
      console.error(err);
      toast("삭제 실패: " + err.message, "error");
    }
  }

  function escapeHtml(s) {
    return String(s ?? "").replace(/[&<>"']/g, (c) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
    }[c]));
  }

  function isPanelVisible(id) {
    const el = document.getElementById(id);
    return !!(el && !el.hidden);
  }

  // 렌더 디바운스 — trailing edge: 마지막 호출 후 150ms 뒤 1회 실행
  // (saveMany 청크 커밋마다 onSnapshot이 발동되므로 선착순이 아닌 최후 호출 기준으로 렌더해야 완전한 데이터를 반영)
  let _renderTimer = 0;
  function renderDebounced() {
    clearTimeout(_renderTimer);
    _renderTimer = setTimeout(() => { _renderTimer = 0; render(); }, 150);
  }

  function render() {
    renderHomeRanks();
    const list = filteredStudents();
    renderKPIs(list);
    renderSidebarStudentList(list);
    updateUnassignedAlert();
    // 통계 패널이 보일 때만 렌더 (숨겨진 DOM 재구성 비용 제거)
    if (isPanelVisible("dashboard-panel")) renderStats(list, "dashboard-body");
    if (isPanelVisible("progress-panel")) renderProgressPanel();
    // 학생 상세도 패널이 보일 때만 렌더
    if (state.selectedEmpNo && isPanelVisible("student-detail-panel")) renderStudentDetail();
  }

  // ========== 통계 렌더링 ==========
  function renderStats(list, targetId) {
    const body = targetId ? document.getElementById(targetId) : $("#dashboard-body") || $("#stats-body");
    const scope = $("#stats-scope");
    if (!body) return;
    const f = state.filter;
    const scopeText = [f.region, f.center, f.branch, f.cohort].filter(Boolean).join(" · ") || "전체";
    if (scope) scope.textContent = scopeText;

    if (!list.length) {
      body.innerHTML = `<div class="empty-state">조건에 맞는 교육생이 없습니다.</div>`;
      return;
    }

    // 비전센터별 인원·실적 합계
    const byCenter = {};
    list.forEach((s) => {
      const k = s.center || "(미지정)";
      if (!byCenter[k]) byCenter[k] = { count: 0, base: 0, target: 0, honors: 0 };
      byCenter[k].count++;
      byCenter[k].base += Number(s.base || 0);
      byCenter[k].target += Number(s.target || 0);
      byCenter[k].honors += Number(s.honors || 0);
    });

    // 지점별 인원
    const byBranch = {};
    list.forEach((s) => {
      const k = s.branch || "(미지정)";
      byBranch[k] = (byBranch[k] || 0) + 1;
    });

    // 기수별 인원
    const byCohort = {};
    list.forEach((s) => {
      const k = s.cohort || "(미지정)";
      byCohort[k] = (byCohort[k] || 0) + 1;
    });

    // 평균실적 상위 10명
    const top10 = list.slice().sort((a, b) => Number(b.base || 0) - Number(a.base || 0)).slice(0, 10);

    const maxCenter = Math.max(...Object.values(byCenter).map((v) => v.count), 1);
    const maxBranch = Math.max(...Object.values(byBranch), 1);
    const maxCohort = Math.max(...Object.values(byCohort), 1);
    const totalBase = list.reduce((a, s) => a + Number(s.base || 0), 0);
    const totalHonors = list.reduce((a, s) => a + Number(s.honors || 0), 0);
    const avgBase = Math.round(totalBase / list.length);
    const avgHonors = Math.round(totalHonors / list.length);

    body.innerHTML = `
      <div class="stats-grid">
        <div class="stats-card">
          <div class="stats-card-head">평균 실적</div>
          <div class="stats-big">${avgBase.toLocaleString()} <span>원</span></div>
          <div class="stats-sub">교육생 1인당</div>
        </div>
        <div class="stats-card">
          <div class="stats-card-head">평균 순증목표</div>
          <div class="stats-big">${avgHonors.toLocaleString()} <span>원</span></div>
          <div class="stats-sub">교육생 1인당</div>
        </div>
        <div class="stats-card">
          <div class="stats-card-head">비전센터 수</div>
          <div class="stats-big">${Object.keys(byCenter).length} <span>곳</span></div>
        </div>
        <div class="stats-card">
          <div class="stats-card-head">지점 수</div>
          <div class="stats-big">${Object.keys(byBranch).length} <span>곳</span></div>
        </div>
      </div>

      <div class="stats-row">
        <div class="stats-block">
          <h3>비전센터별 분포</h3>
          ${Object.entries(byCenter)
            .sort((a, b) => b[1].count - a[1].count)
            .map(([name, v]) => `
              <div class="stats-bar-row">
                <div class="sbr-label">${escapeHtml(name)}</div>
                <div class="sbr-track">
                  <div class="sbr-fill" style="width:${(v.count / maxCenter * 100).toFixed(1)}%"></div>
                </div>
                <div class="sbr-val">${v.count}명</div>
              </div>
            `).join("")}
        </div>

        <div class="stats-block">
          <h3>기수별 분포</h3>
          ${Object.entries(byCohort)
            .sort((a, b) => (a[0] || "").localeCompare(b[0] || ""))
            .map(([name, cnt]) => `
              <div class="stats-bar-row">
                <div class="sbr-label">${escapeHtml(name)}</div>
                <div class="sbr-track">
                  <div class="sbr-fill blue" style="width:${(cnt / maxCohort * 100).toFixed(1)}%"></div>
                </div>
                <div class="sbr-val">${cnt}명</div>
              </div>
            `).join("")}
        </div>
      </div>

      <div class="stats-block">
        <h3>지점별 분포 (${Object.keys(byBranch).length}개 지점)</h3>
        ${Object.entries(byBranch)
          .sort((a, b) => b[1] - a[1])
          .map(([name, cnt]) => `
            <div class="stats-bar-row">
              <div class="sbr-label">${escapeHtml(name)}</div>
              <div class="sbr-track">
                <div class="sbr-fill amber" style="width:${(cnt / maxBranch * 100).toFixed(1)}%"></div>
              </div>
              <div class="sbr-val">${cnt}명</div>
            </div>
          `).join("")}
      </div>

      <div class="stats-block">
        <h3>기준실적 상위 ${top10.length}명</h3>
        <table class="stats-table">
          <thead><tr>
            <th>#</th><th>이름</th><th>지점</th><th>기수</th><th class="r">기준실적</th><th class="r">순증목표</th>
          </tr></thead>
          <tbody>
            ${top10.map((s, i) => `
              <tr>
                <td>${i + 1}</td>
                <td><strong>${escapeHtml(s.name || "")}</strong></td>
                <td>${escapeHtml(s.branch || "")}</td>
                <td>${escapeHtml(s.cohort || "")}</td>
                <td class="r">${Number(s.base || 0).toLocaleString()}</td>
                <td class="r">${Number(s.honors || 0).toLocaleString()}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    `;
  }

  function getMasterTargetDefault(region) {
    try {
      const stored = JSON.parse(localStorage.getItem(LS_DEFAULTS_KEY) || "{}");
      return stored[region] !== undefined ? stored[region] : DEFAULT_MASTER_TARGET;
    } catch (e) { return DEFAULT_MASTER_TARGET; }
  }

  function renderMasterTargetSettings() {
    const container = document.getElementById("settings-default-targets");
    if (!container) return;
    const regions = [...new Set(state.students.map((s) => s.region).filter((r) => r && r.endsWith("지역단")))].sort();
    if (!regions.length) {
      container.innerHTML = `<p class="settings-desc" style="color:#999;">등록된 교육생이 없어 지역단 목록을 불러올 수 없습니다.</p>`;
      return;
    }
    let stored = {};
    try { stored = JSON.parse(localStorage.getItem(LS_DEFAULTS_KEY) || "{}"); } catch (e) {}
    const zeroTargetCount = state.students.filter((s) => !s.target || Number(s.target) === 0).length;
    const nonHonamCount = state.students.filter((s) => s.region && s.region !== "호남지역단").length;
    container.innerHTML = `
      <div class="settings-defaults-grid">${regions.map((r) => {
        const val = stored[r] !== undefined ? stored[r] : DEFAULT_MASTER_TARGET;
        return `<div class="settings-defaults-item">
          <span class="settings-defaults-label">${escapeHtml(r)}</span>
          <span style="display:flex;align-items:center;gap:4px;flex-shrink:0;">
            <input type="number" class="settings-target-input" data-region="${escapeHtml(r)}" value="${val}" min="0" step="1000" style="width:100px;text-align:right;"> 원
          </span>
        </div>`;
      }).join("")}</div>
      <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
        <button class="btn-primary small" id="btn-save-default-targets">💾 지역단별 기본값 저장</button>
        <button class="btn-outline small" id="btn-bulk-set-target">🔄 마스터목표 일괄 적용 (평균실적+20만원, ${zeroTargetCount}명 대상)</button>
        <button class="btn-outline small" id="btn-bulk-set-target-nonhonam">📌 호남 외 마스터목표 (기준실적+5만원, ${nonHonamCount}명)</button>
        <span id="settings-default-targets-msg" class="pg-msg"></span>
      </div>
    `;
    document.getElementById("btn-save-default-targets")?.addEventListener("click", () => {
      const newDefaults = {};
      container.querySelectorAll(".settings-target-input").forEach((inp) => {
        const r = inp.dataset.region;
        const v = parseInt(inp.value, 10);
        if (r && !isNaN(v)) newDefaults[r] = v;
      });
      localStorage.setItem(LS_DEFAULTS_KEY, JSON.stringify(newDefaults));
      const msg = document.getElementById("settings-default-targets-msg");
      if (msg) { msg.textContent = "✅ 저장됨"; msg.className = "pg-msg ok"; setTimeout(() => { msg.textContent = ""; }, 3000); }
      toast("기본값 저장 완료", "success");
    });

    document.getElementById("btn-bulk-set-target")?.addEventListener("click", async () => {
      const targets = state.students.filter((s) => !s.target || Number(s.target) === 0);
      if (!targets.length) { toast("마스터목표가 0인 교육생이 없습니다.", ""); return; }
      if (!confirm(`마스터목표가 0인 교육생 ${targets.length}명에게\n평균실적 + 200,000원을 일괄 적용합니다.\n계속하시겠습니까?`)) return;

      const btn = document.getElementById("btn-bulk-set-target");
      const msg = document.getElementById("settings-default-targets-msg");
      if (btn) btn.disabled = true;
      if (msg) { msg.textContent = "저장중..."; msg.className = "pg-msg"; }

      const updated = targets.map((s) => ({ ...s, target: Number(s.base || 0) + DEFAULT_MASTER_TARGET }));
      try {
        if (typeof window.DataAPI.saveMany === "function") {
          await window.DataAPI.saveMany(updated);
        } else {
          for (const r of updated) await window.DataAPI.save(r);
        }
        if (msg) { msg.textContent = `✅ ${updated.length}명 업데이트 완료`; msg.className = "pg-msg ok"; setTimeout(() => { msg.textContent = ""; }, 5000); }
        toast(`마스터목표 ${updated.length}명 일괄 적용 완료`, "success");
        renderMasterTargetSettings();
      } catch (e) {
        console.error(e);
        if (msg) { msg.textContent = "❌ 저장 실패: " + e.message; msg.className = "pg-msg err"; }
        toast("저장 실패: " + e.message, "error");
        if (btn) btn.disabled = false;
      }
    });

    document.getElementById("btn-bulk-set-target-nonhonam")?.addEventListener("click", async () => {
      const targets = state.students.filter((s) => s.region && s.region !== "호남지역단");
      if (!targets.length) { toast("호남 외 교육생이 없습니다.", ""); return; }
      if (!confirm(`호남지역단 제외 교육생 ${targets.length}명의 마스터목표를\n기준실적 + 50,000원으로 일괄 적용합니다.\n계속하시겠습니까?`)) return;

      const btn = document.getElementById("btn-bulk-set-target-nonhonam");
      const msg = document.getElementById("settings-default-targets-msg");
      if (btn) btn.disabled = true;
      if (msg) { msg.textContent = "저장중..."; msg.className = "pg-msg"; }

      const updated = targets.map((s) => {
        const { base } = getProgressStat(s);
        return { ...s, target: base + 50000 };
      });
      try {
        if (typeof window.DataAPI.saveMany === "function") {
          await window.DataAPI.saveMany(updated);
        } else {
          for (const r of updated) await window.DataAPI.save(r);
        }
        if (msg) { msg.textContent = `✅ ${updated.length}명 업데이트 완료`; msg.className = "pg-msg ok"; setTimeout(() => { msg.textContent = ""; }, 5000); }
        toast(`호남 외 마스터목표 ${updated.length}명 적용 완료`, "success");
        renderMasterTargetSettings();
      } catch (e) {
        console.error(e);
        if (msg) { msg.textContent = "❌ 저장 실패: " + e.message; msg.className = "pg-msg err"; }
        toast("저장 실패: " + e.message, "error");
        if (btn) btn.disabled = false;
      }
    });
  }

  // ========== 목표 금액 설정 모달 ==========
  function _tgMakeKey(region, cohort, step) {
    return `TG:${region}:${cohort}:${step}`;
  }

  function _tgGetStored() {
    try { return JSON.parse(localStorage.getItem(LS_TARGET_GOALS_KEY) || "{}"); } catch { return {}; }
  }

  function _tgFilteredStudents() {
    const region = document.getElementById("tg-region")?.value || "";
    const cohort = document.getElementById("tg-cohort")?.value || "";
    return state.students.filter((s) => {
      if (region && s.region !== region) return false;
      if (cohort) {
        const sc = String(s.cohort || "").replace(/기$/, "");
        if (sc !== cohort) return false;
      }
      return true;
    }).slice().sort((a, b) => (a.branch || "").localeCompare(b.branch || "") || (a.name || "").localeCompare(b.name || ""));
  }

  function _tgRenderStudentList() {
    const list = _tgFilteredStudents();
    const container = document.getElementById("tg-student-list");
    const countLabel = document.getElementById("tg-count-label");
    if (countLabel) countLabel.textContent = `${list.length}명`;
    if (!container) return;
    if (!list.length) {
      container.innerHTML = `<p style="color:#999;font-size:13px;padding:12px 0;">선택한 조건에 해당하는 교육생이 없습니다.</p>`;
      return;
    }
    container.innerHTML = `
      <p style="font-size:12px;color:#6b7280;margin:0 0 6px;">
        ※ <strong>증가액</strong>을 입력하세요. 저장 시 <strong>기준실적 + 증가액 = 마스터목표</strong>로 계산됩니다.
      </p>
      <table style="width:100%;border-collapse:collapse;font-size:13px;">
        <thead>
          <tr style="background:#f3f4f6;border-bottom:2px solid #d1d5db;">
            <th style="padding:7px 8px;text-align:left;font-weight:600;">이름</th>
            <th style="padding:7px 8px;text-align:left;font-weight:600;">지점</th>
            <th style="padding:7px 8px;text-align:right;font-weight:600;">기준실적</th>
            <th style="padding:7px 8px;text-align:right;font-weight:600;">현재 증가액</th>
            <th style="padding:7px 8px;text-align:right;font-weight:600;">증가액 입력 (원)</th>
            <th style="padding:7px 8px;text-align:right;font-weight:600;color:#2563eb;">저장될 목표</th>
          </tr>
        </thead>
        <tbody>
          ${list.map((s) => {
            const base    = Number(s.base   || 0);
            const target  = Number(s.target || 0);
            const curIncr = target > 0 ? target - base : 0; // 현재 증가액
            return `
            <tr style="border-bottom:1px solid #e5e7eb;">
              <td style="padding:6px 8px;">${escapeHtml(s.name || "")}</td>
              <td style="padding:6px 8px;color:#6b7280;">${escapeHtml(s.branch || "")}</td>
              <td style="padding:6px 8px;text-align:right;">${base.toLocaleString()}</td>
              <td style="padding:6px 8px;text-align:right;color:#9ca3af;">${curIncr > 0 ? "+" + curIncr.toLocaleString() : "-"}</td>
              <td style="padding:6px 8px;text-align:right;">
                <input type="number" class="tg-target-input"
                  data-id="${escapeHtml(s.id)}" data-base="${base}"
                  value="${Math.max(0, curIncr)}" min="0" step="10000"
                  style="width:110px;text-align:right;padding:4px 6px;border:1px solid #d1d5db;border-radius:5px;font-size:13px;"
                  oninput="(function(el){const p=el.closest('tr');const lbl=p?.querySelector('.tg-calc-lbl');if(lbl){const b=parseInt(el.dataset.base||0);const v=parseInt(el.value||0);lbl.textContent=(isNaN(v)||v<0)?'-':(b+v).toLocaleString()+'원';}})(this)">
              </td>
              <td style="padding:6px 8px;text-align:right;color:#2563eb;font-weight:600;">
                <span class="tg-calc-lbl">${curIncr > 0 ? (base + curIncr).toLocaleString() + "원" : (base > 0 ? base.toLocaleString() + "원" : "-")}</span>
              </td>
            </tr>`;
          }).join("")}
        </tbody>
      </table>
    `;
  }

  function openTargetGoalsModal() {
    const regionSel = document.getElementById("tg-region");
    const cohortSel = document.getElementById("tg-cohort");
    const stepSel   = document.getElementById("tg-step");
    const regions = [...new Set(state.students.map((s) => s.region).filter((r) => r && (r.endsWith("지역단") || r.endsWith("사업부"))))].sort();
    if (!regions.length) { toast("등록된 교육생이 없습니다.", "error"); return; }
    regionSel.innerHTML = regions.map((r) => `<option value="${escapeHtml(r)}">${escapeHtml(r)}</option>`).join("");

    // 현재 필터 값으로 초기화
    const curRegion = state.filter.region || state.progressRegion || regions[0];
    if (curRegion && regions.includes(curRegion)) regionSel.value = curRegion;
    if (state.filter.cohort && cohortSel) {
      const c = String(state.filter.cohort).replace(/기$/, "");
      cohortSel.value = c || "";
    }
    if (state.progressStep && stepSel) stepSel.value = state.progressStep;

    _tgRenderStudentList();
    [regionSel, cohortSel, stepSel].forEach((sel) => {
      if (sel) sel.onchange = () => _tgRenderStudentList();
    });

    const bulkBtn = document.getElementById("btn-tg-bulk-apply");
    if (bulkBtn) {
      bulkBtn.onclick = () => {
        const val = parseInt(document.getElementById("tg-bulk-val")?.value || "0", 10);
        if (isNaN(val) || val < 0) { toast("올바른 금액을 입력하세요.", "error"); return; }
        document.querySelectorAll(".tg-target-input").forEach((inp) => { inp.value = val; });
      };
    }

    const doSave = async () => {
      const inputs = [...document.querySelectorAll(".tg-target-input")];
      if (!inputs.length) { toast("저장할 데이터가 없습니다.", ""); return; }
      const saveBtn  = document.getElementById("btn-tg-save");
      const saveBtn2 = document.getElementById("btn-tg-save2");
      const saveMsg  = document.getElementById("tg-save-msg");
      if (saveBtn)  saveBtn.disabled  = true;
      if (saveBtn2) saveBtn2.disabled = true;
      if (saveMsg)  { saveMsg.textContent = "저장 중..."; saveMsg.className = "pg-msg"; }

      // 변경된 학생 수집 — 입력값은 증가액, 저장값 = base + 증가액
      const updated = [];
      inputs.forEach((inp) => {
        const id    = inp.dataset.id;
        const incr  = parseInt(inp.value, 10);
        const s     = state.students.find((x) => x.id === id);
        if (!s || isNaN(incr) || incr < 0) return;
        const base      = Number(s.base || 0);
        const newTarget = base + incr;
        if (newTarget !== Number(s.target || 0)) {
          updated.push({ ...s, target: newTarget });
        }
      });

      if (!updated.length) {
        if (saveMsg) { saveMsg.textContent = "변경된 항목이 없습니다."; saveMsg.className = "pg-msg"; setTimeout(() => { saveMsg.textContent = ""; }, 2500); }
        if (saveBtn)  saveBtn.disabled  = false;
        if (saveBtn2) saveBtn2.disabled = false;
        return;
      }

      try {
        if (typeof window.DataAPI?.saveMany === "function") {
          await window.DataAPI.saveMany(updated);
        } else {
          for (const r of updated) await window.DataAPI.save(r);
        }
        // state 업데이트
        updated.forEach((u) => {
          const idx = state.students.findIndex((x) => x.id === u.id);
          if (idx >= 0) state.students[idx] = u;
        });
        if (saveMsg) { saveMsg.textContent = `✅ ${updated.length}명 저장 완료`; saveMsg.className = "pg-msg ok"; setTimeout(() => { saveMsg.textContent = ""; }, 4000); }
        toast(`목표 금액 ${updated.length}명 저장 완료`, "success");
        _tgRenderStudentList(); // 목록 갱신 (현재 목표 열 업데이트)
      } catch (e) {
        console.error("[TargetGoals] 저장 실패:", e);
        if (saveMsg) { saveMsg.textContent = "❌ 저장 실패: " + e.message; saveMsg.className = "pg-msg err"; }
        toast("저장 실패: " + e.message, "error");
      } finally {
        if (saveBtn)  saveBtn.disabled  = false;
        if (saveBtn2) saveBtn2.disabled = false;
      }
    };

    const closeTgModal = () => closeModal("#modal-target-goals");
    document.getElementById("btn-tg-save").onclick   = doSave;
    document.getElementById("btn-tg-save2").onclick  = doSave;
    document.getElementById("btn-tg-close").onclick  = closeTgModal;
    document.getElementById("btn-tg-close2").onclick = closeTgModal;
    document.querySelector("#modal-target-goals .modal-backdrop").onclick = closeTgModal;

    openModal("#modal-target-goals");
  }

  // ========== 폼 ==========
  function resetForm() {
    state.form = { region: "", center: "", branch: "" };
    ["form-empno","form-name","form-phone","form-base","form-honors"].forEach((id) => $("#" + id).value = "");
    const ft = $("#form-target");
    if (ft) {
      ft.value = getMasterTargetDefault(state.filter.region || DEFAULT_REGION);
      ft.removeAttribute("readonly");
    }
    $("#form-cohort").value = "";
    $("#form-bulk").value = "";
    const _bulkCohortSel = document.getElementById("bulk-cohort-sel");
    if (_bulkCohortSel) _bulkCohortSel.value = "";
    state._bulkCohort = "";
    editingEmpNo = null;
    state.formTgtAddAmount = null;
    const _delBtn = document.getElementById("btn-modal-del-student");
    if (_delBtn) _delBtn.hidden = true;
    syncOrgLabels();
  }

  function openEditForm(empNo) {
    const s = state.students.find((x) => x.empNo === empNo);
    if (!s) return;
    state.form = { region: s.region || "", center: s.center || "", branch: s.branch || "" };
    $("#form-empno").value = s.empNo;
    $("#form-name").value = s.name || "";
    $("#form-phone").value = s.phone || "";
    $("#form-base").value   = s.base || "";
    const computedTarget = (s.region !== "호남지역단" && !Number(s.target))
      ? getProgressStat(s).base + 50000
      : Number(s.target) || "";
    $("#form-target").value = computedTarget;
    $("#form-honors").value = s.honors || "";
    $("#form-cohort").value = s.cohort || "";
    const teamNum = parseInt(s.team, 10);
    const formTeamEl = $("#form-team");
    if (formTeamEl) formTeamEl.value = teamNum > 0 ? teamNum : "";
    editingEmpNo = s.empNo;
    state.formTgtAddAmount = null;
    const ft = $("#form-target"); if (ft) ft.removeAttribute("readonly");
    syncOrgLabels();
    switchTab("single");
    openModal("#modal-add");
    // 수정 모드: 삭제 버튼 표시 및 핸들러 바인딩
    const _delBtn = document.getElementById("btn-modal-del-student");
    if (_delBtn) {
      _delBtn.hidden = false;
      _delBtn.onclick = async () => {
        const _s = state.students.find((x) => x.empNo === empNo);
        const _label = _s ? `${_s.name}(${empNo})` : empNo;
        if (!confirm(`"${_label}" 교육생을 완전히 삭제하시겠습니까?\n삭제된 데이터는 복구할 수 없습니다.`)) return;
        await window.DataAPI.removeStudentWithConsultations(empNo);
        closeModal("#modal-add");
        toast(`${_label} 삭제되었습니다.`, "success");
      };
    }
  }

  function switchTab(name) {
    $$(".tab").forEach((t) => t.classList.toggle("active", t.dataset.tab === name));
    $$(".tab-panel").forEach((p) => (p.hidden = p.dataset.panel !== name));
  }

  function buildSingleRecord() {
    const empNo = $("#form-empno").value.trim();
    if (!empNo) return null;
    const region = state.form.region;
    const base   = Number($("#form-base").value   || 0);
    let   target = Number($("#form-target").value || 0);
    if (!target && region !== "호남지역단" && base > 0) target = base + 50000;
    const teamInput = parseInt($("#form-team")?.value, 10);
    const existingStudent = state.students.find((x) => x.empNo === empNo);
    const existingTeam = existingStudent?.team || "";
    return {
      region,
      center: state.form.center,
      branch: state.form.branch,
      cohort: $("#form-cohort").value,
      empNo,
      name: $("#form-name").value.trim(),
      phone: $("#form-phone").value.trim(),
      base,
      target,
      honors: Number($("#form-honors").value || 0),
      team: teamInput > 0 ? String(teamInput) : existingTeam
    };
  }

  // 수정 모드(기존 사번 편집) 시 중복 확인 건너뛰기 위한 플래그
  let editingEmpNo = null;

  async function saveSingle() {
    const rec = buildSingleRecord();
    if (!rec) { toast("사번을 입력하세요.", "error"); return false; }
    // 신규 등록인데 동일 사번이 이미 있으면 중복 확인 모달
    if (editingEmpNo !== rec.empNo) {
      const existing = state.students.find((s) => s.empNo === rec.empNo);
      if (existing) {
        const confirmed = await confirmDuplicate(existing, rec);
        if (!confirmed) {
          toast("저장이 취소되었습니다.", "");
          return false;
        }
      }
    }
    await window.DataAPI.save(rec);
    return true;
  }

  // ========== 중복 사번 확인 모달 ==========
  function confirmDuplicate(oldRec, newRec) {
    return new Promise((resolve) => {
      $("#dup-msg").textContent =
        `사번 ${newRec.empNo} 은(는) 이미 등록되어 있습니다. 기존 값을 덮어쓰시겠습니까?`;
      $("#dup-old").innerHTML = renderDupTable(oldRec, newRec, "old");
      $("#dup-new").innerHTML = renderDupTable(newRec, oldRec, "new");
      openModal("#modal-duplicate");

      const overwriteBtn = $("#btn-dup-overwrite");
      const cleanup = () => {
        overwriteBtn.removeEventListener("click", onOverwrite);
        $$("#modal-duplicate [data-close]").forEach((el) => el.removeEventListener("click", onCancel));
      };
      const onOverwrite = () => { cleanup(); closeModal("#modal-duplicate"); resolve(true); };
      const onCancel = () => { cleanup(); closeModal("#modal-duplicate"); resolve(false); };

      overwriteBtn.addEventListener("click", onOverwrite);
      $$("#modal-duplicate [data-close]").forEach((el) => el.addEventListener("click", onCancel));
    });
  }

  function renderDupTable(rec, other, side) {
    const fields = [
      ["지역단", "region"], ["비전센터", "center"], ["지점", "branch"],
      ["기수", "cohort"], ["이름", "name"], ["연락처", "phone"],
      ["기준실적", "base"], ["마스터목표", "target"], ["아너스목표", "honors"]
    ];
    return fields.map(([label, key]) => {
      const v = rec[key];
      const ov = other ? other[key] : "";
      const isNum = ["base", "target", "honors"].includes(key);
      const display = v === undefined || v === null || v === ""
        ? "<span class='muted'>(비어있음)</span>"
        : escapeHtml(isNum ? Number(v || 0).toLocaleString() : String(v));
      const changed = String(v ?? "") !== String(ov ?? "");
      return `<tr class="${changed ? "diff" : ""}"><th>${label}</th><td>${display}</td></tr>`;
    }).join("");
  }

  function setBulkProgress(text, type) {
    const el = $("#bulk-progress");
    if (!el) return;
    if (!text) { el.hidden = true; el.textContent = ""; el.className = "bulk-progress"; return; }
    el.hidden = false;
    el.textContent = text;
    el.className = "bulk-progress" + (type ? " " + type : "");
  }

  function toNum(v) {
    if (v === undefined || v === null || v === "") return 0;
    const n = Number(String(v).replace(/[,\s]/g, ""));
    return Number.isFinite(n) ? n : 0;
  }

  // ── 일괄입력 공통 헬퍼 ──────────────────────────────────────────
  const _bulkParseAmt = (v) => parseInt((v || "").replace(/[,%\s]/g, ""), 10) || 0;
  const _bulkInferCenter = (region, rawCenter, branch) => {
    if (rawCenter) return rawCenter;
    const nb = (b) => (b || "").replace(/\s/g, "");  // 공백 제거 후 비교
    const branchNorm = nb(branch);
    if (!branchNorm) return "";
    // 1순위: 같은 지역단 + 지점명 일치 (공백 무시)
    let hint = state.students.find((s) => nb(s.branch) === branchNorm && s.region === region && s.center);
    if (hint) return hint.center;
    // 2순위: 지점명만으로 전체 학생에서 역추적 (지역단 무관 — 지점명은 회사 내 고유)
    hint = state.students.find((s) => nb(s.branch) === branchNorm && s.center);
    return hint ? hint.center : "";
  };

  // 파싱 전용 함수 — preview와 saveBulk 공용
  // colOverride: { base, current, pgIpumCount, pgIpumAmt } → 열 인덱스 재지정
  function parseBulkRecords(raw, colOverride = {}) {
    const allLines = raw.split(/\r?\n/).filter((l) => l.trim());
    const records = [];
    const parseErrors = [];
    let isFormatA = false;
    let sampleCols = [];
    let detectedColShift = 0;
    let isHeaderBased = false;

    // BULK_HEADER_MAP 기반 헤더 자동감지 — 첫 행에 3개 이상 매핑 가능 필드가 있으면 헤더 행으로 처리
    let headerFields = null;
    if (allLines.length > 0) {
      const firstCols = (allLines[0].includes("\t") ? allLines[0].split("\t") : allLines[0].split(",")).map((c) => c.trim());
      const mapped = firstCols.map((h) => BULK_HEADER_MAP[h] || null);
      if (mapped.filter(Boolean).length >= 3) {
        isHeaderBased = true;
        headerFields = mapped;
      }
    }

    if (isHeaderBased && headerFields) {
      allLines.slice(1).forEach((line, lineIdx) => {
        const cols = (line.includes("\t") ? line.split("\t") : line.split(","))
          .map((c) => (c || "").trim().replace(/^"(.*)"$/, "$1"));
        const rec = {};
        headerFields.forEach((field, j) => { if (field) rec[field] = cols[j] || ""; });
        const empNo = (rec.empNo || "").replace(/[\s\/\\]/g, "");
        if (!empNo) { parseErrors.push(`${lineIdx + 2}행 사번 누락`); return; }
        let cohort = rec.cohort || state._bulkCohort || state.filter.cohort || "";
        if (cohort && /^\d+$/.test(cohort)) cohort = cohort + "기";
        const region = rec.region || state.filter.region || "";
        const existSt = state.students.find((s) => s.empNo === empNo);
        const center = rec.center || existSt?.center || _bulkInferCenter(region, "", rec.branch || "");
        const base = _bulkParseAmt(rec.base);
        records.push({
          region, center,
          branch: rec.branch || "",
          cohort, empNo,
          name: rec.name || "",
          phone: rec.phone || existSt?.phone || "",
          base,
          target: _bulkParseAmt(rec.target) || (existSt?.target ?? 0),
          honors: _bulkParseAmt(rec.honors) || (existSt?.honors ?? 0),
          tenureMonths: _bulkParseAmt(rec.tenureMonths) || 0,
          current: Number(existSt?.current || 0),
          pgIpumCount: 0, pgIpumAmt: 0,
          team: existSt?.team || "",
          _isNew: !existSt,
        });
      });
      return { records, parseErrors, isFormatA: false, sampleCols: [], colShift: 0, isHeaderBased: true };
    }

    const lines = allLines;
    lines.forEach((line, i) => {
      const cols = line.includes("\t") ? line.split("\t") : line.split(",");
      const cl = cols.map((c) => (c || "").trim().replace(/^"(.*)"$/, "$1"));
      if (cl.length >= 16) {
        if (cl[0] === "지역단" || cl[3] === "사원번호" || cl[3] === "사번") return;
        const colShift = (!cl[18]?.includes('%') && cl[19]?.includes('%')) ? 1 : 0;
        if (!isFormatA) { isFormatA = true; detectedColShift = colShift; sampleCols = cl; }

        const baseIdx        = colOverride.base        ?? (16 + colShift);
        const currentIdx     = colOverride.current     ?? (17 + colShift);
        const pgIpumCountIdx = colOverride.pgIpumCount ?? (20 + colShift);
        const pgIpumAmtIdx   = colOverride.pgIpumAmt   ?? (21 + colShift);

        const region   = cl[0] || state.filter.region || "";
        const rawCtr   = cl[1];
        const branch   = cl[2];
        const empNo    = cl[3].replace(/[\s\/\\]/g, "");
        const name     = cl[4 + colShift];
        const tenureM  = cl[5 + colShift];
        const pgLeader = cl[6 + colShift];
        const pgPreIns    = _bulkParseAmt(cl[7 + colShift]);
        const pgPreConv   = _bulkParseAmt(cl[8 + colShift]);
        const pgPreIncome = _bulkParseAmt(cl[9 + colShift]);
        const base        = _bulkParseAmt(cl[baseIdx]);
        const current     = _bulkParseAmt(cl[currentIdx]);
        const pgIpumCount = _bulkParseAmt(cl[pgIpumCountIdx] || "0");
        const pgIpumAmt   = _bulkParseAmt(cl[pgIpumAmtIdx]   || "0");
        if (!empNo) { parseErrors.push(`${i + 1}행 사번 누락`); return; }
        // 기수: 일괄등록 폼 선택값 우선, 없으면 사이드바 필터 기준
        const cohort  = state._bulkCohort || state.filter.cohort || "";
        const existSt = state.students.find((s) => s.empNo === empNo);
        const center  = rawCtr || existSt?.center || _bulkInferCenter(region, "", branch);
        const phone   = existSt?.phone  || "";
        const honors  = existSt?.honors ?? 0;
        const target  = existSt?.target ?? (region !== "호남지역단" && base > 0 ? base + 50000 : 0);
        records.push({
          region, center, branch, cohort, empNo, name, phone, honors,
          base, target, current,
          tenureMonths: toNum(tenureM),
          pgLeader, pgPreIns, pgPreConv, pgPreIncome,
          pgIpumCount, pgIpumAmt,
          team: existSt?.team || "",
          _isNew: !existSt,
        });
      } else {
        let [region, rawCenter, branch, cohort, empNo, name, phone, base, target, honors, tenureMonths] = cl;
        if (!region) region = state.filter.region || "";
        if (cohort && /^\d+$/.test(cohort)) cohort = cohort + "기";
        if (region === "지역단" && rawCenter === "비전센터" && branch === "지점") return;
        if (!empNo) { parseErrors.push(`${i + 1}행 사번 누락`); return; }
        const existSt = state.students.find((s) => s.empNo === empNo.replace(/[\s\/\\]/g, ""));
        const center  = rawCenter || existSt?.center || _bulkInferCenter(region, "", branch);
        const _baseNum = toNum(base);
        const rec = {
          region, center, branch, cohort,
          empNo: empNo.replace(/[\s\/\\]/g, ""),
          name, phone,
          base: _baseNum, target: toNum(target), honors: toNum(honors),
          team: existSt?.team || "",
          _isNew: !existSt,
        };
        if (tenureMonths) rec.tenureMonths = toNum(tenureMonths);
        records.push(rec);
      }
    });

    return { records, parseErrors, isFormatA, sampleCols, colShift: detectedColShift, isHeaderBased: false };
  }

  // ── 미리보기 오버레이 ──────────────────────────────────────────
  function openBulkPreview() {
    const raw = $("#form-bulk").value.trim();
    if (!raw) { toast("미리보기할 내용이 없습니다.", "error"); return; }
    if (!state._bulkColOverride) state._bulkColOverride = {};
    state._bulkPreviewRaw = raw;  // 헤더 select 변경 시 재렌더에 사용
    _renderBulkPreview(raw);
    document.getElementById("bulk-preview-overlay").hidden = false;
  }

  function _renderBulkPreview(raw) {
    const colOverride = state._bulkColOverride || {};
    const { records, parseErrors, isFormatA, sampleCols, colShift, isHeaderBased } = parseBulkRecords(raw, colOverride);

    const remapBar = document.getElementById("bulk-remap-bar");
    if (remapBar) remapBar.hidden = true;

    const tblWrap = document.getElementById("bulk-preview-tbl-wrap");
    if (!records.length) {
      const err = parseErrors[0] || "데이터 없음";
      tblWrap.innerHTML = `<div class="pg-empty">파싱된 행이 없습니다. ${escapeHtml(err)}</div>`;
    } else if (!isFormatA) {
      // ── 등록 형식 (헤더 기반 또는 11열 기본) — 인라인 편집 가능 표 ──
      const inp = (field, val, type = "text", extra = "") =>
        `<input class="bulk-inp${type === "number" ? " r" : ""}" data-field="${field}" type="${type}" value="${escapeHtml(String(val ?? ""))}" ${extra}>`;
      const rows = records.map((r, i) => {
        const badge = r._isNew
          ? `<span class="bulk-badge-new">신규</span>`
          : `<span class="bulk-badge-exist">기존</span>`;
        const ctrWarn = !r.center ? ' title="비전센터 미확인 — 직접 입력하세요"' : "";
        return `<tr class="bulk-editable-row" data-row="${i}">
          <td style="text-align:center;color:#999">${i + 1}</td>
          <td style="text-align:center">${badge}</td>
          <td>${inp("region", r.region || "", "text", 'style="width:90px"')}</td>
          <td>${inp("center", r.center || "", "text", `style="width:110px"${ctrWarn}`)}</td>
          <td>${inp("branch", r.branch || "", "text", 'style="width:90px"')}</td>
          <td>${inp("cohort", r.cohort || "", "text", 'style="width:44px"')}</td>
          <td>${inp("empNo",  r.empNo  || "", "text", 'style="width:68px"')}</td>
          <td>${inp("name",   r.name   || "", "text", 'style="width:64px"')}</td>
          <td>${inp("phone",  r.phone  || "", "text", 'style="width:110px"')}</td>
          <td>${inp("base",   r.base   || 0,  "number", 'style="width:80px"')}</td>
          <td>${inp("target", r.target || 0,  "number", 'style="width:80px"')}</td>
          <td>${inp("honors", r.honors || 0,  "number", 'style="width:80px"')}</td>
          <td>${inp("tenureMonths", r.tenureMonths || 0, "number", 'style="width:44px"')}</td>
          <td><button type="button" class="bulk-del-row-btn" title="이 행 제외">🗑</button></td>
        </tr>`;
      }).join("");
      const errRows = parseErrors.map((e) =>
        `<tr class="bulk-row-err"><td colspan="14">⚠️ ${escapeHtml(e)}</td></tr>`
      ).join("");
      tblWrap.innerHTML = `<table class="bulk-preview-tbl">
        <thead><tr>
          <th>#</th><th>상태</th>
          <th>지역단</th><th>비전센터</th><th>지점</th><th>기수</th>
          <th>사번</th><th>성명</th><th>연락처</th>
          <th>기준실적(원)</th><th>마스터목표(원)</th><th>아너스목표(원)</th><th>차월</th>
          <th></th>
        </tr></thead>
        <tbody>${rows}${errRows}</tbody>
      </table>`;

      // 행 삭제 버튼
      tblWrap.querySelectorAll(".bulk-del-row-btn").forEach((btn) => {
        btn.addEventListener("click", () => {
          btn.closest("tr").classList.add("bulk-row-deleted");
          _updateBulkPreviewCount(tblWrap);
        });
      });
    } else {
      // ── 형식 A (22열 성과 데이터) — 열 재매핑 + 읽기전용 미리보기 ──
      const remapTh = (key, label, def) => {
        const curIdx = colOverride[key] ?? def;
        const optCols = [];
        for (let idx = 7; idx < Math.min(sampleCols.length, 25); idx++) {
          optCols.push({ idx, label: `열${idx}: ${sampleCols[idx] || "—"}` });
        }
        if (!optCols.length) return `<th>${label}</th>`;
        const opts = optCols.map((c) => `<option value="${c.idx}"${c.idx === curIdx ? " selected" : ""}>${c.label}</option>`).join("");
        return `<th class="bulk-remap-th">${label}<select class="bulk-remap-sel" data-field="${key}">${opts}</select></th>`;
      };
      const rows = records.map((r, i) => {
        const badge = r._isNew
          ? `<span class="bulk-badge-new">신규</span>`
          : `<span class="bulk-badge-exist">기존</span>`;
        const ctrCell = r.center ? escapeHtml(r.center) : `<span class="bulk-warn-center">⚠ 미확인</span>`;
        return `<tr>
          <td style="text-align:center">${i + 1}</td>
          <td>${badge}</td>
          <td>${escapeHtml(r.region || "")}</td><td>${ctrCell}</td><td>${escapeHtml(r.branch || "")}</td>
          <td>${escapeHtml(r.empNo || "")}</td><td>${escapeHtml(r.name || "")}</td>
          <td style="text-align:center">${escapeHtml(r.cohort || "—")}</td>
          <td class="r">${r.base ? Nf(r.base) : "—"}</td>
          <td class="r">${r.current ? Nf(r.current) : "—"}</td>
          <td class="r">${r.pgIpumCount != null ? r.pgIpumCount : "—"}</td>
          <td class="r">${r.pgIpumAmt ? Nf(r.pgIpumAmt) : "—"}</td>
        </tr>`;
      }).join("");
      const errRows = parseErrors.map((e) =>
        `<tr class="bulk-row-err"><td colspan="12">⚠️ ${escapeHtml(e)}</td></tr>`
      ).join("");
      tblWrap.innerHTML = `<table class="bulk-preview-tbl">
        <thead><tr>
          <th>#</th><th>상태</th><th>지역단</th><th>비전센터</th><th>지점</th>
          <th>사번</th><th>성명</th><th>기수</th>
          ${remapTh("base","기준실적",16+colShift)}
          ${remapTh("current","현재실적",17+colShift)}
          ${remapTh("pgIpumCount","인품건수",20+colShift)}
          ${remapTh("pgIpumAmt","인품실적",21+colShift)}
        </tr></thead>
        <tbody>${rows}${errRows}</tbody>
      </table>`;

      tblWrap.querySelectorAll(".bulk-remap-sel").forEach((sel) => {
        sel.addEventListener("change", () => {
          tblWrap.querySelectorAll(".bulk-remap-sel").forEach((s) => {
            state._bulkColOverride[s.dataset.field] = Number(s.value);
          });
          _renderBulkPreview(state._bulkPreviewRaw);
        });
      });
    }

    // 카운트
    const countEl = document.getElementById("bulk-preview-count");
    if (countEl) countEl.textContent = `${records.length}명 파싱${parseErrors.length ? ` · 오류 ${parseErrors.length}건` : ""}`;
    const saveBtn = document.getElementById("btn-bulk-preview-save");
    if (saveBtn) saveBtn.dataset.ready = "1";
  }

  function _updateBulkPreviewCount(tblWrap) {
    const active = tblWrap.querySelectorAll(".bulk-editable-row:not(.bulk-row-deleted)").length;
    const countEl = document.getElementById("bulk-preview-count");
    if (countEl) countEl.textContent = `${active}명 저장 예정`;
  }

  // 인라인 편집 미리보기에서 수집 후 저장
  async function _saveBulkFromPreview() {
    const tblWrap = document.getElementById("bulk-preview-tbl-wrap");
    const rows = tblWrap?.querySelectorAll(".bulk-editable-row:not(.bulk-row-deleted)") || [];
    if (!rows.length) { toast("저장할 데이터가 없습니다.", "error"); return { ok: 0, fail: 0, total: 0 }; }

    const records = [];
    rows.forEach((tr) => {
      const g = (f) => tr.querySelector(`[data-field="${f}"]`)?.value?.trim() || "";
      const empNo = g("empNo").replace(/[\s\/\\]/g, "");
      if (!empNo) return;
      const existSt = state.students.find((s) => s.empNo === empNo);
      const region = g("region") || state.filter.region || "";
      const center = g("center") || existSt?.center || _bulkInferCenter(region, "", g("branch"));
      const base = _bulkParseAmt(g("base"));
      let cohort = g("cohort") || state._bulkCohort || "";
      if (cohort && /^\d+$/.test(cohort)) cohort = cohort + "기";
      records.push({
        region, center,
        branch: g("branch"),
        cohort, empNo,
        name: g("name"),
        phone: g("phone"),
        base,
        target: _bulkParseAmt(g("target")) || (existSt?.target ?? 0),
        honors: _bulkParseAmt(g("honors")) || (existSt?.honors ?? 0),
        tenureMonths: _bulkParseAmt(g("tenureMonths")) || 0,
        current: Number(existSt?.current || 0),
        pgIpumCount: Number(existSt?.pgIpumCount || 0),
        pgIpumAmt: Number(existSt?.pgIpumAmt || 0),
        team: existSt?.team || "",
      });
    });

    if (!records.length) { toast("저장할 행이 없습니다.", "error"); return { ok: 0, fail: 0, total: 0 }; }

    setBulkProgress(`${records.length}건 저장중...`);
    let ok = 0, fail = 0;
    try {
      if (typeof window.DataAPI.saveMany === "function") {
        const { committed, errors } = await window.DataAPI.saveMany(records);
        ok = committed; fail = errors.length;
        if (errors.length) console.warn("[bulk edit save] 실패:", errors);
      } else {
        for (const r of records) { try { await window.DataAPI.save(r); ok++; } catch { fail++; } }
      }
    } catch (err) {
      fail = records.length; ok = 0;
      console.error("[bulk edit save]", err);
    }
    const summary = fail ? `${ok}건 저장 / ${fail}건 실패` : `${ok}건 저장 완료`;
    toast(summary, fail ? "error" : "success");
    setBulkProgress(summary, fail ? "error" : "success");
    return { ok, fail, total: records.length };
  }

  async function saveBulk(colOverride = {}) {
    const raw = $("#form-bulk").value.trim();
    if (!raw) {
      toast("붙여넣을 데이터가 없습니다.", "error");
      setBulkProgress("붙여넣을 데이터가 없습니다.", "error");
      return { ok: 0, fail: 0, total: 0 };
    }

    // 1. 파싱
    const { records, parseErrors, isFormatA } = parseBulkRecords(raw, colOverride);

    // Format A 에서 기수 미선택 경고
    if (isFormatA && !state._bulkCohort && !state.filter.cohort) {
      toast("기수를 선택하세요! (형식 A는 기수가 데이터에 없어 반드시 선택 필요)", "error");
      setBulkProgress("기수를 선택 후 다시 시도하세요.", "error");
      return { ok: 0, fail: 0, total: 0 };
    }

    if (records.length === 0) {
      const msg = `저장할 행이 없습니다. (${parseErrors[0] || "데이터 없음"})`;
      toast(msg, "error");
      setBulkProgress(msg, "error");
      return { ok: 0, fail: 0, total: 0 };
    }

    // 평균실적 단위 오입력 경고 (0 < base < 1000 은 의심)
    const baseWarnNames = records.filter((r) => r.base > 0 && r.base < 1000).map((r) => `${r.name || r.empNo}(${r.base}원)`);
    if (baseWarnNames.length) {
      toast(`⚠️ 평균실적 단위 확인: ${baseWarnNames.join(", ")} — 원(₩) 단위로 입력하세요`, "error");
    }

    // 2. 저장 — writeBatch 일괄 (1회 네트워크 호출)
    const btn = $("#btn-save");
    const origLabel = btn.textContent;
    btn.disabled = true;
    btn.textContent = "저장중...";
    setBulkProgress(`서버에 ${records.length}건 일괄 전송중...`);

    let ok = 0, fail = 0;
    const failMsgs = [];

    if (typeof window.DataAPI.saveMany === "function") {
      // 일괄 저장 경로
      try {
        const { committed, errors } = await window.DataAPI.saveMany(records);
        ok = committed;
        fail = errors.length;
        errors.forEach((e) => {
          failMsgs.push(`${e.empNo}: ${e.message}`);
          console.error("[bulk save] 실패:", e.empNo, e.message);
        });
      } catch (err) {
        fail = records.length;
        failMsgs.push(err.message || String(err));
        console.error("[bulk save] 일괄 저장 실패:", err);
      }
    } else {
      // 폴백: 5건씩 병렬
      const concurrency = 5;
      for (let i = 0; i < records.length; i += concurrency) {
        const chunk = records.slice(i, i + concurrency);
        setBulkProgress(`저장중... ${i}/${records.length}`);
        const results = await Promise.allSettled(chunk.map((r) => window.DataAPI.save(r)));
        results.forEach((r, j) => {
          if (r.status === "fulfilled") ok++;
          else {
            fail++;
            failMsgs.push(`${chunk[j].empNo}: ${r.reason && r.reason.message || r.reason}`);
            console.error("[bulk save] 실패:", chunk[j].empNo, r.reason);
          }
        });
      }
    }

    btn.disabled = false;
    btn.textContent = origLabel;

    const totalFail = fail + parseErrors.length;
    const summary = totalFail
      ? `${ok}건 저장 / ${totalFail}건 실패 — ${(failMsgs[0] || parseErrors[0])}`
      : `${ok}건 저장 완료`;
    toast(summary, totalFail ? "error" : "success");
    setBulkProgress(summary, totalFail ? "error" : "success");
    if (totalFail) console.warn("[bulk save] 실패 목록:", { parseErrors, failMsgs });
    return { ok, fail: totalFail, total: records.length + parseErrors.length };
  }

  // ========== 시드 파일 불러오기 ==========
  async function openSeedPicker() {
    openModal("#modal-seed");
    const list = $("#seed-list");
    list.innerHTML = `<li class="disabled">불러오는 중...</li>`;
    try {
      const res = await fetch("seed_data/index.json", { cache: "no-cache" });
      if (!res.ok) throw new Error("HTTP " + res.status);
      const data = await res.json();
      const seeds = (data && data.seeds) || [];
      list.innerHTML = "";
      if (seeds.length === 0) {
        list.innerHTML = `<li class="disabled">등록된 시드 파일이 없습니다.</li>`;
        return;
      }
      seeds.forEach((seed) => {
        const li = document.createElement("li");
        const countText = seed.count ? `${seed.count}명` : "";
        li.innerHTML = `<strong>${escapeHtml(seed.label)}</strong>` +
          `<div style="font-size:12px;color:#999;margin-top:2px;">${escapeHtml(seed.file)}${countText ? " · " + countText : ""}</div>`;
        li.style.cursor = "pointer";
        li.addEventListener("click", () => loadSeed(seed));
        list.appendChild(li);
      });
    } catch (err) {
      console.error("[seed] 목록 로드 실패:", err);
      list.innerHTML = `<li class="disabled">시드 목록 로드 실패: ${escapeHtml(err.message)}<br><small>로컬에서 file:// 로 여신 경우 CORS로 차단됩니다. 웹서버로 접속하세요.</small></li>`;
    }
  }

  async function loadSeed(seed) {
    try {
      const res = await fetch("seed_data/" + seed.file, { cache: "no-cache" });
      if (!res.ok) throw new Error("HTTP " + res.status);
      const text = await res.text();
      $("#form-bulk").value = text.trim();
      closeModal("#modal-seed");
      switchTab("bulk");
      toast(`[${seed.label}] 불러옴. [저장] 버튼을 눌러 등록하세요.`, "success");
    } catch (err) {
      console.error("[seed] 파일 로드 실패:", err);
      toast("시드 파일 로드 실패: " + err.message, "error");
    }
  }

  async function removeStudent(empNo) {
    if (!confirm(`사번 ${empNo} 교육생을 삭제하시겠습니까?`)) return;
    await window.DataAPI.remove(empNo);
    toast("삭제되었습니다.", "success");
  }

  function exportCSV() {
    const list = filteredStudents();
    if (list.length === 0) { toast("내보낼 데이터가 없습니다.", "error"); return; }
    const headers = ["지역단","비전센터","지점","기수","사번","이름","연락처","기준실적","마스터목표","아너스목표"];
    const rows = list.map((s) => [
      s.region, s.center, s.branch, s.cohort, s.empNo, s.name, s.phone,
      Math.round(Number(s.base||0)), Math.round(Number(s.target||0)), Math.round(Number(s.honors||0))
    ].map((v) => `"${String(v ?? "").replace(/"/g, '""')}"`).join(","));
    const csv = "\uFEFF" + [headers.join(","), ...rows].join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `master_students_${new Date().toISOString().slice(0,10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }

  // ========== 실적진도 시상내역 엑셀 저장 ==========
  function loadSheetJS() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    return new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
      script.onload  = () => resolve(window.XLSX);
      script.onerror = () => reject(new Error("SheetJS 로드 실패"));
      document.head.appendChild(script);
    });
  }

  function loadTesseract() {
    if (window.Tesseract) return Promise.resolve(window.Tesseract);
    return new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = "https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js";
      script.onload  = () => resolve(window.Tesseract);
      script.onerror = () => reject(new Error("Tesseract.js 로드 실패"));
      document.head.appendChild(script);
    });
  }

  // OCR 텍스트에서 탭 구분 테이블 데이터 재구성
  function _ocrTextToTsv(rawText) {
    const lines = rawText.split(/\n/).map(l => l.trimEnd()).filter(l => l.trim());
    return lines.map(line => {
      // 2칸 이상 공백을 탭으로 치환 (열 구분자)
      const cols = line.trim().split(/\s{2,}/);
      return cols.length >= 3 ? cols.join("\t") : line.trim().replace(/\s+/g, "\t");
    }).join("\n");
  }

  // TSV 문자열을 HTML 미리보기 테이블로 변환
  function _tsvToPreviewTable(tsv) {
    const lines = tsv.split(/\n/).filter(l => l.trim());
    if (!lines.length) return "<div class='rm-img-ocr-placeholder'>인식된 데이터 없음</div>";
    const isNumeric = v => /^[0-9,.\-+]+$/.test(v.replace(/\s/g, ""));
    let html = "<table><thead><tr>";
    const firstCols = lines[0].split("\t");
    // 첫 행이 한글 등 헤더처럼 보이면 thead로
    const hasHeader = firstCols.some(c => /[가-힣]/.test(c));
    if (hasHeader) {
      firstCols.forEach(c => { html += `<th>${escapeHtml(c)}</th>`; });
      html += "</tr></thead><tbody>";
      lines.slice(1).forEach(line => {
        html += "<tr>";
        line.split("\t").forEach(c => {
          html += `<td class="${isNumeric(c) ? "rm-cell-num" : ""}">${escapeHtml(c)}</td>`;
        });
        html += "</tr>";
      });
    } else {
      html += "<tr></tr></thead><tbody>";
      lines.forEach(line => {
        html += "<tr>";
        line.split("\t").forEach(c => {
          html += `<td class="${isNumeric(c) ? "rm-cell-num" : ""}">${escapeHtml(c)}</td>`;
        });
        html += "</tr>";
      });
    }
    html += "</tbody></table>";
    return html;
  }

  function _bindRmImagePaste() {
    const btn     = document.getElementById("btn-rm-img-paste");
    const wrap    = document.getElementById("rm-img-preview-wrap");
    const canvas  = document.getElementById("rm-img-canvas");
    const ocrText = document.getElementById("rm-img-ocr-text");
    const statusEl = document.getElementById("rm-img-status");
    const applyBtn = document.getElementById("btn-rm-img-apply");
    const closeBtn  = document.getElementById("btn-rm-img-close");
    const cancelBtn = document.getElementById("btn-rm-img-cancel");
    const tablePreview = document.getElementById("rm-img-table-preview");
    if (!btn || !canvas) return;

    let _pasteActive = false;
    let _pasteHandler = null;

    function _deactivate() {
      if (_pasteHandler) document.removeEventListener("paste", _pasteHandler);
      _pasteHandler = null;
      _pasteActive = false;
      btn.textContent = "📷 사진으로 붙여넣기";
      btn.classList.remove("active");
      btn.disabled = false;
    }

    async function _processImageFile(file) {
      const url = URL.createObjectURL(file);
      const img = new Image();
      img.onload = async () => {
        canvas.width  = img.naturalWidth;
        canvas.height = img.naturalHeight;
        canvas.getContext("2d").drawImage(img, 0, 0);
        URL.revokeObjectURL(url);

        ocrText.value = "";
        applyBtn.disabled = true;
        statusEl.textContent = "OCR 초기화 중… (첫 실행 시 언어 데이터 다운로드)";
        if (tablePreview) tablePreview.innerHTML = "<div class='rm-img-ocr-placeholder'>인식 중...</div>";
        wrap.hidden = false;

        try {
          const Tesseract = await loadTesseract();
          const worker = await Tesseract.createWorker(["kor", "eng"], 1, {
            logger: (m) => {
              if (m.status === "recognizing text") {
                statusEl.textContent = `인식 중… ${Math.round((m.progress || 0) * 100)}%`;
              } else if (m.status && m.status !== "initialized api") {
                statusEl.textContent = m.status;
              }
            }
          });
          await worker.setParameters({ tessedit_pageseg_mode: "6" });
          const { data } = await worker.recognize(canvas);
          await worker.terminate();

          const tsv = _ocrTextToTsv(data.text);
          ocrText.value = tsv;
          if (tablePreview) tablePreview.innerHTML = _tsvToPreviewTable(tsv);
          applyBtn.disabled = false;
          statusEl.textContent = `✅ 인식 완료 — 이미지와 비교해 오류를 수정하세요`;
        } catch (err) {
          console.error("[OCR]", err);
          statusEl.textContent = `❌ 인식 실패: ${err.message}`;
          if (tablePreview) tablePreview.innerHTML = "<div class='rm-img-ocr-placeholder'>인식 실패. 텍스트를 직접 입력하세요.</div>";
        }
      };
      img.onerror = () => { URL.revokeObjectURL(url); toast("이미지를 불러올 수 없습니다.", "error"); };
      img.src = url;
    }

    btn.addEventListener("click", () => {
      if (_pasteActive) { _deactivate(); return; }
      _pasteActive = true;
      btn.textContent = "⏳ Ctrl+V로 이미지 붙여넣기…";
      btn.classList.add("active");
      btn.disabled = false;

      _pasteHandler = (e) => {
        const items = e.clipboardData?.items || [];
        let imgFile = null;
        for (const item of items) {
          if (item.type.startsWith("image/")) { imgFile = item.getAsFile(); break; }
        }
        if (!imgFile) { toast("이미지를 찾을 수 없습니다. 이미지를 먼저 복사(Ctrl+C)해 주세요.", "warn"); return; }
        e.preventDefault();
        _deactivate();
        _processImageFile(imgFile);
      };
      document.addEventListener("paste", _pasteHandler);

      // 15초 후 자동 해제
      setTimeout(() => { if (_pasteActive) { _deactivate(); } }, 15000);
    });

    // textarea에 직접 이미지 드래그 드롭
    const pasteArea = document.getElementById("rm-paste-area");
    if (pasteArea) {
      pasteArea.addEventListener("dragover", e => e.preventDefault());
      pasteArea.addEventListener("drop", e => {
        const file = e.dataTransfer?.files?.[0];
        if (file?.type.startsWith("image/")) {
          e.preventDefault();
          _processImageFile(file);
        }
      });
    }

    closeBtn?.addEventListener("click", () => {
      wrap.hidden = true;
      ocrText.value = "";
      statusEl.textContent = "";
      if (tablePreview) tablePreview.innerHTML = "<div class='rm-img-ocr-placeholder'>인식 결과가 여기 표시됩니다</div>";
    });
    cancelBtn?.addEventListener("click", () => closeBtn?.click());

    applyBtn?.addEventListener("click", () => {
      const text = ocrText.value.trim();
      if (!text) return;
      const area = document.getElementById("rm-paste-area");
      if (area) area.value = text;
      wrap.hidden = true;
      ocrText.value = "";
      statusEl.textContent = "";
      // 자동으로 저장 플로우 시작
      document.getElementById("btn-rm-paste-apply")?.click();
    });
  }

  async function exportProgressAwardExcel() {
    const region    = state.progressRegion || state.filter.region || "";
    const _pgCohort = (state.filter.cohort || state.progressCohort || "").replace(/기$/, "");
    const _pgStep   = state.filter.step    || state.progressStep   || "1";
    const center    = state.filter.center || "";

    const cohortLabel = _pgCohort ? `${_pgCohort}기` : "";
    const stepLabel   = `Step${_pgStep}`;
    const centerPart  = center ? ` ${center}` : "";
    const baseTitle   = `${region}${cohortLabel ? " " + cohortLabel : ""} ${stepLabel}`;
    const msg = `${region || "지역단미선택"}${cohortLabel ? " " + cohortLabel : ""}${centerPart} ${stepLabel} 수상내역을 인쇄 하시겠습니까?`;

    if (!confirm(msg)) return;

    let XLSX;
    try {
      toast("엑셀 파일 준비 중…", "");
      XLSX = await loadSheetJS();
    } catch (_e) {
      toast("SheetJS 로드 실패. 인터넷 연결을 확인하세요.", "error");
      return;
    }

    const list = state.students.filter((s) => {
      if (s.region !== region) return false;
      if (_pgCohort && s.cohort && String(s.cohort).replace(/기$/, "") !== String(_pgCohort)) return false;
      return true;
    });
    if (!list.length) { toast("해당하는 교육생이 없습니다.", "error"); return; }

    const _pa       = getProgressAwardConfig(region, _pgCohort, _pgStep);
    const plan      = _pa.plan;
    const stats     = list.map(getProgressStat);
    const byRate    = [...stats].sort((a, b) => (b.net / (b.base || 1)) - (a.net / (a.base || 1)));
    const byAmt     = [...stats].sort((a, b) => b.net - a.net);
    const byIpum    = [...stats].filter((s) => s.ipumAmt > 0).sort((a, b) => b.ipumAmt - a.ipumAmt || b.ipumCount - a.ipumCount);
    const _bothAsgn = computeBothAwardAssignments(byRate, byAmt, _pa);
    const ipumRankMap = new Map(byIpum.map((st, i) => [st.s.empNo, i + 1]));

    // 그룹 순위 계산
    const hasAnyTeam = stats.some((s) => (s.s.team || "").toString().trim());
    const groupKeyFn = hasAnyTeam
      ? ((s) => (s.s.team || "").toString().trim() || "(팀 미배정)")
      : ((s) => s.s.branch || "(미지정)");
    const groupMap = {};
    stats.forEach((st) => {
      const k = groupKeyFn(st);
      if (!groupMap[k]) groupMap[k] = { base: 0, current: 0, members: [], memberStats: [] };
      groupMap[k].base += st.base;
      groupMap[k].current += st.current;
      groupMap[k].members.push(st.s.name || "");
      groupMap[k].memberStats.push(st);
    });
    const groupRanking = Object.entries(groupMap)
      .map(([name, g]) => ({ name, rate: g.base > 0 ? (g.current / g.base) * 100 : 0, ...g }))
      .sort((a, b) => b.rate - a.rate);
    const groupLabel = hasAnyTeam ? "팀" : "지점";
    const ga1 = plan?.groupAward1;
    const ga2 = plan?.groupAward2;
    const ga1En = !!ga1?.enabled;
    const ga2En = !!ga2?.enabled;
    const ga1ItemList = _ga1Items(ga1);
    const ga2ItemList = _ga2Items(ga2);

    // 순위별 시상 라벨 (물품명 포함)
    const rankAwardLabel = (config, rank) => {
      if (!config?.enabled || !rank || rank < 1) return "-";
      const v = (config.payouts || [])[rank - 1];
      return v != null ? payoutLabel(v) : "-";
    };

    // 비전센터 + 지점 합산 표시 헬퍼
    const ctbr = (s) => [s.center, s.branch].filter(Boolean).join(" / ") || "";

    // 시트 생성 헬퍼 (첫 행 전체 병합)
    const makeWs = (rows, colWidths) => {
      const ws = XLSX.utils.aoa_to_sheet(rows);
      if (colWidths) ws["!cols"] = colWidths;
      const ncols = colWidths?.length || 5;
      ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: ncols - 1 } }];
      return ws;
    };
    // 숫자 열에 3자리 콤마 포맷 적용 헬퍼 (0-indexed 열 배열)
    const applyNumFmt = (ws, numCols) => {
      if (!numCols || !numCols.length) return ws;
      const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
      for (let R = range.s.r; R <= range.e.r; R++) {
        for (const C of numCols) {
          const addr = XLSX.utils.encode_cell({ r: R, c: C });
          const cell = ws[addr];
          if (cell && cell.t === "n") cell.z = "#,##0";
        }
      }
      return ws;
    };

    const wb = XLSX.utils.book_new();

    // ── Sheet 0: 시상안 (기수 시상 계획 전체) ──
    {
      const rows = [];
      rows.push([`${baseTitle} 시상안`]);
      rows.push([]);

      const pushSection = (title) => rows.push(["[ " + title + " ]", "", ""]);
      const pushHdr2 = (a, b) => rows.push(["구분", a, b]);
      const pushRow2 = (label, a, b) => rows.push([label, a, b ?? ""]);

      // 개인순증시상
      pushSection("개인순증시상");
      if (!_pa.tiers.length) {
        rows.push(["", "(미설정)", ""]);
      } else {
        pushHdr2("순증 조건", "시상 내역");
        const sortedTiers = (plan.personalIncr?.items || [])
          .filter(it => Number(it.critVal) > 0)
          .slice().sort((a, b) => Number(b.critVal) - Number(a.critVal));
        sortedTiers.forEach(it => {
          const cond = `순증 ${it.critVal}만원↑`;
          const pay = it.payType === "pct" ? `${it.payVal}%` : `${it.payVal}만원`;
          pushRow2("", cond, pay);
        });
      }
      rows.push([]);

      // 신장률 TOP
      pushSection("신장률 TOP 시상");
      if (!_pa.rateConfig) {
        rows.push(["", "(미설정)", ""]);
      } else {
        const rN = Number(_pa.rateConfig.n) || 0;
        pushHdr2("순위", `TOP${rN} 시상 내역`);
        Array.from({ length: rN }, (_, i) => {
          pushRow2("", `${i + 1}위`, rankAwardLabel(_pa.rateConfig, i + 1));
        });
      }
      rows.push([]);

      // 신장액 TOP
      pushSection("신장액 TOP 시상");
      if (!_pa.amtConfig) {
        rows.push(["", "(미설정)", ""]);
      } else {
        const aN = Number(_pa.amtConfig.n) || 0;
        pushHdr2("순위", `TOP${aN} 시상 내역`);
        Array.from({ length: aN }, (_, i) => {
          pushRow2("", `${i + 1}위`, rankAwardLabel(_pa.amtConfig, i + 1));
        });
      }
      rows.push([]);

      // 그룹시상
      if (ga1En || ga2En) {
        pushSection("그룹시상");
        if (ga1En) {
          pushHdr2("그룹시상1 조건", "시상 내역");
          ga1ItemList.forEach(it => {
            pushRow2("", `전원 순증 ${it.threshold || 5}만원↑`, payoutLabel(it.payout));
          });
        }
        if (ga2En) {
          pushHdr2("그룹시상2 조건", "시상 내역");
          ga2ItemList.forEach(it => {
            pushRow2("", `달성률 ${it.rateThreshold || 110}%↑`, payoutLabel(it.payout ?? 15));
          });
          if (ga2?.linkToGroup1 && ga1En) rows.push(["", "(그룹시상1 달성 시에만 지급)", ""]);
        }
        rows.push([]);
      }

      // 자격조건
      if (plan.eligibility?.enabled) {
        const conds = plan.eligibility.conditions || [];
        if (conds.length) {
          pushSection("시상 자격조건");
          const fLabel = (f) => ({ converted: "환산실적", hiCap: "하이캡", monthly: "월납보험료" }[f] || f);
          const fUnit  = (f) => f === "hiCap" ? "" : "만원";
          const op = plan.eligibility.operator === "or" ? "또는" : "그리고";
          conds.forEach((c, i) => {
            pushRow2(i === 0 ? "제외 조건" : op, `${fLabel(c.field)} ${c.threshold}${fUnit(c.field)} 이하`, "→ 시상 제외");
          });
          rows.push([]);
        }
      }

      // 중복시상 없음
      if (_pa.bothEnabled) {
        pushSection("중복시상 없음");
        rows.push(["", "신장률·신장액 TOP 중복 시 더 큰 시상 1회만 지급", ""]);
        rows.push([]);
      }

      // 비고
      if (plan.notes) {
        pushSection("비고");
        rows.push(["", plan.notes, ""]);
        rows.push([]);
      }

      XLSX.utils.book_append_sheet(wb, makeWs(rows,
        [{ wch: 12 }, { wch: 20 }, { wch: 24 }]), "시상안");
    }

    // ── Sheet 1: 개인순증시상 ──
    {
      const rows = [];
      rows.push([`${baseTitle} 개인순증시상`]);
      rows.push([]);
      if (!_pa.tiers.length) {
        rows.push(["개인순증시상 미설정"]);
      } else {
        rows.push(["순번", "이름", "비전센터/지점", "기준실적(원)", "현재실적(원)", "순증(원)", "달성률(%)", "시상내역"]);
        let cnt = 0;
        byAmt.forEach((st) => {
          const _tlbl = tierLabel(st.net, undefined, _pa);
          if (_tlbl === "-") return;
          const _tamt = tierAward(st.net, undefined, _pa);
          const tierDisp = _tamt > 0 ? `${_tlbl} (${Math.round(_tamt / 10000)}만원)` : _tlbl;
          rows.push([++cnt, st.s.name || "", ctbr(st.s),
            st.base, st.current, st.net,
            parseFloat(st.rate.toFixed(1)), tierDisp]);
        });
        if (!cnt) rows.push(["", "(해당자 없음)"]);
        rows.push([]);
        rows.push(["", "", "", "", "", "", "", `총 ${cnt}명 시상`]);
      }
      const _ws1 = makeWs(rows, [{ wch: 5 }, { wch: 10 }, { wch: 18 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 10 }, { wch: 22 }]);
      applyNumFmt(_ws1, [3, 4, 5]);
      XLSX.utils.book_append_sheet(wb, _ws1, "개인순증시상");
    }

    // ── Sheet 2: 신장TOP (신장률+신장액 통합, 순위별 두 행 interleaved) ──
    {
      const rows = [];
      rows.push([`${baseTitle} 신장률·신장액 TOP 시상`]);
      rows.push([]);
      const _colHdr = ["순위", "구분", "이름", "비전센터/지점", "기준실적(원)", "현재실적(원)", "순증(원)", "달성률(%)", "시상내역"];
      if (!_pa.rateConfig && !_pa.amtConfig) {
        rows.push(["신장TOP 시상 미설정"]);
      } else {
        rows.push(_colHdr);
        // 수상자 목록을 rank → student 맵으로 정리
        const _rateWinMap = new Map(); // effectiveRank → st
        byRate.forEach(st => {
          const rA = _bothAsgn.rateAsgn.get(st.s.empNo);
          if (rA && rA.status === "mine") _rateWinMap.set(rA.effectiveRank, { st, rA });
        });
        const _amtWinMap = new Map();
        byAmt.forEach(st => {
          const aA = _bothAsgn.amtAsgn.get(st.s.empNo);
          if (aA && aA.status === "mine") _amtWinMap.set(aA.effectiveRank, { st, aA });
        });
        const maxN = Math.max(
          _pa.rateConfig ? Number(_pa.rateConfig.n) || 0 : 0,
          _pa.amtConfig  ? Number(_pa.amtConfig.n)  || 0 : 0
        );
        let rateCnt = 0, amtCnt = 0;
        for (let rank = 1; rank <= maxN; rank++) {
          const rEntry = _rateWinMap.get(rank);
          const aEntry = _amtWinMap.get(rank);
          const hasR = !!rEntry && !!_pa.rateConfig && rank <= Number(_pa.rateConfig.n);
          const hasA = !!aEntry && !!_pa.amtConfig  && rank <= Number(_pa.amtConfig.n);
          if (hasR) {
            const { st, rA } = rEntry;
            rows.push([rank, "신장률", st.s.name || "", ctbr(st.s),
              st.base, st.current, st.net, parseFloat(st.rate.toFixed(1)),
              rankAwardLabel(_pa.rateConfig, rA.effectiveRank)]);
            rateCnt++;
          }
          if (hasA) {
            const { st, aA } = aEntry;
            rows.push([hasR ? "" : rank, "신장액", st.s.name || "", ctbr(st.s),
              st.base, st.current, st.net, parseFloat(st.rate.toFixed(1)),
              rankAwardLabel(_pa.amtConfig, aA.effectiveRank)]);
            amtCnt++;
          }
          if (!hasR && !hasA && rank <= maxN) break;
          if (hasR || hasA) rows.push(["", "", "", "", "", "", "", "", ""]);
        }
        rows.push([]);
        rows.push(["", "", "", "", "", "", "", "", `신장률 ${rateCnt}명 · 신장액 ${amtCnt}명 시상`]);
      }
      const _ws2 = makeWs(rows, [{ wch: 5 }, { wch: 7 }, { wch: 10 }, { wch: 18 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 10 }, { wch: 22 }]);
      applyNumFmt(_ws2, [4, 5, 6]);
      XLSX.utils.book_append_sheet(wb, _ws2, "신장TOP");
    }

    // ── Sheet 4: 그룹시상 (설정된 경우만) ──
    if (ga1En || ga2En) {
      const rows = [];
      rows.push([`${baseTitle} 그룹시상`]);
      rows.push([]);
      const hdr = ["순위", groupLabel, "구성원", "기준합계(원)", "현재합계(원)", "달성률(%)"];
      if (ga1En) hdr.push("그룹시상1");
      if (ga2En) hdr.push("그룹시상2");
      rows.push(hdr);
      groupRanking.forEach((g, i) => {
        const row = [i + 1, g.name, g.members.join(", "),
          g.base, g.current, parseFloat(g.rate.toFixed(1))];
        if (ga1En) {
          const total = g.memberStats.length || g.members.length;
          const metItems = ga1ItemList.filter(it =>
            g.memberStats.every(st => (st.net || 0) >= Number(it.threshold || 5) * 10000) && total > 0
          );
          row.push(metItems.length ? metItems.map(it => payoutLabel(it.payout)).join("+") : "미달");
        }
        if (ga2En) {
          // ga1 활성화 시 항상 연계 — ga1 미달이면 ga2도 미달
          const ga1Met = ga1En && ga1ItemList.some(it =>
            g.memberStats.every(st => (st.net || 0) >= Number(it.threshold || 5) * 10000) && g.memberStats.length > 0
          );
          if (ga1En && !ga1Met) {
            row.push("미달(그룹1↑필요)");
          } else {
            const metGa2 = ga2ItemList.filter(it => g.rate >= Number(it.rateThreshold || 110));
            if (metGa2.length) {
              const bestPay = metGa2.reduce((best, it) => {
                const np = normPayout(it.payout ?? 15);
                if (np.type === "item") return best || np;
                return (!best || Number(np.val) > Number(normPayout(best || 0).val)) ? np : best;
              }, null);
              row.push(payoutLabel(bestPay));
            } else {
              row.push(`미달(${ga2ItemList[0]?.rateThreshold || 110}%↑)`);
            }
          }
        }
        rows.push(row);
      });
      const colW = [{ wch: 5 }, { wch: 12 }, { wch: 30 }, { wch: 14 }, { wch: 14 }, { wch: 10 }];
      if (ga1En) colW.push({ wch: 20 });
      if (ga2En) colW.push({ wch: 20 });
      const _wsGrp = makeWs(rows, colW);
      applyNumFmt(_wsGrp, [3, 4]);
      XLSX.utils.book_append_sheet(wb, _wsGrp, "그룹시상");
    }

    // ── Sheet 5: 인품왕 (데이터 있는 경우만) ──
    if (byIpum.length > 0) {
      const rows = [];
      rows.push([`${baseTitle} 인품왕 순위`]);
      rows.push([]);
      rows.push(["순위", "이름", "비전센터/지점", "기준실적(원)", "현재실적(원)", "달성률(%)", "인품건수", "인품실적(원)"]);
      byIpum.forEach((st, i) => {
        const grade = ["인품의 황제", "인품의 제왕", "인품의 왕"][i] || `${i + 1}위`;
        rows.push([grade, st.s.name || "", ctbr(st.s),
          st.base, st.current, parseFloat(st.rate.toFixed(1)),
          st.ipumCount, st.ipumAmt]);
      });
      const _wsIpum = makeWs(rows,
        [{ wch: 12 }, { wch: 10 }, { wch: 18 }, { wch: 14 }, { wch: 14 }, { wch: 10 }, { wch: 10 }, { wch: 14 }]);
      applyNumFmt(_wsIpum, [3, 4, 7]);
      XLSX.utils.book_append_sheet(wb, _wsIpum, "인품왕");
    }

    // ── Sheet 6: 전체검증 ──
    {
      const rows = [];
      rows.push([`${baseTitle} 전체 교육생 시상 검증`]);
      rows.push([]);
      rows.push(["순번", "이름", "비전센터/지점", "기준(원)", "현재(원)", "순증(원)", "달성률(%)",
        "개인순증시상", "신장률TOP", "신장액TOP", "인품왕"]);
      byAmt.forEach((st, i) => {
        const _tamt = tierAward(st.net, undefined, _pa);
        const _tlbl = tierLabel(st.net, undefined, _pa);
        const tierText = _tamt > 0
          ? `${_tlbl} (${Math.round(_tamt / 10000)}만원)`
          : _tlbl !== "-" ? _tlbl
          : (_pa.tiers.length > 0 ? "해당없음" : "-");
        const rA = _bothAsgn.rateAsgn.get(st.s.empNo);
        const aA = _bothAsgn.amtAsgn.get(st.s.empNo);
        const rateTxt = !rA ? "-"
          : rA.status === "mine"       ? `${rA.effectiveRank}위 (${rankAwardLabel(_pa.rateConfig, rA.effectiveRank)})`
          : rA.status === "other"      ? "→신장액시상"
          : rA.status === "ineligible" ? "기준미달" : "-";
        const amtTxt = !aA ? "-"
          : aA.status === "mine"       ? `${aA.effectiveRank}위 (${rankAwardLabel(_pa.amtConfig, aA.effectiveRank)})`
          : aA.status === "other"      ? "→신장률시상"
          : aA.status === "ineligible" ? "기준미달" : "-";
        const ipumRank = ipumRankMap.get(st.s.empNo);
        rows.push([i + 1, st.s.name || "", ctbr(st.s),
          st.base, st.current, st.net,
          parseFloat(st.rate.toFixed(1)),
          tierText, rateTxt, amtTxt, ipumRank ? `${ipumRank}위` : "-"]);
      });
      const _wsAll = makeWs(rows, [
        { wch: 5 }, { wch: 10 }, { wch: 18 }, { wch: 14 }, { wch: 14 },
        { wch: 14 }, { wch: 10 }, { wch: 20 }, { wch: 18 }, { wch: 18 }, { wch: 10 },
      ]);
      applyNumFmt(_wsAll, [3, 4, 5]);
      XLSX.utils.book_append_sheet(wb, _wsAll, "전체검증");
    }

    const today   = new Date();
    const dateStr = `${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, "0")}${String(today.getDate()).padStart(2, "0")}`;
    const cohortPart = cohortLabel ? `_${cohortLabel}` : "";
    const filename = `${region}${cohortPart}_${stepLabel}_수상내역_${dateStr}.xlsx`;
    XLSX.writeFile(wb, filename);
    toast("엑셀 파일이 저장되었습니다.", "success");
  }

  // ========== 초기화 ==========
  function bindEvents() {
    // 상단 헤더 — 캐시 초기화 후 새로고침
    const hrBtn = $("#btn-hard-refresh");
    if (hrBtn) hrBtn.addEventListener("click", async () => {
      if (!confirm("캐시를 초기화하고 페이지를 새로고침 합니다. 계속할까요?")) return;
      hrBtn.classList.add("spinning");
      hrBtn.disabled = true;
      try {
        // 1) Service Worker 등록 전체 해제
        if ("serviceWorker" in navigator) {
          const regs = await navigator.serviceWorker.getRegistrations();
          await Promise.all(regs.map((r) => r.unregister()));
        }
        // 2) Cache Storage 전체 삭제
        if ("caches" in window) {
          const keys = await caches.keys();
          await Promise.all(keys.map((k) => caches.delete(k)));
        }
      } catch (e) {
        console.warn("[hard-refresh] 캐시 정리 중 오류:", e);
      }
      // 3) 쿼리스트링에 타임스탬프 붙여 강제 새로고침
      const url = new URL(window.location.href);
      url.searchParams.set("_r", Date.now().toString());
      window.location.replace(url.toString());
    });

    const errBtn = document.getElementById("btn-error-report");
    if (errBtn) errBtn.addEventListener("click", openErrorReportModal);

    // 사이드바 인라인 셀렉트
    populateFilterRegionSelect();
    $("#filter-region-select").addEventListener("change", (e) => {
      state.filter.region = e.target.value;
      state.filter.center = "";
      state.filter.branch = "";
      syncFilterOrgSelects();
      persistFilter();
      render();
    });
    $("#filter-center-select").addEventListener("change", (e) => {
      state.filter.center = e.target.value;
      state.filter.branch = "";
      syncFilterOrgSelects();
      persistFilter();
      render();
    });
    $("#filter-branch-select").addEventListener("change", (e) => {
      state.filter.branch = e.target.value;
      persistFilter();
      render();
    });
    $("#btn-reset-filter").addEventListener("click", () => {
      state.filter = { region: DEFAULT_REGION, center: "", branch: "", cohort: "", step: "1", q: "" };
      $("#filter-cohort").value = "";
      $("#search-box-side").value = "";
      state.progressStep = "1";
      const filterStep = document.getElementById("filter-step");
      if (filterStep) filterStep.value = "1";
      const pgStepSel = document.getElementById("pg-step-sel");
      if (pgStepSel) pgStepSel.value = "1";
      syncOrgLabels();
      persistFilter();
      render();
    });
    $("#filter-cohort").addEventListener("change", (e) => {
      const prevCohortNum = parseInt(state.filter.cohort, 10) || 0;
      const newCohortNum  = parseInt(e.target.value, 10) || 0;
      state.filter.cohort = e.target.value;
      // 뒤 기수 → 앞 기수: Step 2 / 앞 기수 → 뒤 기수: Step 1
      const autoStep = (newCohortNum > 0 && prevCohortNum > 0 && newCohortNum < prevCohortNum) ? "2" : "1";
      state.filter.step = autoStep;
      state.progressStep = autoStep;
      const fsEl = document.getElementById("filter-step");
      if (fsEl) fsEl.value = autoStep;
      const pgStepSel = document.getElementById("pg-step-sel");
      if (pgStepSel) pgStepSel.value = autoStep;
      persistFilter();
      if (isPanelVisible("progress-panel")) renderProgressPanel();
      render();
    });
    document.getElementById("filter-step")?.addEventListener("change", (e) => {
      state.filter.step = e.target.value;
      state.progressStep = e.target.value;
      persistFilter();
      const pgStepSel = document.getElementById("pg-step-sel");
      if (pgStepSel) pgStepSel.value = e.target.value;
      if (isPanelVisible("progress-panel")) renderProgressPanel();
      render();
    });
    $("#search-box-side").addEventListener("input", (e) => {
      state.filter.q = e.target.value;
      render();
    });
    $("#search-box-side")?.addEventListener("click", (e) => {
      if (e.target.value) { e.target.value = ""; e.target.dispatchEvent(new Event("input")); }
    });

    // 미지정 교육생 알림 버튼
    document.getElementById("btn-unassigned-alert")?.addEventListener("click", openUnassignedModal);

    // 모달 닫기
    $$("[data-close]").forEach((el) => el.addEventListener("click", (e) => {
      e.target.closest(".modal").hidden = true;
    }));

    // 상단 탭 네비게이션 — 클릭한 view 만 표시
    $$(".top-nav a").forEach((a) => {
      a.addEventListener("click", (e) => {
        e.preventDefault();
        switchView(a.getAttribute("href") || "#students");
      });
    });

    // 휴대폰 하단 바 — 홈 버튼은 사이드바 토글, 나머지는 switchView
    $$(".mobile-bottom-nav .mbn-btn").forEach((a) => {
      a.addEventListener("click", (e) => {
        e.preventDefault();
        if (a.dataset.action === "toggleSidebar") {
          toggleMobileSidebar();
          return;
        }
        closeMobileSidebar();
        switchView(a.getAttribute("href") || "#dashboard");
      });
    });
    // 백드롭 클릭 시 사이드바 닫기
    const bd = $("#mobile-sidebar-backdrop");
    if (bd) bd.addEventListener("click", closeMobileSidebar);
    // 좌우 스와이프 제스처 (모바일 한정)
    bindMobileSwipe();

    // 등록 버튼
    bindFormTargetPopup(); // 마스터목표 팝업 버튼 초기화 (DOM 1회 바인딩)
    $("#btn-open-add").addEventListener("click", () => {
      resetForm();
      switchTab("single");
      openModal("#modal-add");
    });

    // 폼 조직 선택 (네이티브 드롭다운)
    populateRegionSelect();
    $("#form-region-select").addEventListener("change", (e) => {
      state.form.region = e.target.value;
      state.form.center = "";
      state.form.branch = "";
      syncFormOrgSelects();
    });
    $("#form-center-select").addEventListener("change", (e) => {
      state.form.center = e.target.value;
      state.form.branch = "";
      syncFormOrgSelects();
    });
    $("#form-branch-select").addEventListener("change", (e) => {
      state.form.branch = e.target.value;
    });

    // 사번 입력 시 기존 교육생 자동 로드 (첫 등록 제외, 사번 기준 매칭)
    $("#form-empno")?.addEventListener("blur", () => {
      if (editingEmpNo) return;
      const empNo = ($("#form-empno")?.value || "").trim();
      if (!empNo) return;
      const s = state.students.find((x) => x.empNo === empNo);
      if (!s) return;
      state.form = { region: s.region || "", center: s.center || "", branch: s.branch || "" };
      $("#form-name").value   = s.name   || "";
      $("#form-phone").value  = s.phone  || "";
      $("#form-base").value   = s.base   || "";
      const computedTarget = (s.region !== "호남지역단" && !Number(s.target))
        ? getProgressStat(s).base + 50000
        : Number(s.target) || "";
      $("#form-target").value  = computedTarget;
      $("#form-honors").value  = s.honors || "";
      $("#form-cohort").value  = s.cohort || "";
      const teamNum = parseInt(s.team, 10);
      const formTeamEl = $("#form-team");
      if (formTeamEl) formTeamEl.value = teamNum > 0 ? teamNum : "";
      editingEmpNo = s.empNo;
      state.formTgtAddAmount = null;
      const ft = $("#form-target"); if (ft) ft.removeAttribute("readonly");
      syncOrgLabels();
      toast(`${s.name}(${empNo}) 정보를 불러왔습니다.`);
    });

    // 탭 전환
    $$(".tab").forEach((t) => t.addEventListener("click", () => switchTab(t.dataset.tab)));

    // 일괄등록 기수 선택
    document.getElementById("bulk-cohort-sel")?.addEventListener("change", (e) => {
      state._bulkCohort = e.target.value;
    });

    // 시드 파일 불러오기
    $("#btn-load-seed").addEventListener("click", openSeedPicker);

    // 일괄입력 미리보기
    document.getElementById("btn-bulk-preview")?.addEventListener("click", openBulkPreview);
    document.getElementById("btn-bulk-preview-close")?.addEventListener("click", () => {
      document.getElementById("bulk-preview-overlay").hidden = true;
    });
    document.getElementById("btn-bulk-preview-cancel")?.addEventListener("click", () => {
      document.getElementById("bulk-preview-overlay").hidden = true;
    });
    document.getElementById("btn-bulk-preview-save")?.addEventListener("click", async () => {
      const overlay = document.getElementById("bulk-preview-overlay");
      const hasEditable = !!overlay?.querySelector(".bulk-editable-row");
      overlay.hidden = true;
      try {
        const { ok, fail } = hasEditable
          ? await _saveBulkFromPreview()
          : await saveBulk(state._bulkColOverride || {});
        if (ok > 0 && fail === 0) closeModal("#modal-add");
      } catch (err) {
        console.error(err);
        toast("저장 실패: " + err.message, "error");
        setBulkProgress("저장 실패: " + err.message, "error");
      }
    });

    // 저장
    $("#btn-save").addEventListener("click", async () => {
      const activeTab = document.querySelector(".tab.active").dataset.tab;
      try {
        if (activeTab === "single") {
          const ok = await saveSingle();
          if (ok) {
            closeModal("#modal-add");
            const pgFull = document.getElementById("pg-full-modal");
            if (pgFull) pgFull.hidden = true;
            toast("저장되었습니다.", "success");
          }
        } else {
          // 벌크: 결과를 모달 안에 표시하고, 전부 성공할 때만 자동 닫음
          const { ok, fail } = await saveBulk();
          if (ok > 0 && fail === 0) {
            closeModal("#modal-add");
          }
        }
      } catch (err) {
        console.error(err);
        toast("저장 실패: " + err.message, "error");
        setBulkProgress("저장 실패: " + err.message, "error");
      }
    });

    // CSV
    $("#btn-export-csv").addEventListener("click", exportCSV);

    // 통계 — 시상안 출력 (현재 필터 범위 전체 교육생)
    const awardBtn = $("#btn-print-awards");
    if (awardBtn) awardBtn.addEventListener("click", printAwardSheets);

    // 실적진도 — 지역단 선택
    document.getElementById("pg-region-sel")?.addEventListener("change", (e) => {
      state.progressRegion = e.target.value;
      if (isPanelVisible("progress-panel")) renderProgressPanel();
    });
    // 실적진도 — 시상내역 엑셀 저장하기
    document.getElementById("btn-pg-excel")?.addEventListener("click", exportProgressAwardExcel);

    // 설정 탭 / 푸터 / 헤더 — 앱 버전 (커밋마다 +0.01)
    const v = $("#app-version"); if (v) v.textContent = `v${APP_VERSION} (build 20260610q)`;
    const fv = $("#app-footer-ver"); if (fv) fv.textContent = APP_VERSION;
    const hv = $("#app-header-ver"); if (hv) hv.textContent = APP_VERSION;
    $("#btn-open-backup-modal").addEventListener("click", openBackupModal);
    $("#btn-open-target-goals")?.addEventListener("click", openTargetGoalsModal);
    $("#btn-open-award-plan")?.addEventListener("click", () =>
      openAwardPlanModal({ region: state.filter.region, cohort: state.filter.cohort, step: state.filter.step || "1" })
    );
    bindAwardPlanModal();
    document.getElementById("pg-cohort-sel")?.addEventListener("change", (e) => {
      const prevCohortNum = parseInt(state.filter.cohort, 10) || 0;  // "2기" → 2
      const newCohortNum  = parseInt(e.target.value, 10) || 0;       // "1" → 1
      state.progressCohort = e.target.value;
      state.filter.cohort  = e.target.value ? `${e.target.value}기` : "";
      // 사이드바 filter-cohort 동기화
      const sidebarCohort = document.getElementById("filter-cohort");
      if (sidebarCohort) sidebarCohort.value = state.filter.cohort;
      // 자동 스텝 전환: 상위기수(낮은 번호) 선택 → Step 2 / 하위기수(높은 번호) 선택 → Step 1
      const autoStep = (newCohortNum > 0 && prevCohortNum > 0 && newCohortNum < prevCohortNum) ? "2" : "1";
      state.filter.step  = autoStep;
      state.progressStep = autoStep;
      const fsEl = document.getElementById("filter-step");
      if (fsEl) fsEl.value = autoStep;
      const pgStepSel = document.getElementById("pg-step-sel");
      if (pgStepSel) pgStepSel.value = autoStep;
      persistFilter();
      if (isPanelVisible("progress-panel")) renderProgressPanel();
      renderHomeRanks();
    });
    document.getElementById("pg-step-sel")?.addEventListener("change", (e) => {
      state.progressStep = e.target.value;
      state.filter.step = e.target.value;
      persistFilter();
      const filterStep = document.getElementById("filter-step");
      if (filterStep) filterStep.value = e.target.value;
      if (isPanelVisible("progress-panel")) renderProgressPanel();
      renderHomeRanks();
    });
    $("#btn-import-json").addEventListener("click", () => $("#file-import-json").click());
    $("#btn-go-progress-admin")?.addEventListener("click", openProgressAdminOverlay);
    $("#btn-pg-admin-overlay-close")?.addEventListener("click", () => {
      const ov = document.getElementById("pg-admin-overlay");
      if (ov) ov.hidden = true;
    });
    $("#file-import-json").addEventListener("change", onImportJSONFile);
    $("#btn-delete-filtered").addEventListener("click", onDeleteFiltered);
    $("#btn-set-cohort-1")?.addEventListener("click", onSetCohort1);
    const openSdBtn = $("#btn-open-student-delete");
    if (openSdBtn) openSdBtn.addEventListener("click", openStudentDeleteModal);
    const sdSearch = $("#sd-search");
    if (sdSearch) sdSearch.addEventListener("input", (e) => renderStudentDeleteList(e.target.value));
    if (sdSearch) sdSearch.addEventListener("click", (e) => {
      if (e.target.value) { e.target.value = ""; e.target.dispatchEvent(new Event("input")); }
    });
    const sdClear = $("#btn-sd-clear");
    if (sdClear) sdClear.addEventListener("click", () => {
      state.sdSelected = new Set();
      renderStudentDeleteList($("#sd-search").value);
      updateSdCounts();
    });
    const sdDelete = $("#btn-sd-delete");
    if (sdDelete) sdDelete.addEventListener("click", doStudentsDeleteFromModal);
  }

  // ========== 시상안 편집 ==========
  async function openAwardPlanModal(opts = {}) {
    const modal = document.getElementById("modal-award-plan");
    const regionSel = document.getElementById("award-plan-region");
    const regions = [...new Set(state.students.map((s) => s.region).filter((r) => r && (r.endsWith("지역단") || r.endsWith("사업부"))))].sort();
    const current = opts.region || state.progressRegion || state.filter.region || (regions[0] || "");
    regionSel.innerHTML = regions.map((r) =>
      `<option value="${escapeHtml(r)}"${r === current ? " selected" : ""}>${escapeHtml(r)}</option>`
    ).join("");
    // Sync year selector to current progress year
    const yearSel = document.getElementById("award-plan-year");
    if (yearSel && state.progressYear) yearSel.value = state.progressYear;
    // Sync cohort selector to current progress cohort
    const cohortSel = document.getElementById("award-plan-cohort");
    if (cohortSel && state.progressCohort) cohortSel.value = state.progressCohort;
    // Sync step selector
    const stepSel = document.getElementById("award-plan-step");
    if (stepSel && state.progressStep) stepSel.value = state.progressStep;
    // 마지막 선택값 복원
    try {
      const last = JSON.parse(localStorage.getItem(LS_AP_LAST_SEL) || "null");
      if (last) {
        if (last.year && yearSel) yearSel.value = last.year;
        if (last.region && regions.includes(last.region)) regionSel.value = last.region;
        if (last.cohort && cohortSel) cohortSel.value = last.cohort;
        if (last.step && stepSel) stepSel.value = last.step;
      }
    } catch {}
    // 호출자 지정 값이 최우선 (필터/rm모달 값)
    if (opts.region && regions.includes(opts.region)) regionSel.value = opts.region;
    if (opts.cohort && cohortSel) cohortSel.value = String(opts.cohort).replace(/기$/, "");
    if (opts.step && stepSel) stepSel.value = opts.step;
    // 로컬 캐시가 없으면 모달 열기 전에 Firestore에서 먼저 가져옴
    const cached = JSON.parse(localStorage.getItem(LS_AWARD_PLANS_KEY) || "{}");
    if (!Object.keys(cached).length && window.DataAPI?.loadAwardPlans) {
      try {
        const fsPlans = await window.DataAPI.loadAwardPlans();
        if (fsPlans && Object.keys(fsPlans).length) {
          localStorage.setItem(LS_AWARD_PLANS_KEY, JSON.stringify(fsPlans));
        }
      } catch (e) {
        console.warn("[AwardPlan] 모달 열기 전 Firestore 로드 실패:", e);
      }
    }
    _apRefreshFromSelectors();
    [yearSel, cohortSel, stepSel, regionSel].forEach((sel) => {
      if (sel) sel.onchange = () => { _apSaveLastSel(); _apMaybeSaveConfirm(_apRefreshFromSelectors); };
    });
    // 모달 열 때 오류 배너 초기화
    const errBanner = document.getElementById("ap-save-error-banner");
    if (errBanner) errBanner.hidden = true;
    openModal("#modal-award-plan");
  }

  function makeAwardPlanKey(year, region, cohort, step) {
    return `AP:${year}:${region}:${cohort}:${step}`;
  }

  function _apGetCurrentKey() {
    const year   = document.getElementById("award-plan-year")?.value   || "";
    const region = document.getElementById("award-plan-region")?.value || "";
    const cohort = document.getElementById("award-plan-cohort")?.value || "";
    const step   = document.getElementById("award-plan-step")?.value   || "";
    return makeAwardPlanKey(year, region, cohort, step);
  }

  function _apGenerateTitle() {
    const year   = document.getElementById("award-plan-year")?.value   || "";
    const region = document.getElementById("award-plan-region")?.value || "";
    const cohort = document.getElementById("award-plan-cohort")?.value || "";
    const step   = document.getElementById("award-plan-step")?.value   || "";
    const centers = [...new Set(
      state.students
        .filter((s) => s.region === region &&
          String(s.cohort || "").replace("기", "") === cohort)
        .map((s) => s.center).filter(Boolean)
    )];
    const centerStr = centers.length ? `(${centers.join(",")})` : "";
    const title = `${year}년 ${region} ${cohort}기${centerStr} Step${step} 활성화 시상안`;
    const el = document.getElementById("ap-title-display");
    if (el) el.textContent = title;
    return title;
  }

  let _apPendingAction = null;
  let _apOriginalJSON = null;
  let _apDirectSaveMode = false;
  const LS_AP_LAST_SEL = "cmf.apLastSel";

  function _apIsDirty() {
    if (!_apOriginalJSON) return false;
    try { return JSON.stringify(_apCollect()) !== _apOriginalJSON; } catch { return false; }
  }

  function _apSaveLastSel() {
    try {
      localStorage.setItem(LS_AP_LAST_SEL, JSON.stringify({
        year:   document.getElementById("award-plan-year")?.value   || "",
        region: document.getElementById("award-plan-region")?.value || "",
        cohort: document.getElementById("award-plan-cohort")?.value || "",
        step:   document.getElementById("award-plan-step")?.value   || ""
      }));
    } catch {}
  }

  function _apShowSaveConfirm(planTitle, msgText, noLabel) {
    document.getElementById("ap-save-confirm-plan-title").textContent = `"${planTitle}"`;
    document.getElementById("ap-save-confirm-msg").textContent = msgText;
    document.getElementById("ap-save-confirm-no").textContent = noLabel;
    document.getElementById("ap-save-confirm").hidden = false;
  }

  function _apMaybeSaveConfirm(onProceed) {
    if (!_apIsDirty()) { onProceed(); return; }
    _apPendingAction = onProceed;
    const title = document.getElementById("ap-title-display")?.textContent || "시상안";
    _apShowSaveConfirm(title, "변경사항이 있습니다. 저장하시겠습니까?", "저장없이 닫기");
  }

  function _apRenderPlanSummary(plan) {
    let html = `<div class="ap-psum-title">${escapeHtml(plan.title || "(제목 없음)")}</div><ul class="ap-psum-list">`;
    if (plan.personalIncr?.enabled) {
      const items = plan.personalIncr.items || [];
      const desc = items.map(it => {
        const val = it.payType === "item" ? it.payVal : (it.payVal + (it.payType === "pct" ? "%" : "만원"));
        return `${it.critVal}만↑→${val}`;
      }).join(" / ");
      html += `<li>📌 개인순증: ${escapeHtml(desc)}</li>`;
    }
    ["topAward1", "topAward2"].forEach(k => {
      const t = plan[k];
      if (!t?.enabled) return;
      const ps = (t.payouts || []).slice(0, 3).map((p, i) => `${i+1}위 ${payoutLabel(p)}`).join(" · ");
      const more = (t.payouts?.length || 0) > 3 ? ` 외 ${t.payouts.length - 3}개` : "";
      html += `<li>${k === "topAward1" ? "🥇" : "🥈"} 신장${t.type === "rate" ? "률" : "액"} Top${t.n}: ${escapeHtml(ps + more)}</li>`;
    });
    if (plan.groupAward1?.enabled) {
      const items = _ga1Items(plan.groupAward1);
      const desc = items.map(it => `${it.threshold}만↑→${payoutLabel(it.payout)}/인`).join(" / ");
      html += `<li>🏅 그룹시상1: ${escapeHtml(desc)}</li>`;
    }
    if (plan.groupAward2?.enabled) {
      const items = plan.groupAward2?.items?.length ? plan.groupAward2.items
        : [{ rateThreshold: plan.groupAward2?.rateThreshold ?? 110, payout: plan.groupAward2?.payout ?? 15 }];
      const desc = items.map(it => `${it.rateThreshold}%↑→${payoutLabel(it.payout)}`).join(" / ");
      html += `<li>🏅 그룹시상2: ${escapeHtml(desc)}</li>`;
    }
    if (plan.notes) html += `<li>📝 ${escapeHtml(plan.notes)}</li>`;
    html += "</ul>";
    return html;
  }

  function _apRenderLoadList(stored) {
    const keys = Object.keys(stored);
    const el = document.getElementById("ap-load-list");
    if (!keys.length) {
      el.innerHTML = `<div class="ap-load-empty">저장된 시상안이 없습니다.</div>`;
    } else {
      el.innerHTML = keys.map(k => {
        const plan = stored[k] || {};
        let label = k;
        if (k.startsWith("AP:")) {
          const parts = k.split(":");
          label = `${parts[2]} ${parts[3]}기 Step${parts[4]} (${parts[1]}년)`;
        }
        return `<div class="ap-load-item" data-key="${escapeHtml(k)}">
          <span class="ap-load-item-lbl">${escapeHtml(label)}</span>
          <span class="ap-load-item-title">${escapeHtml(plan.title || "")}</span>
        </div>`;
      }).join("");
      el.querySelectorAll(".ap-load-item").forEach(item => {
        item.addEventListener("click", () => _apShowLoadPreview(item.dataset.key, stored[item.dataset.key]));
      });
    }
  }

  async function _apOpenLoadPopup() {
    const el = document.getElementById("ap-load-list");
    document.getElementById("ap-load-popup").hidden = false;

    // Firestore가 정본 — 항상 최신 데이터를 조회
    if (window.DataAPI?.loadAwardPlans) {
      el.innerHTML = `<div class="ap-load-empty">Firestore에서 불러오는 중...</div>`;
      try {
        const fsPlans = await window.DataAPI.loadAwardPlans();
        if (fsPlans && Object.keys(fsPlans).length) {
          localStorage.setItem(LS_AWARD_PLANS_KEY, JSON.stringify(fsPlans));
          _apRenderLoadList(fsPlans);
          return;
        }
      } catch (e) {
        console.warn("[AwardPlan] Firestore 불러오기 실패, 로컬 캐시 사용:", e);
      }
    }
    // Firestore 미연결이거나 데이터 없음 → 로컬 캐시 표시
    let stored = {};
    try { stored = JSON.parse(localStorage.getItem(LS_AWARD_PLANS_KEY) || "{}"); } catch {}
    _apRenderLoadList(stored);
  }

  function _apShowLoadPreview(key, plan) {
    document.getElementById("ap-load-preview-body").innerHTML = _apRenderPlanSummary(plan || {});
    document.getElementById("ap-load-preview").hidden = false;
    document.getElementById("ap-load-preview-select").onclick = () => {
      document.getElementById("ap-load-preview").hidden = true;
      document.getElementById("ap-load-popup").hidden = true;
      loadAwardPlanForm(key);
      toast("시상안을 불러왔습니다.", "success");
    };
    document.getElementById("ap-load-preview-cancel").onclick = () => {
      document.getElementById("ap-load-preview").hidden = true;
    };
  }

  async function _apRefreshFromSelectors() {
    const key = _apGetCurrentKey();
    _apGenerateTitle();
    // localStorage에 없으면 Firestore에서 전체 재로드 시도
    if (key) {
      let stored = {};
      try { stored = JSON.parse(localStorage.getItem(LS_AWARD_PLANS_KEY) || "{}"); } catch {}
      if (!stored[key] && window.DataAPI?.loadAwardPlans) {
        try {
          const fsPlans = await window.DataAPI.loadAwardPlans();
          if (fsPlans && Object.keys(fsPlans).length) {
            const merged = { ...stored, ...fsPlans };
            localStorage.setItem(LS_AWARD_PLANS_KEY, JSON.stringify(merged));
          }
        } catch (e) {
          console.warn("[AwardPlan] 셀렉터 변경 시 Firestore 로드 실패:", e);
        }
      }
    }
    loadAwardPlanForm(key);
  }

  function _apRenderGa1(items) {
    const el = document.getElementById("ap-ga1-list");
    if (!el) return;
    el.innerHTML = items.map((it, i) => {
      const np = normPayout(it.payout ?? 5);
      const isCash = np.type !== "item";
      return `
      <div class="ap-row ap-ga2-row ap-ga1-row" data-i="${i}">
        <span class="ap-row-prefix">팀원 전원</span>
        <input type="number" class="pg-input ap-ga1-thr" value="${it.threshold || 5}" min="1" max="500" step="5" style="width:60px;" placeholder="5">
        <span class="ap-row-suffix">만원↑ 달성 시</span>
        <button type="button" class="ap-toggle-btn ap-pay-type" data-val="${isCash ? "cash" : "item"}">${isCash ? "만원" : "물품"}</button>
        <input type="number" class="pg-input ap-pay-cash ap-ga1-pay" value="${isCash ? escapeHtml(String(np.val)) : 5}" min="1" max="500" step="5" style="width:60px;${isCash ? "" : "display:none"}" placeholder="5">
        <input type="text" class="pg-input ap-pay-item" value="${isCash ? "" : escapeHtml(String(np.val || ""))}" placeholder="물품명" style="width:100px;${isCash ? "display:none" : ""}">
        <span class="ap-row-suffix">지급/인</span>
        <button type="button" class="ap-del-btn" title="삭제">✕</button>
      </div>`;
    }).join("");
  }

  function _apRenderGa2(items) {
    const el = document.getElementById("ap-ga2-list");
    if (!el) return;
    el.innerHTML = items.map((it, i) => {
      const np = normPayout(it.payout ?? 15);
      const isCash = np.type !== "item";
      return `
      <div class="ap-row ap-ga2-row" data-i="${i}">
        <span class="ap-row-prefix">팀달성률</span>
        <input type="number" class="pg-input ap-ga2-rate" value="${it.rateThreshold || 110}" min="100" max="300" step="5" style="width:70px;" placeholder="110">
        <span class="ap-row-suffix">%↑ 달성 시</span>
        <button type="button" class="ap-toggle-btn ap-pay-type" data-val="${isCash ? "cash" : "item"}">${isCash ? "만원" : "물품"}</button>
        <input type="number" class="pg-input ap-pay-cash ap-ga2-pay" value="${isCash ? escapeHtml(String(np.val)) : 15}" min="1" max="500" step="5" style="width:60px;${isCash ? "" : "display:none"}" placeholder="15">
        <input type="text" class="pg-input ap-pay-item" value="${isCash ? "" : escapeHtml(String(np.val || ""))}" placeholder="물품명" style="width:80px;${isCash ? "display:none" : ""}">
        <span class="ap-row-suffix">지급</span>
        <button type="button" class="ap-del-btn" title="삭제">✕</button>
      </div>`;
    }).join("");
  }

  function _apRenderPi(items) {
    const el = document.getElementById("ap-pi-list");
    const sorted = [...items].sort((a, b) => a.critVal - b.critVal);
    const rateOpts = [
      { v: "full", t: "100%" },
      { v: "half", t: "50%" },
      { v: "none", t: "부지급" }
    ];
    el.innerHTML = sorted.map((it, i) => {
      const typ = it.payType || "fixed";
      const isItem = typ === "item";
      const rate = it.payRate || "full";
      const typeLabel = typ === "pct" ? "%" : (isItem ? "물품" : "만");
      const optHtml = rateOpts.map((o) => `<option value="${o.v}"${rate === o.v ? " selected" : ""}>${o.t}</option>`).join("");
      return `<div class="ap-pi-card" data-i="${i}">
        <div class="ap-pi-row">
          <span class="ap-pi-lbl">순증</span>
          <input type="number" class="pg-input ap-pi-crit" value="${escapeHtml(String(it.critVal))}" min="0" step="10" style="flex:1;min-width:0;">
          <span class="ap-pi-unit">만↑</span>
        </div>
        <div class="ap-pi-row">
          <span class="ap-pi-lbl">시상</span>
          <input type="number" class="pg-input ap-pi-pay" value="${isItem ? 0 : escapeHtml(String(it.payVal))}" min="0" step="1" style="flex:1;min-width:0;${isItem ? "display:none" : ""}">
          <input type="text" class="pg-input ap-pi-pay-item" value="${isItem ? escapeHtml(String(it.payVal || "")) : ""}" placeholder="물품명" style="flex:1;min-width:0;${isItem ? "" : "display:none"}">
          <button type="button" class="ap-toggle-btn ap-pi-type ap-pi-type-sm" data-val="${escapeHtml(typ)}">${typeLabel}</button>
        </div>
        <div class="ap-pi-footer">
          <select class="pg-input ap-pi-payrate"${isItem ? ' style="display:none"' : ""}>${optHtml}</select>
          <button type="button" class="ap-del-btn" title="삭제">✕</button>
        </div>
      </div>`;
    }).join("");
  }

  function _apRenderTop(slot, payouts) {
    const el = document.getElementById(`ap-${slot}-list`);
    el.innerHTML = payouts.map((p, i) => {
      const np = normPayout(p);
      const isCash = np.type !== "item";
      return `
      <div class="ap-row ap-row-compact" data-i="${i}">
        <span class="ap-row-prefix">${i + 1}위 시상</span>
        <button type="button" class="ap-toggle-btn ap-pay-type" data-val="${isCash ? "cash" : "item"}">${isCash ? "현금" : "물품"}</button>
        <input type="number" class="pg-input ap-pay-cash" value="${isCash ? escapeHtml(String(np.val)) : 0}" min="0" step="1" style="width:52px;${isCash ? "" : "display:none"}">
        <span class="ap-pay-unit" style="${isCash ? "" : "display:none"}">만원</span>
        <input type="text" class="pg-input ap-pay-item" value="${isCash ? "" : escapeHtml(String(np.val || ""))}" placeholder="물품명" style="width:80px;${isCash ? "display:none" : ""}">
        <button type="button" class="ap-del-btn" title="삭제">✕</button>
      </div>`;
    }).join("");
  }

  function _apUpdateMinnetState(slot) {
    const en = document.getElementById(`ap-${slot}-minnet-en`)?.checked;
    const cond = document.getElementById(`ap-${slot}-minnet-cond`);
    if (cond) cond.classList.toggle("ap-disabled", !en);
  }

  function _apRenderElig(conds) {
    const el = document.getElementById("ap-elig-list");
    const fieldOpts = [
      { v: "converted", t: "환산실적" },
      { v: "hiCap",     t: "하이캡" },
      { v: "monthly",   t: "월납보험료" }
    ];
    el.innerHTML = conds.map((c, i) => `
      <div class="ap-row" data-i="${i}">
        <select class="pg-input ap-elig-field" style="width:130px;">
          ${fieldOpts.map((o) => `<option value="${o.v}"${c.field === o.v ? " selected" : ""}>${o.t}</option>`).join("")}
        </select>
        <input type="number" class="pg-input ap-elig-th" value="${escapeHtml(String(c.threshold))}" min="0" step="1" style="width:100px;">
        <span class="ap-row-suffix">이하 제외</span>
        <button type="button" class="ap-del-btn" title="삭제">✕</button>
      </div>`).join("");
  }

  function _apSetToggle(btn, val, choices) {
    btn.dataset.val = val;
    const found = choices.find((c) => c.v === val);
    btn.textContent = found ? found.t : val;
  }

  function loadAwardPlanForm(key) {
    // 저장 여부 확인 → 배너 표시/숨김
    const noBanner   = document.getElementById("ap-no-plan-banner");
    const noBannerKey = document.getElementById("ap-no-plan-key");
    if (key) {
      let stored = {};
      try { stored = JSON.parse(localStorage.getItem(LS_AWARD_PLANS_KEY) || "{}"); } catch {}
      const hasSaved = !!stored[key];
      if (noBanner)    noBanner.hidden = hasSaved;
      if (noBannerKey && !hasSaved) noBannerKey.textContent = key.replace(/^AP:/, "").replace(/:/g, " / ");
    } else {
      if (noBanner) noBanner.hidden = true;
    }
    const plan = key ? getAwardPlan(key) : JSON.parse(JSON.stringify(DEFAULT_AWARD_PLAN));
    document.getElementById("ap-notes").value = plan.notes || "";
    // 개인순증시상
    document.getElementById("ap-pi-en").checked = !!plan.personalIncr?.enabled;
    _apRenderPi((plan.personalIncr?.items?.length ? plan.personalIncr.items : [{ critVal: 5, payType: "fixed", payVal: 5 }]));
    // Top1
    document.getElementById("ap-t1-en").checked = !!plan.topAward1?.enabled;
    _apSetToggle(document.getElementById("ap-t1-type"), plan.topAward1?.type || "rate",
      [{ v: "rate", t: "률(%)" }, { v: "amt", t: "금액(원)" }]);
    document.getElementById("ap-t1-n").value = plan.topAward1?.n || 10;
    document.getElementById("ap-t1-minnet-en").checked = !!plan.topAward1?.minNetEnabled;
    document.getElementById("ap-t1-minnet").value = plan.topAward1?.minNet ?? 300000;
    _apRenderTop("t1", (plan.topAward1?.payouts?.length ? plan.topAward1.payouts : [30]));
    // Top2
    document.getElementById("ap-t2-en").checked = !!plan.topAward2?.enabled;
    _apSetToggle(document.getElementById("ap-t2-type"), plan.topAward2?.type || "amt",
      [{ v: "rate", t: "률(%)" }, { v: "amt", t: "금액(원)" }]);
    document.getElementById("ap-t2-n").value = plan.topAward2?.n || 10;
    document.getElementById("ap-t2-minnet-en").checked = !!plan.topAward2?.minNetEnabled;
    document.getElementById("ap-t2-minnet").value = plan.topAward2?.minNet ?? 300000;
    _apRenderTop("t2", (plan.topAward2?.payouts?.length ? plan.topAward2.payouts : [50]));
    // 순증 기준 활성 상태 초기화
    _apUpdateMinnetState("t1");
    _apUpdateMinnetState("t2");
    // 중복시상 불가 체크박스
    document.getElementById("ap-both-nodup").checked = !!plan.bothNodup;
    // Eligibility
    document.getElementById("ap-elig-en").checked = !!plan.eligibility?.enabled;
    _apSetToggle(document.getElementById("ap-elig-op"), plan.eligibility?.operator || "and",
      [{ v: "and", t: "AND" }, { v: "or", t: "OR" }]);
    _apRenderElig((plan.eligibility?.conditions?.length ? plan.eligibility.conditions : [{ field: "converted", threshold: 80 }]));
    document.getElementById("ap-ga1-en").checked = !!plan.groupAward1?.enabled;
    // ga1 items (구형: threshold+payout → items 배열로 변환)
    const ga1RawItems = plan.groupAward1?.items?.length
      ? plan.groupAward1.items
      : [{ threshold: plan.groupAward1?.threshold ?? 5, payout: normPayout(plan.groupAward1?.payout ?? 5) }];
    _apRenderGa1(ga1RawItems);
    document.getElementById("ap-ga2-en").checked   = !!plan.groupAward2?.enabled;
    const ga2LinkEl = document.getElementById("ap-ga2-link");
    if (ga2LinkEl) ga2LinkEl.checked = !!plan.groupAward2?.linkToGroup1;
    const ga2Items = plan.groupAward2?.items?.length
      ? plan.groupAward2.items
      : [{ rateThreshold: plan.groupAward2?.rateThreshold ?? 110, payout: plan.groupAward2?.payout ?? 15 }];
    _apRenderGa2(ga2Items);
    // Snapshot for dirty check
    _apOriginalJSON = JSON.stringify(_apCollect());
  }

  function _apCollect() {
    // Personal increment items
    const piItems = [];
    document.querySelectorAll("#ap-pi-list .ap-pi-card").forEach((card) => {
      const crit = Number(card.querySelector(".ap-pi-crit").value) || 0;
      const typ  = card.querySelector(".ap-pi-type").dataset.val;
      const rate = card.querySelector(".ap-pi-payrate")?.value || "full";
      const pay  = typ === "item"
        ? (card.querySelector(".ap-pi-pay-item")?.value.trim() || "")
        : (Number(card.querySelector(".ap-pi-pay")?.value) || 0);
      piItems.push({ critVal: crit, payType: typ, payVal: pay, payRate: rate });
    });
    const topPayouts = (slot) => {
      const arr = [];
      document.querySelectorAll(`#ap-${slot}-list .ap-row`).forEach((row) => {
        const type = row.querySelector(".ap-pay-type")?.dataset.val || "cash";
        if (type === "item") {
          arr.push({ type: "item", val: row.querySelector(".ap-pay-item")?.value.trim() || "" });
        } else {
          arr.push({ type: "cash", val: Number(row.querySelector(".ap-pay-cash")?.value) || 0 });
        }
      });
      return arr;
    };
    const eligConds = [];
    document.querySelectorAll("#ap-elig-list .ap-row").forEach((row) => {
      eligConds.push({
        field: row.querySelector(".ap-elig-field").value,
        threshold: Number(row.querySelector(".ap-elig-th").value) || 0
      });
    });
    return {
      title: document.getElementById("ap-title-display")?.textContent?.trim() || "",
      personalIncr: {
        enabled: document.getElementById("ap-pi-en").checked,
        items: piItems
      },
      topAward1: {
        enabled: document.getElementById("ap-t1-en").checked,
        type: document.getElementById("ap-t1-type").dataset.val,
        n: Number(document.getElementById("ap-t1-n").value) || 1,
        minNetEnabled: document.getElementById("ap-t1-minnet-en").checked,
        minNet: Number(document.getElementById("ap-t1-minnet").value) || 0,
        payouts: topPayouts("t1")
      },
      topAward2: {
        enabled: document.getElementById("ap-t2-en").checked,
        type: document.getElementById("ap-t2-type").dataset.val,
        n: Number(document.getElementById("ap-t2-n").value) || 1,
        minNetEnabled: document.getElementById("ap-t2-minnet-en").checked,
        minNet: Number(document.getElementById("ap-t2-minnet").value) || 0,
        payouts: topPayouts("t2")
      },
      bothNodup: document.getElementById("ap-both-nodup").checked,
      eligibility: {
        enabled: document.getElementById("ap-elig-en").checked,
        operator: document.getElementById("ap-elig-op").dataset.val,
        conditions: eligConds
      },
      groupAward1: {
        enabled: document.getElementById("ap-ga1-en")?.checked ?? false,
        items: (() => {
          const arr = [];
          document.querySelectorAll("#ap-ga1-list .ap-ga1-row").forEach((row) => {
            const type = row.querySelector(".ap-pay-type")?.dataset.val || "cash";
            arr.push({
              threshold: Number(row.querySelector(".ap-ga1-thr")?.value) || 5,
              payout: type === "item"
                ? { type: "item", val: row.querySelector(".ap-pay-item")?.value.trim() || "" }
                : { type: "cash", val: Number(row.querySelector(".ap-pay-cash")?.value) || 5 }
            });
          });
          return arr.length ? arr : [{ threshold: 5, payout: { type: "cash", val: 5 } }];
        })()
      },
      groupAward2: {
        enabled: document.getElementById("ap-ga2-en")?.checked ?? false,
        linkToGroup1: document.getElementById("ap-ga2-link")?.checked ?? false,
        items: (() => {
          const arr = [];
          document.querySelectorAll("#ap-ga2-list .ap-ga2-row").forEach((row) => {
            const type = row.querySelector(".ap-pay-type")?.dataset.val || "cash";
            arr.push({
              rateThreshold: Number(row.querySelector(".ap-ga2-rate")?.value) || 110,
              payout: type === "item"
                ? { type: "item", val: row.querySelector(".ap-pay-item")?.value.trim() || "" }
                : { type: "cash", val: Number(row.querySelector(".ap-pay-cash")?.value) || 15 }
            });
          });
          return arr.length ? arr : [{ rateThreshold: 110, payout: { type: "cash", val: 15 } }];
        })()
      },
      notes: document.getElementById("ap-notes").value.trim()
    };
  }

  function bindAwardPlanModal() {
    const modal = document.getElementById("modal-award-plan");
    if (!modal) return;

    // 토글 버튼들 (type, op)
    modal.addEventListener("click", (e) => {
      const tg = e.target.closest(".ap-toggle-btn");
      if (tg) {
        if (tg.id === "ap-t1-type" || tg.id === "ap-t2-type") {
          const next = tg.dataset.val === "rate" ? "amt" : "rate";
          _apSetToggle(tg, next, [{ v: "rate", t: "률(%)" }, { v: "amt", t: "금액(원)" }]);
        } else if (tg.id === "ap-elig-op") {
          const next = tg.dataset.val === "and" ? "or" : "and";
          _apSetToggle(tg, next, [{ v: "and", t: "AND" }, { v: "or", t: "OR" }]);
        } else if (tg.classList.contains("ap-pay-type")) {
          // 현금 ↔ 물품 토글
          const newType = tg.dataset.val === "item" ? "cash" : "item";
          tg.dataset.val = newType;
          tg.textContent = newType === "item" ? "물품" : "현금";
          const container = tg.closest(".ap-row, .ap-ga-row, .ap-ga1-row, .ap-ga2-row");
          const cashInp = container?.querySelector(".ap-pay-cash");
          const unitSpan = container?.querySelector(".ap-pay-unit");
          const itemInp = container?.querySelector(".ap-pay-item");
          if (cashInp) cashInp.style.display = newType === "item" ? "none" : "";
          if (unitSpan) unitSpan.style.display = newType === "item" ? "none" : "";
          if (itemInp) itemInp.style.display = newType === "item" ? "" : "none";
        } else if (tg.classList.contains("ap-pi-type")) {
          // 3-way 순환: pct → fixed → item → pct
          const cur = tg.dataset.val;
          const next = cur === "pct" ? "fixed" : (cur === "fixed" ? "item" : "pct");
          const nextLabel = next === "pct" ? "%" : (next === "item" ? "물품" : "만");
          tg.dataset.val = next;
          tg.textContent = nextLabel;
          const card = tg.closest(".ap-pi-card");
          if (card) {
            const numInp = card.querySelector(".ap-pi-pay");
            const txtInp = card.querySelector(".ap-pi-pay-item");
            const rateEl = card.querySelector(".ap-pi-payrate");
            if (numInp) numInp.style.display = next === "item" ? "none" : "";
            if (txtInp) txtInp.style.display = next === "item" ? "" : "none";
            if (rateEl) rateEl.style.display = next === "item" ? "none" : "";
          }
        }
        return;
      }
      // 삭제 버튼
      const del = e.target.closest(".ap-del-btn");
      if (del) {
        const row = del.closest(".ap-row, .ap-pi-card");
        if (row) row.remove();
      }
    });

    // 순증 기준 체크박스 토글
    document.getElementById("ap-t1-minnet-en")?.addEventListener("change", () => _apUpdateMinnetState("t1"));
    document.getElementById("ap-t2-minnet-en")?.addEventListener("change", () => _apUpdateMinnetState("t2"));

    // 추가 버튼들
    document.getElementById("ap-pi-add")?.addEventListener("click", () => {
      const cur = _apCollect().personalIncr.items;
      cur.push({ critVal: 0, payType: "fixed", payVal: 0, payRate: "full" });
      _apRenderPi(cur);
    });
    document.getElementById("ap-t1-add")?.addEventListener("click", () => {
      const cur = _apCollect().topAward1.payouts;
      cur.push(0);
      _apRenderTop("t1", cur);
    });
    document.getElementById("ap-t2-add")?.addEventListener("click", () => {
      const cur = _apCollect().topAward2.payouts;
      cur.push(0);
      _apRenderTop("t2", cur);
    });
    document.getElementById("ap-elig-add")?.addEventListener("click", () => {
      const cur = _apCollect().eligibility.conditions;
      cur.push({ field: "converted", threshold: 0 });
      _apRenderElig(cur);
    });
    document.getElementById("ap-ga1-add")?.addEventListener("click", () => {
      const cur = _apCollect().groupAward1.items || [];
      cur.push({ threshold: 5, payout: { type: "cash", val: 5 } });
      _apRenderGa1(cur);
    });
    document.getElementById("ap-ga2-add")?.addEventListener("click", () => {
      const cur = _apCollect().groupAward2.items || [];
      cur.push({ rateThreshold: 110, payout: { type: "cash", val: 15 } });
      _apRenderGa2(cur);
    });

    // 저장 확인 오버레이 버튼
    document.getElementById("ap-save-confirm-yes")?.addEventListener("click", async () => {
      document.getElementById("ap-save-confirm").hidden = true;
      // 저장 실행
      const region = document.getElementById("award-plan-region").value;
      if (region) {
        const plan = _apCollect();
        plan.title = _apGenerateTitle();
        let fsSaveErr = null;
        try {
          await saveAwardPlan(_apGetCurrentKey(), plan);
        } catch (e) {
          fsSaveErr = e;
          console.error("[AwardPlan] Firestore 저장 실패 (로컬 저장 완료):", e);
        }
        // localStorage에는 이미 저장됨 — no-plan 배너 즉시 숨김
        const noPlanBanner = document.getElementById("ap-no-plan-banner");
        if (noPlanBanner) noPlanBanner.hidden = true;
        _apOriginalJSON = JSON.stringify(_apCollect());
        _apSaveLastSel();
        renderDebounced();
        if (fsSaveErr) {
          // Firestore 실패 — 배너로 알리고 모달 유지
          const code = fsSaveErr?.code || "";
          const reason = code === "permission-denied"
            ? "권한 없음 — Firebase 보안 규칙 확인 필요"
            : code === "unavailable" || code === "deadline-exceeded"
            ? "서버 연결 불가"
            : code || fsSaveErr?.message || String(fsSaveErr) || "알 수 없는 오류";
          const banner = document.getElementById("ap-save-error-banner");
          const bannerMsg = document.getElementById("ap-save-error-msg");
          if (banner && bannerMsg) {
            bannerMsg.textContent = `로컬에는 저장되었으나 서버 동기화 실패: ${reason}`;
            banner.hidden = false;
          }
          _apDirectSaveMode = false;
          return; // 모달 유지, 후속 액션 취소
        }
        // 성공 — 배너 초기화 후 모달 닫기
        document.getElementById("ap-save-error-banner").hidden = true;
        toast(`시상안 저장 완료: ${plan.title}`, "success");
      }
      _apDirectSaveMode = false;
      document.getElementById("modal-award-plan").hidden = true;
      // 저장 후 원래 대기 중이던 액션 실행
      if (_apPendingAction) { const fn = _apPendingAction; _apPendingAction = null; fn(); }
    });
    document.getElementById("ap-save-confirm-no")?.addEventListener("click", () => {
      document.getElementById("ap-save-confirm").hidden = true;
      if (_apDirectSaveMode) {
        _apDirectSaveMode = false;
        return; // 직접 저장 취소 — dirty 유지, 후속 액션 없음
      }
      _apOriginalJSON = null; // dirty 초기화하여 재질문 방지
      if (_apPendingAction) { const fn = _apPendingAction; _apPendingAction = null; fn(); }
    });

    document.getElementById("btn-award-plan-save")?.addEventListener("click", () => {
      const region = document.getElementById("award-plan-region").value;
      if (!region) { toast("지역단을 선택하세요.", "error"); return; }
      const title = _apGenerateTitle();
      _apDirectSaveMode = true;
      _apPendingAction = null;
      _apShowSaveConfirm(title, "이곳에 저장하는 게 맞습니까?", "취소");
    });
    document.getElementById("btn-award-plan-load")?.addEventListener("click", _apOpenLoadPopup);
    document.querySelector("#ap-load-popup .ap-load-popup-close")?.addEventListener("click", () => {
      document.getElementById("ap-load-popup").hidden = true;
    });
    document.getElementById("btn-award-plan-reset")?.addEventListener("click", () => {
      if (!confirm("시상안을 기본값으로 초기화합니까?")) return;
      loadAwardPlanForm(null);
    });
    document.getElementById("btn-award-plan-close")?.addEventListener("click", () => {
      _apMaybeSaveConfirm(() => { modal.hidden = true; });
    });
    // 백드롭 클릭 시도 dirty 체크
    modal.querySelector(".modal-backdrop")?.addEventListener("click", (e) => {
      e.stopPropagation();
      _apMaybeSaveConfirm(() => { modal.hidden = true; });
    });
  }

  // ========== 설정 ==========
  function exportJSONForRegion(list, regionLabel) {
    if (!list.length) { toast("내보낼 데이터가 없습니다.", "error"); return; }
    const now = new Date();
    const dateStr = String(now.getFullYear()) +
      String(now.getMonth() + 1).padStart(2, "0") +
      String(now.getDate()).padStart(2, "0") +
      String(now.getHours()).padStart(2, "0") +
      String(now.getMinutes()).padStart(2, "0") +
      String(now.getSeconds()).padStart(2, "0");
    const filename = `${regionLabel}_${dateStr}_backup.json`;
    const payload = {
      exportedAt: now.toISOString(),
      region: regionLabel,
      count: list.length,
      students: list.map((s) => ({
        region: s.region || "", center: s.center || "", branch: s.branch || "",
        cohort: s.cohort || "", empNo: s.empNo, name: s.name || "", phone: s.phone || "",
        base: Number(s.base || 0), target: Number(s.target || 0), honors: Number(s.honors || 0),
        insAvg: Number(s.insAvg || 0), curAct: Number(s.curAct || 0)
      }))
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
    toast(`${list.length}건 백업 완료: ${filename}`, "success");
  }

  function openBackupModal() {
    const modal = document.getElementById("modal-backup");
    const body = document.getElementById("backup-modal-body");
    const regions = [...new Set(state.students.map((s) => s.region).filter((r) => r && r.endsWith("지역단")))].sort();
    body.innerHTML = `
      <div style="display:flex;flex-direction:column;gap:8px;">
        <div class="backup-row backup-row-all">
          <span style="font-weight:800;">📦 전체 (${state.students.length}명)</span>
          <button class="btn-primary small" data-br="전체">전체 백업</button>
        </div>
        ${regions.map((r) => {
          const cnt = state.students.filter((s) => s.region === r).length;
          return `<div class="backup-row">
            <span>${escapeHtml(r)} <small style="color:var(--ink-3);">(${cnt}명)</small></span>
            <button class="btn-outline small" data-br="${escapeHtml(r)}">백업</button>
          </div>`;
        }).join("")}
      </div>`;
    body.querySelectorAll("[data-br]").forEach((btn) => {
      btn.addEventListener("click", () => {
        const region = btn.dataset.br;
        const list = region === "전체" ? state.students : state.students.filter((s) => s.region === region);
        exportJSONForRegion(list, region);
      });
    });
    openModal("#modal-backup");
  }

  function showRestoreConfirm(arr, detectedRegion) {
    const modal = document.getElementById("modal-restore-confirm");
    const msgEl = document.getElementById("restore-confirm-msg");
    const regionWrap = document.getElementById("restore-region-select-wrap");
    const regionSelect = document.getElementById("restore-region-select");
    const yesBtn = document.getElementById("btn-restore-yes");
    const noBtn = document.getElementById("btn-restore-no");
    const cancelBtn = document.getElementById("btn-restore-cancel");

    let pendingRegion = detectedRegion;

    const populateRegionSelect = () => {
      const regions = ["전체", ...[...new Set(state.students.map((s) => s.region).filter((r) => r && r.endsWith("지역단")))].sort()];
      regionSelect.innerHTML = regions.map((r) => `<option value="${r}">${r}</option>`).join("");
      if (pendingRegion && regions.includes(pendingRegion)) regionSelect.value = pendingRegion;
    };

    if (detectedRegion) {
      msgEl.textContent = `"${detectedRegion}"으로 ${arr.length}건 복원하시겠습니까?`;
      regionWrap.hidden = true;
      noBtn.hidden = false;
      yesBtn.textContent = "예, 복원합니다";
    } else {
      msgEl.textContent = `${arr.length}건의 교육생 데이터를 복원합니다. 복원할 지역단을 선택하세요.`;
      populateRegionSelect();
      regionWrap.hidden = false;
      noBtn.hidden = true;
      yesBtn.textContent = "선택한 지역단으로 복원";
    }

    modal.hidden = false;

    yesBtn.onclick = async () => {
      const target = regionWrap.hidden ? (pendingRegion || "전체") : regionSelect.value;
      modal.hidden = true;
      await doRestore(arr, target);
    };

    noBtn.onclick = () => {
      msgEl.textContent = `${arr.length}건의 교육생 데이터를 복원합니다. 복원할 지역단을 선택하세요.`;
      populateRegionSelect();
      regionWrap.hidden = false;
      noBtn.hidden = true;
      yesBtn.textContent = "선택한 지역단으로 복원";
    };

    cancelBtn.onclick = () => { modal.hidden = true; };
  }

  async function doRestore(arr, targetRegion) {
    const backupList = (targetRegion && targetRegion !== "전체")
      ? state.students.filter((s) => s.region === targetRegion)
      : state.students;
    toast("만약을 위해 백업 파일을 만듭니다.", "");
    await new Promise((r) => setTimeout(r, 600));
    exportJSONForRegion(backupList, targetRegion || "전체");
    await new Promise((r) => setTimeout(r, 800));
    try {
      if (typeof window.DataAPI.saveMany === "function") {
        const { committed, errors } = await window.DataAPI.saveMany(arr);
        toast(`${committed}건 복원 완료${errors.length ? ` / ${errors.length}건 실패` : ""}`, errors.length ? "error" : "success");
      } else {
        let ok = 0, fail = 0;
        for (const s of arr) {
          try { await window.DataAPI.save(s); ok++; } catch { fail++; }
        }
        toast(`${ok}건 복원, ${fail}건 실패`, fail ? "error" : "success");
      }
    } catch (err) {
      console.error(err);
      toast("복원 실패: " + err.message, "error");
    }
  }

  async function onImportJSONFile(e) {
    const file = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!file) return;
    try {
      const text = await file.text();
      const data = JSON.parse(text);
      const arr = Array.isArray(data) ? data : (Array.isArray(data.students) ? data.students : null);
      if (!arr || !arr.length) { toast("유효한 students 배열이 없습니다.", "error"); return; }
      // 파일명에서 지역단 자동 감지: "지역단명_YYYYMMDDHHMMSS_backup.json"
      const match = file.name.match(/^(.+?)_\d{8,14}_backup\.json$/i);
      const detectedRegion = match ? match[1] : null;
      showRestoreConfirm(arr, detectedRegion);
    } catch (err) {
      console.error(err);
      toast("파일 파싱 실패: " + err.message, "error");
    }
  }

  async function onDeleteFiltered() {
    const list = filteredStudents();
    await confirmAndDeleteStudents(list, {
      label: `현재 필터(${[state.filter.region, state.filter.center, state.filter.branch, state.filter.cohort].filter(Boolean).join(" · ") || "전체"})`
    });
  }

  async function onSetCohort1() {
    const targets = state.students.filter((s) => !s.cohort);
    if (!targets.length) {
      toast("기수 미설정 교육생이 없습니다.", "");
      return;
    }
    // 지역단별 현황 요약
    const byRegion = {};
    targets.forEach((s) => { byRegion[s.region || "(미지정)"] = (byRegion[s.region || "(미지정)"] || 0) + 1; });
    const summary = Object.entries(byRegion).sort().map(([r, n]) => `${r} ${n}명`).join("\n");
    const ok = await openConfirmModal(`기수 미설정 교육생 총 ${targets.length}명을 "1기"로 저장합니다.\n\n${summary}\n\n진행하시겠습니까?`);
    if (!ok) return;

    const btn = document.getElementById("btn-set-cohort-1");
    if (btn) { btn.disabled = true; btn.textContent = "저장 중..."; }
    try {
      const records = targets.map((s) => ({ ...s, cohort: "1기" }));
      if (typeof window.DataAPI.saveMany === "function") {
        await window.DataAPI.saveMany(records);
      } else {
        for (const r of records) await window.DataAPI.save(r);
      }
      toast(`✅ ${targets.length}명의 기수를 "1기"로 저장했습니다.`, "");
    } catch (err) {
      toast("저장 오류: " + err.message, "error");
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = "📌 기수 미설정 → 1기 일괄 저장"; }
    }
  }

  // 공통: 2단계 확인 후 학생 + 면담 원자적 삭제
  async function confirmAndDeleteStudents(list, opts) {
    const lbl = opts?.label || "선택한 범위";
    if (!list.length) { toast("삭제 대상이 없습니다.", "error"); return false; }
    const msg = `경고: ${lbl}에 해당하는 ${list.length}명의 교육생과 그 면담 기록을 모두 삭제합니다.\n\n복구할 수 없습니다. 정말 삭제하시겠습니까?`;
    if (!confirm(msg)) return false;
    const typed = prompt(`다시 한 번 확인합니다. 삭제를 진행하려면 아래와 동일하게 입력하세요:\n\n삭제 ${list.length}명`);
    if (typed !== `삭제 ${list.length}명`) {
      toast("입력이 일치하지 않아 취소되었습니다.", "");
      return false;
    }
    let ok = 0, fail = 0;
    const useBatch = typeof window.DataAPI.removeStudentWithConsultations === "function";
    for (let i = 0; i < list.length; i++) {
      const s = list[i];
      try {
        if (useBatch) await window.DataAPI.removeStudentWithConsultations(s.empNo);
        else await window.DataAPI.remove(s.empNo);
        ok++;
        if (i % 5 === 4 || i === list.length - 1) {
          toast(`삭제중... ${i + 1}/${list.length}`, "");
        }
      } catch (e) { console.error(e); fail++; }
    }
    toast(`${ok}명 삭제 완료${fail ? ` / ${fail}명 실패` : ""}`, fail ? "error" : "success");
    return true;
  }

  // ========== 교육생 개별 선택 삭제 모달 ==========
  function openStudentDeleteModal() {
    state.sdSelected = new Set();
    $("#sd-search").value = "";
    renderStudentDeleteList("");
    openModal("#modal-student-delete");
    updateSdCounts();
    setTimeout(() => $("#sd-search").focus(), 60);
  }

  function updateSdCounts() {
    const total = state.students.length;
    const sel = state.sdSelected ? state.sdSelected.size : 0;
    $("#sd-total").textContent = total;
    $("#sd-sel-cnt").textContent = sel;
    const btn = $("#btn-sd-delete");
    if (btn) { btn.disabled = sel === 0; btn.textContent = sel > 0 ? `🗑️ ${sel}명 삭제` : "🗑️ 선택 교육생 삭제"; }
  }

  function renderStudentDeleteList(q) {
    const container = $("#sd-list");
    if (!container) return;
    const needle = (q || "").trim().toLowerCase();
    const all = state.students.slice();
    const filtered = needle
      ? all.filter((s) => [s.empNo, s.name, s.branch, s.center, s.region].join(" ").toLowerCase().includes(needle))
      : all;
    if (!filtered.length) {
      container.innerHTML = `<div class="empty-mini">일치하는 교육생이 없습니다.</div>`;
      return;
    }
    // 지역단 > 비전센터 > 지점 순 그룹핑
    const byRegion = {};
    filtered.forEach((s) => {
      const r = s.region || "(미지정)";
      const c = s.center || "(미지정)";
      if (!byRegion[r]) byRegion[r] = {};
      if (!byRegion[r][c]) byRegion[r][c] = [];
      byRegion[r][c].push(s);
    });
    const regions = Object.keys(byRegion).sort();
    container.innerHTML = regions.map((r) => {
      const centers = byRegion[r];
      const regionCount = Object.values(centers).reduce((a, arr) => a + arr.length, 0);
      const regionSelCount = Object.values(centers).flat().filter((s) => state.sdSelected.has(s.empNo)).length;
      return `
        <div class="sd-region">
          <div class="sd-region-head">
            <label class="sd-chk-region">
              <input type="checkbox" data-region="${escapeHtml(r)}" ${regionSelCount > 0 && regionSelCount === regionCount ? "checked" : ""}>
              <strong>${escapeHtml(r)}</strong>
              <span class="sd-count">${regionSelCount}/${regionCount}</span>
            </label>
          </div>
          ${Object.keys(centers).sort().map((c) => `
            <div class="sd-center">
              <div class="sd-center-head">↳ ${escapeHtml(c)}</div>
              <div class="sd-student-list">
                ${centers[c].slice().sort((a, b) => (a.name || "").localeCompare(b.name || "")).map((s) => `
                  <label class="sd-item ${state.sdSelected.has(s.empNo) ? "selected" : ""}">
                    <input type="checkbox" data-emp="${escapeHtml(s.empNo)}" ${state.sdSelected.has(s.empNo) ? "checked" : ""}>
                    <span class="sd-name">${escapeHtml(s.name || "(이름 미입력)")}</span>
                    <span class="sd-emp">${escapeHtml(s.empNo)}</span>
                    <span class="sd-branch">${escapeHtml(s.branch || "")}</span>
                    <span class="sd-cohort">${escapeHtml(s.cohort || "")}</span>
                  </label>
                `).join("")}
              </div>
            </div>
          `).join("")}
        </div>
      `;
    }).join("");

    // 개별 체크
    container.querySelectorAll("input[data-emp]").forEach((inp) => {
      inp.addEventListener("change", (e) => {
        const emp = e.target.dataset.emp;
        if (e.target.checked) state.sdSelected.add(emp);
        else state.sdSelected.delete(emp);
        renderStudentDeleteList($("#sd-search").value);
        updateSdCounts();
      });
    });
    // 지역단 전체 토글
    container.querySelectorAll("input[data-region]").forEach((inp) => {
      inp.addEventListener("change", (e) => {
        const region = e.target.dataset.region;
        const centers = byRegion[region];
        const emps = Object.values(centers).flat().map((s) => s.empNo);
        if (e.target.checked) emps.forEach((emp) => state.sdSelected.add(emp));
        else emps.forEach((emp) => state.sdSelected.delete(emp));
        renderStudentDeleteList($("#sd-search").value);
        updateSdCounts();
      });
    });
  }

  async function doStudentsDeleteFromModal() {
    const selected = state.students.filter((s) => state.sdSelected.has(s.empNo));
    if (!selected.length) { toast("선택된 교육생이 없습니다.", "error"); return; }
    const ok = await confirmAndDeleteStudents(selected, { label: "선택한" });
    if (ok) {
      state.sdSelected = new Set();
      closeModal("#modal-student-delete");
    }
  }

  // ========== 모바일 사이드바 제어 ==========
  function isMobileViewport() { return window.matchMedia("(max-width: 640px)").matches; }
  function openMobileSidebar() {
    state.mobileSidebarOpen = true;
    document.documentElement.classList.add("mobile-sidebar-open");
    document.body.classList.add("mobile-sidebar-open");
    const bd = document.getElementById("mobile-sidebar-backdrop");
    if (bd) bd.hidden = false;
    const hb = document.getElementById("mbn-home-btn");
    if (hb) hb.classList.add("mbn-sidebar-open");
  }
  function closeMobileSidebar() {
    state.mobileSidebarOpen = false;
    document.documentElement.classList.remove("mobile-sidebar-open");
    document.body.classList.remove("mobile-sidebar-open");
    const bd = document.getElementById("mobile-sidebar-backdrop");
    if (bd) bd.hidden = true;
    const hb = document.getElementById("mbn-home-btn");
    if (hb) hb.classList.remove("mbn-sidebar-open");
  }
  function toggleMobileSidebar() {
    if (state.mobileSidebarOpen) closeMobileSidebar();
    else openMobileSidebar();
  }

  // 좌우 스와이프로 사이드바 열고 닫기
  function bindMobileSwipe() {
    let startX = 0, startY = 0, dragging = null;
    const THRESHOLD = 60;   // 성공 거리
    const EDGE = 24;        // 왼쪽 가장자리 스와이프 열기 영역
    document.addEventListener("touchstart", (e) => {
      if (!isMobileViewport()) { dragging = null; return; }
      const t = e.touches[0];
      startX = t.clientX; startY = t.clientY;
      if (!state.mobileSidebarOpen && startX < EDGE) dragging = "open";
      else if (state.mobileSidebarOpen && e.target.closest(".sidebar, .mobile-sidebar-backdrop")) dragging = "close";
      else dragging = null;
    }, { passive: true });
    document.addEventListener("touchmove", (e) => {
      if (!dragging) return;
      const t = e.touches[0];
      const dx = t.clientX - startX;
      const dy = t.clientY - startY;
      if (Math.abs(dy) > Math.abs(dx)) { dragging = null; return; }
    }, { passive: true });
    document.addEventListener("touchend", (e) => {
      if (!dragging) return;
      const t = e.changedTouches[0];
      const dx = t.clientX - startX;
      if (dragging === "open" && dx > THRESHOLD) openMobileSidebar();
      else if (dragging === "close" && dx < -THRESHOLD) closeMobileSidebar();
      dragging = null;
    });
  }

  function switchView(href) {
    const map = {
      "#dashboard": "dashboard",
      "#students":  "students",
      "#progress":  "progress",
      "#stats":     "stats",
      "#settings":  "settings"
    };
    const target = map[href];
    // 관리자(설정) 진입 시 암호 체크 — 매번 묻도록
    if (target === "settings") {
      const pwd = prompt("관리자 암호를 입력하세요:");
      if (pwd !== "2051") {
        toast("암호가 일치하지 않습니다.", "error");
        return;
      }
    }
    // nav active 토글 (데스크톱 상단 + 모바일 하단 동시)
    $$(".top-nav a, .mobile-bottom-nav .mbn-btn").forEach((a) => {
      a.classList.toggle("active", a.getAttribute("href") === href);
    });
    // 모든 view-panel 숨기고 target 만 노출
    $$(".view-panel").forEach((p) => {
      p.hidden = (p.dataset.view !== target);
    });
    // 노출된 패널로 스크롤 + 해당 패널 최신 렌더 보장
    const el = document.querySelector(`.view-panel[data-view="${target}"]`);
    if (el) el.scrollIntoView({ behavior: "smooth", block: "start" });
    if (target === "dashboard") renderStats(filteredStudents(), "dashboard-body");
    if (target === "progress") renderProgressPanel();
    if (target === "settings") renderMasterTargetSettings();
      if (target === "settings") subscribeErrorReportsIfNeeded();
    if (target === "students" && state.selectedEmpNo) renderStudentDetail();
    if (target === "students" && !state.selectedEmpNo) {
      toast("좌측 [지점별 교육생] 목록에서 교육생을 선택하세요.", "");
    }
  }

  function retrySubscription() {
    toast("Firestore 재구독 시도중...", "");
    state.studentsLoaded = false;
    state._consultCountsFetched = false;
    if (state._subscribeUnsub) {
      try { state._subscribeUnsub(); } catch (e) {}
    }
    state._subscribeUnsub = window.DataAPI.subscribe((list, meta) => {
      state.students = list || [];
      state.studentsLoaded = true;
      state.syncMeta = meta || { fromCache: false };
      updateDataTimestamp();
      renderDebounced();
      if (list && list.length > 0) toast(`${list.length}건 동기화 완료`, "success");
      prefetchConsultCountsOnce();
    });
    render();
  }

  async function clearCacheAndReload() {
    try {
      // IndexedDB 의 firestore 캐시 삭제
      if (window.indexedDB) {
        const dbs = await (indexedDB.databases ? indexedDB.databases() : Promise.resolve([]));
        for (const d of dbs) {
          if (d.name && d.name.includes("firestore")) {
            try { indexedDB.deleteDatabase(d.name); } catch (e) {}
          }
        }
      }
      localStorage.removeItem("cmf.filter.v1");
      // 강제 캐시 무시 새로고침
      window.location.reload();
    } catch (e) {
      console.error(e);
      window.location.reload();
    }
  }

  // 공유 링크 파라미터 감지 — #share?r=지역단&c=기수&s=스텝
  function _checkShareMode() {
    const hash = location.hash;
    if (!hash.startsWith("#share?")) return;
    const params = new URLSearchParams(hash.slice(7));
    const r = decodeURIComponent(params.get("r") || "");
    if (!r) return;
    const c = params.get("c") || "";
    const s = params.get("s") || "1";
    state.pgShareMode    = true;
    state.progressRegion = r;
    state.progressCohort = c;
    state.progressStep   = s;
    state.filter.cohort  = c ? `${c}기` : "";
    state.filter.step    = s;
    document.body.classList.add("pg-share-mode");
    switchView("#progress");
  }

  function init() {
    bindEvents();
    bindRegionMgrEvents();
    initDraggableModals();
    initErrorReportModal();
    // 기본 view = 교육생 관리 (공유 링크인 경우 _checkShareMode가 오버라이드)
    switchView("#students");
    _checkShareMode();
    // localStorage에서 복원된 필터값을 UI에 반영
    $("#filter-cohort").value = state.filter.cohort || "";
    const fsEl = document.getElementById("filter-step"); if (fsEl) fsEl.value = state.filter.step || "1";
    syncOrgLabels();
    // 첫 응답 5초 안에 안 오면 사용자에게 안내
    const slowTimer = setTimeout(() => {
      if (!state.studentsLoaded) {
        toast("Firebase 응답이 느립니다. 네트워크/방화벽을 확인하세요.", "error");
      }
    }, 5000);
    // 구독 에러 hook
    window.__onSubscribeError = (err) => {
      toast("Firebase 연결 실패: " + (err.message || err.code), "error");
    };
    state._subscribeUnsub = window.DataAPI.subscribe((list, meta) => {
      clearTimeout(slowTimer);
      state.students = list || [];
      state.studentsLoaded = true;
      state.syncMeta = meta || { fromCache: false };
      migrateStudentBaseValuesIfNeeded();
      updateDataTimestamp();
      renderDebounced();
      prefetchConsultCountsOnce();
    });
    // 시상안 Firestore → localStorage 동기화 (1회)
    syncAwardPlansFromFirestore();
  }

  // 구 단위 마이그레이션 — 원 단위 저장으로 전환됨에 따라 비활성화
  async function migrateStudentBaseValuesIfNeeded() { /* no-op: v0.92+ stores all values in 원 */ }

  function updateDataTimestamp() {
    const el = document.getElementById("data-last-updated");
    if (!el) return;
    const now = new Date();
    const pad = (n) => String(n).padStart(2, "0");
    el.textContent = `${now.getFullYear()}.${pad(now.getMonth() + 1)}.${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}`;
  }

  // 최초 1회만 전체 면담 횟수 사전 수집 → state.students 에 consultCount merge → 사이드바 재렌더
  async function prefetchConsultCountsOnce() {
    if (state._consultCountsFetched) return;
    if (!state.students.length) return;
    if (!window.DataAPI || typeof window.DataAPI.fetchAllConsultCounts !== "function") return;
    state._consultCountsFetched = true;
    try {
      const empNos = state.students.map((s) => s.empNo).filter(Boolean);
      const counts = await window.DataAPI.fetchAllConsultCounts(empNos);
      let changed = false;
      state.students = state.students.map((s) => {
        const c = counts[s.empNo] || 0;
        if (Number(s.consultCount || 0) !== c) changed = true;
        return { ...s, consultCount: c };
      });
      if (changed) renderDebounced();
      // 서버측 consultCount 와 실제 값이 다르면 비동기 자가치유 (best-effort)
      if (typeof window.DataAPI.syncConsultCount === "function") {
        state.students.forEach((s) => {
          window.DataAPI.syncConsultCount(s.empNo, s.consultCount || 0).catch(() => {});
        });
      }
    } catch (e) {
      console.warn("[app] 면담 횟수 사전 수집 실패:", e);
      state._consultCountsFetched = false; // 재시도 가능
    }
  }

  // ========== 오류신고 ==========

  function compressImageToBase64(file, maxW, quality) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const img = new Image();
        img.onload = () => {
          const scale = img.width > maxW ? maxW / img.width : 1;
          const w = Math.round(img.width * scale);
          const h = Math.round(img.height * scale);
          const canvas = document.createElement("canvas");
          canvas.width = w; canvas.height = h;
          canvas.getContext("2d").drawImage(img, 0, 0, w, h);
          resolve(canvas.toDataURL("image/jpeg", quality));
        };
        img.onerror = reject;
        img.src = e.target.result;
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }

  function openErrorReportModal() {
    const modal = document.getElementById("modal-error-report");
    if (!modal) return;
    // Reset form
    ["er-title","er-content","er-reporter-name","er-reporter-contact"].forEach((id) => {
      const el = document.getElementById(id); if (el) el.value = "";
    });
    const fileEl = document.getElementById("er-image");
    if (fileEl) fileEl.value = "";
    const preview = document.getElementById("er-image-preview");
    if (preview) preview.hidden = true;
    const previewImg = document.getElementById("er-preview-img");
    if (previewImg) previewImg.src = "";
    modal._imageBase64 = null;
    modal.hidden = false;
  }

  function initErrorReportModal() {
    const modal = document.getElementById("modal-error-report");
    if (!modal) return;
    modal.querySelectorAll("[data-close]").forEach((el) => el.addEventListener("click", () => { modal.hidden = true; }));

    const fileEl = document.getElementById("er-image");
    if (fileEl) {
      fileEl.addEventListener("change", async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        try {
          const base64 = await compressImageToBase64(file, 1200, 0.75);
          modal._imageBase64 = base64;
          const previewImg = document.getElementById("er-preview-img");
          const preview = document.getElementById("er-image-preview");
          if (previewImg) previewImg.src = base64;
          if (preview) preview.hidden = false;
        } catch (e) { toast("이미지 처리 실패: " + e.message, "error"); }
      });
    }

    // 연락처 자동 하이픈 포맷 (010-0000-0000)
    const contactEl = document.getElementById("er-reporter-contact");
    if (contactEl) {
      contactEl.addEventListener("input", () => {
        let v = contactEl.value.replace(/\D/g, "").slice(0, 11);
        if (v.length > 7) v = v.slice(0, 3) + "-" + v.slice(3, 7) + "-" + v.slice(7);
        else if (v.length > 3) v = v.slice(0, 3) + "-" + v.slice(3);
        contactEl.value = v;
      });
    }

    const removeBtn = document.getElementById("er-remove-img");
    if (removeBtn) {
      removeBtn.addEventListener("click", () => {
        modal._imageBase64 = null;
        const fileEl2 = document.getElementById("er-image"); if (fileEl2) fileEl2.value = "";
        const previewImg = document.getElementById("er-preview-img"); if (previewImg) previewImg.src = "";
        const preview = document.getElementById("er-image-preview"); if (preview) preview.hidden = true;
      });
    }

    const submitBtn = document.getElementById("btn-er-submit");
    if (submitBtn) {
      submitBtn.addEventListener("click", async () => {
        const title = (document.getElementById("er-title")?.value || "").trim();
        const content = (document.getElementById("er-content")?.value || "").trim();
        const reporterName = (document.getElementById("er-reporter-name")?.value || "").trim();
        const reporterContact = (document.getElementById("er-reporter-contact")?.value || "").trim();
        if (!title) { toast("오류 제목을 입력하세요.", "error"); return; }
        if (!content) { toast("오류 내용을 입력하세요.", "error"); return; }
        if (!reporterName) { toast("신고자 이름을 입력하세요.", "error"); return; }
        if (!/^\d{3}-\d{3,4}-\d{4}$/.test(reporterContact)) { toast("연락처를 010-0000-0000 형식으로 입력하세요.", "error"); return; }
        submitBtn.disabled = true;
        try {
          await window.DataAPI.addErrorReport({ title, content, reporterName, reporterContact, imageBase64: modal._imageBase64 || null });
          toast("오류신고가 저장되었습니다. 감사합니다!", "success");
          modal.hidden = true;
        } catch (e) { toast("저장 실패: " + e.message, "error"); }
        finally { submitBtn.disabled = false; }
      });
    }
  }

  function subscribeErrorReportsIfNeeded() {
    if (state.errorReportUnsub) return; // 이미 구독 중
    if (!window.DataAPI?.subscribeErrorReports) return;
    state.errorReportUnsub = window.DataAPI.subscribeErrorReports((list) => {
      state.errorReports = list;
      renderErrorReportsBoard();
    });
  }

  function renderErrorReportsBoard() {
    const board = document.getElementById("error-reports-board");
    if (!board) return;
    const delBtn = document.getElementById("btn-er-delete-sel");
    const list = state.errorReports || [];
    if (!list.length) {
      board.innerHTML = `<p class="settings-desc" style="color:#999;">접수된 오류신고가 없습니다.</p>`;
      if (delBtn) delBtn.hidden = true;
      return;
    }
    if (delBtn) delBtn.hidden = false;

    board.innerHTML = list.map((r) => {
      const dateStr = r.createdAt?.toDate ? r.createdAt.toDate().toLocaleDateString("ko-KR") : "—";
      const resolvedCls = r.resolved ? "er-item-resolved" : "";
      const badge = r.resolved
        ? `<span class="er-confirmed-badge">확인</span>`
        : `<span class="er-item-status" title="확인 처리">⬜</span>`;
      return `
        <div class="er-item ${resolvedCls}" data-erid="${escapeHtml(r.id)}">
          <div class="er-item-head">
            <input type="checkbox" class="er-item-chk" data-erid="${escapeHtml(r.id)}">
            ${badge}
            <span class="er-item-date">${dateStr}</span>
            <span class="er-item-reporter">${escapeHtml(r.reporterName || "—")}</span>
            <span class="er-item-title">${escapeHtml(r.title || "(제목 없음)")}</span>
            <button class="er-item-print btn-outline small" data-erid="${escapeHtml(r.id)}">🖨 인쇄</button>
          </div>
          <div class="er-item-body" hidden>
            ${r.imageBase64 ? `<img src="${r.imageBase64}" class="er-item-img" alt="첨부 이미지">` : ""}
            <div class="er-item-content">${escapeHtml(r.content || "")}</div>
            <div class="er-item-meta">신고자: ${escapeHtml(r.reporterName || "—")} / 연락처: ${escapeHtml(r.reporterContact || "—")}</div>
          </div>
        </div>
      `;
    }).join("");

    // Toggle expand (제목/날짜 클릭)
    board.querySelectorAll(".er-item-title, .er-item-date").forEach((el) => {
      el.style.cursor = "pointer";
      el.addEventListener("click", () => {
        const body = el.closest(".er-item").querySelector(".er-item-body");
        if (body) body.hidden = !body.hidden;
      });
    });

    // 확인 처리 토글 (⬜ 아이콘 클릭)
    board.querySelectorAll(".er-item-status").forEach((el) => {
      el.style.cursor = "pointer";
      el.addEventListener("click", async () => {
        const id = el.closest(".er-item").dataset.erid;
        const rep = state.errorReports.find((r) => r.id === id);
        if (!rep) return;
        try {
          await window.DataAPI.updateErrorReport(id, { resolved: !rep.resolved });
        } catch (e) { toast("업데이트 실패: " + e.message, "error"); }
      });
    });

    // 인쇄
    board.querySelectorAll(".er-item-print").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        const id = btn.dataset.erid;
        const rep = state.errorReports.find((r) => r.id === id);
        if (rep) printErrorReport(rep);
      });
    });

    // 선택 삭제 버튼
    if (delBtn) {
      delBtn.onclick = async () => {
        const checked = [...board.querySelectorAll(".er-item-chk:checked")].map((el) => el.dataset.erid);
        if (!checked.length) { toast("삭제할 항목을 선택하세요.", "error"); return; }
        if (!confirm(`선택한 ${checked.length}건을 삭제합니다. 되돌릴 수 없습니다.`)) return;
        delBtn.disabled = true;
        try {
          await Promise.all(checked.map((id) => window.DataAPI.deleteErrorReport(id)));
          toast(`${checked.length}건 삭제 완료`, "success");
        } catch (e) { toast("삭제 실패: " + e.message, "error"); }
        finally { delBtn.disabled = false; }
      };
    }
  }

  function printErrorReport(rep) {
    const dateStr = rep.createdAt?.toDate ? rep.createdAt.toDate().toLocaleString("ko-KR") : "—";
    const resolvedStr = rep.resolved ? "✅ 해결됨" : "⬜ 미해결";
    const imgHtml = rep.imageBase64
      ? `<div class="er-print-img-wrap"><img src="${rep.imageBase64}" class="er-print-img" alt="첨부 이미지"></div>`
      : `<div class="er-print-img-wrap er-print-no-img">첨부 이미지 없음</div>`;
    const html = `<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"><title>오류신고</title>
    <style>
      *{box-sizing:border-box;margin:0;padding:0;}
      body{font-family:'Noto Sans KR','Malgun Gothic',sans-serif;font-size:12px;color:#1a1a1a;}
      .er-print-wrap{padding:10mm 12mm;width:210mm;min-height:297mm;}
      .er-print-hdr{background:#1A2744;color:#fff;padding:8px 14px;border-radius:6px;margin-bottom:8px;display:flex;justify-content:space-between;align-items:center;}
      .er-print-hdr-title{font-size:16px;font-weight:900;}
      .er-print-hdr-meta{font-size:11px;color:rgba(255,255,255,.75);}
      .er-print-img-wrap{height:50%;display:flex;align-items:center;justify-content:center;background:#f5f5f5;border:1px solid #ddd;border-radius:6px;margin-bottom:8px;overflow:hidden;}
      .er-print-img{max-width:100%;max-height:100%;object-fit:contain;}
      .er-print-no-img{height:120px;color:#999;font-size:14px;}
      .er-print-section{margin-bottom:8px;padding:8px 12px;background:#F8F9FF;border-radius:6px;border:1px solid #E0E0E0;}
      .er-print-label{font-size:10px;font-weight:700;color:#5C6BC0;margin-bottom:3px;}
      .er-print-value{font-size:13px;white-space:pre-wrap;}
      .er-print-row{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:8px;}
      .er-print-status{display:inline-block;padding:3px 10px;border-radius:20px;font-weight:700;font-size:12px;}
      .er-print-status.resolved{background:#E8F5E9;color:#2E7D32;border:1px solid #A5D6A7;}
      .er-print-status.unresolved{background:#FFF3E0;color:#E65100;border:1px solid #FFCC80;}
      @media print{@page{size:A4 portrait;margin:0;}body{margin:0;}.er-print-wrap{padding:8mm 10mm;}}
    </style></head><body>
    <div class="er-print-wrap">
      <div class="er-print-hdr">
        <span class="er-print-hdr-title">🚨 오류신고</span>
        <span class="er-print-hdr-meta">${escapeHtml(dateStr)}</span>
      </div>
      ${imgHtml}
      <div class="er-print-row">
        <div class="er-print-section">
          <div class="er-print-label">신고자</div>
          <div class="er-print-value">${escapeHtml(rep.reporterName || "—")}</div>
        </div>
        <div class="er-print-section">
          <div class="er-print-label">연락처</div>
          <div class="er-print-value">${escapeHtml(rep.reporterContact || "—")}</div>
        </div>
      </div>
      <div class="er-print-section">
        <div class="er-print-label">제목</div>
        <div class="er-print-value">${escapeHtml(rep.title || "(제목 없음)")}</div>
      </div>
      <div class="er-print-section">
        <div class="er-print-label">오류 내용</div>
        <div class="er-print-value">${escapeHtml(rep.content || "")}</div>
      </div>
      <div class="er-print-section">
        <div class="er-print-label">처리 상태</div>
        <span class="er-print-status ${rep.resolved ? "resolved" : "unresolved"}">${resolvedStr}</span>
      </div>
    </div>
    </body></html>`;
    const w = window.open("", "_blank", "width=900,height=1200");
    if (!w) { toast("팝업이 차단되었습니다.", "warn"); return; }
    w.document.write(html);
    w.document.close();
    w.addEventListener("load", () => setTimeout(() => { w.focus(); w.print(); }, 300));
  }

  // ========== 모달 드래그 이동 (전역, 이벤트 위임) ==========

  function initDraggableModals() {
    let dragging = false, dragPanel = null, dragHead = null, ox = 0, oy = 0;

    document.addEventListener("mousedown", (e) => {
      const head = e.target.closest(".modal-head");
      if (!head || e.target.closest("button,a,select,input,textarea")) return;
      const panel = head.closest(".modal-panel");
      if (!panel) return;
      e.preventDefault();
      const r = panel.getBoundingClientRect();
      panel.style.position = "fixed";
      panel.style.margin   = "0";
      panel.style.left     = r.left + "px";
      panel.style.top      = r.top  + "px";
      ox = e.clientX - r.left;
      oy = e.clientY - r.top;
      dragging = true;
      dragPanel = panel;
      dragHead  = head;
      head.style.cursor = "grabbing";
    });

    document.addEventListener("mousemove", (e) => {
      if (!dragging || !dragPanel) return;
      const x = Math.max(0, Math.min(e.clientX - ox, window.innerWidth  - dragPanel.offsetWidth));
      const y = Math.max(0, Math.min(e.clientY - oy, window.innerHeight - dragPanel.offsetHeight));
      dragPanel.style.left = x + "px";
      dragPanel.style.top  = y + "px";
    });

    document.addEventListener("mouseup", () => {
      if (!dragging) return;
      dragging = false;
      if (dragHead) { dragHead.style.cursor = ""; dragHead = null; }
      dragPanel = null;
    });

    // 정적 모달이 재오픈될 때 위치 초기화 (MutationObserver)
    new MutationObserver((muts) => {
      muts.forEach((m) => {
        if (m.attributeName !== "hidden") return;
        const modal = m.target;
        if (!modal.hidden || !modal.classList?.contains("modal")) return;
        const panel = modal.querySelector(".modal-panel");
        if (panel) { panel.style.position = ""; panel.style.margin = ""; panel.style.left = ""; panel.style.top = ""; }
      });
    }).observe(document.body, { attributes: true, attributeFilter: ["hidden"], subtree: true });
  }

  // ========== 지역단 관리 모달 ==========

  function openRegionMgrModal() {
    const region = state.filter.region || "";
    if (!region) { toast("좌측 필터에서 지역단을 먼저 선택해주세요.", "warn"); return; }
    const modal = document.getElementById("modal-region-mgr");
    if (!modal) return;
    document.getElementById("rm-region-label").textContent = region;
    // sync cohort/step from current filter
    const cohortSel = document.getElementById("rm-cohort-sel");
    if (cohortSel && state.filter.cohort) cohortSel.value = state.filter.cohort;
    const stepVal = state.filter.step || "1";
    const stepRadio = modal.querySelector(`input[name="rm-step"][value="${stepVal}"]`);
    if (stepRadio) stepRadio.checked = true;
    modal.hidden = false;
    _rmRenderTable();
  }

  function _rmClose() {
    const modal = document.getElementById("modal-region-mgr");
    if (modal) modal.hidden = true;
  }

  function _rmGetState() {
    const region  = (document.getElementById("rm-region-label")?.textContent || "").trim();
    const cohort  = document.getElementById("rm-cohort-sel")?.value || "";
    const stepVal = (document.querySelector('input[name="rm-step"]:checked') || {}).value || "1";
    const sfx     = stepVal === "1" ? "" : stepVal;
    return { region, cohort, stepVal, sfx };
  }

  function _rmRenderTable() {
    const { region, cohort, sfx } = _rmGetState();
    const wrap = document.getElementById("rm-table-wrap");
    if (!wrap) return;
    const rows = state.students
      .filter((s) => s.region === region && (!cohort || !s.cohort || s.cohort === cohort))
      .sort((a, b) => (a.branch || "").localeCompare(b.branch || "", "ko") || (a.name || "").localeCompare(b.name || "", "ko"));
    if (!rows.length) {
      wrap.innerHTML = `<div class="rm-table-empty">해당 기수/스텝의 교육생이 없습니다.</div>`;
      return;
    }
    wrap.innerHTML = `<table class="pg-tbl pg-admin-tbl rm-tbl">
      <thead><tr>
        <th>#</th><th>지점</th><th>성명</th>
        <th>기준실적(원)</th><th>현재실적(원)</th><th>마스터목표(원)</th>
        <th>달성률</th><th>순증</th>
        <th>인품건</th><th>인품실적(원)</th>
        <th>전화번호</th>
      </tr></thead>
      <tbody>${rows.map((s, i) => {
        const base  = Number(s.base || 0);
        const cur   = sfx ? Number(s[`pgCurrent${sfx}`] || 0) : (s.pgCurrent !== undefined ? Number(s.pgCurrent) : Number(s.current || 0));
        const goal  = Number(s.target) > 0 ? Number(s.target) : base;
        const iCnt  = Number(s[`pgIpumCount${sfx}`] || s.pgIpumCount || 0);
        const iAmt  = Number(s[`pgIpumAmt${sfx}`]   || s.pgIpumAmt  || 0);
        const net   = cur - goal;
        const rate  = goal > 0 ? (cur / goal * 100).toFixed(1) : "0.0";
        const fCur  = sfx ? `pgCurrent${sfx}` : "pgCurrent";
        const fIcnt = sfx ? `pgIpumCount${sfx}` : "pgIpumCount";
        const fIamt = sfx ? `pgIpumAmt${sfx}`   : "pgIpumAmt";
        return `<tr>
          <td>${i + 1}</td>
          <td>${escapeHtml(s.branch || "")}</td>
          <td><strong>${escapeHtml(s.name || "")}</strong></td>
          <td class="r pg-base-ref">${base > 0 ? Nf(base) : "—"}</td>
          <td><input type="number" class="pg-input rm-inp" data-emp="${escapeHtml(s.empNo)}" data-f="${escapeHtml(fCur)}" data-pg-role="current" value="${cur}" min="0" step="1"></td>
          <td><input type="number" class="pg-input rm-inp rm-base-inp" data-emp="${escapeHtml(s.empNo)}" data-f="target" data-pg-role="goal" data-pgbase="${base}" value="${goal}" min="0" step="1"></td>
          <td class="r" data-rm-calc="rate-${escapeHtml(s.empNo)}">${rate}%</td>
          <td class="r" data-rm-calc="net-${escapeHtml(s.empNo)}">${net >= 0 ? "+" : ""}${Nf(net)}</td>
          <td><input type="number" class="pg-input rm-inp" data-emp="${escapeHtml(s.empNo)}" data-f="${escapeHtml(fIcnt)}" value="${iCnt}" min="0"></td>
          <td><input type="number" class="pg-input rm-inp" data-emp="${escapeHtml(s.empNo)}" data-f="${escapeHtml(fIamt)}" value="${iAmt}" min="0"></td>
          <td><input type="text" class="pg-input rm-inp rm-phone" data-emp="${escapeHtml(s.empNo)}" data-f="phone" value="${escapeHtml(s.phone || "")}" placeholder="연락처" style="width:110px"></td>
        </tr>`;
      }).join("")}</tbody>
    </table>`;

    // 행 재계산: 현재실적·마스터목표 입력값 기반으로 달성률·순증 업데이트
    function _rmRecalcRow(emp) {
      const curInp  = wrap.querySelector(`.rm-inp[data-pg-role="current"][data-emp="${emp}"]`);
      const goalInp = wrap.querySelector(`.rm-inp[data-pg-role="goal"][data-emp="${emp}"]`);
      const cur  = Number(curInp?.value)  || 0;
      const goal = Number(goalInp?.value) || 0;
      const net  = cur - goal;
      const rate = goal > 0 ? (cur / goal * 100).toFixed(1) : "0.0";
      const rateEl = wrap.querySelector(`[data-rm-calc="rate-${emp}"]`);
      const netEl  = wrap.querySelector(`[data-rm-calc="net-${emp}"]`);
      if (rateEl) rateEl.textContent = `${rate}%`;
      if (netEl)  netEl.textContent  = `${net >= 0 ? "+" : ""}${Nf(net)}`;
    }

    wrap.querySelectorAll('.rm-inp[data-pg-role="current"], .rm-inp[data-pg-role="goal"]').forEach((inp) => {
      inp.addEventListener("input", () => _rmRecalcRow(inp.dataset.emp));
    });

    // 목표일괄수정: 기준실적(pgBase) + 선택금액 → 마스터목표(target)
    const bulkBtn = document.getElementById("btn-rm-bulk-base");
    const bulkPop = document.getElementById("rm-bulk-popup");
    if (bulkBtn && bulkPop) {
      bulkBtn.onclick = (e) => { e.stopPropagation(); bulkPop.hidden = !bulkPop.hidden; };
      bulkPop.querySelectorAll(".btn-bulk-amt").forEach((b) => {
        b.onclick = () => {
          const amt = Number(b.dataset.amt) || 0;
          wrap.querySelectorAll('.rm-inp[data-pg-role="goal"]').forEach((inp) => {
            const pgBase = Number(inp.dataset.pgbase) || 0;
            inp.value = pgBase + amt;
            _rmRecalcRow(inp.dataset.emp);
          });
          bulkPop.hidden = true;
          const msg = document.getElementById("rm-table-save-msg");
          if (msg) { msg.textContent = `마스터목표 기준실적+${Nf(amt)}원 적용 — 저장 버튼으로 확정`; msg.className = "pg-msg ok"; setTimeout(() => { msg.textContent = ""; }, 4000); }
        };
      });
      document.addEventListener("click", function _closeBulk(e) {
        if (!bulkPop.hidden && !bulkPop.contains(e.target) && e.target !== bulkBtn) {
          bulkPop.hidden = true;
          document.removeEventListener("click", _closeBulk);
        }
      });
    }
  }

  async function _rmSaveTable() {
    const btn = document.getElementById("btn-rm-table-save");
    const msg = document.getElementById("rm-table-save-msg");
    const wrap = document.getElementById("rm-table-wrap");
    if (!wrap) return;
    const inputs = wrap.querySelectorAll(".rm-inp");
    if (!inputs.length) return;
    const changes = {};
    inputs.forEach((inp) => {
      const emp = inp.dataset.emp;
      const f   = inp.dataset.f;
      if (!emp || !f) return;
      if (!changes[emp]) changes[emp] = {};
      changes[emp][f] = f === "phone" ? inp.value.trim() : (Number(inp.value) || 0);
    });
    const empNos = Object.keys(changes);
    if (!empNos.length) { msg.textContent = "변경 내용이 없습니다."; msg.className = "pg-msg"; return; }
    if (btn) { btn.disabled = true; btn.textContent = "저장중..."; }
    try {
      const updates = empNos.map((emp) => {
        const st = state.students.find((x) => x.empNo === emp);
        if (!st) return null;
        return { ...st, ...changes[emp] };
      }).filter(Boolean);
      await window.DataAPI.saveMany(updates);
      if (msg) { msg.textContent = `✅ ${updates.length}명 저장 완료`; msg.className = "pg-msg ok"; }
      setTimeout(() => { if (msg) msg.textContent = ""; }, 3000);
      renderDebounced();
    } catch (e) {
      if (msg) { msg.textContent = "❌ 저장 실패: " + e.message; msg.className = "pg-msg err"; }
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = "💾 저장 (Firestore)"; }
    }
  }

  // 연락처 업데이트 모달
  (function _initContactUpdate() {
    let _cuParsed = []; // { name, empNo, phone, matched: student|null }

    function _cuOpen() {
      const modal = document.getElementById("modal-contact-update");
      if (!modal) return;
      // 미리보기 초기화
      document.getElementById("cu-preview-wrap").hidden = true;
      document.getElementById("cu-paste-area").value = "";
      document.getElementById("cu-parse-msg").textContent = "";
      document.getElementById("cu-save-msg").textContent = "";
      _cuParsed = [];
      modal.hidden = false;
      setTimeout(() => document.getElementById("cu-paste-area")?.focus(), 80);
    }

    function _cuClose() {
      const modal = document.getElementById("modal-contact-update");
      if (modal) modal.hidden = true;
    }

    function _cuParse() {
      const txt = (document.getElementById("cu-paste-area")?.value || "").trim();
      const msgEl = document.getElementById("cu-parse-msg");
      const previewWrap = document.getElementById("cu-preview-wrap");
      const tbody = document.getElementById("cu-preview-body");
      const setMsg = (t, cls = "") => { msgEl.textContent = t; msgEl.className = "pg-msg" + (cls ? " " + cls : ""); };

      if (!txt) { setMsg("❌ 붙여넣기 내용이 없습니다.", " err"); return; }

      const lines = txt.split(/\r?\n/).filter((l) => l.trim());
      if (!lines.length) { setMsg("❌ 내용이 없습니다.", " err"); return; }

      // 헤더 자동 감지
      const HEADER_KEYS = { "성명": "name", "이름": "name", "사번": "empNo", "사원번호": "empNo", "전화번호": "phone", "연락처": "phone", "휴대폰": "phone", "핸드폰": "phone" };
      const firstCols = lines[0].trim().split(/\t/).map((c) => c.trim());
      const isHeader = firstCols.some((h) => HEADER_KEYS[h]);
      let dataLines, cols;
      if (isHeader) {
        cols = firstCols.map((h) => HEADER_KEYS[h] || "ignore");
        dataLines = lines.slice(1);
      } else {
        // 기본: 성명 사번 전화번호
        cols = ["name", "empNo", "phone"];
        dataLines = lines;
      }

      if (!cols.includes("phone")) { setMsg("❌ 전화번호 열을 찾을 수 없습니다.", " err"); return; }

      _cuParsed = dataLines.map((line) => {
        const parts = line.split(/\t/).map((c) => c.trim());
        const rec = { name: "", empNo: "", phone: "", matched: null };
        cols.forEach((f, i) => { if (f !== "ignore" && parts[i] !== undefined) rec[f] = parts[i]; });
        // 매칭: 사번 우선, 없으면 성명
        if (rec.empNo) rec.matched = state.students.find((s) => s.empNo === rec.empNo) || null;
        if (!rec.matched && rec.name) rec.matched = state.students.find((s) => s.name === rec.name) || null;
        return rec;
      }).filter((r) => r.phone);

      if (!_cuParsed.length) { setMsg("❌ 유효한 데이터가 없습니다.", " err"); previewWrap.hidden = true; return; }

      const matched = _cuParsed.filter((r) => r.matched).length;
      setMsg(`총 ${_cuParsed.length}건 — 매칭 ${matched}건 / 미매칭 ${_cuParsed.length - matched}건`, matched > 0 ? " ok" : " err");

      tbody.innerHTML = _cuParsed.map((r, i) => {
        const st = r.matched;
        const status = st ? `<span style="color:#1976d2;font-weight:600">✔ 매칭</span>` : `<span style="color:#e53935">✘ 미매칭</span>`;
        const curPhone = st ? escapeHtml(st.phone || "—") : `<span style="color:#999">—</span>`;
        const bgStyle = !st ? 'style="background:#fff8f8"' : (st.phone === r.phone ? 'style="background:#f5f5f5;color:#aaa"' : '');
        return `<tr ${bgStyle}>
          <td style="text-align:center">${i + 1}</td>
          <td><strong>${escapeHtml(r.name || (st?.name || ""))}</strong></td>
          <td style="font-family:monospace">${escapeHtml(r.empNo || (st?.empNo || ""))}</td>
          <td>${curPhone}</td>
          <td><strong>${escapeHtml(r.phone)}</strong></td>
          <td style="text-align:center">${status}</td>
        </tr>`;
      }).join("");

      previewWrap.hidden = false;
      document.getElementById("cu-save-msg").textContent = "";
    }

    async function _cuSave() {
      const toSave = _cuParsed.filter((r) => r.matched);
      const msgEl = document.getElementById("cu-save-msg");
      const saveBtn = document.getElementById("cu-save-btn");
      if (!toSave.length) { msgEl.textContent = "❌ 저장할 매칭된 데이터가 없습니다."; msgEl.className = "pg-msg err"; return; }
      saveBtn.disabled = true; saveBtn.textContent = "저장중...";
      try {
        const updates = toSave.map((r) => ({ ...r.matched, phone: r.phone }));
        await window.DataAPI.saveMany(updates);
        msgEl.textContent = `✅ ${updates.length}명 연락처 저장 완료`;
        msgEl.className = "pg-msg ok";
        // 편집 테이블 전화번호 인풋도 즉시 반영
        updates.forEach((u) => {
          const inp = document.querySelector(`.rm-inp[data-emp="${u.empNo}"][data-f="phone"]`);
          if (inp) inp.value = u.phone;
        });
        setTimeout(() => _cuClose(), 1800);
      } catch (e) {
        msgEl.textContent = "❌ 저장 실패: " + e.message;
        msgEl.className = "pg-msg err";
      } finally {
        saveBtn.disabled = false; saveBtn.textContent = "💾 저장";
      }
    }

    document.getElementById("btn-rm-contact-update")?.addEventListener("click", _cuOpen);
    document.getElementById("cu-close")?.addEventListener("click", _cuClose);
    document.getElementById("cu-backdrop")?.addEventListener("click", _cuClose);
    document.getElementById("cu-cancel-btn")?.addEventListener("click", _cuClose);
    document.getElementById("cu-parse-btn")?.addEventListener("click", _cuParse);
    document.getElementById("cu-clear-btn")?.addEventListener("click", () => {
      document.getElementById("cu-paste-area").value = "";
      document.getElementById("cu-parse-msg").textContent = "";
      document.getElementById("cu-preview-wrap").hidden = true;
      _cuParsed = [];
    });
    document.getElementById("cu-save-btn")?.addEventListener("click", _cuSave);
    document.getElementById("cu-paste-area")?.addEventListener("keydown", (e) => {
      if (e.key === "Enter" && e.ctrlKey) _cuParse();
    });
  })();

  // ========== 조편성 모달 ==========
  (function _initTeamAssign() {
    let _taStudents = [];

    function openTeamAssignModal() {
      const { region, cohort, stepVal } = _rmGetState();
      if (!cohort) { toast("기수를 선택해주세요.", "warn"); return; }
      _taStudents = state.students.filter((s) => s.region === region && s.cohort === cohort);
      if (!_taStudents.length) { toast("해당 기수의 교육생이 없습니다.", "warn"); return; }
      const titleEl = document.getElementById("ta-title");
      if (titleEl) titleEl.textContent = `${region} ${cohort} Step${stepVal}`;
      _taRenderTable();
      document.getElementById("modal-team-assign").hidden = false;
    }

    function _taRenderTable() {
      const tbody = document.getElementById("ta-tbody");
      if (!tbody) return;
      const rows = _taStudents.map((s) => {
        const base = Number(s.base || 0);
        const team = Number(s.team) || 0;
        const opts = Array.from({ length: 10 }, (_, i) => {
          const n = i + 1;
          return `<option value="${n}"${team === n ? " selected" : ""}>${n}조</option>`;
        }).join("");
        return `<tr>
          <td style="text-align:center"><input type="checkbox" class="ta-chk"></td>
          <td>${escapeHtml(s.region || "")}</td>
          <td>${escapeHtml(s.center || "")}</td>
          <td>${escapeHtml(s.branch || "")}</td>
          <td><strong>${escapeHtml(s.name || "")}</strong></td>
          <td>${escapeHtml(s.empNo || "")}</td>
          <td class="r">${Nf(base)}</td>
          <td>${escapeHtml(s.phone || "")}</td>
          <td><select class="ta-grp-sel${team ? " ta-grp-set" : ""}" data-emp="${escapeHtml(s.empNo)}" data-base="${base}">
            <option value="0"${!team ? " selected" : ""}>-</option>
            ${opts}
          </select></td>
        </tr>`;
      });
      tbody.innerHTML = rows.join("");
      tbody.querySelectorAll(".ta-grp-sel").forEach((sel) => {
        sel.addEventListener("change", (e) => {
          e.target.classList.toggle("ta-grp-set", Number(e.target.value) > 0);
          _taUpdateSums();
        });
      });
      _taUpdateSums();
    }

    function _taUpdateSums() {
      const bar = document.getElementById("ta-sum-bar");
      if (!bar) return;
      const sums = {};
      document.querySelectorAll(".ta-grp-sel").forEach((sel) => {
        const g = Number(sel.value);
        const base = Number(sel.dataset.base || 0);
        if (g > 0) sums[g] = (sums[g] || 0) + base;
      });
      const keys = Object.keys(sums).map(Number).sort((a, b) => a - b);
      if (!keys.length) {
        bar.innerHTML = `<span class="ta-sum-empty">조를 배정하면 합계가 표시됩니다.</span>`;
      } else {
        bar.innerHTML = keys.map((g) =>
          `<span class="ta-sum-item">${g}조: ${sums[g].toLocaleString()}원</span>`
        ).join(" &nbsp;|&nbsp; ");
      }
    }

    async function _taSave() {
      const btn = document.getElementById("btn-ta-save");
      const msgEl = document.getElementById("ta-save-msg");
      const sels = document.querySelectorAll(".ta-grp-sel");
      const updates = [];
      sels.forEach((sel) => {
        const emp = sel.dataset.emp;
        const g = Number(sel.value);
        const st = state.students.find((x) => x.empNo === emp);
        if (!st) return;
        const newTeam = g > 0 ? String(g) : "";
        if ((st.team || "") !== newTeam) updates.push({ ...st, team: newTeam });
      });
      if (!updates.length) {
        if (msgEl) { msgEl.textContent = "변경된 조 배정이 없습니다."; msgEl.className = "pg-msg"; }
        setTimeout(() => { if (msgEl) msgEl.textContent = ""; }, 2000);
        return;
      }
      if (btn) { btn.disabled = true; btn.textContent = "저장중..."; }
      try {
        await window.DataAPI.saveMany(updates);
        if (msgEl) { msgEl.textContent = `✅ ${updates.length}명 저장 완료`; msgEl.className = "pg-msg ok"; }
        setTimeout(() => { if (msgEl) msgEl.textContent = ""; }, 3000);
        renderDebounced();
      } catch (e) {
        if (msgEl) { msgEl.textContent = "❌ 저장 실패: " + e.message; msgEl.className = "pg-msg err"; }
      } finally {
        if (btn) { btn.disabled = false; btn.textContent = "💾 저장"; }
      }
    }

    document.getElementById("btn-rm-team-assign")?.addEventListener("click", openTeamAssignModal);
    document.getElementById("ta-close")?.addEventListener("click", () => {
      document.getElementById("modal-team-assign").hidden = true;
    });
    document.getElementById("ta-backdrop")?.addEventListener("click", () => {
      document.getElementById("modal-team-assign").hidden = true;
    });
    document.getElementById("btn-ta-save")?.addEventListener("click", _taSave);
    document.getElementById("ta-chk-all")?.addEventListener("change", (e) => {
      document.querySelectorAll(".ta-chk").forEach((chk) => { chk.checked = e.target.checked; });
    });
    document.querySelector(".ta-action-bar")?.addEventListener("click", (e) => {
      const btn = e.target.closest(".ta-assign-btn");
      if (!btn) return;
      const g = btn.dataset.g;
      const checkedRows = [...document.querySelectorAll(".ta-chk:checked")].map((chk) => chk.closest("tr"));
      if (!checkedRows.length) { toast("배정할 교육생을 체크해주세요.", "warn"); return; }
      checkedRows.forEach((row) => {
        const sel = row.querySelector(".ta-grp-sel");
        if (!sel) return;
        sel.value = g;
        sel.classList.toggle("ta-grp-set", Number(g) > 0);
      });
      _taUpdateSums();
    });
  })();

  async function _rmPasteApply() {
    const { region, cohort, stepVal, sfx } = _rmGetState();
    const txt = (document.getElementById("rm-paste-area")?.value || "").trim();
    const msgEl = document.getElementById("rm-paste-msg");
    const setMsg = (t, cls = "") => { if (msgEl) { msgEl.textContent = t; msgEl.className = "pg-msg" + (cls ? " " + cls : ""); } };
    if (!txt) { setMsg("❌ 붙여넣을 내용이 없습니다.", "err"); return; }
    if (!region) { setMsg("❌ 지역단이 설정되지 않았습니다.", "err"); return; }

    const lines = txt.split(/\r?\n/).filter((l) => l.trim());
    if (!lines.length) return;

    // 헤더 자동 감지
    const firstRow = lines[0].trim().split(/\t/).map((c) => c.trim());
    const isHeader = firstRow.some((h) => Object.prototype.hasOwnProperty.call(PG_HEADER_AUTOMAP, h));
    let dataLines, colDefs;
    if (isHeader) {
      dataLines = lines.slice(1);
      colDefs = firstRow.map((h) => ({ label: h, field: PG_HEADER_AUTOMAP[h] ?? "ignore" }));
    } else {
      dataLines = lines;
      // 열 수에 따라 기본 매핑 결정 (관리자 붙여넣기와 동일 로직)
      let effectiveCols;
      if (firstRow.length === 2) {
        // 2열 단순 포맷: 사번 + 현재실적 (스텝에 따라 pgCurrent 또는 pgCurrent2)
        effectiveCols = ["empNo", `pgCurrent${sfx}`];
      } else if (stepVal === "2" && firstRow.length >= 18) {
        effectiveCols = PG_STEP2_COMBINED_COLS;
      } else if (firstRow.length === 10) {
        effectiveCols = PG_STEP_UNIFIED_COLS;
      } else if (firstRow.length === 13) {
        effectiveCols = PG_STEP1_SHORT_COLS;
      } else {
        effectiveCols = PG_DEFAULT_COLS;
      }
      colDefs = firstRow.map((_, i) => ({
        label: PG_FIELD_OPTIONS.find((o) => o.value === (effectiveCols[i] || "ignore"))?.label ?? `열 ${i + 1}`,
        field: effectiveCols[i] || "ignore",
      }));
    }
    if (!dataLines.length) { setMsg("❌ 데이터가 없습니다.", "err"); return; }

    // 2열 단순 포맷(사번+현재실적) 감지: 열 매핑 팝업 생략하고 바로 저장 확인으로
    const _isSimple2Col = colDefs.length === 2 &&
      colDefs.some((c) => c.field === "empNo") &&
      colDefs.some((c) => c.field === `pgCurrent${sfx}`);

    // 랜덤 3개 샘플 행
    const _rmSampleIdxs = (() => {
      const n = dataLines.length;
      if (n <= 3) return Array.from({ length: n }, (_, i) => i);
      const s = new Set();
      while (s.size < 3) s.add(Math.floor(Math.random() * n));
      return [...s].sort((a, b) => a - b);
    })();
    const sampleRows = _rmSampleIdxs.map((i) => dataLines[i].trim().split(/\t/).map((c) => c.trim()));

    // 열 매핑 확인 팝업 — 2열 단순 포맷은 생략 (매핑 자명)
    let fieldMapping;
    if (_isSimple2Col) {
      fieldMapping = colDefs.map((c) => c.field);
    } else {
      fieldMapping = await openPgColMapModal(colDefs, sampleRows);
      if (!fieldMapping) return;
    }

    const parseAmt = (v) => parseInt((v || "").replace(/,/g, "").trim(), 10) || 0;
    const getCol = (p, f) => { const i = fieldMapping.indexOf(f); return i >= 0 ? (p[i] || "") : ""; };
    const getAmt = (p, f) => { const i = fieldMapping.indexOf(f); return i >= 0 ? parseAmt(p[i]) : 0; };

    const updates = [];
    const unmatched = [];
    dataLines.forEach((line) => {
      const p = line.trim().split(/\t/).map((c) => c.replace(/,/g, "").trim());
      const empNo = getCol(p, "empNo").replace(/[/\\\s]/g, "");
      if (!empNo) return;
      const st = state.students.find((s) => s.empNo === empNo && (!cohort || !s.cohort || s.cohort === cohort));
      if (!st) { unmatched.push(empNo); return; }

      const pgBase      = getAmt(p, "pgBase");
      const pgCurrent   = getAmt(p, "pgCurrent");
      const pgIpumCount = getAmt(p, "pgIpumCount");
      const pgIpumAmt   = getAmt(p, "pgIpumAmt");
      const name        = getCol(p, "name");
      const branch      = getCol(p, "branch");
      const pgMonth     = getCol(p, "pgMonth");

      const pgFields = {};
      if (fieldMapping.includes("pgBase"))       pgFields.base                  = pgBase;
      if (fieldMapping.includes("pgCurrent"))    pgFields[`pgCurrent${sfx}`]   = pgCurrent;
      if (fieldMapping.includes("pgIpumCount"))  pgFields[`pgIpumCount${sfx}`] = pgIpumCount;
      if (fieldMapping.includes("pgIpumAmt"))    pgFields[`pgIpumAmt${sfx}`]   = pgIpumAmt;

      updates.push({
        ...st,
        ...(branch  ? { branch }  : {}),
        ...(pgMonth ? { pgMonth } : {}),
        ...pgFields,
      });
    });

    if (!updates.length) {
      setMsg(`❌ 매칭된 교육생이 없습니다.${unmatched.length ? ` (미매칭 사번: ${unmatched.slice(0,5).join(", ")})` : ""}`, "err");
      return;
    }

    // 기수·스텝 저장 확인 — 관리자 붙여넣기와 동일한 확인 팝업
    if (!await openPasteSaveConfirmModal(region, cohort, stepVal, updates, [])) return;

    const btn = document.getElementById("btn-rm-paste-apply");
    if (btn) { btn.disabled = true; btn.textContent = "저장중..."; }
    try {
      const result = await window.DataAPI.saveMany(updates);
      const committed  = result?.committed ?? updates.length;
      const warnPart = unmatched.length ? ` (미매칭 ${unmatched.length}건 제외)` : "";
      setMsg(`✅ ${committed}명 저장 완료${warnPart}`, "ok");
      if (committed > 0) document.getElementById("rm-paste-area").value = "";
      setTimeout(() => { if (msgEl) msgEl.textContent = ""; }, 4000);
      if (committed > 0) renderDebounced();
    } catch (e) {
      setMsg("❌ 저장 실패: " + e.message, "err");
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = "📥 실적진도 저장"; }
    }
  }

  function bindRegionMgrEvents() {
    document.getElementById("btn-open-region-mgr")?.addEventListener("click", openRegionMgrModal);
    document.getElementById("rm-close")?.addEventListener("click", _rmClose);
    document.getElementById("rm-backdrop")?.addEventListener("click", _rmClose);
    document.getElementById("btn-rm-paste-apply")?.addEventListener("click", _rmPasteApply);
    document.getElementById("btn-rm-paste-clear")?.addEventListener("click", () => {
      const a = document.getElementById("rm-paste-area"); if (a) a.value = "";
      const m = document.getElementById("rm-paste-msg"); if (m) m.textContent = "";
    });
    document.getElementById("btn-rm-table-save")?.addEventListener("click", _rmSaveTable);
    document.getElementById("btn-rm-open-award")?.addEventListener("click", () => {
      const { region, cohort, stepVal } = _rmGetState();
      _rmClose();
      openAwardPlanModal({ region, cohort, step: stepVal });
    });
    // tab switching
    document.querySelectorAll(".rm-tab").forEach((btn) => {
      btn.addEventListener("click", () => {
        document.querySelectorAll(".rm-tab").forEach((b) => b.classList.remove("active"));
        btn.classList.add("active");
        const tab = btn.dataset.tab;
        ["paste","table","award"].forEach((t) => {
          const el = document.getElementById(`rm-tab-${t}`);
          if (el) el.hidden = t !== tab;
        });
        if (tab === "table") _rmRenderTable();
      });
    });
    // re-render table on cohort/step change
    document.getElementById("rm-cohort-sel")?.addEventListener("change", () => {
      const activeTab = document.querySelector(".rm-tab.active")?.dataset.tab;
      if (activeTab === "table") _rmRenderTable();
    });
    document.querySelectorAll('input[name="rm-step"]').forEach((r) => {
      r.addEventListener("change", () => {
        const activeTab = document.querySelector(".rm-tab.active")?.dataset.tab;
        if (activeTab === "table") _rmRenderTable();
      });
    });
    _bindRmImagePaste();
  }

  document.addEventListener("DOMContentLoaded", init);
})();
