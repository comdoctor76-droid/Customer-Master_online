/* 고객컨설팅 마스터과정 운영관리 - 메인 앱 로직 */

(function () {
  const LS_KEY = "cmf.filter.v1";
  const LS_DEFAULTS_KEY = "cmf.masterTargetDefaults.v1";
  const DEFAULT_MASTER_TARGET = 200000; // 원 (= 200,000원)
  const DEFAULT_REGION = "호남지역단";
  // 앱 버전 — 코드 수정(커밋)마다 0.01 씩 증가
  const APP_VERSION = "0.94";

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

  function loadFilter() {
    const base = { region: DEFAULT_REGION, center: "", branch: "", cohort: "", q: "" };
    try {
      const saved = JSON.parse(localStorage.getItem(LS_KEY) || "null");
      if (saved && typeof saved === "object") {
        return { ...base, ...saved, q: "" };
      }
    } catch (e) {}
    return base;
  }

  function persistFilter() {
    const { region, center, branch, cohort } = state.filter;
    try {
      localStorage.setItem(LS_KEY, JSON.stringify({ region, center, branch, cohort }));
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
      if (f.cohort && s.cohort !== f.cohort) return false;
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
    $("#kpi-base").textContent = sum("base").toLocaleString();
    $("#kpi-target").textContent = sum("target").toLocaleString();
    $("#kpi-honors").textContent = sum("honors").toLocaleString();
    $("#mini-total").textContent = state.students.length;
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
        return `
          <details class="branch-mini" data-branch="${escapeHtml(branch)}"${bOpen ? " open" : ""}>
            <summary class="branch-mini-head">
              <span class="branch-name">${escapeHtml(branch)}</span>
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
                    <span class="s-phone">${escapeHtml(s.phone || "")}</span>
                  </span>
                  ${ccBadge}
                </li>
              `;}).join("")}
            </ul>
          </details>
        `;
      }).join("");
      return `
        <details class="center-mini" data-center="${escapeHtml(center)}"${centerOpen ? " open" : ""}>
          <summary class="center-mini-head">
            <span class="center-name">🏢 ${escapeHtml(center)}</span>
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
  }

  // ========== 교육생 선택 → 면담 관리 ==========
  function selectStudent(empNo) {
    state.selectedEmpNo = empNo;
    // 기존 구독 해제
    if (state.consultUnsub) { state.consultUnsub(); state.consultUnsub = null; }
    state.consultations = [];
    renderSidebarStudentList(filteredStudents());
    renderStudentDetail();
    // 모바일: 교육생 선택 시 사이드바 닫고 교육생 관리 패널로 전환
    if (isMobileViewport()) {
      closeMobileSidebar();
      switchView("#students");
    }
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
          <div class="sib-meta">${escapeHtml([s.center, s.branch, s.cohort, s.phone].filter(Boolean).join(" · "))}</div>
        </div>
        <div class="sib-stats">
          <div><span>평균실적</span><strong>${fmt(Number(s.base))}</strong></div>
          <div><span>마스터목표</span><strong>${fmt(s.region !== "호남지역단" && !Number(s.target) ? getProgressStat(s).base + 50000 : Number(s.target))}</strong></div>
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
            <input type="number" id="iv-curAct" placeholder="원" step="1000">
          </div>
          <div class="iv-field">
            <label>진도 <span class="iv-hint" id="iv-pct-hint"></span></label>
            <input type="number" id="iv-pct" placeholder="%">
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
                <label>개인 평균실적 <em>*</em> <span class="iv-hint">(원)</span></label>
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

    const honorRows = HONORS.map((h, i) => {
      const achieved = tgt >= h.critVal;
      const isActive = i === award1Idx;
      const cls = isActive ? "rs-on" : achieved ? "rs-done" : "rs-miss";
      const icon = isActive ? "🏆" : achieved ? "✅" : "⬜";
      // 모바일에선 현재 등급 ± 1 만 노출 (data-dist 로 CSS 필터)
      const dist = award1Idx >= 0 ? Math.abs(i - award1Idx) : 99;
      // 한글만 표시 (영문 괄호 제거) + 줄바꿈 방지
      const gradeKo = h.grade.replace(/\s*\([^)]*\)\s*/g, "").trim();
      // 시상금 짧게: 500만 → "5백만원", 100만 → "1백만원", 나머지 "NN만원"
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
      </div>
    `;
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
    await doSaveInterview(rec);
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
    const awardHtml = (state.printMode === "personal" && calcItv) ? buildPrintAwardHtml(calcItv, allItvs) : "";
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

    const apvHtml = `
      <div class="apv-wrap">
        <div class="apv-box"><div class="apv-title">면담자</div><div class="apv-sign"></div></div>
        <div class="apv-box"><div class="apv-title">조직파트장</div><div class="apv-sign"></div></div>
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
          ${isVeryFirst ? `<div class="print-logo-wrap">
            <div class="print-logo-sub">고객컨설팅 MASTER과정 지점별 면담일지<br><small>비전센터장 활동관리 시스템</small></div>
          </div>` : ""}
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
  function buildAwardSheetPageHtml(student, calcItv, lastInsItv) {
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

    return `
      <div class="pg">
        <div class="hdr">
          <div class="hdr-title">🏆 ${escapeHtml(vc || region)} 시상 예상답안지</div>
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
            <div class="stat-row"><span class="stat-key bl">고객마스터:</span><span class="stat-val bl">${fmtWon(b.insRaw)}</span></div>
          </div>
          <div class="info-card stat">
            <div class="stat-lbl">🎯 기본순증목표</div>
            <div class="stat-row"><span class="stat-key">아너스기본목표:</span><span class="stat-val">${b.baseTgtRaw ? fmtWon(b.baseTgtRaw) : "—"}</span></div>
            <div class="stat-row"><span class="stat-key bl">고객마스터 희망:</span><span class="stat-val bl">${fmtWon(b.tgtRaw)}</span></div>
          </div>
        </div>
        <div class="sec-title">📌 아너스 희망목표금액 기준 시상 (${fmtWon(b.tgtRaw)})</div>
        ${award1Html}${award2Html}
        <div class="sec-title bl">🎯 고객컨설팅마스터 과정 개인시상</div>
        ${award3Html}
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
      </div>
    `;
  }

  // 시상안 출력용 CSS (스탠드얼론 인쇄창에 삽입)
  const AWARD_PRINT_CSS = `
    *{box-sizing:border-box;margin:0;padding:0;}
    body{font-family:'Noto Sans KR','Malgun Gothic',sans-serif;background:#fff;color:#1A1A1A;font-size:12px;line-height:1.4;}
    .pg{padding:8mm 10mm;page-break-after:always;}
    .pg:last-child{page-break-after:auto;}
    .hdr{background:linear-gradient(135deg,#1A2744,#2C3F6E);border-radius:8px;padding:8px 12px;margin-bottom:8px;display:flex;align-items:center;justify-content:space-between;}
    .hdr-title{color:#fff;font-size:14px;font-weight:900;}
    .hdr-date{color:rgba(255,255,255,.65);font-size:11px;}
    .info-row1{display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:5px;margin-bottom:5px;}
    .info-row2{display:grid;grid-template-columns:1fr 1fr;gap:5px;margin-bottom:8px;}
    .info-card{background:#F5F7FF;border-radius:6px;padding:6px 8px;text-align:center;border:1px solid #D5DCF5;}
    .info-lbl{font-size:10px;color:#5C6BC0;font-weight:700;margin-bottom:2px;}
    .info-val{font-size:13px;font-weight:900;color:#1A2744;}
    .info-card.key{background:#1A2744;border-color:#1A2744;}
    .info-card.key .info-lbl{color:rgba(255,255,255,.65);}
    .info-card.key .info-val{color:#fff;font-size:15px;}
    .info-card.stat{text-align:left;padding:7px 10px;background:#F8F9FF;}
    .stat-lbl{font-size:11px;color:#5C6BC0;font-weight:800;margin-bottom:4px;padding-bottom:3px;border-bottom:1px solid #D5DCF5;}
    .stat-row{display:flex;justify-content:space-between;margin-bottom:2px;}
    .stat-key{font-size:11px;font-weight:700;color:#4A148C;}
    .stat-key.bl{color:#1565C0;}
    .stat-val{font-size:14px;font-weight:900;color:#1A2744;}
    .stat-val.bl{color:#1565C0;}
    .sec-title{font-size:12px;font-weight:800;color:#4A148C;margin:6px 0 3px;}
    .sec-title.bl{color:#1565C0;}
    .hl-row{display:flex;align-items:center;gap:8px;background:#E8F5E9;border-radius:7px;padding:6px 10px;margin-bottom:4px;border-left:3px solid #2E7D32;}
    .hl-row.warn{background:#FFEBEE;border-left-color:#C62828;}
    .hl-row.blue{background:#E3F2FD;border-left-color:#1565C0;}
    .hl-icon{font-size:16px;flex-shrink:0;}
    .hl-info{flex:1;min-width:0;}
    .hl-grade{font-size:12px;font-weight:700;color:#1B5E20;}
    .hl-row.warn .hl-grade{color:#C62828;}
    .hl-row.blue .hl-grade{color:#1565C0;}
    .hl-crit{font-size:11px;color:#388E3C;margin-top:1px;line-height:1.5;}
    .hl-row.warn .hl-crit{color:#E57373;}
    .hl-row.blue .hl-crit{color:#1976D2;}
    .hl-crit .lbl{background:#1565C0;color:#fff;font-size:10px;padding:1px 6px;border-radius:8px;font-weight:700;}
    .hl-sub{font-size:12px;font-weight:800;color:#0D47A1;margin-top:2px;}
    .hl-amt{font-size:16px;font-weight:900;color:#2E7D32;white-space:nowrap;}
    .hl-amt.rd{color:#C62828;}
    .hl-amt.bl{color:#1565C0;font-size:18px;}
    .hl-none{padding:5px 10px;color:#999;font-size:11px;background:#F5F5F5;border-radius:6px;margin-bottom:4px;}
    .total-bar{background:linear-gradient(135deg,#1A2744,#2C3F6E);border-radius:7px;padding:8px 14px;display:flex;align-items:center;justify-content:space-between;margin:6px 0 6px;}
    .total-lbl{color:rgba(255,255,255,.8);font-size:11px;font-weight:700;}
    .total-sub{color:rgba(255,255,255,.55);font-size:10px;margin-top:2px;}
    .total-val{color:#FFE082;font-size:22px;font-weight:900;}
    .up-title{font-size:11px;font-weight:800;color:#1565C0;margin:5px 0 3px;}
    .up-table{width:100%;border-collapse:collapse;font-size:10px;}
    .up-table th{background:#E3F2FD;padding:3px 4px;border:1px solid #BBDEFB;color:#1565C0;font-weight:700;font-size:10px;}
    .up-table td{padding:3px 4px;border:1px solid #E3E3E3;text-align:center;font-size:10px;}
    .up-table td.rd{color:#C62828;font-weight:700;}
    .up-table td.bl{color:#1565C0;}
    .up-table td.grn{color:#1B5E20;font-size:11px;}
    .up-table tr.up-next td{background:#FFF9C4;font-weight:700;}
    .warn{color:#C62828;font-size:9px;font-weight:800;}
    .note{font-size:9.5px;color:#888;margin-top:4px;line-height:1.4;}
    @media print{@page{size:A4 portrait;margin:6mm 8mm;}}
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
    if (f.cohort) students = students.filter((s) => s.cohort === f.cohort);
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

  const PROGRESS_AWARDS = {
    // 순증액별 고정시상/비율시상 (원)
    tiers: [
      { min: 500000, type: "pct", val: 1.5 },
      { min: 300000, type: "pct", val: 1.2 },
      { min: 200000, type: "fixed", val: 200000 },
      { min: 100000, type: "fixed", val: 100000 },
      { min: 50000,  type: "fixed", val: 50000 }
    ],
    rateTop10: [300000, 200000, 200000, 50000, 50000, 50000, 50000, 50000, 50000, 50000], // 1위 30만 / 2~3위 20만 / 4~10위 주유권 5만
    amtTop10:  [500000, 300000, 300000, 100000, 100000, 100000, 100000, 100000, 100000, 100000], // 1위 50만 / 2~3위 30만 / 4~10위 주유권 10만
    minNetForRank: 300000 // 개인 기준: 순증 30만 이상
  };

  function openProgressRegionPicker() {
    // 학생이 존재하는 지역단만 추출
    const regions = [...new Set(state.students.map((s) => s.region).filter(Boolean))].sort();
    if (!regions.length) { toast("등록된 교육생이 없습니다.", "error"); return; }
    // 간단 모달 대신 prompt + select — 모달 재사용 대신 빠른 구현
    openPickerModal("지역단 선택", regions, (picked) => {
      state.progressRegion = picked;
      document.getElementById("progress-region-label").textContent = picked;
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

  function renderProgressPanel() {
    const body = $("#progress-body");
    if (!body) return;
    // 좌측 필터에서 선택한 지역단 기준
    const region = state.filter.region || "";
    state.progressRegion = region;
    const label = document.getElementById("progress-region-label");
    if (label) label.textContent = region || "지역단 미선택";

    if (!region) {
      body.innerHTML = `<div class="empty-state">좌측 필터에서 <strong>지역단</strong>을 선택하면 해당 지역단의 실적진도가 표시됩니다.</div>`;
      return;
    }
    const list = state.students.filter((s) => s.region === region);
    if (!list.length) {
      body.innerHTML = `<div class="empty-state">${escapeHtml(region)} 에 등록된 교육생이 없습니다.</div>`;
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

  // 학생 데이터에서 계산된 지표 얻기
  function getProgressStat(s) {
    // pg* 필드(원 단위, 실적진도현황 붙여넣기)가 있으면 우선 사용; 없으면 천원→원 변환
    const base    = s.pgBase    !== undefined ? Number(s.pgBase)    : Number(s.base    || 0);
    const current = s.pgCurrent !== undefined ? Number(s.pgCurrent) : Number(s.current || 0);
    const hiCap   = Number(s.hiCap || 0);
    const ipumCount = s.pgIpumCount !== undefined ? Number(s.pgIpumCount) : Number(s.ipumCount || 0);
    const ipumAmt   = s.pgIpumAmt   !== undefined ? Number(s.pgIpumAmt)   : Number(s.ipumAmt   || 0);
    const net  = current - base;
    const rate = base > 0 ? (current / base) * 100 : 0;
    return { s, base, current, hiCap, net, rate, ipumCount, ipumAmt };
  }

  function tierAward(net) {
    for (const t of PROGRESS_AWARDS.tiers) {
      if (net >= t.min) {
        return t.type === "pct" ? Math.round(net * t.val) : t.val;
      }
    }
    return 0;
  }
  function tierLabel(net) {
    if (net >= 500000) return "150% 지급";
    if (net >= 300000) return "120% 지급";
    if (net >= 200000) return "20만원";
    if (net >= 100000) return "10만원";
    if (net >= 50000)  return "5만원";
    return "-";
  }
  const Nf = (v) => Math.round(Number(v) || 0).toLocaleString();
  const RB = (r) => {
    if (!r) return `<span style="color:#ccc;font-size:10px;">-</span>`;
    const cls = r === 1 ? "r1" : r === 2 ? "r2" : r === 3 ? "r3" : "rt";
    return `<span class="pg-rb ${cls}">${r}</span>`;
  };

  function renderProgressHome(list) {
    const stats = list.map(getProgressStat);
    const total = stats.length;
    const avgR = stats.reduce((a, s) => a + s.rate, 0) / total;
    const over5 = stats.filter((s) => s.net >= 50000).length;
    const elig = stats.filter((s) => s.net >= PROGRESS_AWARDS.minNetForRank);
    const byRate = [...stats].sort((a, b) => (b.net / (b.base || 1)) - (a.net / (a.base || 1)));
    const byAmt  = [...stats].sort((a, b) => b.net - a.net);
    const byIpum = [...stats].filter((s) => s.ipumAmt > 0).sort((a, b) => b.ipumAmt - a.ipumAmt || b.ipumCount - a.ipumCount);

    // 신장액 TOP2 (기준 30만 이상)는 신장률 시상에서 제외
    const amtExcludeIds = new Set(byAmt.slice(0, 2).filter((s) => s.net >= PROGRESS_AWARDS.minNetForRank).map((s) => s.s.empNo));
    const rateFinalList = byRate.filter((s) => !amtExcludeIds.has(s.s.empNo));

    const a5 = stats.filter((s) => s.net >= 500000).length;
    const a4 = stats.filter((s) => s.net >= 300000 && s.net < 500000).length;

    // 그룹 시상 — team 필드가 설정된 학생이 있으면 team 기준, 아니면 branch(지점)
    const hasAnyTeam = stats.some((s) => (s.s.team || "").toString().trim());
    const groupKeyFn = hasAnyTeam
      ? ((s) => (s.s.team || "").toString().trim() || "(팀 미배정)")
      : ((s) => s.s.branch || "(미지정)");
    const groupMap = {};
    stats.forEach((st) => {
      const k = groupKeyFn(st);
      if (!groupMap[k]) groupMap[k] = { base: 0, current: 0, members: [] };
      groupMap[k].base += st.base;
      groupMap[k].current += st.current;
      groupMap[k].members.push(st.s.name || "");
    });
    const groupRanking = Object.entries(groupMap).map(([name, g]) => ({
      name,
      rate: g.base > 0 ? (g.current / g.base) * 100 : 0,
      members: g.members,
      base: g.base,
      current: g.current
    })).sort((a, b) => b.rate - a.rate);
    const groupLabel = hasAnyTeam ? "팀별 인보험 순증" : "지점별 인보험 순증 (팀 미배정)";

    // TOP3 미리보기 행 (신장률/신장액/인품/그룹 공통)
    const pcardRateTop3 = rateFinalList.slice(0, 3).map((st, i) => {
      const rate = st.base > 0 ? (st.net / st.base) * 100 : 0;
      const belowMin = st.net < PROGRESS_AWARDS.minNetForRank;
      const prizeAmt = PROGRESS_AWARDS.rateTop10[i] || 0;
      const prizeTxt = belowMin ? "기준미달" :
        (prizeAmt >= 100000 ? `시상 ${Math.round(prizeAmt/10000)}만원` : `주유권 ${Math.round(prizeAmt/10000)}만`);
      return `<li class="pg-pcard-row" data-emp="${escapeHtml(st.s.empNo)}">
        <span class="pg-rb ${i===0?"r1":i===1?"r2":"r3"}">${i+1}</span>
        <div class="pg-pcard-content">
          <div class="pg-pcard-nm"><strong>${escapeHtml(st.s.name||"")}</strong> <span class="pg-pcard-val">${rate.toFixed(1)}%</span></div>
          <span class="pg-pcard-prize ${belowMin?"pg-b-no":""}">${prizeTxt}</span>
        </div>
      </li>`;
    }).join("");

    const pcardAmtTop3 = byAmt.slice(0, 3).map((st, i) => {
      const belowMin = st.net < PROGRESS_AWARDS.minNetForRank;
      const prizeAmt = PROGRESS_AWARDS.amtTop10[i] || 0;
      const prizeTxt = belowMin ? "기준미달" :
        (prizeAmt >= 200000 ? `시상 ${Math.round(prizeAmt/10000)}만원` : `주유권 ${Math.round(prizeAmt/10000)}만`);
      return `<li class="pg-pcard-row" data-emp="${escapeHtml(st.s.empNo)}">
        <span class="pg-rb ${i===0?"r1":i===1?"r2":"r3"}">${i+1}</span>
        <div class="pg-pcard-content">
          <div class="pg-pcard-nm"><strong>${escapeHtml(st.s.name||"")}</strong> <span class="pg-pcard-val">${st.net>=0?"+":""}${Nf(st.net)}원</span></div>
          <span class="pg-pcard-prize ${belowMin?"pg-b-no":""}">${prizeTxt}</span>
        </div>
      </li>`;
    }).join("");

    const pcardIpumTop3 = byIpum.slice(0, 3).map((st, i) => {
      const grade = ["인품의 황제", "인품의 제왕", "인품의 왕"][i] || "";
      return `<li class="pg-pcard-row" data-emp="${escapeHtml(st.s.empNo)}">
        <span class="pg-rb ${i===0?"r1":i===1?"r2":"r3"}">${i+1}</span>
        <div class="pg-pcard-content">
          <div class="pg-pcard-nm"><strong>${escapeHtml(st.s.name||"")}</strong> <span class="pg-pcard-val">${Nf(st.ipumAmt)}원</span></div>
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

    // 풀스크린 모달에 띄울 데이터 캐시 — 전체 순위 노출 (Infinity 사용)
    state._pgCardFullData = {
      rate: {
        title: `📈 신장률 전체 순위 (${rateFinalList.length}명)`,
        subtitle: "기준실적 대비 순증률 — 신장액 TOP2 제외",
        bodyHTML: renderProgressTop10(rateFinalList, "rate", Infinity)
      },
      amt: {
        title: `💰 신장액 전체 순위 (${byAmt.length}명)`,
        subtitle: "순증 금액 절대값",
        bodyHTML: renderProgressTop10(byAmt, "amt", Infinity)
      },
      ipum: {
        title: `✨ 인품왕 전체 순위 (${byIpum.length}명)`,
        subtitle: "신상품 판매액 기준",
        bodyHTML: byIpum.length ? renderProgressTop10(byIpum, "ipum", Infinity) : `<div class="pg-empty">실적관리 탭에서 인품 데이터를 입력하세요.</div>`
      },
      group: {
        title: `🏅 그룹 순증 전체 순위 (${groupRanking.length}${hasAnyTeam ? "팀" : "개 지점"})`,
        subtitle: groupLabel,
        bodyHTML: `
          <table class="pg-tbl pg-tbl-wide">
            <thead><tr><th>#</th><th>${hasAnyTeam ? "팀" : "지점"}</th><th>인원</th><th>기준</th><th>현재</th><th>달성률</th></tr></thead>
            <tbody>${groupRanking.map((g, i) => `<tr><td>${RB(i+1)}</td><td><strong>${escapeHtml(g.name)}</strong></td><td>${g.members.length}명</td><td class="r">${Nf(g.base)}</td><td class="r">${Nf(g.current)}</td><td>${g.rate.toFixed(1)}%</td></tr>`).join("")}</tbody>
          </table>
        `
      }
    };

    // 시상안 박스 + KPI 영역
    return `
      <div class="pg-wrap">
        <div class="pg-award-box">
          <h3>📋 ${escapeHtml(state.progressRegion)} 고객마스터 활성화 시상안</h3>
          <div class="pg-tier-row">
            <div class="pg-tier"><div class="pg-tc">순증 5만원↑</div><div class="pg-tn">5만원</div></div>
            <div class="pg-tier"><div class="pg-tc">순증 10만원↑</div><div class="pg-tn">10만원</div></div>
            <div class="pg-tier"><div class="pg-tc">순증 20만원↑</div><div class="pg-tn">20만원</div></div>
            <div class="pg-tier"><div class="pg-tc">순증 30만원↑</div><div class="pg-tn">120%</div></div>
            <div class="pg-tier"><div class="pg-tc">순증 50만원↑</div><div class="pg-tn">150%</div></div>
          </div>
          <div class="pg-an"><strong>📈 신장률 TOP10:</strong> 1위 30만 / 2~3위 20만 / 4~10위 주유권 5만</div>
          <div class="pg-an"><strong>💰 신장액 TOP10:</strong> 1위 50만 / 2~3위 30만 / 4~10위 주유권 10만</div>
          <div class="pg-an-crit">⚠️ 개인 시상 기준: 기준실적 대비 <strong>순증 30만원 이상</strong> 달성자</div>
          <div class="pg-an-warn">※ 신장률·신장액 중복시상 없음 — 신장액 TOP2 적용 시 신장률에서 제외</div>
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
          <div class="pg-pcard pg-pcard-rate" data-pcard="rate" role="button" tabindex="0">
            <div class="pg-pcard-head">
              <div class="pg-pcard-icon">📈</div>
              <div class="pg-pcard-titles">
                <h5>최고 신장률 TOP10</h5>
                <p>기준실적 대비 순증률</p>
              </div>
              <span class="pg-pcard-chev">›</span>
            </div>
            <ol class="pg-top3-list">${pcardRateTop3}</ol>
          </div>

          <div class="pg-pcard pg-pcard-amt" data-pcard="amt" role="button" tabindex="0">
            <div class="pg-pcard-head">
              <div class="pg-pcard-icon">💰</div>
              <div class="pg-pcard-titles">
                <h5>최고 신장액 TOP10</h5>
                <p>순증 금액 절대값</p>
              </div>
              <span class="pg-pcard-chev">›</span>
            </div>
            <ol class="pg-top3-list">${pcardAmtTop3}</ol>
          </div>

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

          <div class="pg-pcard pg-pcard-group" data-pcard="group" role="button" tabindex="0">
            <div class="pg-pcard-head">
              <div class="pg-pcard-icon">🏅</div>
              <div class="pg-pcard-titles">
                <h5>그룹 순증 시상</h5>
                <p>${escapeHtml(groupLabel)}</p>
              </div>
              <span class="pg-pcard-chev">›</span>
            </div>
            <ol class="pg-top3-list">${pcardGroupTop3}</ol>
          </div>
        </div>

        <div class="pg-grid2 pg-desktop-only">
          <div class="pg-card">
            <h4>📈 신장률 TOP10 <small>(신장액 TOP2 제외)</small></h4>
            ${renderProgressTop10(rateFinalList, "rate")}
          </div>
          <div class="pg-card">
            <h4>💰 신장액 TOP10</h4>
            ${renderProgressTop10(byAmt, "amt")}
          </div>
        </div>

        <div class="pg-card pg-desktop-only">
          <h4>✨ 인품왕 TOP10 <small>(신상품 판매액 기준)</small></h4>
          ${byIpum.length ? renderProgressTop10(byIpum, "ipum") : `<div class="pg-empty">실적관리 탭에서 인품 데이터를 입력하세요.</div>`}
        </div>

        <div class="pg-card pg-full-tbl-card pg-desktop-only" id="pg-full-tbl-card">
          <h4>📊 전체 교육생 실적표 <small>(신장액 내림차순, 클릭 시 상세)</small>
            <button type="button" class="pg-full-tbl-toggle btn-outline small" id="btn-pg-full-toggle" style="display:none;">펼쳐보기</button>
          </h4>
          <div class="pg-tbl-wrap"><table class="pg-tbl">
            <thead><tr><th>순위</th><th>성명</th><th>지점</th><th>기준실적</th><th>장기하이캡</th><th>현재실적</th><th>달성률</th><th>순증</th><th>시상</th></tr></thead>
            <tbody>${byAmt.map((st, i) => {
              const nc = st.rate >= 120 ? "pg-c-over" : st.rate >= 100 ? "pg-c-good" : st.rate >= 80 ? "pg-c-mid" : "pg-c-low";
              const netC = st.net > 0 ? "pg-net-p" : st.net < 0 ? "pg-net-m" : "";
              const aw = tierLabel(st.net);
              const baseDisp = st.base > 0 ? Nf(st.base) : "—";
              const rateDisp = st.base > 0 ? st.rate.toFixed(1) + "%" : "—";
              return `<tr data-emp="${escapeHtml(st.s.empNo)}" class="pg-tr-click"><td>${RB(i + 1)}</td><td><strong>${escapeHtml(st.s.name || "")}</strong></td><td>${escapeHtml(st.s.branch || "")}</td><td class="r">${baseDisp}</td><td class="r">${st.hiCap ? Nf(st.hiCap) : "—"}</td><td class="r">${Nf(st.current)}</td><td class="${nc}">${rateDisp}</td><td class="r ${netC}">${st.net >= 0 ? "+" : ""}${Nf(st.net)}</td><td>${aw}</td></tr>`;
            }).join("")}</tbody>
          </table></div>
        </div>
      </div>
    `;
  }

  function renderProgressTop10(list, kind, limit) {
    const max = (limit === undefined || limit === null) ? 10 : limit;
    const top = (max === Infinity || max <= 0) ? list.slice() : list.slice(0, max);
    if (!top.length) return `<div class="pg-empty">데이터 없음</div>`;
    return `
      <table class="pg-tbl">
        <thead><tr><th>#</th><th>성명</th><th>지점</th><th>${kind === "ipum" ? "인품실적" : kind === "rate" ? "신장률" : "순증"}</th><th>시상</th></tr></thead>
        <tbody>${top.map((st, i) => {
          let value, prize;
          if (kind === "ipum") {
            value = Nf(st.ipumAmt) + "원";
            const grade = ["인품의 황제", "인품의 제왕", "인품의 왕"][i];
            prize = grade ? `<span class="pg-bdg pg-b-p">${grade}</span>` : "-";
          } else if (kind === "rate") {
            const rate = st.base > 0 ? (st.net / st.base) * 100 : 0;
            value = rate.toFixed(1) + "%";
            if (st.net < PROGRESS_AWARDS.minNetForRank) {
              prize = `<span class="pg-bdg pg-b-no">기준미달</span>`;
            } else {
              const amt = PROGRESS_AWARDS.rateTop10[i] || 0;
              prize = amt ? `<span class="pg-bdg pg-b-g">${amt >= 100000 ? Math.round(amt/10000)+"만원" : "주유권"+Math.round(amt/10000)+"만"}</span>` : "-";
            }
          } else {
            value = (st.net >= 0 ? "+" : "") + Nf(st.net);
            if (st.net < PROGRESS_AWARDS.minNetForRank) {
              prize = `<span class="pg-bdg pg-b-no">기준미달</span>`;
            } else {
              const amt = PROGRESS_AWARDS.amtTop10[i] || 0;
              prize = amt ? `<span class="pg-bdg pg-b-g">${amt >= 200000 ? Math.round(amt/10000)+"만원" : "주유권"+Math.round(amt/10000)+"만"}</span>` : "-";
            }
          }
          // 모달에서 클릭 가능한 행으로 표시 (data-emp 필요)
          return `<tr class="pg-tr-click" data-emp="${escapeHtml(st.s.empNo)}"><td>${RB(i + 1)}</td><td><strong>${escapeHtml(st.s.name || "")}</strong></td><td>${escapeHtml(st.s.branch || "")}</td><td class="r">${value}</td><td>${prize}</td></tr>`;
        }).join("")}</tbody>
      </table>
    `;
  }

  function bindProgressHomeEvents(list) {
    document.querySelectorAll("#progress-body .pg-tr-click").forEach((tr) => {
      tr.addEventListener("click", () => openProgressStudentPopup(tr.dataset.emp));
    });
    // 모바일 TOP3 미리보기 카드의 이름 행 클릭 → 교육생 팝업 (버블링 방지)
    document.querySelectorAll("#progress-body .pg-pcard-row[data-emp]").forEach((row) => {
      row.addEventListener("click", (e) => {
        e.stopPropagation();
        e.preventDefault();
        openProgressStudentPopup(row.dataset.emp);
      });
    });
    // 모바일 카드 자체 클릭 → 풀스크린 TOP10 모달
    document.querySelectorAll("#progress-body .pg-pcard[data-pcard]").forEach((card) => {
      card.addEventListener("click", () => {
        const key = card.dataset.pcard;
        const data = state._pgCardFullData && state._pgCardFullData[key];
        if (!data) return;
        openPgFullModal(data);
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
  }

  function openProgressStudentPopup(empNo, pushStack) {
    const s = state.students.find((x) => x.empNo === empNo);
    if (!s) return;
    const st = getProgressStat(s);
    const netAwd = tierAward(st.net);
    const awdText = netAwd > 0 ? `${Nf(netAwd)}원 (${tierLabel(st.net)})` : "해당없음";
    const rateC = st.rate >= 120 ? "#0040b0" : st.rate >= 100 ? "#006030" : st.rate >= 80 ? "#884400" : "#880000";
    const netC  = st.net > 0 ? "#0040b0" : st.net < 0 ? "#880000" : "#333";
    const initial = (s.name || "?").trim().charAt(0) || "?";
    openPgFullModal({
      title: `👤 ${escapeHtml(s.name || "")} 실적 상세`,
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
              ${s.team ? `<span class="pg-dm-id-team">${escapeHtml(s.team)}</span>` : ""}
            </div>
          </div>
        </div>
        <div class="pg-dm-grid">
          <div class="pg-dm-cell"><div class="pg-dm-l">기준실적</div><div class="pg-dm-v">${Nf(st.base)}원</div></div>
          ${st.hiCap ? `<div class="pg-dm-cell"><div class="pg-dm-l">장기하이캡</div><div class="pg-dm-v">${Nf(st.hiCap)}원</div></div>` : ""}
          <div class="pg-dm-cell"><div class="pg-dm-l">현재실적</div><div class="pg-dm-v">${Nf(st.current)}원</div></div>
          <div class="pg-dm-cell"><div class="pg-dm-l">달성률</div><div class="pg-dm-v" style="color:${rateC};">${st.rate.toFixed(1)}%</div></div>
          <div class="pg-dm-cell"><div class="pg-dm-l">순증</div><div class="pg-dm-v" style="color:${netC};">${st.net >= 0 ? "+" : ""}${Nf(st.net)}원</div></div>
          <div class="pg-dm-cell pg-dm-wide"><div class="pg-dm-l">🏆 예상 시상</div><div class="pg-dm-v" style="color:var(--orange);">${awdText}</div></div>
          ${st.ipumAmt ? `<div class="pg-dm-cell pg-dm-wide"><div class="pg-dm-l">✨ 인품 (신상품)</div><div class="pg-dm-v">${st.ipumCount}건 · ${Nf(st.ipumAmt)}원</div></div>` : ""}
        </div>
      `,
      closeLabel: pushStack ? "← TOP10 으로" : "← 돌아가기",
      pushStack: !!pushStack
    });
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
        <div class="modal-backdrop" data-pg-close></div>
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
            <button class="btn-primary" id="pg-full-modal-close" data-pg-close>닫기</button>
          </div>
        </div>
      `;
      document.body.appendChild(modal);

      // 닫기 요소(백드롭, ×, 하단 버튼)
      modal.querySelectorAll("[data-pg-close]").forEach((el) => {
        el.addEventListener("click", (e) => { e.stopPropagation(); closePgFullModal(); });
      });

      // 패널 내부의 빈 공간(인터랙티브 요소 제외) 탭 시에도 닫기
      modal.querySelector(".modal-panel").addEventListener("click", (e) => {
        if (e.target.closest(".pg-tr-click, .pg-pcard-row, .modal-close, button, a, input, textarea, select, .pg-rb, .pg-pcard-prize, .pg-pcard-chev")) return;
        closePgFullModal();
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
    modal.hidden = false;
    // 모달 바디의 이름 클릭 → 교육생 상세 팝업(스택 push)
    modal.querySelectorAll(".pg-tr-click, .pg-pcard-row[data-emp]").forEach((el) => {
      el.addEventListener("click", (e) => {
        e.stopPropagation();
        const emp = el.dataset.emp;
        if (emp) openProgressStudentPopup(emp, true /* pushStack */);
      });
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
      // 복원된 바디의 이름 클릭 재바인딩
      modal.querySelectorAll(".pg-tr-click, .pg-pcard-row[data-emp]").forEach((el) => {
        el.addEventListener("click", (ev) => {
          ev.stopPropagation();
          const emp = el.dataset.emp;
          if (emp) openProgressStudentPopup(emp, true);
        });
      });
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
          <td class="r">${Nf(r.pgBase)}</td>
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
        const records = newRecords.map((r, i) => ({
          ...r,
          name:   (modal.querySelector(`.pgn-name[data-idx="${i}"]`)?.value.trim()  || r.name),
          cohort: (modal.querySelector(`.pgn-cohort[data-idx="${i}"]`)?.value.trim() || ""),
          phone:  (modal.querySelector(`.pgn-phone[data-idx="${i}"]`)?.value.trim()  || ""),
          target: 0, honors: 0
        }));
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
    const elig = stats.filter((s) => s.net >= PROGRESS_AWARDS.minNetForRank).length;
    const a5 = stats.filter((s) => s.net >= 500000).length;
    const a4 = stats.filter((s) => s.net >= 300000 && s.net < 500000).length;
    const cash = stats.filter((s) => s.net >= 50000 && s.net < 300000).length;
    const exc = total - elig - over5 + a5 + a4; // 제외: 순증 5만 미만
    const exclude = stats.filter((s) => s.net < 50000).length;
    const today = new Date();
    const baseDate = today.getFullYear() + "-" + String(today.getMonth() + 1).padStart(2, "0") + "-" + String(today.getDate()).padStart(2, "0");
    const firstCohort = rows[0]?.cohort || "";
    const cohortTitle = firstCohort ? `${firstCohort} ` : "";

    return `
      <div class="pg-wrap">
        <!-- [1] 과정 기본 정보 카드 -->
        <div class="pg-info-grid">
          <div class="pg-info-card">
            <h5>📘 과정 기본 정보</h5>
            <dl>
              <dt>과정명</dt><dd>${escapeHtml(cohortTitle)}고객마스터</dd>
              <dt>지역단</dt><dd>${escapeHtml(state.progressRegion)}</dd>
              <dt>기준일</dt><dd>${baseDate}</dd>
              <dt>총 인원</dt><dd><strong>${total}명</strong></dd>
            </dl>
          </div>
          <!-- [2] 달성 현황 카드 -->
          <div class="pg-info-card">
            <h5>🎯 달성 현황</h5>
            <dl>
              <dt>순증 5만원↑</dt><dd class="ok"><strong>${over5}명</strong></dd>
              <dt>80~100%</dt><dd>${mid80}명</dd>
              <dt>50~80%</dt><dd>${mid50}명</dd>
              <dt>50% 미만</dt><dd class="warn">${low50}명</dd>
              <dt>평균 달성률</dt><dd><strong>${avgR.toFixed(1)}%</strong></dd>
            </dl>
          </div>
          <!-- [3] 시상 현황 카드 -->
          <div class="pg-info-card">
            <h5>🏆 시상 현황</h5>
            <dl>
              <dt>순증 30만원↑</dt><dd class="ok"><strong>${elig}명</strong></dd>
              <dt>150% 지급</dt><dd>${a5}명</dd>
              <dt>120% 지급</dt><dd>${a4}명</dd>
              <dt>현금 시상</dt><dd>${cash}명</dd>
              <dt>제외</dt><dd class="warn">${exclude}명</dd>
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
              <div class="pg-paste-desc">지역단·비전센터·지점·사원번호·성명·차월·육성리더·직전6개월인보험·직전6개월환산·직전6개월육성소득·기준실적·현재실적·인품건수·인품실적 (탭 구분, 금액단위: 원)</div>
              <textarea id="pg-progress-paste" rows="6" placeholder="엑셀에서 복사 후 붙여넣기"></textarea>
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
                const base  = s.pgBase !== undefined ? Number(s.pgBase) : Number(s.base || 0);  // 원
                const hiCap = Number(s.hiCap   || 0);
                const cur   = Number(s.current  || 0);         // 원
                const net   = cur - base;                       // 원
                const rate  = base > 0 ? (cur / base) * 100 : 0;
                return `<tr>
                  <td>${i + 1}</td>
                  <td>${escapeHtml(s.branch || "")}</td>
                  <td><strong>${escapeHtml(s.name || "")}</strong></td>
                  <td class="r">${Nf(base)}</td>
                  <td><input type="number" class="pg-input" data-emp="${escapeHtml(s.empNo)}" data-f="hiCap" value="${hiCap}" min="0" step="1"></td>
                  <td><input type="number" class="pg-input" data-emp="${escapeHtml(s.empNo)}" data-f="current" value="${cur}" min="0" step="1"></td>
                  <td class="r" data-calc="rate-${escapeHtml(s.empNo)}">${rate.toFixed(1)}%</td>
                  <td class="r" data-calc="net-${escapeHtml(s.empNo)}">${net >= 0 ? "+" : ""}${Nf(net)}</td>
                  <td><input type="number" class="pg-input" data-emp="${escapeHtml(s.empNo)}" data-f="ipumCount" value="${Number(s.ipumCount || 0)}" min="0"></td>
                  <td><input type="number" class="pg-input" data-emp="${escapeHtml(s.empNo)}" data-f="ipumAmt" value="${Number(s.ipumAmt || 0)}" min="0"></td>
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
              <strong>📋 인품 붙여넣기 — "이름 인품건 인품실적" (탭/공백 구분)</strong>
              <textarea id="pg-ipum-paste" rows="4" placeholder="예)
정경화  3  450000
박희자  2  300000"></textarea>
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
            <div class="pg-tbl-wrap"><table class="pg-tbl pg-admin-tbl">
              <thead><tr>
                <th>#</th><th>지점</th><th>성명</th>
                <th>기준실적</th><th>현재실적</th><th>팀</th>
              </tr></thead>
              <tbody>${rows.map((s, i) => {
                const base = Number(s.base || 0);
                const cur = Number(s.current || 0);
                return `<tr>
                  <td>${i + 1}</td>
                  <td>${escapeHtml(s.branch || "")}</td>
                  <td><strong>${escapeHtml(s.name || "")}</strong></td>
                  <td class="r">${Nf(base)}</td>
                  <td class="r">${Nf(cur)}</td>
                  <td><input type="text" class="pg-input pg-team-input" data-emp="${escapeHtml(s.empNo)}" value="${escapeHtml(s.team || "")}" placeholder="예) 1팀"></td>
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
  function computeTeamSummary(list) {
    const byTeam = {};
    list.forEach((s) => {
      const inp = document.querySelector(`#progress-body .pg-team-input[data-emp="${s.empNo}"]`);
      const team = (inp ? inp.value : s.team || "").toString().trim();
      if (!team) return;
      if (!byTeam[team]) byTeam[team] = { base: 0, current: 0, members: [] };
      byTeam[team].base += Number(s.base || 0);
      byTeam[team].current += Number(s.current || 0);
      byTeam[team].members.push(s.name || "");
    });
    return byTeam;
  }

  function renderTeamSummary(list) {
    const box = document.getElementById("pg-team-summary");
    if (!box) return;
    const byTeam = computeTeamSummary(list);
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

  function bindProgressAdminEvents(list) {
    // 실시간 계산 (현재실적 변경 시 달성률/순증 재계산)
    document.querySelectorAll("#progress-body .pg-input[data-f='current']").forEach((inp) => {
      inp.addEventListener("input", (e) => {
        const emp = e.target.dataset.emp;
        const s = state.students.find((x) => x.empNo === emp);
        const base = (s?.pgBase !== undefined) ? Number(s.pgBase) : Number(s?.base || 0);   // 원
        const cur  = parseFloat(e.target.value) || 0; // 원
        const net  = cur - base;                       // 원
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
      document.querySelectorAll("#progress-body .pg-input").forEach((inp) => {
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
    document.querySelectorAll("#progress-body .pg-paste-mode-btn").forEach((btn) => {
      btn.addEventListener("click", () => {
        document.querySelectorAll("#progress-body .pg-paste-mode-btn").forEach((b) => b.classList.remove("active"));
        btn.classList.add("active");
        const mode = btn.dataset.mode;
        const monthlyDiv = document.getElementById("pg-paste-mode-monthly");
        const progressDiv = document.getElementById("pg-paste-mode-progress");
        const honorsDiv  = document.getElementById("pg-paste-mode-honors");
        if (monthlyDiv)  monthlyDiv.style.display  = mode === "monthly"  ? "" : "none";
        if (progressDiv) progressDiv.style.display = mode === "progress" ? "" : "none";
        if (honorsDiv)   honorsDiv.style.display   = mode === "honors"   ? "" : "none";
      });
    });

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

    // ── 실적진도현황 붙여넣기 핸들러 ──────────────────────────────
    const progressPasteApply = $("#btn-pg-progress-paste-apply");
    if (progressPasteApply) progressPasteApply.addEventListener("click", async () => {
      const txt = $("#pg-progress-paste").value.trim();
      const m = $("#pg-progress-paste-msg");
      if (!txt) { if (m) { m.textContent = "❌ 붙여넣을 내용이 없습니다."; m.className = "pg-msg err"; } return; }

      // 컬럼: 지역단(0)|비전센터(1)|지점(2)|사원번호(3)|성명(4)|차월(5)|육성리더(6)
      //        직전6개월인보험(7)|직전6개월환산(8)|직전6개월육성소득(9)
      //        기준실적(10)|현재실적(11)|인품건수(12)|인품실적(13)|순증실적(14-이후 무시)
      const parseAmt = (v) => parseInt((v || "").replace(/,/g, "").trim(), 10) || 0;
      const updateRecords = [];
      const newRecords    = [];  // 미매칭: 신규 등록 대상

      txt.split(/\r?\n/).forEach((line) => {
        const p = line.split(/\t/).map((c) => c.replace(/,/g, "").replace(/[ ​﻿]/g, "").trim());
        if (p.length < 12) return;
        const region  = p[0] || "";
        const center  = p[1] || "";
        const branch  = p[2] || "";
        const empNo   = p[3] || "";
        const name    = p[4] || "";
        const pgMonth = p[5] || "";
        const pgLeader= p[6] || "";
        const pgPreIns    = parseAmt(p[7]);
        const pgPreConv   = parseAmt(p[8]);
        const pgPreIncome = parseAmt(p[9]);
        const pgBase      = parseAmt(p[10]);
        const pgCurrent   = parseAmt(p[11]);
        const pgIpumCount = p.length > 12 ? parseAmt(p[12]) : 0;  // 인품건수
        const pgIpumAmt   = p.length > 13 ? parseAmt(p[13]) : 0;  // 인품실적

        if (!empNo) return;
        // 전체 교육생에서 조회 (필터 범위 제한 없이)
        const existing = state.students.find((x) => x.empNo === empNo);
        const pgFields = { pgBase, pgCurrent, pgIpumCount, pgIpumAmt, pgPreIns, pgPreConv, pgPreIncome, pgMonth, pgLeader };

        if (existing) {
          const baseUpdate = pgBase > 0 ? { base: pgBase, current: pgCurrent } : {};
          const targetUpdate = (region !== "호남지역단" && pgBase > 0) ? { target: pgBase + 50000 } : {};
          updateRecords.push({ ...existing, region, center, branch, name, ...pgFields, ...baseUpdate, ...targetUpdate });
        } else {
          const newTarget = region !== "호남지역단" && pgBase > 0 ? pgBase + 50000 : 0;
          newRecords.push({ region, center, branch, empNo, name, ...pgFields, base: pgBase, current: pgCurrent, target: newTarget });
        }
      });

      if (updateRecords.length === 0 && newRecords.length === 0) {
        if (m) { m.textContent = "❌ 파싱된 행이 없습니다. 탭 구분 및 열 수(12+)를 확인하세요."; m.className = "pg-msg err"; }
        return;
      }

      // 기존 학생 업데이트
      if (updateRecords.length > 0) {
        progressPasteApply.disabled = true;
        if (m) { m.textContent = "저장중..."; m.className = "pg-msg"; }
        try {
          if (typeof window.DataAPI.saveMany === "function") {
            await window.DataAPI.saveMany(updateRecords);
          } else {
            for (const r of updateRecords) await window.DataAPI.save(r);
          }
          let msg = `✅ ${updateRecords.length}명 저장 완료`;
          if (newRecords.length) msg += ` · 신규 ${newRecords.length}명 팝업 확인 필요`;
          if (m) { m.textContent = msg; m.className = "pg-msg ok"; setTimeout(() => { if (m) m.textContent = ""; }, 8000); }
          toast(`${updateRecords.length}명 실적진도현황 저장`, "success");
        } catch (e) {
          console.error(e);
          if (m) { m.textContent = "❌ 저장 실패: " + e.message; m.className = "pg-msg err"; }
          toast("저장 실패: " + e.message, "error");
          progressPasteApply.disabled = false;
          return;
        }
        progressPasteApply.disabled = false;
      }

      // 신규 등록 대상 팝업
      if (newRecords.length > 0) {
        openPgNewStudentModal(newRecords);
      }
    });

    // 인품 테이블 ↔ 메인 테이블 동기화 (같은 empNo 의 두 입력을 동시 반영)
    document.querySelectorAll("#progress-body .pg-input[data-f2]").forEach((inp) => {
      inp.addEventListener("input", (e) => {
        const emp = e.target.dataset.emp;
        const f = e.target.dataset.f2; // "ipumCount" or "ipumAmt"
        const twin = document.querySelector(`#progress-body .pg-input[data-emp="${emp}"][data-f="${f}"]`);
        if (twin) twin.value = e.target.value;
      });
    });
    document.querySelectorAll("#progress-body .pg-input[data-f='ipumCount'], #progress-body .pg-input[data-f='ipumAmt']").forEach((inp) => {
      inp.addEventListener("input", (e) => {
        const emp = e.target.dataset.emp;
        const f = e.target.dataset.f;
        const twin = document.querySelector(`#progress-body .pg-input[data-emp="${emp}"][data-f2="${f}"]`);
        if (twin) twin.value = e.target.value;
      });
    });

    // 인품 붙여넣기 반영: "이름 건수 실적"
    const ipumPasteApply = $("#btn-pg-ipum-paste-apply");
    if (ipumPasteApply) ipumPasteApply.addEventListener("click", () => {
      const txt = $("#pg-ipum-paste").value;
      // 학생 이름 사전 정규화 (공백/제로폭 문자 제거) 후 lookup map
      const normName = (n) => (n || "").replace(/\s+/g, "").replace(/[​-‍﻿]/g, "");
      const nameMap = new Map();
      list.forEach((s) => nameMap.set(normName(s.name), s));
      let cnt = 0;
      const unmatched = [];
      txt.split(/\r?\n/).forEach((line) => {
        const parts = line.trim().split(/\s+/);
        if (parts.length < 3) return;
        const rawName = parts[0];
        const count = parseInt(parts[1].replace(/,/g, ""), 10);
        const amt = parseInt(parts[parts.length - 1].replace(/,/g, ""), 10);
        if (isNaN(count) || isNaN(amt)) return;
        const s = nameMap.get(normName(rawName));
        if (s) {
          ["f2", "f"].forEach((attr) => {
            const cEl = document.querySelector(`.pg-input[data-emp="${s.empNo}"][data-${attr}="ipumCount"]`);
            const aEl = document.querySelector(`.pg-input[data-emp="${s.empNo}"][data-${attr}="ipumAmt"]`);
            if (cEl) cEl.value = count;
            if (aEl) aEl.value = amt;
          });
          cnt++;
        } else {
          unmatched.push(rawName);
        }
      });
      const m = $("#pg-ipum-paste-msg");
      if (m) {
        const warnPart = unmatched.length ? ` · ⚠️ 이름 미일치 ${unmatched.length}명 (${unmatched.slice(0,3).join(", ")}${unmatched.length>3?"…":""})` : "";
        m.innerHTML = `✅ ${cnt}명 반영 (아래 [💾 인품 저장] 눌러야 확정)${warnPart}`;
        m.className = unmatched.length ? "pg-msg warn" : "pg-msg ok";
        setTimeout(() => { m.textContent = ""; }, 8000);
      }
      if (cnt === 0) {
        toast("매칭되는 이름이 없습니다. 좌측 지역단과 이름을 확인하세요.", "error");
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
      document.querySelectorAll("#progress-body .pg-input[data-f2]").forEach((inp) => {
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
    renderTeamSummary(list);

    // 팀 입력 실시간 반영
    document.querySelectorAll("#progress-body .pg-team-input").forEach((inp) => {
      inp.addEventListener("input", () => renderTeamSummary(list));
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
        const teamName = `${teamNo}팀`;
        const inp = document.querySelector(`#progress-body .pg-team-input[data-emp="${s.empNo}"]`);
        if (inp) inp.value = teamName;
      });
      renderTeamSummary(list);
      const msg = $("#pg-team-msg");
      if (msg) { msg.textContent = `✅ ${n}개 팀으로 고르게 배정 완료 (저장 버튼을 눌러 확정)`; msg.className = "pg-msg ok"; setTimeout(() => { msg.textContent = ""; }, 4000); }
    });

    // 전체 초기화
    const clrBtn = $("#btn-pg-team-clear");
    if (clrBtn) clrBtn.addEventListener("click", () => {
      if (!confirm("모든 교육생의 팀 배정을 비웁니다. 저장 버튼을 눌러야 반영됩니다. 계속할까요?")) return;
      document.querySelectorAll("#progress-body .pg-team-input").forEach((inp) => { inp.value = ""; });
      renderTeamSummary(list);
    });

    // 팀 저장
    const teamSaveBtn = $("#btn-pg-team-save");
    if (teamSaveBtn) teamSaveBtn.addEventListener("click", async () => {
      const updates = [];
      document.querySelectorAll("#progress-body .pg-team-input").forEach((inp) => {
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
      } catch (e) {
        console.error(e);
        if (msg) { msg.textContent = "❌ 저장 실패: " + e.message; msg.className = "pg-msg err"; }
        toast("저장 실패: " + e.message, "error");
      }
      teamSaveBtn.disabled = false;
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

  // 렌더 디바운스 (연쇄 쓰기 시 트레일링 실행으로 1회만)
  let _renderTimer = 0;
  function renderDebounced() {
    if (_renderTimer) return;
    _renderTimer = setTimeout(() => { _renderTimer = 0; render(); }, 60);
  }

  function render() {
    const list = filteredStudents();
    renderKPIs(list);
    renderSidebarStudentList(list);
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
        <h3>평균실적 상위 ${top10.length}명</h3>
        <table class="stats-table">
          <thead><tr>
            <th>#</th><th>이름</th><th>지점</th><th>기수</th><th class="r">평균실적</th><th class="r">순증목표</th>
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
    const regions = [...new Set(state.students.map((s) => s.region).filter(Boolean))].sort();
    if (!regions.length) {
      container.innerHTML = `<p class="settings-desc" style="color:#999;">등록된 교육생이 없어 지역단 목록을 불러올 수 없습니다.</p>`;
      return;
    }
    let stored = {};
    try { stored = JSON.parse(localStorage.getItem(LS_DEFAULTS_KEY) || "{}"); } catch (e) {}
    const zeroTargetCount = state.students.filter((s) => !s.target || Number(s.target) === 0).length;
    const nonHonamCount = state.students.filter((s) => s.region && s.region !== "호남지역단").length;
    container.innerHTML = `
      <table class="settings-info" style="width:auto;margin-bottom:8px;">
        <thead><tr><th style="min-width:120px;">지역단</th><th>마스터목표 기본값</th></tr></thead>
        <tbody>${regions.map((r) => {
          const val = stored[r] !== undefined ? stored[r] : DEFAULT_MASTER_TARGET;
          return `<tr>
            <td>${escapeHtml(r)}</td>
            <td><input type="number" class="settings-target-input" data-region="${escapeHtml(r)}" value="${val}" min="0" step="1000" style="width:120px;"> 원</td>
          </tr>`;
        }).join("")}</tbody>
      </table>
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
    editingEmpNo = null;
    state.formTgtAddAmount = null;
    syncOrgLabels();
  }

  function openEditForm(empNo) {
    const s = state.students.find((x) => x.empNo === empNo);
    if (!s) return;
    state.form = { region: s.region || "", center: s.center || "", branch: s.branch || "" };
    $("#form-empno").value = s.empNo;
    $("#form-name").value = s.name || "";
    $("#form-phone").value = s.phone || "";
    $("#form-base").value   = s.base   || "";
    const computedTarget = (s.region !== "호남지역단" && !Number(s.target))
      ? getProgressStat(s).base + 50000
      : Number(s.target) || "";
    $("#form-target").value = computedTarget;
    $("#form-honors").value = s.honors || "";
    $("#form-cohort").value = s.cohort || "";
    editingEmpNo = s.empNo;
    state.formTgtAddAmount = null;
    const ft = $("#form-target"); if (ft) ft.removeAttribute("readonly");
    syncOrgLabels();
    switchTab("single");
    openModal("#modal-add");
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
      honors: Number($("#form-honors").value || 0)
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
      ["평균실적", "base"], ["마스터목표", "target"], ["아너스목표", "honors"]
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

  async function saveBulk() {
    const raw = $("#form-bulk").value.trim();
    if (!raw) {
      toast("붙여넣을 데이터가 없습니다.", "error");
      setBulkProgress("붙여넣을 데이터가 없습니다.", "error");
      return { ok: 0, fail: 0, total: 0 };
    }

    // 1. 파싱
    const lines = raw.split(/\r?\n/).filter((l) => l.trim());
    const records = [];
    const parseErrors = [];
    lines.forEach((line, i) => {
      const cols = line.includes("\t") ? line.split("\t") : line.split(",");
      const [region, center, branch, cohort, empNo, name, phone, base, target, honors, tenureMonths] =
        cols.map((c) => (c || "").trim().replace(/^"(.*)"$/, "$1"));
      // 헤더 행 자동 스킵
      if (region === "지역단" && center === "비전센터" && branch === "지점") return;
      if (!empNo) { parseErrors.push(`${i + 1}행 사번 누락`); return; }
      const rec = {
        region, center, branch, cohort,
        empNo: empNo.replace(/[\s\/\\]/g, ""),
        name, phone,
        base: toNum(base), target: toNum(target), honors: toNum(honors)
      };
      if (tenureMonths) rec.tenureMonths = toNum(tenureMonths);
      records.push(rec);
    });

    if (records.length === 0) {
      const msg = `저장할 행이 없습니다. (${parseErrors[0] || "데이터 없음"})`;
      toast(msg, "error");
      setBulkProgress(msg, "error");
      return { ok: 0, fail: 0, total: 0 };
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
    const headers = ["지역단","비전센터","지점","기수","사번","이름","연락처","평균실적","마스터목표","아너스목표"];
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
      state.filter = { region: DEFAULT_REGION, center: "", branch: "", cohort: "", q: "" };
      $("#filter-cohort").value = "";
      $("#search-box-side").value = "";
      syncOrgLabels();
      persistFilter();
      render();
    });
    $("#filter-cohort").addEventListener("change", (e) => {
      state.filter.cohort = e.target.value;
      persistFilter();
      render();
    });
    $("#search-box-side").addEventListener("input", (e) => {
      state.filter.q = e.target.value;
      render();
    });

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

    // 탭 전환
    $$(".tab").forEach((t) => t.addEventListener("click", () => switchTab(t.dataset.tab)));

    // 시드 파일 불러오기
    $("#btn-load-seed").addEventListener("click", openSeedPicker);

    // 저장
    $("#btn-save").addEventListener("click", async () => {
      const activeTab = document.querySelector(".tab.active").dataset.tab;
      try {
        if (activeTab === "single") {
          const ok = await saveSingle();
          if (ok) {
            closeModal("#modal-add");
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

    // 실적진도 서브탭 (지역단은 좌측 필터에서 선택)
    document.querySelectorAll("#progress-panel .sub-tab").forEach((btn) => {
      btn.addEventListener("click", () => {
        state.progressSubTab = btn.dataset.psub;
        document.querySelectorAll("#progress-panel .sub-tab").forEach((b) => b.classList.toggle("active", b.dataset.psub === state.progressSubTab));
        renderProgressPanel();
      });
    });

    // 설정 탭 / 푸터 / 헤더 — 앱 버전 (커밋마다 +0.01)
    const v = $("#app-version"); if (v) v.textContent = `v${APP_VERSION} (build 20260425m)`;
    const fv = $("#app-footer-ver"); if (fv) fv.textContent = APP_VERSION;
    const hv = $("#app-header-ver"); if (hv) hv.textContent = APP_VERSION;
    $("#btn-export-json").addEventListener("click", () => exportJSON(filteredStudents(), "filtered"));
    $("#btn-export-json-all").addEventListener("click", () => exportJSON(state.students, "all"));
    $("#btn-import-json").addEventListener("click", () => $("#file-import-json").click());
    $("#file-import-json").addEventListener("change", onImportJSONFile);
    $("#btn-delete-filtered").addEventListener("click", onDeleteFiltered);
    const openSdBtn = $("#btn-open-student-delete");
    if (openSdBtn) openSdBtn.addEventListener("click", openStudentDeleteModal);
    const sdSearch = $("#sd-search");
    if (sdSearch) sdSearch.addEventListener("input", (e) => renderStudentDeleteList(e.target.value));
    const sdClear = $("#btn-sd-clear");
    if (sdClear) sdClear.addEventListener("click", () => {
      state.sdSelected = new Set();
      renderStudentDeleteList($("#sd-search").value);
      updateSdCounts();
    });
    const sdDelete = $("#btn-sd-delete");
    if (sdDelete) sdDelete.addEventListener("click", doStudentsDeleteFromModal);
  }

  // ========== 설정 ==========
  function exportJSON(list, scope) {
    if (!list.length) { toast("내보낼 데이터가 없습니다.", "error"); return; }
    const payload = {
      exportedAt: new Date().toISOString(),
      scope,
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
    a.download = `students_${scope}_${new Date().toISOString().slice(0,10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
    toast(`${list.length}건 JSON 내보내기 완료.`, "success");
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
      if (!confirm(`${arr.length}건의 교육생 데이터를 복원합니다. 동일 사번은 merge 됩니다. 진행할까요?`)) return;
      if (typeof window.DataAPI.saveMany === "function") {
        const { committed, errors } = await window.DataAPI.saveMany(arr);
        toast(`${committed}건 복원 완료${errors.length ? ` / ${errors.length}건 실패` : ""}`, errors.length ? "error" : "success");
      } else {
        let ok = 0, fail = 0;
        for (const s of arr) {
          try { await window.DataAPI.save(s); ok++; } catch (err) { fail++; }
        }
        toast(`${ok}건 복원, ${fail}건 실패`, fail ? "error" : "success");
      }
    } catch (err) {
      console.error(err);
      toast("복원 실패: " + err.message, "error");
    }
  }

  async function onDeleteFiltered() {
    const list = filteredStudents();
    await confirmAndDeleteStudents(list, {
      label: `현재 필터(${[state.filter.region, state.filter.center, state.filter.branch, state.filter.cohort].filter(Boolean).join(" · ") || "전체"})`
    });
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
    document.body.classList.add("mobile-sidebar-open");
    const bd = document.getElementById("mobile-sidebar-backdrop");
    if (bd) bd.hidden = false;
    const hb = document.getElementById("mbn-home-btn");
    if (hb) hb.classList.add("mbn-sidebar-open");
  }
  function closeMobileSidebar() {
    state.mobileSidebarOpen = false;
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
    // 관리자(설정) 진입 시 암호 체크 (세션 1회 인증으로 캐시)
    if (target === "settings" && !sessionStorage.getItem("adminAuth")) {
      const pwd = prompt("관리자 암호를 입력하세요:");
      if (pwd !== "2051") {
        toast("암호가 일치하지 않습니다.", "error");
        return;
      }
      sessionStorage.setItem("adminAuth", "1");
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

  function init() {
    bindEvents();
    // 기본 view = 교육생 관리
    switchView("#students");
    // localStorage에서 복원된 필터값을 UI에 반영
    $("#filter-cohort").value = state.filter.cohort || "";
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
      renderDebounced();
      prefetchConsultCountsOnce();
    });
  }

  // 구 단위 마이그레이션 — 원 단위 저장으로 전환됨에 따라 비활성화
  async function migrateStudentBaseValuesIfNeeded() { /* no-op: v0.92+ stores all values in 원 */ }

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

  document.addEventListener("DOMContentLoaded", init);
})();
