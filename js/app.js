/* 고객컨설팅 마스터과정 운영관리 - 메인 앱 로직 */

(function () {
  const LS_KEY = "cmf.filter.v1";
  const DEFAULT_REGION = "호남지역단";

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
    orgPickerTarget: null, // 'filter-region' | 'form-region' 등
    selectedEmpNo: null,
    consultations: [],
    consultUnsub: null,
    // Phase 1 interview form
    tgtAutoMode: true,       // ins→tgt 자동 계산 on/off
    editingConsultId: null,  // Phase 4용 예약 (이번엔 사용 안 함)
    lastDetailEmpNo: null,   // 마지막으로 완전 렌더한 교육생 (폼 보존용)
    // Phase 2 clients
    crData: [],              // 현재 폼의 상담고객 배열 (최대 5)
    // Phase 3 시상 계산기
    calcOpen: false,         // 계산기 접힘/펼침 상태
    calcTgtUserEditing: false, // 희망목표 직접입력 중 플래그
    // 동기화 상태
    studentsLoaded: false,   // Firestore 첫 응답 여부
    syncMeta: { fromCache: false },
    // 교육생 패널 서브뷰 (form | history | print)
    studentSubView: "form",
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

  // ========== 조직 선택 팝업 ==========
  function openOrgPicker(target) {
    state.orgPickerTarget = target;
    const titleMap = {
      "filter-region": "지역단 선택", "filter-center": "비전센터 선택", "filter-branch": "지점 선택",
      "form-region": "지역단 선택", "form-center": "비전센터 선택", "form-branch": "지점 선택"
    };
    $("#modal-org-title").textContent = titleMap[target] || "선택";
    renderOrgList("");
    $("#org-search").value = "";
    openModal("#modal-org");
    setTimeout(() => $("#org-search").focus(), 50);
  }

  function getOrgOptions(target) {
    const isFilter = target.startsWith("filter-");
    const scope = isFilter ? state.filter : state.form;
    const data = window.ORG_DATA;
    if (target.endsWith("region")) {
      return data.regions.map((r) => r.name);
    }
    if (target.endsWith("center")) {
      if (!scope.region) return [];
      const reg = data.regions.find((r) => r.name === scope.region);
      return reg ? reg.centers.map((c) => c.name) : [];
    }
    if (target.endsWith("branch")) {
      if (!scope.region || !scope.center) return [];
      const reg = data.regions.find((r) => r.name === scope.region);
      const ctr = reg && reg.centers.find((c) => c.name === scope.center);
      return ctr ? ctr.branches : [];
    }
    return [];
  }

  function renderOrgList(q) {
    const list = $("#org-list");
    const target = state.orgPickerTarget;
    if (!target) return;
    const opts = getOrgOptions(target);
    list.innerHTML = "";

    if (target.endsWith("center") && !(target.startsWith("filter") ? state.filter.region : state.form.region)) {
      list.innerHTML = `<li class="disabled">먼저 지역단을 선택하세요.</li>`;
      return;
    }
    if (target.endsWith("branch")) {
      const scope = target.startsWith("filter") ? state.filter : state.form;
      if (!scope.region || !scope.center) {
        list.innerHTML = `<li class="disabled">먼저 지역단/비전센터를 선택하세요.</li>`;
        return;
      }
    }

    const filtered = opts.filter((o) => o.toLowerCase().includes((q || "").toLowerCase()));
    // 지역단은 필수 선택이므로 '전체' 옵션 없음. 비전센터/지점만 '전체' 허용.
    if (target.startsWith("filter-") && !target.endsWith("region")) {
      const li = document.createElement("li");
      li.textContent = "전체";
      li.addEventListener("click", () => selectOrg(""));
      list.appendChild(li);
    }
    filtered.forEach((name) => {
      const li = document.createElement("li");
      li.textContent = name;
      li.addEventListener("click", () => selectOrg(name));
      list.appendChild(li);
    });
    if (filtered.length === 0 && opts.length > 0) {
      list.innerHTML += `<li class="disabled">검색 결과가 없습니다.</li>`;
    }
  }

  function selectOrg(name) {
    const target = state.orgPickerTarget;
    const isFilter = target.startsWith("filter-");
    const field = target.split("-")[1]; // region | center | branch
    const scope = isFilter ? state.filter : state.form;

    // 필터의 지역단은 비울 수 없음 (필수 선택)
    if (isFilter && field === "region" && !name) {
      closeModal("#modal-org");
      return;
    }

    scope[field] = name;
    // 상위 변경 시 하위 초기화
    if (field === "region") { scope.center = ""; scope.branch = ""; }
    if (field === "center") { scope.branch = ""; }

    syncOrgLabels();
    closeModal("#modal-org");
    if (isFilter) { persistFilter(); render(); }
  }

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
    const groups = {};
    list.forEach((s) => {
      const key = s.branch || "(지점 미지정)";
      if (!groups[key]) groups[key] = [];
      groups[key].push(s);
    });

    container.innerHTML = Object.keys(groups).sort().map((branch) => {
      const rows = groups[branch].slice().sort((a, b) => (a.name || "").localeCompare(b.name || ""));
      return `
        <div class="branch-mini">
          <div class="branch-mini-head">
            <span class="branch-name">${escapeHtml(branch)}</span>
            <span class="branch-cnt">${rows.length}</span>
          </div>
          <ul class="student-mini-list">
            ${rows.map((s) => {
              const nm = s.name || "(이름 미입력)";
              const initial = (s.name || "?").trim().charAt(0) || "?";
              return `
              <li class="${state.selectedEmpNo === s.empNo ? "selected" : ""}" data-emp="${escapeHtml(s.empNo)}" data-initial="${escapeHtml(initial)}">
                <span class="s-name-wrap">
                  <span class="s-name">${escapeHtml(nm)}</span>
                  <span class="s-phone">${escapeHtml(s.phone || "")}</span>
                </span>
              </li>
            `;}).join("")}
          </ul>
        </div>
      `;
    }).join("");

    container.querySelectorAll("li[data-emp]").forEach((li) => {
      li.addEventListener("click", () => selectStudent(li.dataset.emp));
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
    // 면담 기록 실시간 구독
    if (window.DataAPI && typeof window.DataAPI.subscribeConsultations === "function") {
      state.consultUnsub = window.DataAPI.subscribeConsultations(empNo, (list) => {
        state.consultations = list;
        renderConsultations();
        // 차수/ins/tgt 자동채움 재실행 (빈 필드만 보정)
        const s = state.students.find((x) => x.empNo === empNo);
        if (s) autoFillInterviewForm(s);
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
          <div><span>평균실적</span><strong>${fmt(s.base)}</strong></div>
          <div><span>순증목표</span><strong>${fmt(s.honors)}</strong></div>
          <div><span>인보험평균</span><strong>${fmt(s.insAvg)}</strong></div>
        </div>
        <div class="sib-actions">
          <button class="btn-outline small" id="btn-detail-edit">정보 수정</button>
          <button class="btn-outline small danger" id="btn-detail-del">삭제</button>
        </div>
      </div>
    `;
  }

  // 같은 학생의 기본정보만 부분 갱신 (폼 입력 보존)
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
            <label>비젼센터</label>
            <input type="text" id="iv-vc" value="${escapeHtml(s.center || "")}" readonly>
          </div>
          <div class="iv-field">
            <label>지점</label>
            <input type="text" id="iv-branch" value="${escapeHtml(s.branch || "")}" readonly>
          </div>
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
            <label>인보험 평균 <span class="iv-hint" id="iv-ins-hint"></span></label>
            <input type="number" id="iv-ins" placeholder="천원" step="10">
          </div>

          <div class="iv-field">
            <label>당월목표 <span class="iv-hint" id="iv-tgt-hint">평균+20만원 자동</span></label>
            <input type="number" id="iv-tgt" placeholder="천원" step="10">
          </div>
          <div class="iv-field">
            <label>현재실적</label>
            <input type="number" id="iv-curAct" placeholder="천원" step="10">
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
            <input type="number" id="iv-exp" placeholder="천원" step="10">
          </div>
        </div>

        <div class="iv-clients">
          <div class="iv-clients-head">
            <span class="iv-clients-title">주간 활동 점검 — 상담고객</span>
            <button type="button" class="btn-outline small" id="btn-cr-add">+ 고객 추가 (최대 5)</button>
          </div>
          <div id="cr-rows"></div>
        </div>

        <div class="iv-field iv-coach">
          <label>핵심 코칭포인트 / 후속조치 / 다음주 계획 <em>*</em></label>
          <textarea id="iv-coach" rows="5" placeholder="핵심 코칭포인트, 후속조치, 다음주 계획을 상세히 기록하세요"></textarea>
        </div>

        <div class="iv-calc">
          <div class="iv-calc-head" id="btn-calc-toggle">
            <span class="iv-calc-title">📊 시상 계산기 — '26년 2분기 매출아너스</span>
            <span class="iv-calc-icon" id="calc-toggle-icon">▾</span>
          </div>
          <div class="iv-calc-body" id="calc-section" style="display:none">
            <div class="iv-field">
              <label>✍️ 면담자 의견 <span class="iv-hint">저장 시 이력에 함께 보관</span></label>
              <textarea id="calc-comment" rows="2" placeholder="면담자 의견을 입력하세요 (저장 시 이력에 포함)"></textarea>
            </div>
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
          <button class="btn-outline" id="btn-iv-clear">초기화</button>
          <button class="btn-primary" id="btn-iv-save">💾 저장</button>
        </div>
      </div>
    `;
  }

  function bindInterviewFormEvents() {
    $("#iv-seq").addEventListener("input", updateIvTitle);
    $("#iv-ins").addEventListener("input", onIvInsInput);
    $("#iv-tgt").addEventListener("input", onIvTgtInput);
    $("#iv-curAct").addEventListener("input", calcIvPct);
    $("#btn-iv-clear").addEventListener("click", () => {
      const s = state.students.find((x) => x.empNo === state.selectedEmpNo);
      state.editingConsultId = null;
      clearInterviewForm();
      if (s) autoFillInterviewForm(s);
      renderConsultations();
      updateSaveButtonLabel();
    });
    $("#btn-iv-save").addEventListener("click", saveInterview);
    $("#btn-iv-cancel-edit").addEventListener("click", cancelEditInterview);
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
    // 계산기 접힘 상태 복원
    if (state.calcOpen) {
      $("#calc-section").style.display = "block";
      $("#calc-toggle-icon").style.transform = "rotate(0deg)";
    }
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
      if (s?.base  && avgEl && !avgEl.value)       avgEl.value = Math.round(Number(s.base)).toLocaleString();
      if (s?.honors && baseTgtEl && !baseTgtEl.value) baseTgtEl.value = Math.round(Number(s.honors)).toLocaleString();
      if (tgtEl && !tgtEl.value) {
        const fTgt = parseFloat($("#iv-tgt")?.value) || 0;
        if (fTgt) tgtEl.value = Math.round(fTgt * 1000).toLocaleString();
        else if (s?.honors) tgtEl.value = Math.round(Number(s.honors)).toLocaleString();
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
    const avgEl = $("#calc-avg");
    const insEl = $("#iv-ins");
    if (!avgEl || !insEl) return;
    const raw = parseFloat((avgEl.value || "").replace(/,/g, "")) || 0;
    if (raw > 0) {
      insEl.value = Math.round(raw / 1000); // 원 → 천원
      const hint = $("#iv-ins-hint");
      if (hint) hint.textContent = "▲ 계산기에서 입력";
      // tgt 재계산
      if (state.tgtAutoMode) {
        const tgtEl = $("#iv-tgt");
        if (tgtEl) {
          tgtEl.value = Math.round(raw / 1000) + 200;
          const th = $("#iv-tgt-hint"); if (th) th.textContent = "▲ 자동";
        }
      }
    }
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

    // ③ 마스터과정 개인시상 (iv-ins 천원 → 원)
    const insRaw3 = (parseFloat($("#iv-ins")?.value || "0") || 0) * 1000;
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
      return `<tr class="${cls}">
        <td class="c">${icon}</td>
        <td>${escapeHtml(h.grade)}</td>
        <td class="crit">${h.criteria}</td>
        <td class="prize">${(h.prize * 10000).toLocaleString()}원</td>
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
        <div class="rs-applied">→ 현재 등급: <strong>${escapeHtml(award1Grade)}</strong> · 시상금: <strong>${award1 ? fmtW(award1) : "해당없음"}</strong></div>

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

  function onIvInsInput() {
    const insEl = $("#iv-ins");
    const tgtEl = $("#iv-tgt");
    if (!insEl || !tgtEl) return;
    const insVal = parseFloat(insEl.value) || 0;
    if (state.tgtAutoMode && insVal > 0) {
      tgtEl.value = insVal + 200;
      $("#iv-tgt-hint").textContent = "▲ 자동";
    }
    calcIvPct();
  }

  function onIvTgtInput() {
    state.tgtAutoMode = false;
    $("#iv-tgt-hint").textContent = "✏️ 수동";
    calcIvPct();
  }

  function calcIvPct() {
    const curAct = parseFloat($("#iv-curAct")?.value) || 0;
    const tgt = parseFloat($("#iv-tgt")?.value) || 0;
    if (!curAct || !tgt) return;
    const pct = Math.round((curAct / tgt) * 100);
    $("#iv-pct").value = pct;
    $("#iv-pct-hint").textContent = `▲ 자동 (${curAct}/${tgt})`;
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

    // 인보험 평균 우선순위: student.insAvg → 최근 면담의 ins → student.base/1000
    const insEl = $("#iv-ins");
    const tgtEl = $("#iv-tgt");
    const hintIns = $("#iv-ins-hint");
    const hintTgt = $("#iv-tgt-hint");
    if (!insEl || !tgtEl) return;

    const lastWithIns = state.consultations.find((c) => c.ins);
    state.tgtAutoMode = true;

    if (!insEl.value) {
      if (Number(s.insAvg) > 0) {
        insEl.value = s.insAvg;
        if (hintIns) hintIns.textContent = "▲ 기본 인보험 평균";
      } else if (lastWithIns) {
        insEl.value = lastWithIns.ins;
        if (hintIns) hintIns.textContent = `▲ ${lastWithIns.seq || ""}차 면담값`;
      } else if (Number(s.base) > 0) {
        insEl.value = Math.round(Number(s.base) / 1000);
        if (hintIns) hintIns.textContent = "▲ 평균실적에서 변환";
      }
    }

    if (!tgtEl.value) {
      // 최근 면담에 tgt 값이 있고 ins+200 이 아니면 수동값 → 복원하되 mode 수동
      if (lastWithIns && lastWithIns.tgt) {
        const insV = Number(lastWithIns.ins) || 0;
        const tgtV = Number(lastWithIns.tgt) || 0;
        tgtEl.value = tgtV;
        if (tgtV !== insV + 200) {
          state.tgtAutoMode = false;
          if (hintTgt) hintTgt.textContent = "✏️ 수동";
        } else {
          if (hintTgt) hintTgt.textContent = "▲ 자동";
        }
      } else {
        const insV = parseFloat(insEl.value) || 0;
        if (insV > 0) {
          tgtEl.value = insV + 200;
          if (hintTgt) hintTgt.textContent = "▲ 자동";
        }
      }
    }

    // 현재실적 복원
    const curActEl = $("#iv-curAct");
    if (curActEl && !curActEl.value && Number(s.curAct) > 0) {
      curActEl.value = s.curAct;
    }

    calcIvPct();

    // 시상 계산기: 직전 consultation 의 calc 값 → 없으면 student.base/honors prefill
    const avgEl = $("#calc-avg");
    const baseTgtEl = $("#calc-base-tgt");
    const tgtCalcEl = $("#calc-tgt");
    const commentEl = $("#calc-comment");
    const lastCalc = state.consultations.find(
      (c) => c.calcAvg || c.calcBaseTgt || c.calcTgt || c.calcComment
    );
    if (lastCalc) {
      if (avgEl && !avgEl.value && lastCalc.calcAvg) avgEl.value = lastCalc.calcAvg;
      if (baseTgtEl && !baseTgtEl.value && lastCalc.calcBaseTgt) baseTgtEl.value = lastCalc.calcBaseTgt;
      if (tgtCalcEl && !tgtCalcEl.value && lastCalc.calcTgt) {
        const raw = parseFloat(String(lastCalc.calcTgt).replace(/,/g, "")) || 0;
        const fixed = (raw > 0 && raw < 1000) ? raw * 1000 : raw;
        tgtCalcEl.value = fixed ? Math.round(fixed).toLocaleString() : lastCalc.calcTgt;
      }
      if (commentEl && !commentEl.value && lastCalc.calcComment) commentEl.value = lastCalc.calcComment;
    } else {
      if (avgEl && !avgEl.value && Number(s.base) > 0) avgEl.value = Math.round(Number(s.base)).toLocaleString();
      if (baseTgtEl && !baseTgtEl.value && Number(s.honors) > 0) baseTgtEl.value = Math.round(Number(s.honors)).toLocaleString();
    }
    if (state.calcOpen) calc();
  }

  function clearInterviewForm() {
    ["iv-seq","iv-ins","iv-tgt","iv-curAct","iv-pct","iv-plan","iv-hap","iv-exp","iv-coach"]
      .forEach((id) => { const el = $("#" + id); if (el) el.value = ""; });
    const today = new Date().toISOString().slice(0, 10);
    const d = $("#iv-date"); if (d) d.value = today;
    const ihint = $("#iv-ins-hint"); if (ihint) ihint.textContent = "";
    const thint = $("#iv-tgt-hint"); if (thint) thint.textContent = "평균+20만원 자동";
    const phint = $("#iv-pct-hint"); if (phint) phint.textContent = "";
    state.tgtAutoMode = true;
    initCR([]); // 상담고객 리셋
    // 시상 계산기 필드 리셋
    ["calc-avg","calc-base-tgt","calc-tgt","calc-comment"].forEach((id) => {
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
      ins: num("iv-ins"),
      tgt: num("iv-tgt"),
      pct: num("iv-pct"),
      curAct: num("iv-curAct"),
      plan: num("iv-plan"),
      hap: num("iv-hap"),
      exp: num("iv-exp"),
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
      calcTgt: read("calc-tgt"),
      calcComment: read("calc-comment")
    };
  }

  async function saveInterview() {
    const empNo = state.selectedEmpNo;
    if (!empNo) return;
    const rec = buildInterviewRecord();
    if (!rec.date) { toast("면담일시를 입력하세요.", "error"); return; }
    const coachEl = $("#iv-coach");
    if (!rec.coach) {
      toast("핵심 코칭포인트는 필수입니다.", "error");
      if (coachEl) {
        coachEl.closest(".iv-field")?.classList.add("error");
        coachEl.scrollIntoView({ behavior: "smooth", block: "center" });
        setTimeout(() => coachEl.focus(), 300);
        coachEl.addEventListener("input", () => {
          coachEl.closest(".iv-field")?.classList.remove("error");
        }, { once: true });
      }
      return;
    }
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
      if (c.ins)    badges.push(`<span class="cs-badge">인보험 ${fmt(c.ins)}</span>`);
      if (c.tgt)    badges.push(`<span class="cs-badge">당월 ${fmt(c.tgt)}</span>`);
      if (c.curAct) badges.push(`<span class="cs-badge">현재 ${fmt(c.curAct)}</span>`);
      if (c.pct)    badges.push(`<span class="cs-badge blue">진도 ${fmt(c.pct)}%</span>`);
      if (c.plan)   badges.push(`<span class="cs-badge">가입 ${fmt(c.plan)}</span>`);
      if (c.hap)    badges.push(`<span class="cs-badge">행복 ${fmt(c.hap)}</span>`);
      if (c.exp)    badges.push(`<span class="cs-badge">예상 ${fmt(c.exp)}</span>`);
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
                <td>${ci + 1}</td>
                <td class="cc-name-td">${escapeHtml(c.name || "")}</td>
                <td>${escapeHtml((c.types || []).join(", ")) || "-"}</td>
                <td>${escapeHtml((c.consult || []).join(", ")) || "-"}</td>
                <td>${escapeHtml((c.material || []).join(", ")) || "-"}</td>
                <td class="hi-ct-amt">${escapeHtml((c.amount || []).join(", ")) || "-"}${c.amountDirect ? ` <span class="amt-direct">+${escapeHtml(c.amountDirect)}만</span>` : ""}</td>
                <td class="hi-ct-bj">${escapeHtml((c.bj || []).join(", ")) || "-"}</td>
                <td class="cc-memo-td">${escapeHtml(c.memo || "")}</td>
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
              <div class="hf"><div class="hfl">인보험 평균</div><div class="hfv">${fn(e.ins)} <small>천원</small></div></div>
              <div class="hf"><div class="hfl">당월 목표</div><div class="hfv">${fn(e.tgt)} <small>천원</small></div></div>
              <div class="hf"><div class="hfl">현재실적</div><div class="hfv">${fn(e.curAct)} <small>천원</small></div></div>
              <div class="hf"><div class="hfl">진도</div><div class="hfv">${fn(e.pct)} <small>%</small></div></div>
              <div class="hf"><div class="hfl">가입설계</div><div class="hfv">${fn(e.plan)} <small>건</small></div></div>
              <div class="hf"><div class="hfl">행복보장</div><div class="hfv">${fn(e.hap)} <small>건</small></div></div>
              <div class="hf"><div class="hfl">주간예상</div><div class="hfv">${fn(e.exp)} <small>천원</small></div></div>
            </div>
            ${clientsHtml}
            ${e.coach ? `<div class="note nv"><strong>📌 코칭포인트</strong><p>${nl(e.coach)}</p></div>` : ""}
            ${e.calcComment ? `<div class="note blue"><strong>✍️ 면담자 의견</strong><p>${nl(e.calcComment)}</p></div>` : ""}
            ${calcHtml}
          </div>
        </div>
      `;
    }).join("");

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
    const clientsHtml = clients.length ? `
      <section class="pr-section">
        <h3>상담고객 (${clients.length}명)</h3>
        <table class="pr-table">
          <thead><tr>
            <th>#</th><th>성명</th><th>고객유형</th><th>상담단계</th><th>활용자료</th><th>제안금액</th><th>보종</th><th>면담 내용</th>
          </tr></thead>
          <tbody>
            ${clients.map((cl, i) => `
              <tr>
                <td>${i + 1}</td>
                <td>${escapeHtml(cl.name || "")}</td>
                <td>${escapeHtml((cl.types || []).join(", "))}</td>
                <td>${escapeHtml((cl.consult || []).join(", "))}</td>
                <td>${escapeHtml((cl.material || []).join(", "))}</td>
                <td>${escapeHtml((cl.amount || []).join(", "))}${cl.amountDirect ? ` / ${escapeHtml(cl.amountDirect)}만원` : ""}</td>
                <td>${escapeHtml((cl.bj || []).join(", "))}</td>
                <td class="memo">${escapeHtml(cl.memo || "")}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </section>
    ` : "";
    const calcHtml = (c.calcAvg || c.calcTgt || c.calcBaseTgt || c.calcComment) ? `
      <section class="pr-section">
        <h3>시상 계산기 스냅샷</h3>
        <table class="pr-info">
          <tr><th>개인 평균실적</th><td>${escapeHtml(c.calcAvg || "-")}</td>
              <th>아너스 기본순증목표</th><td>${escapeHtml(c.calcBaseTgt || "-")}</td>
              <th>희망목표금액</th><td>${escapeHtml(c.calcTgt || "-")}</td></tr>
        </table>
        ${c.calcComment ? `<div class="pr-comment"><strong>✍️ 면담자 의견</strong><p>${escapeHtml(c.calcComment)}</p></div>` : ""}
      </section>
    ` : "";

    const html = `<!DOCTYPE html>
<html lang="ko"><head><meta charset="UTF-8">
<title>면담일지 - ${escapeHtml(s.name || "")} ${escapeHtml(c.seq || "")}차</title>
<style>
  @page { size: A4; margin: 14mm; }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: "Malgun Gothic", "Noto Sans KR", sans-serif; font-size: 12px; color: #1A1A1A; line-height: 1.5; }
  header { text-align: center; border-bottom: 3px solid #E8651A; padding-bottom: 10px; margin-bottom: 14px; }
  header h1 { font-size: 18px; font-weight: 900; color: #1A2744; letter-spacing: -0.5px; }
  header .sub { font-size: 12px; color: #666; margin-top: 4px; }
  .pr-section { margin-bottom: 16px; page-break-inside: avoid; }
  .pr-section h3 { font-size: 13px; font-weight: 800; color: #E8651A; padding: 6px 10px; background: #FFF3EC; border-left: 3px solid #E8651A; margin-bottom: 8px; }
  .pr-info { width: 100%; border-collapse: collapse; font-size: 11px; }
  .pr-info th { background: #F5F5F5; color: #444; font-weight: 700; padding: 5px 8px; text-align: left; border: 1px solid #D5D5D5; white-space: nowrap; width: 1%; }
  .pr-info td { padding: 5px 8px; border: 1px solid #D5D5D5; }
  .pr-table { width: 100%; border-collapse: collapse; font-size: 11px; }
  .pr-table th, .pr-table td { border: 1px solid #D5D5D5; padding: 4px 6px; text-align: left; vertical-align: top; }
  .pr-table th { background: #F5F5F5; font-weight: 700; }
  .pr-table td.memo { white-space: pre-wrap; max-width: 200px; }
  .pr-coach { padding: 10px 12px; background: #FFF8F5; border-left: 4px solid #E8651A; border-radius: 0 6px 6px 0; white-space: pre-wrap; font-size: 12px; line-height: 1.6; }
  .pr-comment { margin-top: 8px; padding: 8px 12px; background: #F5F7FF; border-left: 4px solid #1A2744; }
  .pr-comment p { white-space: pre-wrap; margin-top: 4px; }
  footer { margin-top: 20px; padding-top: 10px; border-top: 1px solid #D5D5D5; font-size: 10px; color: #999; text-align: center; }
</style>
</head><body>
  <header>
    <h1>고객컨설팅 MASTER과정 면담일지 (${escapeHtml(c.seq || "")}차 / 활동점검)</h1>
    <div class="sub">${escapeHtml(c.date || "")} · ${escapeHtml(s.region || "")} · ${escapeHtml(s.center || "")} · ${escapeHtml(s.branch || "")}</div>
  </header>

  <section class="pr-section">
    <h3>기본 정보</h3>
    <table class="pr-info">
      <tr><th>사번</th><td>${escapeHtml(s.empNo)}</td>
          <th>성명</th><td>${escapeHtml(s.name || "")}</td>
          <th>연락처</th><td>${escapeHtml(s.phone || "")}</td></tr>
      <tr><th>기수</th><td>${escapeHtml(s.cohort || "")}</td>
          <th>인보험 평균</th><td>${fmt(c.ins)} 천원</td>
          <th>당월목표</th><td>${fmt(c.tgt)} 천원</td></tr>
      <tr><th>현재실적</th><td>${fmt(c.curAct)} 천원</td>
          <th>진도</th><td>${fmt(c.pct)} %</td>
          <th>주간예상</th><td>${fmt(c.exp)} 천원</td></tr>
      <tr><th>가입설계</th><td>${fmt(c.plan)} 건</td>
          <th>행복보장분석</th><td>${fmt(c.hap)} 건</td>
          <th>차수</th><td>${escapeHtml(c.seq || "")}차</td></tr>
    </table>
  </section>

  ${clientsHtml}

  <section class="pr-section">
    <h3>핵심 코칭포인트 / 후속조치 / 다음주 계획</h3>
    <div class="pr-coach">${escapeHtml(c.coach || c.content || "-")}</div>
  </section>

  ${calcHtml}

  <footer>출력일시: ${new Date().toLocaleString("ko-KR")}</footer>
<script>window.onload = () => { window.print(); };</script>
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
      const stuIns = fn(firstItv?.ins);
      const stuTgt = fn(firstItv?.tgt);
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
      return { s, itvs, stuIns, stuTgt, rows };
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
      });

      const summaryHtml = isLast ? `<tr class="sr">
        <td class="c">합계</td>
        <td class="c strong">${fn(sumIns)}</td>
        <td class="c strong">${fn(sumTgt)}</td>
        <td colspan="9" class="sr-text">
          달성률 D(월평균/목표): <strong>${dRate}</strong> &nbsp;
          L(예상/목표): <strong>${lRate}</strong> &nbsp;&nbsp;
          잔여목표: <strong>${fn(remain)}천원</strong>
        </td>
        <td></td>
        <td class="c strong">${fn(sumExp)}</td>
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
          ${isLast ? cmtHtml : ""}
          ${isLast ? awardHtml : ""}
        </div>
      `;
    }).join("");
  }

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

  // 이력 항목을 폼으로 불러와 수정 모드 진입
  function editInterview(consultId) {
    const c = state.consultations.find((x) => x.id === consultId);
    if (!c) return;
    state.editingConsultId = consultId;

    const setVal = (id, v) => { const el = $("#" + id); if (el) el.value = v ?? ""; };
    setVal("iv-date", c.date || "");
    setVal("iv-seq", c.seq || "");
    setVal("iv-ins", c.ins || "");
    setVal("iv-tgt", c.tgt || "");
    setVal("iv-pct", c.pct || "");
    setVal("iv-curAct", c.curAct || "");
    setVal("iv-plan", c.plan || "");
    setVal("iv-hap", c.hap || "");
    setVal("iv-exp", c.exp || "");
    setVal("iv-coach", c.coach || c.content || "");

    // tgt 자동/수동 모드 판별
    const ins = Number(c.ins) || 0;
    const tgt = Number(c.tgt) || 0;
    state.tgtAutoMode = (ins > 0 && tgt === ins + 200);
    const th = $("#iv-tgt-hint");
    if (th) th.textContent = tgt ? (state.tgtAutoMode ? "▲ 자동" : "✏️ 수동") : "";
    const ih = $("#iv-ins-hint");
    if (ih) ih.textContent = c.seq ? `▲ ${c.seq}차 면담값` : "";

    // 상담고객 복원
    initCR(Array.isArray(c.clients) ? c.clients : []);

    // 시상 계산기 값 복원
    setVal("calc-avg", c.calcAvg || "");
    setVal("calc-base-tgt", c.calcBaseTgt || "");
    setVal("calc-tgt", c.calcTgt || "");
    setVal("calc-comment", c.calcComment || "");
    if (c.calcAvg || c.calcTgt) {
      state.calcOpen = true;
      const sec = $("#calc-section"); if (sec) sec.style.display = "block";
      const icon = $("#calc-toggle-icon"); if (icon) icon.style.transform = "rotate(0deg)";
      calc();
    }

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
    const cancelBtn = $("#btn-iv-cancel-edit");
    if (!btn) return;
    if (state.editingConsultId) {
      btn.textContent = "✏️ 수정 저장";
      btn.classList.add("editing");
      if (cancelBtn) cancelBtn.hidden = false;
    } else {
      btn.textContent = "💾 저장";
      btn.classList.remove("editing");
      if (cancelBtn) cancelBtn.hidden = true;
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

  function render() {
    const list = filteredStudents();
    renderKPIs(list);
    renderSidebarStudentList(list);
    renderStats(list);
    // 선택된 교육생 정보가 갱신되면 상세도 다시 그리기
    if (state.selectedEmpNo) renderStudentDetail();
  }

  // ========== 통계 렌더링 ==========
  function renderStats(list) {
    const body = $("#stats-body");
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

  // ========== 폼 ==========
  function resetForm() {
    state.form = { region: "", center: "", branch: "" };
    ["form-empno","form-name","form-phone","form-base","form-target","form-honors"].forEach((id) => $("#" + id).value = "");
    $("#form-cohort").value = "";
    $("#form-bulk").value = "";
    editingEmpNo = null;
    syncOrgLabels();
  }

  function openEditForm(empNo) {
    const s = state.students.find((x) => x.empNo === empNo);
    if (!s) return;
    state.form = { region: s.region || "", center: s.center || "", branch: s.branch || "" };
    $("#form-empno").value = s.empNo;
    $("#form-name").value = s.name || "";
    $("#form-phone").value = s.phone || "";
    $("#form-base").value = s.base || "";
    $("#form-target").value = s.target || "";
    $("#form-honors").value = s.honors || "";
    $("#form-cohort").value = s.cohort || "";
    editingEmpNo = s.empNo;
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
    return {
      region: state.form.region,
      center: state.form.center,
      branch: state.form.branch,
      cohort: $("#form-cohort").value,
      empNo,
      name: $("#form-name").value.trim(),
      phone: $("#form-phone").value.trim(),
      base: Number($("#form-base").value || 0),
      target: Number($("#form-target").value || 0),
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
      ["기준실적", "base"], ["목표실적", "target"], ["아너스실적", "honors"]
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
    const headers = ["지역단","비전센터","지점","기수","사번","이름","연락처","기준실적","목표실적","아너스실적"];
    const rows = list.map((s) => [
      s.region, s.center, s.branch, s.cohort, s.empNo, s.name, s.phone, s.base, s.target, s.honors
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

    // 등록 버튼
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

    // 설정 탭
    const v = $("#app-version"); if (v) v.textContent = "20260418m";
    $("#btn-export-json").addEventListener("click", () => exportJSON(filteredStudents(), "filtered"));
    $("#btn-export-json-all").addEventListener("click", () => exportJSON(state.students, "all"));
    $("#btn-import-json").addEventListener("click", () => $("#file-import-json").click());
    $("#file-import-json").addEventListener("change", onImportJSONFile);
    $("#btn-delete-filtered").addEventListener("click", onDeleteFiltered);
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
    if (!list.length) { toast("삭제 대상이 없습니다.", "error"); return; }
    const msg = `경고: 현재 필터(${[state.filter.region, state.filter.center, state.filter.branch, state.filter.cohort].filter(Boolean).join(" · ") || "전체"})에 해당하는 ${list.length}명의 교육생을 삭제합니다.\n\n복구할 수 없습니다. 정말 삭제하시겠습니까?`;
    if (!confirm(msg)) return;
    const typed = prompt(`다시 한 번 확인합니다. 삭제를 진행하려면 아래와 동일하게 입력하세요:\n\n삭제 ${list.length}명`);
    if (typed !== `삭제 ${list.length}명`) {
      toast("입력이 일치하지 않아 취소되었습니다.", "");
      return;
    }
    let ok = 0, fail = 0;
    for (const s of list) {
      try { await window.DataAPI.remove(s.empNo); ok++; } catch (e) { console.error(e); fail++; }
    }
    toast(`${ok}명 삭제 완료${fail ? ` / ${fail}명 실패` : ""}`, fail ? "error" : "success");
  }

  function switchView(href) {
    const map = {
      "#dashboard": null,            // KPI 만 표시
      "#students":  "students",      // 교육생 면담 관리
      "#stats":     "stats",         // 통계
      "#settings":  "settings"       // 설정
    };
    const target = map[href];
    // nav active 토글
    $$(".top-nav a").forEach((a) => {
      a.classList.toggle("active", a.getAttribute("href") === href);
    });
    // 모든 view-panel 숨기고 target 만 노출
    $$(".view-panel").forEach((p) => {
      p.hidden = (p.dataset.view !== target);
    });
    // 노출된 패널로 스크롤
    if (target) {
      const el = document.querySelector(`.view-panel[data-view="${target}"]`);
      if (el) el.scrollIntoView({ behavior: "smooth", block: "start" });
      if (target === "students" && !state.selectedEmpNo) {
        toast("좌측 [지점별 교육생] 목록에서 교육생을 선택하세요.", "");
      }
    } else {
      // dashboard — KPI 영역으로 스크롤
      const el = document.getElementById("dashboard");
      if (el) el.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  }

  function retrySubscription() {
    toast("Firestore 재구독 시도중...", "");
    state.studentsLoaded = false;
    if (state._subscribeUnsub) {
      try { state._subscribeUnsub(); } catch (e) {}
    }
    state._subscribeUnsub = window.DataAPI.subscribe((list, meta) => {
      state.students = list || [];
      state.studentsLoaded = true;
      state.syncMeta = meta || { fromCache: false };
      render();
      if (list && list.length > 0) toast(`${list.length}건 동기화 완료`, "success");
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
      render();
    });
  }

  document.addEventListener("DOMContentLoaded", init);
})();
