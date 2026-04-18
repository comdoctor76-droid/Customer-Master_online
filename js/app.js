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
    calcTgtUserEditing: false // 희망목표 직접입력 중 플래그
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
      container.innerHTML = `<div class="empty-mini">조건에 맞는 교육생 없음</div>`;
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
    // 같은 학생 재렌더 요청이면 폼 입력 보존 (학생 정보 표만 갱신)
    if (state.lastDetailEmpNo === s.empNo && document.getElementById("iv-coach")) {
      updateStudentInfoCard(s);
      renderConsultations();
      return;
    }
    state.lastDetailEmpNo = s.empNo;
    body.innerHTML = `
      <div class="detail-grid">
        <div class="detail-card">
          <h3>교육생 정보</h3>
          <table class="detail-info">
            <tr><th>사번</th><td>${escapeHtml(s.empNo)}</td></tr>
            <tr><th>이름</th><td>${escapeHtml(s.name || "")}</td></tr>
            <tr><th>연락처</th><td>${escapeHtml(s.phone || "")}</td></tr>
            <tr><th>지역단</th><td>${escapeHtml(s.region || "")}</td></tr>
            <tr><th>비전센터</th><td>${escapeHtml(s.center || "")}</td></tr>
            <tr><th>지점</th><td>${escapeHtml(s.branch || "")}</td></tr>
            <tr><th>기수</th><td>${escapeHtml(s.cohort || "")}</td></tr>
            <tr><th>평균실적</th><td>${Number(s.base || 0).toLocaleString()}</td></tr>
            <tr><th>목표실적</th><td>${Number(s.target || 0).toLocaleString()}</td></tr>
            <tr><th>순증목표</th><td>${Number(s.honors || 0).toLocaleString()}</td></tr>
          </table>
          <div class="detail-actions">
            <button class="btn-outline" id="btn-detail-edit">교육생 정보 수정</button>
            <button class="btn-outline danger" id="btn-detail-del">교육생 삭제</button>
          </div>
        </div>

        ${renderInterviewFormHtml(s)}

        <div class="detail-card consult-history-card">
          <h3>면담 기록</h3>
          <div id="consult-history" class="consult-history">
            <div class="empty-mini">면담 기록 불러오는 중...</div>
          </div>
        </div>
      </div>
    `;

    $("#btn-detail-edit").addEventListener("click", () => openEditForm(s.empNo));
    $("#btn-detail-del").addEventListener("click", () => removeStudent(s.empNo));
    bindInterviewFormEvents();
    autoFillInterviewForm(s);
    renderConsultations();
  }

  // 같은 학생의 기본정보만 부분 갱신 (폼 입력 보존)
  function updateStudentInfoCard(s) {
    const tbl = document.querySelector(".detail-card .detail-info");
    if (!tbl) return;
    tbl.innerHTML = `
      <tr><th>사번</th><td>${escapeHtml(s.empNo)}</td></tr>
      <tr><th>이름</th><td>${escapeHtml(s.name || "")}</td></tr>
      <tr><th>연락처</th><td>${escapeHtml(s.phone || "")}</td></tr>
      <tr><th>지역단</th><td>${escapeHtml(s.region || "")}</td></tr>
      <tr><th>비전센터</th><td>${escapeHtml(s.center || "")}</td></tr>
      <tr><th>지점</th><td>${escapeHtml(s.branch || "")}</td></tr>
      <tr><th>기수</th><td>${escapeHtml(s.cohort || "")}</td></tr>
      <tr><th>평균실적</th><td>${Number(s.base || 0).toLocaleString()}</td></tr>
      <tr><th>목표실적</th><td>${Number(s.target || 0).toLocaleString()}</td></tr>
      <tr><th>순증목표</th><td>${Number(s.honors || 0).toLocaleString()}</td></tr>
    `;
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
      clearInterviewForm();
      if (s) autoFillInterviewForm(s);
    });
    $("#btn-iv-save").addEventListener("click", saveInterview);
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
    // 차수: 이미 저장된 면담 수 + 1
    const seqEl = $("#iv-seq");
    if (seqEl && !seqEl.value) {
      seqEl.value = String(state.consultations.length + 1);
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
    const btn = $("#btn-iv-save");
    if (btn) { btn.disabled = true; btn.textContent = "저장중..."; }
    try {
      await window.DataAPI.addConsultation(empNo, rec);
      if (rec.ins > 0 && typeof window.DataAPI.updateStudentInsAvg === "function") {
        window.DataAPI.updateStudentInsAvg(empNo, rec.ins).catch((e) => {
          console.warn("[insAvg sync]", e);
        });
      }
      toast("면담 기록이 저장되었습니다.", "success");
      clearInterviewForm();
      // 차수 재채움을 위해 선택한 학생 기준 자동채움 재실행
      const s = state.students.find((x) => x.empNo === empNo);
      if (s) autoFillInterviewForm(s);
    } catch (err) {
      console.error(err);
      toast("저장 실패: " + err.message, "error");
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = "💾 저장"; }
    }
  }

  // 이력 목록: 주요 수치를 배지로 요약, 펼치면 코칭 전문
  function renderConsultations() {
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
      return `
        <div class="consult-entry">
          <div class="consult-head">
            <span class="consult-date">${escapeHtml(c.date || "")}${seqLabel ? " · " + seqLabel : ""}</span>
            <button class="consult-del" data-id="${escapeHtml(c.id)}" title="삭제">×</button>
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
    // 선택된 교육생 정보가 갱신되면 상세도 다시 그리기
    if (state.selectedEmpNo) renderStudentDetail();
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
      $("#search-box").value = "";
      syncOrgLabels();
      persistFilter();
      render();
    });
    $("#filter-cohort").addEventListener("change", (e) => {
      state.filter.cohort = e.target.value;
      persistFilter();
      render();
    });
    $("#search-box").addEventListener("input", (e) => {
      state.filter.q = e.target.value;
      render();
    });

    // 모달 닫기
    $$("[data-close]").forEach((el) => el.addEventListener("click", (e) => {
      e.target.closest(".modal").hidden = true;
    }));

    // 상단 탭 네비게이션
    $$(".top-nav a").forEach((a) => {
      a.addEventListener("click", (e) => {
        e.preventDefault();
        const href = a.getAttribute("href") || "";
        $$(".top-nav a").forEach((x) => x.classList.remove("active"));
        a.classList.add("active");

        if (href === "#dashboard") {
          const el = document.getElementById("dashboard");
          if (el) el.scrollIntoView({ behavior: "smooth", block: "start" });
        } else if (href === "#students") {
          const el = document.getElementById("student-detail-panel");
          if (el) {
            el.scrollIntoView({ behavior: "smooth", block: "start" });
            if (!state.selectedEmpNo) {
              toast("좌측 [지점별 교육생] 목록에서 교육생을 선택하세요.", "");
            }
          }
        } else if (href === "#stats" || href === "#settings") {
          toast("준비중입니다.", "");
        }
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
  }

  function init() {
    bindEvents();
    // localStorage에서 복원된 필터값을 UI에 반영
    $("#filter-cohort").value = state.filter.cohort || "";
    syncOrgLabels();
    window.DataAPI.subscribe((list) => {
      state.students = list || [];
      render();
    });
  }

  document.addEventListener("DOMContentLoaded", init);
})();
