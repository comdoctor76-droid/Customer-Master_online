/* 고객컨설팅 마스터과정 운영관리 - 메인 앱 로직 */

(function () {
  const state = {
    students: [],
    filter: { region: "", center: "", branch: "", cohort: "", q: "" },
    form: { region: "", center: "", branch: "" },
    orgPickerTarget: null // 'filter-region' | 'form-region' 등
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
    if (target.startsWith("filter-")) {
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

    scope[field] = name;
    // 상위 변경 시 하위 초기화
    if (field === "region") { scope.center = ""; scope.branch = ""; }
    if (field === "center") { scope.branch = ""; }

    syncOrgLabels();
    closeModal("#modal-org");
    if (isFilter) render();
  }

  function syncOrgLabels() {
    $("#sel-region-text").textContent = state.filter.region || "전체";
    $("#sel-center-text").textContent = state.filter.center || "전체";
    $("#sel-branch-text").textContent = state.filter.branch || "전체";
    $("#form-region-text").textContent = state.form.region || "선택";
    $("#form-center-text").textContent = state.form.center || "선택";
    $("#form-branch-text").textContent = state.form.branch || "선택";
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
    $("#mini-region").textContent = state.filter.region
      ? state.students.filter((s) => s.region === state.filter.region).length
      : state.students.length;
    $("#mini-branch").textContent = state.filter.branch
      ? state.students.filter((s) => s.branch === state.filter.branch).length
      : 0;
  }

  function renderBranchGroups(list) {
    const container = $("#branch-group-list");
    if (list.length === 0) {
      container.innerHTML = `<div class="empty-state">등록된 교육생이 없습니다. 좌측 [교육생 등록] 버튼으로 추가하세요.</div>`;
      return;
    }
    const groups = {};
    list.forEach((s) => {
      const key = s.branch || "(지점 미지정)";
      if (!groups[key]) groups[key] = [];
      groups[key].push(s);
    });

    container.innerHTML = "";
    Object.keys(groups).sort().forEach((branch) => {
      const rows = groups[branch];
      const first = rows[0];
      const el = document.createElement("div");
      el.className = "branch-group";
      el.innerHTML = `
        <div class="branch-group-head">
          <div>
            <span class="title">${escapeHtml(branch)}</span>
            <span class="sub">${escapeHtml(first.region || "")} · ${escapeHtml(first.center || "")}</span>
          </div>
          <div class="count">${rows.length}명</div>
        </div>
        <div class="branch-group-body">
          <table class="student-table">
            <thead>
              <tr>
                <th>사번</th><th>이름</th><th>기수</th><th>연락처</th>
                <th>기준실적</th><th>목표실적</th><th>아너스실적</th><th>관리</th>
              </tr>
            </thead>
            <tbody>
              ${rows.map((s) => `
                <tr>
                  <td><strong>${escapeHtml(s.empNo)}</strong></td>
                  <td>${escapeHtml(s.name || "")}</td>
                  <td>${escapeHtml(s.cohort || "")}</td>
                  <td>${escapeHtml(s.phone || "")}</td>
                  <td>${Number(s.base || 0).toLocaleString()}</td>
                  <td>${Number(s.target || 0).toLocaleString()}</td>
                  <td>${Number(s.honors || 0).toLocaleString()}</td>
                  <td class="row-actions">
                    <button data-act="edit" data-emp="${escapeHtml(s.empNo)}">수정</button>
                    <button data-act="del" data-emp="${escapeHtml(s.empNo)}">삭제</button>
                  </td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        </div>
      `;
      container.appendChild(el);
    });

    container.querySelectorAll("button[data-act]").forEach((btn) => {
      btn.addEventListener("click", () => {
        const empNo = btn.dataset.emp;
        if (btn.dataset.act === "edit") openEditForm(empNo);
        else if (btn.dataset.act === "del") removeStudent(empNo);
      });
    });
  }

  function escapeHtml(s) {
    return String(s ?? "").replace(/[&<>"']/g, (c) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
    }[c]));
  }

  function render() {
    const list = filteredStudents();
    renderKPIs(list);
    renderBranchGroups(list);
  }

  // ========== 폼 ==========
  function resetForm() {
    state.form = { region: "", center: "", branch: "" };
    ["form-empno","form-name","form-phone","form-base","form-target","form-honors"].forEach((id) => $("#" + id).value = "");
    $("#form-cohort").value = "";
    $("#form-bulk").value = "";
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
    syncOrgLabels();
    switchTab("single");
    openModal("#modal-add");
  }

  function switchTab(name) {
    $$(".tab").forEach((t) => t.classList.toggle("active", t.dataset.tab === name));
    $$(".tab-panel").forEach((p) => (p.hidden = p.dataset.panel !== name));
  }

  async function saveSingle() {
    const empNo = $("#form-empno").value.trim();
    if (!empNo) { toast("사번을 입력하세요.", "error"); return false; }
    const rec = {
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
    await window.DataAPI.save(rec);
    return true;
  }

  async function saveBulk() {
    const raw = $("#form-bulk").value.trim();
    if (!raw) { toast("붙여넣을 데이터가 없습니다.", "error"); return false; }
    const lines = raw.split(/\r?\n/).filter((l) => l.trim());
    let ok = 0, fail = 0;
    for (const line of lines) {
      const cols = line.includes("\t") ? line.split("\t") : line.split(",");
      const [region, center, branch, cohort, empNo, name, phone, base, target, honors, tenureMonths] = cols.map((c) => (c || "").trim());
      if (!empNo) { fail++; continue; }
      const rec = {
        region, center, branch, cohort, empNo, name, phone,
        base: Number(base || 0), target: Number(target || 0), honors: Number(honors || 0)
      };
      if (tenureMonths) rec.tenureMonths = Number(tenureMonths);
      await window.DataAPI.save(rec);
      ok++;
    }
    toast(`${ok}건 저장, ${fail}건 실패`, fail ? "error" : "success");
    return true;
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
    // 사이드바 필터
    $("#btn-select-region").addEventListener("click", () => openOrgPicker("filter-region"));
    $("#btn-select-center").addEventListener("click", () => openOrgPicker("filter-center"));
    $("#btn-select-branch").addEventListener("click", () => openOrgPicker("filter-branch"));
    $("#btn-reset-filter").addEventListener("click", () => {
      state.filter = { region: "", center: "", branch: "", cohort: "", q: "" };
      $("#filter-cohort").value = "";
      $("#search-box").value = "";
      syncOrgLabels();
      render();
    });
    $("#filter-cohort").addEventListener("change", (e) => {
      state.filter.cohort = e.target.value;
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

    // 조직 검색
    $("#org-search").addEventListener("input", (e) => renderOrgList(e.target.value));

    // 등록 버튼
    $("#btn-open-add").addEventListener("click", () => {
      resetForm();
      switchTab("single");
      openModal("#modal-add");
    });

    // 폼 조직 선택
    $("#form-region-btn").addEventListener("click", () => openOrgPicker("form-region"));
    $("#form-center-btn").addEventListener("click", () => openOrgPicker("form-center"));
    $("#form-branch-btn").addEventListener("click", () => openOrgPicker("form-branch"));

    // 탭 전환
    $$(".tab").forEach((t) => t.addEventListener("click", () => switchTab(t.dataset.tab)));

    // 저장
    $("#btn-save").addEventListener("click", async () => {
      const activeTab = document.querySelector(".tab.active").dataset.tab;
      let ok = false;
      try {
        ok = activeTab === "single" ? await saveSingle() : await saveBulk();
        if (ok) {
          if (activeTab === "single") toast("저장되었습니다.", "success");
          closeModal("#modal-add");
        }
      } catch (err) {
        console.error(err);
        toast("저장 실패: " + err.message, "error");
      }
    });

    // CSV
    $("#btn-export-csv").addEventListener("click", exportCSV);
  }

  function init() {
    bindEvents();
    syncOrgLabels();
    window.DataAPI.subscribe((list) => {
      state.students = list || [];
      render();
    });
  }

  document.addEventListener("DOMContentLoaded", init);
})();
