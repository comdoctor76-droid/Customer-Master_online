/* 고객컨설팅 마스터과정 운영관리 - 메인 앱 로직 */

(function () {
  const LS_KEY = "cmf.filter.v1";
  const DEFAULT_REGION = "호남지역단";

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
    $("#sel-region-text").textContent = state.filter.region || DEFAULT_REGION;
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
      const f = state.filter;
      const path = [f.region, f.center, f.branch].filter(Boolean).join(" > ");
      container.innerHTML = `<div class="empty-state">[${escapeHtml(path)}] 에 등록된 교육생이 없습니다.</div>`;
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
    // 사이드바 필터
    $("#btn-select-region").addEventListener("click", () => openOrgPicker("filter-region"));
    $("#btn-select-center").addEventListener("click", () => openOrgPicker("filter-center"));
    $("#btn-select-branch").addEventListener("click", () => openOrgPicker("filter-branch"));
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

    // 시드 파일 불러오기
    $("#btn-load-seed").addEventListener("click", openSeedPicker);

    // 저장
    $("#btn-save").addEventListener("click", async () => {
      const activeTab = document.querySelector(".tab.active").dataset.tab;
      try {
        if (activeTab === "single") {
          const ok = await saveSingle();
          if (ok) {
            toast("저장되었습니다.", "success");
            closeModal("#modal-add");
          }
        } else {
          // 벌크: 결과를 모달 안에 표시하고, 전부 성공할 때만 자동 닫음
          const { ok, fail } = await saveBulk();
          if (ok > 0 && fail === 0) {
            setTimeout(() => closeModal("#modal-add"), 800);
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
