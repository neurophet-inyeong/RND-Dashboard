
    // file:// 로 열렸을 때 서버가 살아있으면 즉시 리다이렉트
    if (location.protocol === 'file:') {
      fetch('http://localhost:8080/', { method: 'HEAD', mode: 'no-cors', cache: 'no-store',
        signal: AbortSignal.timeout(800) })
        .then(() => location.replace('http://localhost:8080/dashboard.html'))
        .catch(() => {}); // 서버 없으면 File System API로 계속 진행
    }

    const state = {
      projects: [],
      activeProject: "",
      activeSheetTab: "사업비",
      view: "project",   // "project" | "allOverview" | "budgetLedger"
      ledger: null,
      ledgerSelectedNum: null,
      ledgerMode: "table"   // "table" | "detail"
    };
    let lastReloadToken = "";

    const LEDGER_PATH = "ledger/총괄표_원본데이터.xlsx";
    let ledgerYearlyCostChart = null;
    let ledgerYearChart = null;

    // ── IndexedDB: 폴더 핸들 영속 저장 ──────────────────────────────────
    function openHandleDB() {
      return new Promise((resolve, reject) => {
        const req = indexedDB.open('rnd-dashboard', 1);
        req.onupgradeneeded = () => req.result.createObjectStore('handles');
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      });
    }
    async function getSavedHandle() {
      try {
        const db = await openHandleDB();
        return await new Promise(resolve => {
          const req = db.transaction('handles', 'readonly').objectStore('handles').get('projectsDir');
          req.onsuccess = () => resolve(req.result || null);
          req.onerror = () => resolve(null);
        });
      } catch { return null; }
    }
    async function saveHandle(handle) {
      try {
        const db = await openHandleDB();
        await new Promise((resolve, reject) => {
          const tx = db.transaction('handles', 'readwrite');
          tx.objectStore('handles').put(handle, 'projectsDir');
          tx.oncomplete = resolve;
          tx.onerror = reject;
        });
      } catch {}
    }

    // ── File System Access API: 폴더에서 엑셀 읽기 ──────────────────────
    async function loadFromDirectoryHandle(dirHandle) {
      const files = [];
      for await (const [name, handle] of dirHandle.entries()) {
        if (handle.kind === 'file' && isExcelPath(name)) {
          files.push(await handle.getFile());
        }
      }
      if (!files.length) {
        setMessage("폴더에 Excel 파일(.xlsx, .xls, .xlsm)이 없습니다.");
        return;
      }
      await loadFilesIntoDashboard(files, "projects 폴더");
    }

    async function pickAndLoadDirectory() {
      if (!('showDirectoryPicker' in window)) {
        setMessage("이 브라우저는 폴더 선택을 지원하지 않습니다. '엑셀 직접 업로드'를 이용해주세요.");
        return;
      }
      let handle = await getSavedHandle();
      if (handle) {
        const perm = await handle.requestPermission({ mode: 'read' });
        if (perm !== 'granted') handle = null;
      }
      if (!handle) {
        try {
          handle = await window.showDirectoryPicker({ mode: 'read' });
          await saveHandle(handle);
        } catch (err) {
          if (err.name !== 'AbortError') setMessage(`폴더 선택 오류: ${err.message}`);
          return;
        }
      }
      reloadProjectBtn.textContent = "📂 폴더 다시 읽기";
      reloadProjectBtn.classList.remove("active");
      await loadFromDirectoryHandle(handle);
    }

    async function autoLoadDirectory() {
      if (!('showDirectoryPicker' in window)) return false;
      const handle = await getSavedHandle();
      if (!handle) return false;
      try {
        const perm = await handle.requestPermission({ mode: 'read' });
        if (perm !== 'granted') return false;
        await loadFromDirectoryHandle(handle);
        reloadProjectBtn.textContent = "📂 폴더 다시 읽기";
        return true;
      } catch {
        return false;
      }
    }

    const projectMenu = document.getElementById("projectMenu");
    const allOverviewBtn = document.getElementById("allOverviewBtn");
    const reloadProjectBtn = document.getElementById("reloadProjectBtn");
    const uploadBtn = document.getElementById("uploadBtn");
    const refreshBtn = document.getElementById("refreshBtn");
    const fileInput = document.getElementById("fileInput");
    const messageEl = document.getElementById("message");
    const baseDate = document.getElementById("baseDate");
    const pageTitle = document.getElementById("pageTitle");

    const overviewGrid = document.getElementById("overviewGrid");
    const sheetTabs = document.getElementById("sheetTabs");
    const sheetHint = document.getElementById("sheetHint");
    const activeSheetWrap = document.getElementById("activeSheetWrap");
    const allOverviewSection = document.getElementById("allOverviewSection");
    const allOverviewGrid = document.getElementById("allOverviewGrid");
    const projectSection = document.getElementById("projectSection");
    const budgetLedgerBtn = document.getElementById("budgetLedgerBtn");
    const budgetLedgerSection = document.getElementById("budgetLedgerSection");
    const budgetLedgerGrid = document.getElementById("budgetLedgerGrid");

    const SHEET_TABS = [
      { key: "사업비", label: "사업비", hint: "사업비 시트 상위 120행" },
      { key: "마일스톤", label: "마일스톤", hint: "마일스톤 시트 상위 120행" },
      { key: "정량지표묶음", label: "정량지표", hint: "정량지표 + 정량지표상세 시트 상위 120행" }
    ];

    function setMessage(text) {
      if (!text) {
        messageEl.style.display = "none";
        messageEl.textContent = "";
        return;
      }
      messageEl.style.display = "block";
      messageEl.textContent = text;
    }

    function formatDateToday() {
      const d = new Date();
      const yyyy = d.getFullYear();
      const mm = String(d.getMonth() + 1).padStart(2, "0");
      const dd = String(d.getDate()).padStart(2, "0");
      return `${yyyy}/${mm}/${dd}`;
    }

    function normalizeKey(key) {
      return String(key || "")
        .replace(/[\n\r]/g, " ")
        .replace(/\s+/g, " ")
        .trim();
    }

    function isExcelPath(path) {
      const name = path.split(/[\\/]/).pop() || "";
      if (name.startsWith("~$")) return false; // 엑셀 실행 중 생성되는 잠금 파일 제외
      return /\.(xlsx|xls|xlsm)$/i.test(path);
    }

    function sheetRows(workbook, sheetName, headerRow = 1) {
      if (!workbook.Sheets[sheetName]) return [];
      const ws = workbook.Sheets[sheetName];

      if (headerRow <= 1) {
        const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
        return rows.map((row) => {
          const out = {};
          for (const [k, v] of Object.entries(row)) {
            out[normalizeKey(k) || "미지정"] = v;
          }
          return out;
        });
      }

      const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
      const headerIdx = Math.max(0, headerRow - 1);
      const headerRaw = matrix[headerIdx] || [];
      const header = headerRaw.map((h, i) => normalizeKey(h) || `미지정_${i + 1}`);
      const body = matrix.slice(headerIdx + 1);

      return body
        .filter((r) => Array.isArray(r) && r.some((v) => v !== null && String(v).trim() !== ""))
        .map((r) => {
          const out = {};
          header.forEach((h, i) => {
            out[h] = r[i] ?? null;
          });
          return out;
        });
    }

    function makeUniqueHeaders(headers) {
      const counter = new Map();
      return headers.map((h, idx) => {
        const base = normalizeKey(h) || `미지정_${idx + 1}`;
        const n = (counter.get(base) || 0) + 1;
        counter.set(base, n);
        return n === 1 ? base : `${base}_${n}`;
      });
    }

    // raw matrix(첫 행=헤더) → row 객체 배열
    function parseMatrixToRows(matrix) {
      if (matrix.length < 2) return [];
      const header = makeUniqueHeaders(
        (matrix[0] || []).map((h, i) => normalizeKey(h) || `미지정_${i + 1}`)
      );
      return matrix.slice(1)
        .filter(r => r.some(v => v !== null && String(v).trim() !== ""))
        .map(r => {
          const out = {};
          header.forEach((h, i) => { out[h] = r[i] ?? null; });
          return out;
        });
    }

    // 개요+사업비가 합쳐진 시트를 분리
    // 개요: A2:B15 고정 (matrix 인덱스 1~14)
    // 사업비: 단계-연차 + 시작일 + 종료일 헤더 행부터
    function parseCombinedSheet(workbook, sheetName) {
      const ws = workbook.Sheets[sheetName];
      if (!ws) return null;

      const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
      if (!matrix.length) return null;

      // 개요: A2:B15 고정 범위 (matrix[1]~matrix[14])
      const overviewRows = [];
      for (let r = 1; r <= 14 && r < matrix.length; r++) {
        const row = matrix[r] || [];
        const a = row[0] ?? null;
        const b = row[1] ?? null;
        if (a !== null && String(a).trim()) {
          overviewRows.push({ '항목': a, '내용': b });
        }
      }

      // 사업비: 단계-연차 + 시작일 + 종료일 헤더 감지
      let budgetHeaderIdx = -1;
      for (let i = 1; i < matrix.length; i++) {
        const texts = (matrix[i] || []).map(v => normalizeKey(v));
        const hasYearCol  = texts.some(t => /단계.{0,5}연차|단계-연차/.test(t));
        const hasStartCol = texts.some(t => t === "시작일" || /^시작일$/.test(t));
        const hasEndCol   = texts.some(t => t === "종료일" || /^종료일$/.test(t));
        if (hasYearCol && hasStartCol && hasEndCol) {
          budgetHeaderIdx = i;
          break;
        }
      }

      const nonEmpty = r => r.some(v => v !== null && String(v).trim() !== "");
      const budgetMatrix = budgetHeaderIdx > 0
        ? matrix.slice(budgetHeaderIdx).filter(nonEmpty)
        : [];

      if (!overviewRows.length && !budgetMatrix.length) return null;

      return {
        overviewRows,
        budgetRows: parseMatrixToRows(budgetMatrix),
      };
    }

    function parseMilestoneRows(workbook, sheetName) {
      const ws = workbook.Sheets[sheetName];
      if (!ws) return [];

      const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
      const groupRaw = (matrix[0] || []).map((v) => normalizeKey(v));
      const subRaw = (matrix[1] || []).map((v, i) => normalizeKey(v) || `미지정_${i + 1}`);

      const groups = [];
      let carry = "";
      for (const g of groupRaw) {
        if (g) carry = g;
        groups.push(carry);
      }

      const combined = subRaw.map((sub, i) => {
        const grp = groups[i];
        if (grp === "계획" || grp === "실행") {
          return `${grp}|${sub}`;
        }
        return sub;
      });

      const header = makeUniqueHeaders(combined);
      const body = matrix.slice(2);

      return body
        .filter((r) => Array.isArray(r) && r.some((v) => v !== null && String(v).trim() !== ""))
        .map((r) => {
          const out = {};
          header.forEach((h, i) => {
            out[h] = r[i] ?? null;
          });
          return out;
        });
    }

    function parseMetricRows(workbook, sheetName) {
      const ws = workbook.Sheets[sheetName];
      if (!ws) return [];

      const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
      const topRaw = matrix[0] || [];
      const subRaw = matrix[1] || [];

      const top = topRaw.map((v) => normalizeKey(v));
      const sub = subRaw.map((v) => normalizeKey(v));
      const groupedKeys = new Set(["목표", "달성"]);

      const groups = [];
      let carry = "";
      for (const t of top) {
        if (t && !/^Unnamed/i.test(t)) {
          carry = t;
        }
        groups.push(carry);
      }

      const combined = top.map((t, i) => {
        const g = groups[i];
        const s = sub[i] || "";

        if (groupedKeys.has(g)) {
          return `${g}|${s || `항목_${i + 1}`}`;
        }

        if (g.includes("연구개발") && g.includes("달성률")) {
          return `연구개발달성률|${s || `항목_${i + 1}`}`;
        }

        if (g && !/^Unnamed/i.test(g)) {
          return g;
        }

        if (s) {
          return s;
        }

        return `미지정_${i + 1}`;
      });

      const header = makeUniqueHeaders(combined);
      const body = matrix.slice(2);

      return body
        .filter((r) => Array.isArray(r) && r.some((v) => v !== null && String(v).trim() !== ""))
        .map((r) => {
          const out = {};
          header.forEach((h, i) => {
            out[h] = r[i] ?? null;
          });
          return out;
        });
    }

    function getActiveProject() {
      if (!state.projects.length) return null;
      if (!state.activeProject) return state.projects[0];
      return state.projects.find((p) => p.name === state.activeProject) || state.projects[0];
    }

    // ── 통합 개요: 개요 필드 맵 추출 ─────────────────────────────────────
    const OV_FIELDS = [
      "사업명", "당해기간", "주관/공동(책임자)"
    ];

    function extractOverviewMap(project) {
      const rows = project.sheets["개요"] || [];
      const map = {};
      rows.forEach((row) => {
        const entries = Object.entries(row);
        if (!entries.length) return;
        const key = String(entries[0][1] ?? entries[0][0] ?? "").trim();
        if (!key) return;
        const rawVal = entries.length >= 2 ? (entries[1][1] ?? "") : "";
        map[key] = key === "공고일"
          ? (toYmd(rawVal) || String(rawVal ?? "").trim())
          : String(rawVal ?? "").trim();
      });
      return map;
    }

    function renderAllOverview() {
      allOverviewGrid.innerHTML = "";

      if (!state.projects.length) {
        allOverviewGrid.innerHTML = "<p style='color:var(--muted)'>불러온 과제가 없습니다.</p>";
        return;
      }

      const ALL_FIELDS = ["과제번호", "과제명", ...OV_FIELDS];
      const maps = state.projects.map(p => extractOverviewMap(p));

      // 헤더: 필드명
      const thead = `<thead><tr>
        ${ALL_FIELDS.map(f => `<th>${f}</th>`).join("")}
      </tr></thead>`;

      // 행: 과제별 데이터
      const tbody = `<tbody>${state.projects.map((project, i) => {
        const map = maps[i];
        const isActive = project.name === state.activeProject;
        const cells = ALL_FIELDS.map(f => `<td>${map[f] || ""}</td>`).join("");
        return `<tr class="ov-row${isActive ? " ov-row-active" : ""}" data-name="${project.name}">${cells}</tr>`;
      }).join("")}</tbody>`;

      const wrap = document.createElement("div");
      wrap.style.cssText = "overflow-y:auto;width:100%;height:calc(100vh - 160px);border:1px solid var(--line);border-radius:14px;background:#fff;box-sizing:border-box;box-shadow:0 10px 24px rgba(27,45,75,0.08);";
      wrap.innerHTML = `<table class="sheet-table ov-table">${thead}${tbody}</table>`;

      // 행 클릭 → 과제 상세로 이동
      wrap.querySelectorAll("tr.ov-row").forEach(tr => {
        tr.style.cursor = "pointer";
        tr.addEventListener("click", () => {
          state.activeProject = tr.dataset.name;
          switchView("project");
          renderProjectMenu();
          renderDashboard();
        });
      });

      allOverviewGrid.appendChild(wrap);
    }

    // ── 총괄표(예산 현황) ────────────────────────────────────────────────
    function formatKrw(value) {
      const num = Number(value) || 0;
      return `${Math.round(num).toLocaleString("ko-KR")} 원`;
    }

    function loadLedgerSheetRows(workbook, sheetName) {
      const ws = workbook.Sheets[sheetName];
      if (!ws) return [];
      const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
      return parseMatrixToRows(matrix);
    }

    function sumBy(rows, key) {
      return rows.reduce((acc, r) => acc + (parseNumericLike(r[key]) || 0), 0);
    }

    const LEDGER_COST_CATEGORIES = [
      { key: "인건비", label: "인건비", color: "#4a66ff" },
      { key: "연구활동비", label: "연구활동비", color: "#27ae6f" },
      { key: "연구재료비", label: "연구재료비", color: "#f1a94d" },
      { key: "연구시설장비비", label: "연구시설장비비", color: "#e0763f" },
      { key: "연구수당", label: "연구수당", color: "#9b6bdb" },
      { key: "간접비", label: "간접비", color: "#4dbfc7" }
    ];

    const LEDGER_GOV_FUND_KEY = "정부지원금 (현금)";

    // 원본데이터의 시작일 연도 기준으로 비용 항목(인건비/연구활동비 등) + 정부지원금을 합산
    function groupYearlyCosts(detailRows) {
      const byYear = new Map();
      for (const r of detailRows) {
        const d = parseToDate(r["시작일"]);
        if (!d) continue;
        const year = d.getFullYear();
        if (!byYear.has(year)) byYear.set(year, {});
        const bucket = byYear.get(year);
        for (const cat of LEDGER_COST_CATEGORIES) {
          bucket[cat.key] = (bucket[cat.key] || 0) + (parseNumericLike(r[cat.key]) || 0);
        }
        bucket[LEDGER_GOV_FUND_KEY] = (bucket[LEDGER_GOV_FUND_KEY] || 0) + (parseNumericLike(r[LEDGER_GOV_FUND_KEY]) || 0);
      }
      const years = [...byYear.keys()].sort((a, b) => a - b);
      return { years, categories: LEDGER_COST_CATEGORIES, map: byYear };
    }

    async function loadLedgerData() {
      if (state.ledger) return state.ledger;

      const resp = await fetch(LEDGER_PATH, { cache: "no-store" });
      if (!resp.ok) {
        throw new Error(
          "총괄표 파일을 찾을 수 없습니다 (ledger/총괄표_원본데이터.xlsx). sync_ledger.py가 실행 중인지 확인해주세요."
        );
      }
      const data = await resp.arrayBuffer();
      const wb = XLSX.read(data, { type: "array" });

      const summaryAll = loadLedgerSheetRows(wb, "총괄표");
      const summary = summaryAll.filter((r) => {
        const dept = String(r["부처"] ?? "").trim();
        return dept && dept !== "0" && r["총 사업비"] != null;
      });

      const detailAll = loadLedgerSheetRows(wb, "원본데이터");
      const detail = detailAll.filter((r) => r["시트명"] != null);

      state.ledger = { summary, detail };
      return state.ledger;
    }

    const LEDGER_TABLE_COLUMNS = [
      { key: "번호", label: "번호", cls: "ledger-col-num" },
      { key: "유형", label: "유형" },
      { key: "사업명", label: "사업명" },
      { key: "과제명", label: "과제명" },
      { key: "총 수행기간", label: "총 수행기간" },
      { key: "총 사업비", label: "총 사업비", money: true }
    ];

    function renderLedgerBody(summary, detail) {
      const body = document.getElementById("ledgerBody");
      if (!body) return;

      if (state.ledgerMode === "detail" && state.ledgerSelectedNum != null && summary.some((r) => Number(r["번호"]) === state.ledgerSelectedNum)) {
        renderLedgerDetailView(body, summary, detail);
      } else {
        state.ledgerMode = "table";
        renderLedgerTableView(body, summary, detail);
      }
    }

    function renderLedgerTableView(body, summary, detail) {
      if (ledgerYearChart) { ledgerYearChart.destroy(); ledgerYearChart = null; }

      const sorted = summary.slice().sort((a, b) => Number(a["번호"]) - Number(b["번호"]));
      const thead = `<thead><tr>${LEDGER_TABLE_COLUMNS.map((c) => `<th class="${c.cls || ""}">${c.label}</th>`).join("")}</tr></thead>`;
      const tbody = `<tbody>${sorted
        .map((row) => {
          const cells = LEDGER_TABLE_COLUMNS.map((c) => {
            const raw = row[c.key];
            const text = c.money ? formatKrw(raw) : String(raw ?? "");
            if (c.key === "과제명") {
              return `<td class="${c.cls || ""}"><a href="#" class="ledger-open-detail" data-num="${row["번호"]}" style="color:var(--accent);font-weight:700;text-decoration:none;">${text || "(과제명 미입력)"}</a></td>`;
            }
            return `<td class="${c.cls || ""}">${text}</td>`;
          }).join("");
          return `<tr class="ov-row" data-num="${row["번호"]}">${cells}</tr>`;
        })
        .join("")}</tbody>`;

      body.innerHTML = `<h3 class="card-title">과제별 상세 (총괄표 + 연차별 예산 breakdown)</h3>
        <p class="hint">과제명을 클릭하면 상세로 이동합니다.</p>
        <div class="table-wrap" style="max-height:520px;">
          <table class="sheet-table ledger-table">${thead}${tbody}</table>
        </div>`;

      const openDetail = (num) => {
        state.ledgerSelectedNum = num;
        state.ledgerMode = "detail";
        renderLedgerBody(summary, detail);
      };

      body.querySelectorAll("tr.ov-row").forEach((tr) => {
        tr.style.cursor = "pointer";
        tr.addEventListener("click", (e) => {
          if (e.target.closest("a.ledger-open-detail")) return; // 링크 클릭은 자체 핸들러가 처리
          openDetail(Number(tr.dataset.num));
        });
      });
      body.querySelectorAll("a.ledger-open-detail").forEach((a) => {
        a.addEventListener("click", (e) => {
          e.preventDefault();
          openDetail(Number(a.dataset.num));
        });
      });
    }

    function renderLedgerDetailView(body, summary, detail) {
      const num = state.ledgerSelectedNum;
      const row = summary.find((r) => Number(r["번호"]) === num);
      if (!row) {
        state.ledgerMode = "table";
        renderLedgerTableView(body, summary, detail);
        return;
      }

      const fields = [
        "유형", "부처", "전문기관", "사업명", "과제명", "과제번호",
        "주관/공동(책임자)", "총 수행기간", "총 사업비", "정부지원금 (현금)", "총 민간부담금"
      ];
      const moneyFields = new Set(["총 사업비", "정부지원금 (현금)", "총 민간부담금"]);
      const detailTableHtml = `<table class="overview-table"><tbody>${fields
        .map((f) => {
          const text = moneyFields.has(f) ? formatKrw(row[f]) : String(row[f] ?? "");
          return `<tr><th>${f}</th><td colspan="3">${text}</td></tr>`;
        })
        .join("")}</tbody></table>`;

      const yearly = detail.filter((r) => Number(r["시트명"]) === num);
      const yearlyHtml = yearly.length
        ? `<h3 class="card-title" style="margin-top:14px;">연차별 사업비</h3>
           <canvas id="ledgerYearChart" height="180"></canvas>
           <div class="table-wrap" style="margin-top:10px;">${buildSheetTableHtml(yearly, "사업비")}</div>`
        : `<p class="hint" style="margin-top:14px;">이 과제의 연차별 상세 데이터(원본데이터)가 없습니다.</p>`;

      body.innerHTML = `
        <button id="ledgerBackBtn" type="button" class="refresh-btn" style="margin-bottom:14px;">← 목록으로</button>
        <h3 class="card-title">[${row["번호"]}] ${row["사업명"] ?? ""}</h3>
        ${detailTableHtml}
        ${yearlyHtml}
      `;

      document.getElementById("ledgerBackBtn").addEventListener("click", () => {
        state.ledgerMode = "table";
        renderLedgerBody(summary, detail);
      });

      if (ledgerYearChart) { ledgerYearChart.destroy(); ledgerYearChart = null; }
      if (yearly.length) {
        const yearCtx = document.getElementById("ledgerYearChart");
        ledgerYearChart = new Chart(yearCtx, {
          type: "bar",
          data: {
            labels: yearly.map((r) => String(r["단계-연차"] ?? "")),
            datasets: [{
              label: "합계",
              data: yearly.map((r) => parseNumericLike(r["합계"]) || 0),
              backgroundColor: "#27ae6f"
            }]
          },
          options: {
            plugins: { legend: { display: false } },
            scales: { y: { ticks: { callback: (v) => Number(v).toLocaleString("ko-KR") } } }
          }
        });
      }
    }

    async function renderBudgetLedger() {
      if (state.view !== "budgetLedger") return;
      budgetLedgerGrid.innerHTML = `<p style="color:var(--muted)">불러오는 중...</p>`;

      let ledger;
      try {
        ledger = await loadLedgerData();
      } catch (err) {
        budgetLedgerGrid.innerHTML = `<div class="warn" style="display:block">${err.message}</div>`;
        return;
      }

      const { summary, detail } = ledger;
      if (!summary.length) {
        budgetLedgerGrid.innerHTML = `<p style="color:var(--muted)">총괄표에서 유효한 과제 데이터를 찾지 못했습니다.</p>`;
        return;
      }

      const totalBudget = sumBy(summary, "총 사업비");
      const govFund = sumBy(summary, "정부지원금 (현금)");
      const privateFund = sumBy(summary, "총 민간부담금");
      const yearlyCost = groupYearlyCosts(detail);

      budgetLedgerGrid.innerHTML = `
        <div class="card" style="margin-bottom:16px;">
          <h2 class="card-title">전사 과제 예산 현황 (SharePoint 연동 · 총괄표/원본데이터)</h2>
          <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-top:10px;">
            <div class="ov-card"><div class="ov-card-badge">전체 과제 수</div><div class="ov-card-name">${summary.length.toLocaleString("ko-KR")}건</div></div>
            <div class="ov-card"><div class="ov-card-badge">총 사업비 합계</div><div class="ov-card-name">${formatKrw(totalBudget)}</div></div>
            <div class="ov-card"><div class="ov-card-badge">정부지원금 합계</div><div class="ov-card-name">${formatKrw(govFund)}</div></div>
            <div class="ov-card"><div class="ov-card-badge">민간부담금 합계</div><div class="ov-card-name">${formatKrw(privateFund)}</div></div>
          </div>
        </div>
        <div class="card" style="margin-bottom:16px;">
          <div id="ledgerBody"></div>
        </div>
        <div class="card" style="margin-bottom:16px;">
          <h3 class="card-title">연도별 비용 항목 합계 (원본데이터)</h3>
          <canvas id="ledgerYearlyCostChart" height="220"></canvas>
        </div>
      `;

      if (ledgerYearlyCostChart) { ledgerYearlyCostChart.destroy(); ledgerYearlyCostChart = null; }
      const yearlyCostCtx = document.getElementById("ledgerYearlyCostChart");
      const barCount = yearlyCost.categories.length;
      const stackTotalLabelPlugin = {
        id: "stackTotalLabel",
        afterDatasetsDraw(chart) {
          const { ctx, data } = chart;
          const lastBarMeta = chart.getDatasetMeta(barCount - 1);
          ctx.save();
          ctx.font = "bold 11px sans-serif";
          ctx.fillStyle = "#33486f";
          ctx.textAlign = "center";
          ctx.textBaseline = "bottom";
          lastBarMeta.data.forEach((bar, index) => {
            const total = data.datasets
              .slice(0, barCount)
              .reduce((sum, ds) => sum + (ds.data[index] || 0), 0);
            if (total > 0) {
              ctx.fillText(total.toLocaleString("ko-KR"), bar.x, bar.y - 4);
            }
          });
          ctx.restore();
        }
      };
      ledgerYearlyCostChart = new Chart(yearlyCostCtx, {
        data: {
          labels: yearlyCost.years.map(String),
          datasets: [
            ...yearlyCost.categories.map((cat) => ({
              type: "bar",
              label: cat.label,
              data: yearlyCost.years.map((y) => yearlyCost.map.get(y)[cat.key] || 0),
              backgroundColor: cat.color,
              yAxisID: "y",
              order: 1
            })),
            {
              type: "line",
              label: "정부지원금 (현금)",
              data: yearlyCost.years.map((y) => yearlyCost.map.get(y)[LEDGER_GOV_FUND_KEY] || 0),
              borderColor: "#d9345c",
              backgroundColor: "#d9345c",
              yAxisID: "y1",
              tension: 0.3,
              pointRadius: 3,
              order: 2
            }
          ]
        },
        plugins: [stackTotalLabelPlugin],
        options: {
          layout: { padding: { top: 22 } },
          plugins: { legend: { position: "bottom" } },
          scales: {
            x: { stacked: true },
            y: {
              stacked: true,
              ticks: { callback: (v) => Number(v).toLocaleString("ko-KR") }
            },
            y1: {
              position: "right",
              grid: { drawOnChartArea: false },
              ticks: { callback: (v) => Number(v).toLocaleString("ko-KR") }
            }
          }
        }
      });

      renderLedgerBody(summary, detail);
    }

    function switchView(view) {
      state.view = view;
      const mainEl = document.querySelector(".main");

      allOverviewSection.classList.add("hidden");
      budgetLedgerSection.classList.add("hidden");
      projectSection.classList.add("hidden");
      allOverviewBtn.classList.remove("active");
      budgetLedgerBtn.classList.remove("active");

      if (view === "allOverview") {
        allOverviewSection.classList.remove("hidden");
        pageTitle.textContent = "통합 개요";
        allOverviewBtn.classList.add("active");
        mainEl.style.padding = "24px 24px 0";
        renderProjectMenu();
        renderAllOverview();
      } else if (view === "budgetLedger") {
        budgetLedgerSection.classList.remove("hidden");
        pageTitle.textContent = "총괄표 (예산 현황)";
        budgetLedgerBtn.classList.add("active");
        mainEl.style.padding = "";
        renderProjectMenu();
        renderBudgetLedger();
      } else {
        projectSection.classList.remove("hidden");
        pageTitle.textContent = "과제 현황";
        mainEl.style.padding = "";
        renderProjectMenu();
      }
    }

    function renderProjectMenu() {
      projectMenu.innerHTML = "";
      for (const project of state.projects) {
        const btn = document.createElement("button");
        const isActive = state.view === "project" && project.name === state.activeProject;
        btn.className = `menu-btn${isActive ? " active" : ""}`;
        btn.textContent = `📊 ${project.name}`;
        btn.addEventListener("click", () => {
          state.activeProject = project.name;
          switchView("project");
          renderProjectMenu();
          renderDashboard();
        });
        projectMenu.appendChild(btn);
      }
    }

    function renderOverview(project) {
      const rows = project.sheets["개요"] || [];
      if (!rows.length) {
        overviewGrid.innerHTML = "개요 데이터가 없습니다.";
        return;
      }

      const fields = [];
      rows.forEach((row) => {
        const allEntries = Object.entries(row);
        if (!allEntries.length) return;
        const key = String(allEntries[0][1] ?? allEntries[0][0] ?? "").trim();
        if (!key) return;
        const rawValue = allEntries.length >= 2 ? (allEntries[1][1] ?? "") : "";
        const value = String(rawValue ?? "").trim();
        fields.push({ key, value, rawValue });
      });

      const dedup = [];
      const seen = new Set();
      for (const item of fields) {
        const signature = `${item.key}::${item.value}`;
        if (!seen.has(signature)) {
          dedup.push(item);
          seen.add(signature);
        }
      }

      const finalItems = dedup.length ? dedup : [{ key: "안내", value: "개요 시트 형식을 확인해주세요." }];
      const fmtVal = (item) => item.key === "공고일" ? toYmd(item.rawValue) || item.value : item.value;
      const body = [];
      for (let i = 0; i < finalItems.length; i += 2) {
        const a = finalItems[i];
        const b = finalItems[i + 1];
        if (b) {
          body.push(`<tr><th>${a.key}</th><td>${fmtVal(a)}</td><th>${b.key}</th><td>${fmtVal(b)}</td></tr>`);
        } else {
          body.push(`<tr><th>${a.key}</th><td colspan="3">${fmtVal(a)}</td></tr>`);
        }
      }

      overviewGrid.innerHTML = `<table class="overview-table"><tbody>${body.join("")}</tbody></table>`;
    }

    function toYmd(value) {
      if (value === null || value === undefined || value === "") return "";

      if (value instanceof Date && !Number.isNaN(value.getTime())) {
        const yyyy = value.getFullYear();
        const mm = String(value.getMonth() + 1).padStart(2, "0");
        const dd = String(value.getDate()).padStart(2, "0");
        return `${yyyy}-${mm}-${dd}`;
      }

      if (typeof value === "number") {
        const parsed = XLSX.SSF.parse_date_code(value);
        if (parsed) {
          const yyyy = String(parsed.y).padStart(4, "0");
          const mm = String(parsed.m).padStart(2, "0");
          const dd = String(parsed.d).padStart(2, "0");
          return `${yyyy}-${mm}-${dd}`;
        }
      }

      const text = String(value).trim();
      if (!text) return "";

      const normalized = text.replace(/\./g, "-").replace(/\//g, "-");
      const exact = normalized.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
      if (exact) {
        const yyyy = exact[1];
        const mm = String(Number(exact[2])).padStart(2, "0");
        const dd = String(Number(exact[3])).padStart(2, "0");
        return `${yyyy}-${mm}-${dd}`;
      }

      const d = new Date(text);
      if (!Number.isNaN(d.getTime())) {
        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, "0");
        const dd = String(d.getDate()).padStart(2, "0");
        return `${yyyy}-${mm}-${dd}`;
      }

      return text;
    }

    function parseToDate(value) {
      if (value === null || value === undefined || value === "") return null;
      if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
      if (typeof value === "number") {
        if (value <= 0) return null; // 날짜 포맷 적용된 빈 셀 → 0 반환 → 무시
        const parsed = XLSX.SSF.parse_date_code(value);
        if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
      }
      const text = String(value).trim();
      if (!text) return null;
      const normalized = text.replace(/\./g, "-").replace(/\//g, "-");
      const exact = normalized.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
      if (exact) return new Date(Number(exact[1]), Number(exact[2]) - 1, Number(exact[3]));
      const d = new Date(text);
      return isNaN(d.getTime()) ? null : d;
    }

    // 사업비 시트에서 단계-연차별 기간 추출 (A=단계-연차, B=시작일, C=종료일)
    function parseBudgetYears(budgetRows) {
      if (!budgetRows.length) return [];
      const columns = Object.keys(budgetRows[0]);
      const labelCol = columns.find(c => /단계.*연차|연차.*단계|단계|연차/.test(c)) || columns[0];
      const startCol = columns.find(c => /시작일/.test(c)) || columns[1];
      const endCol   = columns.find(c => /종료일/.test(c)) || columns[2];

      const seen = new Set();
      const years = [];
      for (const row of budgetRows) {
        const label = String(row[labelCol] ?? "").trim();
        if (!label || seen.has(label)) continue;
        const start = parseToDate(row[startCol]);
        const end   = parseToDate(row[endCol]);
        if (!start || !end || end < start) continue;
        seen.add(label);
        years.push({ label, start, end });
      }
      return years;
    }

    function buildProgressStatusCardHtml(progressRows) {
      if (!progressRows || !progressRows.length) return "";

      const columns = Object.keys(progressRows[0]);
      const dateCol    = columns.find(c => /날짜/.test(c)) || columns[0];
      const statusCol  = columns.find(c => /상태/.test(c)) || columns[1];
      const contentCol = columns.find(c => /내용/.test(c)) || columns[2];

      const statusClass = (status) => {
        const s = String(status ?? "").trim().toLowerCase();
        if (/risk|위험|지연|미달/.test(s)) return "progress-status-risk";
        if (/done|완료/.test(s)) return "progress-status-done";
        return "progress-status-progress";
      };

      const validRows = progressRows
        .filter(row => String(row[statusCol] ?? "").trim() || String(row[contentCol] ?? "").trim());

      const latestTime = validRows.reduce((max, row) => {
        const d = parseToDate(row[dateCol]);
        return d && (max === null || d.getTime() > max) ? d.getTime() : max;
      }, null);

      const items = validRows
        .filter(row => {
          if (latestTime === null) return true;
          const d = parseToDate(row[dateCol]);
          return d && d.getTime() === latestTime;
        })
        .map(row => {
          const date = toYmd(row[dateCol]) || String(row[dateCol] ?? "").trim();
          const status = String(row[statusCol] ?? "").trim();
          const content = String(row[contentCol] ?? "").trim();
          return `<div class="progress-status-item">
            <span class="progress-status-date">${date}</span>
            <span class="progress-status-badge ${statusClass(status)}">${status}</span>
            <span class="progress-status-content">${content}</span>
          </div>`;
        }).join("");

      if (!items) return "";

      return `<div class="card progress-status-card">
        <h3 class="card-title">진행상황</h3>
        <div class="progress-status-list">${items}</div>
      </div>`;
    }

    function buildMilestoneGanttHtml(rows, budgetRows, progressRows) {
      if (!rows.length) return "<p>데이터가 없습니다.</p>";

      const columns = Object.keys(rows[0]);
      const yearCol   = columns.find(c => /연차|차수|연도|년차/.test(c));
      const nameCol   = columns.find(c => /마일스톤|세부내용|내용|명칭|항목/.test(c)) || columns[1];
      const planStart = columns.find(c => c.startsWith("계획|") && /시작/.test(c));
      const planEnd   = columns.find(c => c.startsWith("계획|") && /종료/.test(c));
      const actStart  = columns.find(c => c.startsWith("실행|") && /시작/.test(c));
      const actEnd    = columns.find(c => c.startsWith("실행|") && /종료/.test(c));
      const hasActual = !!(actStart && actEnd);

      const dateFmt = (d) => d
        ? `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`
        : "";

      // 사업비에서 연차별 기간 정의 가져오기
      const budgetYears = parseBudgetYears(budgetRows || []);

      // groups: [{ label, start, end, rows[] }]
      let groups;

      if (budgetYears.length) {
        // 사업비 연차 기준으로 그룹 구성 (순서 유지)
        groups = budgetYears.map(by => ({ label: by.label, start: by.start, end: by.end, rows: [] }));

        for (const row of rows) {
          const ps = parseToDate(planStart ? row[planStart] : null);
          const pe = parseToDate(planEnd   ? row[planEnd]   : null);

          if (ps && pe) {
            // 날짜가 겹치는 모든 연차에 추가
            let placed = false;
            for (const grp of groups) {
              if (ps <= grp.end && pe >= grp.start) {
                grp.rows.push(row);
                placed = true;
              }
            }
            // 겹치는 연차가 없으면 연차 컬럼 기준 fallback
            if (!placed && yearCol) {
              const yr = String(row[yearCol] ?? "").trim();
              const g = groups.find(g => g.label === yr);
              if (g) g.rows.push(row);
            }
          } else {
            // 날짜 없는 경우 연차 컬럼으로 배치
            const yr = yearCol ? String(row[yearCol] ?? "").trim() : "";
            const g = groups.find(g => g.label === yr);
            if (g) g.rows.push(row);
          }
        }
      } else {
        // fallback: 마일스톤의 연차 컬럼으로 그룹화, 날짜 범위는 해당 연차 데이터에서 계산
        const map = new Map();
        for (const row of rows) {
          const yr = yearCol ? String(row[yearCol] ?? "").trim() : "전체";
          if (!yr) continue;
          if (!map.has(yr)) map.set(yr, { label: yr, start: null, end: null, rows: [] });
          map.get(yr).rows.push(row);
        }
        for (const grp of map.values()) {
          for (const row of grp.rows) {
            for (const col of [planStart, planEnd, actStart, actEnd]) {
              if (!col) continue;
              const d = parseToDate(row[col]);
              if (!d) continue;
              if (!grp.start || d < grp.start) grp.start = d;
              if (!grp.end   || d > grp.end)   grp.end   = d;
            }
          }
          if (grp.start) {
            grp.start = new Date(grp.start.getFullYear(), grp.start.getMonth(), 1);
            grp.end   = new Date(grp.end.getFullYear(), grp.end.getMonth() + 1, 0);
          }
        }
        groups = [...map.values()];
      }

      const legend = `<div class="gantt-legend">
        <div class="gantt-legend-item"><div class="gantt-legend-dot plan"></div><span>계획</span></div>
        ${hasActual ? '<div class="gantt-legend-item"><div class="gantt-legend-dot actual"></div><span>실행</span></div>' : ""}
        ${hasActual ? '<div class="gantt-legend-item"><div class="gantt-legend-dot" style="background:transparent;border:2px dashed #27ae6f;border-right:none;height:9px;width:22px;border-radius:0"></div><span>실행중(종료일 미입력)</span></div>' : ""}
        <div class="gantt-legend-item"><div class="gantt-legend-dot" style="background:#e03e3e;height:10px;"></div><span>미착수 초과</span></div>
        <div class="gantt-legend-item"><div class="gantt-legend-dot" style="background:rgba(224,62,62,0.12);border:2px dashed #e03e3e;box-sizing:border-box;height:9px;border-radius:4px;"></div><span>늦은 시작</span></div>
      </div>`;

      const parts = [legend];
      const progressCardHtml = buildProgressStatusCardHtml(progressRows);

      for (const grp of groups) {
        if (!grp.rows.length || !grp.start || !grp.end) continue;

        const rStart  = grp.start;
        const rEnd    = grp.end;
        const totalMs = rEnd - rStart;
        const pct     = (d) => Math.max(0, Math.min(100, (d - rStart) / totalMs * 100));

        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const todayInRange = today >= rStart && today <= rEnd;
        const todayPct = todayInRange ? pct(today).toFixed(2) : null;

        const months = [];
        let cur = new Date(rStart.getFullYear(), rStart.getMonth(), 1);
        while (cur <= rEnd) {
          const nextMonth = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
          const isToday = todayInRange
            && today >= cur && today < nextMonth;
          months.push({
            label: `${cur.getMonth()+1}월`,
            left: pct(cur).toFixed(2),
            right: pct(new Date(Math.min(nextMonth, rEnd))).toFixed(2),
            isToday,
          });
          cur = nextMonth;
        }

        // 현재 월 배경 (헤더용 + 바 트랙용)
        const todayMonth = months.find(m => m.isToday);
        const todayMonthBg = todayMonth
          ? `<div class="gantt-today-month-bg" style="left:${todayMonth.left}%;width:${(todayMonth.right - todayMonth.left).toFixed(2)}%"></div>`
          : "";
        const todayHeaderBg = todayMonth
          ? `<div class="gantt-month-header-today-bg" style="left:${todayMonth.left}%;width:${(todayMonth.right - todayMonth.left).toFixed(2)}%"></div>`
          : "";
        // 오늘 날짜 세로선
        const todayLine = todayPct !== null
          ? `<div class="gantt-today-line" style="left:${todayPct}%"></div>`
          : "";

        const monthLines  = months.map(m => `<div class="gantt-month-line" style="left:${m.left}%"></div>`).join("");
        const monthLabels = months.map(m => {
          const center = ((parseFloat(m.left) + parseFloat(m.right)) / 2).toFixed(2);
          return `<div class="gantt-month-label${m.isToday ? " is-today" : ""}" style="left:${center}%">${m.label}</div>`;
        }).join("");

        const tbody = grp.rows.map(row => {
          const name = nameCol ? String(row[nameCol] ?? "").trim() : "";
          const ps = parseToDate(planStart ? row[planStart] : null);
          const pe = parseToDate(planEnd   ? row[planEnd]   : null);
          const as = parseToDate(actStart  ? row[actStart]  : null);
          const ae = parseToDate(actEnd    ? row[actEnd]    : null);

          // 규칙 1·3: 시작~종료 바 (연차 범위 클램핑)
          const makeBar = (s, e, cls) => {
            if (!s || !e || e < s) return "";
            const cs = new Date(Math.max(s.getTime(), rStart.getTime()));
            const ce = new Date(Math.min(e.getTime(), rEnd.getTime()));
            if (ce <= cs) return "";
            const l = pct(cs).toFixed(2);
            const w = Math.max(0.3, pct(ce) - pct(cs)).toFixed(2);
            return `<div class="gantt-bar ${cls}" style="left:${l}%;width:${w}%" title="${dateFmt(s)} ~ ${dateFmt(e)}"></div>`;
          };

          // 미착수 초과: 계획 시작일이 오늘 이전인데 실행 시작일 없음 → 다홍색
          const isOverdue = ps && !as && ps <= today;
          // 늦은 시작: 실행 시작일이 계획 시작일보다 늦음 → 다홍색 점선
          const isLateStart = as && ps && as > ps;

          // 규칙 2: 계획 종료일이 사업기간 이전이면 표시 안 함 (시작일만 이전이면 rStart부터 표시)
          const planBar = makeBar(ps, pe, isOverdue ? "gantt-overdue" : "gantt-plan");

          // 규칙 3·4: 실행 종료일 있으면 실선 바, 시작일만 있으면 점선 바
          const actBar = (as && ae)
            ? makeBar(as, ae, isLateStart ? "gantt-late-start" : "gantt-actual")
            : (as && !ae)
              ? (() => {
                  const cs = new Date(Math.max(as.getTime(), rStart.getTime()));
                  const ce = new Date(Math.min(today.getTime(), rEnd.getTime()));
                  if (ce <= cs) return "";
                  const l = pct(cs).toFixed(2);
                  const w = Math.max(0.3, pct(ce) - pct(cs)).toFixed(2);
                  const cls = isLateStart ? "gantt-late-start-open" : "gantt-actual-open";
                  return `<div class="gantt-bar ${cls}" style="left:${l}%;width:${w}%" title="${dateFmt(as)} ~ 오늘(${dateFmt(today)})"></div>`;
                })()
              : "";

          // 바가 하나도 없으면 행 자체를 표시하지 않음 (사업기간 전 마일스톤 등)
          if (!planBar && !actBar) return null;

          return `<tr>
            <td class="gantt-name-cell">${name}</td>
            <td class="gantt-bar-cell">
              <div class="gantt-bar-track">${todayMonthBg}${monthLines}${planBar}${actBar}${todayLine}</div>
            </td>
          </tr>`;
        }).filter(Boolean).join("");

        if (todayInRange && progressCardHtml) {
          parts.push(progressCardHtml);
        }

        parts.push(`<div class="gantt-year-block">
          <div class="gantt-year-title">${grp.label} <span style="font-weight:500;font-size:0.85rem;color:#6a7b99;">(${dateFmt(rStart)} ~ ${dateFmt(rEnd)})</span></div>
          <table class="gantt-table">
            <thead><tr>
              <th class="gantt-name-col-head">마일스톤</th>
              <th class="gantt-month-header-cell"><div class="gantt-month-header">${todayHeaderBg}${monthLabels}</div></th>
            </tr></thead>
            <tbody>${tbody}</tbody>
          </table>
        </div>`);
      }

      return `<div class="gantt-wrap">${parts.join("")}</div>`;
    }

    function parseNumericLike(value) {
      if (typeof value === "number") return Number.isFinite(value) ? value : Number.NaN;
      const text = String(value ?? "").replace(/,/g, "").trim();
      if (!text) return Number.NaN;
      const num = Number(text);
      return Number.isFinite(num) ? num : Number.NaN;
    }

    function formatBudgetCell(column, value) {
      if (value === null || value === undefined) return "";

      if (column.includes("시작일") || column.includes("종료일")) {
        return toYmd(value);
      }

      const num = parseNumericLike(value);
      if (!Number.isNaN(num)) {
        if (num === 0) return "";
        if (Number.isInteger(num)) return num.toLocaleString("ko-KR");
        return num.toLocaleString("ko-KR", { maximumFractionDigits: 4 });
      }

      const text = String(value);
      return text.trim() === "0" ? "" : text;
    }

    function formatMilestoneCell(column, value) {
      if (value === null || value === undefined) return "";
      if (column.includes("시작일") || column.includes("종료일")) {
        return toYmd(value);
      }
      return value ?? "";
    }

    function milestoneThead(columns) {
      const tokens = columns.map((col) => {
        const parts = String(col).split("|");
        if (parts.length >= 2 && (parts[0] === "계획" || parts[0] === "실행")) {
          return { grouped: true, group: parts[0], label: parts.slice(1).join("|") };
        }
        return { grouped: false, group: "", label: col };
      });

      const hasGrouped = tokens.some((t) => t.grouped);
      if (!hasGrouped) {
        return `<thead><tr>${columns.map((c) => `<th>${c}</th>`).join("")}</tr></thead>`;
      }

      let row1 = "";
      let i = 0;
      while (i < tokens.length) {
        const t = tokens[i];
        if (!t.grouped) {
          row1 += `<th rowspan="2">${t.label}</th>`;
          i += 1;
          continue;
        }

        let span = 1;
        while (i + span < tokens.length && tokens[i + span].grouped && tokens[i + span].group === t.group) {
          span += 1;
        }
        row1 += `<th colspan="${span}">${t.group}</th>`;
        i += span;
      }

      const row2 = tokens
        .filter((t) => t.grouped)
        .map((t) => `<th>${t.label}</th>`)
        .join("");

      return `<thead><tr>${row1}</tr><tr>${row2}</tr></thead>`;
    }

    function metricThead(columns) {
      const tokens = columns.map((col) => {
        const parts = String(col).split("|");
        if (parts.length >= 2 && (parts[0] === "목표" || parts[0] === "달성" || parts[0] === "연구개발달성률")) {
          return { grouped: true, group: parts[0], label: parts.slice(1).join("|") };
        }
        return { grouped: false, group: "", label: col };
      });

      // 텍스트 컬럼은 넓게, 숫자 컬럼은 좁게 — 합계가 정확히 100%가 되도록 계산
      const widePattern = /지표|항목|내용|명칭?|구분|연차|번호|단계|과제/;
      const isWideCol   = (t) => !t.grouped && widePattern.test(t.label);
      const wideCount   = tokens.filter(isWideCol).length || 1;
      const narrowCount = tokens.length - wideCount;
      const wideW       = Math.min(18, Math.floor(40 / wideCount));
      const narrowW     = narrowCount > 0
        ? ((100 - wideW * wideCount) / narrowCount)
        : 0;

      const colgroup = `<colgroup>${tokens.map(t =>
        `<col style="width:${isWideCol(t) ? wideW.toFixed(1) : narrowW.toFixed(2)}%">`
      ).join("")}</colgroup>`;

      const hasGrouped = tokens.some((t) => t.grouped);
      if (!hasGrouped) {
        return colgroup + `<thead><tr>${columns.map((c) => `<th>${c}</th>`).join("")}</tr></thead>`;
      }

      let row1 = "";
      let i = 0;
      while (i < tokens.length) {
        const t = tokens[i];
        if (!t.grouped) {
          row1 += `<th rowspan="2">${t.label}</th>`;
          i += 1;
          continue;
        }

        let span = 1;
        while (i + span < tokens.length && tokens[i + span].grouped && tokens[i + span].group === t.group) {
          span += 1;
        }
        const groupCls = t.group === "목표"
          ? "metric-target"
          : t.group === "달성"
            ? "metric-achieve"
            : "metric-rate-head";
        row1 += `<th colspan="${span}" class="${groupCls}">${t.group}</th>`;
        i += span;
      }

      const row2 = tokens
        .filter((t) => t.grouped)
        .map((t) => {
          const groupCls = t.group === "목표"
            ? "metric-target"
            : t.group === "달성"
              ? "metric-achieve"
              : "metric-rate-head";
          return `<th class="${groupCls}">${t.label}</th>`;
        })
        .join("");

      return colgroup + `<thead><tr>${row1}</tr><tr>${row2}</tr></thead>`;
    }

    function buildSheetTableHtml(rows, sectionName = "") {
      if (!rows.length) {
        return "데이터가 없습니다.";
      }

      const columns = Object.keys(rows[0]);
      const display = rows.slice(0, 120);

      const formatter = sectionName === "사업비"
        ? (col, val) => formatBudgetCell(col, val)
        : sectionName === "마일스톤"
          ? (col, val) => formatMilestoneCell(col, val)
          : (_col, val) => (val ?? "");

      const commaNumberRe = /^-?\d{1,3}(,\d{3})+(\.\d+)?$/;

      const thead = sectionName === "마일스톤"
        ? milestoneThead(columns)
        : sectionName === "정량지표"
          ? metricThead(columns)
          : `<thead><tr>${columns.map((c) => `<th>${c}</th>`).join("")}</tr></thead>`;
      const tbody = `<tbody>${display
        .map((row) => `<tr>${columns.map((c) => {
          const formatted = formatter(c, row[c]);
          const text = String(formatted ?? "").trim();
          const rawNum = parseNumericLike(row[c]);
          const formattedNum = parseNumericLike(formatted);

          let cls = "txt-cell";
          if (commaNumberRe.test(text)) {
            cls = "num-cell";
          } else if (!Number.isNaN(rawNum) || !Number.isNaN(formattedNum)) {
            cls = "mid-cell";
          }

          const clsList = [cls];
          if (sectionName === "정량지표") {
            if (c.startsWith("목표|")) {
              clsList.push("metric-target");
            } else if (c.startsWith("달성|")) {
              clsList.push("metric-achieve");
            }

            if (c.includes("달성률")) {
              clsList.push("metric-rate");
              if (text.includes("달성") && !text.includes("미달성")) {
                clsList.push("status-ok");
              } else if (text.includes("미달성")) {
                clsList.push("status-miss");
              }
            }
          }

          return `<td class="${clsList.join(" ")}">${formatted}</td>`;
        }).join("")}</tr>`)
        .join("")}</tbody>`;
      const classList = ["sheet-table"];
      if (sectionName === "마일스톤" || sectionName === "정량지표") {
        classList.push("multi-head");
      }
      const tableClass = classList.join(" ");
      return `<table class="${tableClass}">${thead}${tbody}</table>`;
    }

    function renderSheetTable(rows, targetWrap, sectionName = "") {
      targetWrap.innerHTML = buildSheetTableHtml(rows, sectionName);
    }

    function renderSheetTabs(project) {
      sheetTabs.innerHTML = "";

      if (!SHEET_TABS.some((t) => t.key === state.activeSheetTab)) {
        state.activeSheetTab = "사업비";
      }

      for (const tab of SHEET_TABS) {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = `subtab-btn${tab.key === state.activeSheetTab ? " active" : ""}`;
        btn.textContent = tab.label;
        btn.addEventListener("click", () => {
          state.activeSheetTab = tab.key;
          renderSheetTabs(project);
          renderActiveSheet(project);
        });
        sheetTabs.appendChild(btn);
      }
    }

    function renderActiveSheet(project) {
      const tab = SHEET_TABS.find((t) => t.key === state.activeSheetTab) || SHEET_TABS[0];
      sheetHint.textContent = tab.hint;

      if (tab.key === "정량지표묶음") {
        const metricRows = project.sheets["정량지표"] || [];
        const metricDetailRows = project.sheets["정량지표상세"] || [];

        activeSheetWrap.style.cssText =
          "overflow:hidden;max-height:none;border:none;border-radius:0;background:transparent;";

        activeSheetWrap.innerHTML = `
          <div class="merged-sheet">
            <div>
              <div class="merged-title">정량지표</div>
              <div class="metric-main-wrap">
                ${buildSheetTableHtml(metricRows, "정량지표")}
              </div>
            </div>
            <div>
              <div class="merged-title">정량지표상세</div>
              <div class="metric-detail-wrap">
                ${buildSheetTableHtml(metricDetailRows, "정량지표상세")}
              </div>
            </div>
          </div>
        `;
        return;
      }

      if (tab.key === "마일스톤") {
        const rows = project.sheets["마일스톤"] || [];
        const budgetRows = project.sheets["사업비"] || [];
        const progressRows = project.sheets["진행상황"] || [];
        activeSheetWrap.style.cssText =
          "overflow:visible;max-height:none;border:none;border-radius:0;background:transparent;";
        activeSheetWrap.innerHTML = buildMilestoneGanttHtml(rows, budgetRows, progressRows);
        return;
      }

      activeSheetWrap.removeAttribute("style");
      const rows = project.sheets[tab.key] || [];
      renderSheetTable(rows, activeSheetWrap, tab.key);
    }

    function renderDashboard() {
      if (state.view === "allOverview") {
        renderAllOverview();
        return;
      }
      if (state.view === "budgetLedger") {
        renderBudgetLedger();
        return;
      }
      const project = getActiveProject();
      if (!project) return;

      renderOverview(project);
      renderSheetTabs(project);
      renderActiveSheet(project);
    }

    // XLSX.js가 버린 t='d'(ISO 날짜) 셀을 ZIP→XML 직접 파싱으로 복원
    async function patchIsoDateCells(wb, arrayBuffer) {
      try {
        const u8 = new Uint8Array(arrayBuffer);

        // ZIP Central Directory 스캔 — 중복 항목이 있을 경우 마지막(최신) 항목 사용
        const entryMap = new Map();
        let eocd = -1;
        for (let j = u8.length - 22; j >= Math.max(0, u8.length - 65558); j--) {
          if (u8[j]===0x50 && u8[j+1]===0x4B && u8[j+2]===0x05 && u8[j+3]===0x06) { eocd = j; break; }
        }
        if (eocd >= 0) {
          const cdOff = u8[eocd+16]|u8[eocd+17]<<8|u8[eocd+18]<<16|u8[eocd+19]<<24;
          const cdSz  = u8[eocd+12]|u8[eocd+13]<<8|u8[eocd+14]<<16|u8[eocd+15]<<24;
          let pos = cdOff;
          while (pos < cdOff + cdSz && pos + 46 <= u8.length) {
            if (!(u8[pos]===0x50&&u8[pos+1]===0x4B&&u8[pos+2]===0x01&&u8[pos+3]===0x02)) break;
            const method  = u8[pos+10]|u8[pos+11]<<8;
            const csize   = u8[pos+20]|u8[pos+21]<<8|u8[pos+22]<<16|u8[pos+23]<<24;
            const fnLen   = u8[pos+28]|u8[pos+29]<<8;
            const exLen   = u8[pos+30]|u8[pos+31]<<8;
            const cmLen   = u8[pos+32]|u8[pos+33]<<8;
            const lhOff   = u8[pos+42]|u8[pos+43]<<8|u8[pos+44]<<16|u8[pos+45]<<24;
            const fname   = new TextDecoder().decode(u8.subarray(pos+46, pos+46+fnLen));
            if (/^xl\/worksheets\/sheet\d+\.xml$/.test(fname)) {
              const lhFnLen = u8[lhOff+26]|u8[lhOff+27]<<8;
              const lhExLen = u8[lhOff+28]|u8[lhOff+29]<<8;
              const dStart  = lhOff + 30 + lhFnLen + lhExLen;
              entryMap.set(fname, { fname, method, data: u8.subarray(dStart, dStart + csize) });
            }
            pos += 46 + fnLen + exLen + cmLen;
          }
        }
        const entries = [...entryMap.values()];

        // 엑셀 날짜 직렬 변환 (Dec 30, 1899 기준)
        const EPOCH = new Date(1899, 11, 30).getTime();
        const toSerial = d => (d.getTime() - EPOCH) / 86400000;

        for (const entry of entries) {
          // 시트 인덱스 → 시트명 매핑 (sheet1.xml → SheetNames[0])
          const idx = parseInt(entry.fname.match(/sheet(\d+)\.xml/)[1]) - 1;
          const sname = wb.SheetNames[idx];
          if (!sname) continue;
          const ws = wb.Sheets[sname];
          if (!ws) continue;

          // 압축 해제
          let xml;
          if (entry.method === 0) {
            xml = new TextDecoder().decode(entry.data);
          } else {
            const ds = new DecompressionStream('deflate-raw');
            const w = ds.writable.getWriter();
            const r = ds.readable.getReader();
            w.write(entry.data); w.close();
            const chunks = [];
            while (true) {
              const { done, value } = await r.read();
              if (done) break;
              chunks.push(value);
            }
            const buf = new Uint8Array(chunks.reduce((n,c)=>n+c.length,0));
            let off = 0;
            for (const c of chunks) { buf.set(c, off); off += c.length; }
            xml = new TextDecoder().decode(buf);
          }

          // t="d" 셀 추출 → 워크시트에 추가
          if (sname === "마일스톤") {
            const row3match = xml.match(/<row[^>]*r="3"[^>]*>[\s\S]*?<\/row>/);
            console.log("[ROW3-XML]", sname, row3match ? row3match[0] : "row3 not found");
          }
          const re = /<c\b([^>]*)t="d"([^>]*)>([\s\S]*?)<\/c>/g;
          let m;
          let patchCount = 0;
          while ((m = re.exec(xml)) !== null) {
            const attrs = m[1] + m[2];
            const inner = m[3];
            const ref   = (attrs.match(/\br="([A-Z]+\d+)"/) || [])[1];
            const iso   = (inner.match(/<v>([^<]+)<\/v>/) || [])[1];
            console.log("[PATCH-CELL]", sname, ref, iso, "exists?", !!ws[ref]);
            if (!ref || !iso || ws[ref]) continue;
            const d = new Date(iso);
            if (isNaN(d.getTime())) continue;
            const local = new Date(d.getFullYear(), d.getMonth(), d.getDate());
            ws[ref] = { t: 'n', v: toSerial(local) };
            patchCount++;
          }
          if (patchCount > 0) console.log("[PATCH]", sname, patchCount, "cells patched");
        }
      } catch (e) {
        console.warn('patchIsoDateCells:', e);
      }
    }

    async function parseFiles(fileEntries) {
      if (typeof XLSX === "undefined") {
        throw new Error("XLSX 라이브러리를 불러오지 못했습니다. 네트워크/CDN 접근을 확인해주세요.");
      }

      const projects = [];
      for (const fileEntry of fileEntries) {
        const name = fileEntry.name.replace(/\.[^.]+$/, "");
        const data = await fileEntry.arrayBuffer();
        const wb = XLSX.read(data, { type: "array" });
        // XLSX.js v0.18.x는 t='d'(ISO 날짜) 셀을 파싱 시 버림 → ZIP에서 직접 복원
        await patchIsoDateCells(wb, data);
        const sheets = {};
        wb.SheetNames.forEach((sheetName) => {
          if (sheetName === "마일스톤") {
            sheets[sheetName] = parseMilestoneRows(wb, sheetName);
          } else if (sheetName === "정량지표") {
            sheets[sheetName] = parseMetricRows(wb, sheetName);
          } else if (sheetName === "진행상황") {
            sheets[sheetName] = sheetRows(wb, sheetName, 1);
          } else {
            const combined = parseCombinedSheet(wb, sheetName);
            if (combined) {
              // 개요 패널용 virtual key — 아직 없을 때만 설정
              if (!sheets["개요"]) sheets["개요"] = combined.overviewRows;
              // 사업비(간트 기간 참조)용 virtual key — 아직 없을 때만 설정
              if (!sheets["사업비"]) sheets["사업비"] = combined.budgetRows;
              // 원래 시트명으로는 사업비(budget) 부분만 저장 (탭 테이블용)
              // 단, sheetName이 "개요"인 경우 "개요" 키를 덮어쓰지 않도록 주의
              if (sheetName !== "개요") {
                sheets[sheetName] = combined.budgetRows;
              }
              // sheetName === "개요"인 경우: sheets["개요"]는 overviewRows,
              // sheets["사업비"]는 budgetRows 로 이미 올바르게 설정됨
            } else {
              sheets[sheetName] = sheetRows(wb, sheetName, 1);
            }
          }
        });
        projects.push({ name, sheets });
      }
      return projects;
    }

    async function discoverProjectExcelPaths() {
      const resp = await fetch("projects/");
      if (!resp.ok) throw new Error("projects 폴더 접근 실패");

      const html = await resp.text();
      const doc = new DOMParser().parseFromString(html, "text/html");
      const anchors = Array.from(doc.querySelectorAll("a[href]"));
      const paths = anchors
        .map((a) => a.getAttribute("href") || "")
        .filter((href) => href && !href.startsWith("?") && !href.startsWith("#"))
        .map((href) => decodeURIComponent(href))
        .filter((href) => !href.startsWith("../") && isExcelPath(href))
        .map((href) => `projects/${href.replace(/^\.\//, "")}`);

      return Array.from(new Set(paths));
    }

    async function loadFilesIntoDashboard(files, sourceLabel) {
      state.projects = await parseFiles(files);
      if (!state.projects.length) {
        setMessage("불러온 엑셀 데이터가 없습니다.");
        return;
      }

      if (!state.projects.some((p) => p.name === state.activeProject)) {
        state.activeProject = state.projects[0].name;
      }

      renderProjectMenu();
      renderDashboard();
      setMessage(`${sourceLabel}에서 ${state.projects.length}개 과제를 불러왔습니다.`);
    }

    async function loadProjectFolder(silent = false) {
      try {
        const paths = await discoverProjectExcelPaths();
        if (!paths.length) throw new Error("projects 폴더에서 엑셀 파일을 찾지 못했습니다.");

        const files = [];
        for (const path of paths) {
          const resp = await fetch(path, { cache: "no-store" });
          if (!resp.ok) throw new Error(`${path} 로딩 실패`);
          const blob = await resp.blob();
          files.push(new File([blob], path.split("/").pop(), { type: blob.type }));
        }

        await loadFilesIntoDashboard(files, "projects 폴더");
      } catch (err) {
        if (!silent) {
          setMessage(`projects 자동 로딩 실패: ${err.message}`);
        }
      }
    }

    async function loadFromInput() {
      if (!fileInput.files || !fileInput.files.length) {
        setMessage("업로드할 엑셀 파일을 선택해주세요.");
        return;
      }
      const files = Array.from(fileInput.files);
      await loadFilesIntoDashboard(files, "업로드 파일");
    }

    async function checkReloadToken() {
      try {
        const resp = await fetch(`.watch/reload-token.json?t=${Date.now()}`, { cache: "no-store" });
        if (!resp.ok) return;
        const payload = await resp.json();
        const token = String(payload.token || "");
        if (!token) return;

        if (!lastReloadToken) {
          lastReloadToken = token;
          return;
        }

        if (token !== lastReloadToken) {
          location.reload();
        }
      } catch (_) {
        // Watcher not running or token file not available.
      }
    }

    function bindEvents() {
      allOverviewBtn.addEventListener("click", () => {
        if (state.view === "allOverview") {
          switchView("project");
          renderDashboard();
        } else {
          switchView("allOverview");
        }
      });

      budgetLedgerBtn.addEventListener("click", () => {
        if (state.view === "budgetLedger") {
          switchView("project");
          renderDashboard();
        } else {
          switchView("budgetLedger");
        }
      });

      uploadBtn.addEventListener("click", () => fileInput.click());
      fileInput.addEventListener("change", () => {
        loadFromInput().catch((err) => setMessage(`파일 읽기 오류: ${err.message}`));
      });

      refreshBtn.addEventListener("click", () => {
        if (location.protocol === "file:") {
          pickAndLoadDirectory().catch((err) => setMessage(`새로고침 오류: ${err.message}`));
        } else {
          loadProjectFolder(false).catch((err) => setMessage(`새로고침 오류: ${err.message}`));
        }
      });

      reloadProjectBtn.addEventListener("click", () => {
        if (location.protocol === "file:") {
          pickAndLoadDirectory().catch((err) => setMessage(`폴더 로딩 오류: ${err.message}`));
        } else {
          loadProjectFolder(false).catch((err) => setMessage(`폴더 로딩 오류: ${err.message}`));
        }
      });
    }

    function init() {
      baseDate.value = formatDateToday();
      bindEvents();
      checkReloadToken();
      setInterval(checkReloadToken, 3000);
      if (location.protocol === "file:") {
        // 서버 리다이렉트(800ms)가 없으면 File System API로 자동 로드 시도
        setTimeout(async () => {
          if ('showDirectoryPicker' in window) {
            const loaded = await autoLoadDirectory();
            if (!loaded) {
              const saved = await getSavedHandle();
              reloadProjectBtn.textContent = saved ? "📂 데이터 불러오기" : "📁 프로젝트 폴더 연결하기";
              reloadProjectBtn.classList.add("active");
              reloadProjectBtn.focus();
              setMessage(
                saved
                  ? "왼쪽의 '📂 데이터 불러오기' 버튼을 눌러 프로젝트 폴더 접근 권한을 다시 허용해주세요."
                  : "브라우저 보안 정책상 폴더를 자동으로 읽을 수 없습니다. 왼쪽의 '📁 프로젝트 폴더 연결하기' 버튼을 눌러 projects 폴더를 선택해주세요. (또는 open-dashboard.bat 실행 권장)"
              );
            }
          } else {
            reloadProjectBtn.style.display = "none";
            setMessage("'엑셀 직접 업로드' 버튼으로 Excel 파일을 선택해주세요.");
          }
        }, 900);
      } else {
        loadProjectFolder(false).catch((err) => {
          setMessage(`자동 로딩에 실패했습니다: ${err.message}`);
        });
      }
    }

    init();
  