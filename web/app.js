const state = {
  view: "in-use",
  keyword: "",
  rows: [],
  selectedId: "",
  sort: {
    key: "",
    direction: "asc"
  },
  options: {
    colleagues: [],
    departments: [],
    models: [],
    employeeTypes: []
  },
  editor: {
    kind: "",
    mode: "create",
    item: null
  }
};

const tableHead = document.getElementById("table-head");
const tableBody = document.getElementById("table-body");
const statsGrid = document.getElementById("stats-grid");
const recentList = document.getElementById("recent-list");
const listCard = document.querySelector(".list-card");
const sideStack = document.querySelector(".side-stack");
const searchInput = document.getElementById("search-input");
const currentUser = document.getElementById("current-user");
const currentRole = document.getElementById("current-role");
const workspaceTitle = document.getElementById("workspace-title");
const listSummary = document.getElementById("list-summary");
const addBtn = document.getElementById("add-btn");
const editBtn = document.getElementById("edit-btn");
const deleteBtn = document.getElementById("delete-btn");
const mailCheckBtn = document.getElementById("mail-check-btn");
const detailBtn = document.getElementById("detail-btn");
const refreshBtn = document.getElementById("refresh-btn");
const detailPanelTitle = document.getElementById("detail-panel-title");
const detailPanelContent = document.getElementById("detail-panel-content");
const detailModal = document.getElementById("detail-modal");
const detailTitle = document.getElementById("detail-title");
const detailContent = document.getElementById("detail-content");
const editorModal = document.getElementById("editor-modal");
const editorForm = document.getElementById("editor-form");
const editorTitle = document.getElementById("editor-title");
const editorFields = document.getElementById("editor-fields");
const editorAlert = document.getElementById("editor-alert");
const submitBtn = document.getElementById("submit-btn");
const mailModal = document.getElementById("mail-modal");
const mailAlert = document.getElementById("mail-alert");
const mailHint = document.getElementById("mail-hint");
const mailList = document.getElementById("mail-list");
const mailSkipped = document.getElementById("mail-skipped");
const mailResult = document.getElementById("mail-result");

let workspaceResizeObserver = null;

const viewConfig = {
  "in-use": {
    title: "在用电脑列表",
    searchPlaceholder: "搜索电脑名称、序列号、归属人、备注等信息",
    addLabel: "新增电脑",
    editLabel: "编辑电脑",
    deleteLabel: "转入库存",
    detailType: "computer",
    sortableColumns: ["model", "ownerName"],
    columns: [
      ["computerName", "电脑名称"],
      ["assetNumber", "固定资产号"],
      ["model", "型号"],
      ["macAddress", "MAC 地址"],
      ["ownerName", "归属人"],
      ["updatedAt", "更新时间"]
    ]
  },
  inventory: {
    title: "库存电脑列表",
    searchPlaceholder: "搜索库存电脑名称、序列号、型号、备注等信息",
    addLabel: "新增库存电脑",
    editLabel: "编辑库存电脑",
    deleteLabel: "删除库存电脑",
    detailType: "computer",
    sortableColumns: ["model"],
    columns: [
      ["computerName", "电脑名称"],
      ["assetNumber", "固定资产号"],
      ["model", "型号"],
      ["macAddress", "MAC 地址"],
      ["remark", "备注"],
      ["updatedAt", "更新时间"]
    ]
  },
  people: {
    title: "人员名单",
    searchPlaceholder: "搜索中文名、拼音、邮箱、部门、Mentor 等信息",
    addLabel: "新增人员",
    editLabel: "编辑人员",
    deleteLabel: "删除人员",
    detailType: "person",
    sortableColumns: [],
    columns: [
      ["displayName", "中文名"],
      ["pinyin", "拼音"],
      ["email", "邮箱"],
      ["department", "部门"],
      ["employeeType", "员工类型"],
      ["mentorName", "Mentor"],
      ["ownedCount", "名下电脑"]
    ]
  }
};

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

async function fetchJson(url, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: {
      Accept: "application/json",
      ...(options.headers || {})
    }
  });

  let payload = {};
  try {
    payload = await response.json();
  } catch {
    payload = {};
  }

  if (!response.ok) {
    throw new Error(payload.error || "请求失败，请稍后重试。");
  }

  return payload;
}

function formatTag(value, kind = "na") {
  return `<span class="tag ${kind}">${escapeHtml(value)}</span>`;
}

function normalizeEmployeeType(value) {
  const text = String(value || "").trim();
  if (!text) return "N/A";
  if (text.includes("实习") || text.includes("瀹炰範")) return "实习生";
  if (text.includes("正式") || text.includes("姝ｅ紡")) return "正式员工";
  return text;
}

function normalizeInventoryText(value) {
  const text = String(value || "").trim();
  if (!text) return "N/A";
  if (text === "N/A") return "N/A";
  if (text.includes("库存") || text.includes("搴撳瓨")) return "库存";
  if (text.includes("未知") || text.includes("鏈煡")) return "未知人员";
  return text;
}

function formatPinyinDisplay(value) {
  return String(value || "")
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase())
    .join(" ");
}

function formatMacAddressDisplay(value) {
  const raw = String(value ?? "").trim();
  if (!raw || raw.toUpperCase() === "N/A") {
    return "N/A";
  }

  const normalized = raw.replace(/[^0-9a-fA-F]/g, "").toUpperCase();
  if (normalized.length !== 12) {
    return raw.toUpperCase();
  }

  return normalized.match(/.{1,2}/g).join("-");
}

function getSortValue(row, key) {
  if (key === "ownerName") {
    return row.ownerId ? normalizeInventoryText(row.ownerName) : "库存";
  }
  if (key === "employeeType") {
    return normalizeEmployeeType(row.employeeType);
  }
  if (key === "pinyin") {
    return formatPinyinDisplay(row.pinyin);
  }
  return String(row[key] || "").trim();
}

function sortRows(rows) {
  const { key, direction } = state.sort;
  if (!key) return rows;

  const sorted = [...rows].sort((left, right) => {
    const leftValue = getSortValue(left, key);
    const rightValue = getSortValue(right, key);
    const leftEmpty = !leftValue || leftValue === "N/A";
    const rightEmpty = !rightValue || rightValue === "N/A";

    if (leftEmpty && rightEmpty) return 0;
    if (leftEmpty) return 1;
    if (rightEmpty) return -1;

    return leftValue.localeCompare(rightValue, "zh-Hans-CN", {
      numeric: true,
      sensitivity: "base"
    });
  });

  return direction === "desc" ? sorted.reverse() : sorted;
}

function renderCell(key, row) {
  const value = row[key] ?? "";

  if (key === "computerName" || key === "displayName") {
    return `<span class="cell-strong">${escapeHtml(value)}</span>`;
  }

  if (key === "pinyin") {
    return escapeHtml(formatPinyinDisplay(value));
  }

  if (key === "assetNumber") {
    return value === "N/A" ? formatTag("N/A", "na") : escapeHtml(value);
  }

  if (key === "macAddress") {
    const macAddress = formatMacAddressDisplay(value);
    return macAddress === "N/A"
      ? formatTag("N/A", "na")
      : `<span class="mac-address">${escapeHtml(macAddress)}</span>`;
  }

  if (key === "ownerName") {
    return row.ownerId
      ? `<span class="owner-name">${escapeHtml(normalizeInventoryText(value))}</span>`
      : formatTag("库存", "inventory");
  }

  if (key === "employeeType") {
    const employeeType = normalizeEmployeeType(value);
    return employeeType === "实习生" ? formatTag(employeeType, "intern") : escapeHtml(employeeType);
  }

  if (key === "mentorName") {
    const mentorName = normalizeInventoryText(value);
    return mentorName === "N/A" ? formatTag("N/A", "na") : escapeHtml(mentorName);
  }

  if (value === "" || value === null || value === undefined) {
    return formatTag("N/A", "na");
  }

  return escapeHtml(value);
}

function getSelectedRow() {
  return state.rows.find((row) => row.id === state.selectedId) || null;
}

function setEditorAlert(message = "") {
  if (!message) {
    editorAlert.textContent = "";
    editorAlert.classList.add("hidden");
    return;
  }
  editorAlert.textContent = message;
  editorAlert.classList.remove("hidden");
}

function setMailAlert(message = "") {
  if (!message) {
    mailAlert.textContent = "";
    mailAlert.classList.add("hidden");
    return;
  }
  mailAlert.textContent = message;
  mailAlert.classList.remove("hidden");
}

function syncWorkspaceHeights() {
  if (!listCard || !sideStack) return;

  if (window.innerWidth <= 1180) {
    listCard.style.height = "";
    return;
  }

  const sideHeight = Math.ceil(sideStack.getBoundingClientRect().height);
  if (sideHeight > 0) {
    listCard.style.height = `${sideHeight}px`;
  }
}

function buildDataList(id, values) {
  if (!values.length) return "";
  return `
    <datalist id="${id}">
      ${values.map((value) => `<option value="${escapeHtml(value)}"></option>`).join("")}
    </datalist>
  `;
}

function normalizeLookupText(value) {
  return String(value || "").trim().toLowerCase();
}

function buildOwnerLookupItems() {
  return state.options.colleagues.map((item) => {
    const searchValues = [
      item.displayName,
      item.pinyin,
      item.email,
      item.department,
      `${item.displayName} (${item.department})`,
      item.label
    ].filter(Boolean);

    return {
      id: item.id,
      displayName: item.displayName,
      label: `${item.displayName} (${item.department})`,
      searchValues,
      normalizedSet: new Set(searchValues.map((value) => normalizeLookupText(value)))
    };
  });
}

function searchOwnerCandidates(rawValue) {
  const normalized = normalizeLookupText(rawValue);
  if (!normalized || normalized === "库存") {
    return [];
  }

  return buildOwnerLookupItems().filter((item) =>
    item.searchValues.some((value) => normalizeLookupText(value).includes(normalized))
  );
}

function resolveOwnerInputToId(rawValue, selectedId = "") {
  const normalized = normalizeLookupText(rawValue);
  if (!normalized || normalized === "库存") {
    return { id: "", matched: true };
  }

  const lookupItems = buildOwnerLookupItems();
  const selectedMatch = selectedId ? lookupItems.find((item) => item.id === selectedId) : null;
  if (selectedMatch && selectedMatch.normalizedSet.has(normalized)) {
    return { id: selectedMatch.id, matched: true };
  }

  const exactMatches = lookupItems.filter((item) => item.normalizedSet.has(normalized));
  if (exactMatches.length === 1) {
    return { id: exactMatches[0].id, matched: true };
  }

  const prefixMatches = lookupItems.filter((item) =>
    item.searchValues.some((value) => normalizeLookupText(value).startsWith(normalized))
  );
  if (prefixMatches.length === 1) {
    return { id: prefixMatches[0].id, matched: true };
  }

  const fuzzyMatches = lookupItems.filter((item) =>
    item.searchValues.some((value) => normalizeLookupText(value).includes(normalized))
  );
  if (fuzzyMatches.length === 1) {
    return { id: fuzzyMatches[0].id, matched: true };
  }

  return { id: "", matched: false };
}

function buildOwnerSuggestions(selectedId = "") {
  const values = ["库存"];
  const seen = new Set(values.map((value) => normalizeLookupText(value)));

  state.options.colleagues.forEach((item) => {
    const suggestionValues = [item.displayName, `${item.displayName} (${item.department})`, item.email, item.pinyin];
    suggestionValues.forEach((value) => {
      const normalized = normalizeLookupText(value);
      if (!normalized || seen.has(normalized)) return;
      seen.add(normalized);
      values.push(value);
    });
  });

  const selectedOwner = state.options.colleagues.find((item) => item.id === selectedId);
  if (selectedOwner) {
    const selectedLabel = `${selectedOwner.displayName} (${selectedOwner.department})`;
    const normalized = normalizeLookupText(selectedLabel);
    if (!seen.has(normalized)) {
      values.unshift(selectedLabel);
    }
  }

  return values;
}

function getOwnerInputValue(selectedId = "") {
  if (!selectedId) return "库存";
  const selectedOwner = state.options.colleagues.find((item) => item.id === selectedId);
  return selectedOwner ? `${selectedOwner.displayName} (${selectedOwner.department})` : "";
}

function createField({ name, label, type = "text", value = "", options = [], note = "", required = false, wide = false, list = "", disabled = false }) {
  if (type === "textarea") {
    return `
      <div class="field-wrap ${wide ? "wide" : ""}">
        <label for="${name}">${label}${required ? " *" : ""}</label>
        <textarea id="${name}" name="${name}" ${required ? "required" : ""} ${disabled ? "disabled" : ""}>${escapeHtml(value)}</textarea>
        ${note ? `<div class="field-note">${escapeHtml(note)}</div>` : ""}
      </div>
    `;
  }

  if (type === "select") {
    return `
      <div class="field-wrap ${wide ? "wide" : ""}">
        <label for="${name}">${label}${required ? " *" : ""}</label>
        <select id="${name}" name="${name}" ${required ? "required" : ""} ${disabled ? "disabled" : ""}>
          ${options
            .map(
              (option) => `
                <option value="${escapeHtml(option.value)}" ${String(option.value) === String(value) ? "selected" : ""}>
                  ${escapeHtml(option.label)}
                </option>
              `
            )
            .join("")}
        </select>
        ${note ? `<div class="field-note">${escapeHtml(note)}</div>` : ""}
      </div>
    `;
  }

  return `
    <div class="field-wrap ${wide ? "wide" : ""}">
      <label for="${name}">${label}${required ? " *" : ""}</label>
      <input id="${name}" name="${name}" type="${type}" value="${escapeHtml(value)}" ${required ? "required" : ""} ${list ? `list="${list}"` : ""} ${disabled ? "disabled" : ""} />
      ${note ? `<div class="field-note">${escapeHtml(note)}</div>` : ""}
    </div>
  `;
}

function createOwnerAutocompleteField(selectedId = "") {
  return `
    <div class="field-wrap">
      <label for="ownerInput">归属人</label>
      <input id="ownerInput" name="ownerInput" type="text" value="${escapeHtml(getOwnerInputValue(selectedId))}" />
      <div id="owner-suggestions" class="owner-suggestions hidden"></div>
      <div class="field-note">输入姓名、拼音或邮箱自动匹配；填写“库存”或留空表示仍在库存。</div>
    </div>
  `;
}

function renderTable() {
  const config = viewConfig[state.view];

  tableHead.innerHTML = `<tr>${config.columns
    .map(([key, label]) => {
      if (!config.sortableColumns.includes(key)) {
        return `<th>${label}</th>`;
      }

      const isActive = state.sort.key === key;
      const indicator = isActive ? (state.sort.direction === "asc" ? "^" : "v") : "<>";
      return `<th><button type="button" class="sort-btn ${isActive ? "active" : ""}" data-sort-key="${key}">${label}<span class="sort-indicator">${indicator}</span></button></th>`;
    })
    .join("")}</tr>`;

  if (!state.rows.length) {
    tableBody.innerHTML = `<tr><td class="empty-row" colspan="${config.columns.length}">当前没有可显示的数据。</td></tr>`;
    return;
  }

  tableBody.innerHTML = state.rows
    .map((row) => {
      const selected = row.id === state.selectedId ? "selected" : "";
      return `
        <tr class="${selected}" data-id="${escapeHtml(row.id)}">
          ${config.columns.map(([key]) => `<td>${renderCell(key, row)}</td>`).join("")}
        </tr>
      `;
    })
    .join("");
}

function renderStats(cards) {
  statsGrid.innerHTML = cards
    .map(
      (card) => `
        <article class="stat-card stat-${escapeHtml(card.tone || "blue")}">
          <div class="stat-label">${escapeHtml(card.label)}</div>
          <div class="stat-value">${escapeHtml(card.value)}</div>
        </article>
      `
    )
    .join("");
}

function renderRecentList(items) {
  if (!items.length) {
    recentList.innerHTML = `<article class="recent-item"><h3>暂无更新</h3><div class="recent-meta">最近暂无电脑变更记录。</div></article>`;
    return;
  }

  recentList.innerHTML = items
    .map(
      (item) => `
        <article class="recent-item">
          <h3>${escapeHtml(item.computerName || "N/A")}</h3>
          <div class="recent-meta">归属人：${escapeHtml(normalizeInventoryText(item.ownerName || "库存"))}</div>
          <div class="recent-meta">型号：${escapeHtml(item.model || "N/A")}</div>
          <div class="recent-meta">更新时间：${escapeHtml(item.updatedAt || "N/A")}</div>
        </article>
      `
    )
    .join("");
}

function renderDetailFields(fields) {
  return `
    <div class="detail-grid">
      ${fields
        .map(
          ({ label, value, wide = false }) => `
            <div class="detail-field ${wide ? "wide" : ""}">
              <label>${escapeHtml(label)}</label>
              <p>${escapeHtml(value || "N/A")}</p>
            </div>
          `
        )
        .join("")}
    </div>
  `;
}

function renderComputerDetail(item) {
  const history = (item.ownerHistory || []).length
    ? `
        <div class="detail-field wide">
          <label>归属历史</label>
          <div class="history-list">
            ${item.ownerHistory
              .map(
                (entry) => `
                  <div class="history-item">
                    <strong>${escapeHtml(entry.changed_at || "N/A")}</strong><br />
                    ${escapeHtml(normalizeInventoryText(entry.old_owner_name || "库存"))} -> ${escapeHtml(normalizeInventoryText(entry.new_owner_name || "库存"))}
                  </div>
                `
              )
              .join("")}
          </div>
        </div>
      `
    : `
        <div class="detail-field wide">
          <label>归属历史</label>
          <p>暂无归属变更记录。</p>
        </div>
      `;

  return `
    ${renderDetailFields([
      { label: "电脑名称", value: item.computerName },
      { label: "序列号", value: item.serialNumber },
      { label: "固定资产号", value: item.assetNumber },
      { label: "型号", value: item.model || "N/A" },
      { label: "MAC 地址", value: formatMacAddressDisplay(item.macAddress) },
      { label: "归属人", value: normalizeInventoryText(item.ownerName) },
      { label: "更新时间", value: item.updatedAt },
      { label: "备注", value: item.remark || "N/A", wide: true }
    ])}
    ${history}
  `;
}

function renderPersonDetail(item) {
  return renderDetailFields([
    { label: "中文名", value: item.displayName },
    { label: "拼音", value: formatPinyinDisplay(item.pinyin) },
    { label: "邮箱", value: item.email },
    { label: "部门", value: item.department },
    { label: "员工类型", value: normalizeEmployeeType(item.employeeType) },
    { label: "Mentor", value: normalizeInventoryText(item.mentorName) },
    { label: "名下电脑", value: String(item.ownedCount) }
  ]);
}

async function refreshDetailPanel() {
  if (!state.selectedId) {
    detailPanelTitle.textContent = "未选择记录";
    detailPanelContent.className = "detail-content detail-panel-empty";
    detailPanelContent.textContent = "选择一条记录后，这里会显示详细信息和归属历史。";
    syncWorkspaceHeights();
    return;
  }

  const type = viewConfig[state.view].detailType;
  const data = await fetchJson(`/api/details?type=${type}&id=${encodeURIComponent(state.selectedId)}`);

  detailPanelTitle.textContent = type === "computer" ? data.computerName : data.displayName;
  detailPanelContent.className = "detail-content";
  detailPanelContent.innerHTML = type === "computer" ? renderComputerDetail(data) : renderPersonDetail(data);
  syncWorkspaceHeights();
}

async function openDetailModal() {
  const row = getSelectedRow();
  if (!row) {
    window.alert("请先选择一条记录。");
    return;
  }

  const type = viewConfig[state.view].detailType;
  const data = await fetchJson(`/api/details?type=${type}&id=${encodeURIComponent(row.id)}`);

  detailTitle.textContent = type === "computer" ? data.computerName : data.displayName;
  detailContent.innerHTML = type === "computer" ? renderComputerDetail(data) : renderPersonDetail(data);
  detailModal.showModal();
}

function buildOwnerOptions(selectedId = "") {
  return [
    { value: "", label: "库存" },
    ...state.options.colleagues.map((item) => ({
      value: item.id,
      label: item.id === selectedId ? item.label : `${item.displayName} (${item.department})`
    }))
  ];
}

function buildMentorOptions(excludeId = "") {
  const options = state.options.colleagues.filter(
    (item) => item.id !== excludeId && normalizeEmployeeType(item.employeeType) === "正式员工"
  );

  return [{ value: "", label: "请选择 Mentor" }].concat(
    options.map((item) => ({
      value: item.id,
      label: `${item.displayName} (${item.department})`
    }))
  );
}

function openComputerEditor(mode, item = null) {
  const isInventory = state.view === "inventory";

  state.editor = {
    kind: "computer",
    mode,
    item
  };

  setEditorAlert("");
  editorTitle.textContent = mode === "edit" ? (isInventory ? "编辑库存电脑" : "编辑电脑") : (isInventory ? "新增库存电脑" : "新增电脑");
  submitBtn.textContent = "保存";

  editorFields.innerHTML = `
    <div class="form-grid">
      ${createField({ name: "computerName", label: "电脑名称", value: item?.computerName || "", required: true })}
      ${createField({ name: "serialNumber", label: "序列号", value: item?.serialNumber || "", required: true })}
      ${createField({ name: "assetNumber", label: "固定资产号", value: item?.assetNumber === "N/A" ? "" : item?.assetNumber || "", note: "留空或填写“暂无”都会按 N/A 展示。" })}
      ${createField({
        name: "model",
        label: "型号",
        value: item?.model || "",
        list: "model-list"
      })}
      ${createField({ name: "macAddress", label: "MAC 地址", value: item?.macAddress === "N/A" ? "" : item?.macAddress || "" })}
      ${createOwnerAutocompleteField(item?.ownerId || "")}
      ${createField({ name: "remark", label: "备注", type: "textarea", value: item?.remark || "", wide: true })}
      ${createField({
        name: "authorizationPassword",
        label: "授权密码",
        type: "password",
        value: "",
        required: mode === "edit",
        wide: true
      })}
    </div>
    <input id="ownerId" name="ownerId" type="hidden" value="${escapeHtml(item?.ownerId || "")}" />
    ${buildDataList("model-list", state.options.models)}
  `;

  bindOwnerField();
  editorModal.showModal();
}

function bindOwnerField() {
  const ownerInput = document.getElementById("ownerInput");
  const ownerIdField = document.getElementById("ownerId");
  const ownerSuggestions = document.getElementById("owner-suggestions");
  if (!ownerInput || !ownerIdField || !ownerSuggestions) return;

  const hideSuggestions = () => {
    ownerSuggestions.innerHTML = "";
    ownerSuggestions.classList.add("hidden");
  };

  const selectOwner = (candidate) => {
    ownerInput.value = candidate ? candidate.label : "库存";
    ownerIdField.value = candidate ? candidate.id : "";
    hideSuggestions();
  };

  const renderSuggestions = () => {
    const candidates = searchOwnerCandidates(ownerInput.value).slice(0, 8);
    if (!candidates.length) {
      hideSuggestions();
      return;
    }

    ownerSuggestions.innerHTML = candidates
      .map(
        (candidate) => `
          <button type="button" class="owner-suggestion-item" data-owner-id="${escapeHtml(candidate.id)}">
            <span class="owner-suggestion-name">${escapeHtml(candidate.displayName)}</span>
            <span class="owner-suggestion-meta">${escapeHtml(candidate.searchValues.filter(Boolean).slice(0, 3).join(" / "))}</span>
          </button>
        `
      )
      .join("");
    ownerSuggestions.classList.remove("hidden");
  };

  const syncOwner = () => {
    const result = resolveOwnerInputToId(ownerInput.value, ownerIdField.value);
    ownerIdField.value = result.matched ? result.id : "";
    renderSuggestions();
  };

  ownerInput.addEventListener("input", syncOwner);
  ownerInput.addEventListener("change", syncOwner);
  ownerInput.addEventListener("focus", () => {
    if (normalizeLookupText(ownerInput.value) === "库存") {
      ownerInput.value = "";
      ownerIdField.value = "";
    }
    renderSuggestions();
  });
  ownerInput.addEventListener("blur", () => {
    setTimeout(() => {
      if (!normalizeLookupText(ownerInput.value)) {
        ownerInput.value = "库存";
        ownerIdField.value = "";
      }
      hideSuggestions();
    }, 120);
  });

  ownerSuggestions.addEventListener("click", (event) => {
    const button = event.target.closest("[data-owner-id]");
    if (!button) return;
    const candidate = buildOwnerLookupItems().find((item) => item.id === button.dataset.ownerId);
    if (candidate) {
      selectOwner(candidate);
    }
  });

  syncOwner();
}

function updateMentorFieldState() {
  const employeeTypeField = document.getElementById("employeeType");
  const mentorField = document.getElementById("mentorId");
  if (!employeeTypeField || !mentorField) return;

  const isIntern = normalizeEmployeeType(employeeTypeField.value) === "实习生";
  mentorField.disabled = !isIntern;
  if (!isIntern) {
    mentorField.value = "";
  }
}

async function openPersonEditor(mode, item = null) {
  await loadOptions();
  state.editor = {
    kind: "person",
    mode,
    item
  };

  setEditorAlert("");
  editorTitle.textContent = mode === "edit" ? "编辑人员" : "新增人员";
  submitBtn.textContent = "保存";

  editorFields.innerHTML = `
    <div class="form-grid">
      ${createField({ name: "displayName", label: "中文名", value: item?.displayName || "", required: true })}
      ${createField({ name: "pinyin", label: "拼音", value: item?.pinyin || "", required: true, note: "Web 端当前请手动填写拼音。" })}
      ${createField({ name: "email", label: "邮箱", type: "email", value: item?.email || "@itk-engineering.com", required: true })}
      ${createField({
        name: "department",
        label: "部门",
        value: item?.department || "",
        required: true,
        list: "department-list"
      })}
      ${createField({
        name: "employeeType",
        label: "员工类型",
        type: "select",
        value: item?.employeeType || "",
        options: state.options.employeeTypes.map((value) => ({
          value,
          label: normalizeEmployeeType(value)
        })),
        required: true
      })}
      ${createField({
        name: "mentorId",
        label: "Mentor",
        type: "select",
        value: item?.mentorId || "",
        options: buildMentorOptions(item?.id || ""),
        note: "只有实习生需要选择 Mentor。"
      })}
    </div>
    ${buildDataList("department-list", state.options.departments)}
  `;

  document.getElementById("employeeType").addEventListener("change", updateMentorFieldState);
  updateMentorFieldState();
  editorModal.showModal();
}

async function openEditorForCreate() {
  if (state.view === "people") {
    await openPersonEditor("create");
    return;
  }
  await loadOptions();
  openComputerEditor("create");
}

async function openEditorForEdit() {
  const item = getSelectedRow();
  if (!item) {
    setEditorAlert("");
    window.alert("请先选择一条记录。");
    return;
  }

  if (state.view === "people") {
    await openPersonEditor("edit", item);
    return;
  }

  await loadOptions();
  openComputerEditor("edit", item);
}

async function submitComputerForm(formData) {
  const ownerInput = formData.get("ownerInput");
  const ownerMatch = resolveOwnerInputToId(ownerInput, formData.get("ownerId"));
  if (!ownerMatch.matched) {
    throw new Error("归属人未正确匹配，请从自动提示中选择，或填写“库存”。");
  }

  const payload = {
    computerName: formData.get("computerName"),
    serialNumber: formData.get("serialNumber"),
    assetNumber: formData.get("assetNumber"),
    model: formData.get("model"),
    macAddress: formatMacAddressDisplay(formData.get("macAddress")),
    ownerId: ownerMatch.id,
    remark: formData.get("remark"),
    authorizationPassword: formData.get("authorizationPassword")
  };

  const isEdit = state.editor.mode === "edit";
  const url = isEdit
    ? `/api/computers/${encodeURIComponent(state.editor.item.id)}`
    : "/api/computers";
  const method = isEdit ? "PUT" : "POST";

  const data = await fetchJson(url, {
    method,
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });

  state.selectedId = data.item.id;
}

async function submitPersonForm(formData) {
  const payload = {
    displayName: formData.get("displayName"),
    pinyin: formData.get("pinyin"),
    email: formData.get("email"),
    department: formData.get("department"),
    employeeType: formData.get("employeeType"),
    mentorId: formData.get("mentorId")
  };

  const isEdit = state.editor.mode === "edit";
  const url = isEdit
    ? `/api/colleagues/${encodeURIComponent(state.editor.item.id)}`
    : "/api/colleagues";
  const method = isEdit ? "PUT" : "POST";

  const data = await fetchJson(url, {
    method,
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });

  state.selectedId = data.item.id;
}

async function handleDeleteAction() {
  const item = getSelectedRow();
  if (!item) {
    window.alert("请先选择一条记录。");
    return;
  }

  if (state.view === "in-use") {
    if (!window.confirm(`确定将电脑“${item.computerName}”转入库存吗？`)) return;
    await fetchJson(`/api/computers/${encodeURIComponent(item.id)}/transfer-to-inventory`, {
      method: "POST"
    });
  } else if (state.view === "inventory") {
    if (!window.confirm(`确定删除库存电脑“${item.computerName}”吗？`)) return;
    await fetchJson(`/api/inventory/${encodeURIComponent(item.id)}`, {
      method: "DELETE"
    });
  } else {
    const extra = item.ownedCount ? `\n\n该人员名下的 ${item.ownedCount} 台电脑会自动转入库存。` : "";
    if (!window.confirm(`确定删除人员“${item.displayName}”吗？${extra}`)) return;
    await fetchJson(`/api/colleagues/${encodeURIComponent(item.id)}`, {
      method: "DELETE"
    });
  }

  state.selectedId = "";
  await reloadAll();
}

function renderMailPreview(data) {
  setMailAlert("");
  mailResult.innerHTML = "";
  mailResult.classList.add("hidden");
  mailHint.textContent = `以下收件人会按桌面端规则生成邮件草稿。有效收件人 ${data.validRecipients.length} 位，跳过 ${data.skippedRecipients.length} 位。`;

  mailList.innerHTML = data.validRecipients.length
    ? data.validRecipients
        .map(
          (recipient) => `
            <label class="mail-item">
              <div class="mail-item-head">
                <input type="checkbox" class="mail-checkbox" value="${escapeHtml(recipient.colleague_id)}" checked />
                <div>
                  <div class="mail-item-title">${escapeHtml(recipient.display_name)} (${escapeHtml(recipient.email)})</div>
                  <div class="mail-item-meta">${escapeHtml(recipient.computers.length)} 台电脑，来源：${escapeHtml((recipient.source_labels || []).join("、"))}</div>
                </div>
              </div>
            </label>
          `
        )
        .join("")
    : `<div class="mail-item">当前没有可发送盘点邮件的对象。</div>`;

  mailSkipped.innerHTML = data.skippedRecipients.length
    ? `
        <div class="mail-skip-title">已跳过</div>
        ${data.skippedRecipients
          .map(
            (item) => `
              <div class="mail-skip-item">
                <div class="mail-item-title">${escapeHtml(item.display_name)}</div>
                <div class="mail-skip-meta">${escapeHtml(item.result_message)}</div>
              </div>
            `
          )
          .join("")}
      `
    : "";
}

async function openMailCheckModal() {
  const data = await fetchJson("/api/inventory-mail/preview");
  renderMailPreview(data);
  mailModal.showModal();
}

function getSelectedMailRecipientIds() {
  return [...mailList.querySelectorAll(".mail-checkbox:checked")].map((input) => input.value);
}

function openMailtoLink(mailto) {
  const link = document.createElement("a");
  link.href = mailto;
  link.style.display = "none";
  document.body.appendChild(link);
  link.click();
  link.remove();
}

function renderMailResult(result) {
  const createdCount = result.createdItems.length;
  const skippedCount = result.skippedItems.length;
  const mailtoPayload = escapeHtml(JSON.stringify(result.selectedRecipients.map((item) => item.mailto)));

  mailResult.innerHTML = `
    <div class="mail-result-summary">
      已生成 ${createdCount} 封邮件草稿入口，跳过 ${skippedCount} 位。
    </div>
    <div class="mail-result-actions">
      <button type="button" class="primary-btn mail-open-all-btn" data-mailto-list="${mailtoPayload}">一键打开全部邮件</button>
    </div>
    <div class="mail-result-list">
      ${result.selectedRecipients
        .map(
          (recipient) => `
            <div class="mail-result-item">
              <div>
                <div class="mail-item-title">${escapeHtml(recipient.display_name)} (${escapeHtml(recipient.email)})</div>
                <div class="mail-item-meta">${escapeHtml(recipient.computer_count)} 台电脑，来源：${escapeHtml((recipient.source_labels || []).join("、"))}</div>
              </div>
              <button type="button" class="secondary-btn mail-open-btn" data-mailto="${escapeHtml(recipient.mailto)}">打开邮件</button>
            </div>
          `
        )
        .join("")}
    </div>
  `;
  mailResult.classList.remove("hidden");
}

function openMailtoLinks(mailtoList) {
  mailtoList.forEach((mailto, index) => {
    setTimeout(() => {
      openMailtoLink(mailto);
    }, index * 250);
  });
}

async function generateMailDrafts() {
  setMailAlert("");
  const selectedRecipientIds = getSelectedMailRecipientIds();
  if (!selectedRecipientIds.length) {
    setMailAlert("请至少勾选一位收件人。");
    return;
  }

  const result = await fetchJson("/api/inventory-mail/batches", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ selectedRecipientIds })
  });

  renderMailResult(result);
  if (result.selectedRecipients.length === 1) {
    openMailtoLink(result.selectedRecipients[0].mailto);
  }
}

function syncViewUI() {
  const config = viewConfig[state.view];
  workspaceTitle.textContent = config.title;
  searchInput.placeholder = config.searchPlaceholder;
  addBtn.textContent = config.addLabel;
  editBtn.textContent = config.editLabel;
  deleteBtn.textContent = config.deleteLabel;
  mailCheckBtn.style.display = state.view === "in-use" ? "" : "none";
}

async function loadMeta() {
  const data = await fetchJson("/api/meta");
  currentUser.textContent = data.currentUser || "Admin";
  currentRole.textContent = data.role || "admin";
  document.title = data.systemName || "ITK China 电脑信息管理系统";
}

async function loadOptions() {
  const data = await fetchJson("/api/options");
  state.options = {
    colleagues: Array.isArray(data.colleagues) ? data.colleagues : [],
    departments: Array.isArray(data.departments) ? data.departments : [],
    models: Array.isArray(data.models) ? data.models : [],
    employeeTypes: Array.isArray(data.employeeTypes) ? data.employeeTypes : []
  };
}

async function loadDashboard() {
  const data = await fetchJson("/api/dashboard");
  renderStats(data.cards || []);
  renderRecentList(data.recentUpdates || []);
  syncWorkspaceHeights();
}

async function loadList() {
  syncViewUI();

  const params = new URLSearchParams({
    view: state.view
  });

  if (state.keyword) {
    params.set("q", state.keyword);
  }

  const data = await fetchJson(`/api/list?${params.toString()}`);
  state.rows = sortRows(data.rows || []);

  if (state.selectedId && !state.rows.some((row) => row.id === state.selectedId)) {
    state.selectedId = state.rows[0]?.id || "";
  }

  if (!state.selectedId && state.rows.length) {
    state.selectedId = state.rows[0].id;
  }

  listSummary.textContent = `共 ${state.rows.length} 条`;
  renderTable();
  await refreshDetailPanel();
}

async function reloadAll() {
  await Promise.all([loadOptions(), loadDashboard()]);
  await loadList();
}

document.querySelectorAll(".tab-btn").forEach((button) => {
  button.addEventListener("click", async () => {
    document.querySelectorAll(".tab-btn").forEach((item) => item.classList.remove("active"));
    button.classList.add("active");
    state.view = button.dataset.view;
    state.keyword = "";
    state.selectedId = "";
    state.sort = { key: "", direction: "asc" };
    searchInput.value = "";
    await loadList();
  });
});

searchInput.addEventListener("input", async (event) => {
  state.keyword = event.target.value.trim();
  await loadList();
});

tableHead.addEventListener("click", async (event) => {
  const button = event.target.closest("[data-sort-key]");
  if (!button) return;

  const key = button.dataset.sortKey;
  const config = viewConfig[state.view];
  if (!config.sortableColumns.includes(key)) return;

  if (state.sort.key === key) {
    state.sort.direction = state.sort.direction === "asc" ? "desc" : "asc";
  } else {
    state.sort = {
      key,
      direction: "asc"
    };
  }

  state.rows = sortRows(state.rows);
  renderTable();
  await refreshDetailPanel();
});

tableBody.addEventListener("click", async (event) => {
  const row = event.target.closest("tr[data-id]");
  if (!row) return;
  state.selectedId = row.dataset.id;
  renderTable();
  await refreshDetailPanel();
});

tableBody.addEventListener("dblclick", async (event) => {
  const row = event.target.closest("tr[data-id]");
  if (!row) return;
  state.selectedId = row.dataset.id;
  renderTable();
  await openDetailModal();
});

addBtn.addEventListener("click", () => {
  openEditorForCreate().catch((error) => window.alert(error.message));
});
editBtn.addEventListener("click", () => {
  openEditorForEdit().catch((error) => window.alert(error.message));
});
deleteBtn.addEventListener("click", () => {
  handleDeleteAction().catch((error) => window.alert(error.message));
});
detailBtn.addEventListener("click", () => {
  openDetailModal().catch((error) => window.alert(error.message));
});
refreshBtn.addEventListener("click", () => {
  reloadAll().catch((error) => window.alert(error.message));
});
mailCheckBtn.addEventListener("click", () => {
  openMailCheckModal().catch((error) => window.alert(error.message));
});

document.getElementById("close-detail").addEventListener("click", () => detailModal.close());
document.getElementById("close-editor").addEventListener("click", () => editorModal.close());
document.getElementById("cancel-editor").addEventListener("click", () => editorModal.close());
document.getElementById("close-mail").addEventListener("click", () => mailModal.close());
document.getElementById("mail-cancel").addEventListener("click", () => mailModal.close());
document.getElementById("mail-select-all").addEventListener("click", () => {
  mailList.querySelectorAll(".mail-checkbox").forEach((input) => {
    input.checked = true;
  });
});
document.getElementById("mail-clear-all").addEventListener("click", () => {
  mailList.querySelectorAll(".mail-checkbox").forEach((input) => {
    input.checked = false;
  });
});
document.getElementById("mail-generate").addEventListener("click", () => {
  generateMailDrafts().catch((error) => setMailAlert(error.message));
});
mailResult.addEventListener("click", (event) => {
  const openAllButton = event.target.closest("[data-mailto-list]");
  if (openAllButton) {
    try {
      const mailtoList = JSON.parse(openAllButton.dataset.mailtoList);
      openMailtoLinks(mailtoList);
    } catch (error) {
      setMailAlert("批量打开邮件失败，请逐个打开。");
    }
    return;
  }

  const button = event.target.closest("[data-mailto]");
  if (!button) return;
  openMailtoLink(button.dataset.mailto);
});

window.addEventListener("resize", syncWorkspaceHeights);

editorForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  setEditorAlert("");
  const formData = new FormData(editorForm);

  try {
    if (state.editor.kind === "person") {
      await submitPersonForm(formData);
    } else {
      await submitComputerForm(formData);
    }
    editorModal.close();
    await reloadAll();
  } catch (error) {
    setEditorAlert(error.message);
  }
});

async function boot() {
  await Promise.all([loadMeta(), loadOptions(), loadDashboard()]);
  syncViewUI();
  await loadList();
  syncWorkspaceHeights();

  if (window.ResizeObserver && sideStack) {
    workspaceResizeObserver = new ResizeObserver(() => {
      syncWorkspaceHeights();
    });
    workspaceResizeObserver.observe(sideStack);
  }
}

boot().catch((error) => {
  console.error(error);
  recentList.innerHTML = `<article class="recent-item"><h3>加载失败</h3><div class="recent-meta">${escapeHtml(error.message)}</div></article>`;
  detailPanelContent.className = "detail-content detail-panel-empty";
  detailPanelContent.textContent = error.message;
});
