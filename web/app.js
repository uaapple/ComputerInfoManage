const state = {
  view: "in-use",
  keyword: "",
  rows: []
};

const tableHead = document.getElementById("table-head");
const tableBody = document.getElementById("table-body");
const statsGrid = document.getElementById("stats-grid");
const recentList = document.getElementById("recent-list");
const searchInput = document.getElementById("search-input");
const modal = document.getElementById("detail-modal");
const detailTitle = document.getElementById("detail-title");
const detailContent = document.getElementById("detail-content");
const currentUser = document.getElementById("current-user");
const currentRole = document.getElementById("current-role");

const viewColumns = {
  "in-use": [
    ["computerName", "电脑名称"],
    ["serialNumber", "序列号"],
    ["assetNumber", "固定资产号"],
    ["model", "型号"],
    ["macAddress", "MAC 地址"],
    ["ownerName", "归属人"],
    ["updatedAt", "更新时间"]
  ],
  inventory: [
    ["computerName", "电脑名称"],
    ["serialNumber", "序列号"],
    ["assetNumber", "固定资产号"],
    ["model", "型号"],
    ["macAddress", "MAC 地址"],
    ["remark", "备注"],
    ["updatedAt", "更新时间"]
  ],
  people: [
    ["displayName", "姓名"],
    ["department", "部门"],
    ["employeeType", "员工类型"],
    ["mentorName", "Mentor"],
    ["email", "邮箱"],
    ["ownedCount", "名下电脑"]
  ]
};

function formatTag(value, kind = "na") {
  return `<span class="tag ${kind}">${value}</span>`;
}

function renderCell(key, row) {
  const value = row[key] ?? "";
  if (key === "computerName" || key === "displayName") {
    return `<span class="cell-strong">${value}</span>`;
  }
  if (key === "assetNumber" || key === "macAddress") {
    return value === "N/A" ? formatTag("N/A", "na") : value;
  }
  if (key === "ownerName") {
    return row.ownerId ? value : formatTag("库存", "inventory");
  }
  if (key === "employeeType") {
    return value === "实习生" ? formatTag(value, "intern") : value;
  }
  if (key === "mentorName") {
    return value === "N/A" ? formatTag("N/A", "na") : value;
  }
  return value || formatTag("N/A", "na");
}

function renderTable() {
  const columns = viewColumns[state.view];
  tableHead.innerHTML = `<tr>${columns.map(([, label]) => `<th>${label}</th>`).join("")}</tr>`;
  tableBody.innerHTML = state.rows
    .map(
      (row) => `
        <tr data-id="${row.id}">
          ${columns.map(([key]) => `<td>${renderCell(key, row)}</td>`).join("")}
        </tr>
      `
    )
    .join("");
}

async function fetchJson(url) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`Request failed: ${response.status}`);
  return response.json();
}

async function loadMeta() {
  const data = await fetchJson("/api/meta");
  currentUser.textContent = data.currentUser;
  currentRole.textContent = data.role;
}

async function loadDashboard() {
  const data = await fetchJson("/api/dashboard");
  statsGrid.innerHTML = data.cards
    .map(
      (card) => `
        <article class="stat-card ${card.tone}">
          <div class="stat-label">${card.label}</div>
          <div class="stat-value">${card.value}</div>
        </article>
      `
    )
    .join("");

  recentList.innerHTML = data.recentUpdates
    .map(
      (item) => `
        <article class="recent-item">
          <h3>${item.computerName}</h3>
          <div class="recent-meta">归属人：${item.ownerName}</div>
          <div class="recent-meta">型号：${item.model || "N/A"}</div>
          <div class="recent-meta">更新时间：${item.updatedAt || "N/A"}</div>
        </article>
      `
    )
    .join("");
}

async function loadList() {
  const query = new URLSearchParams({ view: state.view, q: state.keyword });
  const data = await fetchJson(`/api/list?${query.toString()}`);
  state.rows = data.rows;
  renderTable();
}

function renderDetailFields(fields) {
  return `<div class="detail-grid">${fields
    .map(
      ([label, value]) => `
        <div class="detail-field">
          <label>${label}</label>
          <p>${value || "N/A"}</p>
        </div>
      `
    )
    .join("")}</div>`;
}

async function openDetail(id) {
  const type = state.view === "people" ? "person" : "computer";
  const data = await fetchJson(`/api/details?type=${type}&id=${id}`);

  if (type === "computer") {
    detailTitle.textContent = data.computerName;
    detailContent.innerHTML = `
      ${renderDetailFields([
        ["序列号", data.serialNumber],
        ["固定资产号", data.assetNumber],
        ["型号", data.model],
        ["MAC 地址", data.macAddress],
        ["归属人", data.ownerName],
        ["更新时间", data.updatedAt],
        ["备注", data.remark || "N/A"]
      ])}
      <div class="detail-field">
        <label>归属历史</label>
        <div class="history-list">
          ${(data.ownerHistory || [])
            .map(
              (item) => `
                <div class="history-item">
                  <strong>${item.changed_at || "N/A"}</strong><br />
                  ${item.old_owner_name || "库存"} → ${item.new_owner_name || "库存"}
                </div>
              `
            )
            .join("") || "<p>暂无归属历史</p>"}
        </div>
      </div>
    `;
  } else {
    detailTitle.textContent = data.displayName;
    detailContent.innerHTML = renderDetailFields([
      ["拼音", data.pinyin],
      ["邮箱", data.email],
      ["部门", data.department],
      ["员工类型", data.employeeType],
      ["Mentor", data.mentorName],
      ["名下电脑", String(data.ownedCount)]
    ]);
  }

  modal.showModal();
}

document.querySelectorAll(".tab-btn").forEach((button) => {
  button.addEventListener("click", async () => {
    document.querySelectorAll(".tab-btn").forEach((item) => item.classList.remove("active"));
    button.classList.add("active");
    state.view = button.dataset.view;
    await loadList();
  });
});

searchInput.addEventListener("input", async (event) => {
  state.keyword = event.target.value.trim();
  await loadList();
});

tableBody.addEventListener("click", async (event) => {
  const row = event.target.closest("tr[data-id]");
  if (!row) return;
  await openDetail(row.dataset.id);
});

document.getElementById("close-detail").addEventListener("click", () => {
  modal.close();
});

modal.addEventListener("click", (event) => {
  const rect = modal.getBoundingClientRect();
  const inside =
    rect.top <= event.clientY &&
    event.clientY <= rect.top + rect.height &&
    rect.left <= event.clientX &&
    event.clientX <= rect.left + rect.width;
  if (!inside) modal.close();
});

async function boot() {
  await Promise.all([loadMeta(), loadDashboard(), loadList()]);
}

boot().catch((error) => {
  console.error(error);
  recentList.innerHTML = `<article class="recent-item"><h3>加载失败</h3><div class="recent-meta">${error.message}</div></article>`;
});
