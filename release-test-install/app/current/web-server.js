const http = require("http");
const fs = require("fs");
const path = require("path");
const { URL } = require("url");
const { randomUUID } = require("crypto");

const HOST = process.env.HOST || "0.0.0.0";
const PORT = Number(process.env.PORT || 8099);
const ROOT = __dirname;
const PUBLIC_DIR = path.join(ROOT, "web");
const DATA_DIR = path.resolve(process.env.DATA_DIR || path.join(ROOT, "data"));
const EDIT_PASSWORD = "1";

function nowString() {
  const date = new Date();
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const hour = String(date.getHours()).padStart(2, "0");
  const minute = String(date.getMinutes()).padStart(2, "0");
  const second = String(date.getSeconds()).padStart(2, "0");
  return `${year}-${month}-${day} ${hour}:${minute}:${second}`;
}

function stripBom(text) {
  return String(text || "").replace(/^\uFEFF/, "");
}

function sendJson(res, statusCode, payload) {
  const body = JSON.stringify(payload);
  res.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(body),
    "Cache-Control": "no-store"
  });
  res.end(body);
}

function sendFile(req, res, filePath, contentType) {
  fs.readFile(filePath, (err, data) => {
    if (err) {
      sendJson(res, 404, { error: "Not Found" });
      return;
    }

    res.writeHead(200, { "Content-Type": contentType });
    if (req.method === "HEAD") {
      res.end();
      return;
    }
    res.end(data);
  });
}

function readJson(fileName) {
  const filePath = path.join(DATA_DIR, fileName);
  if (!fs.existsSync(filePath)) {
    return [];
  }

  const raw = stripBom(fs.readFileSync(filePath, "utf8"));
  if (!raw.trim()) {
    return [];
  }

  const parsed = JSON.parse(raw);
  return Array.isArray(parsed) ? parsed : [parsed];
}

function writeJson(fileName, data) {
  const filePath = path.join(DATA_DIR, fileName);
  const body = `${JSON.stringify(data, null, 2)}\n`;
  fs.writeFileSync(filePath, body, "utf8");
}

function normalizeNullableText(value) {
  return String(value || "").trim();
}

function normalizeDisplay(value) {
  const text = normalizeNullableText(value);
  if (!text || ["none", "null", "n/a", "na", "暂无"].includes(text.toLowerCase())) {
    return "N/A";
  }
  return text;
}

function normalizeAssetNumberInput(value) {
  const text = normalizeNullableText(value);
  if (!text || ["暂无", "n/a", "na", "none", "null"].includes(text.toLowerCase())) {
    return "";
  }
  return text;
}

function normalizeMacAddressInput(value) {
  const text = normalizeNullableText(value);
  if (!text || ["暂无", "n/a", "na", "none", "null"].includes(text.toLowerCase())) {
    return "";
  }
  return text;
}

function normalizeColleagueRecord(record) {
  const employeeType = normalizeNullableText(record.employee_type) || "正式员工";
  const normalized = {
    id: normalizeNullableText(record.id) || randomUUID(),
    display_name: normalizeNullableText(record.display_name),
    pinyin: normalizeNullableText(record.pinyin),
    email: normalizeNullableText(record.email) || "@itk-engineering.com",
    department: normalizeNullableText(record.department),
    employee_type: employeeType,
    mentor_id: employeeType === "实习生" ? normalizeNullableText(record.mentor_id) : ""
  };
  return normalized;
}

function getOwnerLabel(colleagueMap, ownerId) {
  const id = normalizeNullableText(ownerId);
  if (!id) return "库存";
  const owner = colleagueMap.get(id);
  return owner ? owner.display_name || "未知人员" : "未知人员";
}

function normalizeOwnerHistory(history) {
  if (!Array.isArray(history)) return [];
  return history.map((item) => ({
    changed_at: normalizeNullableText(item.changed_at),
    old_owner_id: normalizeNullableText(item.old_owner_id),
    old_owner_name: normalizeNullableText(item.old_owner_name),
    new_owner_id: normalizeNullableText(item.new_owner_id),
    new_owner_name: normalizeNullableText(item.new_owner_name)
  }));
}

function normalizeComputerRecord(record, colleagueMap) {
  const updatedAt = normalizeNullableText(record.updated_at) || nowString();
  const ownerId = normalizeNullableText(record.owner_id);
  let ownerHistory = normalizeOwnerHistory(record.owner_history);

  if (!ownerHistory.length && ownerId) {
    ownerHistory = [
      {
        changed_at: updatedAt,
        old_owner_id: "",
        old_owner_name: "库存",
        new_owner_id: ownerId,
        new_owner_name: getOwnerLabel(colleagueMap, ownerId)
      }
    ];
  }

  return {
    id: normalizeNullableText(record.id) || randomUUID(),
    computer_name: normalizeNullableText(record.computer_name),
    serial_number: normalizeNullableText(record.serial_number),
    asset_number: normalizeAssetNumberInput(record.asset_number),
    model: normalizeNullableText(record.model),
    mac_address: normalizeMacAddressInput(record.mac_address),
    owner_id: ownerId,
    owner_history: ownerHistory,
    remark: normalizeNullableText(record.remark),
    updated_at: updatedAt
  };
}

function loadStore() {
  const colleagues = readJson("colleagues.json").map(normalizeColleagueRecord);
  const colleagueMap = new Map(colleagues.map((item) => [item.id, item]));
  const computers = readJson("computers.json").map((item) => normalizeComputerRecord(item, colleagueMap));
  const mailBatches = readJson("inventory_mail_batches.json");

  return {
    computers,
    colleagues,
    mailBatches
  };
}

function saveStore(store) {
  writeJson("computers.json", store.computers);
  writeJson("colleagues.json", store.colleagues);
  writeJson("inventory_mail_batches.json", store.mailBatches);
}

function createColleagueMap(colleagues) {
  return new Map(colleagues.map((item) => [item.id, item]));
}

function hydrateComputer(item, colleagueMap) {
  return {
    id: item.id,
    computerName: item.computer_name,
    serialNumber: item.serial_number,
    assetNumber: normalizeDisplay(item.asset_number),
    model: item.model,
    macAddress: normalizeDisplay(item.mac_address),
    ownerId: item.owner_id,
    ownerName: getOwnerLabel(colleagueMap, item.owner_id),
    remark: item.remark,
    updatedAt: item.updated_at,
    ownerHistory: normalizeOwnerHistory(item.owner_history)
  };
}

function hydrateColleague(item, computers, colleagueMap) {
  const ownedCount = computers.filter((computer) => computer.owner_id === item.id).length;
  return {
    id: item.id,
    displayName: item.display_name,
    pinyin: item.pinyin,
    email: item.email,
    department: item.department,
    employeeType: item.employee_type,
    mentorId: item.mentor_id,
    mentorName: item.mentor_id ? getOwnerLabel(colleagueMap, item.mentor_id) : "N/A",
    ownedCount
  };
}

function filterByKeyword(rows, keyword, fields) {
  const query = normalizeNullableText(keyword).toLowerCase();
  if (!query) return rows;

  return rows.filter((row) =>
    fields.some((field) => String(row[field] || "").toLowerCase().includes(query))
  );
}

function addComputerOwnerHistoryEntry({ computer, oldOwnerId, newOwnerId, colleagueMap, changedAt }) {
  computer.owner_history = [...normalizeOwnerHistory(computer.owner_history), {
    changed_at: changedAt,
    old_owner_id: oldOwnerId,
    old_owner_name: getOwnerLabel(colleagueMap, oldOwnerId),
    new_owner_id: newOwnerId,
    new_owner_name: getOwnerLabel(colleagueMap, newOwnerId)
  }];
}

function setComputerOwner({ computer, newOwnerId, colleagueMap, changedAt }) {
  const oldOwnerId = normalizeNullableText(computer.owner_id);
  const nextOwnerId = normalizeNullableText(newOwnerId);
  if (oldOwnerId === nextOwnerId) {
    return false;
  }

  computer.owner_id = nextOwnerId;
  addComputerOwnerHistoryEntry({
    computer,
    oldOwnerId,
    newOwnerId: nextOwnerId,
    colleagueMap,
    changedAt
  });
  return true;
}

function assertEmailValid(email) {
  return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email);
}

function validateComputerPayload(store, payload, editingId = "") {
  const name = normalizeNullableText(payload.computerName);
  const serial = normalizeNullableText(payload.serialNumber);
  const asset = normalizeAssetNumberInput(payload.assetNumber);
  const ownerId = normalizeNullableText(payload.ownerId);

  if (!name) {
    throw new Error("请输入电脑名称。");
  }
  if (!serial) {
    throw new Error("请输入序列号。");
  }
  if (ownerId && !store.colleagues.some((item) => item.id === ownerId)) {
    throw new Error("归属同事未正确匹配，请重新选择一位人员。");
  }

  const duplicate = store.computers.find((item) => {
    if (item.id === editingId) return false;
    if (item.serial_number === serial) return true;
    if (asset && normalizeAssetNumberInput(item.asset_number) === asset) return true;
    return false;
  });

  if (duplicate) {
    throw new Error("序列号或固定资产号已存在，请确认后再保存。");
  }

  return {
    computer_name: name,
    serial_number: serial,
    asset_number: asset,
    model: normalizeNullableText(payload.model),
    mac_address: normalizeMacAddressInput(payload.macAddress),
    owner_id: ownerId,
    remark: normalizeNullableText(payload.remark)
  };
}

function validateColleaguePayload(store, payload, editingId = "") {
  const displayName = normalizeNullableText(payload.displayName);
  const pinyin = normalizeNullableText(payload.pinyin).toLowerCase();
  const email = normalizeNullableText(payload.email);
  const department = normalizeNullableText(payload.department);
  const employeeType = normalizeNullableText(payload.employeeType);
  let mentorId = normalizeNullableText(payload.mentorId);

  if (!displayName) throw new Error("请输入人员中文名。");
  if (!pinyin) throw new Error("请输入人员拼音。");
  if (!email) throw new Error("请输入邮箱。");
  if (!assertEmailValid(email)) throw new Error("邮箱格式不正确，请至少包含 @ 和域名。");
  if (!department) throw new Error("请输入部门。");
  if (!["正式员工", "实习生"].includes(employeeType)) throw new Error("请选择员工类型。");

  if (employeeType === "实习生") {
    if (!mentorId) throw new Error("实习生必须关联一位正式员工 Mentor。");
    const mentor = store.colleagues.find((item) => item.id === mentorId);
    if (!mentor || mentor.employee_type !== "正式员工") {
      throw new Error("Mentor 必须是已存在的正式员工。");
    }
    if (mentorId === editingId) {
      throw new Error("Mentor 不能选择本人。");
    }
  } else {
    mentorId = "";
  }

  if (editingId) {
    const hasDependents = store.colleagues.some((item) => item.mentor_id === editingId);
    if (hasDependents && employeeType !== "正式员工") {
      throw new Error("该人员当前已被其他实习生关联为 Mentor，请先调整相关 Mentor 关系。");
    }
  }

  return {
    display_name: displayName,
    pinyin,
    email,
    department,
    employee_type: employeeType,
    mentor_id: mentorId
  };
}

function buildOptionsPayload(store) {
  const colleagueMap = createColleagueMap(store.colleagues);
  const departments = [...new Set(store.colleagues.map((item) => item.department).filter(Boolean))].sort();
  const models = [...new Set(store.computers.map((item) => item.model).filter(Boolean))].sort();
  const colleagues = store.colleagues
    .map((item) => ({
      id: item.id,
      displayName: item.display_name,
      pinyin: item.pinyin,
      department: item.department,
      employeeType: item.employee_type,
      mentorId: item.mentor_id,
      label: `${item.display_name} (${item.pinyin}, ${item.department})`,
      ownedCount: store.computers.filter((computer) => computer.owner_id === item.id).length,
      mentorName: item.mentor_id ? getOwnerLabel(colleagueMap, item.mentor_id) : "N/A"
    }))
    .sort((a, b) => a.displayName.localeCompare(b.displayName, "zh-CN"));

  return {
    colleagues,
    departments,
    models,
    employeeTypes: ["正式员工", "实习生"],
    editPasswordHint: "1"
  };
}

function testEmailValid(email) {
  if (!normalizeNullableText(email)) return false;
  return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(normalizeNullableText(email));
}

function getInventoryMailSubject() {
  return "[电脑盘点] 请确认您名下电脑信息";
}

function newInventoryMailComputerSnapshot(computer, colleagueMap) {
  return {
    id: computer.id,
    computer_name: computer.computer_name,
    serial_number: computer.serial_number,
    asset_number: normalizeDisplay(computer.asset_number),
    model: computer.model,
    mac_address: normalizeDisplay(computer.mac_address),
    owner_id: computer.owner_id,
    owner_name: getOwnerLabel(colleagueMap, computer.owner_id),
    remark: computer.remark,
    updated_at: computer.updated_at
  };
}

function formatInventoryMailComputerList(computers, colleagueMap) {
  const lines = [];
  computers.forEach((computer, index) => {
    lines.push(`${index + 1}. 电脑名称：${computer.computer_name}`);
    lines.push(`   当前归属人：${getOwnerLabel(colleagueMap, computer.owner_id)}`);
    lines.push(`   序列号：${computer.serial_number}`);
    lines.push(`   固定资产号：${normalizeDisplay(computer.asset_number)}`);
    lines.push(`   型号：${computer.model || "N/A"}`);
    lines.push(`   MAC 地址：${normalizeDisplay(computer.mac_address)}`);
    lines.push(`   备注：${computer.remark || "无"}`);
    lines.push("");
  });
  return lines.join("\n").trimEnd();
}

function getInventoryMailBody(recipient, colleagueMap) {
  const computerList = formatInventoryMailComputerList(recipient.computers, colleagueMap);
  const displayName = recipient.display_name;
  const selfComputers = recipient.computers.filter((item) => item.owner_id === recipient.colleague_id);
  const delegatedComputers = recipient.computers.filter((item) => item.owner_id !== recipient.colleague_id);
  const ownerNames = [...new Set(delegatedComputers.map((item) => getOwnerLabel(colleagueMap, item.owner_id)))];

  let introText = "为便于完成当前电脑资产盘点，请您确认以下登记在您名下的电脑目前仍由您本人使用，且设备状态正常。";
  if (delegatedComputers.length && !selfComputers.length) {
    introText = "为便于完成当前电脑资产盘点，请您协助确认以下登记在您所负责实习生名下的电脑信息。";
  } else if (delegatedComputers.length && selfComputers.length) {
    introText = "为便于完成当前电脑资产盘点，请您一并确认以下登记在您本人名下及您所负责实习生名下的电脑信息。";
  }

  const ownerHint = ownerNames.length ? `\n涉及实习生：${ownerNames.join("、")}` : "";

  return `${displayName}，您好：

${introText}${ownerHint}

${computerList}

如以上电脑信息无误，则无需额外回复；自本邮件发出之时起两天内未回复，即视为您确认上述信息准确无误。

如您发现电脑归属、设备状态或清单内容存在异议，请直接回复本邮件反馈，我们会及时更新。

谢谢配合。`;
}

function getInventoryMailRecipients(store) {
  const validRecipients = [];
  const skippedRecipients = [];
  const recipientMap = new Map();
  const colleagueMap = createColleagueMap(store.colleagues);
  const inUseComputers = store.computers.filter((item) => item.owner_id);
  const groups = new Map();

  for (const computer of inUseComputers) {
    const key = computer.owner_id;
    const list = groups.get(key) || [];
    list.push(computer);
    groups.set(key, list);
  }

  for (const [ownerId, computers] of [...groups.entries()].sort((a, b) => a[0].localeCompare(b[0]))) {
    const ownerColleague = store.colleagues.find((item) => item.id === ownerId);
    const sortedComputers = [...computers].sort((a, b) =>
      `${a.computer_name}${a.serial_number}${a.asset_number}`.localeCompare(`${b.computer_name}${b.serial_number}${b.asset_number}`, "zh-CN")
    );

    if (!ownerColleague) {
      skippedRecipients.push({
        colleague_id: ownerId,
        display_name: "未知人员",
        email: "",
        computers: sortedComputers,
        result: "skipped_missing_colleague",
        result_message: "未找到对应的人员记录。"
      });
      continue;
    }

    let recipientColleague = ownerColleague;
    let sourceLabel = "正式员工本人";

    if (ownerColleague.employee_type === "实习生") {
      const mentorId = normalizeNullableText(ownerColleague.mentor_id);
      if (!mentorId) {
        skippedRecipients.push({
          colleague_id: ownerColleague.id,
          display_name: ownerColleague.display_name,
          email: ownerColleague.email,
          computers: sortedComputers,
          result: "skipped_missing_mentor",
          result_message: "该实习生未关联 Mentor，无法生成盘点邮件。"
        });
        continue;
      }

      const mentorRecord = store.colleagues.find((item) => item.id === mentorId);
      if (!mentorRecord || mentorRecord.employee_type !== "正式员工") {
        skippedRecipients.push({
          colleague_id: ownerColleague.id,
          display_name: ownerColleague.display_name,
          email: ownerColleague.email,
          computers: sortedComputers,
          result: "skipped_invalid_mentor",
          result_message: "该实习生关联的 Mentor 不存在或不是正式员工。"
        });
        continue;
      }

      recipientColleague = mentorRecord;
      sourceLabel = `实习生 Mentor（${ownerColleague.display_name}）`;
    }

    const email = normalizeNullableText(recipientColleague.email);
    if (!testEmailValid(email)) {
      skippedRecipients.push({
        colleague_id: recipientColleague.id,
        display_name: recipientColleague.display_name,
        email,
        computers: sortedComputers,
        result: "skipped_invalid_email",
        result_message: ownerColleague.employee_type === "实习生"
          ? "对应 Mentor 的邮箱为空或格式不正确。"
          : "邮箱为空或格式不正确。"
      });
      continue;
    }

    const recipientKey = recipientColleague.id;
    if (!recipientMap.has(recipientKey)) {
      recipientMap.set(recipientKey, {
        colleague: recipientColleague,
        colleague_id: recipientColleague.id,
        display_name: recipientColleague.display_name,
        email,
        computers: [],
        source_labels: []
      });
    }

    const recipient = recipientMap.get(recipientKey);
    recipient.computers.push(...sortedComputers);
    recipient.source_labels.push(sourceLabel);
  }

  for (const recipient of recipientMap.values()) {
    recipient.computers = recipient.computers.sort((a, b) =>
      `${a.computer_name}${a.serial_number}${a.asset_number}`.localeCompare(`${b.computer_name}${b.serial_number}${b.asset_number}`, "zh-CN")
    );
    recipient.source_labels = [...new Set(recipient.source_labels)];
    const subject = getInventoryMailSubject();
    const body = getInventoryMailBody(recipient, colleagueMap);
    const mailto = `mailto:${encodeURIComponent(recipient.email)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;

    validRecipients.push({
      ...recipient,
      subject,
      body,
      mailto
    });
  }

  return {
    validRecipients: validRecipients.sort((a, b) => `${a.display_name}${a.email}`.localeCompare(`${b.display_name}${b.email}`, "zh-CN")),
    skippedRecipients: skippedRecipients.sort((a, b) => `${a.display_name}${a.colleague_id}`.localeCompare(`${b.display_name}${b.colleague_id}`, "zh-CN"))
  };
}

function newInventoryMailBatchItem(recipient, result, resultMessage, draftCreatedAt, colleagueMap) {
  return {
    colleague_id: recipient.colleague_id,
    display_name: recipient.display_name,
    email: recipient.email,
    computer_ids: recipient.computers.map((item) => item.id),
    computer_snapshot: recipient.computers.map((item) => newInventoryMailComputerSnapshot(item, colleagueMap)),
    result,
    result_message: resultMessage,
    draft_created_at: draftCreatedAt || ""
  };
}

function saveInventoryMailBatchRecord(store, createdAt, status, subjectTemplate, items) {
  const batch = {
    id: randomUUID(),
    created_at: createdAt,
    status,
    recipient_count: items.filter((item) => item.result === "draft_created").length,
    skipped_count: items.filter((item) => item.result.startsWith("skipped_")).length,
    subject_template: subjectTemplate,
    items
  };
  store.mailBatches.push(batch);
  saveStore(store);
  return batch;
}

function handleList(res, store, url) {
  const colleagueMap = createColleagueMap(store.colleagues);
  const view = url.searchParams.get("view") || "in-use";
  const keyword = url.searchParams.get("q") || "";

  if (view === "people") {
    const rows = filterByKeyword(
      store.colleagues.map((item) => hydrateColleague(item, store.computers, colleagueMap)),
      keyword,
      ["displayName", "pinyin", "email", "department", "employeeType", "mentorName"]
    );
    sendJson(res, 200, { rows });
    return;
  }

  const source = view === "inventory"
    ? store.computers.filter((item) => !item.owner_id)
    : store.computers.filter((item) => item.owner_id);

  const rows = filterByKeyword(
    source.map((item) => hydrateComputer(item, colleagueMap)),
    keyword,
    ["computerName", "serialNumber", "assetNumber", "model", "macAddress", "ownerName", "remark"]
  );

  sendJson(res, 200, { rows });
}

function handleDetails(res, store, url) {
  const type = url.searchParams.get("type");
  const id = url.searchParams.get("id");
  const colleagueMap = createColleagueMap(store.colleagues);

  if (type === "computer") {
    const computer = store.computers.find((item) => item.id === id);
    if (!computer) {
      sendJson(res, 404, { error: "Not Found" });
      return;
    }
    sendJson(res, 200, hydrateComputer(computer, colleagueMap));
    return;
  }

  if (type === "person") {
    const person = store.colleagues.find((item) => item.id === id);
    if (!person) {
      sendJson(res, 404, { error: "Not Found" });
      return;
    }
    sendJson(res, 200, hydrateColleague(person, store.computers, colleagueMap));
    return;
  }

  sendJson(res, 404, { error: "Not Found" });
}

async function readRequestBody(req) {
  const chunks = [];
  for await (const chunk of req) {
    chunks.push(chunk);
  }
  const raw = Buffer.concat(chunks).toString("utf8");
  if (!raw.trim()) return {};
  return JSON.parse(raw);
}

async function handleApi(req, res, url) {
  const store = loadStore();
  const colleagueMap = createColleagueMap(store.colleagues);

  if (req.method === "GET" && url.pathname === "/api/meta") {
    sendJson(res, 200, {
      currentUser: process.env.USERNAME || "Admin",
      role: "admin",
      systemName: "ITK China 电脑信息管理系统"
    });
    return;
  }

  if (req.method === "GET" && url.pathname === "/api/options") {
    sendJson(res, 200, buildOptionsPayload(store));
    return;
  }

  if (req.method === "GET" && url.pathname === "/api/dashboard") {
    const hydratedColleagues = store.colleagues.map((item) => hydrateColleague(item, store.computers, colleagueMap));
    const hydratedComputers = store.computers.map((item) => hydrateComputer(item, colleagueMap));
    const inUseCount = hydratedComputers.filter((item) => item.ownerId).length;
    const inventoryCount = hydratedComputers.length - inUseCount;
    const internCount = hydratedColleagues.filter((item) => item.employeeType === "实习生").length;
    const recentUpdates = [...hydratedComputers]
      .sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt)))
      .slice(0, 6);

    sendJson(res, 200, {
      cards: [
        { key: "inUse", label: "在用电脑", value: inUseCount, tone: "blue" },
        { key: "inventory", label: "库存电脑", value: inventoryCount, tone: "teal" },
        { key: "people", label: "人员总数", value: hydratedColleagues.length, tone: "slate" },
        { key: "interns", label: "实习生数量", value: internCount, tone: "amber" }
      ],
      recentUpdates
    });
    return;
  }

  if (req.method === "GET" && url.pathname === "/api/inventory-mail/preview") {
    const summary = getInventoryMailRecipients(store);
    sendJson(res, 200, summary);
    return;
  }

  if (req.method === "GET" && url.pathname === "/api/list") {
    handleList(res, store, url);
    return;
  }

  if (req.method === "GET" && url.pathname === "/api/details") {
    handleDetails(res, store, url);
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/computers") {
    const payload = await readRequestBody(req);
    const data = validateComputerPayload(store, payload);
    const now = nowString();
    const computer = {
      id: randomUUID(),
      ...data,
      owner_history: data.owner_id ? [{
        changed_at: now,
        old_owner_id: "",
        old_owner_name: "库存",
        new_owner_id: data.owner_id,
        new_owner_name: getOwnerLabel(colleagueMap, data.owner_id)
      }] : [],
      updated_at: now
    };

    store.computers.push(computer);
    saveStore(store);
    sendJson(res, 201, { item: hydrateComputer(computer, createColleagueMap(store.colleagues)) });
    return;
  }

  const updateComputerMatch = url.pathname.match(/^\/api\/computers\/([^/]+)$/);
  if (updateComputerMatch && req.method === "PUT") {
    const id = decodeURIComponent(updateComputerMatch[1]);
    const computer = store.computers.find((item) => item.id === id);
    if (!computer) {
      sendJson(res, 404, { error: "未找到对应电脑记录。" });
      return;
    }

    const payload = await readRequestBody(req);
    if (normalizeNullableText(payload.authorizationPassword) !== EDIT_PASSWORD) {
      sendJson(res, 400, { error: "密码不正确，未保存修改。" });
      return;
    }

    const data = validateComputerPayload(store, payload, id);
    const now = nowString();

    computer.computer_name = data.computer_name;
    computer.serial_number = data.serial_number;
    computer.asset_number = data.asset_number;
    computer.model = data.model;
    computer.mac_address = data.mac_address;
    computer.remark = data.remark;
    setComputerOwner({
      computer,
      newOwnerId: data.owner_id,
      colleagueMap,
      changedAt: now
    });
    computer.updated_at = now;

    saveStore(store);
    sendJson(res, 200, { item: hydrateComputer(computer, createColleagueMap(store.colleagues)) });
    return;
  }

  const transferMatch = url.pathname.match(/^\/api\/computers\/([^/]+)\/transfer-to-inventory$/);
  if (transferMatch && req.method === "POST") {
    const id = decodeURIComponent(transferMatch[1]);
    const computer = store.computers.find((item) => item.id === id);
    if (!computer) {
      sendJson(res, 404, { error: "未找到选中的电脑记录。" });
      return;
    }
    if (!computer.owner_id) {
      sendJson(res, 400, { error: "这台电脑已经在库存中。如需彻底删除，请到库存电脑里删除。" });
      return;
    }

    const now = nowString();
    setComputerOwner({
      computer,
      newOwnerId: "",
      colleagueMap,
      changedAt: now
    });
    computer.updated_at = now;
    saveStore(store);
    sendJson(res, 200, { item: hydrateComputer(computer, createColleagueMap(store.colleagues)) });
    return;
  }

  const inventoryDeleteMatch = url.pathname.match(/^\/api\/inventory\/([^/]+)$/);
  if (inventoryDeleteMatch && req.method === "DELETE") {
    const id = decodeURIComponent(inventoryDeleteMatch[1]);
    const computer = store.computers.find((item) => item.id === id);
    if (!computer) {
      sendJson(res, 404, { error: "未找到库存电脑记录。" });
      return;
    }
    if (computer.owner_id) {
      sendJson(res, 400, { error: "只能删除库存电脑记录。" });
      return;
    }

    store.computers = store.computers.filter((item) => item.id !== id);
    saveStore(store);
    sendJson(res, 200, { ok: true });
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/colleagues") {
    const payload = await readRequestBody(req);
    const data = validateColleaguePayload(store, payload);
    const colleague = {
      id: randomUUID(),
      ...data
    };
    store.colleagues.push(colleague);
    saveStore(store);
    sendJson(res, 201, { item: hydrateColleague(colleague, store.computers, createColleagueMap(store.colleagues)) });
    return;
  }

  const updateColleagueMatch = url.pathname.match(/^\/api\/colleagues\/([^/]+)$/);
  if (updateColleagueMatch && req.method === "PUT") {
    const id = decodeURIComponent(updateColleagueMatch[1]);
    const colleague = store.colleagues.find((item) => item.id === id);
    if (!colleague) {
      sendJson(res, 404, { error: "未找到对应人员记录。" });
      return;
    }

    const payload = await readRequestBody(req);
    const data = validateColleaguePayload(store, payload, id);

    colleague.display_name = data.display_name;
    colleague.pinyin = data.pinyin;
    colleague.email = data.email;
    colleague.department = data.department;
    colleague.employee_type = data.employee_type;
    colleague.mentor_id = data.mentor_id;

    saveStore(store);
    sendJson(res, 200, { item: hydrateColleague(colleague, store.computers, createColleagueMap(store.colleagues)) });
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/inventory-mail/batches") {
    const payload = await readRequestBody(req);
    const selectedRecipientIds = Array.isArray(payload.selectedRecipientIds)
      ? payload.selectedRecipientIds.map((item) => normalizeNullableText(item)).filter(Boolean)
      : [];

    const summary = getInventoryMailRecipients(store);
    const colleagueMapForBatch = createColleagueMap(store.colleagues);
    const selectedRecipients = summary.validRecipients.filter((item) => selectedRecipientIds.includes(item.colleague_id));

    if (!selectedRecipients.length) {
      sendJson(res, 400, { error: "请至少勾选一位收件人。" });
      return;
    }

    const createdAt = nowString();
    const subjectTemplate = getInventoryMailSubject();
    const batchItems = [];
    const createdItems = [];
    const skippedItems = [];

    for (const recipient of selectedRecipients) {
      const draftCreatedAt = nowString();
      const createdItem = newInventoryMailBatchItem(
        recipient,
        "draft_created",
        "已打开默认邮件客户端的撰写窗口。",
        draftCreatedAt,
        colleagueMapForBatch
      );
      batchItems.push(createdItem);
      createdItems.push(createdItem);
    }

    for (const recipient of summary.skippedRecipients) {
      const skippedItem = newInventoryMailBatchItem(
        recipient,
        recipient.result,
        recipient.result_message,
        "",
        colleagueMapForBatch
      );
      batchItems.push(skippedItem);
      skippedItems.push(skippedItem);
    }

    const batch = saveInventoryMailBatchRecord(store, createdAt, "draft_created", subjectTemplate, batchItems);

    sendJson(res, 200, {
      batch,
      subjectTemplate,
      createdItems,
      skippedItems,
      selectedRecipients: selectedRecipients.map((item) => ({
        colleague_id: item.colleague_id,
        display_name: item.display_name,
        email: item.email,
        source_labels: item.source_labels,
        computer_count: item.computers.length,
        mailto: item.mailto
      }))
    });
    return;
  }

  const deleteColleagueMatch = url.pathname.match(/^\/api\/colleagues\/([^/]+)$/);
  if (deleteColleagueMatch && req.method === "DELETE") {
    const id = decodeURIComponent(deleteColleagueMatch[1]);
    const colleague = store.colleagues.find((item) => item.id === id);
    if (!colleague) {
      sendJson(res, 404, { error: "未找到对应人员记录。" });
      return;
    }

    const usedAsMentor = store.colleagues.some((item) => item.mentor_id === id);
    if (usedAsMentor) {
      sendJson(res, 400, { error: "该人员仍被实习生关联为 Mentor，请先调整相关 Mentor 关系后再删除。" });
      return;
    }

    const now = nowString();
    const refreshedMap = createColleagueMap(store.colleagues);
    const ownedComputers = store.computers.filter((item) => item.owner_id === id);
    for (const computer of ownedComputers) {
      setComputerOwner({
        computer,
        newOwnerId: "",
        colleagueMap: refreshedMap,
        changedAt: now
      });
      computer.updated_at = now;
    }

    store.colleagues = store.colleagues.filter((item) => item.id !== id);
    saveStore(store);
    sendJson(res, 200, { ok: true, movedCount: ownedComputers.length });
    return;
  }

  sendJson(res, 404, { error: "Not Found" });
}

function getContentType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === ".css") return "text/css; charset=utf-8";
  if (ext === ".js") return "application/javascript; charset=utf-8";
  if (ext === ".jpg" || ext === ".jpeg") return "image/jpeg";
  if (ext === ".png") return "image/png";
  if (ext === ".svg") return "image/svg+xml";
  if (ext === ".ico") return "image/x-icon";
  return "text/html; charset=utf-8";
}

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url, `http://${req.headers.host}`);

    if (url.pathname.startsWith("/api/")) {
      await handleApi(req, res, url);
      return;
    }

    if (url.pathname === "/assets/logo.jpg") {
      sendFile(req, res, path.join(ROOT, "ITK_Logo_RGB.jpg"), "image/jpeg");
      return;
    }

    const safePath = url.pathname === "/" ? "/index.html" : url.pathname;
    const filePath = path.normalize(path.join(PUBLIC_DIR, safePath));
    if (!filePath.startsWith(PUBLIC_DIR)) {
      sendJson(res, 403, { error: "Forbidden" });
      return;
    }

    sendFile(req, res, filePath, getContentType(filePath));
  } catch (error) {
    console.error(error);
    sendJson(res, 500, {
      error: error && error.message ? error.message : "Server Error"
    });
  }
});

server.listen(PORT, HOST, () => {
  console.log(`ITK Web preview running at http://${HOST}:${PORT}`);
});
