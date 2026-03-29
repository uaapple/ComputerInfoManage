const http = require("http");
const fs = require("fs");
const path = require("path");
const { URL } = require("url");

const HOST = process.env.HOST || "0.0.0.0";
const PORT = process.env.PORT || 8099;
const ROOT = __dirname;
const PUBLIC_DIR = path.join(ROOT, "web");
const DATA_DIR = path.join(ROOT, "data");

function sendJson(res, statusCode, payload) {
  const body = JSON.stringify(payload);
  res.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(body),
    "Cache-Control": "no-store"
  });
  res.end(body);
}

function sendFile(res, filePath, contentType) {
  fs.readFile(filePath, (err, data) => {
    if (err) {
      sendJson(res, 404, { error: "Not Found" });
      return;
    }
    res.writeHead(200, { "Content-Type": contentType });
    res.end(data);
  });
}

function readJson(fileName) {
  const filePath = path.join(DATA_DIR, fileName);
  return JSON.parse(fs.readFileSync(filePath, "utf8"));
}

function normalizeDisplay(value) {
  const text = String(value || "").trim();
  if (!text || ["none", "null", "n/a", "na", "暂无"].includes(text.toLowerCase())) {
    return "N/A";
  }
  return text;
}

function loadStore() {
  const computers = readJson("computers.json");
  const colleagues = readJson("colleagues.json");
  const mailBatches = readJson("inventory_mail_batches.json");
  const colleagueMap = new Map(colleagues.map((item) => [String(item.id), item]));

  const hydratedComputers = computers.map((item) => {
    const owner = colleagueMap.get(String(item.owner_id || ""));
    return {
      id: String(item.id),
      computerName: String(item.computer_name || ""),
      serialNumber: String(item.serial_number || ""),
      assetNumber: normalizeDisplay(item.asset_number),
      model: String(item.model || ""),
      macAddress: normalizeDisplay(item.mac_address),
      ownerId: String(item.owner_id || ""),
      ownerName: owner ? String(owner.display_name || "") : "库存",
      remark: String(item.remark || ""),
      updatedAt: String(item.updated_at || ""),
      ownerHistory: Array.isArray(item.owner_history) ? item.owner_history : []
    };
  });

  const hydratedColleagues = colleagues.map((item) => {
    const ownedCount = hydratedComputers.filter((computer) => computer.ownerId === String(item.id)).length;
    const mentor = colleagueMap.get(String(item.mentor_id || ""));
    return {
      id: String(item.id),
      displayName: String(item.display_name || ""),
      pinyin: String(item.pinyin || ""),
      email: String(item.email || ""),
      department: String(item.department || ""),
      employeeType: String(item.employee_type || ""),
      mentorName: mentor ? String(mentor.display_name || "") : "N/A",
      ownedCount
    };
  });

  return { computers: hydratedComputers, colleagues: hydratedColleagues, mailBatches };
}

function filterByKeyword(rows, keyword, fields) {
  const query = String(keyword || "").trim().toLowerCase();
  if (!query) return rows;
  return rows.filter((row) =>
    fields.some((field) => String(row[field] || "").toLowerCase().includes(query))
  );
}

function handleApi(req, res, url) {
  const store = loadStore();

  if (url.pathname === "/api/meta") {
    sendJson(res, 200, {
      currentUser: process.env.USERNAME || "Admin",
      role: "admin",
      systemName: "ITK China 电脑信息管理系统"
    });
    return;
  }

  if (url.pathname === "/api/dashboard") {
    const inUseCount = store.computers.filter((item) => item.ownerId).length;
    const inventoryCount = store.computers.length - inUseCount;
    const internCount = store.colleagues.filter((item) => item.employeeType === "实习生").length;
    const recentUpdates = [...store.computers]
      .sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt)))
      .slice(0, 6);

    sendJson(res, 200, {
      cards: [
        { key: "inUse", label: "在用电脑", value: inUseCount, tone: "blue" },
        { key: "inventory", label: "库存电脑", value: inventoryCount, tone: "teal" },
        { key: "people", label: "人员总数", value: store.colleagues.length, tone: "slate" },
        { key: "interns", label: "实习生数量", value: internCount, tone: "amber" }
      ],
      recentUpdates
    });
    return;
  }

  if (url.pathname === "/api/list") {
    const view = url.searchParams.get("view") || "in-use";
    const q = url.searchParams.get("q") || "";

    if (view === "people") {
      const rows = filterByKeyword(store.colleagues, q, [
        "displayName",
        "pinyin",
        "email",
        "department",
        "employeeType",
        "mentorName"
      ]);
      sendJson(res, 200, { rows });
      return;
    }

    const baseRows =
      view === "inventory"
        ? store.computers.filter((item) => !item.ownerId)
        : store.computers.filter((item) => item.ownerId);

    const rows = filterByKeyword(baseRows, q, [
      "computerName",
      "serialNumber",
      "assetNumber",
      "model",
      "macAddress",
      "ownerName",
      "remark"
    ]);
    sendJson(res, 200, { rows });
    return;
  }

  if (url.pathname === "/api/details") {
    const type = url.searchParams.get("type");
    const id = url.searchParams.get("id");

    if (type === "computer") {
      const item = store.computers.find((computer) => computer.id === id);
      if (!item) {
        sendJson(res, 404, { error: "Not Found" });
        return;
      }
      sendJson(res, 200, item);
      return;
    }

    if (type === "person") {
      const item = store.colleagues.find((person) => person.id === id);
      if (!item) {
        sendJson(res, 404, { error: "Not Found" });
        return;
      }
      sendJson(res, 200, item);
      return;
    }
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

const server = http.createServer((req, res) => {
  const url = new URL(req.url, `http://${req.headers.host}`);
  if (url.pathname.startsWith("/api/")) {
    handleApi(req, res, url);
    return;
  }

  if (url.pathname === "/assets/logo.jpg") {
    sendFile(res, path.join(ROOT, "ITK_Logo_RGB.jpg"), "image/jpeg");
    return;
  }

  const safePath = url.pathname === "/" ? "/index.html" : url.pathname;
  const filePath = path.normalize(path.join(PUBLIC_DIR, safePath));
  if (!filePath.startsWith(PUBLIC_DIR)) {
    sendJson(res, 403, { error: "Forbidden" });
    return;
  }

  sendFile(res, filePath, getContentType(filePath));
});

server.listen(PORT, HOST, () => {
  console.log(`ITK Web preview running at http://${HOST}:${PORT}`);
});
