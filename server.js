// ==============================================================================
// SMART PUSH-UP MULTI-DEVICE API SERVER & WEB SERVER (BÀI 9 + TỰ ĐỘNG XUẤT EXCEL)
// - ROUTING WEB SERVER CHUẨN MỰC (/, /guide, /about, 404) THEO BÀI 9
// - PHỤC VỤ FILE TĨNH & REST API /sessions LƯU LỊCH SỬ TẬP
// - TỰ ĐỘNG CẬP NHẬT VÀ GHI FILE EXCEL (.CSV UTF-8 BOM) CẢ TRÊN MÁY TÍNH VÀ ĐIỆN THOẠI
// - TÍCH HỢP MQTT CLIENT LẮNG NGHE ESP32 THẬT & ẢO
// ==============================================================================

const http = require("http");
const net = require("net");
const ws = require("ws");
const fs = require("fs");
const path = require("path");
const mqtt = require("mqtt");

const PORT = process.env.PORT || 3000;
const SESSIONS_FILE = path.join(__dirname, "sessions.json");
const EXCEL_FILE_BACKEND = path.join(__dirname, "workout_history.csv");
const EXCEL_FILE_ROOT = path.join(__dirname, "..", "workout_history.csv");

const USERS_FILE = path.join(__dirname, "users.json");

function getUsers() {
  try {
    if (fs.existsSync(USERS_FILE)) {
      return JSON.parse(fs.readFileSync(USERS_FILE, "utf8"));
    }
  } catch (err) {
    console.error("Error reading users.json:", err.message);
  }
  return [
    { id: 1, username: "admin", password: "123456", fullName: "Quản Trị Viên (Admin)", role: "admin", height: 175, weight: 70 },
    { id: 2, username: "user", password: "123456", fullName: "Học Viên Tập Luyện", role: "user", height: 170, weight: 65 }
  ];
}

function saveUsers(users) {
  try {
    fs.writeFileSync(USERS_FILE, JSON.stringify(users, null, 2), "utf8");
  } catch (err) {
    console.error("Error saving users.json:", err.message);
  }
}


// Khởi tạo sessions.json nếu chưa tồn tại
if (!fs.existsSync(SESSIONS_FILE)) {
  const initialData = [
    {
      id: 1,
      deviceId: "ESP32_01",
      startTime: new Date(Date.now() - 3600000).toISOString(),
      endTime: new Date(Date.now() - 3600000 + 120000).toISOString(),
      rep: 15,
      validRep: 14,
      invalidRep: 1,
      accuracy: 93.3,
      cadence: 27.5
    }
  ];
  fs.writeFileSync(SESSIONS_FILE, JSON.stringify(initialData, null, 2), "utf8");
}

// ------------------------------------------------------------------------------
// CÁC HÀM XỬ LÝ DỮ LIỆU & XUẤT EXCEL (CSV UTF-8 BOM)
// ------------------------------------------------------------------------------

// Định dạng thời gian theo múi giờ Việt Nam (GMT+7): YYYY-MM-DD HH:mm:ss
function formatDateTimeVN(isoStr) {
  if (!isoStr) return "";
  try {
    const d = new Date(isoStr);
    if (isNaN(d.getTime())) return String(isoStr);
    const pad = (n) => String(n).padStart(2, "0");
    const utc = d.getTime() + d.getTimezoneOffset() * 60000;
    const vn = new Date(utc + 7 * 3600000);
    return (
      vn.getFullYear() +
      "-" +
      pad(vn.getMonth() + 1) +
      "-" +
      pad(vn.getDate()) +
      " " +
      pad(vn.getHours()) +
      ":" +
      pad(vn.getMinutes()) +
      ":" +
      pad(vn.getSeconds())
    );
  } catch (e) {
    return String(isoStr);
  }
}

// Tính thời lượng buổi tập
function calculateDuration(startIso, endIso) {
  if (!startIso || !endIso) return "0s";
  try {
    const s = new Date(startIso).getTime();
    const e = new Date(endIso).getTime();
    const diffSec = Math.max(0, Math.round((e - s) / 1000));
    const m = Math.floor(diffSec / 60);
    const sec = diffSec % 60;
    if (m > 0) return `${m}p ${sec}s`;
    return `${sec}s`;
  } catch (e) {
    return "0s";
  }
}

// Đánh giá xếp loại buổi tập
function getRating(validRep, invalidRep) {
  const total = Number(validRep || 0) + Number(invalidRep || 0);
  if (total === 0) return "Chưa hoàn thành";
  const ratio = (Number(validRep || 0) / total) * 100;
  if (ratio >= 90 && validRep >= 15) return "Kỹ thuật xuất sắc";
  if (ratio >= 80 && validRep >= 8) return "Tốt";
  if (ratio >= 65) return "Đạt yêu cầu";
  return "Cần cải thiện";
}

// Sinh nội dung file Excel (CSV có ký tự UTF-8 BOM \uFEFF hiển thị tiếng Việt hoàn hảo)

// ------------------------------------------------------------------------------
// CHUẨN HÓA VÀ LÀM GIÀU DỮ LIỆU JSON CHI TIẾT (THỂ HIỆN RÕ RÀNG MỌI THÔNG SỐ)
// ------------------------------------------------------------------------------
function enrichSession(s, index = 0) {
  const totalRep = Number(s.rep != null ? s.rep : (Number(s.validRep || 0) + Number(s.invalidRep || 0)));
  const validRep = Number(s.validRep || 0);
  const invalidRep = Number(s.invalidRep || 0);
  const accuracy = (totalRep > 0) ? Number(((validRep / totalRep) * 100).toFixed(1)) : (totalRep === 0 ? 0 : 100);
  const devId = s.deviceId || "ESP32_01";
  const isVirtual = (devId === "ESP32_02" || devId.toLowerCase().includes("virtual") || devId.toLowerCase().includes("02"));
  const devName = isVirtual ? "ESP32 Ảo (Wokwi Simulator)" : "ESP32 Thật (Cảm biến HC-SR04)";
  const devType = isVirtual ? "Mạch Ảo (Virtual)" : "Mạch Thật (Hardware)";
  const startIso = s.startTime || "";
  const endIso = s.endTime || "";
  const durationStr = calculateDuration(startIso, endIso);
  const cadenceVal = Number(s.cadence || 0);
  const calories = Number((validRep * 0.45).toFixed(1));
  const ratingText = getRating(validRep, invalidRep);

  let feedbackText = "Cần cải thiện độ sâu khi hạ ngực để ghi nhận rep chuẩn.";
  if (accuracy >= 90 && validRep >= 15) {
    feedbackText = "Kỹ thuật tuyệt vời! Biên độ chuyển động hoàn hảo, nhịp độ và tư thế đạt chuẩn cao.";
  } else if (accuracy >= 80 && validRep >= 8) {
    feedbackText = "Động tác rất tốt! Kiểm soát cơ thể ổn định trong suốt buổi tập.";
  } else if (accuracy >= 65) {
    feedbackText = "Đạt yêu cầu thể lực. Chú ý hạ ngực sát cảm biến hơn để tăng số rep chuẩn.";
  }

  let durationSec = 0;
  if (startIso && endIso) {
    const sTime = new Date(startIso).getTime();
    const eTime = new Date(endIso).getTime();
    if (!isNaN(sTime) && !isNaN(eTime)) {
      durationSec = Math.max(0, Math.round((eTime - sTime) / 1000));
    }
  }

  const datePart = startIso ? formatDateTimeVN(startIso).split(" ")[0].replace(/-/g, "") : "DATE";
  const idStr = String(s.id || (Date.now() + index)).slice(-4);
  const sessionCode = `PUSHUP-${datePart}-${idStr}`;

  return {
    id: s.id || Date.now(),
    sessionCode: sessionCode,
    athlete: {
      username: s.username || "user",
      fullName: s.fullName || s.username || "Học Viên Tập Luyện"
    },
    device: {
      deviceId: devId,
      deviceName: devName,
      deviceType: devType
    },
    time: {
      startTime: startIso,
      endTime: endIso,
      startTimeFormatted: formatDateTimeVN(startIso),
      endTimeFormatted: formatDateTimeVN(endIso),
      durationSeconds: durationSec,
      durationFormatted: durationStr
    },
    performance: {
      totalRep: totalRep,
      validRep: validRep,
      invalidRep: invalidRep,
      accuracyPercent: accuracy,
      cadenceRpm: Number(cadenceVal.toFixed(1)),
      caloriesBurnedKcal: calories
    },
    evaluation: {
      rating: ratingText,
      feedback: feedbackText
    },
    // Các trường tương thích ngược (Flat Fields) cho Dashboard & Scripts cũ:
    rep: totalRep,
    validRep: validRep,
    invalidRep: invalidRep,
    cadence: Number(cadenceVal.toFixed(1)),
    accuracy: accuracy,
    deviceId: devId,
    username: s.username || "user",
    fullName: s.fullName || s.username || "Học Viên Tập Luyện",
    startTime: startIso,
    endTime: endIso,
    duration: durationStr,
    calories: calories,
    rating: ratingText
  };
}

function generateDetailedJson(sessions, reqUser = "", isAdmin = false) {
  const enriched = (sessions || []).map((s, idx) => enrichSession(s, idx));

  let totalReps = 0;
  let totalValid = 0;
  let totalInvalid = 0;
  let totalCalo = 0;
  let totalCadence = 0;

  enriched.forEach(s => {
    totalReps += s.performance.totalRep;
    totalValid += s.performance.validRep;
    totalInvalid += s.performance.invalidRep;
    totalCalo += s.performance.caloriesBurnedKcal;
    totalCadence += s.performance.cadenceRpm;
  });

  const avgAccuracy = (totalReps > 0) ? Number(((totalValid / totalReps) * 100).toFixed(1)) : 0;
  const avgCadence = (enriched.length > 0) ? Number((totalCadence / enriched.length).toFixed(1)) : 0;

  return {
    system: "Smart Push-Up IoT Analytics Platform",
    version: "2.5.0",
    description: "Báo cáo chi tiết lịch sử tập luyện hít đất thông minh",
    exportedAt: formatDateTimeVN(new Date().toISOString()),
    exportScope: isAdmin ? (reqUser && reqUser !== "all" ? `Học viên @${reqUser}` : "Toàn bộ hệ thống (Admin)") : `Cá nhân học viên @${reqUser || 'user'}`,
    summary: {
      totalSessions: enriched.length,
      totalReps: totalReps,
      totalValidReps: totalValid,
      totalInvalidReps: totalInvalid,
      overallAccuracyPercent: avgAccuracy,
      totalCaloriesBurnedKcal: Number(totalCalo.toFixed(1)),
      averageCadenceRpm: avgCadence
    },
    sessions: enriched
  };
}

function generateExcelCsv(sessions, isAdmin = false) {
  const BOM = "\uFEFF";
  const headers = isAdmin
    ? [
        "STT",
        "Học Viên",
        "Tài Khoản",
        "Mã Buổi Tập",
        "Thiết Bị",
        "Thời Gian Bắt Đầu",
        "Thời Gian Kết Thúc",
        "Thời Lượng",
        "Tổng Số Rep",
        "Rep Chuẩn",
        "Rep Chưa Đạt",
        "Nhịp Độ (rep/phút)",
        "Calo Ước Tính (Kcal)",
        "Đánh Giá"
      ]
    : [
        "STT",
        "Mã Buổi Tập",
        "Thiết Bị",
        "Thời Gian Bắt Đầu",
        "Thời Gian Kết Thúc",
        "Thời Lượng",
        "Tổng Số Rep",
        "Rep Chuẩn",
        "Rep Chưa Đạt",
        "Nhịp Độ (rep/phút)",
        "Calo Ước Tính (Kcal)",
        "Đánh Giá"
      ];

  const escapeCol = (val) => {
    const s = String(val ?? "").replace(/"/g, '""');
    return `"${s}"`;
  };

  const rows = (sessions || []).map((s, idx) => {
    const stt = idx + 1;
    const id = s.id || "";
    const dev = (s.deviceId === "ESP32_02") ? "ESP32 Ảo (02)" : "ESP32 Thật (01)";
    const start = formatDateTimeVN(s.startTime);
    const end = formatDateTimeVN(s.endTime);
    const duration = calculateDuration(s.startTime, s.endTime);
    const totalRep = Number(s.rep || 0);
    const validRep = Number(s.validRep || 0);
    const invalidRep = Number(s.invalidRep || 0);
    const cadence = Number(s.cadence || 0).toFixed(1);
    const calo = (validRep * 0.45).toFixed(1);
    const rating = getRating(validRep, invalidRep);

    if (isAdmin) {
      const athleteName = s.fullName || s.username || "Học Viên";
      const athleteUser = s.username || "user";
      return [
        stt,
        escapeCol(athleteName),
        escapeCol(athleteUser),
        escapeCol(id),
        escapeCol(dev),
        escapeCol(start),
        escapeCol(end),
        escapeCol(duration),
        totalRep,
        validRep,
        invalidRep,
        cadence,
        calo,
        escapeCol(rating)
      ].join(",");
    }

    return [
      stt,
      escapeCol(id),
      escapeCol(dev),
      escapeCol(start),
      escapeCol(end),
      escapeCol(duration),
      totalRep,
      validRep,
      invalidRep,
      cadence,
      calo,
      escapeCol(rating)
    ].join(",");
  });

  return BOM + headers.join(",") + "\r\n" + rows.join("\r\n") + "\r\n";
}

// Đồng bộ và ghi ra file Excel trên ổ cứng máy tính
function syncExcelFiles(sessions) {
  try {
    const csvContent = generateExcelCsv(sessions);
    fs.writeFileSync(EXCEL_FILE_BACKEND, csvContent, "utf8");
    fs.writeFileSync(EXCEL_FILE_ROOT, csvContent, "utf8");
    console.log(`>>> [EXCEL SYNC] Da tu dong cap nhat file Excel: ${EXCEL_FILE_ROOT} (${sessions.length} buoi tap)`);
  } catch (err) {
    console.error(">>> [EXCEL SYNC ERROR]", err.message);
  }
}

// Đọc danh sách buổi tập từ JSON
function getSessions() {
  try {
    const raw = fs.readFileSync(SESSIONS_FILE, "utf8");
    return JSON.parse(raw);
  } catch (err) {
    return [];
  }
}

// Lưu danh sách buổi tập và tự động cập nhật Excel
function saveSessions(sessions) {
  const enrichedList = (sessions || []).map((s, idx) => enrichSession(s, idx));
  fs.writeFileSync(SESSIONS_FILE, JSON.stringify(enrichedList, null, 2), "utf8");
  syncExcelFiles(enrichedList);
}

// Tự động khởi tạo file Excel ngay khi bật server
syncExcelFiles(getSessions());

const activeSessions = {};

// ==============================================================================
// 1. MQTT CLIENT
// ==============================================================================
const MQTT_BROKER = process.env.MQTT_BROKER || "mqtt://broker.hivemq.com:1883";
const mqttClient = mqtt.connect(MQTT_BROKER, {
  reconnectPeriod: 3000
});

mqttClient.on("connect", () => {
  console.log("--------------------------------------------------");
  console.log(">>> [MQTT] API Server da ket noi Broker thanh cong");
  mqttClient.subscribe("pushup_ntp113/+/command");
  mqttClient.subscribe("pushup_ntp113/+/live");
  mqttClient.subscribe("pushup/+/command");
  mqttClient.subscribe("pushup/+/live");
  console.log(">>> [MQTT] Dang lang nghe multi-device: pushup/+/command & pushup/+/live");
  console.log("--------------------------------------------------");
});

mqttClient.on("message", (topic, message) => {
  const text = message.toString();
  const parts = topic.split("/");
  // Format: pushup_ntp113/device01/live -> parts[1] is device01
  const devKey = parts[1] || "device01";
  const devIdFormatted = (devKey === "device02") ? "ESP32_02" : "ESP32_01";

  if (topic.endsWith("/command")) {
    const cmd = text.trim();

    if (cmd.startsWith("START")) {
      let username = "user";
      let fullName = "Học Viên Tập Luyện";
      const userMatch = cmd.match(/USER:([^:\s]+):?([^\r\n]*)/i);
      if (userMatch) {
        username = userMatch[1].trim().toLowerCase();
        fullName = userMatch[2].trim() || username;
      }

      activeSessions[devKey] = {
        id: Date.now(),
        deviceId: devIdFormatted,
        username: username,
        fullName: fullName,
        startTime: new Date().toISOString(),
        endTime: null,
        rep: 0,
        validRep: 0,
        invalidRep: 0,
        cadence: 0.0
      };
      console.log(`\n[START] Bat dau buoi tap cho [${devKey} - ${devIdFormatted}] - Hoc vien: [${fullName} (@${username})] luc: ${new Date().toLocaleTimeString("vi-VN")}`);
    } 
    else if (cmd === "STOP" || cmd === "RESET") {
      const sess = activeSessions[devKey];
      if (sess) {
        sess.endTime = new Date().toISOString();
        if (sess.rep > 0 || cmd === "STOP") {
          const sessions = getSessions();
          sess.username = sess.username || "user";
          sess.fullName = sess.fullName || "Học Viên Tập Luyện";
          delete sess.accuracy;
          sessions.push(sess);
          saveSessions(sessions);
          console.log(`\n[SAVED & EXCEL UPDATED] Da luu buoi tap cua [${sess.fullName} - ${devIdFormatted}]: ${sess.rep} rep (Chuan: ${sess.validRep}) vao sessions.json va workout_history.csv!`);
        }
        delete activeSessions[devKey];
      }
    }
    return;
  }

  if (topic.endsWith("/live")) {
    let sess = activeSessions[devKey];
    try {
      const data = JSON.parse(text);

      // TỰ ĐỘNG PHỤC HỒI BUỔI TẬP: Nếu Wokwi hoặc ESP32 đang tập mà chưa kịp nhận lệnh START
      if (!sess && (data.status === "TRAINING" || (data.rep && data.rep > 0))) {
        activeSessions[devKey] = {
          id: Date.now(),
          deviceId: devIdFormatted,
          username: "user",
          fullName: (devKey === "device02") ? "Mạch Ảo Wokwi" : "Học Viên Tập Luyện",
          startTime: new Date().toISOString(),
          endTime: null,
          rep: data.rep || 0,
          validRep: data.validRep || 0,
          invalidRep: data.invalidRep || 0,
          cadence: Number(data.cadence || 0)
        };
        sess = activeSessions[devKey];
        console.log(`>>> [AUTO-SESSION] Tự động kích hoạt phiên tập cho [${devKey} - ${devIdFormatted}] khi nhận tín hiệu live!`);
      }

      if (sess) {
        sess.rep = data.rep ?? sess.rep;
        sess.validRep = data.validRep ?? sess.validRep;
        sess.invalidRep = data.invalidRep ?? sess.invalidRep;
        sess.cadence = Number(data.cadence ?? sess.cadence);
        if (data.deviceId) sess.deviceId = data.deviceId;
      }
    } catch (e) {}
  }
});

mqttClient.on("error", (err) => console.log("[MQTT Error]", err.message));

// ==============================================================================
// 2. WEB SERVER & ROUTING THEO BÀI 9 (PORT 3000)
// ==============================================================================
const server = http.createServer((req, res) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, DELETE, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") {
    res.writeHead(204);
    res.end();
    return;
  }

  const url = new URL(req.url, `http://${req.headers.host}`);
  console.log(`[HTTP ${req.method}] ${url.pathname}`);

  // ROUTE 1: REST API LẤY LỊCH SỬ BUỔI TẬP (JSON)
  
  // ROUTE 0.1: REST API DANG NHAP (LOGIN)
  if (req.method === "POST" && (url.pathname === "/api/login" || url.pathname === "/login")) {
    let body = "";
    req.on("data", chunk => (body += chunk));
    req.on("end", () => {
      try {
        const { username, password } = JSON.parse(body);
        const u = (username || "").trim().toLowerCase();
        const p = (password || "").trim();
        const users = getUsers();
        const found = users.find(x => x.username.toLowerCase() === u && x.password === p);
        if (found) {
          const safeUser = {
            id: found.id,
            username: found.username,
            fullName: found.fullName || found.username,
            role: found.role || "user",
            height: found.height || 170,
            weight: found.weight || 65
          };
          console.log(`>>> [AUTH] Dang nhap thanh cong: ${found.username} (${found.role})`);
          res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
          return res.end(JSON.stringify({ success: true, message: "Đăng nhập thành công!", user: safeUser }));
        } else {
          res.writeHead(401, { "Content-Type": "application/json; charset=utf-8" });
          return res.end(JSON.stringify({ success: false, error: "Tên tài khoản hoặc mật khẩu không đúng!" }));
        }
      } catch (err) {
        res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
        return res.end(JSON.stringify({ success: false, error: "Dữ liệu không hợp lệ" }));
      }
    });
    return;
  }

  // ROUTE 0.2: REST API DANG KY TAI KHOAN MOI (REGISTER)
  if (req.method === "POST" && (url.pathname === "/api/register" || url.pathname === "/register")) {
    let body = "";
    req.on("data", chunk => (body += chunk));
    req.on("end", () => {
      try {
        const { username, password, fullName, height, weight } = JSON.parse(body);
        const u = (username || "").trim().toLowerCase();
        const p = (password || "").trim();
        const name = (fullName || "").trim() || u;
        const h = Number(height) || 170;
        const w = Number(weight) || 65;

        if (!u || u.length < 3) {
          res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
          return res.end(JSON.stringify({ success: false, error: "Tên đăng nhập phải có ít nhất 3 ký tự!" }));
        }
        if (!p || p.length < 4) {
          res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
          return res.end(JSON.stringify({ success: false, error: "Mật khẩu phải có ít nhất 4 ký tự!" }));
        }

        const users = getUsers();
        if (users.some(x => x.username.toLowerCase() === u)) {
          res.writeHead(409, { "Content-Type": "application/json; charset=utf-8" });
          return res.end(JSON.stringify({ success: false, error: `Tài khoản '${u}' đã tồn tại! Vui lòng chọn tên khác.` }));
        }

        const newUser = {
          id: Date.now(),
          username: u,
          password: p,
          fullName: name,
          role: "user",
          height: h,
          weight: w,
          createdAt: new Date().toISOString()
        };

        users.push(newUser);
        saveUsers(users);

        const safeUser = {
          id: newUser.id,
          username: newUser.username,
          fullName: newUser.fullName,
          role: newUser.role,
          height: newUser.height,
          weight: newUser.weight
        };

        console.log(`>>> [REGISTER] Nguoi dung moi dang ky thanh cong: ${u} (${name})`);
        res.writeHead(201, { "Content-Type": "application/json; charset=utf-8" });
        return res.end(JSON.stringify({ success: true, message: "Đăng ký tài khoản thành công!", user: safeUser }));
      } catch (err) {
        res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
        return res.end(JSON.stringify({ success: false, error: "Dữ liệu đăng ký không hợp lệ!" }));
      }
    });
    return;
  }

  // ROUTE 0.3: REST API LẤY DANH SÁCH USERS (CHO ADMIN LỌC HỌC VIÊN)
  if (req.method === "GET" && (url.pathname === "/api/users" || url.pathname === "/users")) {
    const users = getUsers().map(u => ({
      id: u.id,
      username: u.username,
      fullName: u.fullName || u.username,
      role: u.role || "user",
      height: u.height,
      weight: u.weight
    }));
    res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
    return res.end(JSON.stringify(users));
  }

  // ROUTE 1: REST API LẤY LỊCH SỬ BUỔI TẬP (PHÂN QUYỀN: USER CHỈ XEM CỦA MÌNH, ADMIN XEM HẾT)
  if (req.method === "GET" && url.pathname === "/sessions") {
    const reqUser = (url.searchParams.get("username") || "").trim().toLowerCase();
    const reqRole = (url.searchParams.get("role") || "").trim().toLowerCase();
    const filterUser = (url.searchParams.get("filterUser") || "").trim().toLowerCase();
    let sessions = getSessions();

    if (reqRole !== "admin") {
      // User thường chỉ xem lịch sử của chính mình
      sessions = sessions.filter(s => {
        const u = (s.username || "user").toLowerCase();
        return u === reqUser || (reqUser === "user" && !s.username);
      });
    } else {
      // Admin: có thể xem tất cả hoặc lọc theo 1 học viên cụ thể
      if (filterUser && filterUser !== "all") {
        sessions = sessions.filter(s => {
          const u = (s.username || "user").toLowerCase();
          return u === filterUser;
        });
      }
    }

    res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
    return res.end(JSON.stringify(sessions));
  }

  // ROUTE 2: REST API LƯU BUỔI TẬP MỚI
  if (req.method === "POST" && url.pathname === "/sessions") {
    let body = "";
    req.on("data", chunk => (body += chunk));
    req.on("end", () => {
      try {
        const newSession = JSON.parse(body);
        newSession.id = newSession.id || Date.now();
        if (!newSession.username) newSession.username = "user";
        if (!newSession.fullName) newSession.fullName = "Học Viên Tập Luyện";
        delete newSession.accuracy; // Bỏ độ chính xác
        const sessions = getSessions();
        sessions.push(newSession);
        saveSessions(sessions);
        res.writeHead(201, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ success: true, session: newSession }));
      } catch (err) {
        res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ success: false, error: "Invalid JSON" }));
      }
    });
    return;
  }

  // ROUTE 2.5: REST API RESET / XÓA LỊCH SỬ BUỔI TẬP (DELETE /sessions)
  if (req.method === "DELETE" && url.pathname === "/sessions") {
    const reqUser = (url.searchParams.get("username") || "").trim().toLowerCase();
    const reqRole = (url.searchParams.get("role") || "").trim().toLowerCase();
    const filterUser = (url.searchParams.get("filterUser") || "").trim().toLowerCase();

    let sessions = getSessions();

    if (reqRole === "admin" && (!filterUser || filterUser === "all")) {
      // Admin reset TOÀN BỘ buổi tập của tất cả học viên
      saveSessions([]);
      console.log(`>>> [RESET ALL] Admin da reset toan bo lich su buoi tap cua he thong ve 0!`);
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      return res.end(JSON.stringify({ success: true, message: "Đã reset toàn bộ lịch sử buổi tập của tất cả học viên về 0!" }));
    } else if (reqRole === "admin" && filterUser) {
      // Admin reset riêng 1 học viên
      const before = sessions.length;
      sessions = sessions.filter(s => (s.username || "user").toLowerCase() !== filterUser);
      saveSessions(sessions);
      console.log(`>>> [RESET ATHLETE] Admin da reset ${before - sessions.length} buoi tap cua @${filterUser}`);
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      return res.end(JSON.stringify({ success: true, message: `Đã xóa toàn bộ lịch sử tập luyện của học viên @${filterUser}!` }));
    } else if (reqUser) {
      // User thường chỉ xóa lịch sử của chính mình
      const before = sessions.length;
      sessions = sessions.filter(s => {
        const u = (s.username || "user").toLowerCase();
        return u !== reqUser && !(reqUser === "user" && !s.username);
      });
      saveSessions(sessions);
      console.log(`>>> [RESET USER] Nguoi dung @${reqUser} da reset ${before - sessions.length} buoi tap cua minh!`);
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      return res.end(JSON.stringify({ success: true, message: "Đã xóa lịch sử tập luyện cá nhân của bạn!" }));
    } else {
      res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
      return res.end(JSON.stringify({ success: false, error: "Yêu cầu không hợp lệ!" }));
    }
  }

  // ROUTE 3: XUẤT / TẢI FILE EXCEL (.CSV CHUẨN UTF-8 BOM TƯƠNG THÍCH MỌI EXCEL TRÊN PC VÀ ĐIỆN THOẠI)
  // PHÂN QUYỀN: USER CHỈ TẢI LỊCH SỬ CỦA MÌNH, ADMIN TẢI TOÀN BỘ HOẶC THEO HỌC VIÊN
  if (req.method === "GET" && (url.pathname === "/export-excel" || url.pathname === "/export-csv" || url.pathname === "/workout_history.csv")) {
    const reqUser = (url.searchParams.get("username") || "").trim().toLowerCase();
    const reqRole = (url.searchParams.get("role") || "").trim().toLowerCase();
    const filterUser = (url.searchParams.get("filterUser") || "").trim().toLowerCase();
    let sessions = getSessions();
    const isAdmin = (reqRole === "admin");
    let exportFileName = "SmartPushUp_ToanBoLichSu.csv";

    if (!isAdmin && reqUser) {
      sessions = sessions.filter(s => {
        const u = (s.username || "user").toLowerCase();
        return u === reqUser || (reqUser === "user" && !s.username);
      });
      exportFileName = `SmartPushUp_LichSu_${reqUser}.csv`;
    } else if (isAdmin && filterUser && filterUser !== "all") {
      sessions = sessions.filter(s => {
        const u = (s.username || "user").toLowerCase();
        return u === filterUser;
      });
      exportFileName = `SmartPushUp_LichSu_${filterUser}.csv`;
    }

    const csvContent = generateExcelCsv(sessions, isAdmin);
    res.writeHead(200, {
      "Content-Type": "text/csv; charset=utf-8",
      "Content-Disposition": `attachment; filename="${exportFileName}"`,
      "Cache-Control": "no-cache, no-store, must-revalidate"
    });
    res.end(csvContent);
    console.log(`>>> [EXCEL DOWNLOAD] Da xuat file ${exportFileName} (${sessions.length} buoi tap) cho [${reqUser || 'admin'}]`);
    return;
  }

  
  // ROUTE 3.5: XUẤT / TẢI FILE JSON CHI TIẾT ĐẦY ĐỦ THÔNG SỐ (THỂ HIỆN RÕ RÀNG MỌI THÔNG TIN)
  if (req.method === "GET" && (url.pathname === "/export-json" || url.pathname === "/sessions.json")) {
    const reqUser = (url.searchParams.get("username") || "").trim().toLowerCase();
    const reqRole = (url.searchParams.get("role") || "").trim().toLowerCase();
    const filterUser = (url.searchParams.get("filterUser") || "").trim().toLowerCase();
    let sessions = getSessions();
    const isAdmin = (reqRole === "admin");
    let exportFileName = "SmartPushUp_ToanBoLichSu.json";
    let targetUser = reqUser;

    if (!isAdmin && reqUser) {
      sessions = sessions.filter(s => {
        const u = (s.username || "user").toLowerCase();
        return u === reqUser || (reqUser === "user" && !s.username);
      });
      exportFileName = `SmartPushUp_LichSu_${reqUser}.json`;
    } else if (isAdmin && filterUser && filterUser !== "all") {
      sessions = sessions.filter(s => {
        const u = (s.username || "user").toLowerCase();
        return u === filterUser;
      });
      exportFileName = `SmartPushUp_LichSu_${filterUser}.json`;
      targetUser = filterUser;
    }

    const jsonData = generateDetailedJson(sessions, targetUser, isAdmin);
    const jsonStr = JSON.stringify(jsonData, null, 2);

    res.writeHead(200, {
      "Content-Type": "application/json; charset=utf-8",
      "Content-Disposition": `attachment; filename="${exportFileName}"`,
      "Cache-Control": "no-cache, no-store, must-revalidate"
    });
    res.end(jsonStr);
    console.log(`>>> [JSON DOWNLOAD] Đã xuất file JSON chi tiết: ${exportFileName} (${sessions.length} buổi tập) cho [${reqUser || 'admin'}]`);
    return;
  }

  // ROUTE 4: XUẤT DỮ LIỆU CẢM BIẾN LIVE (NẾU CÓ pushup_data.csv)
  if (req.method === "GET" && url.pathname === "/export-live-csv") {
    const liveCsvPath = path.join(__dirname, "pushup_data.csv");
    if (fs.existsSync(liveCsvPath)) {
      res.writeHead(200, {
        "Content-Type": "text/csv; charset=utf-8",
        "Content-Disposition": 'attachment; filename="SmartPushUp_DuLieuCamBien_Live.csv"'
      });
      return res.end(fs.readFileSync(liveCsvPath));
    } else {
      res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      return res.end("Chua co du lieu cam bien live.");
    }
  }

  // ROUTE 5 (BÀI 9): TRANG CHỦ DASHBOARD (/)
  if (req.method === "GET" && (url.pathname === "/" || url.pathname === "/index.html")) {
    const htmlFile = path.join(__dirname, "public", "index.html");
    if (fs.existsSync(htmlFile)) {
      res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
      return res.end(fs.readFileSync(htmlFile));
    }
  }

  // ROUTE 6 (BÀI 9): TRANG HƯỚNG DẪN & CÀI ĐẶT (/guide hoặc /about)
  if (req.method === "GET" && (url.pathname === "/guide" || url.pathname === "/about")) {
    const guideFile = path.join(__dirname, "public", "guide.html");
    if (fs.existsSync(guideFile)) {
      res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
      return res.end(fs.readFileSync(guideFile));
    }
  }

  // ROUTE 6.5: PHỤC VỤ STATIC FILES CHO PWA & ASSETS (/manifest.json, /icon-192.jpg, /icon-512.jpg, /apple-touch-icon.png, v.v.)
  const dashboardDir = path.join(__dirname, "public");
  const safeFilename = path.basename(url.pathname);
  const potentialFile = path.join(dashboardDir, safeFilename);

  if (req.method === "GET" && fs.existsSync(potentialFile) && fs.statSync(potentialFile).isFile()) {
    const ext = path.extname(potentialFile).toLowerCase();
    const mimeTypes = {
      ".json": "application/json; charset=utf-8",
      ".jpg": "image/jpeg",
      ".jpeg": "image/jpeg",
      ".png": "image/png",
      ".svg": "image/svg+xml",
      ".ico": "image/x-icon",
      ".js": "application/javascript",
      ".css": "text/css"
    };
    const contentType = mimeTypes[ext] || "application/octet-stream";
    res.writeHead(200, { "Content-Type": contentType });
    return res.end(fs.readFileSync(potentialFile));
  }

  // ROUTE 7 (BÀI 9): BẮT LỖI 404 NẾU GÕ SAI ĐƯỜNG DẪN
  res.writeHead(404, { "Content-Type": "text/html; charset=utf-8" });
  res.end(`
    <!DOCTYPE html><html><head><meta charset="UTF-8"><title>404 Not Found</title></head>
    <body style="font-family: Arial; text-align: center; padding: 50px; background: #0f172a; color: white;">
      <h1 style="color: #ef4444; font-size: 45px;">⚠️ 404 - Trang Không Tồn Tại</h1>
      <p style="color: #94a3b8;">Đường dẫn <code>${url.pathname}</code> không tồn tại trên Smart Push-Up Server.</p>
      <br><a href="/" style="color: #38bdf8; text-decoration: none; font-weight: bold; border: 1px solid #38bdf8; padding: 10px 20px; border-radius: 8px;">&larr; Về Dashboard Tập Luyện</a>
    </body></html>
  `);
});


// ==============================================================================

server.listen(PORT, () => {
  console.log("==============================================================");
  console.log("🚀 SMART PUSH-UP 24/7 CLOUD SERVER IS RUNNING!");
  console.log("   - Port           : " + PORT);
  console.log("   - Cloud Broker   : " + MQTT_BROKER);
  console.log("   - Active Prefix  : pushup_ntp113");
  console.log("==============================================================");
});
