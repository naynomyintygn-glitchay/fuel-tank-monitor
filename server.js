const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const http = require('http');
const socketIo = require('socket.io');
const path = require('path');
const fs = require('fs');
const { createClient } = require('@libsql/client');

const app = express();
const server = http.createServer(app);
const io = socketIo(server, { cors: { origin: "*" } });
const upload = multer({ dest: 'uploads/' });

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

const ADMIN_PASSWORD = "htoo2024";
const PORT = process.env.PORT || 3000;

const db = createClient({
  url: "libsql://fuelmonitor-naynomyintygn-glitchay.aws-ap-northeast-1.turso.io",
  authToken: "eyJhbGciOiJFZERTQSIsInR5cCI6IkpXVCJ9.eyJhIjoicnciLCJpYXQiOjE3ODExODc4MjQsImlkIjoiMDE5ZWI3MGItZGEwMS03NjI0LWE5OTgtZWQ4ZDc2NDk5YjliIiwicmlkIjoiZmQxMDhjYjItMGI1ZC00N2IxLTg1MjUtODVkOGEyMjFkMzg3In0.YdLgQQnyL9qx6h1hxUOrU8ZB1pOU9CpP2PX473ssUtLScZo7vLI8PX1EAw_Yc4fompj6eIIMpeIdWXoXv0nAAA"
});

const defaultStationData = {
  name: "Htoo Fuel Station",
  branchName: "ပြင်ဦးလွင်",
  phoneNumber: "09-123456789",
  logoUrl: "[placehold.co](https://placehold.co/100x100/orange/white?text=HTOO)"
};

const defaultTanksData = [
  { tankNumber: 1, fuelType: "HSD", currentCM: "220", currentLiter: "22500", maxCapacity: "30500" },
  { tankNumber: 2, fuelType: "Premium Diesel", currentCM: "220", currentLiter: "22500", maxCapacity: "30500" },
  { tankNumber: 3, fuelType: "92 RON", currentCM: "220", currentLiter: "22500", maxCapacity: "30500" },
  { tankNumber: 4, fuelType: "92 RON", currentCM: "220", currentLiter: "22500", maxCapacity: "30500" },
  { tankNumber: 5, fuelType: "95 RON", currentCM: "142", currentLiter: "11000", maxCapacity: "15000" },
  { tankNumber: 6, fuelType: "92 RON", currentCM: "142", currentLiter: "11000", maxCapacity: "15000" }
];

let systemDatabase = {
  lastUpdated: "မရှိသေးပါ",
  station: { ...defaultStationData },
  tanks: [...defaultTanksData]
};

function getNowString() {
  const now = new Date();
  return now.toLocaleDateString('my-MM') + " - " + now.toLocaleTimeString('my-MM');
}

async function initDatabase() {
  await db.execute(`
    CREATE TABLE IF NOT EXISTS station (
      id TEXT PRIMARY KEY,
      name TEXT,
      branchName TEXT,
      phoneNumber TEXT,
      logoUrl TEXT,
      lastUpdated TEXT
    )
  `);

  await db.execute(`
    CREATE TABLE IF NOT EXISTS tanks (
      tankNumber INTEGER PRIMARY KEY,
      fuelType TEXT,
      currentCM TEXT,
      currentLiter TEXT,
      maxCapacity TEXT
    )
  `);

  const stationResult = await db.execute("SELECT * FROM station WHERE id = 'station_settings'");
  if (stationResult.rows.length === 0) {
    await db.execute({
      sql: "INSERT INTO station (id, name, branchName, phoneNumber, logoUrl, lastUpdated) VALUES (?, ?, ?, ?, ?, ?)",
      args: [
        "station_settings",
        defaultStationData.name,
        defaultStationData.branchName,
        defaultStationData.phoneNumber,
        defaultStationData.logoUrl,
        "မရှိသေးပါ"
      ]
    });
  }

  const tanksResult = await db.execute("SELECT COUNT(*) as count FROM tanks");
  if (tanksResult.rows[0].count === 0) {
    for (const tank of defaultTanksData) {
      await db.execute({
        sql: "INSERT INTO tanks (tankNumber, fuelType, currentCM, currentLiter, maxCapacity) VALUES (?, ?, ?, ?, ?)",
        args: [tank.tankNumber, tank.fuelType, tank.currentCM, tank.currentLiter, tank.maxCapacity]
      });
    }
  } else {
    for (const tank of defaultTanksData) {
      const exists = await db.execute({
        sql: "SELECT tankNumber FROM tanks WHERE tankNumber = ?",
        args: [tank.tankNumber]
      });
      if (exists.rows.length === 0) {
        await db.execute({
          sql: "INSERT INTO tanks (tankNumber, fuelType, currentCM, currentLiter, maxCapacity) VALUES (?, ?, ?, ?, ?)",
          args: [tank.tankNumber, tank.fuelType, tank.currentCM, tank.currentLiter, tank.maxCapacity]
        });
      }
    }
  }
}

async function loadDatabaseFromTurso() {
  const stationResult = await db.execute("SELECT * FROM station WHERE id = 'station_settings'");
  const tanksResult = await db.execute("SELECT * FROM tanks ORDER BY tankNumber ASC");

  if (stationResult.rows.length > 0) {
    const s = stationResult.rows[0];
    systemDatabase.station = {
      name: s.name,
      branchName: s.branchName,
      phoneNumber: s.phoneNumber,
      logoUrl: s.logoUrl
    };
    systemDatabase.lastUpdated = s.lastUpdated || "မရှိသေးပါ";
  }

  if (tanksResult.rows.length > 0) {
    systemDatabase.tanks = tanksResult.rows.map(t => ({
      tankNumber: t.tankNumber,
      fuelType: t.fuelType,
      currentCM: t.currentCM,
      currentLiter: t.currentLiter,
      maxCapacity: t.maxCapacity
    }));
  }
}

async function updateAndEmit() {
  await db.execute({
    sql: `INSERT INTO station (id, name, branchName, phoneNumber, logoUrl, lastUpdated)
          VALUES (?, ?, ?, ?, ?, ?)
          ON CONFLICT(id) DO UPDATE SET
          name=excluded.name,
          branchName=excluded.branchName,
          phoneNumber=excluded.phoneNumber,
          logoUrl=excluded.logoUrl,
          lastUpdated=excluded.lastUpdated`,
    args: [
      "station_settings",
      systemDatabase.station.name,
      systemDatabase.station.branchName,
      systemDatabase.station.phoneNumber,
      systemDatabase.station.logoUrl,
      systemDatabase.lastUpdated
    ]
  });

  for (const tank of systemDatabase.tanks) {
    await db.execute({
      sql: `INSERT INTO tanks (tankNumber, fuelType, currentCM, currentLiter, maxCapacity)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(tankNumber) DO UPDATE SET
            fuelType=excluded.fuelType,
            currentCM=excluded.currentCM,
            currentLiter=excluded.currentLiter,
            maxCapacity=excluded.maxCapacity`,
      args: [
        Number(tank.tankNumber),
        String(tank.fuelType || ""),
        String(tank.currentCM || "0"),
        String(tank.currentLiter || "0"),
        String(tank.maxCapacity || "0")
      ]
    });
  }

  await loadDatabaseFromTurso();
  io.emit("dataUpdate", systemDatabase);
}

function autoCorrectValue(val) {
  if (val === null || val === undefined || val === "") return null;
  let str = String(val).trim();
  if (str.includes("1900-") || (str.includes("T") && str.includes("Z"))) return null;
  str = str.replace(/\.{2,}/g, '.');
  str = str.replace(/,/g, '');
  str = str.replace(/[^0-9.-]/g, '');
  const parts = str.split('.');
  if (parts.length > 2) str = parts[0] + '.' + parts.slice(1).join('');
  if (!str || str === '.' || str === '-') return null;
  const num = parseFloat(str);
  return Number.isNaN(num) ? null : num;
}

function parseSheetData(matrixData) {
  let lastValidCM = "0";
  let lastValidLiter = "0";
  let detectedCapacity = "30500";
  let detectedProduct = "Unknown";
  let tankNumber = null;

  for (let i = 0; i < Math.min(matrixData.length, 20); i++) {
    const row = matrixData[i];
    if (!row || row.length === 0) continue;
    const lineStr = row.map(cell => String(cell || "").toUpperCase()).join(" ");

    const capacityMatch = lineStr.match(/\(?\s*(\d{1,3}(?:,?\d{3})*)\s*\)?\s*LITER/i);
    if (capacityMatch?.[1]) detectedCapacity = capacityMatch[1].replace(/,/g, '');

    const tankMatch = lineStr.match(/TANK\s*(?:NO\.?|NUMBER)?\s*\(?(\d+)\)?/i);
    if (tankMatch?.[1]) tankNumber = parseInt(tankMatch[1], 10);

    if (lineStr.includes("95 RON")) detectedProduct = "95 RON";
    else if (lineStr.includes("92 RON")) detectedProduct = "92 RON";
    else if (lineStr.includes("PREMIUM DIESEL") || lineStr.includes("PDO")) detectedProduct = "Premium Diesel";
    else if (lineStr.includes("HSD") || (lineStr.includes("DIESEL") && !lineStr.includes("PREMIUM"))) detectedProduct = "HSD";
  }

  for (let i = matrixData.length - 1; i >= 0; i--) {
    const row = matrixData[i];
    if (!row || row.length < 3) continue;

    for (const [cmCol, literCol] of [[1,2],[4,5],[7,8]]) {
      if (row.length <= literCol) continue;
      const cmVal = autoCorrectValue(row[cmCol]);
      const literVal = autoCorrectValue(row[literCol]);
      if (cmVal && cmVal > 0 && cmVal < 300 && literVal !== null && literVal >= 0 && literVal <= 50000) {
        lastValidCM = String(cmVal.toFixed(1));
        lastValidLiter = String(Math.round(literVal));
        break;
      }
    }
    if (lastValidCM !== "0") break;
  }

  return { cm: lastValidCM, liter: lastValidLiter, capacity: detectedCapacity, fuelType: detectedProduct, tankNumber };
}

app.post('/api/verify-password', (req, res) => {
  res.json({ success: req.body.password === ADMIN_PASSWORD });
});

app.get('/api/data', async (req, res) => {
  try {
    await loadDatabaseFromTurso();
    res.json(systemDatabase);
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/manual-save', async (req, res) => {
  try {
    const { station, tanks, password } = req.body;
    if (password !== ADMIN_PASSWORD) {
      return res.status(401).json({ success: false, message: "Password မှားနေပါသည်" });
    }
    systemDatabase.lastUpdated = getNowString();
    if (station) {
      systemDatabase.station = {
        ...systemDatabase.station,
        name: station.name ?? systemDatabase.station.name,
        branchName: station.branchName ?? systemDatabase.station.branchName,
        phoneNumber: station.phoneNumber ?? systemDatabase.station.phoneNumber,
        logoUrl: station.logoUrl ?? systemDatabase.station.logoUrl
      };
    }
    if (Array.isArray(tanks)) {
      systemDatabase.tanks = tanks.map(tank => ({
        tankNumber: Number(tank.tankNumber),
        fuelType: String(tank.fuelType || ""),
        currentCM: String(tank.currentCM || "0"),
        currentLiter: String(tank.currentLiter || "0"),
        maxCapacity: String(tank.maxCapacity || "0")
      }));
    }
    await updateAndEmit();
    res.json({ success: true, data: systemDatabase });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/excel-upload', upload.single('excelFile'), async (req, res) => {
  try {
    const { password, tankNumber: selectedTankNum } = req.body;
    if (password !== ADMIN_PASSWORD) {
      return res.status(401).json({ success: false, message: "Password မှားနေပါသည်" });
    }
    if (!req.file) {
      return res.status(400).json({ success: false, message: "No Excel file uploaded." });
    }

    const workbook = xlsx.readFile(req.file.path);
    const sheetNames = workbook.SheetNames;
    const tempTanks = [...systemDatabase.tanks];
    const sheetsToProcess = [];

    if (selectedTankNum && selectedTankNum !== 'all') {
      if (sheetNames.length > 0) {
        sheetsToProcess.push({ sheet: workbook.Sheets[sheetNames[0]], targetTank: parseInt(selectedTankNum, 10) });
      }
    } else {
      for (let i = 0; i < Math.min(sheetNames.length, 6); i++) {
        sheetsToProcess.push({ sheet: workbook.Sheets[sheetNames[i]], targetTank: i + 1 });
      }
    }

    for (const { sheet, targetTank } of sheetsToProcess) {
      const matrixData = xlsx.utils.sheet_to_json(sheet, { header: 1, raw: false });
      const parsed = parseSheetData(matrixData);
      const finalTankNum = parsed.tankNumber || targetTank;
      const tankIndex = tempTanks.findIndex(t => Number(t.tankNumber) === Number(finalTankNum));
      if (tankIndex !== -1) {
        if (parsed.cm !== "0") tempTanks[tankIndex].currentCM = parsed.cm;
        if (parsed.liter !== "0") tempTanks[tankIndex].currentLiter = parsed.liter;
        if (parsed.capacity !== "30500") tempTanks[tankIndex].maxCapacity = parsed.capacity;
        if (parsed.fuelType !== "Unknown") tempTanks[tankIndex].fuelType = parsed.fuelType;
      }
    }

    systemDatabase.tanks = tempTanks;
    systemDatabase.lastUpdated = getNowString();
    await updateAndEmit();
    fs.unlink(req.file.path, () => {});
    res.json({ success: true, data: systemDatabase });
  } catch (err) {
    if (req.file?.path) fs.unlink(req.file.path, () => {});
    res.status(500).json({ success: false, message: err.message });
  }
});

io.on('connection', socket => {
  socket.emit('dataUpdate', systemDatabase);
});

async function startServer() {
  try {
    console.log("Connecting to Turso database...");
    await initDatabase();
    await loadDatabaseFromTurso();
    console.log("Turso database connected and loaded.");

    server.listen(PORT, () => {
      console.log(`Server running on port ${PORT}`);
    });
  } catch (err) {
    console.error("Failed to start server:", err);
    process.exit(1);
  }
}

startServer();
