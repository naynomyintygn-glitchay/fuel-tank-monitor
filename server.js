const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const http = require('http');
const socketIo = require('socket.io');
const path = require('path');
const fs = require('fs');
const mongoose = require('mongoose');

const app = express();
const server = http.createServer(app);
const io = socketIo(server, {
  cors: { origin: "*" }
});

const upload = multer({ dest: 'uploads/' });

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

const ADMIN_PASSWORD = "htoo2024";
const PORT = process.env.PORT || 3000;
const MONGODB_URI = process.env.MONGODB_URI;

if (!MONGODB_URI) {
  console.error("MONGODB_URI is missing in Render Environment Variables");
  process.exit(1);
}

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

const tankSchema = new mongoose.Schema({
  tankNumber: { type: Number, required: true, unique: true },
  fuelType: { type: String, default: "" },
  currentCM: { type: String, default: "0" },
  currentLiter: { type: String, default: "0" },
  maxCapacity: { type: String, default: "0" }
});

const stationSchema = new mongoose.Schema({
  _id: { type: String, default: "station_settings" },
  name: { type: String, default: defaultStationData.name },
  branchName: { type: String, default: defaultStationData.branchName },
  phoneNumber: { type: String, default: defaultStationData.phoneNumber },
  logoUrl: { type: String, default: defaultStationData.logoUrl },
  lastUpdated: { type: String, default: "မရှိသေးပါ" }
});

const Tank = mongoose.model('Tank', tankSchema);
const Station = mongoose.model('Station', stationSchema);

let systemDatabase = {
  lastUpdated: "မရှိသေးပါ",
  station: defaultStationData,
  tanks: defaultTanksData
};

function getNowString() {
  const now = new Date();
  return now.toLocaleDateString('my-MM') + " - " + now.toLocaleTimeString('my-MM');
}

async function seedDatabaseIfNeeded() {
  let stationDoc = await Station.findById("station_settings");
  if (!stationDoc) {
    await Station.create({
      _id: "station_settings",
      ...defaultStationData,
      lastUpdated: "မရှိသေးပါ"
    });
  }

  const tankCount = await Tank.countDocuments();
  if (tankCount === 0) {
    await Tank.insertMany(defaultTanksData);
  } else {
    for (const tank of defaultTanksData) {
      const exists = await Tank.findOne({ tankNumber: tank.tankNumber });
      if (!exists) {
        await Tank.create(tank);
      }
    }
  }
}

async function loadDatabaseFromMongo() {
  const stationDoc = await Station.findById("station_settings");
  const tanksDocs = await Tank.find({}).sort({ tankNumber: 1 });

  systemDatabase.station = stationDoc
    ? {
        name: stationDoc.name,
        branchName: stationDoc.branchName,
        phoneNumber: stationDoc.phoneNumber,
        logoUrl: stationDoc.logoUrl
      }
    : defaultStationData;

  systemDatabase.tanks = tanksDocs.length
    ? tanksDocs.map(t => ({
        tankNumber: t.tankNumber,
        fuelType: t.fuelType,
        currentCM: t.currentCM,
        currentLiter: t.currentLiter,
        maxCapacity: t.maxCapacity
      }))
    : defaultTanksData;

  systemDatabase.lastUpdated = stationDoc?.lastUpdated || "မရှိသေးပါ";
}

async function updateAndEmit() {
  await Station.findByIdAndUpdate(
    "station_settings",
    {
      _id: "station_settings",
      name: systemDatabase.station.name,
      branchName: systemDatabase.station.branchName,
      phoneNumber: systemDatabase.station.phoneNumber,
      logoUrl: systemDatabase.station.logoUrl,
      lastUpdated: systemDatabase.lastUpdated
    },
    { upsert: true, new: true }
  );

  for (const tank of systemDatabase.tanks) {
    await Tank.findOneAndUpdate(
      { tankNumber: Number(tank.tankNumber) },
      {
        tankNumber: Number(tank.tankNumber),
        fuelType: String(tank.fuelType || ""),
        currentCM: String(tank.currentCM || "0"),
        currentLiter: String(tank.currentLiter || "0"),
        maxCapacity: String(tank.maxCapacity || "0")
      },
      { upsert: true, new: true }
    );
  }

  await loadDatabaseFromMongo();
  io.emit("dataUpdate", systemDatabase);
}

function autoCorrectValue(val) {
  if (val === null || val === undefined || val === "") return null;

  let str = String(val).trim();

  if (str.includes("1900-") || (str.includes("T") && str.includes("Z"))) {
    return null;
  }

  str = str.replace(/\.{2,}/g, '.');
  str = str.replace(/,/g, '');
  str = str.replace(/[^0-9.-]/g, '');

  const parts = str.split('.');
  if (parts.length > 2) {
    str = parts[0] + '.' + parts.slice(1).join('');
  }

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
    if (capacityMatch?.[1]) {
      detectedCapacity = capacityMatch[1].replace(/,/g, '');
    }

    const tankMatch = lineStr.match(/TANK\s*(?:NO\.?|NUMBER)?\s*\(?(\d+)\)?/i);
    if (tankMatch?.[1]) {
      tankNumber = parseInt(tankMatch[1], 10);
    }

    if (lineStr.includes("95 RON")) detectedProduct = "95 RON";
    else if (lineStr.includes("92 RON")) detectedProduct = "92 RON";
    else if (lineStr.includes("PREMIUM DIESEL") || lineStr.includes("PDO")) detectedProduct = "Premium Diesel";
    else if (lineStr.includes("HSD") || (lineStr.includes("DIESEL") && !lineStr.includes("PREMIUM"))) detectedProduct = "HSD";
  }

  for (let i = matrixData.length - 1; i >= 0; i--) {
    const row = matrixData[i];
    if (!row || row.length < 3) continue;

    const pairs = [
      [1, 2],
      [4, 5],
      [7, 8]
    ];

    for (const [cmCol, literCol] of pairs) {
      if (row.length <= literCol) continue;

      const cmVal = autoCorrectValue(row[cmCol]);
      const literVal = autoCorrectValue(row[literCol]);

      const isCmValid = cmVal !== null && cmVal > 0 && cmVal < 300;
      const isLiterValid = literVal !== null && literVal >= 0 && literVal <= 50000;

      if (isCmValid && isLiterValid) {
        lastValidCM = String(cmVal.toFixed(1));
        lastValidLiter = String(Math.round(literVal));
        break;
      }
    }

    if (lastValidCM !== "0") break;
  }

  return {
    cm: lastValidCM,
    liter: lastValidLiter,
    capacity: detectedCapacity,
    fuelType: detectedProduct,
    tankNumber
  };
}

app.post('/api/verify-password', (req, res) => {
  const { password } = req.body;
  res.json({ success: password === ADMIN_PASSWORD });
});

app.get('/api/data', async (req, res) => {
  try {
    await loadDatabaseFromMongo();
    res.json(systemDatabase);
  } catch (err) {
    console.error("GET /api/data error:", err);
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
    console.error("POST /api/manual-save error:", err);
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
        sheetsToProcess.push({
          sheet: workbook.Sheets[sheetNames[0]],
          targetTank: parseInt(selectedTankNum, 10)
        });
      }
    } else {
      for (let i = 0; i < Math.min(sheetNames.length, 6); i++) {
        sheetsToProcess.push({
          sheet: workbook.Sheets[sheetNames[i]],
          targetTank: i + 1
        });
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
    console.error("POST /api/excel-upload error:", err);
    if (req.file?.path) fs.unlink(req.file.path, () => {});
    res.status(500).json({ success: false, message: err.message });
  }
});

io.on('connection', socket => {
  socket.emit('dataUpdate', systemDatabase);
});

mongoose.connect(MONGODB_URI)
  .then(async () => {
    console.log("MongoDB connected");
    await seedDatabaseIfNeeded();
    await loadDatabaseFromMongo();

    server.listen(PORT, () => {
      console.log(`Server running on port ${PORT}`);
    });
  })
  .catch(err => {
    console.error("MongoDB connection error:", err);
    process.exit(1);
  });
