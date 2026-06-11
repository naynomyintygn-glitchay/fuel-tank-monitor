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
app.use(express.static(path.join(__dirname, 'public')));

const ADMIN_PASSWORD = "htoo2024";
const MONGODB_URI = process.env.MONGODB_URI;

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
    fuelType: String,
    currentCM: String,
    currentLiter: String,
    maxCapacity: String
});

const stationSchema = new mongoose.Schema({
    _id: { type: String, default: 'station_settings' },
    name: String,
    branchName: String,
    phoneNumber: String,
    logoUrl: String,
    lastUpdated: String
});

const Tank = mongoose.model('Tank', tankSchema);
const Station = mongoose.model('Station', stationSchema);

let systemDatabase = {
    lastUpdated: "မရှိသေးပါ",
    station: defaultStationData,
    tanks: defaultTanksData
};

async function loadDatabaseFromMongo() {
    try {
        const stationDoc = await Station.findById('station_settings');
        const tanksDocs = await Tank.find({}).sort({ tankNumber: 1 });

        if (!stationDoc) {
            const newStation = await Station.create({
                _id: 'station_settings',
                ...defaultStationData,
                lastUpdated: "မရှိသေးပါ"
            });
            systemDatabase.station = newStation.toObject();
        } else {
            systemDatabase.station = stationDoc.toObject();
        }

        if (tanksDocs.length === 0) {
            const newTanks = await Tank.insertMany(defaultTanksData);
            systemDatabase.tanks = newTanks.map(doc => doc.toObject());
        } else {
            systemDatabase.tanks = tanksDocs.map(doc => doc.toObject());
        }

        systemDatabase.lastUpdated = systemDatabase.station.lastUpdated || "မရှိသေးပါ";
    } catch (err) {
        console.error('Load MongoDB error:', err);
    }
}

async function updateAndEmit() {
    let stationDoc = await Station.findById('station_settings');

    if (!stationDoc) {
        stationDoc = new Station({
            _id: 'station_settings',
            ...systemDatabase.station,
            lastUpdated: systemDatabase.lastUpdated
        });
    } else {
        stationDoc.name = systemDatabase.station.name;
        stationDoc.branchName = systemDatabase.station.branchName;
        stationDoc.phoneNumber = systemDatabase.station.phoneNumber;
        stationDoc.logoUrl = systemDatabase.station.logoUrl;
        stationDoc.lastUpdated = systemDatabase.lastUpdated;
    }

    await stationDoc.save();

    for (const tank of systemDatabase.tanks) {
        await Tank.updateOne(
            { tankNumber: tank.tankNumber },
            { $set: tank },
            { upsert: true }
        );
    }

    await loadDatabaseFromMongo();
    io.emit('dataUpdate', systemDatabase);
}

function autoCorrectValue(val) {
    if (val === null || val === undefined || val === '') return null;

    let str = String(val).trim();

    if (str.includes('1900-') || (str.includes('T') && str.includes('Z'))) {
        return null;
    }

    str = str.replace(/\.{2,}/g, '.');
    str = str.replace(/,/g, '');
    str = str.replace(/[^0-9.-]/g, '');

    const parts = str.split('.');
    if (parts.length > 2) {
        str = parts[0] + '.' + parts.slice(1).join('');
    }

    if (str === '' || str === '.' || str === '-') return null;

    const num = parseFloat(str);
    return isNaN(num) ? null : num;
}

function parseSheetData(matrixData) {
    let lastValidCM = "0";
    let lastValidLiter = "0";
    let detectedCapacity = "30500";
    let detectedProduct = "Unknown";
    let tankNumber = null;

    for (let i = 0; i < Math.min(matrixData.length, 15); i++) {
        const row = matrixData[i];
        if (!row || row.length === 0) continue;

        const lineStr = row.map(cell => String(cell || '').toUpperCase()).join(" ");

        const capacityMatch = lineStr.match(/\(?\s*(\d{1,3}(?:,?\d{3})*)\s*\)?\s*LITER/i);
        if (capacityMatch && capacityMatch[1]) {
            detectedCapacity = capacityMatch[1].replace(/,/g, '');
        }

        const tankMatch = lineStr.match(/TANK\s*(?:NO\.?|NUMBER)?\s*\(?(\d+)\)?/i);
        if (tankMatch && tankMatch[1]) {
            tankNumber = parseInt(tankMatch[1]);
        }

        if (lineStr.includes("95 RON")) detectedProduct = "95 RON";
        else if (lineStr.includes("92 RON")) detectedProduct = "92 RON";
        else if (lineStr.includes("PREMIUM DIESEL") || lineStr.includes("PDO")) detectedProduct = "Premium Diesel";
        else if (lineStr.includes("HSD") || (lineStr.includes("DIESEL") && !lineStr.includes("PREMIUM"))) detectedProduct = "HSD";
    }

    for (let i = matrixData.length - 1; i >= 0; i--) {
        const currentRow = matrixData[i];
        if (!currentRow || currentRow.length < 3) continue;

        const possibleCmCols = [1, 4, 7];
        const possibleLiterCols = [2, 5, 8];

        for (let j = 0; j < possibleCmCols.length; j++) {
            const cmCol = possibleCmCols[j];
            const literCol = possibleLiterCols[j];

            if (currentRow.length > cmCol && currentRow.length > literCol) {
                const cmVal = autoCorrectValue(currentRow[cmCol]);
                const literVal = autoCorrectValue(currentRow[literCol]);

                const isCmReasonable = cmVal !== null && cmVal > 0 && cmVal < 300;
                const isLiterReasonable = literVal !== null && literVal >= 0 && literVal <= 50000;

                if (isCmReasonable && isLiterReasonable) {
                    lastValidCM = String(cmVal.toFixed(1));
                    lastValidLiter = String(literVal.toFixed(0));
                    break;
                }
            }
        }

        if (lastValidCM !== "0") break;
    }

    return {
        cm: lastValidCM,
        liter: lastValidLiter,
        capacity: detectedCapacity,
        fuelType: detectedProduct,
        tankNumber: tankNumber
    };
}

app.post('/api/verify-password', (req, res) => {
    const { password } = req.body;
    res.json({ success: password === ADMIN_PASSWORD });
});

app.get('/api/data', (req, res) => {
    res.json(systemDatabase);
});

app.post('/api/manual-save', async (req, res) => {
    try {
        const { station, tanks, password } = req.body;

        if (password !== ADMIN_PASSWORD) {
            return res.status(401).json({ success: false, message: "Password မှားနေပါသည်" });
        }

        const now = new Date();
        systemDatabase.lastUpdated = now.toLocaleDateString('my-MM') + " - " + now.toLocaleTimeString('my-MM');

        if (station) {
            systemDatabase.station = {
                ...systemDatabase.station,
                ...station
            };
        }

        if (tanks && Array.isArray(tanks)) {
            systemDatabase.tanks = tanks;
        }

        await updateAndEmit();

        res.json({ success: true, data: systemDatabase });
    } catch (err) {
        console.error('manual-save error:', err);
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
        let tempTanks = [...systemDatabase.tanks];

        const sheetsToProcess = [];

        if (selectedTankNum && selectedTankNum !== 'all') {
            if (sheetNames.length > 0) {
                sheetsToProcess.push({
                    sheet: workbook.Sheets[sheetNames[0]],
                    targetTank: parseInt(selectedTankNum)
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

        const now = new Date();
        systemDatabase.lastUpdated = now.toLocaleDateString('my-MM') + " - " + now.toLocaleTimeString('my-MM');

        await updateAndEmit();

        fs.unlink(req.file.path, () => {});
        res.json({ success: true, data: systemDatabase });
    } catch (err) {
        console.error('excel-upload error:', err);
        if (req.file && req.file.path) {
            fs.unlink(req.file.path, () => {});
        }
        res.status(500).json({ success: false, message: err.message });
    }
});

io.on('connection', (socket) => {
    socket.emit('dataUpdate', systemDatabase);
});

mongoose.connect(MONGODB_URI)
    .then(async () => {
        console.log('MongoDB connected');
        await loadDatabaseFromMongo();

        const PORT = process.env.PORT || 3000;
        server.listen(PORT, () => {
            console.log(`Server running on port ${PORT}`);
        });
    })
    .catch(err => {
        console.error('MongoDB connection failed:', err);
    });
