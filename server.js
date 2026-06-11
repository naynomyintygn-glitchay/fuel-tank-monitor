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

// IMPORTANT: Ensure MONGODB_URI is set as an Environment Variable in Render Dashboard!
const MONGODB_URI = process.env.MONGODB_URI;

// Default data (used only if database is completely empty upon first connection)
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

// Mongoose Schemas (should be correct from previous versions)
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
            
            // Check for any missing tanks from defaultTanksData and add them if necessary
            const existingTankNumbers = new Set(tanksDocs.map(t => t.tankNumber));
            const tanksToAdd = defaultTanksData.filter(dTank => !existingTankNumbers.has(dTank.tankNumber));
            if (tanksToAdd.length > 0) {
              await Tank.insertMany(tanksToAdd);
            }
            // Re-fetch all tanks to ensure systemDatabase holds all tanks including newly added defaults
            systemDatabase.tanks = (await Tank.find({}).sort({ tankNumber: 1 })).map(doc => doc.toObject());
        }

        systemDatabase.lastUpdated = systemDatabase.station.lastUpdated || "မရှိသေးပါ";
    } catch (err) {
        console.error('Load MongoDB error in loadDatabaseFromMongo:', err);
        // It's crucial to handle this. Maybe default to in-memory data if DB load fails completely.
        // For now, it will use the initialized systemDatabase.
    }
}

async function updateAndEmit() {
    try {
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
                { upsert: true } // Create document if it doesn't exist
            );
        }

        // Reload data from Mongo to ensure systemDatabase is always fresh from DB before emitting
        await loadDatabaseFromMongo();
        io.emit('dataUpdate', systemDatabase);
    } catch (err) {
        console.error('Update and emit error:', err);
        throw err; // Re-throw to propagate potential errors to API handlers
    }
}

function autoCorrectValue(val) {
    if (val === null || val === undefined || val === '') return null;

    let str = String(val).trim();

    // Specific logic for Excel date serials or time-like values
    if (str.includes('1900-') || (str.includes('T') && str.includes('Z'))) {
         return null; // Ignore actual date/time strings in CM/Liter columns
    }
    
    // Attempt to handle time-like values that look like floats (e.g., 5.02)
    const timeAsFloatMatch = str.match(/^(\d{1,2})\.(\d{1,2})$/); // e.g., 5.02, 10.30
    if (timeAsFloatMatch) {
        const hour = parseInt(timeAsFloatMatch[1]);
        const minute = parseInt(timeAsFloatMatch[2]);
        if (hour < 250 && minute < 60) { // Reasonable CM/Liter ranges for parts
            return parseFloat(str); // Treat as a valid float
        }
    }


    str = str.replace(/\.{2,}/g, '.'); // Replace multiple dots with a single dot
    str = str.replace(/,/g, ''); // Remove thousands separators (commas)
    str = str.replace(/[^0-9.-]/g, ''); // Remove any non-numeric characters except dot and minus

    const parts = str.split('.'); // Split by dot
    if (parts.length > 2) { // If there are more than one dot (e.g., 10.0.5)
        str = parts[0] + '.' + parts.slice(1).join(''); // Keep first part, join rest after first dot
    }

    if (str === '' || str === '.' || str === '-') return null; // Handle cases where cleanup leaves only symbols

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
                ...systemDatabase.station, // Keep existing station properties, only update provided ones
                ...station
            };
        }

        if (tanks && Array.isArray(tanks)) {
            // Ensure tank data from frontend is clean and valid before assigning
            const cleanedTanks = tanks.map(tank => ({
                tankNumber: Number(tank.tankNumber),
                fuelType: String(tank.fuelType),
                currentCM: String(tank.currentCM),
                currentLiter: String(tank.currentLiter),
                maxCapacity: String(tank.maxCapacity)
            }));
            systemDatabase.tanks = cleanedTanks;
        }

        await updateAndEmit();

        res.json({ success: true, data: systemDatabase });
    } catch (err) {
        console.error('manual-save error:', err);
        res.status(500).json({ success: false, message: err.message || "Unknown error during manual save." });
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
        let changesMade = false;
        let updateLog = [];


        const sheetsToProcess = [];

        if (selectedTankNum && selectedTankNum !== 'all') {
            if (sheetNames.length > 0) {
                sheetsToProcess.push({
                    sheet: workbook.Sheets[sheetNames[0]],
                    targetTank: parseInt(selectedTankNum)
                });
            }
        } else {
            // Process up to 6 sheets for automatic updates
            for (let i = 0; i < Math.min(sheetNames.length, 6); i++) {
                sheetsToProcess.push({
                    sheet: workbook.Sheets[sheetNames[i]],
                    targetTank: i + 1 // Default target tank if not detected in sheet
                });
            }
        }
        
        if (sheetsToProcess.length === 0) {
            fs.unlink(req.file.path, () => {}); // Clean up temp file
            return res.json({ success: true, message: "No data found or processed from Excel file.", log: [] });
        }

        for (const { sheet, targetTank } of sheetsToProcess) {
            const matrixData = xlsx.utils.sheet_to_json(sheet, { header: 1, raw: false });
            const parsed = parseSheetData(matrixData);

            const finalTankNum = parsed.tankNumber || targetTank;
            const tankIndex = tempTanks.findIndex(t => Number(t.tankNumber) === Number(finalTankNum));

            if (tankIndex !== -1) {
                if (parsed.cm !== "0" && tempTanks[tankIndex].currentCM !== parsed.cm) {
                    tempTanks[tankIndex].currentCM = parsed.cm; changesMade = true;
                    updateLog.push(`Tank ${finalTankNum}: CM updated to ${parsed.cm}`);
                }
                if (parsed.liter !== "0" && tempTanks[tankIndex].currentLiter !== parsed.liter) {
                    tempTanks[tankIndex].currentLiter = parsed.liter; changesMade = true;
                    updateLog.push(`Tank ${finalTankNum}: Liter updated to ${parsed.liter}`);
                }
                if (parsed.capacity !== "30500" && tempTanks[tankIndex].maxCapacity !== parsed.capacity) {
                    tempTanks[tankIndex].maxCapacity = parsed.capacity; changesMade = true;
                    updateLog.push(`Tank ${finalTankNum}: Max Capacity updated to ${parsed.capacity}`);
                }
                if (parsed.fuelType !== "Unknown" && tempTanks[tankIndex].fuelType !== parsed.fuelType) {
                    tempTanks[tankIndex].fuelType = parsed.fuelType; changesMade = true;
                    updateLog.push(`Tank ${finalTankNum}: Fuel Type updated to ${parsed.fuelType}`);
                }
            } else {
                updateLog.push(`Error: Tank ${finalTankNum} not found in system to update.`);
            }
        }

        if (changesMade) {
            systemDatabase.tanks = tempTanks; // Apply all changes to the main systemDatabase
            const now = new Date();
            systemDatabase.lastUpdated = now.toLocaleDateString('my-MM') + " - " + now.toLocaleTimeString('my-MM');
            await updateAndEmit(); // Persist changes to DB and emit
            res.json({ success: true, data: systemDatabase, log: updateLog });
        } else {
            res.json({ success: true, data: systemDatabase, log: ["No significant changes found in Excel data."] });
        }

        fs.unlink(req.file.path, () => {}); // Clean up temp file
    } catch (err) {
        console.error('excel-upload error:', err);
        if (req.file && req.file.path) {
            fs.unlink(req.file.path, () => {});
        }
        res.status(500).json({ success: false, message: err.message || "Unknown error during Excel upload." });
    }
});


// Mongoose Connect and Server Start (moved here to allow async/await for initial load)
mongoose.connect(MONGODB_URI)
    .then(async () => {
        console.log('MongoDB Connected successfully');
        await loadDatabaseFromMongo(); // Load initial data from DB

        const PORT = process.env.PORT || 3000;
        server.listen(PORT, () => console.log(`Htoo ATG Monitor running on port ${PORT}`));
    })
    .catch(err => {
        console.error('MongoDB connection error:', err);
        // If DB connection fails, ensure server still starts with default in-memory data
        // Or, exit if DB connection is truly critical. For now, we exit.
        process.exit(1); 
    });


