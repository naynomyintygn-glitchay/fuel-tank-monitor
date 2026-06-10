const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const http = require('http');
const socketIo = require('socket.io');
const path = require('path');
const fs = require('fs');

const app = express();
const server = http.createServer(app);
const io = socketIo(server, {
    cors: { origin: "*" }
});
const upload = multer({ dest: 'uploads/' });

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Admin Password
const ADMIN_PASSWORD = "htoo2024";

// Data file path
const DATA_FILE = path.join(__dirname, 'data.json');

// Default database
const defaultDatabase = {
    lastUpdated: "မရှိသေးပါ",
    station: {
        name: "Htoo Fuel Station",
        branchName: "ပြင်ဦးလွင်",
        phoneNumber: "09-123456789",
        logoUrl: "[placehold.co](https://placehold.co/100x100/orange/white?text=HTOO)"
    },
    tanks: [
        { tankNumber: 1, fuelType: "HSD", currentCM: "220", currentLiter: "22500", maxCapacity: "30500" },
        { tankNumber: 2, fuelType: "Premium Diesel", currentCM: "220", currentLiter: "22500", maxCapacity: "30500" },
        { tankNumber: 3, fuelType: "92 RON", currentCM: "220", currentLiter: "22500", maxCapacity: "30500" },
        { tankNumber: 4, fuelType: "92 RON", currentCM: "220", currentLiter: "22500", maxCapacity: "30500" },
        { tankNumber: 5, fuelType: "95 RON", currentCM: "142", currentLiter: "11000", maxCapacity: "15000" },
        { tankNumber: 6, fuelType: "92 RON", currentCM: "142", currentLiter: "11000", maxCapacity: "15000" }
    ]
};

// Load or initialize database
let systemDatabase;
try {
    if (fs.existsSync(DATA_FILE)) {
        systemDatabase = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
    } else {
        systemDatabase = defaultDatabase;
        fs.writeFileSync(DATA_FILE, JSON.stringify(systemDatabase, null, 2));
    }
} catch (err) {
    systemDatabase = defaultDatabase;
}

function saveDatabase() {
    try {
        fs.writeFileSync(DATA_FILE, JSON.stringify(systemDatabase, null, 2));
    } catch (err) {
        console.error('Database save error:', err);
    }
}

// Socket.IO
io.on('connection', (socket) => {
    console.log('Client connected:', socket.id);
    socket.emit('dataUpdate', systemDatabase);
    socket.on('disconnect', () => console.log('Client disconnected:', socket.id));
});

// API: Verify Password
app.post('/api/verify-password', (req, res) => {
    const { password } = req.body;
    res.json({ success: password === ADMIN_PASSWORD });
});

// API: Get Data
app.get('/api/data', (req, res) => {
    res.json(systemDatabase);
});

// API: Save Station
app.post('/api/save-station', (req, res) => {
    const { station, password } = req.body;
    if (password !== ADMIN_PASSWORD) {
        return res.status(401).json({ success: false, message: "Password မှားနေပါသည်" });
    }
    const now = new Date();
    systemDatabase.lastUpdated = now.toLocaleDateString('my-MM') + " " + now.toLocaleTimeString('my-MM');
    systemDatabase.station = { ...systemDatabase.station, ...station };
    saveDatabase();
    io.emit('dataUpdate', systemDatabase);
    res.json({ success: true, data: systemDatabase });
});

// API: Manual Save
app.post('/api/manual-save', (req, res) => {
    const { station, tanks, password } = req.body;
    if (password !== ADMIN_PASSWORD) {
        return res.status(401).json({ success: false, message: "Password မှားနေပါသည်" });
    }
    const now = new Date();
    systemDatabase.lastUpdated = now.toLocaleDateString('my-MM') + " " + now.toLocaleTimeString('my-MM');
    if (station) systemDatabase.station = station;
    if (tanks) systemDatabase.tanks = tanks;
    saveDatabase();
    io.emit('dataUpdate', systemDatabase);
    res.json({ success: true, data: systemDatabase });
});

// API: Update Single Tank
app.post('/api/update-tank', (req, res) => {
    const { tankNumber, currentCM, currentLiter, maxCapacity, password } = req.body;
    if (password !== ADMIN_PASSWORD) {
        return res.status(401).json({ success: false, message: "Password မှားနေပါသည်" });
    }
    const tankIndex = systemDatabase.tanks.findIndex(t => t.tankNumber == tankNumber);
    if (tankIndex === -1) {
        return res.status(404).json({ success: false, message: "Tank မတွေ့ပါ" });
    }
    const now = new Date();
    systemDatabase.lastUpdated = now.toLocaleDateString('my-MM') + " " + now.toLocaleTimeString('my-MM');
    if (currentCM !== undefined) systemDatabase.tanks[tankIndex].currentCM = currentCM;
    if (currentLiter !== undefined) systemDatabase.tanks[tankIndex].currentLiter = currentLiter;
    if (maxCapacity !== undefined) systemDatabase.tanks[tankIndex].maxCapacity = maxCapacity;
    saveDatabase();
    io.emit('dataUpdate', systemDatabase);
    res.json({ success: true, data: systemDatabase });
});

// ========== EXCEL AUTO-CORRECT PARSER ==========
function autoCorrectValue(val) {
    if (val === null || val === undefined || val === '') return null;
    
    let str = String(val).trim();
    
    // Handle Excel date serial numbers (e.g., "1900-01-01T05:02:24.000Z")
    if (str.includes('1900-') || str.includes('T') && str.includes('Z')) {
        // Try to extract numeric part
        const dateMatch = str.match(/(\d+):(\d+):(\d+)/);
        if (dateMatch) {
            // Convert time-like format to decimal (e.g., "05:02:24" -> might be 5.02 or similar)
            return null; // Skip malformed dates
        }
        return null;
    }
    
    // Remove common typos (double dots like "10..5")
    str = str.replace(/\.{2,}/g, '.');
    
    // Remove commas (thousands separator)
    str = str.replace(/,/g, '');
    
    // Remove any non-numeric characters except dot and minus
    str = str.replace(/[^\d.\-]/g, '');
    
    // Handle multiple dots (keep only first)
    const parts = str.split('.');
    if (parts.length > 2) {
        str = parts[0] + '.' + parts.slice(1).join('');
    }
    
    const num = parseFloat(str);
    return isNaN(num) ? null : num;
}

function parseSheetData(matrixData) {
    let lastValidCM = "0";
    let lastValidLiter = "0";
    let detectedCapacity = "30500";
    let detectedProduct = "Unknown";
    let tankNumber = null;
    
    // Scan headers for metadata
    for (let i = 0; i < Math.min(matrixData.length, 15); i++) {
        const row = matrixData[i];
        if (!row || row.length === 0) continue;
        
        const lineStr = row.join(" ").toUpperCase();
        
        // Detect capacity
        if (lineStr.includes("CAPACITY")) {
            const match = lineStr.match(/\(?\s*(\d{1,3}(?:,?\d{3})*)\s*\)?\s*LITER/i);
            if (match) {
                detectedCapacity = match[1].replace(/,/g, '');
            }
        }
        
        // Detect tank number
        const tankMatch = lineStr.match(/TANK\s*(?:NO\.?|NUMBER)?\s*\(?(\d+)\)?/i);
        if (tankMatch) {
            tankNumber = parseInt(tankMatch[1]);
        }
        
        // Detect fuel type
        if (lineStr.includes("95 RON") || lineStr.includes("95RON")) {
            detectedProduct = "95 RON";
        } else if (lineStr.includes("92 RON") || lineStr.includes("92RON")) {
            detectedProduct = "92 RON";
        } else if (lineStr.includes("PREMIUM DIESEL") || lineStr.includes("PDO")) {
            detectedProduct = "Premium Diesel";
        } else if (lineStr.includes("HSD") || (lineStr.includes("DIESEL") && !lineStr.includes("PREMIUM"))) {
            detectedProduct = "HSD";
        }
    }
    
    // Find last valid CM/Liter pair (scan from bottom)
    for (let i = matrixData.length - 1; i >= 0; i--) {
        const row = matrixData[i];
        if (!row || row.length < 3) continue;
        
        // Try multiple column combinations (data might be in different columns)
        const columnPairs = [
            [1, 2],   // Standard: CM in col 1, Liter in col 2
            [4, 5],   // Second set
            [7, 8],   // Third set
        ];
        
        for (const [cmCol, literCol] of columnPairs) {
            if (row.length <= literCol) continue;
            
            const cmVal = autoCorrectValue(row[cmCol]);
            const literVal = autoCorrectValue(row[literCol]);
            
            // Validate: CM should be < 300, Liter should be > 500 and < 35000
            if (cmVal !== null && literVal !== null && 
                cmVal > 0 && cmVal < 300 && 
                literVal >= 500 && literVal <= 35000) {
                lastValidCM = String(cmVal);
                lastValidLiter = String(literVal);
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
        tankNumber: tankNumber
    };
}

// API: Excel Upload with Auto-Correct
app.post('/api/excel-upload', upload.single('excelFile'), (req, res) => {
    const { password, tankNumber } = req.body;
    
    if (password !== ADMIN_PASSWORD) {
        return res.status(401).json({ success: false, message: "Password မှားနေပါသည်" });
    }
    
    try {
        const workbook = xlsx.readFile(req.file.path);
        const sheetNames = workbook.SheetNames;
        let updateLog = [];
        
        if (tankNumber && tankNumber !== 'all') {
            // Single tank update
            const sheet = workbook.Sheets[sheetNames[0]];
            const matrixData = xlsx.utils.sheet_to_json(sheet, { header: 1, raw: false });
            const parsed = parseSheetData(matrixData);
            
            const tankIdx = systemDatabase.tanks.findIndex(t => t.tankNumber == tankNumber);
            if (tankIdx !== -1) {
                if (parsed.cm !== "0") systemDatabase.tanks[tankIdx].currentCM = parsed.cm;
                if (parsed.liter !== "0") systemDatabase.tanks[tankIdx].currentLiter = parsed.liter;
                if (parsed.capacity !== "30500") systemDatabase.tanks[tankIdx].maxCapacity = parsed.capacity;
                if (parsed.fuelType !== "Unknown") systemDatabase.tanks[tankIdx].fuelType = parsed.fuelType;
                updateLog.push(`Tank ${tankNumber}: CM=${parsed.cm}, Liter=${parsed.liter}`);
            }
        } else {
            // Multi-sheet update
            sheetNames.forEach((sheetName, index) => {
                if (index >= 6) return;
                
                const sheet = workbook.Sheets[sheetName];
                const matrixData = xlsx.utils.sheet_to_json(sheet, { header: 1, raw: false });
                const parsed = parseSheetData(matrixData);
                
                // Use detected tank number or fall back to sheet index
                const targetTankNum = parsed.tankNumber || (index + 1);
                const tankIdx = systemDatabase.tanks.findIndex(t => t.tankNumber == targetTankNum);
                
                if (tankIdx !== -1) {
                    if (parsed.cm !== "0") systemDatabase.tanks[tankIdx].currentCM = parsed.cm;
                    if (parsed.liter !== "0") systemDatabase.tanks[tankIdx].currentLiter = parsed.liter;
                    if (parsed.capacity !== "30500") systemDatabase.tanks[tankIdx].maxCapacity = parsed.capacity;
                    if (parsed.fuelType !== "Unknown") systemDatabase.tanks[tankIdx].fuelType = parsed.fuelType;
                    updateLog.push(`Tank ${targetTankNum}: CM=${parsed.cm}, Liter=${parsed.liter}, Type=${parsed.fuelType}`);
                }
            });
        }
        
        const now = new Date();
        systemDatabase.lastUpdated = now.toLocaleDateString('my-MM') + " " + now.toLocaleTimeString('my-MM');
        
        saveDatabase();
        io.emit('dataUpdate', systemDatabase);
        
        // Cleanup
        fs.unlink(req.file.path, () => {});
        
        res.json({ 
            success: true, 
            data: systemDatabase,
            log: updateLog
        });
    } catch (err) {
        console.error('Excel parse error:', err);
        res.status(500).json({ success: false, message: 'Excel ဖိုင်ဖတ်ရာတွင် အမှားရှိသည်: ' + err.message });
    }
});

const PORT = process.env.PORT || 3000;
server.listen(PORT, () => console.log(`Htoo ATG Monitor running on port ${PORT}`));
