const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const http = require('http');
const socketIo = require('socket.io');
const path = require('path');
const fs = require('fs'); // --- NEW: File System module ---

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

// --- NEW: Data storage file ---
const DATA_FILE = path.join(__dirname, 'data.json'); 

// Default database structure --- ADDED defaultTankData for easier management ---
const defaultStationData = {
    name: "Htoo Fuel Station",
    branchName: "ပြင်ဦးလွင်",
    phoneNumber: "09-123456789",
    logoUrl: "[placehold.co](https://placehold.co/100x100/orange/white?text=HTOO)"
};

const defaultTankData = [
    { tankNumber: 1, fuelType: "HSD", currentCM: "220", currentLiter: "22500", maxCapacity: "30500" },
    { tankNumber: 2, fuelType: "Premium Diesel", currentCM: "220", currentLiter: "22500", maxCapacity: "30500" },
    { tankNumber: 3, fuelType: "92 RON", currentCM: "220", currentLiter: "22500", maxCapacity: "30500" },
    { tankNumber: 4, fuelType: "92 RON", currentCM: "220", currentLiter: "22500", maxCapacity: "30500" },
    { tankNumber: 5, fuelType: "95 RON", currentCM: "142", currentLiter: "11000", maxCapacity: "15000" },
    { tankNumber: 6, fuelType: "92 RON", currentCM: "142", currentLiter: "11000", maxCapacity: "15000" }
];

let systemDatabase;

// --- NEW: Function to load data from file ---
function loadDatabase() {
    try {
        if (fs.existsSync(DATA_FILE)) {
            const data = fs.readFileSync(DATA_FILE, 'utf8');
            systemDatabase = JSON.parse(data);
            // Ensure all necessary fields exist, if not, use defaults
            systemDatabase.station = { ...defaultStationData, ...systemDatabase.station };
            systemDatabase.tanks = systemDatabase.tanks || defaultTankData;
            console.log('Database loaded from file.');
        } else {
            console.log('Data file not found, initializing with default data.');
            systemDatabase = {
                lastUpdated: "မရှိသေးပါ",
                station: defaultStationData,
                tanks: defaultTankData
            };
            saveDatabase(); // Save default data to file
        }
    } catch (err) {
        console.error('Error loading database:', err);
        // Fallback to default if there's an error reading/parsing file
        systemDatabase = {
            lastUpdated: "မရှိသေးပါ",
            station: defaultStationData,
            tanks: defaultTankData
        };
        console.warn('Using default database due to load error.');
    }
}

// --- NEW: Function to save data to file ---
function saveDatabase() {
    try {
        fs.writeFileSync(DATA_FILE, JSON.stringify(systemDatabase, null, 2), 'utf8');
        console.log('Database saved to file.');
    } catch (err) {
        console.error('Error saving database:', err);
    }
}

// --- NEW: Load database on server start ---
loadDatabase();

// Socket.IO connection handling
io.on('connection', (socket) => {
    console.log('Client connected:', socket.id);
    socket.emit('dataUpdate', systemDatabase);
    
    socket.on('disconnect', () => {
        console.log('Client disconnected:', socket.id);
    });
});

// API: Verify Admin Password
app.post('/api/verify-password', (req, res) => {
    const { password } = req.body;
    if (password === ADMIN_PASSWORD) {
        res.json({ success: true });
    } else {
        res.json({ success: false, message: "Password မှားနေပါသည်" });
    }
});

// API: Get Live Monitor State
app.get('/api/data', (req, res) => {
    res.json(systemDatabase);
});

// API: Save Manual + Station Configuration Data --- CHANGED: calls saveDatabase() ---
app.post('/api/manual-save', (req, res) => {
    const { station, tanks, password } = req.body;
    
    if (password !== ADMIN_PASSWORD) {
        return res.status(401).json({ success: false, message: "Password မှားနေပါသည်" });
    }
    
    const now = new Date();
    // Using Burmese locale for date/time formatting
    systemDatabase.lastUpdated = now.toLocaleDateString('my-MM') + " - " + now.toLocaleTimeString('my-MM');
    
    // Merge station data instead of overwriting completely
    if (station) systemDatabase.station = { ...systemDatabase.station, ...station };
    if (tanks) systemDatabase.tanks = tanks;
    
    saveDatabase(); // Save to file
    io.emit('dataUpdate', systemDatabase); // Emit live shift
    res.json({ success: true, data: systemDatabase });
});

// API: Update Single Tank (This API wasn't directly used in latest admin.html but good to keep)
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
    systemDatabase.lastUpdated = now.toLocaleDateString('my-MM') + " - " + now.toLocaleTimeString('my-MM');
    
    // Update only provided fields
    if (currentCM !== undefined) systemDatabase.tanks[tankIndex].currentCM = currentCM;
    if (currentLiter !== undefined) systemDatabase.tanks[tankIndex].currentLiter = currentLiter;
    if (maxCapacity !== undefined) systemDatabase.tanks[tankIndex].maxCapacity = maxCapacity;

    saveDatabase(); // Save to file
    io.emit('dataUpdate', systemDatabase);
    res.json({ success: true, data: systemDatabase });
});

// ========== EXCEL AUTO-CORRECT PARSER ==========
function autoCorrectValue(val) {
    if (val === null || val === undefined || val === '') return null;
    
    let str = String(val).trim();
    
    // Handle Excel date serial numbers (e.g., "1900-01-01T05:02:24.000Z")
    if (str.includes('1900-') || (str.includes('T') && str.includes('Z'))) {
        // This regex tries to find time part, which is sometimes CM or Liter when data is malformed
        const timeMatch = str.match(/(\d{1,2}):(\d{2}):(\d{2})/);
        if (timeMatch) {
            // Attempt to convert time-like string to a number if it looks like CM/Liter data
            const hr = parseInt(timeMatch[1]);
            const min = parseInt(timeMatch[2]);
            const sec = parseInt(timeMatch[3]);
            // Heuristic: if time looks like a reasonable CM value (e.g., small hour), use it
            if (hr < 250 && min < 60 && sec < 60) { // Assuming CM won't exceed 250 (roughly)
                return parseFloat(`${hr}.${min}`); // e.g., 5:02:24 -> 5.02
            }
        }
        return null; // Otherwise, ignore date/time values
    }
    
    // Remove common typos (double dots like "10..5")
    str = str.replace(/\.{2,}/g, '.');
    
    // Remove commas (thousands separator), if present
    str = str.replace(/,/g, '');
    
    // Remove any non-numeric characters except dot and minus (but not dots if there are many)
    str = str.replace(/[^0-9.-]/g, ''); 
    
    // Ensure only one dot for decimals
    const parts = str.split('.');
    if (parts.length > 2) {
        str = parts[0] + '.' + parts.slice(1).join('');
    }
    
    if (str === '' || str === '.' || str === '-') return null; // Edge cases like just a dot or minus

    const num = parseFloat(str);
    return isNaN(num) ? null : num;
}

function parseSheetData(matrixData) {
    let lastValidCM = "0";
    let lastValidLiter = "0";
    let detectedCapacity = "30500";
    let detectedProduct = "Unknown";
    let tankNumber = null;
    
    // Scan headers for metadata (up to first 15 rows)
    for (let i = 0; i < Math.min(matrixData.length, 15); i++) {
        const row = matrixData[i];
        if (!row || row.length === 0) continue;
        
        // Convert all elements in the row to string and join for full line scan
        const lineStr = row.map(cell => String(cell || '').toUpperCase()).join(" ");
        
        // Detect Capacity
        const capacityMatch = lineStr.match(/\(?\s*(\d{1,3}(?:,?\d{3})*)\s*\)?\s*LITER/i);
        if (capacityMatch && capacityMatch[1]) {
            detectedCapacity = capacityMatch[1].replace(/,/g, '');
        }
        
        // Detect Tank Number
        const tankMatch = lineStr.match(/TANK\s*(?:NO\.?|NUMBER)?\s*\(?(\d+)\)?/i);
        if (tankMatch && tankMatch[1]) {
            tankNumber = parseInt(tankMatch[1]);
        }
        
        // Detect Fuel Type
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
        const currentRow = matrixData[i];
        if (!currentRow || currentRow.length < 3) continue;
        
        // Define common column indices for CM and Liter in your Excel format
        // This is flexible to account for variations if present
        const possibleCmCols = [1, 4, 7]; // column B, E, H
        const possibleLiterCols = [2, 5, 8]; // column C, F, I

        for (let j = 0; j < possibleCmCols.length; j++) {
            const cmCol = possibleCmCols[j];
            const literCol = possibleLiterCols[j];

            // Ensure columns exist in the current row
            if (currentRow.length > cmCol && currentRow.length > literCol) {
                const cmVal = autoCorrectValue(currentRow[cmCol]);
                const literVal = autoCorrectValue(currentRow[literCol]);
                
                // Heuristic validation: CM in typical range (e.g., 0-300), Liter also in a reasonable range
                // Adjust ranges if your tanks have different max heights/volumes
                const isCmReasonable = cmVal !== null && cmVal > 0 && cmVal < 300; 
                const isLiterReasonable = literVal !== null && literVal >= 0 && literVal <= 50000; // Assuming max liter 50k
                
                if (isCmReasonable && isLiterReasonable) {
                    lastValidCM = String(cmVal.toFixed(1)); // Format to 1 decimal place
                    lastValidLiter = String(literVal.toFixed(0)); // Format to whole number
                    break; // Found valid data, stop searching this row
                }
            }
        }
        
        if (lastValidCM !== "0") break; // Found valid data in this row, stop checking previous rows
    }
    
    return {
        cm: lastValidCM,
        liter: lastValidLiter,
        capacity: detectedCapacity,
        fuelType: detectedProduct,
        tankNumber: tankNumber // This might be null if not detected
    };
}

// API: Excel Upload with Auto-Correct --- CHANGED: calls saveDatabase() ---
app.post('/api/excel-upload', upload.single('excelFile'), (req, res) => {
    const { password, tankNumber: selectedTankNum } = req.body;
    
    if (password !== ADMIN_PASSWORD) {
        return res.status(401).json({ success: false, message: "Password မှားနေပါသည်" });
    }
    
    if (!req.file) {
        return res.status(400).json({ success: false, message: "No Excel file uploaded." });
    }

    try {
        const workbook = xlsx.readFile(req.file.path);
        const sheetNames = workbook.SheetNames;
        let updateLog = [];
        
        // --- Loop through selected sheets based on admin input ---
        const sheetsToProcess = [];
        if (selectedTankNum && selectedTankNum !== 'all') {
            // If admin selected a specific tank, process only the first sheet
            // and try to map it to the selected tank number.
            if (sheetNames.length > 0) {
                sheetsToProcess.push({ sheet: workbook.Sheets[sheetNames[0]], index: 0, targetTank: parseInt(selectedTankNum) });
            }
        } else {
            // If 'all' is selected, process up to 6 sheets, mapping to Tank 1-6
            for (let i = 0; i < Math.min(sheetNames.length, 6); i++) {
                sheetsToProcess.push({ sheet: workbook.Sheets[sheetNames[i]], index: i, targetTank: i + 1 });
            }
        }

        if (sheetsToProcess.length === 0) {
            // Cleanup and respond if no sheets were processed
            fs.unlink(req.file.path, () => {});
            return res.json({ success: true, message: "No data found or processed from Excel file.", log: [] });
        }

        sheetsToProcess.forEach(({ sheet, index, targetTank }) => {
            const matrixData = xlsx.utils.sheet_to_json(sheet, { header: 1, raw: false });
            const parsed = parseSheetData(matrixData);
            
            // Determine actual tank to update: priority to parsed from sheet, then targetTank from admin, then default index
            const finalTankNum = parsed.tankNumber || targetTank;
            const tankIdx = systemDatabase.tanks.findIndex(t => t.tankNumber == finalTankNum);
            
            if (tankIdx !== -1) {
                // Only update if parsed values are valid and not default "0" or "Unknown"
                if (parsed.cm !== "0") systemDatabase.tanks[tankIdx].currentCM = parsed.cm;
                if (parsed.liter !== "0") systemDatabase.tanks[tankIdx].currentLiter = parsed.liter;
                if (parsed.capacity !== "30500") systemDatabase.tanks[tankIdx].maxCapacity = parsed.capacity;
                // Only update fuelType if detected from sheet and not "Unknown"
                if (parsed.fuelType !== "Unknown" && parsed.fuelType !== systemDatabase.tanks[tankIdx].fuelType) {
                    systemDatabase.tanks[tankIdx].fuelType = parsed.fuelType;
                    updateLog.push(`Tank ${finalTankNum}: Fuel Type changed to ${parsed.fuelType}`);
                }
                updateLog.push(`Tank ${finalTankNum}: CM=${parsed.cm}, Liter=${parsed.liter}`);
            } else {
                updateLog.push(`Error: Could not find Tank ${finalTankNum} in system configuration to update.`);
            }
        });
        
        const now = new Date();
        systemDatabase.lastUpdated = now.toLocaleDateString('my-MM') + " - " + now.toLocaleTimeString('my-MM');
        
        saveDatabase(); // Save to file
        io.emit('dataUpdate', systemDatabase);
        
        // Clean up uploaded file
        fs.unlink(req.file.path, () => {}); // Using callback, so it doesn't block
        
        res.json({ 
            success: true, 
            data: systemDatabase,
            log: updateLog.length > 0 ? updateLog : ["Excel data processed successfully."] // Provide a log or default success msg
        });
    } catch (err) {
        console.error('Excel parse error:', err);
        // Clean up even on error
        if (req.file && req.file.path) fs.unlink(req.file.path, () => {});
        res.status(500).json({ success: false, message: 'Excel ဖိုင်ဖတ်ရာတွင် အမှားရှိသည်: ' + (err.message || "Unknown error") });
    }
});

const PORT = process.env.PORT || 3000;
server.listen(PORT, () => console.log(`Htoo ATG Monitor running on port ${PORT}`));

