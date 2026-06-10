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

// Data file path for persistence
const DATA_FILE = path.join(__dirname, 'data.json');

// Default database structure
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

// Save database to file
function saveDatabase() {
    try {
        fs.writeFileSync(DATA_FILE, JSON.stringify(systemDatabase, null, 2));
    } catch (err) {
        console.error('Database save error:', err);
    }
}

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

// API: Save Station Settings
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

// API: Save Manual Tank Data
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

// API: Process uploaded Excel
app.post('/api/excel-upload', upload.single('excelFile'), (req, res) => {
    const { password, tankNumber } = req.body;
    
    if (password !== ADMIN_PASSWORD) {
        return res.status(401).json({ success: false, message: "Password မှားနေပါသည်" });
    }
    
    try {
        const workbook = xlsx.readFile(req.file.path);
        const sheetNames = workbook.SheetNames;
        
        // If specific tank is selected, only update that tank
        if (tankNumber && tankNumber !== 'all') {
            const sheet = workbook.Sheets[sheetNames[0]];
            const matrixData = xlsx.utils.sheet_to_json(sheet, { header: 1 });
            
            const { cm, liter, capacity } = parseSheetData(matrixData);
            
            const tankIdx = systemDatabase.tanks.findIndex(t => t.tankNumber == tankNumber);
            if (tankIdx !== -1) {
                if (cm !== "0") systemDatabase.tanks[tankIdx].currentCM = cm;
                if (liter !== "0") systemDatabase.tanks[tankIdx].currentLiter = liter;
                if (capacity !== "30500") systemDatabase.tanks[tankIdx].maxCapacity = capacity;
            }
        } else {
            // Update all tanks from multiple sheets
            sheetNames.forEach((sheetName, index) => {
                if (index >= 6) return;
                
                const sheet = workbook.Sheets[sheetName];
                const matrixData = xlsx.utils.sheet_to_json(sheet, { header: 1 });
                
                const { cm, liter, capacity, fuelType } = parseSheetData(matrixData);
                const previousTank = systemDatabase.tanks.find(t => t.tankNumber == (index + 1));
                
                if (previousTank) {
                    previousTank.currentCM = cm;
                    previousTank.currentLiter = liter;
                    if (capacity !== "30500") previousTank.maxCapacity = capacity;
                    if (fuelType !== "Unknown") previousTank.fuelType = fuelType;
                }
            });
        }
        
        const now = new Date();
        systemDatabase.lastUpdated = now.toLocaleDateString('my-MM') + " " + now.toLocaleTimeString('my-MM');
        
        saveDatabase();
        io.emit('dataUpdate', systemDatabase);
        
        // Clean up uploaded file
        fs.unlink(req.file.path, () => {});
        
        res.json({ success: true, data: systemDatabase });
    } catch (err) {
        console.error(err);
        res.status(500).json({ success: false, message: err.message });
    }
});

// Helper function to parse sheet data
function parseSheetData(matrixData) {
    let lastValidCM = "0";
    let lastValidLiter = "0";
    let detectedCapacity = "30500";
    let detectedProduct = "Unknown";
    
    // Header parse scanning
    matrixData.forEach(row => {
        if (!row || row.length === 0) return;
        const lineStr = row.join(" ").toUpperCase();
        
        if (lineStr.includes("CAPACITY")) {
            const match = lineStr.match(/(\d{1,3}(,\d{3})*|\d+)/);
            if (match) detectedCapacity = match[0].replace(/,/g, '');
        }
        if (lineStr.includes("92 RON")) detectedProduct = "92 RON";
        else if (lineStr.includes("95 RON")) detectedProduct = "95 RON";
        else if (lineStr.includes("PREMIUM DIESEL") || lineStr.includes("PDO")) detectedProduct = "Premium Diesel";
        else if (lineStr.includes("HSD") || lineStr.includes("DIESEL")) detectedProduct = "HSD";
    });
    
    // Find last valid data row
    for (let i = matrixData.length - 1; i >= 0; i--) {
        const currentRow = matrixData[i];
        if (!currentRow || currentRow.length < 3) continue;
        
        const col1 = parseFloat(String(currentRow[1]).replace(/,/g, ''));
        const col2 = parseFloat(String(currentRow[2]).replace(/,/g, ''));
        
        if (!isNaN(col1) && !isNaN(col2) && col1 > 0 && col2 > 0) {
            lastValidCM = String(currentRow[1]).replace(/,/g, '');
            lastValidLiter = String(currentRow[2]).replace(/,/g, '');
            break;
        }
    }
    
    return {
        cm: lastValidCM,
        liter: lastValidLiter,
        capacity: detectedCapacity,
        fuelType: detectedProduct
    };
}

const PORT = process.env.PORT || 3000;
server.listen(PORT, () => console.log(`Htoo ATG Monitor running on port ${PORT}`));
