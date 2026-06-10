const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const http = require('http');
const socketIo = require('socket.io');
const path = require('path');

const app = express();
const server = http.createServer(app);
const io = socketIo(server);
const upload = multer({ dest: 'uploads/' });

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Persistent memory object store
let systemDatabase = {
    lastUpdated: "မရှိသေးပါ",
    station: {
        name: "Htoo Fuel Station",
        branchName: "ပြင်ဦးလွင်ဘဏ်ခွဲ",
        phoneNumber: "09-123456789",
        logoUrl: "https://placehold.co/100x100/orange/white?text=HTOO"
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

// API: Get Live Monitor State
app.get('/api/data', (req, res) => {
    res.json(systemDatabase);
});

// API: Save Manual + Station Configuration Data
app.post('/api/manual-save', (req, res) => {
    const { station, tanks } = req.body;
    const now = new Date();
   
    systemDatabase.lastUpdated = now.toLocaleDateString() + " " + now.toLocaleTimeString();
    if(station) systemDatabase.station = station;
    if(tanks) systemDatabase.tanks = tanks;
   
    io.emit('dataUpdate', systemDatabase); // Emit live shift
    res.json({ status: "success", data: systemDatabase });
});

// API: Process uploaded excel (Sheets parse last matching numeric row)
app.post('/api/excel-upload', upload.single('excelFile'), (req, res) => {
    try {
        const workbook = xlsx.readFile(req.file.path);
        const sheetNames = workbook.SheetNames;
       
        let updatedTanks = [];

        sheetNames.forEach((sheetName, index) => {
            if(index >= 6) return; // 6 Tanks Bound limit
           
            const sheet = workbook.Sheets[sheetName];
            const matrixData = xlsx.utils.sheet_to_json(sheet, { header: 1 });
           
            let lastValidCM = "0";
            let lastValidLiter = "0";
            let detectedCapacity = 30500;
            let detectedProduct = "Unknown";

            // Header parse scanning loop
            matrixData.forEach(row => {
                if(!row || row.length === 0) return;
                const lineStr = row.join(" ").toUpperCase();
               
                if(lineStr.includes("CAPACITY")) {
                    const match = lineStr.match(/\d+/);
                    if(match) detectedCapacity = parseInt(match[0]);
                }
                if(lineStr.includes("92 RON")) detectedProduct = "92 RON";
                else if(lineStr.includes("95 RON")) detectedProduct = "95 RON";
                else if(lineStr.includes("PREMIUM DIESEL") || lineStr.includes("PDO")) detectedProduct = "Premium Diesel";
                else if(lineStr.includes("HSD") || lineStr.includes("DIESEL")) detectedProduct = "HSD";
            });

            // Parse back-loop algorithm to catch the last string pattern containing numerical values
            for (let i = matrixData.length - 1; i >= 0; i--) {
                const currentRow = matrixData[i];
                if (!currentRow || currentRow.length < 3) continue;
               
                const col1 = parseFloat(String(currentRow[1]).replace(/,/g, ''));
                const col2 = parseFloat(String(currentRow[2]).replace(/,/g, ''));

                if (!isNaN(col1) && !isNaN(col2) && col1 > 0 && col2 > 0) {
                    lastValidCM = String(currentRow[1]);
                    lastValidLiter = String(currentRow[2]);
                    break;
                }
            }

            // Keep the previous static config maximum or type details if missing from text headers
            const previousTank = systemDatabase.tanks.find(t => t.tankNumber == (index + 1));

            updatedTanks.push({
                tankNumber: index + 1,
                fuelType: detectedProduct !== "Unknown" ? detectedProduct : (previousTank ? previousTank.fuelType : `Tank ${index + 1}`),
                currentCM: lastValidCM,
                currentLiter: lastValidLiter,
                maxCapacity: detectedCapacity !== 30500 ? String(detectedCapacity) : (previousTank ? previousTank.maxCapacity : "30500")
            });
        });

        if(updatedTanks.length > 0) {
            const now = new Date();
            systemDatabase.lastUpdated = now.toLocaleDateString() + " " + now.toLocaleTimeString();
            systemDatabase.tanks = updatedTanks;
            io.emit('dataUpdate', systemDatabase);
        }

        res.json({ status: "success", data: systemDatabase });
    } catch (err) {
        console.error(err);
        res.status(500).json({ status: "error", message: err.message });
    }
});

server.listen(3000, () => console.log('Htoo ATG live system loaded on port 3000'));

