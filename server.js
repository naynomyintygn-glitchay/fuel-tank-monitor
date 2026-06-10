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

// Database Fallback Store Memory
let systemDatabase = {
    lastUpdated: "မရှိသေးပါ",
    tanks: [
        { tankNumber: 1, fuelType: "HSD", currentCM: "220", currentLiter: "22500", maxCapacity: "30500" },
        { tankNumber: 2, fuelType: "Premium Diesel", currentCM: "220", currentLiter: "22500", maxCapacity: "30500" },
        { tankNumber: 3, fuelType: "92 RON", currentCM: "220", currentLiter: "22500", maxCapacity: "30500" },
        { tankNumber: 4, fuelType: "92 RON", currentCM: "220", currentLiter: "22500", maxCapacity: "30500" },
        { tankNumber: 5, fuelType: "95 RON", currentCM: "142", currentLiter: "11000", maxCapacity: "15000" },
        { tankNumber: 6, fuelType: "92 RON", currentCM: "142", currentLiter: "11000", maxCapacity: "15000" }
    ]
};

// API: Get Live Monitor Data
app.get('/api/data', (req, res) => {
    res.json(systemDatabase);
});

// API: Save Manual Inputs
app.post('/api/manual-save', (req, res) => {
    const { tanks } = req.body;
    const now = new Date();
    systemDatabase.lastUpdated = now.toLocaleDateString() + " " + now.toLocaleTimeString();
    systemDatabase.tanks = tanks;
   
    io.emit('dataUpdate', systemDatabase); // Broadcast Live Frontend Update
    res.json({ status: "success" });
});

// API: Processing Excel - Read Sheet per Tank and get LAST VALID ROW values
app.post('/api/excel-upload', upload.single('excelFile'), (req, res) => {
    try {
        const workbook = xlsx.readFile(req.file.path);
        const sheetNames = workbook.SheetNames; // Sheet list (Tank 1, Tank 2, ...)
       
        let updatedTanks = [];

        sheetNames.forEach((sheetName, index) => {
            if(index >= 6) return; // စနစ်တွင် တိုင်ကီ ၆ လုံးသာ ကန့်သတ်ဖတ်ရှုမည်
           
            const sheet = workbook.Sheets[sheetName];
            // Sheet တစ်ခုလုံးကို Row မျိုးစုံဖတ်နိုင်ရန် 2D Array ပြောင်းလဲခြင်း
            const matrixData = xlsx.utils.sheet_to_json(sheet, { header: 1 });
           
            let lastValidCM = "0";
            let lastValidLiter = "0";
            let detectedCapacity = 30500; // Fallback default
            let detectedProduct = "Unknown";

            // စာမျက်နှာအပေါ်ပိုင်း ခေါင်းစဉ်ဧရိယာမှ CAPACITY နှင့် PRODUCT ရှာဖွေခြင်း
            matrixData.forEach(row => {
                if(!row || row.length === 0) return;
                const lineStr = row.join(" ").toUpperCase();
               
                if(lineStr.includes("CAPACITY")) {
                    const match = lineStr.match(/\d+/);
                    if(match) detectedCapacity = parseInt(match[0]);
                }
                // ဆီအမျိုးအစားသတ်မှတ်ချက်ကို ရှာဖွေခြင်း
                if(lineStr.includes("92 RON")) detectedProduct = "92 RON";
                else if(lineStr.includes("95 RON")) detectedProduct = "95 RON";
                else if(lineStr.includes("PREMIUM DIESEL") || lineStr.includes("PDO")) detectedProduct = "Premium Diesel";
                else if(lineStr.includes("HSD") || lineStr.includes("DIESEL")) detectedProduct = "HSD";
            });

            // 🌟 [အဓိကစနစ်] အောက်ဆုံး Row များမှ ကိန်းဂဏန်းစစ်စစ် ပါဝင်သော စာကြောင်းကို ပြောင်းပြန်လှန်၍ ရှာဖွေဖတ်ရှုခြင်း
            for (let i = matrixData.length - 1; i >= 0; i--) {
                const currentRow = matrixData[i];
                if (!currentRow || currentRow.length < 3) continue;
               
                // Row ထဲက ကော်လံတွေကို စစ်ဆေးပြီး ဂဏန်းအမှန်ပါတာကို ယူခြင်း
                const col1 = parseFloat(String(currentRow[1]).replace(/,/g, ''));
                const col2 = parseFloat(String(currentRow[2]).replace(/,/g, ''));

                if (!isNaN(col1) && !isNaN(col2) && col1 > 0 && col2 > 0) {
                    lastValidCM = String(currentRow[1]);
                    lastValidLiter = String(currentRow[2]);
                    break; // နောက်ဆုံးစာကြောင်း တွေ့သည်နှင့် ရပ်တန့်မည်
                }
            }

            updatedTanks.push({
                tankNumber: index + 1,
                fuelType: detectedProduct !== "Unknown" ? detectedProduct : `Tank ${index + 1}`,
                currentCM: lastValidCM,
                currentLiter: lastValidLiter,
                maxCapacity: String(detectedCapacity)
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

server.listen(3000, () => console.log('Htoo Station Server running on port 3000'));
