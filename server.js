const express = require('express');
const app = express();
const http = require('http').createServer(app);
const io = require('socket.io')(http);
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

app.use(express.json());
app.use(express.static('public'));

const upload = multer({ dest: 'uploads/' });

// Database initial state
let stationData = {
    name: "Htoo Fuel Station",
    branchName: "ပြင်ဦးလွင်",
    phoneNumber: "09-123456789",
    logoUrl: "https://placehold.co/100x100/orange/white?text=HTOO",
    adminPassword: "123"
};

let tanks = [
    { tankNumber: 1, fuelType: "HSD", maxCapacity: 30500, currentCM: 120.5, currentLiter: 15200 },
    { tankNumber: 2, fuelType: "Premium Diesel", maxCapacity: 30500, currentCM: 175.2, currentLiter: 24100 },
    { tankNumber: 3, fuelType: "92 RON", maxCapacity: 30500, currentCM: 142.0, currentLiter: 18500 },
    { tankNumber: 4, fuelType: "92 RON", maxCapacity: 30500, currentCM: 70.4, currentLiter: 8900 },
    { tankNumber: 5, fuelType: "95 RON", maxCapacity: 15000, currentCM: 165.1, currentLiter: 12000 },
    { tankNumber: 6, fuelType: "95 RON", maxCapacity: 15000, currentCM: 65.3, currentLiter: 4500 }
];

app.get('/api/data', (req, res) => { res.json({ station: stationData, tanks: tanks }); });

app.post('/api/admin/update-station', (req, res) => {
    const { password, name, branchName, phoneNumber, logoUrl } = req.body;
    if (password !== stationData.adminPassword) return res.status(403).json({ success: false, message: "Password မှားယွင်းနေပါသည်။" });
    stationData.name = name; stationData.branchName = branchName; stationData.phoneNumber = phoneNumber; stationData.logoUrl = logoUrl;
    io.emit('dataUpdate', { station: stationData, tanks: tanks });
    res.json({ success: true, message: "ဆိုင်အချက်အလက် ပြင်ဆင်ပြီးပါပြီ။" });
});

app.post('/api/admin/update-tank', (req, res) => {
    const { password, tankNumber, maxCapacity, currentCM, currentLiter } = req.body;
    if (password !== stationData.adminPassword) return res.status(403).json({ success: false, message: "Password မှားယွင်းနေပါသည်။" });
    const index = tanks.findIndex(t => t.tankNumber == tankNumber);
    if (index !== -1) {
        tanks[index].maxCapacity = Number(maxCapacity);
        tanks[index].currentCM = currentCM;
        tanks[index].currentLiter = Number(currentLiter);
        io.emit('dataUpdate', { station: stationData, tanks: tanks });
        return res.json({ success: true, message: "လက်မန်ဖြင့် ပြင်ဆင်မှု အောင်မြင်ပါသည်။" });
    }
    res.status(404).json({ success: false, message: "Tank ရှာမတွေ့ပါ။" });
});

// Excel ဖတ်ရှုသည့် အဆင့်မြင့် စနစ် Logic
app.post('/api/admin/upload-excel', upload.single('excelFile'), (req, res) => {
    const { password, tankNumber } = req.body;
    if (password !== stationData.adminPassword) {
        if(req.file) fs.unlinkSync(req.file.path);
        return res.status(403).json({ success: false, message: "Password မှားယွင်းနေပါသည်။" });
    }
    if (!req.file) return res.status(400).json({ success: false, message: "Excel ဖိုင်တင်ရန် လိုအပ်ပါသည်။" });

    try {
        const workbook = xlsx.readFile(req.file.path);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawData = xlsx.utils.sheet_to_json(worksheet, { defval: "" });
        fs.unlinkSync(req.file.path);

        let latestCM = "N/A";
        let latestLiter = 0;

        rawData.forEach(row => {
            const possibleLiters = [row['Liter'], row['Liter_1'], row['Liter_2']];
            const possibleCMs = [row['CM'], row['CM_1'], row['CM_2']];

            for (let i = 0; i < possibleLiters.length; i++) {
                let litVal = possibleLiters[i];
                let cmVal = possibleCMs[i];

                if (litVal !== undefined && litVal !== null && String(litVal).trim() !== "") {
                    let cleanLitStr = String(litVal).replace(/,/g, '').trim();
                    let currentLit = parseInt(cleanLitStr);

                    if (!isNaN(currentLit) && currentLit > 0) {
                        if (currentLit >= latestLiter) {
                            latestLiter = currentLit;
                            if (cmVal !== undefined && cmVal !== null && String(cmVal).trim() !== "") {
                                let cmStr = String(cmVal).trim();
                                if (cmStr.includes('..')) cmStr = cmStr.replace('..', '.');
                                let parsedCM = parseFloat(cmStr);
                                latestCM = isNaN(parsedCM) ? "N/A" : parsedCM;
                            } else {
                                latestCM = "N/A";
                            }
                        }
                    }
                }
            }
        });

        const index = tanks.findIndex(t => t.tankNumber == tankNumber);
        if (index !== -1) {
            tanks[index].currentCM = latestCM;
            tanks[index].currentLiter = latestLiter;
            io.emit('dataUpdate', { station: stationData, tanks: tanks });
            return res.json({
                success: true,
                message: `Excel ဖတ်ရှုပြီး Tank ${tankNumber} ကို အောင်မြင်စွာ Update ပြုလုပ်ပြီးပါပြီဗျာ။`,
                data: { cm: latestCM, liter: latestLiter }
            });
        }
        res.status(404).json({ success: false, message: "တိုင်ကီနံပါတ် ရှာမတွေ့ပါ။" });
    } catch (error) {
        if(fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.status(500).json({ success: false, message: "Error: " + error.message });
    }
});

io.on('connection', (socket) => { console.log('Client Connected'); });

// Cloud Server ပေါ်တွင် Port ပြောင်းလဲမှုများကို အလိုအလျောက် သိရှိစေမည့် လိုင်း (ပြင်ဆင်ပြီးသား)
const PORT = process.env.PORT || 3000;
http.listen(PORT, () => { console.log(`Server running on port ${PORT}`); });