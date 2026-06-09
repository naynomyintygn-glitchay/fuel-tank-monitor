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

// Database Initial State
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

// API - ဒေတာများ ဆွဲယူရန်
app.get('/api/data', (req, res) => {
    res.json({ station: stationData, tanks: tanks });
});

// API - ဆီဆိုင် Profile ပြင်ဆင်ရန်
app.post('/api/admin/update-station', (req, res) => {
    const { password, name, branchName, phoneNumber, logoUrl } = req.body;
    if (password !== stationData.adminPassword) return res.status(403).json({ success: false, message: "Password မှားယွင်းနေပါသည်။" });
   
    stationData.name = name;
    stationData.branchName = branchName;
    stationData.phoneNumber = phoneNumber;
    stationData.logoUrl = logoUrl;
   
    io.emit('dataUpdate', { station: stationData, tanks: tanks });
    res.json({ success: true, message: "ဆိုင်အချက်အလက် ပြင်ဆင်ပြီးပါပြီ။" });
});

// API - တိုင်ကီ ၆ ခုလုံးကို ပြိုင်တူ တစ်ပြိုင်နက် သိမ်းဆည်းပေးခြင်း Logic
app.post('/api/admin/update-all-tanks', (req, res) => {
    const { password, tanksData } = req.body;
   
    // Password စစ်ဆေးခြင်း
    if (password !== stationData.adminPassword) {
        return res.status(403).json({ success: false, message: "Password မှားယွင်းနေပါသည်။" });
    }
   
    if (!tanksData || !Array.isArray(tanksData)) {
        return res.status(400).json({ success: false, message: "မှန်ကန်သော ဒေတာပုံစံ မဟုတ်ပါ။" });
    }

    // ဝင်လာသော တိုင်ကီဒေတာ ၆ ခုလုံးကို Array ထဲမှာ အစားထိုး သိမ်းဆည်းခြင်း
    tanksData.forEach(incomingTank => {
        const index = tanks.findIndex(t => t.tankNumber == incomingTank.tankNumber);
        if (index !== -1) {
            if (incomingTank.fuelType !== undefined && incomingTank.fuelType !== "") {
                tanks[index].fuelType = String(incomingTank.fuelType).trim(); // ဆီအမျိုးအစားသစ်ကို အစားထိုးသိမ်းဆည်းခြင်း
            }
            tanks[index].maxCapacity = Number(incomingTank.maxCapacity);
            tanks[index].currentCM = incomingTank.currentCM;
            tanks[index].currentLiter = Number(incomingTank.currentLiter);
        }
    });
   
    // Frontend UI များဆီသို့ ဒေတာအသစ် ချက်ချင်း ပို့လွှတ်ခြင်း
    io.emit('dataUpdate', { station: stationData, tanks: tanks });
    return res.json({ success: true, message: "တိုင်ကီအားလုံးကို အောင်မြင်စွာ သိမ်းဆည်းပြီးပါပြီဗျာ။" });
});

// API - Excel Bulk Upload စနစ်
app.post('/api/admin/upload-excel', upload.single('excelFile'), (req, res) => {
    const { password } = req.body;
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

        tanks.forEach((tank, tIndex) => {
            let latestCM = "N/A";
            let latestLiter = 0;
            let possibleLiters = [];
            let possibleCMs = [];

            if(tank.tankNumber === 1) { possibleLiters = ['Liter', 'Liter_1']; possibleCMs = ['CM', 'CM_1']; }
            else if(tank.tankNumber === 2) { possibleLiters = ['Liter_2', 'Liter_3']; possibleCMs = ['CM_2', 'CM_3']; }
            else if(tank.tankNumber === 3) { possibleLiters = ['Liter_4', 'Liter_5']; possibleCMs = ['CM_4', 'CM_5']; }
            else if(tank.tankNumber === 4) { possibleLiters = ['Liter_6', 'Liter_7']; possibleCMs = ['CM_6', 'CM_7']; }
            else if(tank.tankNumber === 5) { possibleLiters = ['Liter_8', 'Liter_9']; possibleCMs = ['CM_8', 'CM_9']; }
            else if(tank.tankNumber === 6) { possibleLiters = ['Liter_10', 'Liter_11']; possibleCMs = ['CM_10', 'CM_11']; }

            rawData.forEach(row => {
                for (let i = 0; i < possibleLiters.length; i++) {
                    let litVal = row[possibleLiters[i]];
                    let cmVal = row[possibleCMs[i]];

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

            if (latestLiter > 0) {
                tanks[tIndex].currentCM = latestCM;
                tanks[tIndex].currentLiter = latestLiter;
            }
        });

        io.emit('dataUpdate', { station: stationData, tanks: tanks });
        res.json({ success: true, message: "Excel ဖိုင်ကို ဖတ်ရှုပြီး တိုင်ကီ (၆) ခုလုံးအား အောင်မြင်စွာ Update ပြုလုပ်ပြီးပါပြီဗျာ။" });

    } catch (error) {
        if(fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.status(500).json({ success: false, message: "Error: " + error.message });
    }
});

io.on('connection', (socket) => { console.log('Client Connected'); });

const PORT = process.env.PORT || 3000;
http.listen(PORT, () => { console.log(`Server running on port ${PORT}`); });
