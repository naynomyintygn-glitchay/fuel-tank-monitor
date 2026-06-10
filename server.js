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

let stationData = {
    name: "Htoo Fuel Station",
    branchName: "ပြင်ဦးလွင်",
    phoneNumber: "09-123456789",
    logoUrl: "https://placehold.co/100x100/orange/white?text=HTOO",
    adminPassword: "123"
};

let tanks = [
    { tankNumber: 1, fuelType: "HSD", maxCapacity: 30500, currentCM: "120.5", currentLiter: 15200 },
    { tankNumber: 2, fuelType: "Premium Diesel", maxCapacity: 30500, currentCM: "175.2", currentLiter: 24100 },
    { tankNumber: 3, fuelType: "92 RON", maxCapacity: 30500, currentCM: "142.0", currentLiter: 18500 },
    { tankNumber: 4, fuelType: "92 RON", maxCapacity: 30500, currentCM: "70.4", currentLiter: 8900 },
    { tankNumber: 5, fuelType: "95 RON", maxCapacity: 15000, currentCM: "165.1", currentLiter: 12000 },
    { tankNumber: 6, fuelType: "92 RON", maxCapacity: 15000, currentCM: "65.3", currentLiter: 4500 }
];

app.get('/api/data', (req, res) => { res.json({ station: stationData, tanks: tanks }); });

app.post('/api/admin/update-station', (req, res) => {
    const { password, name, branchName, phoneNumber, logoUrl } = req.body;
    if (password !== stationData.adminPassword) return res.status(403).json({ success: false, message: "Password မှားယွင်းနေပါသည်။" });
    stationData.name = name; stationData.branchName = branchName; stationData.phoneNumber = phoneNumber; stationData.logoUrl = logoUrl;
    io.emit('dataUpdate', { station: stationData, tanks: tanks });
    res.json({ success: true, message: "Corporate profile updated successfully." });
});

// ✅ လက်မန်ဖြင့် ဝင်ပြင်ခြင်း API
app.post('/api/admin/update-all-tanks', (req, res) => {
    try {
        const password = req.body.password;
        if (password !== stationData.adminPassword) {
            return res.status(403).json({ success: false, message: "Password မှားယွင်းနေပါသည်။" });
        }

        for (let i = 1; i <= 6; i++) {
            const index = tanks.findIndex(t => t.tankNumber === i);
            if (index !== -1) {
                if (req.body[`tank${i}_fuelType`] !== undefined) tanks[index].fuelType = String(req.body[`tank${i}_fuelType`]).trim();
                if (req.body[`tank${i}_maxCapacity`] !== undefined) tanks[index].maxCapacity = Number(req.body[`tank${i}_maxCapacity`]) || 0;
                if (req.body[`tank${i}_currentCM`] !== undefined) tanks[index].currentCM = String(req.body[`tank${i}_currentCM`]).trim();
                if (req.body[`tank${i}_currentLiter`] !== undefined) tanks[index].currentLiter = Number(req.body[`tank${i}_currentLiter`]);
            }
        }

        io.emit('dataUpdate', { station: stationData, tanks: tanks });
        return res.json({ success: true, message: "တိုင်ကီအားလုံးကို အောင်မြင်စွာ သိမ်းဆည်းပြီးပါပြီဗျာ။" });
    } catch (error) {
        return res.status(500).json({ success: false, message: "Error: " + error.message });
    }
});

// ✅ စစ်ဆေးပြီးသား Excel Multi-Sheet & Multi-Column ဖတ်ရှုသည့် Logic အမှန်
app.post('/api/admin/upload-excel', upload.single('excelFile'), (req, res) => {
    const { password } = req.body;
    if (password !== stationData.adminPassword) {
        if(req.file) fs.unlinkSync(req.file.path);
        return res.status(403).json({ success: false, message: "Password မှားယွင်းနေပါသည်။" });
    }
    if (!req.file) return res.status(400).json({ success: false, message: "Excel ဖိုင်တင်ရန် လိုအပ်ပါသည်။" });

    try {
        const workbook = xlsx.readFile(req.file.path);
        const sheetNames = workbook.SheetNames;

        tanks.forEach((tank, tIndex) => {
            // စာလုံးအကြီးသေး၊ Space မရွေး တိုင်ကီနံပါတ်ပါဝင်သော Sheet ကို ရှာဖွေစစ်ဆေးသည်
            const matchedSheetName = sheetNames.find(name => {
                const cleanName = name.toLowerCase().replace(/[^0-9a-z]/g, '');
                return cleanName.includes(`tank${tank.tankNumber}`) || cleanName === `tank${tank.tankNumber}` || (cleanName.includes('tank') && cleanName.includes(String(tank.tankNumber)));
            });

            if (!matchedSheetName) return;

            const worksheet = workbook.Sheets[matchedSheetName];
            // defval: "" ကို သုံးပြီး null/undefined ဖြစ်ခြင်းမှ ကြိုတင်ကာကွယ်ထားသည်
            const rawRows = xlsx.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

            let maxSrNo = -1;
            let finalCM = "0";
            let finalLiter = 0;

            // Sheet ထဲရှိ သုံးရမည့် Row 7 မှစ၍ အကုန်လုံးကို လိုက်ဖတ်ပြီး အကြီးဆုံး Sr.No ရှိသည့် Row ကို ရှာဖွေသည်
            for (let r = 7; r < rawRows.length; r++) {
                let row = rawRows[r];
                if (!row || row.length === 0) continue;

                // ၃ ကော်လံစီ (Sr.No, CM, Liter) ခွဲလျက် ဘေးတိုက်ရှိနေသော ဒေတာအတွဲများကို ပတ်ဖတ်သည်
                for (let colIdx = 0; colIdx < row.length; colIdx += 3) {
                    if (row[colIdx] === undefined || row[colIdx] === null) continue;
                   
                    let srVal = parseInt(String(row[colIdx]).trim());
                    let cmVal = row[colIdx + 1] !== undefined && row[colIdx + 1] !== null ? String(row[colIdx + 1]).trim() : "";
                    let litVal = row[colIdx + 2] !== undefined && row[colIdx + 2] !== null ? String(row[colIdx + 2]).trim() : "";

                    if (!isNaN(srVal) && litVal !== "") {
                        let cleanLit = parseInt(litVal.replace(/,/g, ''));
                        if (!isNaN(cleanLit) && cleanLit >= 0) {
                            // အစဉ်လိုက်ဖြစ်နေသော Sr.No ထဲက အကြီးဆုံး (နောက်ဆုံးထွက်) ဒေတာကို ကောက်ယူသည်
                            if (srVal > maxSrNo) {
                                maxSrNo = srVal;
                                finalLiter = cleanLit;
                                finalCM = cmVal.replace('..', '.') || "0";
                            }
                        }
                    }
                }
            }

            // အကယ်၍ ကိုက်ညီသော ဒေတာအတွဲ တွေ့ရှိပါက သိမ်းဆည်းမည်
            if (maxSrNo !== -1) {
                tanks[tIndex].currentLiter = finalLiter;
                tanks[tIndex].currentCM = finalCM;
            }
        });

        fs.unlinkSync(req.file.path);
        io.emit('dataUpdate', { station: stationData, tanks: tanks });
        return res.json({ success: true, message: "Excel ဖိုင်ရှိ Sheet အားလုံးကို အောင်မြင်စွာ ဖတ်ရှုပြင်ဆင်ပြီးပါပြီဗျာ။" });

    } catch (error) {
        if(req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.status(500).json({ success: false, message: "Excel Read Error: " + error.message });
    }
});

io.on('connection', (socket) => { console.log('Client Connected'); });

const PORT = process.env.PORT || 3000;
http.listen(PORT, () => { console.log(`Server running on port ${PORT}`); });
