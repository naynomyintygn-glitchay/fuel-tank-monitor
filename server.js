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

// ဖိုင်တင်ရန် သတ်မှတ်ချက်
const upload = multer({ dest: 'uploads/' });

// ဆိုင်အချက်အလက် (Default)
let stationData = {
    name: "Htoo Fuel Station",
    branchName: "ပြင်ဦးလွင်ဘဏ်ခွဲ",
    phoneNumber: "09-123456789",
    logoUrl: "https://placehold.co/100x100/orange/white?text=HTOO",
    adminPassword: "123"
};

// တိုင်ကီ (၆) လုံး အချက်အလက် (Default)
let tanks = [
    { tankNumber: 1, fuelType: "HSD", maxCapacity: 30500, currentCM: "0", currentLiter: 0 },
    { tankNumber: 2, fuelType: "Premium Diesel", maxCapacity: 30500, currentCM: "0", currentLiter: 0 },
    { tankNumber: 3, fuelType: "92 RON", maxCapacity: 30500, currentCM: "0", currentLiter: 0 },
    { tankNumber: 4, fuelType: "92 RON", maxCapacity: 30500, currentCM: "0", currentLiter: 0 },
    { tankNumber: 5, fuelType: "95 RON", maxCapacity: 15000, currentCM: "0", currentLiter: 0 },
    { tankNumber: 6, fuelType: "92 RON", maxCapacity: 15000, currentCM: "0", currentLiter: 0 }
];

// Front-end ကနေ ဒေတာလှမ်းတောင်းလျှင် ပေးမည့် API
app.get('/api/data', (req, res) => {
    res.json({ station: stationData, tanks: tanks });
});

// Admin ကနေ တိုင်ကီဒေတာအားလုံးကို Form နှင့် ကိုယ်တိုင်ပြင်လျှင် သုံးမည့် API
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
                if (req.body[`tank${i}_currentLiter`] !== undefined) tanks[index].currentLiter = Number(req.body[`tank${i}_currentLiter`]) || 0;
            }
        }
       
        io.emit('dataUpdate', { station: stationData, tanks: tanks });
        return res.json({ success: true, message: "တိုင်ကီအားလုံးကို အောင်မြင်စွာ သိမ်းဆည်းပြီးပါပြီဗျာ။" });
    } catch (error) {
        return res.status(500).json({ success: false, message: "Error: " + error.message });
    }
});

// ✅ မူရင်း Excel Dip Chart (Sheet ၆ ခုစလုံး) ကို အမှားအယွင်းမရှိ အော်တိုဖတ်ပေးမည့် စွမ်းထက်ဆုံး Excel Upload API
app.post('/api/admin/upload-excel', upload.single('excelFile'), (req, res) => {
    const { password } = req.body;
   
    // Password စစ်ဆေးခြင်း
    if (password !== stationData.adminPassword) {
        if (req.file) fs.unlinkSync(req.file.path);
        return res.status(403).json({ success: false, message: "Password မှားယွင်းနေပါသည်။" });
    }
    if (!req.file) return res.status(400).json({ success: false, message: "Excel ဖိုင်တင်ရန် လိုအပ်ပါသည်။" });

    try {
        // Excel ဖိုင်ကို ဖတ်ခြင်း
        const workbook = xlsx.readFile(req.file.path);
       
        // Sheet အားလုံးကို တစ်ခုချင်းစီပတ်ပြီး သက်ဆိုင်ရာ Tank ဒေတာ ရှာဖွေခြင်း
        workbook.SheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
           
            // Row ဒေတာများကို Matrix (2D Array) အဖြစ် ပြောင်းလဲခြင်း (ခေါင်းစီးမျိုးစုံ ရှိနေနိုင်သောကြောင့် ဖြစ်သည်)
            const rangeData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
           
            let tankNumber = null;
            let maxCapacity = 0;
            let foundHeaderRow = -1;

            // (၁) အပေါ်ပိုင်း Header မှာ Tank နံပါတ် နှင့် Capacity ကို လိုက်ရှာခြင်း
            for (let i = 0; i < Math.min(rangeData.length, 10); i++) {
                const rowStr = JSON.stringify(rangeData[i] || []).toLowerCase();
               
                // Tank No. ရှာဖွေခြင်း
                if (rowStr.includes("tank no") || rowStr.includes("tankno")) {
                    const match = rowStr.match(/tank\s*(?:no\.?|number)?\s*\(?(\d+)\)?/i) || rowStr.match(/tank\s*no\s*(\d+)/i) || rowStr.match(/(\d+)/);
                    if (match) tankNumber = parseInt(match[1]);
                }
               
                // Capacity ရှာဖွေခြင်း
                if (rowStr.includes("capacity")) {
                    const capMatch = rowStr.match(/(\d[\d,]*)/);
                    if (capMatch) maxCapacity = parseInt(capMatch[1].replace(/,/g, ''));
                }
            }

            // အကယ်၍ Sheet နာမည်ထဲမှာ Tank ပါခဲ့ရင်လည်း နံပါတ်ထုတ်ယူခြင်း (ဥပမာ - Tank 1)
            if (!tankNumber) {
                const nameMatch = sheetName.match(/tank\s*(\d+)/i) || sheetName.match(/(\d+)/);
                if (nameMatch) tankNumber = parseInt(nameMatch[1]);
            }

            // တိုင်ကီနံပါတ် ရှိမှသာ ရှေ့ဆက်ဖတ်မည်
            if (tankNumber && tankNumber >= 1 && tankNumber <= 6) {
                const tIndex = tanks.findIndex(t => t.tankNumber === tankNumber);
                if (tIndex !== -1 && maxCapacity > 0) {
                    tanks[tIndex].maxCapacity = maxCapacity; // မူရင်း Capacity အတိုင်း အော်တိုပြင်ပေးသည်
                }

                let maxSrNo = -1;
                let finalCM = "0";
                let finalLiter = 0;

                // (၂) ဇယားကွက်ထဲမှ Sr.No, CM, Liter ဒေတာအတွဲများကို ဒေါင်လိုက် လိုက်ရှာခြင်း
                // ခေါင်းစီးအတန်း (Sr.No ပါဝင်သော အတန်း) ကို အရင်ရှာပါမည်
                for (let i = 0; i < rangeData.length; i++) {
                    const row = rangeData[i] || [];
                    if (row.some(cell => String(cell).toLowerCase().replace(/\s+/g, '') === 'sr.no')) {
                        foundHeaderRow = i;
                        break;
                    }
                }

                if (foundHeaderRow !== -1) {
                    const headerRow = rangeData[foundHeaderRow];
                   
                    // ခေါင်းစီးအတန်းအောက်က ဒေတာအားလုံးကို စစ်ဆေးခြင်း
                    for (let r = foundHeaderRow + 1; r < rangeData.length; r++) {
                        const row = rangeData[r] || [];
                       
                        // တစ်တန်းထဲမှာ ဒေတာအတွဲ ၃ တွဲရှိနိုင်သောကြောင့် ၃ ကော်လံစီ ခွဲထုတ်စစ်ဆေးခြင်း
                        for (let c = 0; c < row.length; c += 3) {
                            let srNoVal = parseInt(String(row[c]).trim());
                            let cmVal = row[c+1];
                            let litVal = row[c+2];

                            if (!isNaN(srNoVal) && cmVal !== undefined && litVal !== undefined) {
                                // စာရင်းထဲတွင် Sr.No အကြီးဆုံး (အောက်ဆုံး) ဖြစ်မည့် နောက်ဆုံးထွက် ဒေတာကို မှတ်သားခြင်း
                                if (srNoVal > maxSrNo) {
                                    maxSrNo = srNoVal;
                                    finalCM = String(cmVal).trim();
                                    finalLiter = parseInt(String(litVal).replace(/,/g, '').trim()) || 0;
                                }
                            }
                        }
                    }
                }

                // (၃) ရှာဖွေတွေ့ရှိသော နောက်ဆုံးထွက်တန်ဖိုးကို Memory ထဲတွင် သွားရောက် အပ်ဒိတ်လုပ်ခြင်း
                if (maxSrNo !== -1 && tIndex !== -1) {
                    tanks[tIndex].currentCM = finalCM;
                    tanks[tIndex].currentLiter = finalLiter;
                }
            }
        });

        // ဖိုင်ကို ရှင်းထုတ်ပြီး Socket.io ကနေ Front-end UI ဆီ Real-time လှမ်းပို့ခြင်း
        fs.unlinkSync(req.file.path);
        io.emit('dataUpdate', { station: stationData, tanks: tanks });
       
        return res.json({ success: true, message: "မူရင်း Excel Dip Chart ထဲမှ နောက်ဆုံးထွက် အချက်အလက်အားလုံးကို အောင်မြင်စွာ ဖတ်ပြီးပါပြီဗျာ။" });

    } catch (error) {
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.status(500).json({ success: false, message: "Excel Read Error: " + error.message });
    }
});

// Socket Connection ငြိမ်မငြိမ် စောင့်ကြည့်ခြင်း
io.on('connection', (socket) => {
    console.log('Client Connected');
});

// Port သတ်မှတ်ချက်
const PORT = process.env.PORT || 3000;
http.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
