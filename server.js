const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const cors = require('cors');

const app = express();
const upload = multer({ dest: '/tmp/uploads/' }); // Render အတွက် tmp သုံးတာ ပိုကောင်းပါတယ်

app.use(cors());
app.use(express.json());
app.use(express.static('public'));

const DATA_FILE = '/tmp/data.json'; // Render မှာ data မပျောက်အောင် tmp ထဲခဏသိမ်းမယ်

if (!fs.existsSync(DATA_FILE)) {
    const initialData = {
        station: { name: "Htoo Fuel Station", branchName: "ပြင်ဦးလွင်ဘဏ်ခွဲ", phoneNumber: "09-123456789", logoUrl: "" },
        tanks: Array.from({ length: 6 }, (_, i) => ({
            tankNumber: i + 1, fuelType: "92 RON", currentCM: "0", currentLiter: "0", maxCapacity: 30500
        }))
    };
    fs.writeFileSync(DATA_FILE, JSON.stringify(initialData));
}

app.get('/api/data', (req, res) => {
    const data = JSON.parse(fs.readFileSync(DATA_FILE));
    res.json(data);
});

app.post('/api/manual-save', (req, res) => {
    fs.writeFileSync(DATA_FILE, JSON.stringify(req.body, null, 2));
    res.json({ success: true });
});

app.post('/api/excel-upload', upload.single('excelFile'), (req, res) => {
    try {
        const workbook = xlsx.readFile(req.file.path);
        const rows = xlsx.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        const lastRow = rows[rows.length - 1];
        let currentData = JSON.parse(fs.readFileSync(DATA_FILE));
        currentData.tanks = currentData.tanks.map((tank, i) => {
            const n = i + 1;
            return {
                ...tank,
                currentLiter: lastRow[`Tank${n}_Liter`] || tank.currentLiter,
                currentCM: lastRow[`Tank${n}_CM`] || tank.currentCM
            };
        });
        fs.writeFileSync(DATA_FILE, JSON.stringify(currentData, null, 2));
        fs.unlinkSync(req.file.path);
        res.json({ success: true });
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server is running on port ${PORT}`));
