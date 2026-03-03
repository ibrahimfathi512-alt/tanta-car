/**
 * Talabat Zone Management System - Dashboard Backend
 * Version: 2026.2.1
 * Author: Ibrahim Fathi & AI Collaborative
 * Description: Fixed compatibility for google-spreadsheet v3 (r.get replacement)
 */

const express = require('express');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
const session = require('express-session');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();

// --- Configuration ---
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// --- Session Management ---
app.use(session({
    secret: process.env.SESSION_SECRET || 'talabat-security-key-2026-pro',
    resave: false,
    saveUninitialized: false,
    cookie: { 
        maxAge: 24 * 60 * 60 * 1000, 
        secure: false 
    }
}));

const SPREADSHEET_ID = '1bNhlUVWnt43Pq1hqDALXbfGDVazD7VhaeKM58hBTsN0';

const zonePasswords = {
    'Ain shams': '754', 'Alexandria': '1234', 'Cairo_city_centre': '909',
    'Giza': '1568', 'Heliopolis': '2161', 'Ismalia city': '1122',
    'Kafr el-sheikh': '3344', 'Maadi': '878', 'Mansoura': '5566',
    'Mohandiseen': '1862', 'Nasr city': '2851', 'New damietta': '7788',
    'October': '2161', 'Portsaid city': '9900', 'Shebin el koom': '4455',
    'Sheikh zayed': '854', 'Suez': '6677', 'Tagammoa south': '1072',
    'Tanta': '8899', 'Zagazig': '2233'
};

const officePasswords = {
    'مكتب طلبات المنصوره': '1010', 'مكتب طلبات الأسكندريه': '2020',
    'مكتب طلبات مدينه نصر': '3030', 'مكتب طلبات أكتوبر': '4040',
    'مكتب طلبات الهرم': '5050', 'مكتب طلبات المعادي': '6060',
    'مكتب طلبات المهندسين': '7070', 'مكتب طلبات التجمع': '8080'
};

// --- Google Sheets Auth ---
async function getDoc() {
    let credsData;
    try {
        if (process.env.GOOGLE_CREDS) {
            credsData = JSON.parse(process.env.GOOGLE_CREDS);
        } else {
            const credsFilePath = path.join(__dirname, 'credentials.json');
            credsData = require(credsFilePath);
        }

        const doc = new GoogleSpreadsheet(SPREADSHEET_ID);
        await doc.useServiceAccountAuth({
            client_email: credsData.client_email,
            private_key: credsData.private_key.replace(/\\n/g, '\n'),
        });
        await doc.loadInfo();
        return doc;
    } catch (error) {
        console.error("FATAL ERROR:", error);
        throw error;
    }
}

const cleanData = (val) => {
    if (val === undefined || val === null || val === '') return 0;
    let strVal = val.toString().trim();
    if (['NA', '#N/A', 'N/A', '0'].includes(strVal)) return 0;
    let res = parseFloat(strVal.replace(/,/g, '').replace(/[^0-9.-]/g, ''));
    return isNaN(res) ? 0 : res;
};

const checkAuth = (req, res, next) => {
    if (!req.session.userZone) return res.redirect('/');
    next();
};

// ==========================================
// --- Routes (Fixed r.get issue) ---
// ==========================================

// [1] Login Page
app.get('/', async (req, res) => {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByIndex[0];
        const rows = await sheet.getRows();
        const allZones = [...new Set(rows.map(r => r.zone_name))].filter(z => z);
        res.render('login', { zones: allZones, error: null });
    } catch (e) {
        res.status(500).send("خطأ في الاتصال: " + e.message);
    }
});

// [2] Login Action
app.post('/login', (req, res) => {
    const { zone, password } = req.body;
    if (zonePasswords[zone] === password) {
        req.session.userZone = zone;
        res.redirect('/dashboard');
    } else {
        res.render('login', { zones: Object.keys(zonePasswords), error: 'خطأ في البيانات' });
    }
});

// [3] Dashboard
app.get('/dashboard', checkAuth, async (req, res) => {
    try {
        const doc = await getDoc();
        const mainSheet = doc.sheetsByIndex[0];
        const rows = await mainSheet.getRows();
        
        let myRiders = rows.filter(r => r.zone_name === req.session.userZone);

        const hiresSheet = doc.sheetsByTitle['تعيينات الشهر'];
        let newCount = 0;
        if (hiresSheet) {
            const hireRows = await hiresSheet.getRows();
            newCount = hireRows.filter(r => r.zone_name === req.session.userZone).length;
        }

        const stats = {
            total: myRiders.length,
            withShifts: myRiders.filter(r => cleanData(r['شيفتات الغد']) > 0).length,
            noShifts: myRiders.filter(r => cleanData(r['شيفتات الغد']) === 0).length,
            highWallet: myRiders.filter(r => cleanData(r['المحفظه']) > 1000).length,
            newCount: newCount
        };

        res.render('dashboard', { 
            riders: myRiders, zone: req.session.userZone, stats, 
            headers: mainSheet.headerValues, cleanData 
        });
    } catch (e) {
        res.status(500).send("خطأ: " + e.message);
    }
});

// [4] Inquiry Gate
app.get('/uploaded-inquiry', checkAuth, async (req, res) => {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['مرفوعين استعلام'];
        const rows = await sheet.getRows();
        const locations = [...new Set(rows.map(r => r['مقر التحضير']))].filter(l => l);
        res.render('inquiry_auth', { zone: req.session.userZone, locations, error: null });
    } catch (e) {
        res.status(500).send("خطأ: " + e.message);
    }
});

// [5] Inquiry Auth
app.post('/uploaded-inquiry-auth', checkAuth, async (req, res) => {
    const { password, location } = req.body;
    if (officePasswords[location] === password) {
        try {
            const doc = await getDoc();
            const sheet = doc.sheetsByTitle['مرفوعين استعلام'];
            const rows = await sheet.getRows();
            const filteredData = rows.filter(r => (r['مقر التحضير'] || "").trim() === location.trim());
            res.render('uploaded_inquiry', { 
                data: filteredData, zone: req.session.userZone, 
                location, headers: sheet.headerValues 
            });
        } catch (e) { res.status(500).send(e.message); }
    } else {
        res.redirect('/uploaded-inquiry?error=true');
    }
});

// [6] Wallets
app.get('/office-wallets', checkAuth, async (req, res) => {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['جميع المحافظ'];
        const rows = await sheet.getRows();
        let lastSeenDate = "";
        const processedWallets = rows.map(row => {
            let rowObj = row.toObject();
            let currentDate = row.Date;
            if (!currentDate || currentDate === '0') rowObj.Date = lastSeenDate;
            else { rowObj.Date = currentDate; lastSeenDate = currentDate; }
            return rowObj;
        });
        res.render('office_wallets', { wallets: processedWallets, zone: req.session.userZone, headers: sheet.headerValues });
    } catch (e) { res.status(500).send(e.message); }
});

// [7] Reconciliations
app.get('/reconciliations', checkAuth, async (req, res) => {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['تصالحات'];
        const rows = await sheet.getRows();
        let lastSeenDate = "";
        const processedData = rows.map(row => {
            let rowObj = row.toObject();
            let currentDate = row['التاريخ'];
            if (!currentDate) rowObj['التاريخ'] = lastSeenDate;
            else { rowObj['التاريخ'] = currentDate; lastSeenDate = currentDate; }
            return rowObj;
        });
        res.render('reconciliations', { data: processedData, zone: req.session.userZone, headers: sheet.headerValues });
    } catch (e) { res.status(500).send(e.message); }
});

// [8] Targets
app.get('/targets', checkAuth, async (req, res) => {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['التارجت'];
        const rows = await sheet.getRows();
        const zoneData = rows.find(r => r.zone_name === req.session.userZone);
        const mainRows = await doc.sheetsByIndex[0].getRows();
        const total = mainRows.filter(r => r.zone_name === req.session.userZone).length;
        res.render('targets', { zone: req.session.userZone, zoneData, stats: { total }, headers: sheet.headerValues, cleanData });
    } catch (e) { res.status(500).send(e.message); }
});

// [9] New Riders
app.get('/new-riders', checkAuth, async (req, res) => {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['تعيينات الشهر'];
        const rows = await sheet.getRows();
        const myRiders = rows.filter(r => r.zone_name === req.session.userZone);
        const stats = {
            total: myRiders.length,
            received: myRiders.filter(r => r['الحاله'] === 'استلم' || r['الحاله'] === 'تم الاستلام').length,
            notReceived: myRiders.filter(r => r['الحاله'] !== 'استلم' && r['الحاله'] !== 'تم الاستلام').length
        };
        res.render('new_riders', { riders: myRiders, zone: req.session.userZone, stats, headers: sheet.headerValues, cleanData });
    } catch (e) { res.status(500).send(e.message); }
});

// [10] Order Responses
app.get('/order-responses', checkAuth, async (req, res) => {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['ردود الأوردات'];
        const rows = await sheet.getRows();
        const myOrders = rows.filter(r => r.zone_name === req.session.userZone);
        res.render('order_responses', { orders: myOrders, zone: req.session.userZone, headers: sheet.headerValues });
    } catch (e) { res.status(500).send(e.message); }
});

// [11] Hiring Feedback
app.get('/new-riders-responses', checkAuth, async (req, res) => {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['ردود التعيينات'];
        const rows = await sheet.getRows();
        const myResponses = rows.filter(r => r['Zone Name'] === req.session.userZone);
        res.render('new_riders_responses', { responses: myResponses, zone: req.session.userZone, headers: sheet.headerValues });
    } catch (e) { res.status(500).send(e.message); }
});

// [12] Rejected Inquiry
app.get('/rejected-inquiry', checkAuth, async (req, res) => {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['مرفوضين استعلام'];
        const rows = await sheet.getRows();
        const allRejectedData = rows.map(row => ({
            date: row['التاريخ'], office: row['مكتب'], prep_office: row['مقر التحضير'],
            name: row['الاسم'], phone: row['رقم الهاتف'], national_id: row['الرقم القومي'],
            supervisor: row['اسم المشرف'], reason: row['سبب الرفض']
        }));
        res.render('rejected_inquiry', { data: allRejectedData, zone: req.session.userZone });
    } catch (e) { res.status(500).send(e.message); }
});

// [13] Download Excel
app.get('/download', checkAuth, async (req, res) => {
    try {
        const doc = await getDoc();
        const rows = await (doc.sheetsByIndex[0]).getRows();
        const myData = rows.filter(r => r.zone_name === req.session.userZone).map(r => r.toObject());
        const worksheet = XLSX.utils.json_to_sheet(myData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Data");
        const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
        res.setHeader('Content-Disposition', `attachment; filename=Report.xlsx`);
        res.type('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);
    } catch (e) { res.status(500).send(e.message); }
});

app.get('/logout', (req, res) => {
    req.session.destroy(() => res.redirect('/'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on ${PORT}`));

process.on('uncaughtException', (err) => console.error('CRITICAL:', err));