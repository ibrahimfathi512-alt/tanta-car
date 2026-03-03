/**
 * Talabat Zone Management System - Dashboard Backend
 * Version: 2026.2.1
 * Author: Ibrahim Fathi & AI Collaborative
 * Description: Node.js server handling Google Sheets data, custom office authentication, 
 * and multi-zone dashboard routing with session management.
 */

const express = require('express');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
const session = require('express-session');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();

// --- إعدادات المحرك وتنسيقات الملفات (Configuration) ---
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// --- إعدادات الجلسة (Session Management) ---
app.use(session({
    secret: process.env.SESSION_SECRET || 'talabat-security-key-2026-pro',
    resave: false,
    saveUninitialized: false,
    cookie: { 
        maxAge: 24 * 60 * 60 * 1000, // صلاحية الجلسة يوم كامل
        secure: false 
    }
}));

// --- معرف الشيت (Google Sheets ID) ---
const SPREADSHEET_ID = '1bNhlUVWnt43Pq1hqDALXbfGDVazD7VhaeKM58hBTsN0';

// --- كلمات مرور المناطق للدخول الرئيسي (Main Zone Passwords) ---
const zonePasswords = {
    'Ain shams': '754',
    'Alexandria': '1234',
    'Cairo_city_centre': '909',
    'Giza': '1568',
    'Heliopolis': '2161',
    'Ismalia city': '1122',
    'Kafr el-sheikh': '3344',
    'Maadi': '878',
    'Mansoura': '5566',
    'Mohandiseen': '1862',
    'Nasr city': '2851',
    'New damietta': '7788',
    'October': '2161',
    'Portsaid city': '9900',
    'Shebin el koom': '4455',
    'Sheikh zayed': '854',
    'Suez': '6677',
    'Tagammoa south': '1072',
    'Tanta': '8899',
    'Zagazig': '2233'
};

// --- كلمات مرور مقرات التحضير لصفحة الاستعلامات (Custom Office Passwords) ---
const officePasswords = {
    'مكتب طلبات المنصوره': '1010',
    'مكتب طلبات الأسكندريه': '2020',
    'مكتب طلبات مدينه نصر': '3030',
    'مكتب طلبات أكتوبر': '4040',
    'مكتب طلبات الهرم': '5050',
    'مكتب طلبات المعادي': '6060',
    'مكتب طلبات المهندسين': '7070',
    'مكتب طلبات التجمع': '8080'
};

// --- وظيفة الاتصال بشيت جوجل (Authentication Function) ---
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
        console.error("FATAL ERROR: Failed to connect to Google Sheets", error);
        throw error;
    }
}

// --- وظيفة تنظيف البيانات الرقمية (Data Sanitization) ---
const cleanData = (val) => {
    if (val === undefined || val === null || val === '') return 0;
    let strVal = val.toString().trim();
    if (['NA', '#N/A', 'N/A', '0'].includes(strVal)) return 0;
    let res = parseFloat(strVal.replace(/,/g, '').replace(/[^0-9.-]/g, ''));
    return isNaN(res) ? 0 : res;
};

// --- الميدل وير للتحقق من الدخول (Auth Middleware) ---
const checkAuth = (req, res, next) => {
    if (!req.session.userZone) return res.redirect('/');
    next();
};

// ==========================================
// --- المسارات الـ 13 (Route Definitions) ---
// ==========================================

// [1] صفحة تسجيل الدخول الرئيسية
app.get('/', async (req, res) => {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByIndex[0];
        const rows = await sheet.getRows();
        const allZones = [...new Set(rows.map(r => r.get('zone_name')))].filter(z => z);
        res.render('login', { zones: allZones, error: null });
    } catch (e) {
        res.status(500).send("عذراً، حدث خطأ في الاتصال بقاعدة البيانات: " + e.message);
    }
});

// [2] معالجة تسجيل الدخول الرئيسي
app.post('/login', (req, res) => {
    const { zone, password } = req.body;
    if (zonePasswords[zone] && zonePasswords[zone] === password) {
        req.session.userZone = zone;
        res.redirect('/dashboard');
    } else {
        const zones = Object.keys(zonePasswords);
        res.render('login', { zones, error: 'كلمة المرور غير صحيحة أو المنطقة غير موجودة' });
    }
});

// [3] لوحة التحكم الرئيسية (Dashboard Analytics)
app.get('/dashboard', checkAuth, async (req, res) => {
    try {
        const doc = await getDoc();
        const mainSheet = doc.sheetsByIndex[0];
        const rows = await mainSheet.getRows();
        
        let myRiders = rows.filter(r => r.get('zone_name') === req.session.userZone);

        const hiresSheet = doc.sheetsByTitle['تعيينات الشهر'];
        let newCount = 0;
        if (hiresSheet) {
            const hireRows = await hiresSheet.getRows();
            newCount = hireRows.filter(r => r.get('zone_name') === req.session.userZone).length;
        }

        const stats = {
            total: myRiders.length,
            withShifts: myRiders.filter(r => cleanData(r.get('شيفتات الغد')) > 0).length,
            noShifts: myRiders.filter(r => cleanData(r.get('شيفتات الغد')) === 0).length,
            highWallet: myRiders.filter(r => cleanData(r.get('المحفظه')) > 1000).length,
            newCount: newCount
        };

        res.render('dashboard', { 
            riders: myRiders, 
            zone: req.session.userZone, 
            stats, 
            headers: mainSheet.headerValues, 
            cleanData 
        });
    } catch (e) {
        res.status(500).send("خطأ في تحميل لوحة التحكم: " + e.message);
    }
});

// [4] بوابة أمان صفحة "مرفوعين استعلام" (Custom Login)
app.get('/uploaded-inquiry', checkAuth, async (req, res) => {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['مرفوعين استعلام'];
        if (!sheet) throw new Error("شيت 'مرفوعين استعلام' غير موجود");
        
        const rows = await sheet.getRows();
        const locations = [...new Set(rows.map(r => r.get('مقر التحضير')))].filter(l => l && l.trim() !== "");

        res.render('inquiry_auth', { 
            zone: req.session.userZone, 
            locations, 
            error: null 
        });
    } catch (e) {
        res.status(500).send("خطأ في جلب بيانات المقرات: " + e.message);
    }
});

// [5] التحقق بباسورد المقر المخصص وعرض البيانات المفلترة
app.post('/uploaded-inquiry-auth', checkAuth, async (req, res) => {
    const { password, location } = req.body;

    if (officePasswords[location] === password) {
        try {
            const doc = await getDoc();
            const sheet = doc.sheetsByTitle['مرفوعين استعلام'];
            const rows = await sheet.getRows();
            
            const filteredData = rows.filter(r => (r.get('مقر التحضير') || "").trim() === location.trim());

            res.render('uploaded_inquiry', { 
                data: filteredData, 
                zone: req.session.userZone, 
                location: location, 
                headers: sheet.headerValues 
            });
        } catch (e) { 
            res.status(500).send("خطأ أثناء معالجة بيانات الاستعلام: " + e.message); 
        }
    } else {
        res.redirect('/uploaded-inquiry?error=true');
    }
});

// [6] صفحة جميع المحافظ (Historical Wallets)
app.get('/office-wallets', checkAuth, async (req, res) => {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['جميع المحافظ'];
        if (!sheet) throw new Error("شيت 'جميع المحافظ' غير موجود");
        
        const rows = await sheet.getRows();
        let lastSeenDate = "";
        
        const processedWallets = rows.map(row => {
            let rowObj = row.toObject();
            let currentDate = row.get('Date');
            if (!currentDate || currentDate === '0' || currentDate === '') {
                rowObj.Date = lastSeenDate;
            } else {
                rowObj.Date = currentDate;
                lastSeenDate = currentDate;
            }
            return rowObj;
        });

        res.render('office_wallets', { 
            wallets: processedWallets, 
            zone: req.session.userZone, 
            headers: sheet.headerValues 
        });
    } catch (e) {
        res.status(500).send("خطأ في شيت المحافظ: " + e.message);
    }
});

// [7] صفحة التصالحات (Reconciliations)
app.get('/reconciliations', checkAuth, async (req, res) => {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['تصالحات'];
        if (!sheet) throw new Error("شيت 'تصالحات' غير موجود");
        
        const rows = await sheet.getRows();
        let lastSeenDate = "";
        
        const processedData = rows.map(row => {
            let rowObj = row.toObject();
            let currentDate = row.get('التاريخ');
            if (!currentDate || currentDate === '') {
                rowObj.التاريخ = lastSeenDate;
            } else {
                rowObj.التاريخ = currentDate;
                lastSeenDate = currentDate;
            }
            return rowObj;
        });

        res.render('reconciliations', { 
            data: processedData, 
            zone: req.session.userZone, 
            headers: sheet.headerValues 
        });
    } catch (e) {
        res.status(500).send("حدث خطأ، تأكد من وجود شيت باسم 'تصالحات'");
    }
});

// [8] صفحة تتبع الأهداف (Target Management)
app.get('/targets', checkAuth, async (req, res) => {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['التارجت'];
        const rows = await sheet.getRows();
        
        const zoneData = rows.find(r => r.get('zone_name') === req.session.userZone);
        
        const mainSheet = doc.sheetsByIndex[0];
        const mainRows = await mainSheet.getRows();
        const myRiders = mainRows.filter(r => r.get('zone_name') === req.session.userZone);
        
        res.render('targets', { 
            zone: req.session.userZone, 
            zoneData, 
            stats: { total: myRiders.length }, 
            headers: sheet.headerValues, 
            cleanData 
        });
    } catch (e) {
        res.status(500).send("خطأ: تأكد من وجود شيت باسم 'التارجت'");
    }
});

// [9] صفحة تعيينات الشهر - (تم الحل: إرسال الـ stats لإصلاح الخطأ)
app.get('/new-riders', checkAuth, async (req, res) => {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['تعيينات الشهر'];
        if (!sheet) throw new Error("شيت 'تعيينات الشهر' غير موجود");
        
        const rows = await sheet.getRows();
        const myRiders = rows.filter(r => r.get('zone_name') === req.session.userZone);

        // إنشاء كائن stats المطلوب في ملف الـ EJS لرسم الـ Pie Chart
        const stats = {
            total: myRiders.length,
            received: myRiders.filter(r => r.get('الحاله') === 'استلم' || r.get('الحاله') === 'تم الاستلام').length,
            notReceived: myRiders.filter(r => r.get('الحاله') !== 'استلم' && r.get('الحاله') !== 'تم الاستلام').length
        };

        res.render('new_riders', { 
            riders: myRiders, 
            zone: req.session.userZone, 
            stats: stats, // الآن المتغير متاح ولن يظهر ReferenceError
            headers: sheet.headerValues, 
            cleanData 
        });
    } catch (e) {
        console.error(e);
        res.status(500).send("تأكد من وجود شيت باسم 'تعيينات الشهر' بشكل صحيح");
    }
});

// [10] صفحة ردود الأوردات (Order Status Responses)
app.get('/order-responses', checkAuth, async (req, res) => {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['ردود الأوردات'];
        const rows = await sheet.getRows();
        
        const myOrders = rows.filter(r => r.get('zone_name') === req.session.userZone);
        
        res.render('order_responses', { 
            orders: myOrders, 
            zone: req.session.userZone, 
            headers: sheet.headerValues 
        });
    } catch (e) {
        res.status(500).send("تأكد من وجود شيت باسم 'ردود الأوردات'");
    }
});

// [11] صفحة ردود التعيينات (Hiring Feedback)
app.get('/new-riders-responses', checkAuth, async (req, res) => {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['ردود التعيينات'];
        const rows = await sheet.getRows();
        
        const myResponses = rows.filter(r => r.get('Zone Name') === req.session.userZone);
        
        res.render('new_riders_responses', { 
            responses: myResponses, 
            zone: req.session.userZone, 
            headers: sheet.headerValues 
        });
    } catch (e) {
        res.status(500).send("خطأ: تأكد من وجود شيت باسم 'ردود التعيينات'.");
    }
});

// [12] صفحة مرفوضين الاستعلام (Security Rejected)
app.get('/rejected-inquiry', checkAuth, async (req, res) => {
    try {
        const doc = await getDoc();
        const sheet = doc.sheetsByTitle['مرفوضين استعلام'];
        const rows = await sheet.getRows();
        
        const allRejectedData = rows.map(row => ({
            date: row.get('التاريخ'),
            office: row.get('مكتب'),
            prep_office: row.get('مقر التحضير'),
            name: row.get('الاسم'),
            phone: row.get('رقم الهاتف'),
            national_id: row.get('الرقم القومي'),
            supervisor: row.get('اسم المشرف'),
            reason: row.get('سبب الرفض')
        }));

        res.render('rejected_inquiry', { 
            data: allRejectedData,
            zone: req.session.userZone 
        });
    } catch (e) {
        res.status(500).send("خطأ في الوصول لشيت 'مرفوضين استعلام'");
    }
});

// [13] مسار تصدير البيانات وتحميلها (Excel Export Service)
app.get('/download', checkAuth, async (req, res) => {
    try {
        const doc = await getDoc();
        const rows = await (doc.sheetsByIndex[0]).getRows();
        
        const myData = rows
            .filter(r => r.get('zone_name') === req.session.userZone)
            .map(r => r.toObject());

        const worksheet = XLSX.utils.json_to_sheet(myData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Riders_Data");
        
        const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

        res.setHeader('Content-Disposition', `attachment; filename=Talabat_${req.session.userZone}_Report.xlsx`);
        res.type('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);
    } catch (e) {
        res.status(500).send("عذراً، فشل تصدير البيانات لملف إكسيل: " + e.message);
    }
});

// --- تسجيل الخروج (Logout) ---
app.get('/logout', (req, res) => {
    req.session.destroy((err) => {
        if(err) console.log(err);
        res.redirect('/');
    });
});

// --- تشغيل السيرفر (Server Listener) ---
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`
    =======================================================
    🚀 Talabat Pro System Started Successfully!
    📍 Mode: Multi-Office Security & Stats Fixed
    📡 Port: ${PORT}
    🔗 URL: http://localhost:${PORT}
    =======================================================
    `);
});

// --- معالجة الأخطاء غير المتوقعة (Global Error Handling) ---
process.on('uncaughtException', (err) => {
    console.error('CRITICAL ERROR:', err);
});