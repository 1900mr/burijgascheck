// استيراد المكتبات المطلوبة
const TelegramBot = require('node-telegram-bot-api');
const ExcelJS = require('exceljs');  // استيراد مكتبة exceljs
require('dotenv').config();  // إذا كنت تستخدم متغيرات بيئية
const express = require('express');  // إضافة Express لتشغيل السيرفر

// إعداد سيرفر Express (لتشغيل التطبيق على Render أو في بيئة محلية)
const app = express();

// تحديد المنفذ باستخدام متغير البيئة PORT
const port = process.env.PORT || 4000;  // إذا لم يكن هناك PORT في البيئة، سيعمل على 4000

// استبدل 'YOUR_BOT_TOKEN_HERE' بالتوكن الخاص بالبوت
const token = process.env.TELEGRAM_BOT_TOKEN || '7201507244:AAFmUzJTZ0CuhWxTE_BjwQJ-XB3RXlYMKYU';

// إنشاء البوت مع التفعيل
const bot = new TelegramBot(token, { polling: true });

// تحميل البيانات من ملف Excel
let gasburij = {};

// قراءة البيانات من ملف Excel باستخدام exceljs
async function loadDataFromExcel() {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile('gas18-11-2024.xlsx');  // تأكد من أن اسم الملف صحيح
        const worksheet = workbook.worksheets[0];  // الحصول على أول ورقة عمل
        
        worksheet.eachRow((row, rowNumber) => {
            const idnumber = row.getCell(1).value;  // أول عمود يحتوي على اسم الطالب
            const name = row.getCell(2).value;  // ثاني عمود يحتوي على النتيجة
            const jawwal = row.getCell(3).value;
            const place = row.getCell(6).value;
            const gasman = row.getCell(8).value;
            const jawwalgasman = row.getCell(9).value;

            if (idnumber && name) {
                gasburij[idnumber.trim()] = name.trim();
            }
        });

        console.log('تم تحميل البيانات بنجاح.');
    } catch (error) {
        console.error('حدث خطأ أثناء قراءة ملف Excel:', error.message);
    }
}

// تحميل البيانات عند بدء التشغيل
loadDataFromExcel();

// الرد عند بدء المحادثة
bot.onText(/\/start/, (msg) => {
    bot.sendMessage(msg.chat.id, "مرحبًا! أدخل اسمك للحصول على نتيجتك.");
});

// الرد عند استقبال رسالة
bot.on('message', (msg) => {
    const chatId = msg.chat.id;
    const idnumber = msg.text.trim();

    if (idnumber === '/start') return; // تجاهل أمر /start

    const name = gasburij[idnumber];
    if (name) {
        bot.sendMessage(chatId, `نرجو التوجه مباشرة للموزع ${idnumber}: ${name} : ${jawwal} : ${place} : ${gasman} : ${jawwalgasman} `);
    } else {
        bot.sendMessage(chatId, "عذرًا، لم أتمكن من العثور على اسمك.");
    }
});

// بدء السيرفر
app.listen(port, () => {
    console.log(`Server is running on port ${port}`);
});
