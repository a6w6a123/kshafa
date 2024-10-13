const express = require('express');
const bodyParser = require('body-parser');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static('public'));

// تحديد مسار ملف Excel
const filePath = path.resolve(__dirname, 'kashafa_registration.xlsx');

// دالة لإنشاء ملف Excel جديد
function createExcelFile() {
    const header = [['الاسم', 'رقم الهوية', 'الصف', 'الشعبة', 'رقم الجوال', 'الهوايات', 'المهارات']];
    const worksheet = XLSX.utils.aoa_to_sheet(header);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    XLSX.writeFile(workbook, filePath);
}

// التحقق من وجود الملف وإنشائه إذا لم يكن موجودًا
if (!fs.existsSync(filePath)) {
    createExcelFile();
    console.log('تم إنشاء ملف Excel جديد.');
}

app.post('/register', (req, res) => {
    try {
        const { fullName, idNumber, grade, classSection, phone, hobbies, skills } = req.body;

        // قراءة ملف Excel
        const workbook = XLSX.readFile(filePath);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];

        const newData = {
            'الاسم': fullName,
            'رقم الهوية': idNumber,
            'الصف': grade,
            'الشعبة': classSection,
            'رقم الجوال': phone,
            'الهوايات': hobbies || '',
            'المهارات': skills || ''
        };

        const data = XLSX.utils.sheet_to_json(worksheet);
        data.push(newData); // إضافة البيانات الجديدة إلى البيانات الحالية

        const newWorksheet = XLSX.utils.json_to_sheet(data);
        workbook.Sheets[workbook.SheetNames[0]] = newWorksheet;
        XLSX.writeFile(workbook, filePath); // حفظ الملف

        res.send('تم التسجيل بنجاح!');
    } catch (error) {
        console.error('حدث خطأ أثناء كتابة البيانات إلى ملف Excel:', error);
        res.status(500).send('حدث خطأ أثناء حفظ البيانات.');
    }
});

// تشغيل الخادم
app.listen(3000, () => {
    console.log('Server running on http://localhost:3000');
});
