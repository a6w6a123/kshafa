const express = require('express');
const bodyParser = require('body-parser');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const filePath = path.resolve(__dirname, 'kashafa_registration.xlsx');

app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static('public'));

function createExcelFile() {
    const header = [['الاسم', 'رقم الهوية', 'الصف', 'الشعبة', 'رقم الجوال', 'الهوايات', 'المهارات']];
    const worksheet = XLSX.utils.aoa_to_sheet(header);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    XLSX.writeFile(workbook, filePath);
}

if (!fs.existsSync(filePath)) {
    createExcelFile();
}

app.post('/register', (req, res) => {
    try {
        const { fullName, idNumber, grade, classSection, phone, hobbies, skills } = req.body;

        const workbook = XLSX.readFile(filePath);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];

        const data = XLSX.utils.sheet_to_json(worksheet);
        data.push({
            'الاسم': fullName,
            'رقم الهوية': idNumber,
            'الصف': grade,
            'الشعبة': classSection,
            'رقم الجوال': phone,
            'الهوايات': hobbies || '',
            'المهارات': skills || ''
        });

        const newWorksheet = XLSX.utils.json_to_sheet(data);
        workbook.Sheets[workbook.SheetNames[0]] = newWorksheet;
        XLSX.writeFile(workbook, filePath);

        res.send('تم التسجيل بنجاح!');
    } catch (error) {
        console.error('Error:', error);
        res.status(500).send('حدث خطأ أثناء التسجيل.');
    }
});

app.listen(3000, () => console.log('Server running on http://localhost:3000'));
