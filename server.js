app.post('/register', (req, res) => {
    try {
        const { fullName, idNumber, grade, classSection, phone, hobbies, skills } = req.body;
        console.log('Received data:', req.body);  // تتبع البيانات المستلمة

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
        data.push(newData);

        const newWorksheet = XLSX.utils.json_to_sheet(data);
        workbook.Sheets[workbook.SheetNames[0]] = newWorksheet;
        XLSX.writeFile(workbook, filePath);  // حفظ الملف
        console.log('Data saved successfully');  // تأكيد حفظ البيانات

        res.send('تم التسجيل بنجاح!');
    } catch (error) {
        console.error('Error writing data to Excel:', error);
        res.status(500).send('حدث خطأ أثناء حفظ البيانات.');
    }
});
