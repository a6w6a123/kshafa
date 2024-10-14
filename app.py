from flask import Flask, request, render_template
import pandas as pd

# إنشاء تطبيق Flask
app = Flask(__name__)

# تحميل بيانات ملف Excel
file_path = 'sample_student_data.xlsx'
data = pd.read_excel(file_path)

# تحويل "رقم الهوية" إلى نص للتوافق مع الإدخال
data['رقم الهوية'] = data['رقم الهوية'].astype(str)

# دالة البحث عن البيانات بناءً على رقم الهوية
def search_by_id(id_number):
    result = data[data['رقم الهوية'] == id_number]
    if not result.empty:
        return {
            "الاسم": result['الاسم الثلاثي'].values[0],
            "رقم الجوال": result.get('رقم الجوال', 'غير متوفر'),
            "الصف": result['الصف'].values[0],
            "الشعبة": result['الشعبة'].values[0],
            "الهويات": result['الهويات'].values[0],
            "المهارات": result['المهارات'].values[0]
        }
    else:
        return None

# الصفحة الرئيسية
@app.route('/', methods=['GET', 'POST'])
def index():
    result = None
    id_number = None
    if request.method == 'POST':
        id_number = request.form['id_number']
        result = search_by_id(id_number)
    return render_template('index.html', result=result, id_number=id_number)

# تشغيل التطبيق
if __name__ == '__main__':
    app.run(debug=True)
