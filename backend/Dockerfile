# استخدم نسخة خفيفة من Python
FROM python:3.11-slim

# تحديد مجلد العمل داخل الحاوية
WORKDIR /app

# نسخ كل ملفات المشروع
COPY . /app

# تثبيت المتطلبات
RUN pip install --no-cache-dir -r backend/requirements.txt

# فتح المنفذ الذي سيعمل عليه Flask
EXPOSE 5000

# تشغيل التطبيق من داخل مجلد backend
CMD ["gunicorn", "app:app", "--chdir", "backend", "--bind", "0.0.0.0:5000"]
