# استخدم نسخة Python خفيفة وسريعة
FROM python:3.11.0

# تحديد مجلد العمل داخل الحاوية
WORKDIR /app

# نسخ ملفات المشروع كلها داخل الحاوية
COPY . /app

# تثبيت المتطلبات
RUN pip install --no-cache-dir -r backend/requirements.txt

# فتح المنفذ الذي سيعمل عليه Flask
EXPOSE 5000

# الأمر الذي يشغل التطبيق
CMD ["gunicorn", "backend.app:app", "--bind", "0.0.0.0:5000"]
