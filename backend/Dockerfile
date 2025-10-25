# استخدم نسخة خفيفة من Python
FROM python:3.11-slim

# تحديد مجلد العمل داخل الحاوية
WORKDIR /app

# نسخ كل ملفات المشروع
COPY . /app

# تثبيت المتطلبات
RUN pip install --no-cache-dir -r backend/requirements.txt

# قراءة المنفذ من متغير البيئة (Deployra يمرره تلقائيًا)
ENV PORT=8080

# فتح المنفذ (اختياري للتوثيق فقط)
EXPOSE 8080

# تشغيل التطبيق عبر gunicorn وربطه بالمنفذ الصحيح
CMD ["sh", "-c", "gunicorn backend.app:app --bind 0.0.0.0:${PORT}"]
