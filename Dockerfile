# استخدام Python 3.11 slim للحجم الأصغر
FROM python:3.11-slim

# تعيين متغيرات البيئة
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PORT=8080

# تثبيت المكتبات النظامية المطلوبة
RUN apt-get update && apt-get install -y \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# إنشاء مجلد العمل
WORKDIR /app

# نسخ ملف requirements أولاً للاستفادة من Docker cache
COPY backend/requirements.txt /app/backend/

# تثبيت المتطلبات
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r backend/requirements.txt

# نسخ كل ملفات المشروع
COPY . /app/

# إنشاء المجلدات المطلوبة
RUN mkdir -p /app/backend/data \
    /app/backend/uploads \
    /app/backend/outputs \
    /app/frontend

# إعطاء صلاحيات الكتابة
RUN chmod -R 755 /app/backend/data \
    /app/backend/uploads \
    /app/backend/outputs

# فتح المنفذ
EXPOSE 8080

# الانتقال لمجلد backend
WORKDIR /app/backend

# تشغيل التطبيق باستخدام gunicorn
CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:8080", "--workers", "2", "--timeout", "300", "--access-logfile", "-", "--error-logfile", "-"]
