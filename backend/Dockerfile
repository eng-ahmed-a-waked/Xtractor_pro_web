FROM python:3.11-slim
WORKDIR /app
COPY . /app
RUN pip install --no-cache-dir -r backend/requirements.txt
EXPOSE 5000
CMD ["gunicorn", "backend.app:app", "--bind", "0.0.0.0:5000"]
