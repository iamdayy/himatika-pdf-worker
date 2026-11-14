# ⚠️ DEPRECATED: This Dockerfile is NOT used in production.
# Production deployment uses Vercel Serverless (see vercel.json).
FROM python:3.12

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    ffmpeg \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8080
ENV PORT=8080
ENV PYTHONUNBUFFERED=1

# Worker mode: same app, serves /api/media/compress-video/run for QStash.
# Deploy this image on a persistent host (Railway/Fly/VPS) so ffmpeg + threads survive.
CMD ["python", "run_waitress.py"]
