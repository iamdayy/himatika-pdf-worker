# Deploy Worker (QStash target)

Repo ini bisa jalan dalam **2 mode**:

| Mode | Deploy | Fungsi |
|------|--------|--------|
| API serverless | Vercel (`vercel.json`) | Semua endpoint **kecuali** kompresi video. `POST /api/media/compress-video` kembali **503** karena ffmpeg tidak tersedia di Vercel Serverless. |
| Worker | Railway / Fly.io / VPS pakai `Dockerfile` ini | Menjalankan `POST /api/media/compress-video/run` yang dipanggil QStash. ffmpeg tersedia di image ini. |

## Jika tetap ingin satu Vercel saja

Ganti blok 503 di `api/handlers/media.py` dengan panggilan HTTP ke transcoding API eksternal (Mux, Cloudflare Stream, AWS MediaConvert, dst.), lalu `POST` hasilnya ke `callbackUrl`. Tidak perlu worker terpisah.

## Langkah deploy worker (kalau later mau aktifkan fitur video)

1. Deploy repo ini di Railway/Fly pakai Dockerfile.
2. Set env di worker: semua env Vercel + `MEDIA_WORKER_SECRET` + `PORT=8080`.
3. Di Vercel set `QSTASH_TOKEN`, `MEDIA_WORKER_URL`, `MEDIA_WORKER_SECRET`.
4. Kembalikan route `/run` di `media.py` dan imbal endpoint 503 dengan enqueue ke QStash.

## Cek worker hidup

```
curl https://<worker>/healthz
# 200 {"status":"success",...} kalau semua env + R2 OK
```

`/run` sengaja tanpa JWT (QStash tidak bawa JWT backend) — diamankan oleh `X-Worker-Secret`.
