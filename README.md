# himatika-pdf-worker

Service pendamping platform **HIMATIKA** (Himpunan Mahasiswa Informatika ITSNU Pekalongan) untuk menghasilkan PDF, spreadsheet, dan QR code.

Service ini dipanggil **server-to-server** oleh backend HIMATIKA — bukan aplikasi yang dibuka langsung oleh pengguna, dan tidak punya antarmuka web publik.

## Fungsi

- **PDF** — tiket, sertifikat, dan dokumen (ReportLab, PyMuPDF, pypdf)
- **Spreadsheet / Excel** — generate dan ekspor data (openpyxl)
- **QR code** — generate dan baca QR (qrcode, zxing-cpp)
- **Media** — pemrosesan file dan penyimpanan ke R2/S3

## Tech Stack

Python 3.12, Flask 3, PyMongo, boto3 (R2/S3), ReportLab, PyMuPDF, pypdf, openpyxl, qrcode, zxing-cpp, PyJWT, Waitress.

## Cara Kerja & Autentikasi

- Entrypoint: `api/index.py` (dijalankan sebagai `api/index.py` di Vercel via `@vercel/python`).
- Semua request — kecuali `GET /` dan `GET /healthz` — wajib membawa header `Authorization: Bearer <token>` (JWT HS256) dengan klaim `service` bernilai `"himatika-backend"`. Tanpa token yang valid, request ditolak 401/403.
- `GET /` dan `GET /healthz` adalah health check tanpa autentikasi (`/healthz` memeriksa dependensi: MongoDB, R2, dan JWT secret).
- Handler dikelompokkan dalam lima blueprint di `api/handlers/`: `pdf`, `sheet`, `sign`, `tools`, `media`.

## Konfigurasi

Salin `.env.example` menjadi `.env` dan isi semua variabel:

| Variabel | Keterangan |
|---|---|
| `HIMATIKA_MONGODB_URI` | Connection string MongoDB HIMATIKA |
| `DBNAME` | Nama database (default di contoh: `himatika`) |
| `R2_ACCOUNT_ID` | Account ID Cloudflare R2 |
| `R2_ACCESS_KEY_ID` | Access key R2 |
| `R2_SECRET_ACCESS_KEY` | Secret key R2 |
| `R2_BUCKET_NAME` | Nama bucket R2 |
| `R2_PUBLIC_DOMAIN` | Domain publik file R2 |
| `HIMATIKA_JWT_SECRET` | Secret JWT yang sama dengan yang dipakai backend HIMATIKA |

Tanpa `HIMATIKA_JWT_SECRET` (atau `JWT_SECRET`), semua request terautentikasi akan ditolak.

## Cara Menjalankan

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env   # isi semua variabel di atas
python api/index.py    # mode pengembangan, port 5000
```

Untuk deployment, service ini dikonfigurasi untuk Vercel — lihat `vercel.json` (build `@vercel/python`, entry `api/index.py`).

## Pengujian API

Koleksi Postman untuk pengecekan API lengkap tersedia di folder `postman/`.

## Repo Terkait

- [himatika](https://github.com/iamdayy/himatika) — platform utama yang memanggil service ini
- [himatika-docs](https://github.com/iamdayy/himatika-docs) — dokumentasi HIMATIKA

## Status

Aktif — commit terakhir 14 November 2026.
