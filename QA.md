# QA & Testing Guide

This document outlines the testing procedures for the **Himatika PDF Worker** service.

## Prerequisites

1. **Environment Variables**:
   Ensure your `.env` file is populated with the following keys:

   ```env
   HIMATIKA_MONGODB_URI=...
   DBNAME=...
   R2_ACCOUNT_ID=...
   R2_ACCESS_KEY_ID=...
   R2_SECRET_ACCESS_KEY=...
   R2_BUCKET_NAME=...
   R2_PUBLIC_DOMAIN=...
   HIMATIKA_JWT_SECRET=...
   ```

2. **Run Server**:

   ```bash
   pip install -r requirements.txt
   python api/index.py
   # Runs on http://localhost:5000
   ```

3. **Authentication**:
   All endpoints (except `GET /`) require a valid JWT token in the `Authorization` header:
   ```
   Authorization: Bearer <jwt_token>
   ```
   The token must contain `{ "service": "himatika-backend" }` and be signed with `HIMATIKA_JWT_SECRET`.

---

## Endpoint Tests

### 1. Generic Excel Export

**Endpoint:** `POST /api/sheet/export`
**Goal:** Verify JSON to Excel conversion.

**Request Payload:**

```json
{
  "title": "Test Export Data",
  "headers": ["Name", "Role", "Active"],
  "data": [
    { "Name": "Alice", "Role": "Admin", "Active": "Yes" },
    { "Name": "Bob", "Role": "User", "Active": "No" }
  ]
}
```

**Expected Result:**

- Status: `200 OK`
- Header: `Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
- Body: Binary Excel file.

---

### 2. Excel Import

**Endpoint:** `POST /api/sheet/import`
**Goal:** Verify Excel file to JSON conversion.

**Request Payload:** Multipart form-data with an `.xlsx` file.

**Expected Result:**

- Status: `200 OK`
- Body: JSON array of parsed rows.

---

### 3. Scan QR from PDF

**Endpoint:** `POST /api/pdf/scan-qr`
**Goal:** Extract QR code data from a PDF document.

**Request Payload:**

```json
{
  "pdf": "https://example.com/document-with-qr.pdf"
}
```

**Expected Result:**

- Status: `200 OK`
- Body: JSON with extracted QR value(s).

---

### 4. Search Text in PDF

**Endpoint:** `POST /api/pdf/search-text`
**Goal:** Find text locations within a PDF document.

**Request Payload:**

```json
{
  "pdf": "https://example.com/document.pdf",
  "searchText": "Signature"
}
```

**Expected Result:**

- Status: `200 OK`
- Body: JSON with text location coordinates.

---

### 5. Create Activiness Letter

**Endpoint:** `POST /api/pdf/activiness-letter`
**Goal:** Generate Surat Keterangan Aktif via ReportLab.

**Request Payload:**

```json
{
  "docNumber": "001/HIMATIKA/I/2024",
  "config": {
    "name": "HIMPUNAN MAHASISWA INFORMATIKA",
    "address": "Jl. Karangdowo No. 9",
    "contact": { "phone": "08123456789", "email": "hima@example.com" }
  },
  "organizer": {
    "period": "2023 - 2024",
    "dailyManagement": [
      { "position": "Ketua Umum", "member": { "fullName": "Ketua Budi", "NIM": "111" } },
      { "position": "Sekretaris Umum", "member": { "fullName": "Sekretaris Ani", "NIM": "222" } }
    ]
  },
  "user": {
    "member": { "fullName": "Anggota Caca", "NIM": "333", "class": "TI-3A" }
  },
  "point": {
    "semester": "V",
    "activities": {
      "agendas": { "committees": 5, "participants": 2 },
      "manualPoints": 10,
      "projects": 1,
      "aspirations": 0
    }
  }
}
```

**Expected Result:**

```json
{
  "success": true,
  "url": "https://.../Surat Keterangan Aktif 333 Semester V.pdf",
  "filename": "...",
  "no": "001/HIMATIKA/I/2024"
}
```

---

### 6. Generate Event Ticket

**Endpoint:** `POST /api/pdf/ticket`
**Goal:** Generate event ticket PDF.

**Request Payload:**

```json
{
  "role": "participant",
  "agenda": {
    "_id": "agenda123",
    "title": "SEMINAR NASIONAL 2024",
    "at": "Auditorium Kampus",
    "date": { "start": "2024-05-20T08:00:00Z" },
    "photos": [{ "image": "https://via.placeholder.com/800" }]
  },
  "participant": {
    "_id": "p123",
    "member": { "fullName": "John Doe" }
  }
}
```

**Expected Result:**

- Status: `200 OK`
- Body: Binary PDF (Ticket-participant-John.pdf)

---

### 7. Certificate Preview

**Endpoint:** `POST /api/pdf/certificate-preview`
**Goal:** Generate a preview of a certificate PDF.

**Request Payload:** Certificate configuration JSON (template items, dimensions, etc.)

**Expected Result:**

- Status: `200 OK`
- Body: Binary PDF preview.

---

### 8. Generate Certificate

**Endpoint:** `POST /api/pdf/certificate`
**Goal:** Generate final certificate PDF for a participant.

**Request Payload:** Certificate configuration + participant data.

**Expected Result:**

- Status: `200 OK`
- Body: JSON with URL to generated certificate.

---

### 9. Signature Overlay (Auto-Detect)

**Endpoint:** `POST /api/sign/process`
**Goal:** Auto-detect text location and overlay QR code signature.

**Request Payload:**

```json
{
  "pdf": "https://pdfobject.com/pdf/sample.pdf",
  "outputBlobPath": "testing/signature-result.pdf",
  "qrValue": "SIGNED-BY-12345",
  "locations": [
    { "page": 1, "x": 50, "y": 50, "width": 100, "height": 100 }
  ]
}
```

**Expected Result:**

```json
{
    "statusCode": 200,
    "statusMessage": "Sign processed successfully",
    "data": "https://.../testing/signature-result.pdf"
}
```

---

### 10. Compress Video

**Endpoint:** `POST /api/media/compress-video`
**Goal:** Compress and transcode video files stored in R2.

**Request Payload:**

```json
{
  "fileKey": "videos/raw/video123.mp4",
  "videoId": "video_db_id",
  "callbackUrl": "http://localhost:3000/api/storage/webhook-media"
}
```

**Expected Result:**

- Status: `200 OK`
- Body: JSON with compressed video URL and status.

---

### 11. Generate QR Code

**Endpoint:** `POST /api/tools/qr`
**Goal:** Generate a QR code image.

**Request Payload:**

```json
{
  "value": "https://himatika.example.com/verify/123"
}
```

**Expected Result:**

- Status: `200 OK`
- Body: PNG image of QR code.

---

### 12. Compress Image

**Endpoint:** `POST /api/tools/compress-image`
**Goal:** Compress and resize an uploaded image.

**Request Payload:** Multipart form-data with image file.

**Expected Result:**

- Status: `200 OK`
- Body: Compressed image binary or URL.

---

### 13. Upload Image

**Endpoint:** `POST /api/tools/upload-image`
**Goal:** Upload an image to R2 storage.

**Request Payload:** Multipart form-data with image file and destination path.

**Expected Result:**

- Status: `200 OK`
- Body: JSON with public URL of uploaded image.
