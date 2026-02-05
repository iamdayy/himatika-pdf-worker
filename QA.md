# QA & Testing Guide

This document outlines the testing procedures for the **Himatika PDF Worker** service.

## Prerequisites

1. **Environment Variables**:
   Ensure your `.env` file is populated with the following keys:

   ```env
   HIMATIKA_MONGODB_URI=...
   DBNAME=...
   R2_ENDPOINT_URL=...
   R2_ACCESS_KEY_ID=...
   R2_SECRET_ACCESS_KEY=...
   R2_BUCKET_NAME=...
   R2_Public_Dev_Url=...
   ```

2. **Run Server**:

   ```bash
   pip install -r requirements.txt
   python api/index.py
   # Runs on http://localhost:5000
   ```

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

### 2. Member Export

**Endpoint:** `POST /api/member/sheet/export`  
**Goal:** Verify database export of members.

**Request Payload:**

```json
{
  "data": ["20200123001", "20200123002"] 
}
```

*(Note: Array of NIMs. If empty `[]`, attempts to export all matching query)*

**Expected Result:**

- Status: `200 OK`
- Body: Binary Excel file containing member data (NIM, Name, etc).

---

### 3. PDF Overlay (QR Code)

**Endpoint:** `POST /api/pdf/overlay-qr`  
**Goal:** Verify QR code stamping on existing PDF.

**Request Payload:**

```json
{
  "pdf": "https://pdfobject.com/pdf/sample.pdf",
  "outputBlobPath": "testing/overlay-result.pdf",
  "qrValue": "VALID-123",
  "locations": [
    { "page": 1, "x": 50, "y": 50, "width": 100, "height": 100 }
  ]
}
```

**Expected Result:**

```json
{
  "success": true,
  "url": "https://.../testing/overlay-result.pdf"
}
```

---

### 4. Create Activiness Letter

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

### 5. Generate Event Ticket

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

### 6. Signature Overlay (Auto-Detect)

**Endpoint:** `POST /api/sign/process`
**Goal:** Auto-detect text location and overlay QR code.

**Request Payload:**

```json
{
  "pdf": "https://pdfobject.com/pdf/sample.pdf",
  "outputBlobPath": "testing/signature-result.pdf",
  "qrValue": "SIGNED-BY-12345",
  "searchText": "/12345signature/"
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
