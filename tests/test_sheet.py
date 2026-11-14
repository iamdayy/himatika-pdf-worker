import pytest
import io
from openpyxl import Workbook


def test_sheet_export_title_sanitized(client, valid_token):
    """Export with title containing invalid chars should not 500."""
    payload = {
        "title": "Laporan: Kegiatan/2026",
        "data": [{"name": "A", "val": 1}],
    }
    response = client.post(
        '/api/sheet/export',
        json=payload,
        headers={"Authorization": f"Bearer {valid_token}"}
    )
    assert response.status_code == 200
    # file should be returned as attachment
    assert response.headers['Content-Type'] == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    # filename should be sanitized
    assert 'Laporan- Kegiatan-2026' in response.headers.get('Content-Disposition', '')


def test_sheet_export_non_dict_items_rejected(client, valid_token):
    """data items must be dicts; non-dict should 400 not 500."""
    payload = {"data": [[1, 2], "not a dict"]}
    response = client.post(
        '/api/sheet/export',
        json=payload,
        headers={"Authorization": f"Bearer {valid_token}"}
    )
    assert response.status_code == 400
    assert "is not an object" in response.json.get('error', '')


def test_sheet_import_non_xlsx_rejected(client, valid_token):
    """Non-.xlsx upload should 400 not 500."""
    data = b"not an xlsx"
    response = client.post(
        '/api/sheet/import',
        data={'file': (io.BytesIO(data), 'test.csv')},
        headers={"Authorization": f"Bearer {valid_token}"},
        content_type='multipart/form-data'
    )
    assert response.status_code == 400
    assert "xlsx" in response.json.get('error', '').lower()


def test_sheet_import_corrupt_xlsx_rejected(client, valid_token):
    """Corrupt .xlsx should 400 not 500."""
    data = b"PK\x03\x04corrupt"
    response = client.post(
        '/api/sheet/import',
        data={'file': (io.BytesIO(data), 'test.xlsx')},
        headers={"Authorization": f"Bearer {valid_token}"},
        content_type='multipart/form-data'
    )
    assert response.status_code == 400
    assert "invalid" in response.json.get('error', '').lower() or "corrupt" in response.json.get('error', '').lower()


def test_sheet_export_include_id(client, valid_token):
    """Export should include _id in flattened output when present."""
    payload = {
        "title": "Test ID",
        "data": [{"_id": "abc123", "name": "X"}],
    }
    response = client.post(
        '/api/sheet/export',
        json=payload,
        headers={"Authorization": f"Bearer {valid_token}"}
    )
    assert response.status_code == 200
    # Verify the returned Excel contains the _id column
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(response.data))
    ws = wb.active
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    assert '_id' in headers