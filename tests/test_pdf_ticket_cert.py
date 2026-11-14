import pytest
import json
from unittest.mock import patch, MagicMock
import io


@patch('api.handlers.pdf.upload_bytes_to_r2', return_value='https://mocked-r2-url.com/ticket.pdf')
def test_ticket_qr_payload_no_nameerror(mock_upload, client, valid_token):
    """Ticket generation with participant._id should not 500 (payload NameError fix)."""
    payload = {
        "role": "participant",
        "agenda": {
            "_id": "agenda123",
            "title": "SEMINAR",
            "at": "Hall",
            "date": {"start": "2024-05-20T08:00:00Z"},
            "photos": [{"image": "https://example.com/img.png"}]
        },
        "participant": {
            "_id": "part123",
            "member": {"fullName": "John Doe"}
        }
    }
    response = client.post(
        '/api/pdf/ticket',
        json=payload,
        headers={"Authorization": f"Bearer {valid_token}"}
    )
    # Expect 200 with PDF binary (send_file), not 500
    assert response.status_code == 200
    assert response.headers['Content-Type'] == 'application/pdf'


@patch('api.handlers.pdf.safe_fetch')
@patch('api.handlers.pdf.upload_bytes_to_r2', return_value='https://mocked-r2-url.com/cert.pdf')
def test_certificate_safe_fetch_used(mock_upload, mock_safe_fetch, client, valid_token):
    """Certificate generation uses safe_fetch (imported at top) - no NameError."""
    # Build a real minimal PDF with reportlab
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    c.drawString(100, 700, "Template")
    c.save()
    buf.seek(0)

    mock_resp = MagicMock()
    mock_resp.status_code = 200
    mock_resp.content = buf.getvalue()
    mock_safe_fetch.return_value = mock_resp

    payload = {
        "templateUrl": "https://example.com/template.pdf",
        "items": [
            {"type": "name", "x": 100, "y": 100, "width": 200, "height": 30, "fontSize": 14, "fontFamily": "Times-Roman", "align": "center", "color": "#000000"}
        ],
        "data": {"name": "Test User"}
    }
    response = client.post(
        '/api/pdf/certificate',
        json=payload,
        headers={"Authorization": f"Bearer {valid_token}"}
    )
    # Should not 500 with NameError: safe_fetch
    assert response.status_code == 200
    data = response.json
    assert data['success'] is True
    assert 'url' in data
    # safe_fetch should have been called with templateUrl
    mock_safe_fetch.assert_called_with("https://example.com/template.pdf")