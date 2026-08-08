import pytest
import requests
import os

BASE_URL = "http://localhost:5000"

def test_home_endpoint():
    response = requests.get(f"{BASE_URL}/")
    assert response.status_code == 200
    assert response.json()["status"] == "success"

def test_auth_rejection():
    # Calling an endpoint without JWT should fail
    response = requests.post(f"{BASE_URL}/api/pdf/overlay-qr", json={"test": "data"})
    assert response.status_code == 401
    assert "Unauthorized" in response.json()["error"]
