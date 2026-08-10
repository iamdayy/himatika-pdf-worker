import pytest
from unittest.mock import patch
import jwt
from api.index import app as flask_app

import os

@pytest.fixture
def app():
    flask_app.config.update({
        "TESTING": True,
    })
    os.environ['JWT_SECRET'] = 'test-secret'
    
    # Mock R2 upload globally
    with patch('api.handlers.pdf.upload_bytes_to_r2', return_value='https://mocked-r2-url.com/file.pdf'):
        yield flask_app

@pytest.fixture
def client(app):
    return app.test_client()

@pytest.fixture
def valid_token():
    return jwt.encode({"service": "himatika-backend"}, "test-secret", algorithm="HS256")

@pytest.fixture
def runner(app):
    return app.test_cli_runner()
