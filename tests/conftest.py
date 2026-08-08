import pytest
from unittest.mock import patch
from api.index import app as flask_app

@pytest.fixture
def app():
    flask_app.config.update({
        "TESTING": True,
    })
    
    # Mock R2 upload globally
    with patch('api.handlers.pdf.upload_bytes_to_r2', return_value='https://mocked-r2-url.com/file.pdf'):
        yield flask_app

@pytest.fixture
def client(app):
    return app.test_client()

@pytest.fixture
def runner(app):
    return app.test_cli_runner()
