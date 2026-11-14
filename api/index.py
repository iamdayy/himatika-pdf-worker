import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from flask import Flask, jsonify, request
from flask_cors import CORS

from api.handlers.pdf import bp as pdf_bp
from api.handlers.sheet import bp as sheet_bp
from api.handlers.sign import bp as sign_bp
from api.handlers.tools import bp as tools_bp
from api.handlers.media import bp as media_bp

import jwt

app = Flask(__name__)
# Worker is server-to-server, open CORS is dangerous.
# CORS(app)

REQUIRED_ENV = [
    "HIMATIKA_MONGODB_URI", "DBNAME",
    "R2_ACCOUNT_ID", "R2_ACCESS_KEY_ID", "R2_SECRET_ACCESS_KEY",
    "R2_BUCKET_NAME", "R2_PUBLIC_DOMAIN",
]


def get_jwt_secret():
    return os.getenv("HIMATIKA_JWT_SECRET") or os.getenv("JWT_SECRET")  # backend uses jwtSecret


_secret = get_jwt_secret()
if not _secret:
    print("WARNING: no HIMATIKA_JWT_SECRET/JWT_SECRET set — every API call will be rejected 401.")
elif _secret.startswith("$2"):
    print("WARNING: JWT secret looks like a bcrypt hash, not an HS256 secret — token verification will fail. Set the same jwtSecret the backend uses.")


@app.before_request
def verify_jwt():
    # Allow health check routes without JWT
    if request.path in ('/', '/healthz') and request.method == 'GET':
        return

    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({"error": "Unauthorized: Missing or invalid token"}), 401

    token = auth_header.split(' ')[1]
    secret = get_jwt_secret()
    if not secret:
        return jsonify({"error": "Unauthorized: server JWT secret not configured"}), 401

    try:
        payload = jwt.decode(token, secret, algorithms=['HS256'])
        if payload.get('service') != 'himatika-backend':
            return jsonify({"error": "Forbidden: Invalid service token"}), 403
    except jwt.ExpiredSignatureError:
        return jsonify({"error": "Unauthorized: Token expired"}), 401
    except jwt.InvalidTokenError:
        return jsonify({"error": "Unauthorized: Invalid token"}), 401
# Register Blueprints
app.register_blueprint(pdf_bp)
app.register_blueprint(sheet_bp)
app.register_blueprint(sign_bp)
app.register_blueprint(tools_bp)
app.register_blueprint(media_bp)

def _health_check(deep=False):
    problems = [k for k in REQUIRED_ENV if not os.getenv(k)]
    if not get_jwt_secret():
        problems.append("HIMATIKA_JWT_SECRET")
    if deep and not problems:
        # Optional deep R2 check (fast, few ms)
        try:
            from utils.storage import get_s3_client
            get_s3_client().head_bucket(Bucket=os.getenv("R2_BUCKET_NAME"))
        except Exception as e:
            problems.append(f"R2: {e}")
    healthy = not problems
    return jsonify({
        "status": "success" if healthy else "unhealthy",
        "message": "HIMATIKA PDF Worker is running" if healthy else "Dependency checks failed",
        "problems": problems
    }), 200 if healthy else 503


@app.route('/', methods=['GET'])
def home():
    return _health_check(deep=False)


@app.route('/healthz', methods=['GET'])
def healthz():
    return _health_check(deep=True)


if __name__ == '__main__':
    app.run(debug=True, port=5000)
