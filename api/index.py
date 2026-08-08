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

@app.before_request
def verify_jwt():
    # Allow home route for health checks
    if request.path == '/' and request.method == 'GET':
        return
        
    auth_header = request.headers.get('Authorization')
    if not auth_header or not auth_header.startswith('Bearer '):
        return jsonify({"error": "Unauthorized: Missing or invalid token"}), 401
        
    token = auth_header.split(' ')[1]
    secret = os.getenv("HIMATIKA_JWT_SECRET") or os.getenv("JWT_SECRET") # backend uses jwtSecret
    
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

@app.route('/', methods=['GET'])
def home():
    return jsonify({
        "status": "success",
        "message": "HIMATIKA PDF Worker is running"
    })


if __name__ == '__main__':
    app.run(debug=True, port=5000)
