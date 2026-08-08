import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from flask import Flask, jsonify
from flask_cors import CORS

from api.handlers.pdf import bp as pdf_bp
from api.handlers.sheet import bp as sheet_bp
from api.handlers.sign import bp as sign_bp
from api.handlers.tools import bp as tools_bp
from api.handlers.media import bp as media_bp

app = Flask(__name__)
CORS(app)

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

# Note: JWT middleware will be added in another branch.

if __name__ == '__main__':
    app.run(debug=True, port=5000)
