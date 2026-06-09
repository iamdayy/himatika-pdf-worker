from waitress import serve
from api.index import app
import os

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    print(f"Serving on port {port}...")
    serve(app, host="0.0.0.0", port=port)
