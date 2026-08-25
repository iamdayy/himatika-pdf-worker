import os
from pymongo import MongoClient
from dotenv import load_dotenv

load_dotenv()

_client = None
_db = None


def _ensure_connection():
    """Connect lazily — importing this module no longer hard-crashes the
    whole service when Mongo env vars are absent (nothing in the codebase
    actually queries Mongo from the worker today)."""
    global _client, _db
    if _db is None:
        mongo_uri = os.getenv("HIMATIKA_MONGODB_URI")
        if not mongo_uri:
            raise ValueError(
                "Environment variable HIMATIKA_MONGODB_URI is not set. "
                "Pastikan Anda sudah mengatur env variables di platform deployment."
            )
        db_name = os.getenv("DBNAME")
        if not db_name:
            raise ValueError("Environment variable DBNAME is not set.")
        _client = MongoClient(mongo_uri)
        _db = _client[db_name]
    return _db


def get_members_collection():
    return _ensure_connection()["members"]
