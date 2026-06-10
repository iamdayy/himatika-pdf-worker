import os
from pymongo import MongoClient
from dotenv import load_dotenv

load_dotenv()

MONGO_URI = os.getenv("HIMATIKA_MONGODB_URI")
if not MONGO_URI:
    raise ValueError("Environment variable HIMATIKA_MONGODB_URI is not set. Pastikan Anda sudah mengatur env variables di Coolify.")

DB_NAME = os.getenv("DBNAME")
if not DB_NAME:
    raise ValueError("Environment variable DBNAME is not set. Pastikan Anda sudah mengatur env variables di Coolify.")

client = MongoClient(MONGO_URI)
db = client[DB_NAME]

def get_members_collection():
    return db["members"]