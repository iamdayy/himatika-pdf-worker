import os
from pymongo import MongoClient
from dotenv import load_dotenv

load_dotenv()

MONGO_URI = os.getenv("HIMATIKA_MONGODB_URI")
DB_NAME = os.getenv("DBNAME")

client = MongoClient(MONGO_URI)
db = client[DB_NAME]

def get_members_collection():
    return db["members"]