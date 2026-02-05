import os
import boto3
from dotenv import load_dotenv

load_dotenv()

def get_s3_client():
    return boto3.client(
        's3',
        endpoint_url=f"https://{os.getenv('R2_ACCOUNT_ID')}.r2.cloudflarestorage.com",
        aws_access_key_id=os.getenv('R2_ACCESS_KEY_ID'),
        aws_secret_access_key=os.getenv('R2_SECRET_ACCESS_KEY'),
        region_name="auto" # R2 biasanya auto
    )

def upload_bytes_to_r2(file_bytes, content_type, key):
    s3 = get_s3_client()
    bucket_name = os.getenv('R2_BUCKET_NAME')
    
    s3.put_object(
        Bucket=bucket_name,
        Key=key,
        Body=file_bytes,
        ContentType=content_type
    )
    
    return f"{os.getenv('R2_PUBLIC_DOMAIN')}/{key}"