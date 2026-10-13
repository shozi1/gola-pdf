"""
Gola PDF - Google Cloud Storage & Ephemeral Auto-Purge Engine
Supports Google Cloud Storage buckets, signed upload/download URLs,
and a 2-hour Zero-Data-Retention ephemeral auto-purge lifecycle worker.
"""

import os
import io
import time
import json
import uuid
import shutil
import re
from datetime import datetime, timezone
from typing import Dict, Any, Optional, Tuple

# Try importing official google-cloud-storage if available
try:
    from google.cloud import storage as gcs_lib
    HAS_GCS_SDK = True
except ImportError:
    gcs_lib = None
    HAS_GCS_SDK = False


class CloudStorageEngine:
    """
    Manages Google Cloud Storage uploads and zero-retention ephemeral lifecycle.
    Falls back gracefully to a localized GCS-compatible ephemeral bucket if GCS
    credentials are not yet configured in the environment.
    """

    def __init__(self):
        self.bucket_name = os.getenv("GCS_BUCKET_NAME", "gola-pdf-storage")
        self.project_id = os.getenv("GCS_PROJECT_ID", "gola-pdf-cloud")
        self.credentials_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "")
        self.ttl_hours = float(os.getenv("GCS_RETENTION_HOURS", "2.0"))
        
        # Local ephemeral storage path for zero-retention cache & dev fallback
        self.local_bucket_dir = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "ephemeral_gcs_storage"
        )
        os.makedirs(self.local_bucket_dir, exist_ok=True)

        self.gcs_client = None
        self._init_client()

    def _init_client(self):
        """Initializes GCS client if SDK is installed and credentials exist."""
        if HAS_GCS_SDK:
            try:
                if self.credentials_path and os.path.exists(self.credentials_path):
                    self.gcs_client = gcs_lib.Client.from_service_account_json(self.credentials_path)
                elif "GOOGLE_APPLICATION_CREDENTIALS" in os.environ:
                    self.gcs_client = gcs_lib.Client()
            except Exception as e:
                print(f"[CloudStorageEngine] GCS SDK init note: {e}. Running in ephemeral cloud mode.")
                self.gcs_client = None
        else:
            self.gcs_client = None

    @property
    def is_gcs_live(self) -> bool:
        """Returns True if connected to an active Google Cloud Storage bucket."""
        return self.gcs_client is not None

    def upload_file(
        self,
        file_bytes: bytes,
        filename: str,
        content_type: str = "application/pdf",
        custom_metadata: Optional[Dict[str, str]] = None,
        file_id: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Uploads a file to Google Cloud Storage with a 2-Hour TTL.
        """
        raw_id = file_id or uuid.uuid4().hex[:12]
        file_id = re.sub(r'[^a-zA-Z0-9_-]', '', raw_id) or uuid.uuid4().hex[:12]
        clean_name = os.path.basename(filename) or "document.pdf"
        blob_name = f"uploads/{file_id}/{clean_name}"
        now_ts = time.time()
        expires_ts = now_ts + (self.ttl_hours * 3600)

        meta = {
            "blob_name": blob_name,
            "filename": clean_name,
            "file_id": file_id,
            "content_type": content_type,
            "size": len(file_bytes),
            "created_at": now_ts,
            "created_iso": datetime.now(timezone.utc).isoformat(),
            "expires_at": expires_ts,
            "ttl_hours": self.ttl_hours,
            "custom": custom_metadata or {}
        }

        # 1. Try upload to live Google Cloud Storage bucket if connected
        if self.is_gcs_live:
            try:
                bucket = self.gcs_client.bucket(self.bucket_name)
                blob = bucket.blob(blob_name)
                blob.metadata = {
                    "ttl_hours": str(self.ttl_hours),
                    "created_at": str(now_ts),
                    "expires_at": str(expires_ts),
                    "original_filename": clean_name
                }
                blob.upload_from_string(file_bytes, content_type=content_type)
                
                return {
                    "status": "success",
                    "storage_backend": "Google Cloud Storage",
                    "bucket": self.bucket_name,
                    "blob_name": blob_name,
                    "filename": clean_name,
                    "file_id": file_id,
                    "size": len(file_bytes),
                    "download_url": f"/api/cloud/gcs/file/{file_id}/{clean_name}",
                    "expires_in_hours": self.ttl_hours,
                    "retention_policy": "2-Hour Zero Retention Auto-Purge"
                }
            except Exception as e:
                print(f"[CloudStorageEngine] GCS upload exception: {e}. Using ephemeral bucket.")

        # 2. Ephemeral Storage Fallback (Zero retention on disk with 2h auto-purge)
        target_dir = os.path.join(self.local_bucket_dir, file_id)
        os.makedirs(target_dir, exist_ok=True)
        file_path = os.path.join(target_dir, clean_name)
        meta_path = os.path.join(target_dir, "meta.json")

        with open(file_path, "wb") as f:
            f.write(file_bytes)

        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump(meta, f, indent=2)

        return {
            "status": "success",
            "storage_backend": "Ephemeral Cloud Bucket (Local Fallback)",
            "bucket": self.bucket_name,
            "blob_name": blob_name,
            "filename": clean_name,
            "file_id": file_id,
            "size": len(file_bytes),
            "download_url": f"/api/cloud/gcs/file/{file_id}/{clean_name}",
            "expires_in_hours": self.ttl_hours,
            "retention_policy": "2-Hour Zero Retention Auto-Purge"
        }

    def download_file(self, file_id: str, filename: str) -> Optional[Tuple[bytes, str]]:
        """
        Retrieves file bytes and content-type by file_id and filename.
        """
        clean_id = re.sub(r'[^a-zA-Z0-9_-]', '', file_id)
        clean_name = os.path.basename(filename)
        if not clean_id or not clean_name:
            return None

        # 1. Check live GCS bucket
        if self.is_gcs_live:
            try:
                bucket = self.gcs_client.bucket(self.bucket_name)
                blob_name = f"uploads/{clean_id}/{clean_name}"
                blob = bucket.blob(blob_name)
                if blob.exists():
                    data = blob.download_as_bytes()
                    return data, blob.content_type or "application/pdf"
            except Exception as e:
                print(f"[CloudStorageEngine] GCS download exception: {e}")

        # 2. Check local ephemeral storage (with path traversal protection)
        safe_base = os.path.abspath(self.local_bucket_dir)
        target_dir = os.path.abspath(os.path.join(self.local_bucket_dir, clean_id))
        file_path = os.path.abspath(os.path.join(target_dir, clean_name))
        meta_path = os.path.join(target_dir, "meta.json")

        if not target_dir.startswith(safe_base) or not file_path.startswith(safe_base):
            return None

        if os.path.exists(file_path):
            content_type = "application/pdf"
            if os.path.exists(meta_path):
                try:
                    with open(meta_path, "r", encoding="utf-8") as f:
                        m = json.load(f)
                        content_type = m.get("content_type", "application/pdf")
                except Exception:
                    pass
            with open(file_path, "rb") as f:
                return f.read(), content_type

        return None

    def generate_signed_upload_url(self, filename: str, content_type: str = "application/pdf") -> Dict[str, Any]:
        """
        Generates a pre-signed URL for direct client-to-GCS upload.
        """
        file_id = uuid.uuid4().hex[:12]
        clean_name = os.path.basename(filename) or "document.pdf"
        blob_name = f"uploads/{file_id}/{clean_name}"

        if self.is_gcs_live:
            try:
                bucket = self.gcs_client.bucket(self.bucket_name)
                blob = bucket.blob(blob_name)
                url = blob.generate_signed_url(
                    version="v4",
                    expiration=7200,  # 2 hours
                    method="PUT",
                    content_type=content_type
                )
                return {
                    "signed_url": url,
                    "blob_name": blob_name,
                    "file_id": file_id,
                    "filename": clean_name,
                    "method": "PUT",
                    "headers": {"Content-Type": content_type},
                    "storage_backend": "Google Cloud Storage"
                }
            except Exception as e:
                print(f"[CloudStorageEngine] Signed URL error: {e}")

        # Fallback to direct app endpoint
        return {
            "signed_url": f"/api/cloud/gcs/upload-direct?file_id={file_id}&filename={clean_name}",
            "blob_name": blob_name,
            "file_id": file_id,
            "filename": clean_name,
            "method": "POST",
            "headers": {"Content-Type": content_type},
            "storage_backend": "Ephemeral Cloud Bucket (Local Fallback)"
        }

    def purge_expired_objects(self, max_age_hours: Optional[float] = None) -> Dict[str, Any]:
        """
        Deletes all stored files and metadata older than the TTL (default 2 hours).
        Guarantees strict Zero-Data-Retention.
        """
        cutoff_hours = max_age_hours if max_age_hours is not None else self.ttl_hours
        now = time.time()
        purged_files = 0
        purged_bytes = 0
        active_files = 0

        # 1. Clean local ephemeral bucket
        if os.path.exists(self.local_bucket_dir):
            for item in os.listdir(self.local_bucket_dir):
                item_path = os.path.join(self.local_bucket_dir, item)
                if os.path.isdir(item_path):
                    meta_path = os.path.join(item_path, "meta.json")
                    is_expired = False
                    
                    if os.path.exists(meta_path):
                        try:
                            with open(meta_path, "r", encoding="utf-8") as f:
                                meta = json.load(f)
                                created = meta.get("created_at", 0)
                                if (now - created) >= (cutoff_hours * 3600):
                                    is_expired = True
                        except Exception:
                            is_expired = True
                    else:
                        # Check directory mtime
                        dir_mtime = os.path.getmtime(item_path)
                        if (now - dir_mtime) >= (cutoff_hours * 3600):
                            is_expired = True

                    if is_expired:
                        for root, _, files in os.walk(item_path):
                            for f in files:
                                fp = os.path.join(root, f)
                                purged_bytes += os.path.getsize(fp)
                                purged_files += 1
                        shutil.rmtree(item_path, ignore_errors=True)
                    else:
                        active_files += 1

        # 2. Clean live GCS bucket if connected
        if self.is_gcs_live:
            try:
                bucket = self.gcs_client.bucket(self.bucket_name)
                blobs = bucket.list_blobs(prefix="uploads/")
                for blob in blobs:
                    created_time = blob.time_created.timestamp() if getattr(blob, 'time_created', None) else now
                    blob_age_seconds = now - created_time
                    if blob_age_seconds >= (cutoff_hours * 3600):
                        purged_bytes += blob.size or 0
                        purged_files += 1
                        blob.delete()
                    else:
                        active_files += 1
            except Exception as e:
                print(f"[CloudStorageEngine] GCS purge exception: {e}")

        return {
            "purged_files": purged_files,
            "purged_bytes": purged_bytes,
            "active_files": active_files,
            "retention_cutoff_hours": cutoff_hours,
            "executed_at": datetime.now(timezone.utc).isoformat()
        }

    def get_storage_status(self) -> Dict[str, Any]:
        """
        Returns real-time status and telemetry of the cloud storage engine.
        """
        total_size = 0
        file_count = 0

        if os.path.exists(self.local_bucket_dir):
            for root, _, files in os.walk(self.local_bucket_dir):
                for f in files:
                    if f != "meta.json":
                        fp = os.path.join(root, f)
                        total_size += os.path.getsize(fp)
                        file_count += 1

        return {
            "backend": "Google Cloud Storage" if self.is_gcs_live else "Ephemeral Cloud Bucket (Local Fallback)",
            "gcs_sdk_available": HAS_GCS_SDK,
            "is_connected_to_gcs": self.is_gcs_live,
            "bucket_name": self.bucket_name,
            "project_id": self.project_id,
            "ttl_retention_hours": self.ttl_hours,
            "zero_data_retention_policy": "Automated 2-Hour Auto-Purge Lifecycle",
            "active_objects": file_count,
            "storage_used_bytes": total_size,
            "storage_used_formatted": f"{total_size / (1024 * 1024):.2f} MB"
        }


# Global singleton instance
cloud_storage = CloudStorageEngine()
