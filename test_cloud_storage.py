"""
Test Suite for Google Cloud Storage & Ephemeral Auto-Purge Engine
Uses standard library urllib.request against the live FastAPI server.
"""

import urllib.request
import json
import io
import fitz

BASE_URL = "http://127.0.0.1:8080"

def run_tests():
    print("=" * 60)
    print("   TESTING GOOGLE CLOUD STORAGE & EPHEMERAL AUTO-PURGE")
    print("=" * 60)

    # 1. Test Storage Status Endpoint
    print("\n1. Testing GET /api/cloud/gcs/status...")
    req = urllib.request.Request(f"{BASE_URL}/api/cloud/gcs/status")
    with urllib.request.urlopen(req) as res:
        assert res.status == 200, f"Failed: {res.status}"
        status = json.loads(res.read().decode())
        print("  -> Backend:", status["backend"])
        print("  -> Bucket:", status["bucket_name"])
        print("  -> Policy:", status["zero_data_retention_policy"])
        print("  -> Retention:", status["ttl_retention_hours"], "hours")
        assert "bucket_name" in status
        assert status["ttl_retention_hours"] == 2.0
        print("  -> PASSED: GCS status returned valid metadata!")

    # 2. Test File Upload to GCS / Ephemeral Bucket
    print("\n2. Testing POST /api/cloud/gcs/upload...")
    doc = fitz.open()
    doc.new_page(width=300, height=300)
    dummy_pdf = doc.tobytes()
    doc.close()

    boundary = "----GcsUploadBoundary98765"
    body = io.BytesIO()
    body.write(f"--{boundary}\r\n".encode("utf-8"))
    body.write(b'Content-Disposition: form-data; name="file"; filename="gcs_test_contract.pdf"\r\n')
    body.write(b'Content-Type: application/pdf\r\n\r\n')
    body.write(dummy_pdf)
    body.write(b"\r\n")
    body.write(f"--{boundary}--\r\n".encode("utf-8"))

    upload_req = urllib.request.Request(
        f"{BASE_URL}/api/cloud/gcs/upload",
        data=body.getvalue(),
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"}
    )
    with urllib.request.urlopen(upload_req) as upload_res:
        assert upload_res.status == 200, f"Failed upload: {upload_res.status}"
        data = json.loads(upload_res.read().decode())
        print("  -> Status:", data["status"])
        print("  -> File ID:", data["file_id"])
        print("  -> Download URL:", data["download_url"])
        print("  -> Retention:", data["retention_policy"])
        file_id = data["file_id"]
        filename = data["filename"]
        assert data["status"] == "success"
        print("  -> PASSED: Document successfully uploaded to Cloud Storage!")

    # 3. Test Download from GCS / Ephemeral Bucket
    print(f"\n3. Testing GET /api/cloud/gcs/file/{file_id}/{filename}...")
    dl_req = urllib.request.Request(f"{BASE_URL}/api/cloud/gcs/file/{file_id}/{filename}")
    with urllib.request.urlopen(dl_req) as dl_res:
        assert dl_res.status == 200, f"Failed download: {dl_res.status}"
        content = dl_res.read()
        assert content == dummy_pdf
        print("  -> Content Size:", len(content), "bytes")
        print("  -> PASSED: Document successfully retrieved and verified from storage!")

    # 4. Test Pre-Signed URL Generation
    print("\n4. Testing GET /api/cloud/gcs/signed-upload-url...")
    signed_req = urllib.request.Request(f"{BASE_URL}/api/cloud/gcs/signed-upload-url?filename=direct_upload.pdf")
    with urllib.request.urlopen(signed_req) as signed_res:
        assert signed_res.status == 200
        s_data = json.loads(signed_res.read().decode())
        print("  -> Signed URL:", s_data["signed_url"][:60], "...")
        print("  -> Method:", s_data["method"])
        print("  -> Backend:", s_data["storage_backend"])
        assert "signed_url" in s_data
        print("  -> PASSED: Pre-signed upload URL successfully generated!")

    # 5. Test 2-Hour Ephemeral Auto-Purge Worker
    print("\n5. Testing POST /api/cloud/gcs/purge (Forced TTL Expire)...")
    purge_req = urllib.request.Request(f"{BASE_URL}/api/cloud/gcs/purge?max_age_hours=0", data=b"")
    with urllib.request.urlopen(purge_req) as purge_res:
        assert purge_res.status == 200
        p_data = json.loads(purge_res.read().decode())
        print("  -> Purged Files:", p_data["purged_files"])
        print("  -> Purged Bytes:", p_data["purged_bytes"])
        print("  -> Executed At:", p_data["executed_at"])
        assert p_data["purged_files"] >= 1
        print("  -> PASSED: 2-Hour Zero Retention auto-purge cleanly wiped expired files!")

    print("\n" + "=" * 60)
    print("   ALL GOOGLE CLOUD STORAGE & AUTO-PURGE TESTS PASSED!")
    print("=" * 60)

if __name__ == "__main__":
    run_tests()
