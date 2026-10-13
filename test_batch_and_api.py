"""
Automated Test Suite for Batch Processing Queue, Bulk ZIP Archive, and Developer REST API v1
Uses standard library urllib.request against the live FastAPI server.
"""

import urllib.request
import json
import io
import zipfile
import fitz

BASE_URL = "http://127.0.0.1:8080"

def run_tests():
    print("=" * 65)
    print("   TESTING BATCH PROCESSING & DEVELOPER REST API (v1)")
    print("=" * 65)

    # 1. Test Developer API Info
    print("\n1. Testing GET /api/v1/info...")
    req = urllib.request.Request(f"{BASE_URL}/api/v1/info")
    with urllib.request.urlopen(req) as res:
        assert res.status == 200
        info = json.loads(res.read().decode())
        print("  -> API Name:", info["api_name"])
        print("  -> Version:", info["version"])
        print("  -> Capabilities:", len(info["capabilities"]), "endpoints")
        assert "merge" in info["capabilities"]
        assert "batch-process" in info["capabilities"]
        print("  -> PASSED: Developer REST API info is active and documented!")

    # 2. Test Multi-File Batch Processing & Bulk ZIP Archive
    print("\n2. Testing POST /api/pdf/batch-process (Compress 3 files into ZIP)...")
    # Generate 3 test PDFs
    def create_dummy_pdf(text):
        d = fitz.open()
        p = d.new_page(width=300, height=300)
        p.insert_text(fitz.Point(30, 50), text, fontsize=12)
        b = d.tobytes()
        d.close()
        return b

    pdf1 = create_dummy_pdf("Batch Doc 1")
    pdf2 = create_dummy_pdf("Batch Doc 2")
    pdf3 = create_dummy_pdf("Batch Doc 3")

    boundary = "----BatchUploadBoundary12345"
    body = io.BytesIO()

    for idx, pdf_b in enumerate([pdf1, pdf2, pdf3], 1):
        body.write(f"--{boundary}\r\n".encode("utf-8"))
        body.write(f'Content-Disposition: form-data; name="files"; filename="doc_{idx}.pdf"\r\n'.encode("utf-8"))
        body.write(b'Content-Type: application/pdf\r\n\r\n')
        body.write(pdf_b)
        body.write(b"\r\n")

    body.write(f"--{boundary}\r\n".encode("utf-8"))
    body.write(b'Content-Disposition: form-data; name="action"\r\n\r\n')
    body.write(b'compress\r\n')

    body.write(f"--{boundary}\r\n".encode("utf-8"))
    body.write(b'Content-Disposition: form-data; name="options"\r\n\r\n')
    body.write(b'{"level":"recommended"}\r\n')

    body.write(f"--{boundary}--\r\n".encode("utf-8"))

    batch_req = urllib.request.Request(
        f"{BASE_URL}/api/pdf/batch-process",
        data=body.getvalue(),
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"}
    )
    with urllib.request.urlopen(batch_req) as batch_res:
        assert batch_res.status == 200
        zip_bytes = batch_res.read()
        print("  -> Received ZIP Archive:", len(zip_bytes), "bytes")
        assert len(zip_bytes) > 0

        # Verify ZIP contains 3 compressed files + manifest
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
            names = z.namelist()
            print("  -> ZIP Contents:", names)
            assert "compress_doc_1.pdf" in names
            assert "compress_doc_2.pdf" in names
            assert "compress_doc_3.pdf" in names
            assert "GOLA_BATCH_MANIFEST.json" in names

            manifest = json.loads(z.read("GOLA_BATCH_MANIFEST.json").decode())
            print("  -> Manifest Total Files:", manifest["total_files"])
            print("  -> Successful Files:", manifest["successful_files"])
            assert manifest["successful_files"] == 3
        print("  -> PASSED: Multi-file batch processing generated organized ZIP with manifest!")

    # 3. Test Developer API v1 /api/v1/compress
    print("\n3. Testing POST /api/v1/compress...")
    c_boundary = "----V1CompressBoundary12345"
    c_body = io.BytesIO()
    c_body.write(f"--{c_boundary}\r\n".encode("utf-8"))
    c_body.write(b'Content-Disposition: form-data; name="file"; filename="api_test.pdf"\r\n')
    c_body.write(b'Content-Type: application/pdf\r\n\r\n')
    c_body.write(pdf1)
    c_body.write(b"\r\n")
    c_body.write(f"--{c_boundary}--\r\n".encode("utf-8"))

    comp_req = urllib.request.Request(
        f"{BASE_URL}/api/v1/compress",
        data=c_body.getvalue(),
        headers={"Content-Type": f"multipart/form-data; boundary={c_boundary}"}
    )
    with urllib.request.urlopen(comp_req) as comp_res:
        assert comp_res.status == 200
        v1_comp_bytes = comp_res.read()
        print("  -> Compressed Output:", len(v1_comp_bytes), "bytes")
        print("  -> Header X-Gola-API-Version:", comp_res.headers.get("X-Gola-API-Version"))
        print("  -> Header X-Processing-Time-Ms:", comp_res.headers.get("X-Processing-Time-Ms"))
        assert comp_res.headers.get("X-Gola-API-Version") == "v1.0"
        print("  -> PASSED: Developer REST API v1 compress endpoint verified!")

    # 4. Test Developer API v1 /api/v1/ocr
    print("\n4. Testing POST /api/v1/ocr...")
    ocr_req = urllib.request.Request(
        f"{BASE_URL}/api/v1/ocr",
        data=c_body.getvalue(),
        headers={"Content-Type": f"multipart/form-data; boundary={c_boundary}"}
    )
    with urllib.request.urlopen(ocr_req) as ocr_res:
        assert ocr_res.status == 200
        ocr_data = json.loads(ocr_res.read().decode())
        print("  -> Status:", ocr_data["status"])
        print("  -> Word Count:", ocr_data["word_count"])
        assert ocr_data["status"] == "success"
        print("  -> PASSED: Developer REST API v1 OCR endpoint verified!")

    print("\n" + "=" * 65)
    print("   ALL BATCH PROCESSING & DEVELOPER REST API TESTS PASSED!")
    print("=" * 65)

if __name__ == "__main__":
    run_tests()
