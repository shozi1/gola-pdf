"""
Quick HTTP API Verifier for live iLovePDF server
"""

import urllib.request
import io
import fitz

def test_api():
    print("Testing live server at http://127.0.0.1:8080 ...")
    # Health check
    h_res = urllib.request.urlopen("http://127.0.0.1:8080/api/health")
    print("Health check response:", h_res.read().decode())

    # Create dummy PDF
    doc = fitz.open()
    doc.new_page(width=400, height=400)
    pdf_bytes = doc.tobytes()

    boundary = "----TestBoundary12345"
    body = io.BytesIO()
    
    # File 1
    body.write(f"--{boundary}\r\n".encode("utf-8"))
    body.write(b'Content-Disposition: form-data; name="files"; filename="file1.pdf"\r\n')
    body.write(b'Content-Type: application/pdf\r\n\r\n')
    body.write(pdf_bytes)
    body.write(b"\r\n")

    # File 2
    body.write(f"--{boundary}\r\n".encode("utf-8"))
    body.write(b'Content-Disposition: form-data; name="files"; filename="file2.pdf"\r\n')
    body.write(b'Content-Type: application/pdf\r\n\r\n')
    body.write(pdf_bytes)
    body.write(b"\r\n")

    body.write(f"--{boundary}--\r\n".encode("utf-8"))

    req = urllib.request.Request(
        "http://127.0.0.1:8080/api/pdf/merge",
        data=body.getvalue(),
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"}
    )

    resp = urllib.request.urlopen(req)
    result_data = resp.read()
    print(f"API Merge endpoint test: HTTP {resp.status} OK (Received {len(result_data)} bytes)")
    merged_doc = fitz.open(stream=result_data, filetype="pdf")
    print(f"Merged PDF page count: {len(merged_doc)} pages")
    merged_doc.close()
    print("LIVE API ENDPOINT TEST PASSED SUCCESSFULLY!")

if __name__ == "__main__":
    test_api()
