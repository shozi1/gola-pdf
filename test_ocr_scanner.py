"""
Test Suite for OCR and Pre-Download Document Security Scanner
"""

import io
import os
import sys
import fitz

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pdf_engine


def create_test_scanned_pdf() -> bytes:
    """Create a sample scanned-style PDF."""
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    page.insert_text(fitz.Point(50, 80), "CONFIDENTIAL INVOICE #89214", fontsize=16, color=(0.1, 0.1, 0.1))
    page.insert_text(fitz.Point(50, 120), "Client: Acme Corporation", fontsize=11)
    page.insert_text(fitz.Point(50, 140), "Amount Due: $14,920.00 USD", fontsize=11)
    page.insert_text(fitz.Point(50, 160), "Due Date: September 15, 2026", fontsize=11)
    
    # Draw simple stamp image
    page.draw_rect(fitz.Rect(50, 200, 250, 260), color=(0.8, 0.2, 0.2), width=1.5)
    page.insert_text(fitz.Point(70, 235), "PAID & VERIFIED", fontsize=14, color=(0.8, 0.2, 0.2))

    buf = io.BytesIO()
    doc.save(buf)
    doc.close()
    return buf.getvalue()


def run_scanner_and_ocr_tests():
    print("=======================================================")
    print("   TESTING OCR & PRE-DOWNLOAD DOCUMENT SCANNER         ")
    print("=======================================================\n")

    sample_pdf = create_test_scanned_pdf()
    print(f"Created test PDF ({len(sample_pdf)} bytes).")

    # 1. TEST DOCUMENT SECURITY & AUDIT SCANNER
    print("\n1. Testing [scan_and_inspect_document] Security Scanner...")
    scan_report = pdf_engine.scan_and_inspect_document(sample_pdf)
    print("  -> Status:", scan_report["status"])
    print("  -> Clean:", scan_report["clean"])
    print("  -> Security Badge:", scan_report["security_badge"])
    print("  -> Total Words:", scan_report["total_words"])
    print("  -> Pages:", scan_report["pages"])
    assert scan_report["clean"] is True
    assert scan_report["pages"] == 1
    assert scan_report["total_words"] > 5
    print("  -> PASSED: Document security scanner verified integrity & metadata successfully!")

    # 2. TEST OCR TEXT EXTRACTION
    print("\n2. Testing [ocr_pdf] Text Extraction Engine...")
    extracted_text = pdf_engine.ocr_pdf(sample_pdf)
    print("  -> Extracted text preview:\n", extracted_text[:120], "...")
    assert "CONFIDENTIAL INVOICE" in extracted_text
    assert "Acme Corporation" in extracted_text
    print("  -> PASSED: OCR extracted text from document successfully!")

    # 3. TEST SEARCHABLE OCR PDF GENERATION
    print("\n3. Testing [create_searchable_ocr_pdf] Searchable PDF Builder...")
    searchable_pdf_bytes, summary = pdf_engine.create_searchable_ocr_pdf(sample_pdf)
    assert len(searchable_pdf_bytes) > 2000
    print("  -> Generated Searchable PDF Size:", len(searchable_pdf_bytes), "bytes")
    print("  -> Total Words Embedded:", summary["total_words"])
    
    # Inspect generated searchable PDF
    check_doc = fitz.open(stream=searchable_pdf_bytes, filetype="pdf")
    p_text = check_doc[0].get_text("text")
    assert len(p_text) > 10
    check_doc.close()
    print("  -> PASSED: Searchable OCR PDF generated with selectable/searchable text layer!")

    print("\n=======================================================")
    print("   ALL OCR & DOCUMENT SCANNER TESTS PASSED!            ")
    print("=======================================================")


if __name__ == "__main__":
    run_scanner_and_ocr_tests()
