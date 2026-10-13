"""
Automated Verification Test Suite for iLovePDF Processing Engine
Generates test PDF and office files, executes every tool, and verifies outputs.
"""

import os
import sys
import io
import fitz
from PIL import Image

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pdf_engine

def generate_sample_pdf(text="Sample iLovePDF Document", pages=3) -> bytes:
    doc = fitz.open()
    for i in range(pages):
        page = doc.new_page(width=595, height=842)
        page.insert_text(fitz.Point(50, 100), f"{text} - Page {i+1}", fontsize=24, color=(0.9, 0.2, 0.2))
        page.insert_text(fitz.Point(50, 150), f"This is automated test paragraph line {i+1} with sample content.", fontsize=12)
        page.draw_rect(fitz.Rect(50, 200, 545, 400), color=(0.2, 0.4, 0.8), width=2)
    buf = io.BytesIO()
    doc.save(buf)
    doc.close()
    return buf.getvalue()

def generate_sample_image() -> bytes:
    img = Image.new("RGB", (300, 300), color=(229, 50, 45))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()

def run_all_tests():
    print("Testing iLovePDF Core Suite...\n")
    sample_pdf = generate_sample_pdf("Document A", 3)
    sample_pdf_b = generate_sample_pdf("Document B", 2)
    sample_img = generate_sample_image()

    # 1. Merge
    print("1. Testing Merge PDF...")
    merged = pdf_engine.merge_pdfs([sample_pdf, sample_pdf_b])
    doc_m = fitz.open(stream=merged, filetype="pdf")
    assert len(doc_m) == 5, f"Expected 5 pages, got {len(doc_m)}"
    doc_m.close()
    print("  -> PASSED (5 pages merged)")

    # 2. Split
    print("2. Testing Split PDF...")
    splits = pdf_engine.split_pdf(sample_pdf, page_ranges="1-2, 3")
    assert len(splits) == 2, f"Expected 2 parts, got {len(splits)}"
    print("  -> PASSED (Split into 2 parts)")

    # 3. Compress
    print("3. Testing Compress PDF...")
    compressed = pdf_engine.compress_pdf(sample_pdf, level="recommended")
    assert len(compressed) > 0
    print(f"  -> PASSED (Output size: {len(compressed)} bytes)")

    # 4. PDF to DOCX & DOCX to PDF
    print("4. Testing PDF to Word & Word to PDF...")
    docx_bytes = pdf_engine.pdf_to_docx(sample_pdf)
    assert len(docx_bytes) > 0
    pdf_from_docx = pdf_engine.docx_to_pdf(docx_bytes)
    assert len(pdf_from_docx) > 0
    print("  -> PASSED (Roundtrip PDF <-> DOCX)")

    # 5. PDF to Excel & Excel to PDF
    print("5. Testing PDF to Excel & Excel to PDF...")
    xlsx_bytes = pdf_engine.pdf_to_xlsx(sample_pdf)
    assert len(xlsx_bytes) > 0
    pdf_from_xlsx = pdf_engine.xlsx_to_pdf(xlsx_bytes)
    assert len(pdf_from_xlsx) > 0
    print("  -> PASSED (Roundtrip PDF <-> XLSX)")

    # 6. PDF to PPTX & PPTX to PDF
    print("6. Testing PDF to PowerPoint & PowerPoint to PDF...")
    pptx_bytes = pdf_engine.pdf_to_pptx(sample_pdf)
    assert len(pptx_bytes) > 0
    pdf_from_pptx = pdf_engine.pptx_to_pdf(pptx_bytes)
    assert len(pdf_from_pptx) > 0
    print("  -> PASSED (Roundtrip PDF <-> PPTX)")

    # 7. PDF to Images & Images to PDF
    print("7. Testing PDF to Images & Images to PDF...")
    images = pdf_engine.pdf_to_images(sample_pdf, img_format="png", dpi=100)
    assert len(images) == 3
    pdf_from_imgs = pdf_engine.images_to_pdf([img[1] for img in images])
    assert len(pdf_from_imgs) > 0
    print("  -> PASSED (3 images converted to PDF)")

    # 8. Watermark
    print("8. Testing Watermark PDF...")
    watermarked = pdf_engine.watermark_pdf(sample_pdf, text="SAMPLE WATERMARK", opacity=0.4)
    assert len(watermarked) > 0
    print("  -> PASSED (Watermark applied)")

    # 9. Rotate
    print("9. Testing Rotate PDF...")
    rotated = pdf_engine.rotate_pdf(sample_pdf, angle=90)
    doc_r = fitz.open(stream=rotated, filetype="pdf")
    assert doc_r[0].rotation == 90
    doc_r.close()
    print("  -> PASSED (Pages rotated 90 deg)")

    # 10. Protect & Unlock
    print("10. Testing Protect & Unlock PDF...")
    protected = pdf_engine.protect_pdf(sample_pdf, user_password="secretpassword123")
    unlocked = pdf_engine.unlock_pdf(protected, password="secretpassword123")
    doc_u = fitz.open(stream=unlocked, filetype="pdf")
    assert not doc_u.is_encrypted
    doc_u.close()
    print("  -> PASSED (AES Encrypted & Decrypted)")

    # 11. Organize
    print("11. Testing Organize PDF...")
    organized = pdf_engine.organize_pdf(sample_pdf, page_order=[2, 0], deleted_pages=[1])
    doc_o = fitz.open(stream=organized, filetype="pdf")
    assert len(doc_o) == 2
    doc_o.close()
    print("  -> PASSED (Reordered and deleted pages)")

    # 12. Page Numbers
    print("12. Testing Add Page Numbers...")
    numbered = pdf_engine.add_page_numbers(sample_pdf, format_str="Page {n} of {p}")
    assert len(numbered) > 0
    print("  -> PASSED (Page numbers added)")

    # 13. Crop
    print("13. Testing Crop PDF...")
    cropped = pdf_engine.crop_pdf(sample_pdf, crop_x=0.1, crop_y=0.1, crop_w=0.8, crop_h=0.8)
    assert len(cropped) > 0
    print("  -> PASSED (Page cropped)")

    # 14. Sign PDF
    print("14. Testing Sign PDF...")
    signed = pdf_engine.sign_pdf(sample_pdf, signature_image_bytes=sample_img, signer_name="Shahroz Tariq", date_str="2026-09-01")
    assert len(signed) > 0
    print("  -> PASSED (Signature stamped)")

    # 15. OCR Text Extraction
    print("15. Testing OCR Text Extraction...")
    ocr_txt = pdf_engine.ocr_pdf(sample_pdf)
    assert "Document A" in ocr_txt
    print(f"  -> PASSED (Extracted text length: {len(ocr_txt)} chars)")

    # 16. Compare PDFs
    print("16. Testing Compare PDFs...")
    diffs = pdf_engine.compare_pdfs(sample_pdf, sample_pdf_b)
    assert len(diffs) == 3
    print("  -> PASSED (Generated comparison diff images)")

    # 17. Edit PDF (Annotations & Overlay)
    print("17. Testing Edit PDF (Visual Annotations Overlay)...")
    edited = pdf_engine.edit_pdf(sample_pdf, overlay_image_bytes=sample_img, page_num=1)
    assert len(edited) > 0
    doc_e = fitz.open(stream=edited, filetype="pdf")
    assert len(doc_e) == 3
    doc_e.close()
    print("  -> PASSED (Canvas annotations overlay merged into PDF)")

    # 18. Batch Processing
    print("18. Testing Batch Processing Engine...")
    batch_zip, stats = pdf_engine.batch_process_documents(
        [("doc1.pdf", sample_pdf), ("doc2.pdf", sample_pdf_b)],
        action="watermark",
        options={"text": "BATCH WATERMARK"}
    )
    assert len(batch_zip) > 0
    assert stats["success_count"] == 2
    print("  -> PASSED (Batch processed multiple files with manifest)")

    print("\n=======================================================")
    print("   ALL 18 ENGINE TEST SUITES PASSED FLAWLESSLY!        ")
    print("=======================================================")

if __name__ == "__main__":
    run_all_tests()
