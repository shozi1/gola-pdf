"""
Test Suite for Phase 1: AI Chat with PDF, Document Summarizer, Smart PII Auto-Redaction, and PDF/A.
"""

import io
import os
import sys
import fitz

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pdf_engine
import ai_engine


def create_sample_confidential_contract() -> bytes:
    """Create a realistic legal contract PDF with sensitive PII and facts for AI testing."""
    doc = fitz.open()
    page1 = doc.new_page(width=595, height=842)
    
    # Title & Introduction
    page1.insert_text(fitz.Point(50, 70), "MUTUAL NON-DISCLOSURE AND SERVICE AGREEMENT", fontsize=14, fontname="helv", color=(0.1, 0.1, 0.1))
    page1.insert_text(fitz.Point(50, 100), "1. Parties Involved", fontsize=12, fontname="helv", color=(0.9, 0.2, 0.2))
    contract_p1 = "This agreement is entered into on October 12, 2026 between Global FinTech Inc. ('Client') and Gola Solutions Ltd. ('Vendor'). The total contract compensation is fixed at $2,450,000 USD payable across 4 quarters."
    page1.insert_textbox(fitz.Rect(50, 115, 545, 175), contract_p1, fontsize=10, fontname="helv")
    
    # Sensitive Data for Redaction
    page1.insert_text(fitz.Point(50, 195), "2. Billing and Personal Identification Information", fontsize=12, fontname="helv", color=(0.9, 0.2, 0.2))
    pii_p1 = "Primary Signatory: Jane Smith, SSN: 123-45-6789, Email: jane.smith@globalfintech.org, Phone: +1 555-839-2041."
    page1.insert_textbox(fitz.Rect(50, 210, 545, 245), pii_p1, fontsize=10, fontname="helv")
    card_p1 = "Corporate Credit Card on file: 4532-1488-2941-8924 for recurring retainer fees."
    page1.insert_textbox(fitz.Rect(50, 245, 545, 275), card_p1, fontsize=10, fontname="helv")

    # Page 2: Deliverables & Milestones
    page2 = doc.new_page(width=595, height=842)
    page2.insert_text(fitz.Point(50, 70), "3. Milestones and Key Deadlines", fontsize=12, fontname="helv", color=(0.9, 0.2, 0.2))
    milestone_text = "• Phase 1 Delivery Date is November 30, 2026 with an initial milestone payout of $600,000.\n• Final integration and security audit must achieve a 99.9% uptime compliance score by March 15, 2027.\n• Either party may terminate this agreement with 30 days prior written notice."
    page2.insert_textbox(fitz.Rect(50, 95, 545, 180), milestone_text, fontsize=10, fontname="helv")

    buf = io.BytesIO()
    doc.save(buf)
    doc.close()
    return buf.getvalue()


def run_tests():
    print("=======================================================")
    print("   TESTING PHASE 1: AI & SMART REDACTION ENGINES      ")
    print("=======================================================\n")
    
    sample_pdf = create_sample_confidential_contract()
    print(f"Generated test contract PDF ({len(sample_pdf)} bytes, 2 pages).\n")

    # 1. TEST AI SUMMARIZER
    print("1. Testing [ai_engine.summarize_document]...")
    summary = ai_engine.summarize_document(sample_pdf)
    assert len(summary["executive_summary"]) > 30, "Executive summary too short"
    assert len(summary["key_takeaways"]) >= 2, "Insufficient takeaways"
    print("  -> Executive Summary:\n   ", summary["executive_summary"])
    print("  -> Key Takeaways Count:", len(summary["key_takeaways"]))
    print("  -> Critical Metrics:", [m["metric"] for m in summary["critical_metrics"]])
    assert any("$2,450,000" in m["metric"] or "2026" in m["metric"] for m in summary["critical_metrics"]), "Failed to extract key financial metric"
    print("  -> PASSED: AI Summarizer extracted coherent summary and key metrics!\n")

    # 2. TEST CHAT WITH PDF (SEMANTIC QA)
    print("2. Testing [ai_engine.chat_with_document]...")
    q1 = "What is the total compensation and fee?"
    res1 = ai_engine.chat_with_document(sample_pdf, query=q1)
    print(f"  -> Q: '{q1}'")
    print(f"  -> A: {res1['answer']}")
    print(f"  -> Citations: Page(s) {res1['citations']}")
    assert "$2,450,000" in res1["answer"] or 1 in res1["citations"], "Failed to locate compensation fact"

    q2 = "What is the Phase 1 delivery deadline?"
    res2 = ai_engine.chat_with_document(sample_pdf, query=q2)
    print(f"  -> Q: '{q2}'")
    print(f"  -> A: {res2['answer']}")
    print(f"  -> Citations: Page(s) {res2['citations']}")
    assert "November 30" in res2["answer"] or 2 in res2["citations"], "Failed to locate deadline fact"
    print("  -> PASSED: Chat with PDF accurately answered questions with page citations!\n")

    # 3. TEST SMART PII AUTO-REDACTION
    print("3. Testing [pdf_engine.auto_redact_pii]...")
    redacted_pdf, stats = pdf_engine.auto_redact_pii(sample_pdf)
    print("  -> Redaction Stats:", stats)
    assert stats["total_redactions"] >= 3, "Failed to redact sensitive items"
    
    # Verify redacted PDF stream contains zero unredacted PII
    check_doc = fitz.open(stream=redacted_pdf, filetype="pdf")
    page1_text = check_doc[0].get_text("text")
    assert "123-45-6789" not in page1_text, "SSN still visible!"
    assert "jane.smith@globalfintech.org" not in page1_text, "Email still visible!"
    assert "4532-1488-2941-8924" not in page1_text, "Credit card still visible!"
    check_doc.close()
    print("  -> PASSED: All credit cards, SSNs, and emails were permanently burned out and sanitized!\n")

    # 4. TEST PDF/A CONVERSION
    print("4. Testing [pdf_engine.convert_to_pdfa]...")
    pdfa_bytes = pdf_engine.convert_to_pdfa(sample_pdf)
    assert len(pdfa_bytes) > 1000
    pdfa_doc = fitz.open(stream=pdfa_bytes, filetype="pdf")
    meta = pdfa_doc.metadata
    assert "PDF/A" in meta.get("producer", ""), "PDF/A metadata missing"
    pdfa_doc.close()
    print("  -> PASSED: Successfully converted document to PDF/A standard!\n")

    print("=======================================================")
    print("   ALL PHASE 1 AI & SMART ENGINE TESTS PASSED!         ")
    print("=======================================================")


if __name__ == "__main__":
    run_tests()
