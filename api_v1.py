"""
Gola PDF - Developer Public REST API (v1)
Exposes standardized endpoints for developers, third-party applications,
and webhooks to integrate PDF tools, OCR, and AI summarization.
"""

import time
import io
import json
from typing import List, Optional
from fastapi import APIRouter, File, UploadFile, Form, HTTPException, Header, Response
from fastapi.responses import JSONResponse

import pdf_engine
import ai_engine

router = APIRouter(prefix="/api/v1", tags=["Developer REST API v1"])


@router.get("/info")
async def api_info():
    """Returns Developer API status, capabilities, and rate-limit details."""
    return {
        "api_name": "Gola PDF Developer REST API",
        "version": "v1.0.0",
        "documentation": "https://golapdf.com/api/v1/docs",
        "capabilities": [
            "merge", "split", "compress", "ocr", "summarize",
            "auto-redact", "to-pdfa", "batch-process", "watermark", "protect"
        ],
        "rate_limits": {
            "free_tier": "60 requests / minute",
            "max_file_size_mb": 100,
            "max_batch_files": 50
        },
        "privacy": "Zero Data Retention Guarantee (Automated 2-Hour Ephemeral Purge)"
    }


@router.post("/merge")
async def v1_merge_pdfs(files: List[UploadFile] = File(...)):
    """Merge 2 or more PDF files in sequential order."""
    t0 = time.time()
    if len(files) < 2:
        raise HTTPException(status_code=400, detail="Minimum of 2 PDF files required to merge.")
    
    file_bytes_list = []
    for f in files:
        file_bytes_list.append(await f.read())

    try:
        merged = pdf_engine.merge_pdfs(file_bytes_list)
        latency = round((time.time() - t0) * 1000, 1)
        return Response(
            content=merged,
            media_type="application/pdf",
            headers={
                "Content-Disposition": "attachment; filename=gola_api_merged.pdf",
                "X-Gola-API-Version": "v1.0",
                "X-Processing-Time-Ms": str(latency),
                "X-Files-Merged": str(len(files))
            }
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/compress")
async def v1_compress_pdf(file: UploadFile = File(...), level: str = Form("recommended")):
    """Compress a PDF with selectable level: 'extreme', 'recommended', 'less'."""
    t0 = time.time()
    try:
        content = await file.read()
        compressed = pdf_engine.compress_pdf(content, level=level)
        orig_len = len(content)
        comp_len = len(compressed)
        savings = round(((orig_len - comp_len) / max(1, orig_len)) * 100, 1)
        latency = round((time.time() - t0) * 1000, 1)

        return Response(
            content=compressed,
            media_type="application/pdf",
            headers={
                "Content-Disposition": f"attachment; filename=compressed_{file.filename or 'document.pdf'}",
                "X-Gola-API-Version": "v1.0",
                "X-Original-Size": str(orig_len),
                "X-Compressed-Size": str(comp_len),
                "X-Savings-Percent": f"{savings}%",
                "X-Processing-Time-Ms": str(latency)
            }
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/ocr")
async def v1_ocr_extract(file: UploadFile = File(...)):
    """Extract searchable text stream from document via OCR."""
    t0 = time.time()
    try:
        content = await file.read()
        text = pdf_engine.ocr_pdf(content)
        latency = round((time.time() - t0) * 1000, 1)
        return {
            "status": "success",
            "filename": file.filename,
            "word_count": len(text.split()),
            "character_count": len(text),
            "text": text,
            "latency_ms": latency
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/summarize")
async def v1_ai_summarize(file: UploadFile = File(...), bullets: int = Form(5)):
    """AI Document Summarizer: returns executive summary, bullet points, and key metrics."""
    t0 = time.time()
    try:
        content = await file.read()
        summary = ai_engine.summarize_document(content, max_bullets=bullets)
        summary["latency_ms"] = round((time.time() - t0) * 1000, 1)
        return JSONResponse(summary)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/auto-redact")
async def v1_auto_redact_pii(file: UploadFile = File(...), types: str = Form("card,ssn,email,phone,iban")):
    """Sanitize and permanently blackout credit cards (Luhn-checked), SSNs, emails, and phones."""
    t0 = time.time()
    try:
        content = await file.read()
        selected_types = [t.strip() for t in types.split(",") if t.strip()]
        redacted_bytes, stats = pdf_engine.auto_redact_pii(content, pii_types=selected_types)
        latency = round((time.time() - t0) * 1000, 1)

        return Response(
            content=redacted_bytes,
            media_type="application/pdf",
            headers={
                "Content-Disposition": f"attachment; filename=sanitized_{file.filename or 'document.pdf'}",
                "X-Total-Redactions": str(stats["total_redactions"]),
                "X-Cards-Redacted": str(stats["cards_redacted"]),
                "X-SSN-Redacted": str(stats["ssn_redacted"]),
                "X-Processing-Time-Ms": str(latency)
            }
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/pdfa")
async def v1_convert_to_pdfa(file: UploadFile = File(...)):
    """Convert standard PDF to ISO 19005 compliant PDF/A archival format."""
    try:
        content = await file.read()
        pdfa_bytes = pdf_engine.convert_to_pdfa(content)
        return Response(
            content=pdfa_bytes,
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename=pdfa_{file.filename or 'document.pdf'}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/batch")
async def v1_batch_process(files: List[UploadFile] = File(...), action: str = Form("compress")):
    """Batch process multiple documents and package all outputs into a single ZIP archive."""
    try:
        file_tuples = []
        for f in files:
            b = await f.read()
            file_tuples.append((f.filename or "doc.pdf", b))

        zip_bytes, stats = pdf_engine.batch_process_documents(file_tuples, action=action)
        return Response(
            content=zip_bytes,
            media_type="application/zip",
            headers={
                "Content-Disposition": f"attachment; filename=gola_batch_{action}.zip",
                "X-Files-Processed": str(stats["total_files"]),
                "X-Success-Count": str(stats["success_count"])
            }
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/split")
async def v1_split_pdf(file: UploadFile = File(...), ranges: str = Form("all")):
    """Split a PDF into individual pages or page ranges."""
    t0 = time.time()
    try:
        content = await file.read()
        results = pdf_engine.split_pdf(content, page_ranges=ranges)
        latency = round((time.time() - t0) * 1000, 1)
        if len(results) == 1:
            name, pdf_data = results[0]
            return Response(
                content=pdf_data,
                media_type="application/pdf",
                headers={
                    "Content-Disposition": f"attachment; filename={name}",
                    "X-Gola-API-Version": "v1.0",
                    "X-Processing-Time-Ms": str(latency)
                }
            )
        else:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
                for name, data in results:
                    z.writestr(name, data)
            return Response(
                content=zip_buf.getvalue(),
                media_type="application/zip",
                headers={
                    "Content-Disposition": "attachment; filename=split_pages.zip",
                    "X-Gola-API-Version": "v1.0",
                    "X-Parts-Count": str(len(results)),
                    "X-Processing-Time-Ms": str(latency)
                }
            )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/watermark")
async def v1_watermark_pdf(
    file: UploadFile = File(...),
    text: str = Form("CONFIDENTIAL"),
    opacity: float = Form(0.3),
    position: str = Form("center")
):
    """Stamp a watermark across all pages of a PDF."""
    t0 = time.time()
    try:
        content = await file.read()
        wm_pdf = pdf_engine.watermark_pdf(content, text=text, opacity=opacity, position=position)
        latency = round((time.time() - t0) * 1000, 1)
        return Response(
            content=wm_pdf,
            media_type="application/pdf",
            headers={
                "Content-Disposition": f"attachment; filename=watermarked_{file.filename or 'document.pdf'}",
                "X-Gola-API-Version": "v1.0",
                "X-Processing-Time-Ms": str(latency)
            }
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/protect")
async def v1_protect_pdf(file: UploadFile = File(...), password: str = Form(...)):
    """Encrypt and password-protect a PDF with AES-256 encryption."""
    t0 = time.time()
    try:
        content = await file.read()
        protected = pdf_engine.protect_pdf(content, user_password=password)
        latency = round((time.time() - t0) * 1000, 1)
        return Response(
            content=protected,
            media_type="application/pdf",
            headers={
                "Content-Disposition": f"attachment; filename=protected_{file.filename or 'document.pdf'}",
                "X-Gola-API-Version": "v1.0",
                "X-Processing-Time-Ms": str(latency)
            }
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
