"""
Gola PDF Full Suite FastAPI Application
Exposes REST endpoints for all 26 PDF manipulation and document conversion tools,
and serves the complete interactive single-page application.
"""

import io
import json
import zipfile
from typing import List, Optional
from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import Response, JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from fastapi.middleware.gzip import GZipMiddleware

import pdf_engine
import ai_engine

app = FastAPI(title="Gola PDF Suite API", version="2.5.0")

# Enable high-speed GZIP compression for all responses > 500 bytes
app.add_middleware(GZipMiddleware, minimum_size=500)

# Enable CORS for cross-origin integration
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Enforce Security Headers & 100MB Upload Protection (Protects $14 VPS Memory)
@app.middleware("http")
async def security_and_limits_middleware(request, call_next):
    # Enforce maximum upload size limit of 100 MB to protect VPS memory
    content_length = request.headers.get("content-length")
    if content_length and int(content_length) > 104_857_600:
        return JSONResponse(
            status_code=413,
            content={"detail": "File size exceeds the 100 MB free processing limit."}
        )
    
    response = await call_next(request)
    # Add modern web security headers
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "SAMEORIGIN"
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    return response


@app.get("/api/health")
async def health_check():
    """Health check endpoint."""
    return {"status": "ok", "app": "Gola PDF", "version": "2.5.0"}


# 1. MERGE PDF
@app.post("/api/pdf/merge")
async def api_merge_pdfs(files: List[UploadFile] = File(...)):
    try:
        bytes_list = [await f.read() for f in files]
        if not bytes_list:
            raise HTTPException(status_code=400, detail="No PDF files uploaded.")
        merged_bytes = pdf_engine.merge_pdfs(bytes_list)
        return Response(
            content=merged_bytes,
            media_type="application/pdf",
            headers={"Content-Disposition": "attachment; filename=merged_document.pdf"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 2. SPLIT PDF
@app.post("/api/pdf/split")
async def api_split_pdf(file: UploadFile = File(...), ranges: str = Form("all")):
    try:
        content = await file.read()
        results = pdf_engine.split_pdf(content, page_ranges=ranges)
        if len(results) == 1:
            name, pdf_data = results[0]
            return Response(
                content=pdf_data,
                media_type="application/pdf",
                headers={"Content-Disposition": f"attachment; filename={name}"}
            )
        else:
            # Return as zip archive
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
                for name, data in results:
                    z.writestr(name, data)
            return Response(
                content=zip_buf.getvalue(),
                media_type="application/zip",
                headers={"Content-Disposition": "attachment; filename=split_pages.zip"}
            )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 3. COMPRESS PDF
@app.post("/api/pdf/compress")
async def api_compress_pdf(file: UploadFile = File(...), level: str = Form("recommended")):
    try:
        content = await file.read()
        compressed = pdf_engine.compress_pdf(content, level=level)
        return Response(
            content=compressed,
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename=compressed_{file.filename or 'document.pdf'}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 4. PDF TO WORD (DOCX)
@app.post("/api/pdf/to-word")
async def api_pdf_to_word(file: UploadFile = File(...)):
    try:
        content = await file.read()
        docx_bytes = pdf_engine.pdf_to_docx(content)
        base_name = (file.filename or "document").rsplit(".", 1)[0]
        return Response(
            content=docx_bytes,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f"attachment; filename={base_name}.docx"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 5. WORD (DOCX) TO PDF
@app.post("/api/pdf/word-to-pdf")
async def api_word_to_pdf(file: UploadFile = File(...)):
    try:
        content = await file.read()
        pdf_bytes = pdf_engine.docx_to_pdf(content)
        base_name = (file.filename or "document").rsplit(".", 1)[0]
        return Response(
            content=pdf_bytes,
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename={base_name}.pdf"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 6. PDF TO EXCEL (XLSX)
@app.post("/api/pdf/to-excel")
async def api_pdf_to_excel(file: UploadFile = File(...)):
    try:
        content = await file.read()
        xlsx_bytes = pdf_engine.pdf_to_xlsx(content)
        base_name = (file.filename or "document").rsplit(".", 1)[0]
        return Response(
            content=xlsx_bytes,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={base_name}.xlsx"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 7. EXCEL (XLSX) TO PDF
@app.post("/api/pdf/excel-to-pdf")
async def api_excel_to_pdf(file: UploadFile = File(...)):
    try:
        content = await file.read()
        pdf_bytes = pdf_engine.xlsx_to_pdf(content)
        base_name = (file.filename or "document").rsplit(".", 1)[0]
        return Response(
            content=pdf_bytes,
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename={base_name}.pdf"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 8. PDF TO POWERPOINT (PPTX)
@app.post("/api/pdf/to-powerpoint")
async def api_pdf_to_powerpoint(file: UploadFile = File(...)):
    try:
        content = await file.read()
        pptx_bytes = pdf_engine.pdf_to_pptx(content)
        base_name = (file.filename or "document").rsplit(".", 1)[0]
        return Response(
            content=pptx_bytes,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": f"attachment; filename={base_name}.pptx"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 9. POWERPOINT (PPTX) TO PDF
@app.post("/api/pdf/powerpoint-to-pdf")
async def api_powerpoint_to_pdf(file: UploadFile = File(...)):
    try:
        content = await file.read()
        pdf_bytes = pdf_engine.pptx_to_pdf(content)
        base_name = (file.filename or "document").rsplit(".", 1)[0]
        return Response(
            content=pdf_bytes,
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename={base_name}.pdf"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 10. PDF TO JPG / PNG
@app.post("/api/pdf/to-jpg")
async def api_pdf_to_jpg(file: UploadFile = File(...), format: str = Form("jpg"), dpi: int = Form(150)):
    try:
        content = await file.read()
        images = pdf_engine.pdf_to_images(content, img_format=format, dpi=dpi)
        if len(images) == 1:
            name, img_data = images[0]
            m_type = "image/jpeg" if "jpg" in format else "image/png"
            return Response(
                content=img_data,
                media_type=m_type,
                headers={"Content-Disposition": f"attachment; filename={name}"}
            )
        else:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
                for name, img_data in images:
                    z.writestr(name, img_data)
            return Response(
                content=zip_buf.getvalue(),
                media_type="application/zip",
                headers={"Content-Disposition": "attachment; filename=pdf_images.zip"}
            )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 11. JPG / IMAGES TO PDF
@app.post("/api/pdf/jpg-to-pdf")
async def api_jpg_to_pdf(
    files: List[UploadFile] = File(...),
    orientation: str = Form("portrait"),
    margin: int = Form(0),
    page_size: str = Form("a4")
):
    try:
        img_list = [await f.read() for f in files]
        pdf_bytes = pdf_engine.images_to_pdf(img_list, orientation=orientation, margin=margin, page_size_name=page_size)
        return Response(
            content=pdf_bytes,
            media_type="application/pdf",
            headers={"Content-Disposition": "attachment; filename=images_converted.pdf"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 12. WATERMARK PDF
@app.post("/api/pdf/watermark")
async def api_watermark_pdf(
    file: UploadFile = File(...),
    text: str = Form("CONFIDENTIAL"),
    opacity: float = Form(0.3),
    position: str = Form("center"),
    rotation: int = Form(45),
    color: str = Form("#e5322d"),
    font_size: int = Form(48)
):
    try:
        content = await file.read()
        wm_pdf = pdf_engine.watermark_pdf(
            content, text=text, opacity=opacity, position=position, rotation=rotation, color_hex=color, font_size=font_size
        )
        return Response(
            content=wm_pdf,
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename=watermarked_{file.filename or 'document.pdf'}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 13. ROTATE PDF
@app.post("/api/pdf/rotate")
async def api_rotate_pdf(file: UploadFile = File(...), angle: int = Form(90)):
    try:
        content = await file.read()
        rotated = pdf_engine.rotate_pdf(content, angle=angle)
        return Response(
            content=rotated,
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename=rotated_{file.filename or 'document.pdf'}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 14. PROTECT PDF
@app.post("/api/pdf/protect")
async def api_protect_pdf(file: UploadFile = File(...), password: str = Form(...)):
    try:
        content = await file.read()
        protected = pdf_engine.protect_pdf(content, user_password=password)
        return Response(
            content=protected,
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename=protected_{file.filename or 'document.pdf'}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 15. UNLOCK PDF
@app.post("/api/pdf/unlock")
async def api_unlock_pdf(file: UploadFile = File(...), password: str = Form("")):
    try:
        content = await file.read()
        unlocked = pdf_engine.unlock_pdf(content, password=password)
        return Response(
            content=unlocked,
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename=unlocked_{file.filename or 'document.pdf'}"}
        )
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


# 16. ORGANIZE PDF
@app.post("/api/pdf/organize")
async def api_organize_pdf(
    file: UploadFile = File(...),
    page_order: str = Form("[]"),
    deleted_pages: str = Form("[]"),
    rotations: str = Form("{}")
):
    try:
        content = await file.read()
        order = json.loads(page_order)
        deleted = json.loads(deleted_pages)
        rot_map = json.loads(rotations)
        organized = pdf_engine.organize_pdf(content, page_order=order, deleted_pages=deleted, rotations=rot_map)
        return Response(
            content=organized,
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename=organized_{file.filename or 'document.pdf'}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 17. ADD PAGE NUMBERS
@app.post("/api/pdf/page-numbers")
async def api_page_numbers(
    file: UploadFile = File(...),
    format_str: str = Form("Page {n} of {p}"),
    position: str = Form("bottom-center"),
    margin: int = Form(30),
    start_page: int = Form(1),
    font_size: int = Form(10)
):
    try:
        content = await file.read()
        numbered = pdf_engine.add_page_numbers(
            content, format_str=format_str, position=position, margin=margin, start_page=start_page, font_size=font_size
        )
        return Response(
            content=numbered,
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename=numbered_{file.filename or 'document.pdf'}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 18. CROP PDF
@app.post("/api/pdf/crop")
async def api_crop_pdf(
    file: UploadFile = File(...),
    x: float = Form(0.0),
    y: float = Form(0.0),
    width: float = Form(1.0),
    height: float = Form(1.0)
):
    try:
        content = await file.read()
        cropped = pdf_engine.crop_pdf(content, crop_x=x, crop_y=y, crop_w=width, crop_h=height)
        return Response(
            content=cropped,
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename=cropped_{file.filename or 'document.pdf'}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 19. REDACT PDF
@app.post("/api/pdf/redact")
async def api_redact_pdf(file: UploadFile = File(...), redactions: str = Form("[]")):
    try:
        content = await file.read()
        redacts = json.loads(redactions)
        redacted = pdf_engine.redact_pdf(content, redactions=redacts)
        return Response(
            content=redacted,
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename=redacted_{file.filename or 'document.pdf'}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 20. REPAIR PDF
@app.post("/api/pdf/repair")
async def api_repair_pdf(file: UploadFile = File(...)):
    try:
        content = await file.read()
        repaired = pdf_engine.repair_pdf(content)
        return Response(
            content=repaired,
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename=repaired_{file.filename or 'document.pdf'}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# EDIT PDF (CANVAS ANNOTATIONS & OVERLAYS)
@app.post("/api/pdf/edit")
async def api_edit_pdf(file: UploadFile = File(...), overlay: Optional[UploadFile] = File(None), page: int = Form(1)):
    try:
        content = await file.read()
        overlay_bytes = await overlay.read() if overlay else None
        edited_bytes = pdf_engine.edit_pdf(content, overlay_image_bytes=overlay_bytes, page_num=page)
        return Response(
            content=edited_bytes,
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename=edited_{file.filename or 'document.pdf'}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 21. OCR / EXTRACT TEXT
@app.post("/api/pdf/ocr")
async def api_ocr_pdf(file: UploadFile = File(...)):
    try:
        content = await file.read()
        extracted_text = pdf_engine.ocr_pdf(content)
        return JSONResponse({"text": extracted_text, "filename": file.filename})
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 22. SIGN PDF
@app.post("/api/pdf/sign")
async def api_sign_pdf(
    file: UploadFile = File(...),
    signature: UploadFile = File(...),
    page: int = Form(1),
    x: float = Form(100.0),
    y: float = Form(500.0),
    width: float = Form(180.0),
    height: float = Form(60.0),
    signer_name: str = Form(""),
    date_str: str = Form("")
):
    try:
        pdf_bytes = await file.read()
        sig_bytes = await signature.read()
        signed = pdf_engine.sign_pdf(
            pdf_bytes,
            signature_image_bytes=sig_bytes,
            page_num=page,
            x=x,
            y=y,
            width=width,
            height=height,
            signer_name=signer_name,
            date_str=date_str
        )
        return Response(
            content=signed,
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename=signed_{file.filename or 'document.pdf'}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 23. COMPARE PDFS
@app.post("/api/pdf/compare")
async def api_compare_pdfs(file1: UploadFile = File(...), file2: UploadFile = File(...)):
    try:
        b1 = await file1.read()
        b2 = await file2.read()
        diffs = pdf_engine.compare_pdfs(b1, b2)
        
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
            for name, img_data in diffs:
                z.writestr(name, img_data)
        return Response(
            content=zip_buf.getvalue(),
            media_type="application/zip",
            headers={"Content-Disposition": "attachment; filename=pdf_comparison.zip"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 24. DOCUMENT SECURITY & PRE-DOWNLOAD SCANNER
@app.post("/api/pdf/scan")
async def api_scan_document(file: UploadFile = File(...)):
    try:
        content = await file.read()
        scan_results = pdf_engine.scan_and_inspect_document(content)
        scan_results["filename"] = file.filename or "document.pdf"
        return JSONResponse(scan_results)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 25. OCR SEARCHABLE PDF GENERATOR
@app.post("/api/pdf/ocr-searchable")
async def api_ocr_searchable(file: UploadFile = File(...)):
    try:
        content = await file.read()
        searchable_pdf, summary = pdf_engine.create_searchable_ocr_pdf(content)
        base_name = (file.filename or "document").rsplit(".", 1)[0]
        return Response(
            content=searchable_pdf,
            media_type="application/pdf",
            headers={
                "Content-Disposition": f"attachment; filename=searchable_ocr_{base_name}.pdf",
                "X-OCR-Words": str(summary["total_words"]),
                "X-OCR-Pages": str(summary["total_pages"])
            }
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 26. AI DOCUMENT SUMMARIZER
@app.post("/api/ai/summarize")
async def api_ai_summarize(file: UploadFile = File(...), bullets: int = Form(5)):
    try:
        content = await file.read()
        summary = ai_engine.summarize_document(content, max_bullets=bullets)
        return JSONResponse(summary)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 27. AI CHAT WITH PDF (SEMANTIC QA)
@app.post("/api/ai/chat")
async def api_ai_chat(file: UploadFile = File(...), query: str = Form(...)):
    try:
        content = await file.read()
        response_data = ai_engine.chat_with_document(content, query=query)
        return JSONResponse(response_data)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 28. SMART AUTO-REDACT PII (CREDIT CARDS, SSN, EMAILS, PHONES)
@app.post("/api/pdf/auto-redact")
async def api_auto_redact_pii(file: UploadFile = File(...), types: str = Form("card,ssn,email,phone,iban")):
    try:
        content = await file.read()
        selected_types = [t.strip() for t in types.split(",") if t.strip()]
        redacted_bytes, stats = pdf_engine.auto_redact_pii(content, pii_types=selected_types)
        base_name = (file.filename or "document").rsplit(".", 1)[0]
        return Response(
            content=redacted_bytes,
            media_type="application/pdf",
            headers={
                "Content-Disposition": f"attachment; filename=redacted_pii_{base_name}.pdf",
                "X-Total-Redactions": str(stats["total_redactions"])
            }
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# 29. CONVERT TO PDF/A (LEGAL & ARCHIVAL COMPLIANCE)
@app.post("/api/pdf/to-pdfa")
async def api_convert_to_pdfa(file: UploadFile = File(...)):
    try:
        content = await file.read()
        pdfa_bytes = pdf_engine.convert_to_pdfa(content)
        base_name = (file.filename or "document").rsplit(".", 1)[0]
        return Response(
            content=pdfa_bytes,
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename=pdfa_{base_name}.pdf"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# =========================================================================
# 30. GOOGLE CLOUD STORAGE & EPHEMERAL ZERO-RETENTION ENDPOINTS
# =========================================================================
from cloud_storage import cloud_storage
import threading
import time

@app.get("/api/cloud/gcs/status")
async def api_gcs_status():
    """Returns status, bucket configuration, and zero-retention statistics."""
    return JSONResponse(cloud_storage.get_storage_status())


@app.post("/api/cloud/gcs/upload")
async def api_gcs_upload(file: UploadFile = File(...)):
    """Uploads a document to Google Cloud Storage with a 2-Hour TTL."""
    try:
        content = await file.read()
        result = cloud_storage.upload_file(
            file_bytes=content,
            filename=file.filename or "document.pdf",
            content_type=file.content_type or "application/pdf"
        )
        return JSONResponse(result)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/cloud/gcs/signed-upload-url")
async def api_gcs_signed_upload_url(filename: str = "document.pdf", content_type: str = "application/pdf"):
    """Generates pre-signed URL for direct browser-to-bucket upload."""
    try:
        data = cloud_storage.generate_signed_upload_url(filename=filename, content_type=content_type)
        return JSONResponse(data)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/cloud/gcs/upload-direct")
async def api_gcs_upload_direct(file_id: str, filename: str, file: UploadFile = File(...)):
    """Direct upload endpoint for fallback ephemeral storage."""
    try:
        content = await file.read()
        res = cloud_storage.upload_file(content, filename=filename, content_type=file.content_type or "application/pdf", file_id=file_id)
        return JSONResponse(res)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/cloud/gcs/file/{file_id}/{filename}")
async def api_gcs_download_file(file_id: str, filename: str):
    """Downloads/streams a file stored in Google Cloud Storage."""
    file_data = cloud_storage.download_file(file_id, filename)
    if not file_data:
        raise HTTPException(status_code=404, detail="Document not found or expired under the 2-Hour Zero Retention policy.")
    data_bytes, ctype = file_data
    return Response(
        content=data_bytes,
        media_type=ctype,
        headers={"Content-Disposition": f"inline; filename={filename}"}
    )


@app.post("/api/cloud/gcs/purge")
async def api_gcs_purge(max_age_hours: Optional[float] = None):
    """Triggers zero-retention auto-purge of expired files (default 2 hours)."""
    try:
        result = cloud_storage.purge_expired_objects(max_age_hours=max_age_hours)
        return JSONResponse(result)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# =========================================================================
# 31. BATCH PROCESSING QUEUE & BULK ZIP ARCHIVE
# =========================================================================
@app.post("/api/pdf/batch-process")
async def api_batch_process(
    files: List[UploadFile] = File(...),
    action: str = Form("compress"),
    options: str = Form("{}")
):
    """Process 10-50+ files in parallel and package into a single ZIP archive."""
    try:
        opts = json.loads(options) if options else {}
        file_tuples = []
        for f in files:
            b = await f.read()
            file_tuples.append((f.filename or "doc.pdf", b))

        zip_bytes, stats = pdf_engine.batch_process_documents(file_tuples, action=action, options=opts)
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


# Mount Developer REST API v1
import api_v1
app.include_router(api_v1.router)


# Background worker for 15-minute periodic auto-purge
def _run_periodic_purge():
    while True:
        try:
            time.sleep(900)  # Every 15 minutes
            cloud_storage.purge_expired_objects()
        except Exception as e:
            print(f"[AutoPurgeWorker] Error: {e}")

purge_thread = threading.Thread(target=_run_periodic_purge, daemon=True)
purge_thread.start()


# Static directory reference
import os
static_dir = os.path.join(os.path.dirname(__file__), "static")
if not os.path.exists(static_dir):
    os.makedirs(static_dir, exist_ok=True)


# SEO & ROBOTS.TXT
@app.get("/robots.txt", response_class=Response)
async def get_robots_txt():
    robots_content = (
        "User-agent: *\n"
        "Allow: /\n"
        "Disallow: /api/\n"
        "Disallow: /download/\n\n"
        "Sitemap: https://golapdf.com/sitemap.xml\n"
    )
    return Response(content=robots_content, media_type="text/plain")


# DYNAMIC SITEMAP.XML GENERATOR
@app.get("/sitemap.xml", response_class=Response)
async def get_sitemap_xml():
    tools = [
        "merge-pdf", "split-pdf", "compress-pdf", "pdf-to-word",
        "pdf-to-powerpoint", "pdf-to-excel", "word-to-pdf", "powerpoint-to-pdf",
        "excel-to-pdf", "edit-pdf", "pdf-to-jpg", "jpg-to-pdf",
        "page-numbers", "watermark-pdf", "rotate-pdf", "unlock-pdf",
        "protect-pdf", "crop-pdf", "sign-pdf", "ocr-pdf",
        "redact-pdf", "organize-pdf", "repair-pdf", "chat-pdf",
        "ai-summarize", "auto-redact", "pdf-to-pdfa", "compare-pdf"
    ]
    xml_lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">',
        '  <url>',
        '    <loc>https://golapdf.com/</loc>',
        '    <changefreq>daily</changefreq>',
        '    <priority>1.0</priority>',
        '  </url>'
    ]
    for t in tools:
        xml_lines.extend([
            '  <url>',
            f'    <loc>https://golapdf.com/{t}</loc>',
            '    <changefreq>weekly</changefreq>',
            '    <priority>0.9</priority>',
            '  </url>'
        ])
    for legal in ["privacy", "terms", "about", "contact"]:
        xml_lines.extend([
            '  <url>',
            f'    <loc>https://golapdf.com/{legal}</loc>',
            '    <changefreq>monthly</changefreq>',
            '    <priority>0.5</priority>',
            '  </url>'
        ])
    xml_lines.append('</urlset>')
    return Response(content="\n".join(xml_lines), media_type="application/xml")


# CLEAN TOOL URL ROUTING (/{tool_id})
VALID_TOOL_SLUGS = {
    "merge-pdf", "split-pdf", "compress-pdf", "pdf-to-word",
    "pdf-to-powerpoint", "pdf-to-excel", "word-to-pdf", "powerpoint-to-pdf",
    "excel-to-pdf", "edit-pdf", "pdf-to-jpg", "jpg-to-pdf",
    "page-numbers", "watermark-pdf", "rotate-pdf", "unlock-pdf",
    "protect-pdf", "crop-pdf", "sign-pdf", "ocr-pdf",
    "redact-pdf", "organize-pdf", "repair-pdf", "chat-pdf",
    "ai-summarize", "auto-redact", "pdf-to-pdfa", "compare-pdf",
    "privacy", "terms", "about", "contact"
}

@app.get("/{tool_id}")
async def get_tool_route(tool_id: str):
    if tool_id in VALID_TOOL_SLUGS:
        index_file = os.path.join(static_dir, "index.html")
        return FileResponse(index_file)
    safe_static_dir = os.path.abspath(static_dir)
    candidate_file = os.path.abspath(os.path.join(safe_static_dir, tool_id))
    if candidate_file.startswith(safe_static_dir) and os.path.exists(candidate_file) and os.path.isfile(candidate_file):
        return FileResponse(candidate_file)
    raise HTTPException(status_code=404, detail="Page not found")


# Mount static files directory
app.mount("/", StaticFiles(directory=static_dir, html=True), name="static")

