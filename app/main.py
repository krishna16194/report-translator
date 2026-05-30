"""FastAPI web app: upload a document, translate it, download a report.

Flow: upload (.txt/.docx/.xlsx) + target language + output format(s)
  -> parse into a content model
  -> translate every text segment
  -> generate .docx and/or .pptx
  -> return a single file, or a .zip when both formats are requested.
"""
from __future__ import annotations

import datetime
import io
import os
import tempfile
import zipfile
from typing import Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles

from . import generators, parsers
from .translator import LANGUAGES, Translator

app = FastAPI(title="Document Translator")

_STATIC_DIR = os.path.join(os.path.dirname(__file__), "..", "static")
_ALLOWED_EXT = {".txt", ".docx", ".xlsx", ".xlsm"}


async def _save_temp(upload: Optional[UploadFile], expected_ext: set[str]) -> Optional[str]:
    """Persist an optional upload to a temp file, validating its extension."""
    if upload is None or not upload.filename:
        return None
    ext = os.path.splitext(upload.filename)[1].lower()
    if ext not in expected_ext:
        raise HTTPException(400, f"Template must be one of {sorted(expected_ext)}, got '{ext}'.")
    with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
        tmp.write(await upload.read())
        return tmp.name


@app.get("/api/languages")
def languages() -> dict[str, str]:
    return LANGUAGES


@app.post("/api/translate")
async def translate(
    file: UploadFile = File(...),
    target_lang: str = Form(...),
    formats: str = Form("docx"),
    docx_template: Optional[UploadFile] = File(None),
    pptx_template: Optional[UploadFile] = File(None),
) -> StreamingResponse:
    ext = os.path.splitext(file.filename or "")[1].lower()
    if ext not in _ALLOWED_EXT:
        raise HTTPException(400, f"Unsupported file type '{ext}'.")
    if target_lang not in LANGUAGES:
        raise HTTPException(400, f"Unsupported language '{target_lang}'.")

    wanted = {f.strip() for f in formats.split(",") if f.strip()}
    if not wanted <= {"docx", "pptx"} or not wanted:
        raise HTTPException(400, "formats must be 'docx', 'pptx', or both.")

    # Optional custom templates (None -> the standard professional design is used).
    docx_tpl = await _save_temp(docx_template, {".docx"})
    pptx_tpl = await _save_temp(pptx_template, {".pptx", ".potx"})

    # Persist upload to a temp file so the parsers can open it by path.
    with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
        tmp.write(await file.read())
        tmp_path = tmp.name

    try:
        parsed = parsers.parse(tmp_path)
    except parsers.UnsupportedFileError as exc:
        raise HTTPException(400, str(exc))
    finally:
        os.unlink(tmp_path)

    # Use the original upload name for the title/download, not the temp file.
    parsed["title"] = os.path.splitext(os.path.basename(file.filename))[0]

    translator = Translator(target_lang=target_lang)
    try:
        translated_blocks = translator.blocks(parsed["blocks"])
    except Exception as exc:  # surfacing translation backend failures cleanly
        _cleanup(docx_tpl, pptx_tpl)
        raise HTTPException(502, f"Translation failed: {exc}")

    title = translator.text(parsed["title"]) or parsed["title"]
    lang_name = LANGUAGES[target_lang]
    base = parsed["title"]
    date_str = datetime.date.today().strftime("%B %d, %Y")

    try:
        outputs: dict[str, bytes] = {}
        if "docx" in wanted:
            outputs[f"{base}_{target_lang}.docx"] = generators.build_docx(
                title, translated_blocks, lang_name, date_str, template_path=docx_tpl
            )
        if "pptx" in wanted:
            outputs[f"{base}_{target_lang}.pptx"] = generators.build_pptx(
                title, translated_blocks, lang_name, date_str, template_path=pptx_tpl
            )
    except Exception as exc:
        raise HTTPException(400, f"Could not build report from template: {exc}")
    finally:
        _cleanup(docx_tpl, pptx_tpl)

    if len(outputs) == 1:
        name, data = next(iter(outputs.items()))
        return _file_response(name, data)

    # Both formats -> zip.
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in outputs.items():
            zf.writestr(name, data)
    zip_buf.seek(0)
    return _file_response(f"{base}_{target_lang}.zip", zip_buf.getvalue())


def _cleanup(*paths: Optional[str]) -> None:
    for path in paths:
        if path and os.path.exists(path):
            try:
                os.unlink(path)
            except OSError:
                pass


def _file_response(filename: str, data: bytes) -> StreamingResponse:
    return StreamingResponse(
        io.BytesIO(data),
        media_type="application/octet-stream",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# Serve the frontend at "/". Mounted last so it doesn't shadow /api routes.
app.mount("/", StaticFiles(directory=_STATIC_DIR, html=True), name="static")
