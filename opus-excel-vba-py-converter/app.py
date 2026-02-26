"""
Excel VBA to Python Converter – FastAPI Application
"""
from __future__ import annotations

import io
import logging
import os
import re
import shutil
import zipfile
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Annotated

import uvicorn
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from werkzeug.utils import secure_filename

from config import get_config
from data_exporter import DataExporter
from formula_extractor import FormulaExtractor
from llm_converter import VBAToPythonConverter
from offline_converter import OfflineConverter
from vba_extractor import VBAExtractor
from workbook_analyzer import WorkbookAnalyzer

# ---------------------------------------------------------------------------
# Configuration & Logging
# ---------------------------------------------------------------------------
cfg = get_config()

logging.basicConfig(
    level=getattr(logging, cfg.LOG_LEVEL, logging.INFO),
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger(__name__)


@asynccontextmanager
async def lifespan(_app: FastAPI):
    """Startup / shutdown lifecycle."""
    cfg.ensure_upload_folder()
    logger.info("Upload folder ready: %s", cfg.UPLOAD_FOLDER)
    yield


# Shared error messages
_ERR_NO_FILE = "No file selected"
_TEMPLATE_PATH = Path("templates/index.html")

# Reusable annotated type for file uploads
FileUpload = Annotated[UploadFile, File(...)]

# Documented HTTP error responses for OpenAPI
_UPLOAD_RESPONSES = {
    400: {"description": "Invalid or missing file"},
    500: {"description": "Server processing error"},
}
_CONVERT_RESPONSES = {
    500: {"description": "Conversion error"},
}

app = FastAPI(
    title="Excel VBA → Python Converter",
    version="1.0.0",
    lifespan=lifespan,
)

# Serve static assets (CSS / JS)
app.mount("/static", StaticFiles(directory="static"), name="static")


# ---------------------------------------------------------------------------
# Pydantic request / response models
# ---------------------------------------------------------------------------

class ConvertRequest(BaseModel):
    vba_code: str
    module_name: str = "converted_module"
    target_library: str = "pandas"
    provider: str | None = None

class ModulePayload(BaseModel):
    """A single VBA module as sent from the frontend."""
    model_config = {"extra": "ignore"}

    code: str
    name: str = "module"
    type: str = "unknown"

class ConvertAllRequest(BaseModel):
    modules: list[ModulePayload]
    target_library: str = "pandas"
    provider: str | None = None

class FormulaConvertRequest(BaseModel):
    formula: str
    cell_address: str = "A1"
    sheet_name: str = "Sheet1"
    provider: str | None = None


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _allowed_file(filename: str, extensions: set[str] | None = None) -> bool:
    if extensions is None:
        extensions = cfg.ALL_EXTENSIONS
    return "." in filename and filename.rsplit(".", 1)[1].lower() in extensions


def _save_upload(upload: UploadFile) -> tuple[str, str]:
    filename = secure_filename(upload.filename or "upload.bin")
    filepath = os.path.join(cfg.UPLOAD_FOLDER, filename)
    with open(filepath, "wb") as f:
        shutil.copyfileobj(upload.file, f)
    logger.info("Saved uploaded file: %s", filename)
    # Fix extension mismatch (e.g. .xls file that is actually OpenXML/xlsx)
    filepath, filename = _normalize_extension(filepath)
    return filename, filepath


def _normalize_extension(filepath: str) -> tuple[str, str]:
    """Detect actual file format and rename when the extension is wrong.

    Common case: a `.xls` file that is actually a ZIP-based OpenXML file
    (i.e. `.xlsx` content).  openpyxl refuses to open such files unless the
    extension is `.xlsx`.

    Returns:
        (filepath, filename) — possibly updated.
    """
    ext = os.path.splitext(filepath)[1].lower()
    if ext in ('.xls', '.xla'):
        try:
            with open(filepath, 'rb') as f:
                magic = f.read(4)
            # ZIP files start with b'PK\x03\x04'
            if magic[:2] == b'PK':
                new_ext = '.xlsx' if ext == '.xls' else '.xlam'
                new_path = os.path.splitext(filepath)[0] + new_ext
                os.rename(filepath, new_path)
                logger.info(
                    "Renamed %s -> %s (detected OpenXML inside old extension)",
                    os.path.basename(filepath), os.path.basename(new_path),
                )
                return new_path, os.path.basename(new_path)
        except OSError:
            pass
    return filepath, os.path.basename(filepath)


def _cleanup(filepath: str) -> None:
    try:
        if os.path.exists(filepath):
            os.remove(filepath)
            logger.debug("Cleaned up file: %s", filepath)
    except OSError as exc:
        logger.warning("Failed to clean up %s: %s", filepath, exc)


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
async def index():
    """Serve the main page."""
    return HTMLResponse(content=_TEMPLATE_PATH.read_text(encoding="utf-8"))


@app.get("/api/health")
async def health_check():
    return {
        "status": "ok",
        "provider": cfg.LLM_PROVIDER,
        "model": cfg.LLM_MODEL,
        "max_file_size_mb": cfg.MAX_FILE_SIZE_MB,
    }


# ── VBA upload & extraction ──────────────────────────────────────────────

@app.post("/api/upload", responses=_UPLOAD_RESPONSES)
async def upload_file(file: FileUpload):
    """Upload an Excel file and extract VBA modules."""
    if not file.filename:
        raise HTTPException(400, _ERR_NO_FILE)

    if not _allowed_file(file.filename, cfg.ALL_EXTENSIONS):
        raise HTTPException(
            400,
            f"Invalid file type. Allowed: {', '.join(sorted(cfg.ALL_EXTENSIONS))}",
        )

    filename, filepath = _save_upload(file)
    try:
        extractor = VBAExtractor(filepath)
        vba_modules = extractor.extract_all()

        if not vba_modules:
            ext = os.path.splitext(filepath)[1].lower()
            hint = (
                "This file has no VBA macros. "
                "If you want to extract formulas or data, use the "
                "'Complete Workbook Analysis' section below."
            )
            if ext == '.xlsx':
                hint += " Note: .xlsx files cannot contain VBA — use .xlsm for macro-enabled workbooks."
            return {
                "success": True,
                "filename": filename,
                "warning": hint,
                "modules": [],
            }

        logger.info("Extracted %d VBA module(s) from %s", len(vba_modules), filename)
        return {"success": True, "filename": filename, "modules": vba_modules}

    except Exception as e:
        logger.exception("Error processing upload %s", filename)
        raise HTTPException(500, str(e)) from e
    finally:
        _cleanup(filepath)


# ── VBA → Python conversion ─────────────────────────────────────────────

@app.post("/api/convert", responses=_CONVERT_RESPONSES)
async def convert_vba(body: ConvertRequest):
    """Convert a single VBA module to Python."""
    try:
        if body.provider == "offline":
            converter = OfflineConverter()
            result = converter.convert(
                body.vba_code, body.module_name, body.target_library,
            )
            logger.info(
                "Offline-converted VBA module '%s' (%d chars → %d chars)",
                body.module_name, len(body.vba_code), len(result.python_code),
            )
            return {
                "success": result.success,
                "python_code": result.python_code,
                "conversion_notes": result.conversion_notes,
                "engine": "offline",
            }

        converter = VBAToPythonConverter(provider=body.provider)
        python_code = converter.convert(
            body.vba_code, body.module_name, body.target_library,
        )

        logger.info(
            "Converted VBA module '%s' (%d chars → %d chars)",
            body.module_name, len(body.vba_code), len(python_code),
        )
        return {
            "success": True,
            "python_code": python_code,
            "conversion_notes": converter.get_conversion_notes(),
            "engine": body.provider or cfg.LLM_PROVIDER,
        }

    except Exception as e:
        logger.exception("Conversion failed for module '%s'", body.module_name)
        raise HTTPException(500, str(e)) from e


@app.post("/api/convert-all", responses=_CONVERT_RESPONSES)
async def convert_all_modules(body: ConvertAllRequest):
    """Batch-convert all VBA modules to Python."""
    converted: list[dict] = []

    try:
        if body.provider == "offline":
            offline = OfflineConverter()
            for mod in body.modules:
                result = offline.convert(mod.code, mod.name, body.target_library)
                converted.append({
                    "name": mod.name,
                    "type": mod.type,
                    "original_code": mod.code,
                    "python_code": result.python_code,
                    "conversion_notes": result.conversion_notes,
                })
        else:
            llm_converter = VBAToPythonConverter(provider=body.provider)
            for mod in body.modules:
                python_code = llm_converter.convert(
                    mod.code, mod.name, body.target_library,
                )
                converted.append({
                    "name": mod.name,
                    "type": mod.type,
                    "original_code": mod.code,
                    "python_code": python_code,
                    "conversion_notes": llm_converter.get_conversion_notes(),
                })

        logger.info("Batch-converted %d module(s)", len(converted))
        return {"success": True, "converted_modules": converted}

    except Exception as e:
        logger.exception("Batch conversion failed")
        raise HTTPException(500, str(e)) from e


# ── Download All as ZIP ──────────────────────────────────────────────────

class ZipFileEntry(BaseModel):
    filename: str
    content: str

class DownloadZipRequest(BaseModel):
    files: list[ZipFileEntry]

@app.post("/api/download-zip")
async def download_zip(body: DownloadZipRequest):
    """Package converted Python files into a ZIP and return as download."""
    if not body.files:
        raise HTTPException(400, "No files to package")

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        seen: set[str] = set()
        for entry in body.files:
            # Sanitise filename
            name = re.sub(r'[^a-zA-Z0-9_\-.]', '_', entry.filename)
            if not name.endswith('.py'):
                name += '.py'
            # Deduplicate
            base = name
            counter = 1
            while name in seen:
                name = f"{base[:-3]}_{counter}.py"
                counter += 1
            seen.add(name)
            zf.writestr(name, entry.content)

    buf.seek(0)
    logger.info("Created ZIP with %d file(s)", len(body.files))
    return StreamingResponse(
        buf,
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename=converted_modules.zip"},
    )


# ── Formula extraction & conversion ─────────────────────────────────────

@app.post("/api/extract-formulas", responses=_UPLOAD_RESPONSES)
async def extract_formulas(file: FileUpload):
    """Extract all formulas from an Excel file."""
    if not file.filename:
        raise HTTPException(400, _ERR_NO_FILE)

    if not _allowed_file(file.filename, cfg.DATA_EXTENSIONS):
        raise HTTPException(
            400,
            f"Invalid file type. Allowed: {', '.join(sorted(cfg.DATA_EXTENSIONS))}",
        )

    filename, filepath = _save_upload(file)
    try:
        extractor = FormulaExtractor(filepath)
        formulas = extractor.extract_all_formulas()
        statistics = extractor.get_formula_statistics(formulas)

        formulas_dict = [
            {
                "sheet_name": f.sheet_name,
                "cell_address": f.cell_address,
                "formula": f.formula,
                "formula_type": f.formula_type,
                "dependencies": f.dependencies,
                "functions": f.contains_functions,
            }
            for f in formulas
        ]

        logger.info("Extracted %d formula(s) from %s", len(formulas), filename)
        return {
            "success": True,
            "filename": filename,
            "formulas": formulas_dict,
            "statistics": statistics,
        }

    except Exception as e:
        logger.exception("Formula extraction failed for %s", filename)
        raise HTTPException(500, str(e)) from e
    finally:
        _cleanup(filepath)


@app.post("/api/convert-formula", responses=_CONVERT_RESPONSES)
async def convert_formula(body: FormulaConvertRequest):
    """Convert a single Excel formula to Python."""
    try:
        if body.provider == "offline":
            offline = OfflineConverter()
            result = offline.convert_formula(
                body.formula, body.cell_address, body.sheet_name,
            )
            logger.info("Offline-converted formula at %s!%s", body.sheet_name, body.cell_address)
            return {
                "success": result.success,
                "python_code": result.python_code,
                "conversion_notes": result.conversion_notes,
            }

        converter = VBAToPythonConverter(provider=body.provider)
        python_code = converter.convert_formula(
            body.formula, body.cell_address, body.sheet_name,
        )

        logger.info("Converted formula at %s!%s", body.sheet_name, body.cell_address)
        return {
            "success": True,
            "python_code": python_code,
            "conversion_notes": converter.get_conversion_notes(),
        }

    except Exception as e:
        logger.exception(
            "Formula conversion failed for %s!%s", body.sheet_name, body.cell_address,
        )
        raise HTTPException(500, str(e)) from e


# ── Data export ──────────────────────────────────────────────────────────

@app.post("/api/export-data", responses=_UPLOAD_RESPONSES)
async def export_data(file: FileUpload):
    """Export Excel sheet data to Python/pandas code."""
    if not file.filename:
        raise HTTPException(400, _ERR_NO_FILE)

    if not _allowed_file(file.filename, cfg.DATA_EXTENSIONS):
        raise HTTPException(
            400,
            f"Invalid file type. Allowed: {', '.join(sorted(cfg.DATA_EXTENSIONS))}",
        )

    filename, filepath = _save_upload(file)
    try:
        exporter = DataExporter(filepath)
        result = exporter.export_all_sheets()

        logger.info(
            "Exported data from %s (%d sheet(s))", filename, len(result.sheet_data),
        )
        return {
            "success": True,
            "filename": filename,
            "python_code": result.python_code,
            "metadata": result.metadata,
        }

    except Exception as e:
        logger.exception("Data export failed for %s", filename)
        raise HTTPException(500, str(e)) from e
    finally:
        _cleanup(filepath)


# ── Workbook analysis ────────────────────────────────────────────────────

@app.post("/api/analyze-workbook", responses=_UPLOAD_RESPONSES)
async def analyze_workbook(file: FileUpload):
    """Full workbook analysis (VBA + formulas + data)."""
    if not file.filename:
        raise HTTPException(400, _ERR_NO_FILE)

    if not _allowed_file(file.filename, cfg.DATA_EXTENSIONS):
        raise HTTPException(
            400,
            f"Invalid file type. Allowed: {', '.join(sorted(cfg.DATA_EXTENSIONS))}",
        )

    filename, filepath = _save_upload(file)
    try:
        analyzer = WorkbookAnalyzer(filepath)
        analysis = analyzer.analyze_complete()
        report = analyzer.generate_analysis_report(analysis)

        response_data = {
            "success": True,
            "filename": analysis.filename,
            "has_vba": analysis.has_vba,
            "vba_modules_count": len(analysis.vba_modules) if analysis.vba_modules else 0,
            "has_formulas": analysis.has_formulas,
            "formulas_count": len(analysis.formulas) if analysis.formulas else 0,
            "sheets_count": (
                len(analysis.data_export.sheet_data) if analysis.data_export else 0
            ),
            "python_script": analysis.python_script,
            "report": report,
            "metadata": analysis.data_export.metadata if analysis.data_export else {},
        }

        logger.info(
            "Analyzed workbook %s: %d VBA modules, %d formulas, %d sheets",
            filename,
            response_data["vba_modules_count"],
            response_data["formulas_count"],
            response_data["sheets_count"],
        )
        return response_data

    except Exception as e:
        logger.exception("Workbook analysis failed for %s", filename)
        raise HTTPException(500, str(e)) from e
    finally:
        _cleanup(filepath)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    """Run the application with Uvicorn."""
    uvicorn.run(
        "app:app",
        host="127.0.0.1",
        port=5000,
        reload=cfg.DEBUG,
    )


if __name__ == "__main__":
    main()
