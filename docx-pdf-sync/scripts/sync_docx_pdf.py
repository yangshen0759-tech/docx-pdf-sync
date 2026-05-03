#!/usr/bin/env python3
"""Synchronize a DOCX to a same-name image-only PDF with Word/WPS and PyMuPDF."""

from __future__ import annotations

import argparse
import importlib
import os
import subprocess
import sys
import uuid
from pathlib import Path


WORD_PROG_ID = "Word.Application"
WPS_PROG_ID = "KWPS.Application"
PDF_FORMAT = 17
DEFAULT_RASTER_DPI = 220
DEFAULT_JPEG_QUALITY = 92


def valid_pdf(path: Path) -> bool:
    try:
        if not path.exists() or path.stat().st_size <= 0:
            return False
        with path.open("rb") as handle:
            return handle.read(5) == b"%PDF-"
    except OSError:
        return False


def ensure_pymupdf(auto_install: bool = True):
    """Import PyMuPDF, installing it first when allowed."""
    try:
        return importlib.import_module("fitz"), False
    except ModuleNotFoundError:
        if not auto_install:
            raise RuntimeError(
                "PyMuPDF is required for image-only PDF output. "
                "Install it with: python -m pip install pymupdf"
            )

    print("PyMuPDF not found; installing with: python -m pip install pymupdf", file=sys.stderr)
    result = subprocess.run(
        [sys.executable, "-m", "pip", "install", "pymupdf"],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if result.returncode != 0:
        details = []
        if result.stdout.strip():
            details.append(f"pip stdout:\n{result.stdout.strip()}")
        if result.stderr.strip():
            details.append(f"pip stderr:\n{result.stderr.strip()}")
        raise RuntimeError(
            "PyMuPDF is required for image-only PDF output, and automatic installation failed. "
            "Run this command manually, then retry: python -m pip install pymupdf\n\n"
            + "\n\n".join(details)
        )

    importlib.invalidate_caches()
    return importlib.import_module("fitz"), True


def kill_process_tree(pid: int) -> None:
    if os.name == "nt":
        subprocess.run(
            ["taskkill", "/F", "/T", "/PID", str(pid)],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )
        return

    try:
        os.kill(pid, 9)
    except OSError:
        pass


def run_office_export(
    prog_id: str,
    label: str,
    input_docx: Path,
    output_pdf: Path,
    timeout_seconds: int,
) -> subprocess.CompletedProcess[str]:
    if os.name != "nt":
        return subprocess.CompletedProcess(
            ["powershell", prog_id],
            127,
            "",
            f"{label} COM PDF export is only available on Windows.",
        )

    ps_script = rf"""
$ErrorActionPreference = 'Stop'
$InputDocx = $env:DOCX_VIEWER_INPUT
$OutputPdf = $env:DOCX_VIEWER_OUTPUT
if ([string]::IsNullOrWhiteSpace($InputDocx) -or [string]::IsNullOrWhiteSpace($OutputPdf)) {{
  throw 'DOCX_VIEWER_INPUT and DOCX_VIEWER_OUTPUT must be set.'
}}
$app = $null
$doc = $null
$wdDoNotSaveChanges = 0

try {{
  $app = New-Object -ComObject '{prog_id}'
  $app.Visible = $false
  try {{ $app.DisplayAlerts = 0 }} catch {{}}
  $doc = $app.Documents.Open($InputDocx, $false, $true, $false)
  $doc.ExportAsFixedFormat(
    $OutputPdf,
    {PDF_FORMAT},
    $false,
    0,
    0,
    1,
    1,
    0,
    $true,
    $true,
    0,
    $true,
    $true,
    $false
  )
}}
finally {{
  if ($null -ne $doc) {{
    $doc.Close([ref]$wdDoNotSaveChanges) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
  }}
  if ($null -ne $app) {{
    try {{ $app.Quit([ref]$wdDoNotSaveChanges) | Out-Null }} catch {{ $app.Quit() | Out-Null }}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
  }}
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}}
""".strip()

    cmd = [
        "powershell",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        ps_script,
    ]
    env = os.environ.copy()
    env["DOCX_VIEWER_INPUT"] = str(input_docx)
    env["DOCX_VIEWER_OUTPUT"] = str(output_pdf)
    proc = subprocess.Popen(
        cmd,
        env=env,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    try:
        stdout, stderr = proc.communicate(timeout=timeout_seconds)
        return subprocess.CompletedProcess(cmd, proc.returncode, stdout, stderr)
    except subprocess.TimeoutExpired:
        kill_process_tree(proc.pid)
        stdout, stderr = proc.communicate()
        stderr = (stderr or "") + f"\n{label} PDF export timed out after {timeout_seconds} seconds."
        return subprocess.CompletedProcess(cmd, 124, stdout, stderr)


def export_with_fallback(
    input_docx: Path,
    staging_pdf: Path,
    timeout_seconds: int,
) -> tuple[str, list[str]]:
    failures: list[str] = []

    for prog_id, label in ((WORD_PROG_ID, "Microsoft Word"), (WPS_PROG_ID, "WPS Writer")):
        staging_pdf.unlink(missing_ok=True)
        result = run_office_export(prog_id, label, input_docx, staging_pdf, timeout_seconds)
        if valid_pdf(staging_pdf):
            return label, failures

        details = [f"{label} export failed; exit code: {result.returncode}"]
        if result.stdout.strip():
            details.append(f"{label} stdout:\n{result.stdout.strip()}")
        if result.stderr.strip():
            details.append(f"{label} stderr:\n{result.stderr.strip()}")
        failures.append("\n".join(details))

    raise RuntimeError(
        "PDF synchronization failed; neither Microsoft Word nor WPS Writer generated a valid PDF.\n\n"
        + "\n\n".join(failures)
    )


def rasterize_pdf(
    vector_pdf: Path,
    raster_pdf: Path,
    fitz,
    dpi: int,
    jpeg_quality: int,
) -> None:
    if dpi < 72:
        raise ValueError(f"Raster DPI must be at least 72, got: {dpi}")
    if not 1 <= jpeg_quality <= 100:
        raise ValueError(f"JPEG quality must be between 1 and 100, got: {jpeg_quality}")
    if not valid_pdf(vector_pdf):
        raise RuntimeError(f"Cannot rasterize invalid source PDF: {vector_pdf}")

    raster_pdf.unlink(missing_ok=True)
    matrix = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    source = fitz.open(vector_pdf)
    output = fitz.open()
    try:
        if len(source) == 0:
            raise RuntimeError(f"Source PDF has no pages: {vector_pdf}")

        for page in source:
            pix = page.get_pixmap(matrix=matrix, alpha=False, colorspace=fitz.csRGB)
            image_bytes = pix.tobytes("jpeg", jpg_quality=jpeg_quality)
            raster_page = output.new_page(width=page.rect.width, height=page.rect.height)
            raster_page.insert_image(raster_page.rect, stream=image_bytes)

        output.save(
            raster_pdf,
            garbage=4,
            deflate=True,
            deflate_images=True,
            use_objstms=1,
            compression_effort=9,
        )
    finally:
        output.close()
        source.close()

    if not valid_pdf(raster_pdf):
        raise RuntimeError(f"Rasterized PDF is invalid: {raster_pdf}")


def sync_docx_pdf(
    source_docx: Path,
    timeout_seconds: int = 180,
    dpi: int = DEFAULT_RASTER_DPI,
    jpeg_quality: int = DEFAULT_JPEG_QUALITY,
    auto_install_pymupdf: bool = True,
) -> tuple[Path, str]:
    source_docx = source_docx.expanduser().resolve()
    if not source_docx.exists():
        raise FileNotFoundError(f"DOCX not found: {source_docx}")
    if source_docx.suffix.lower() != ".docx":
        raise ValueError(f"Expected a .docx file, got: {source_docx}")

    target_pdf = source_docx.with_suffix(".pdf")
    vector_staging_pdf = target_pdf.with_name(f"{target_pdf.name}.vector_tmp_{uuid.uuid4().hex}.pdf")
    raster_staging_pdf = target_pdf.with_name(f"{target_pdf.name}.raster_tmp_{uuid.uuid4().hex}.pdf")

    try:
        fitz, installed_pymupdf = ensure_pymupdf(auto_install=auto_install_pymupdf)

        engine, failures = export_with_fallback(source_docx, vector_staging_pdf, timeout_seconds)
        if not valid_pdf(vector_staging_pdf):
            detail = "\n\n".join(failures) if failures else "No exporter failure details were returned."
            raise RuntimeError(f"Generated vector PDF staging file is invalid: {vector_staging_pdf}\n{detail}")

        rasterize_pdf(vector_staging_pdf, raster_staging_pdf, fitz, dpi=dpi, jpeg_quality=jpeg_quality)

        try:
            os.replace(raster_staging_pdf, target_pdf)
        except PermissionError as exc:
            raise PermissionError(
                f"Could not replace target PDF, likely because it is open or locked: {target_pdf}. "
                "Close Word, WPS, PDF viewers, browsers, or sync tools using the file and retry."
            ) from exc

        install_note = "; PyMuPDF installed automatically" if installed_pymupdf else ""
        return target_pdf, f"{engine} + PyMuPDF image-only PDF ({dpi} dpi, JPEG quality {jpeg_quality}{install_note})"
    finally:
        vector_staging_pdf.unlink(missing_ok=True)
        raster_staging_pdf.unlink(missing_ok=True)


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Synchronize a DOCX to a same-name image-only PDF. The script exports with "
            "Microsoft Word first, falls back to WPS Writer, then rasterizes every page "
            "with PyMuPDF for VS Code/PDF.js-compatible preview."
        )
    )
    parser.add_argument("docx", help="Path to the .docx file")
    parser.add_argument("--timeout", type=int, default=180, help="Seconds before giving up on each office export attempt")
    parser.add_argument("--dpi", type=int, default=DEFAULT_RASTER_DPI, help="Rasterization DPI for each PDF page")
    parser.add_argument(
        "--jpeg-quality",
        type=int,
        default=DEFAULT_JPEG_QUALITY,
        help="JPEG quality for embedded page images, from 1 to 100",
    )
    parser.add_argument(
        "--no-auto-install-pymupdf",
        action="store_true",
        help="Fail instead of running python -m pip install pymupdf when PyMuPDF is missing",
    )
    args = parser.parse_args()

    try:
        pdf, engine = sync_docx_pdf(
            Path(args.docx),
            timeout_seconds=args.timeout,
            dpi=args.dpi,
            jpeg_quality=args.jpeg_quality,
            auto_install_pymupdf=not args.no_auto_install_pymupdf,
        )
        print(f"PDF updated: {pdf}")
        print(f"Export pipeline: {engine}")
        print("PDF type: image-only PDF for VS Code/PDF.js preview compatibility")
        return 0
    except PermissionError as exc:
        print(f"PDF synchronization failed: {exc}", file=os.sys.stderr)
        return 13
    except Exception as exc:
        print(f"PDF synchronization failed: {exc}", file=os.sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
