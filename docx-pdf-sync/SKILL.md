---
name: docx-pdf-sync
description: "After every AI edit to a .docx file, immediately generate or overwrite a same-name image-only .pdf for stable VS Code/PDF.js preview. Applies to DOCX editing, replacement, insertion, deletion, style changes, field updates, headers and footers, tables, images, and batch DOCX edits. The script exports through Microsoft Word COM first, falls back to WPS Writer COM, rasterizes the PDF with PyMuPDF, validates the result, atomically replaces the same-name PDF, and reports the full process."
---

# docx-pdf-sync

## Rules

After any real `.docx` modification, immediately create or overwrite a same-name image-only `.pdf` in the source DOCX directory.

If the DOCX was only read, inspected, or analyzed without saving any modification, do not run this skill.

The generated PDF is image-only. This is intended to keep VS Code/PDF.js preview stable even when Office math fonts or CJK font mappings render incorrectly in vector PDFs. The PDF does not preserve text selection or text search.

## Dependencies

- Windows
- Python 3
- PowerShell
- Microsoft Word or desktop WPS Writer
- PyMuPDF, automatically installed by the script when missing

If PyMuPDF automatic installation fails, tell the user to run:

```powershell
python -m pip install pymupdf
```

If neither Word nor WPS is installed, tell the user to install one desktop version.

## Command

Use the bundled script:

```powershell
python "$env:USERPROFILE\.codex\skills\docx-pdf-sync\scripts\sync_docx_pdf.py" "C:\path\to\modified.docx"
```

Optional timeout:

```powershell
python "$env:USERPROFILE\.codex\skills\docx-pdf-sync\scripts\sync_docx_pdf.py" "C:\path\to\modified.docx" --timeout 300
```

Optional rasterization settings:

```powershell
python "$env:USERPROFILE\.codex\skills\docx-pdf-sync\scripts\sync_docx_pdf.py" "C:\path\to\modified.docx" --dpi 220 --jpeg-quality 92
```

## PDF Sync Process

The script must:

1. Open the original DOCX path directly in read-only mode.
2. Export a temporary PDF with Microsoft Word COM `ExportAsFixedFormat`.
3. If Word does not generate a valid PDF, retry with WPS Writer COM `ExportAsFixedFormat`.
4. Use PyMuPDF to rasterize every temporary PDF page and rebuild an image-only PDF.
5. Validate that the staged PDF exists, is non-empty, and starts with `%PDF-`.
6. Atomically replace the same-name target PDF next to the source DOCX.
7. Report the PDF path and export pipeline.

Never save the source DOCX during PDF synchronization.

## Failure Handling

If PyMuPDF is missing and automatic installation fails, report the pip error details and tell the user to run `python -m pip install pymupdf`.

If both Word and WPS export fail, report both failure details.

If the target PDF is locked, tell the user to close Word, WPS, PDF viewers, browsers, sync tools, or any other program using the file, then retry.

Do not create an alternative PDF filename unless the user explicitly asks for it.

Do not roll back completed DOCX edits because PDF synchronization failed unless the user explicitly asks for rollback.

## Final Response

Whenever a DOCX was modified, the final response must include:

- The modified `.docx` path.
- The same-name image-only `.pdf` path if synchronization succeeded.
- The export pipeline, including Microsoft Word or WPS Writer and PyMuPDF rasterization settings.
- A clear statement that the PDF is an image-only PDF for VS Code/PDF.js preview compatibility.
- If synchronization failed, the failure reason and the recommended next step.
