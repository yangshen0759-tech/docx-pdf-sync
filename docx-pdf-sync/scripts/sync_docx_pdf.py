#!/usr/bin/env python3
"""将 DOCX 同步为同名图像型 PDF，使用 Word/WPS 导出并由 PyMuPDF 栅格化。"""

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
    """导入 PyMuPDF；允许时会先自动安装。"""
    try:
        return importlib.import_module("fitz"), False
    except ModuleNotFoundError:
        if not auto_install:
            raise RuntimeError(
                "生成图像型 PDF 需要 PyMuPDF。"
                "请先安装：python -m pip install pymupdf"
            )

    print("未检测到 PyMuPDF，正在执行：python -m pip install pymupdf", file=sys.stderr)
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
            details.append(f"pip 标准输出：\n{result.stdout.strip()}")
        if result.stderr.strip():
            details.append(f"pip 错误输出：\n{result.stderr.strip()}")
        raise RuntimeError(
            "生成图像型 PDF 需要 PyMuPDF，但自动安装失败。"
            "请手动运行以下命令后重试：python -m pip install pymupdf\n\n"
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
            f"{label} COM PDF 导出仅支持 Windows。",
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
        stderr = (stderr or "") + f"\n{label} PDF 导出超过 {timeout_seconds} 秒后超时。"
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

        details = [f"{label} 导出失败；退出码：{result.returncode}"]
        if result.stdout.strip():
            details.append(f"{label} 标准输出：\n{result.stdout.strip()}")
        if result.stderr.strip():
            details.append(f"{label} 错误输出：\n{result.stderr.strip()}")
        failures.append("\n".join(details))

    raise RuntimeError(
        "PDF 同步失败：Microsoft Word 和 WPS Writer 都没有生成有效 PDF。\n\n"
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
        raise ValueError(f"栅格化 DPI 至少应为 72，当前为：{dpi}")
    if not 1 <= jpeg_quality <= 100:
        raise ValueError(f"JPEG 质量必须在 1 到 100 之间，当前为：{jpeg_quality}")
    if not valid_pdf(vector_pdf):
        raise RuntimeError(f"无法栅格化无效源 PDF：{vector_pdf}")

    raster_pdf.unlink(missing_ok=True)
    matrix = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    source = fitz.open(vector_pdf)
    output = fitz.open()
    try:
        if len(source) == 0:
            raise RuntimeError(f"源 PDF 没有页面：{vector_pdf}")

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
        raise RuntimeError(f"栅格化后的 PDF 无效：{raster_pdf}")


def sync_docx_pdf(
    source_docx: Path,
    timeout_seconds: int = 180,
    dpi: int = DEFAULT_RASTER_DPI,
    jpeg_quality: int = DEFAULT_JPEG_QUALITY,
    auto_install_pymupdf: bool = True,
) -> tuple[Path, str]:
    source_docx = source_docx.expanduser().resolve()
    if not source_docx.exists():
        raise FileNotFoundError(f"未找到 DOCX：{source_docx}")
    if source_docx.suffix.lower() != ".docx":
        raise ValueError(f"需要 .docx 文件，当前为：{source_docx}")

    target_pdf = source_docx.with_suffix(".pdf")
    vector_staging_pdf = target_pdf.with_name(f"{target_pdf.name}.vector_tmp_{uuid.uuid4().hex}.pdf")
    raster_staging_pdf = target_pdf.with_name(f"{target_pdf.name}.raster_tmp_{uuid.uuid4().hex}.pdf")

    try:
        fitz, installed_pymupdf = ensure_pymupdf(auto_install=auto_install_pymupdf)

        engine, failures = export_with_fallback(source_docx, vector_staging_pdf, timeout_seconds)
        if not valid_pdf(vector_staging_pdf):
            detail = "\n\n".join(failures) if failures else "导出器没有返回失败详情。"
            raise RuntimeError(f"生成的临时矢量 PDF 无效：{vector_staging_pdf}\n{detail}")

        rasterize_pdf(vector_staging_pdf, raster_staging_pdf, fitz, dpi=dpi, jpeg_quality=jpeg_quality)

        try:
            os.replace(raster_staging_pdf, target_pdf)
        except PermissionError as exc:
            raise PermissionError(
                f"无法替换目标 PDF，可能是文件正在打开或被锁定：{target_pdf}。"
                "请关闭占用该文件的 Word、WPS、PDF 阅读器、浏览器或同步工具后重试。"
            ) from exc

        install_note = "；PyMuPDF 已自动安装" if installed_pymupdf else ""
        return target_pdf, f"{engine} + PyMuPDF 图像型 PDF（{dpi} dpi，JPEG 质量 {jpeg_quality}{install_note}）"
    finally:
        vector_staging_pdf.unlink(missing_ok=True)
        raster_staging_pdf.unlink(missing_ok=True)


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "将 DOCX 同步为同名图像型 PDF。脚本优先使用 Microsoft Word 导出，失败时回退到 WPS Writer，"
            "随后使用 PyMuPDF 栅格化每一页，以兼容 VS Code/PDF.js 预览。"
        )
    )
    parser.add_argument("docx", help="待同步的 .docx 文件路径")
    parser.add_argument("--timeout", type=int, default=180, help="每次 Office 导出尝试的超时时间，单位为秒")
    parser.add_argument("--dpi", type=int, default=DEFAULT_RASTER_DPI, help="每页 PDF 的栅格化 DPI")
    parser.add_argument(
        "--jpeg-quality",
        type=int,
        default=DEFAULT_JPEG_QUALITY,
        help="嵌入页面图片的 JPEG 质量，范围为 1 到 100",
    )
    parser.add_argument(
        "--no-auto-install-pymupdf",
        action="store_true",
        help="缺少 PyMuPDF 时直接失败，而不是自动执行 python -m pip install pymupdf",
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
        print(f"PDF 已更新：{pdf}")
        print(f"导出流程：{engine}")
        print("PDF 类型：用于兼容 VS Code/PDF.js 预览的图像型 PDF")
        return 0
    except PermissionError as exc:
        print(f"PDF 同步失败：{exc}", file=os.sys.stderr)
        return 13
    except Exception as exc:
        print(f"PDF 同步失败：{exc}", file=os.sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
