---
name: docx-pdf-sync
description: "在每次 AI 修改 .docx 文件后，立即同步生成或覆盖同名图像型 .pdf，确保 VS Code/PDF.js 稳定预览。适用于任何 DOCX 编辑、替换、插入、删除、样式调整、字段更新、页眉页脚修改、表格或图片修改、批量 DOCX 修改等场景。脚本会先用 Microsoft Word COM 导出，失败时回退到 WPS Writer COM，再将 PDF 转成图像型 PDF，验证后原子替换同名 PDF，并报告完整流程。"
---

# docx-pdf-sync

## 规则

任何实际 `.docx` 修改完成后，都要立即在源 DOCX 所在目录创建或覆盖同名图像型 `.pdf`。

如果只是读取、检查或分析 DOCX，没有保存任何修改，不要运行此 skill。

生成的 PDF 是图像型 PDF：即使 Office 数学字体或中日韩字体映射在矢量 PDF 中显示异常，也应能在 VS Code/PDF.js 中稳定显示。该 PDF 不保留文本选择和文本搜索能力。

## 依赖

- Windows
- Python 3
- PowerShell
- Microsoft Word 或 WPS Writer 桌面版
- PyMuPDF（脚本会在缺失时自动安装）

如果 PyMuPDF 自动安装失败，指导用户手动执行：

```powershell
python -m pip install pymupdf
```

如果 Word 或 WPS 没有安装，指导用户安装其中一个桌面版。

## 命令

使用内置脚本：

```powershell
python "$env:USERPROFILE\.codex\skills\docx-pdf-sync\scripts\sync_docx_pdf.py" "C:\path\to\modified.docx"
```

可选超时设置：

```powershell
python "$env:USERPROFILE\.codex\skills\docx-pdf-sync\scripts\sync_docx_pdf.py" "C:\path\to\modified.docx" --timeout 300
```

可选栅格化设置：

```powershell
python "$env:USERPROFILE\.codex\skills\docx-pdf-sync\scripts\sync_docx_pdf.py" "C:\path\to\modified.docx" --dpi 220 --jpeg-quality 92
```

## PDF 同步流程

脚本必须：

1. 以只读方式直接打开原始 DOCX 路径。
2. 使用 Microsoft Word COM 的 `ExportAsFixedFormat` 导出临时 PDF。
3. 如果 Word 未生成有效 PDF，改用 WPS Writer COM 的 `ExportAsFixedFormat` 重试。
4. 使用 PyMuPDF 将临时 PDF 的每一页转为图像并重建为图像型 PDF。
5. 验证生成的暂存 PDF 存在、非空，且以 `%PDF-` 开头。
6. 原子替换源 DOCX 旁边的同名目标 PDF。
7. 报告 PDF 路径和使用的导出流程。

PDF 同步过程中绝不能保存源 DOCX。

## 失败处理

如果 PyMuPDF 缺失且自动安装失败，报告 pip 错误详情，并提示用户手动运行 `python -m pip install pymupdf`。

如果 Word 和 WPS 都导出失败，报告两者的失败详情。

如果目标 PDF 被锁定，告知用户关闭 Word、WPS、PDF 阅读器、浏览器、同步工具或其他占用该文件的程序后重试。

除非用户明确要求，不要创建替代 PDF 文件名。

除非用户明确要求回滚，不要因为 PDF 同步失败而撤销已经完成的 DOCX 修改。

## 最终回复

只要修改过 DOCX，最终回复必须包括：

- 已修改的 `.docx` 路径。
- 如果同步成功，给出同名图像型 `.pdf` 路径。
- 使用的导出流程，包括 Microsoft Word 或 WPS Writer，以及 PyMuPDF 栅格化设置。
- 明确说明该 PDF 是面向 VS Code/PDF.js 预览兼容性的图像型 PDF。
- 如果同步失败，说明失败原因和建议的下一步。
