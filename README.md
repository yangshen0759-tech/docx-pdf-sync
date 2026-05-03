# docx-pdf-sync

docx-pdf-sync 是一个用于论文写作场景的 AI agent skill。

在用 IDE 或终端管理工具写科研论文时，我们经常会让 AI agent 帮忙修改 `.docx` 文件，比如插入内容、调整格式、修改表格、处理页眉页脚等。但很多 IDE 和终端管理工具不能直接、完整地预览 Word 文档，例如 VS Code、Codex 桌面端、tmux、Wave 等。修改完以后还要切到 Word 或 WPS 里查看版式，很麻烦。

PDF 就方便很多。几乎所有 IDE 和终端管理工具都能正常显示 PDF，而且 PDF 更接近论文最后打印、提交或送审时的效果。

所以这个 skill 做的事情很简单：**当 AI agent 修改 DOCX 后，自动生成并覆盖同名图像型 PDF。**

这样你只需要在 IDE 或终端管理工具里打开 PDF。之后让 agent 去修改 DOCX，修改完成后 PDF 也会同步更新。你看到的 PDF，就是当前 DOCX 的最新结果。

## 功能

- 在 AI agent 修改 `.docx` 后，自动创建或更新同名 `.pdf`
- 每次修改后覆盖旧版 PDF，保证当前 PDF 是最新版
- 生成图像型 PDF
- 如果 Word 不可用，自动尝试使用 WPS 导出
- 验证生成的 PDF 是否有效
- 报告本次成功使用的导出引擎
- 让你直接在 IDE 或终端管理工具中查看论文修改后的版式效果

## 适合谁使用

适合这些场景：

- 用 AI agent 辅助写论文、改论文或调整 Word 格式
- 主要工作环境是 VS Code、Codex 桌面端、tmux、Wave 或其他 IDE/终端管理工具
- 希望不用反复切换到 Word/WPS，也能查看 agent 修改后的文档效果
- 希望每次 DOCX 修改后，PDF 自动保持最新

## 工作方式

安装后，通常不需要你手动运行转换命令。

在支持 skills 的 agent 中，agent 对 `.docx` 做了实际修改后，会根据 skill 说明自动调用 `docx-pdf-sync`：

1. 找到被修改的 DOCX 文件。
2. 优先调用 Microsoft Word 生成图像型 PDF。
3. 如果 Word 没有成功生成有效 PDF，则改用 WPS Writer。
4. 检查 PDF 是否存在、是否非空、是否是有效 PDF。
5. 用新 PDF 覆盖旧的同名 PDF。
6. 告诉你 PDF 已更新，以及使用的是 Word 还是 WPS。

例如：

```text
论文.docx
论文.pdf
```

当 agent 修改 `论文.docx` 后，`论文.pdf` 会自动更新。

## 安装要求

- Windows
- Python 3
- PowerShell
- 桌面版 Microsoft Word 或桌面版 WPS Writer

至少安装 Word 或 WPS Writer 中的一个即可。两个都安装时，默认先用 Word，Word 不成功再用 WPS。

生成图像型 PDF 需要 PyMuPDF。你不需要手动安装它；脚本检测到电脑里没有 PyMuPDF 时，会自动尝试安装。

## 安装方法

下面的命令适合 Windows 用户。你只需要复制对应命令，然后粘贴到 PowerShell 里运行。

### 如何打开 PowerShell

1. 按键盘上的 `Win` 键。
2. 输入 `PowerShell`。
3. 打开 `Windows PowerShell`。
4. 复制下面对应的命令。
5. 在 PowerShell 窗口中右键粘贴。
6. 按回车运行。

### Codex

如果你使用的是 Codex，把下面这一整行复制到 PowerShell 中运行：

```powershell
New-Item -ItemType Directory -Force "$env:USERPROFILE\.codex\skills" | Out-Null; if (Test-Path "$env:USERPROFILE\.codex\skills\docx-pdf-sync") { Remove-Item "$env:USERPROFILE\.codex\skills\docx-pdf-sync" -Recurse -Force }; git clone https://github.com/yangshen0759-tech/docx-pdf-sync.git "$env:USERPROFILE\.codex\skills\docx-pdf-sync"
```

这条命令会做两件事：

1. 在你的用户目录下创建 Codex 的 skills 文件夹。
2. 从 GitHub 下载 `docx-pdf-sync`，并放到 Codex 能识别的位置。如果之前安装过，会自动覆盖为最新版。

### Claude Code

如果你使用的是 Claude Code，把下面这一整行复制到 PowerShell 中运行：

```powershell
New-Item -ItemType Directory -Force "$env:USERPROFILE\.claude\skills" | Out-Null; if (Test-Path "$env:USERPROFILE\.claude\skills\docx-pdf-sync") { Remove-Item "$env:USERPROFILE\.claude\skills\docx-pdf-sync" -Recurse -Force }; git clone https://github.com/yangshen0759-tech/docx-pdf-sync.git "$env:USERPROFILE\.claude\skills\docx-pdf-sync"
```

这条命令会做两件事：

1. 在你的用户目录下创建 Claude Code 的 skills 文件夹。
2. 从 GitHub 下载 `docx-pdf-sync`，并放到 Claude Code 能识别的位置。如果之前安装过，会自动覆盖为最新版。

如果你不想自己输入命令，也可以直接把 GitHub 仓库网址复制给 AI：

```text
https://github.com/yangshen0759-tech/docx-pdf-sync
```

然后告诉它：

```text
帮我把这个 GitHub 仓库安装到我的 agent skill 文件夹中。
```

AI 通常可以帮你完成下载和放置。

## 安装后怎么确认

安装完成后，在 AI 对话框里输入 `/`，确认可以检索到：

```text
docx-pdf-sync
```

能检索到就说明 skill 已经被 agent 识别。

之后通常不需要你手动运行任何转换命令。只要 agent 修改 DOCX，它就会根据这个 skill 自动创建或更新同名 PDF。

## 使用时的注意事项

如果同步失败，优先检查是否有软件正在占用 DOCX 文件。

请关闭任何可能正在占用 DOCX 的软件，例如：

- Microsoft Word
- WPS Writer
- 文件同步工具
- 其他正在打开或锁定该 DOCX 的程序

实测在 IDE 或终端管理工具中打开 PDF 通常不影响更新；真正容易影响修改和同步的是 DOCX 被占用。

## 已知限制

- 仅支持 Windows
- 需要桌面版 Microsoft Word 或桌面版 WPS Writer
- 不支持 Word Online、WPS 网页版、macOS、Linux 或 WSL
- 无法绕过 Office/WPS 激活、企业策略限制或被禁用的 COM 自动化

## 许可证

MIT。
