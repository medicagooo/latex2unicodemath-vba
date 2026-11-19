# WORD-001 · Word LaTeX 批量转换和全局安装

状态：实施中。

待执行：WORD-001-002。

## WORD-001-001 · 用户需求 · 2026-11-19

修复大量选中文本转换时 Word 卡住的问题。保留 `$...$`、`\(...\)`、`\[...\]` 和原生可编辑公式。用桌面 Word 验证批量转换、选区边界、独立公式排版和一次撤销。

交付一个内含宏模板的 Windows PowerShell 安装脚本。安装到当前用户的 Word 启动目录，使重启后的新建和已有文档都可以运行宏。用户授权实施、审查、修复和推送到 `https://github.com/medicagooo/latex2unicodemath-vba.git` 的 `main`。

## WORD-001-002 · 初始化 Git · 2026-11-19

依据 WORD-001-001。在 `D:/proj/latex2unicodemath-vba` 初始化 Git 并直接使用 `main`，配置指定 GitHub 仓库为 `origin`。起始目录只有原始 `word latex.vba` 且无本地 Git 仓库，远端无任何 ref。预期第一份本地提交保留原始 VBA 和需求记录，用作修复的对照基线。
