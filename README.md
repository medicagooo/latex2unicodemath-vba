# 在 Word 中批量转换 LaTeX 公式

支持 Windows 桌面 Word，Word 需要提供 LaTeX 公式输入模式。

关闭 Word，用 PowerShell 7 运行安装脚本。

```powershell
pwsh -File .\Install-WordLatex.ps1
```

脚本按照 Word 的 Startup 设置安装内置的 `WordLatexEquations.dotm`。重新打开 Word 即可使用。再次运行脚本会更新模板。要卸载，关闭 Word，然后从安装脚本报告的路径移除该模板。 [Microsoft Word Startup 说明](https://learn.microsoft.com/en-us/office/vba/api/word.application.startuppath)

1. 选中包含公式的正文。
2. 按 `Alt+F8`，运行 `Convert_Selected_LaTeX_To_WordEquations`。
3. `$...$` 和 `\(...\)` 转为行内公式，`\[...\]` 转为独立居中的公式。大批量转换时，Word 会显示进度。用一次 `Ctrl+Z` 撤销整次转换。

如果有同名旧宏，在宏对话框里选择 `WordLatexEquations.dotm`。公式语法以 Word 的 LaTeX 支持范围为准。 [Microsoft Word LaTeX 说明](https://support.microsoft.com/en-gb/word/linear-format-equations-using-unicodemath-and-latex-in-word)

## 更新安装脚本中的模板

1. 在 Word 中创建空白 `.dotm`，在 VBA 编辑器里将 [word latex.vba](word%20latex.vba) 作为 `.bas` 文件导入模板项目。
2. 在同一项目插入空白 UserForm，把 `(Name)` 设为 `LatexBatchProgress`，将 [LatexBatchProgress.vba](LatexBatchProgress.vba) 粘贴到窗体的代码编辑器。
3. 编译项目并保存模板。将模板字节转换成 Base64 后，替换安装脚本的 `$payload` 内容。

安装后，在 PowerShell 7 里运行 `.\tests\Test-WordLatex.ps1`，可以验证三个定界符、选区边界和一次撤销。加上 `-BulkCounts 100,1000,5000` 可测试大批量处理。
