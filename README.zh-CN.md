# 在 Word 中批量转换 LaTeX 公式

[English](README.md) · 简体中文

这个 Word 宏可以把选中文本中的 LaTeX 表达式批量转换为可编辑的 Word 原生公式。行内公式和独立公式可以一起转换。处理大量公式时会显示进度，转换完成后按一次 `Ctrl+Z` 即可撤销整批改动。

PowerShell 安装脚本内置 Word 模板。安装后，当前 Windows 用户在新建或已有的 Word 文档中都能调用这个宏。

## 安装

需要 PowerShell 7 和支持 LaTeX 公式输入的 Windows 桌面版 Word。 [Word 支持的 LaTeX 语法](https://support.microsoft.com/en-gb/word/linear-format-equations-using-unicodemath-and-latex-in-word)

关闭 Word，在 [Install-WordLatex.ps1](Install-WordLatex.ps1) 所在文件夹打开 PowerShell，运行安装脚本。

```powershell
pwsh -File .\Install-WordLatex.ps1
```

脚本会显示 `WordLatexEquations.dotm` 的安装路径。再次启动 Word 时，Word 会从配置的 Startup 文件夹自动加载模板。 [Microsoft Word Startup 说明](https://learn.microsoft.com/en-us/office/vba/api/word.application.startuppath)

## 转换公式

1. 在 Word 中打开或新建文档。
2. 选中一段包含 LaTeX 表达式的连续文本。
3. 按 `Alt+F8`，选择 `Convert_Selected_LaTeX_To_WordEquations`，点击**运行**。如果有同名旧宏，在宏对话框中选择 `WordLatexEquations.dotm`。

| 定界符 | 转换结果 |
| --- | --- |
| `$...$` | 行内公式 |
| `\(...\)` | 行内公式 |
| `\[...\]` | 单独成行的居中公式 |

可以把下面的文本粘贴到 Word，选中后运行宏。

```text
面积是 $A = \pi r^2$，第二项是 \(x_2\)。
\[\frac{a}{b} = c\]
```

示例会生成两个行内公式和一个居中公式。

## 更新或卸载

更新时，关闭 Word 并运行新版安装脚本。卸载时，关闭 Word，然后从安装脚本显示的路径移除 `WordLatexEquations.dotm`。

## 重新构建模板

1. 在 Word 中创建空白 `.dotm` 模板，在 VBA 编辑器中将 [word latex.vba](word%20latex.vba) 作为 `.bas` 文件导入模板项目。
2. 在同一项目中插入空白 UserForm，将 `(Name)` 设为 `LatexBatchProgress`，把 [LatexBatchProgress.vba](LatexBatchProgress.vba) 粘贴到窗体的代码编辑器。
3. 编译 VBA 项目并保存模板。将模板字节转换为 Base64，替换安装脚本中的 `$payload`。

安装后，在 PowerShell 7 中运行 `.\tests\Test-WordLatex.ps1`，可以验证定界符、选区边界和撤销。加上 `-BulkCounts 100,1000,5000` 可测试大批量处理。
