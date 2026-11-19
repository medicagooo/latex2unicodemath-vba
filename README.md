# LaTeX to Word Equations

English · [简体中文](README.zh-CN.md)

Convert LaTeX expressions in selected text into editable Microsoft Word equations. The macro converts inline and display equations together, shows progress on large selections, and lets you undo the whole conversion with one press of `Ctrl+Z`.

The PowerShell installer includes a Word template that makes the macro available in new and existing documents for your Windows account.

## Install

You need PowerShell 7 and the Windows desktop version of Word with LaTeX equation input. [Supported LaTeX syntax in Word](https://support.microsoft.com/en-gb/word/linear-format-equations-using-unicodemath-and-latex-in-word)

Close Word, open PowerShell in the folder containing [Install-WordLatex.ps1](Install-WordLatex.ps1), and run the installer.

```powershell
pwsh -File .\Install-WordLatex.ps1
```

The script reports where it installed `WordLatexEquations.dotm`. Word loads the template from its configured Startup folder the next time it starts. [Microsoft Word Startup documentation](https://learn.microsoft.com/en-us/office/vba/api/word.application.startuppath)

## Convert equations

1. Open or create a document in Word.
2. Select a continuous range of text containing the LaTeX expressions.
3. Press `Alt+F8`, select `Convert_Selected_LaTeX_To_WordEquations`, and click **Run**. If an older macro has the same name, choose `WordLatexEquations.dotm` in the Macros dialog.

| Delimiters | Result |
| --- | --- |
| `$...$` | Inline equation |
| `\(...\)` | Inline equation |
| `\[...\]` | Centered equation on its own line |

To try it, paste the following text into Word and select it before running the macro.

```text
The area is $A = \pi r^2$ and the second term is \(x_2\).
\[\frac{a}{b} = c\]
```

The example produces two inline equations and one centered equation.

## Update or uninstall

To update, close Word and run the newer installer. To uninstall, close Word and remove `WordLatexEquations.dotm` from the path reported by the installer.

## Rebuild the template

1. Create a blank `.dotm` template in Word. In the VBA editor, import [word latex.vba](word%20latex.vba) as a `.bas` file into the template project.
2. Add a blank UserForm to the same project and set its `(Name)` to `LatexBatchProgress`. Paste [LatexBatchProgress.vba](LatexBatchProgress.vba) into the form's code window.
3. Compile the VBA project and save the template. Encode the template bytes as Base64 and replace the installer's `$payload`.

After installation, run `.\tests\Test-WordLatex.ps1` from PowerShell 7 to check delimiter handling, selection boundaries, and undo. Pass `-BulkCounts 100,1000,5000` to test larger selections.
