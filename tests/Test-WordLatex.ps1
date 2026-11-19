#Requires -Version 7.0
[CmdletBinding()]
param(
    [string] $TemplatePath,
    [int[]] $BulkCounts = @()
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$macro = 'WordLatex.Convert_Selected_LaTeX_To_WordEquations'

function Assert-WordResult([bool] $Condition, [string] $Message) {
    if (-not $Condition) { throw $Message }
}

function Get-LatexInputMode {
    $probe = $word.Documents.Add()
    try {
        $probe.Range(0, 0).Text = 'x'
        $range = $probe.Range(0, 1)
        $probeEquation = $range.OMaths.Add($range)
        $probeEquation.Select()
        $word.CommandBars.GetPressedMso('EquationLaTexFormat')
    } finally {
        $probe.Close(0)
    }
}

function Test-Conversion(
    [string] $Text, [int] $Count, [string] $Before = '', [string] $After = '',
    [switch] $CheckLayout
) {
    if ($CheckLayout) { $latexInputMode = Get-LatexInputMode }
    $doc = $word.Documents.Add()
    try {
        $original = $Before + $Text + $After + "`r"
        $doc.Range(0, 0).Text = $original.Substring(0, $original.Length - 1)
        $doc.Range($Before.Length, $Before.Length + $Text.Length).Select()
        $viewType = $word.ActiveWindow.View.Type
        $pagination = $word.Options.Pagination
        $screenUpdating = $word.ScreenUpdating
        $timer = [Diagnostics.Stopwatch]::StartNew()
        $null = $word.Run($macro)
        $timer.Stop()

        Assert-WordResult ($doc.OMaths.Count -eq $Count) "Expected $Count equations, found $($doc.OMaths.Count)."
        Assert-WordResult ($doc.Content.Text.StartsWith($Before)) 'Text before the selection changed.'
        Assert-WordResult ($doc.Content.Text.EndsWith("$After`r")) 'Text after the selection changed.'
        Assert-WordResult ($viewType -eq $word.ActiveWindow.View.Type) 'Word did not restore the document view.'
        Assert-WordResult ($pagination -eq $word.Options.Pagination) 'Word did not restore background pagination.'
        Assert-WordResult ($screenUpdating -eq $word.ScreenUpdating) 'Word did not restore screen updating.'
        if ($CheckLayout) {
            $xml = $doc.Content.WordOpenXML
            Assert-WordResult ($xml -match '<m:f[ >]') 'The LaTeX fraction did not build into a native fraction.'
            Assert-WordResult ($xml -match '<m:sSub[ >]') 'The parenthesized formula did not build into a native subscript.'
            Assert-WordResult ($xml -match '<m:rad[ >]') 'The display formula did not build into a native radical.'
            Assert-WordResult ($doc.OMaths.Item(1).Type -eq 1 -and $doc.OMaths.Item(2).Type -eq 1) 'Expected two inline equations.'
            Assert-WordResult ($doc.OMaths.Item(3).Type -eq 0) 'Expected a display equation.'
            Assert-WordResult ($doc.OMaths.Item(3).Range.ParagraphFormat.Alignment -eq 1) 'The display equation was not centered.'
            Assert-WordResult ($doc.Content.Text.Contains('$$untouched$$ \$literal \( \)')) 'Ignored delimiters were modified.'
            Assert-WordResult ($doc.Content.Text.Contains('\[unfinished prose ')) 'An unmatched display opener consumed surrounding text.'
            $doc.OMaths.Item(1).Range.Select()
            Assert-WordResult ($latexInputMode -eq $word.CommandBars.GetPressedMso('EquationLaTexFormat')) 'Word did not restore the equation input mode.'
        }
        Assert-WordResult ($doc.Undo(1)) 'The conversion could not be undone.'
        Assert-WordResult ($doc.Content.Text -ceq $original -and $doc.OMaths.Count -eq 0) 'One Undo did not restore the original document.'

        [pscustomobject]@{ Equations = $Count; Seconds = [math]::Round($timer.Elapsed.TotalSeconds, 3) }
    } finally {
        $doc.Close(0)
    }
}

$word = New-Object -ComObject Word.Application
try {
    if ($TemplatePath) {
        $template = (Resolve-Path -LiteralPath $TemplatePath).Path
        $null = $word.AddIns.Add($template, $true)
    }
    $word.Visible = $true
    $sample = 'Selected ' + [char]::ConvertFromUtf32(0x1F642) + ' $\frac{a}{b}$, \(x_2\), then \[unfinished prose \[\sqrt{3}\] and $$untouched$$ \$literal \( \).'
    Test-Conversion -Text $sample `
        -Count 3 -Before 'Outside $before$. ' -After ' Outside $after$.' -CheckLayout
    Test-Conversion -Text (('\[unmatched ' * 5000) + '$x$') -Count 1
    foreach ($count in $BulkCounts) {
        $text = (1..$count | ForEach-Object { 'Before $x_{' + $_ + '}+1$ after' }) -join "`r"
        Test-Conversion -Text $text -Count $count
    }
} finally {
    try {
        $word.Quit(0)
    } finally {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
    }
}
