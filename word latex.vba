Option Explicit

'========================
' Main entry
'========================
Public Sub Convert_Selected_LaTeX_To_WordEquations()
    Dim selRng As Range
    If Selection Is Nothing Then Exit Sub
    If Selection.Range Is Nothing Then Exit Sub
    If Selection.Range.Start = Selection.Range.End Then
        MsgBox "请先在文档中框选一段包含 LaTeX 公式定界符的文本区域。", vbExclamation
        Exit Sub
    End If

    Set selRng = Selection.Range.Duplicate

    Dim oldScreenUpdating As Boolean
    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    On Error GoTo CleanFail

    ' Try to switch Equation input to LaTeX mode (best-effort).
    Ensure_EquationMode_LaTeX

    ' Collect matches (0-based indices in selRng.Text)
    Dim starts() As Long, lens() As Long, latex() As String, isDisplay() As Boolean
    Dim n As Long
    n = Collect_LaTeX_Matches(selRng.Text, starts, lens, latex, isDisplay)

    If n = 0 Then
        MsgBox "在所选区域内未找到 $...$、\(...\) 或 \[...\] 形式的 LaTeX 公式。", vbInformation
        GoTo CleanExit
    End If

    ' One undo step (Word 2016+/365 should support UndoRecord)
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "Convert LaTeX to Word Equations"
    On Error GoTo CleanFail

    ' Process from end to start to keep indices stable
    Dim i As Long
    For i = n To 1 Step -1
        Convert_One_Match selRng, starts(i), lens(i), latex(i), isDisplay(i)
    Next i

    On Error Resume Next
    Application.UndoRecord.EndCustomRecord
    On Error GoTo 0

CleanExit:
    Application.ScreenUpdating = oldScreenUpdating
    Exit Sub

CleanFail:
    On Error Resume Next
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = oldScreenUpdating
    MsgBox "转换过程中发生错误：" & vbCrLf & Err.Number & " - " & Err.Description, vbExclamation
End Sub

'========================
' Convert one match into OMath
'========================
Private Sub Convert_One_Match(ByVal baseSel As Range, ByVal start0 As Long, ByVal len0 As Long, _
                              ByVal latexInner As String, ByVal displayMode As Boolean)

    Dim doc As Document
    Set doc = baseSel.Document

    Dim absStart As Long, absEnd As Long
    absStart = baseSel.Start + start0
    absEnd = absStart + len0

    ' Normalize LaTeX (Word OMath can't span paragraphs; replace hard breaks with spaces)
    Dim s As String
    s = latexInner
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Trim$(s)

    ' For display \[...\], force it into its own paragraph (best-effort).
    ' Insert paragraph break BEFORE if needed
    If displayMode Then
        If absStart > 0 Then
            If doc.Range(absStart - 1, absStart).Text <> vbCr Then
                doc.Range(absStart, absStart).InsertBefore vbCr
                absStart = absStart + 1
                absEnd = absEnd + 1
            End If
        End If

        ' Insert paragraph break AFTER if needed
        If absEnd < doc.Content.End Then
            ' Need a safe 1-char range to test; if at boundary, skip
            If absEnd + 1 <= doc.Content.End Then
                If doc.Range(absEnd, absEnd + 1).Text <> vbCr Then
                    doc.Range(absEnd, absEnd).InsertAfter vbCr
                End If
            End If
        End If
    End If

    ' Replace delimiter-wrapped text with pure LaTeX content
    Dim r As Range
    Set r = doc.Range(absStart, absEnd)
    r.Text = s

    ' Turn into OMath and BuildUp
    Dim eqRng As Range
    Set eqRng = r.OMaths.Add(r)          ' returns a range that contains OMath(s) (per Word OMath model)
    Dim eq As OMath
    Set eq = eqRng.OMaths(1)

    ' Justification control
    On Error Resume Next
    If displayMode Then
        eq.Justification = wdOMathJcCenter
        eq.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Else
        eq.Justification = wdOMathJcInline
    End If
    On Error GoTo 0

    ' Build up to professional format
    On Error Resume Next
    eq.BuildUp
    On Error GoTo 0

    ' Important: BuildUp may collapse/shift the working range; normalize to end of equation.
    ' (Common workaround observed in practice.)
    On Error Resume Next
    Dim tmp As Range
    Set tmp = eq.Range
    tmp.Collapse Direction:=wdCollapseEnd
    On Error GoTo 0
End Sub

'========================
' Collect matches using RegExp
' Returns count n; arrays are 1..n
'========================
Private Function Collect_LaTeX_Matches(ByVal textIn As String, _
                                      ByRef starts() As Long, ByRef lens() As Long, _
                                      ByRef latex() As String, ByRef isDisplay() As Boolean) As Long
    Dim re As Object, ms As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")

    ' Pattern supports:
    ' 1) \[ ... \]  (display)     -> capture group 1
    ' 2) \( ... \)  (inline)      -> capture group 2
    ' 3) $ ... $    (inline)      -> capture group 3
    '
    ' Notes:
    ' - [\s\S]*? allows cross-line matching without relying on DOTALL.
    ' - (?!\$) avoids matching $$...$$ as $...$.
    re.Global = True
    re.MultiLine = True
    re.IgnoreCase = False
    re.Pattern = "\\\[([\s\S]*?)\\\]|\\\(([\s\S]*?)\\\)|\$(?!\$)((?:\\\$|[^$])*)\$(?!\$)"

    Set ms = re.Execute(textIn)

    Dim n As Long
    n = 0

    For Each m In ms
        n = n + 1
        ReDim Preserve starts(1 To n)
        ReDim Preserve lens(1 To n)
        ReDim Preserve latex(1 To n)
        ReDim Preserve isDisplay(1 To n)

        starts(n) = CLng(m.FirstIndex)   ' 0-based in textIn
        lens(n) = CLng(m.Length)

        If Len(m.SubMatches(0)) > 0 Then
            isDisplay(n) = True
            latex(n) = CStr(m.SubMatches(0))
        ElseIf Len(m.SubMatches(1)) > 0 Then
            isDisplay(n) = False
            latex(n) = CStr(m.SubMatches(1))
        Else
            isDisplay(n) = False
            latex(n) = CStr(m.SubMatches(2))
        End If
    Next m

    Collect_LaTeX_Matches = n
End Function

'========================
' Best-effort: switch Equation input mode to LaTeX
'========================
Private Sub Ensure_EquationMode_LaTeX()
    On Error Resume Next

    ' Some builds expose an idMso; if unavailable, this fails silently.
    Application.CommandBars.ExecuteMso "EquationModeLaTeX"

    ' Widely used workaround: invoke the Word command via a built-in dialog id.
    ' (Best-effort; may no-op depending on build/UI state.)
    Application.Dialogs(2844).Execute
    DoEvents

    On Error GoTo 0
End Sub
