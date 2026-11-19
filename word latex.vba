Attribute VB_Name = "WordLatex"
Option Explicit

Private Const LATEX_FORMAT As String = "EquationLaTexFormat"
Private Const UNICODE_FORMAT As String = "EquationUnicodeFormat"
Private Const PROGRESS_THRESHOLD As Long = 100

Private Type LatexMatch
    Offset As Long
    Length As Long
    Formula As String
    Display As Boolean
End Type

Private batchActive As Boolean
Private matches() As LatexMatch
Private total As Long
Private remaining As Long
Private baseStart As Long
Private work As Range
Private modeAnchor As Range
Private restoreUnicode As Boolean
Private checkExistingMath As Boolean
Private edited As Boolean
Private batchError As String

Public Sub Convert_Selected_LaTeX_To_WordEquations()
    If batchActive Then Exit Sub
    If Documents.Count = 0 Then Exit Sub

    Dim source As Range
    Set source = Selection.Range.Duplicate
    If source.Start = source.End Then
        MsgBox "Select text containing LaTeX equations first.", vbInformation
        Exit Sub
    End If

    Dim doc As Document
    Dim docView As View
    Dim undo As UndoRecord
    Dim undoOpen As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldViewType As WdViewType
    Dim oldPagination As Boolean
    Dim changedView As Boolean
    Dim changedPagination As Boolean
    Dim progress As LatexBatchProgress
    Dim failure As String

    batchActive = True
    oldScreenUpdating = Application.ScreenUpdating
    On Error GoTo Failed

    Set doc = source.Document
    Dim snapshot As String
    snapshot = source.Text
    If Len(snapshot) <> source.End - source.Start Then
        Err.Raise vbObjectError + 410, , "Word did not return a consistent text range. Select a smaller, continuous area."
    End If

    ScanMatches snapshot
    If total = 0 Then
        MsgBox "No LaTeX equations delimited by $...$, \(...\), or \[...\] were found.", vbInformation
        GoTo Cleanup
    End If

    baseStart = source.Start
    Set work = source.Duplicate
    checkExistingMath = (source.OMaths.Count > 0)
    remaining = total
    Application.ScreenUpdating = False

    If total >= PROGRESS_THRESHOLD Then
        Set docView = ActiveWindow.View
        oldViewType = docView.Type
        If oldViewType = wdPrintView Then
            docView.Type = wdNormalView
            changedView = True
        End If
        If docView.Type = wdNormalView Then
            oldPagination = Options.Pagination
            If oldPagination Then
                Options.Pagination = False
                changedPagination = True
            End If
        End If
    End If

    Set undo = Application.UndoRecord
    undo.StartCustomRecord "Convert LaTeX to Word Equations"
    undoOpen = True

    ' Word enables the LaTeX ribbon command when an equation is selected. Set
    ' the mode on the first new equation, before showing the modal progress window.
    Do While remaining > 0 And modeAnchor Is Nothing
        ConvertNext
    Loop

    If remaining > 0 And total >= PROGRESS_THRESHOLD And Application.Visible Then
        Set progress = New LatexBatchProgress
        progress.Show vbModal
    Else
        ConvertRemainingLatex Nothing
    End If
    failure = batchError
    GoTo Cleanup

Failed:
    failure = Err.Description

Cleanup:
    On Error Resume Next
    If Not progress Is Nothing Then Unload progress
    If restoreUnicode Then
        modeAnchor.Select
        Application.CommandBars.ExecuteMso UNICODE_FORMAT
    End If
    If undoOpen Then undo.EndCustomRecord
    If Len(failure) > 0 And edited And undoOpen Then
        Dim rolledBack As Boolean
        rolledBack = doc.Undo(1)
        If Not rolledBack Then failure = failure & vbCrLf & "Use Undo to review the partial conversion."
    End If
    If changedPagination Then Options.Pagination = oldPagination
    If changedView Then docView.Type = oldViewType
    If edited Then source.Select
    Application.ScreenUpdating = oldScreenUpdating
    On Error GoTo 0

    Erase matches
    Set work = Nothing
    Set modeAnchor = Nothing
    total = 0
    restoreUnicode = False
    edited = False
    batchError = vbNullString
    batchActive = False

    If Len(failure) > 0 Then MsgBox failure, vbExclamation, "Convert LaTeX"
End Sub

' Called by LatexBatchProgress.Activate. The modal window prevents edits to
' the selected document while DoEvents updates progress and dispatches messages.
Public Sub ConvertRemainingLatex(ByVal progress As LatexBatchProgress)
    On Error GoTo Failed
    Dim lastUpdate As Single
    lastUpdate = Timer
    If Not progress Is Nothing Then progress.Report total - remaining, total

    Do While remaining > 0
        ConvertNext
        If Not progress Is Nothing Then
            If Timer < lastUpdate Or Timer - lastUpdate >= 0.15 Or remaining = 0 Then
                progress.Report total - remaining, total
                DoEvents
                lastUpdate = Timer
            End If
        End If
    Loop
    Exit Sub

Failed:
    batchError = Err.Description
End Sub

Private Sub ConvertNext()
    With matches(remaining)
        Dim startAt As Long, endAt As Long
        startAt = baseStart + .Offset
        endAt = startAt + .Length
        work.SetRange startAt, endAt
        If checkExistingMath Then
            If work.OMaths.Count > 0 Then
                remaining = remaining - 1
                Exit Sub
            End If
        End If

        Dim prefix As String, suffix As String
        If .Display Then
            If startAt > 0 Then
                work.SetRange startAt - 1, startAt
                If work.Text <> vbCr And work.Text <> Chr$(7) Then prefix = vbCr
            End If
            If endAt < work.StoryLength Then
                work.SetRange endAt, endAt + 1
                If work.Text <> vbCr And work.Text <> Chr$(7) Then suffix = vbCr
            End If
            work.SetRange startAt, endAt
        End If

        work.Text = prefix & .Formula & suffix
        edited = True
        startAt = startAt + Len(prefix)
        work.SetRange startAt, startAt + Len(.Formula)

        Dim equationRange As Range
        Set equationRange = work.OMaths.Add(work)
        Dim equation As OMath
        Set equation = equationRange.OMaths(1)

        If modeAnchor Is Nothing Then
            Set modeAnchor = equationRange.Duplicate
            modeAnchor.Select
            restoreUnicode = Not Application.CommandBars.GetPressedMso(LATEX_FORMAT)
            If restoreUnicode Then Application.CommandBars.ExecuteMso LATEX_FORMAT
            If Not Application.CommandBars.GetPressedMso(LATEX_FORMAT) Then
                Err.Raise vbObjectError + 411, , "This version of Word could not select LaTeX equation input."
            End If
        End If

        equation.BuildUp
        If .Display Then
            If equation.Type <> wdOMathDisplay Then equation.Type = wdOMathDisplay
            equation.Justification = wdOMathJcCenter
            equation.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Else
            If equation.Type <> wdOMathInline Then equation.Type = wdOMathInline
        End If
    End With
    remaining = remaining - 1
End Sub

Private Sub ScanMatches(ByVal snapshot As String)
    ReDim matches(1 To 64)
    Dim segmentStart As Long, i As Long
    segmentStart = 1
    ' Word uses these characters for cell ends, section breaks and fields.
    ' A replacement must stay on one side of each boundary.
    For i = 1 To Len(snapshot)
        Select Case AscW(Mid$(snapshot, i, 1))
            Case 7, 12, 19, 20, 21
                ScanSegment Mid$(snapshot, segmentStart, i - segmentStart), segmentStart - 1
                segmentStart = i + 1
        End Select
    Next i
    ScanSegment Mid$(snapshot, segmentStart), segmentStart - 1
End Sub

Private Sub ScanSegment(ByVal segment As String, ByVal offset As Long)
    Dim i As Long, closing As Long, delimiterSize As Long
    Dim nextBracket As Long, nextParen As Long, nextDollar As Long
    Dim display As Boolean
    i = 1
    Do While i <= Len(segment)
        closing = 0
        display = False
        delimiterSize = 2
        Select Case Mid$(segment, i, 1)
            Case "\"
                Select Case Mid$(segment, i + 1, 1)
                    Case "["
                        display = True
                        closing = FindCloser(segment, "\]", i + 2, nextBracket)
                    Case "("
                        closing = FindCloser(segment, "\)", i + 2, nextParen)
                End Select
            Case "$"
                delimiterSize = 1
                If Mid$(segment, i + 1, 1) = "$" Then
                    Do While Mid$(segment, i + 1, 1) = "$"
                        i = i + 1
                    Loop
                Else
                    closing = FindCloser(segment, "$", i + 1, nextDollar)
                    If closing > 0 Then
                        If Mid$(segment, closing + 1, 1) = "$" Then closing = 0
                    End If
                End If
            Case Else
                delimiterSize = 1
        End Select

        If closing > 0 Then
            Dim formula As String
            formula = Mid$(segment, i + delimiterSize, closing - i - delimiterSize)
            formula = Replace(formula, vbCr, " ")
            formula = Replace(formula, vbLf, " ")
            formula = Replace(formula, vbTab, " ")
            formula = Trim$(Replace(formula, Chr$(11), " "))
            If Len(formula) > 0 Then
                total = total + 1
                If total > UBound(matches) Then ReDim Preserve matches(1 To UBound(matches) * 2)
                With matches(total)
                    .Offset = offset + i - 1
                    .Length = closing + delimiterSize - i
                    .Formula = formula
                    .Display = display
                End With
            End If
            i = closing + delimiterSize
        Else
            i = i + delimiterSize
        End If
    Loop
End Sub

Private Function FindCloser(ByVal segment As String, ByVal delimiter As String, _
                            ByVal startAt As Long, ByRef cached As Long) As Long
    ' Cache each delimiter's next match, including a failed search. A series
    ' of unmatched openers must not scan the same suffix for each opener.
    If cached = -1 Then Exit Function
    If cached < startAt Then
        cached = InStr(startAt, segment, delimiter, vbBinaryCompare)
        Do While cached > 0
            Dim preceding As Long
            preceding = cached - 1
            Do While preceding > 0
                If Mid$(segment, preceding, 1) <> "\" Then Exit Do
                preceding = preceding - 1
            Loop
            If (cached - 1 - preceding) Mod 2 = 0 Then Exit Do
            cached = InStr(cached + 1, segment, delimiter, vbBinaryCompare)
        Loop
        If cached = 0 Then
            cached = -1
            Exit Function
        End If
    End If
    FindCloser = cached
End Function
