Option Explicit

Private message As MSForms.Label
Private started As Boolean

Private Sub UserForm_Initialize()
    Me.Caption = "Converting LaTeX equations"
    Me.Width = 290
    Me.Height = 74
    Set message = Me.Controls.Add("Forms.Label.1", "ProgressMessage", True)
    message.Move 14, 14, 252, 16
    message.Caption = "Preparing equations..."
End Sub

Private Sub UserForm_Activate()
    If started Then Exit Sub
    started = True
    WordLatex.ConvertRemainingLatex Me
    Me.Hide
End Sub

Public Sub Report(ByVal completed As Long, ByVal total As Long)
    message.Caption = completed & " of " & total & " equations processed"
    Me.Repaint
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Cancel = True
End Sub
