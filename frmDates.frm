VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDates 
   Caption         =   "Select Dates"
   ClientHeight    =   2950
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   4480
   OleObjectBlob   =   "frmDates.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cancelled  As Boolean
Private Sub btnCancel_Click()
    Call OnCancel
End Sub
Public Property Get IsCancelled() As Boolean
    IsCancelled = cancelled
End Property
Private Sub OnCancel()
    cancelled = True
    Hide
End Sub
Private Sub btnCancel_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Call OnCancel
    End If
End Sub
Private Sub btnOK_Click()
    Hide
End Sub
Private Sub lblFrom_Click()
    With frmDatePicker
        .Caption = "From"
        .Show
        If .IsCancelled Then
            Exit Sub
        End If
        Me.lblFrom.Caption = Format(.SelectedDate, "dd/mm/yyyy")
    End With
End Sub
Private Sub lblFrom_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblFrom.ForeColor = vbBlue
End Sub
Private Sub lblTo_Click()
    With frmDatePicker
        .Caption = "To"
        .Show
        If .IsCancelled Then
            Exit Sub
        End If
        Me.lblTo.Caption = Format(.SelectedDate, "dd/mm/yyyy")
    End With
End Sub
Private Sub lblTo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblTo.ForeColor = vbBlue
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblFrom.ForeColor = vbBlack
    lblTo.ForeColor = vbBlack
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        Call OnCancel
    End If
End Sub

