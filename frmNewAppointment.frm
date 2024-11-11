VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNewAppointment 
   Caption         =   "New Appointment"
   ClientHeight    =   2142
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   5880
   OleObjectBlob   =   "frmNewAppointment.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNewAppointment"
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
    hide
End Sub
Private Sub btnOK_Click()
    hide
End Sub
Private Sub cmbTopic_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Call OnCancel
    End If
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        Call OnCancel
    End If
End Sub
