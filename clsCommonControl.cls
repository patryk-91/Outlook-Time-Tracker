VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommonControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1
Public frm As frmDatePicker
Private Sub btn_Click()
    frm.SelectedDate = DateSerial(Year(frm.Label1.Caption), Month(frm.Label1.Caption), Me.btn.Caption)
    frm.Hide
End Sub
Private Sub btn_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        frm.Hide
    End If
End Sub
