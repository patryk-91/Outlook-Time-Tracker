VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDatePicker 
   ClientHeight    =   2490
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   2850
   OleObjectBlob   =   "frmDatePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SelectedDate As Date
Dim mColButtons As New Collection
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
Private Sub SpinButton1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
        Call OnCancel
    End If
End Sub
Private Sub SpinButton1_SpinDown()
    Me.Label1 = Format(DateAdd("m", -1, Me.Label1), "Mmmm yyyy")
    Call Days_Repaint
End Sub
Private Sub SpinButton1_SpinUp()
    Me.Label1 = Format(DateAdd("m", 1, Me.Label1), "Mmmm yyyy")
    Call Days_Repaint
End Sub
Private Sub UserForm_Initialize()
    Dim btnEvent As clsCommonControl
    Dim ctl As MSForms.Control
        Me.Label1 = Format(Date, "Mmmm yyyy")
    For Each ctl In Me.Controls
        If TypeName(ctl) = "CommandButton" Then
            Set btnEvent = New clsCommonControl
            Set btnEvent.btn = ctl
            Set btnEvent.frm = Me
            mColButtons.Add btnEvent
        End If
    Next ctl
    Call Days_Repaint
End Sub
Private Sub Days_Repaint()
    Dim ctl As MSForms.CommandButton
    For i = 1 To 31
        Set ctl = mColButtons.Item(i).btn
        If i <= Day(DateSerial(Year(Me.Label1), Month(Me.Label1) + 1, 0)) Then
            ctl.Caption = i
            ctl.Visible = True
        Else
            ctl.Caption = ""
            ctl.Visible = False
        End If
    Next i
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        Call OnCancel
    End If
End Sub
