VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Application_AdvancedSearchComplete(ByVal SearchObject As Search)
    If SearchObject.Tag = SEARCH_TAG Then
        IsSearchCompleted = True
    End If
End Sub

Private Sub Application_Quit()
    Dim last_appointment As Object
    Set last_appoiment = GetLastAppointment
    If MsgBox("Do you want to align preceding appointment to now?", vbYesNo + vbExclamation + vbDefaultButton2, "Outlook Close") = vbYes Then
        Call Appointment_Align(last_appointment, Now)
    End If
End Sub

Private Sub Application_Reminder(ByVal Item As Object)
    Dim last_appointment As Object
    If Not TypeOf Item Is AppointmentItem Then
        Exit Sub
    End If
    Set last_appointment = GetLastAppointment
    If MsgBox("Do you want to align preceding appointment to " & Item.Subject & "?", vbYesNo + vbExclamation + vbDefaultButton2, "Upcoming Appointment") = vbYes Then
        Call Appointment_Align(last_appointment, Item.Start)
    End If
End Sub

Private Sub Application_Startup()
    Call Appointment_Add("Catch-up on my correspondence")
End Sub
