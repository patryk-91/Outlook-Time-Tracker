Attribute VB_Name = "modCalendar"
Public Const SEARCH_TAG As String = "last7days"
Public IsSearchCompleted As Boolean
Public Sub Appointment_Add(Optional topic As String)
    Dim now_time As Date
    Dim last_appointment As Object
    now_time = Now
    Set last_appointment = GetLastAppointment
    If Not last_appointment Is Nothing Then
        last_subject = last_appointment.Subject
    End If
    If Len(topic) = 0 Then
        With New frmNewAppointment
            .Caption = "New Appointment"
            .cmbTopic.List = GetLastAppointments
            .cmbTopic.Value = last_subject
            .cmbTopic.SelectionMargin = 0
            .cmbTopic.SelLength = Len(.cmbTopic.Value)
            .Show
            If .IsCancelled Then
                Exit Sub
            End If
            new_subject = .cmbTopic.Value
        End With
    Else
        new_subject = topic
    End If
    With CreateItem(olAppointmentItem)
        .Subject = new_subject
        .Start = now_time
        .Duration = 5
        .ReminderSet = False
        .Save
    End With
    If Not last_appointment Is Nothing Then
        Call Appointment_Align(last_appointment, now_time)
    End If
End Sub
Public Sub Appointment_Align(appointment As Object, end_time As Date)
    appointment.End = end_time
    appointment.Save
End Sub
Public Function GetLastAppointment() As Object
    Dim ns As Object
    Dim calendar_folder As Object
    Dim calendar_appointments As Object
    Dim date_appointments As Object
    Dim date_criteria As String
    Dim appointment As Object
    Set ns = Application.GetNamespace("MAPI")
    Set calendar_folder = ns.GetDefaultFolder(olFolderCalendar)
    Set calendar_appointments = calendar_folder.Items
        calendar_appointments.IncludeRecurrences = True
        calendar_appointments.Sort "[Start]", False
        date_criteria = "[Start]>'" & Format(Date, "dd/mm/yyyy") & " 12:00 AM" & "' and [Start] <= '" & Format(Now, "dd/mm/yyyy hh:nn AM/PM") & "'"
    Set date_appointments = calendar_appointments.Restrict(date_criteria)
    For Each appointment In date_appointments
        Set GetLastAppointment = appointment
    Next
End Function
Public Function GetLastAppointments() As Variant
    Dim ns As Object
    Dim search_table As Object
    Dim calendar_folder As Object
    Dim dasl As String
    Set ns = Application.GetNamespace("MAPI")
    Set calendar_folder = ns.GetDefaultFolder(olFolderCalendar)
    dasl = "%last7days(" & """" & "urn:schemas:calendar:dtstart" & """" & ")%"
    IsSearchCompleted = False
    Set search_result = Application.AdvancedSearch("'" & calendar_folder.FolderPath & "'", dasl, False, SEARCH_TAG)
    While Not IsSearchCompleted
        DoEvents
    Wend
    Set search_table = search_result.GetTable
    GetLastAppointments = search_table.GetArray(search_table.GetRowCount)
End Function

