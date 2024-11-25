Attribute VB_Name = "modCalendar"
Public Const SEARCH_TAG As String = "last7days"
Public IsSearchCompleted As Boolean
Public Sub Start()
    Call Appointment_Add
End Sub
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
Public Sub Appointments_Export()
    Dim xlApp As Object
    Dim wb As Object
    Dim ws As Object
    Dim olNS As NameSpace
    Dim olCalendar As Folder
    Dim itms As Items
    Dim filtered_items As Items
    Dim olApt As AppointmentItem
    Dim txtFrom As String
    Dim txtTo As String
    With New frmDates
        .lblFrom.Caption = Format(Date - 7, "dd/mm/yyyy")
        .lblTo.Caption = Format(Date, "dd/mm/yyyy")
        .Show
        If .IsCancelled Then
            Exit Sub
        End If
        txtFrom = .lblFrom.Caption
        txtTo = .lblTo.Caption
    End With
On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If xlApp Is Nothing Then Set xlApp = CreateObject("Excel.Application")
    xlApp.screenupdating = False
On Error GoTo 0
    Set wb = xlApp.workbooks.Add
    Set ws = wb.sheets(1)
    Set olNS = Application.GetNamespace("MAPI")
    Set olCalendar = olNS.GetDefaultFolder(olFolderCalendar)
    'browse calendar
    Set itms = olCalendar.Items
        itms.Sort "[Start]", False
        itms.IncludeRecurrences = True
        filter_string = "[Start] >= '" & txtFrom & " 00:00' and [End] <= '" & txtTo & " 00:00'"
    Set filtered_items = itms.Restrict(filter_string)
    'write results
    ws.Range("A1") = "Topic"
    ws.Range("B1") = "Start"
    ws.Range("C1") = "Duration [min]"
    ws.Range("D1") = "End"
    NextRow = 2
    For Each olApt In filtered_items
        ws.Range("A" & NextRow).Value = olApt.ConversationTopic
        ws.Range("B" & NextRow).Value = olApt.Start
        ws.Range("C" & NextRow).Value = olApt.Duration
        ws.Range("D" & NextRow).Value = olApt.End
        NextRow = NextRow + 1
    Next
    If NextRow > 2 Then
        With ws.Range("A1:D" & NextRow - 1)
            .Borders.Weight = xlThin
            .Columns.AutoFit
        End With
        xlApp.Visible = True
        xlApp.screenupdating = True
        MsgBox "Exported successfully", vbInformation, "Time Tracker"
    Else
        wb.Close False
        MsgBox "No appointments found", vbExclamation
    End If
End Sub
