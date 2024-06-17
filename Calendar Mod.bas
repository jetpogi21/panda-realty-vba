Attribute VB_Name = "Calendar Mod"
Option Compare Database
Option Explicit

Public Function AddTaskToCalendar(frm As Form)

    Dim TaskID, TaskDescription, PropertyListID, StartDate, DueDate, Status, Recipient, StreetAddress
    TaskID = frm("TaskID")
    TaskDescription = frm("TaskDescription")
    PropertyListID = frm("PropertyListID")
    StartDate = frm("StartDate")
    DueDate = frm("DueDate")
    Status = frm("Status")
    Recipient = frm("Recipient")
    
    StreetAddress = GetPropertyAddress(PropertyListID)
    
    If ExitIfTrue(IsNull(Recipient), "Please select a valid recipient..") Then Exit Function
    
    If Recipient = "Both" Then Recipient = "Richard and Maria"
    
    Dim TruncatedTask, Subject
    TruncatedTask = left(TaskDescription, "25")
    Subject = "Note for " & Recipient & "-" & TruncatedTask
    
    If Not IsNull(StreetAddress) Then Subject = Subject & " (" & StreetAddress & ")"
    
    Dim olApp As Object, olMail As Object
    
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(1)
    With olMail
        .Display
        .Location = StreetAddress
        .AllDayEvent = True
        .Subject = Subject
        .Body = TaskDescription
        '.Save
        '.Send
    End With
    
    'MsgBox "Task successfully added to calendar.."
    
End Function
