Attribute VB_Name = "Task Mod"
Option Compare Database
Option Explicit


Public Function SetDueDateField(frm As Form)

    Dim StartDate: StartDate = frm("StartDate")
    frm("DueDate") = StartDate
    
End Function

Public Function SetDueTimeField(frm As Form)

    Dim StartTime: StartTime = frm("StartTime")
    If isFalse(StartTime) Then
        frm("DueTime") = Null
        Exit Function
    End If
    frm("DueTime") = TimeValue(DateAdd("n", 60, StartTime))
    
End Function

Public Function TaskCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function GetTimestamp(vDate, vTime)

    If IsNull(vDate) Then
        GetTimestamp = Null
        Exit Function
    End If
    
    If IsNull(vTime) Then
        GetTimestamp = DateValue(vDate)
        Exit Function
    End If
    
    
    GetTimestamp = DateValue(vDate) + TimeValue(vTime)
    
End Function

Public Function UpdateTaskNote(frm As Form)

    Dim TaskID
    TaskID = frm("TaskID")
    
    If isFalse(TaskID) Then Exit Function
    
    RunSQL "UPDATE tblTasks SET TaskNote = " & EscapeString(GetTaskNotes(TaskID)) & " WHERE TaskID = " & TaskID
    
End Function

''GetTaskNotes(17)
Public Function GetTaskNotes(TaskID) As String
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblGenericNotes WHERE TableName = 'tblTasks' And RecordID = " & TaskID & " ORDER BY [Timestamp]")
    
    Dim noteArr As New clsArray
    
    Do Until rs.EOF
        noteArr.Add rs.fields("NoteDescription")
        rs.MoveNext
    Loop
    
    If noteArr.Count > 0 Then GetTaskNotes = noteArr.JoinArr(" | ")

End Function

Public Function SetDueDate(frm As Form)

    Dim StartDate, DueDate
    StartDate = frm("StartDate")
    
    If isFalse(StartDate) Then
        frm("DueDate") = Null
        Exit Function
    End If
    
    DueDate = DateAdd("n", 30, StartDate)
    frm("DueDate") = DueDate
    
End Function

''GetTaskDateFormat(Forms("frmCustomDashboard")("subTasks").Form)
Public Function GetTaskDateFormat(frm As Form)

    Debug.Print frm("StartDate").Format
    
End Function

''ExportTasksToExcel(Forms("frmCustomDashboard"))
Public Function ExportTasksToExcel(frm As Form)
    
    ''Get all the tasks to export. Make a query for that so that we can export easily (qryTasksToExport)
    Dim sqlStr: sqlStr = "SELECT TaskDescription As Description, " & _
                                "Location, " & _
                                Esc("") & " As Private, " & _
                                "MemberName As [Attendee Name], " & _
                                "MemberPhoneNumber As [Attendee Phone Number], " & _
                                "MemberEmailAddress As [Attendee 1 Email Address], " & _
                                "MyPandaEmail As [My Panda Email], " & _
                                "Reminder, " & _
                                "MinutesBeforeReminder As [Minutes Before Reminder]" & _
                        "FROM qryTasks ORDER BY TaskID"
    
    Dim qDef As QueryDef: Set qDef = CurrentDb.QueryDefs("qryTasksToExport")
    qDef.sql = sqlStr: qDef.Close
    
    ''Title,Start Date,End Date,Location,Note (Leave the note blank for now)
    Dim fileName: fileName = GetExcelFileName & "Book1.xlsx"
    If IsFileOpen(fileName) Then
        MsgBox "The file: " & Esc(fileName) & " is open. Please close it first.", vbCritical + vbYesNo
        Exit Function
    End If
    
    If Dir(fileName) <> "" Then
        Kill fileName
    End If
    
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "qryTasksToExport", fileName
    
    ''FormatExcelFields fileName
    MsgBox "Tasks successfully exported at: " & Esc(fileName)
    
    CreateObject("Shell.Application").Open fileName
    
End Function

Public Function Sync_tblEventTimelines_with_tblTasks(Optional ShowMessage As Boolean = True)
    
    ''Select only those that has number of days indicated
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "qryEventTimelines"
          .AddFilter "Days <> 0 AND tblTasks.TaskID IS NULL"
          .fields = "tblEventTimelines.EventTimelineID,tblEventTimelines.PropertyListID,DateValue(EventTimelineDate) As DueDate," & _
            "TimeValue(EventTimelineDate) AS DueTime, Description AS TaskDescription, iif(IsChecked,""Completed"",""In Progress"") AS Status," & _
            "DateValue(ContractDate) AS StartDate, TimeValue(ContractDate) AS StartTime"
          .Joins.Add GenerateJoinObj("tblTasks", "EventTimelineID", , , "LEFT")
          sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "tblTasks"
          .fields = "EventTimelineID,PropertyListID,DueDate,DueTime,TaskDescription,Status,StartDate,StartTime"
          .InsertSQL = sqlStr
          .InsertFilterField = "EventTimelineID,PropertyListID,DueDate,DueTime,TaskDescription,Status,StartDate,StartTime"
          rowsAffected = .Run
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "qryEventTimelines"
          .AddFilter "Days <> 0"
          .fields = "tblEventTimelines.EventTimelineID,tblEventTimelines.PropertyListID,DateValue(EventTimelineDate) As DueDate," & _
            "TimeValue(EventTimelineDate) AS DueTime, Description AS TaskDescription, iif(IsChecked,""Completed"",""In Progress"") AS Status," & _
            "DateValue(ContractDate) AS StartDate, TimeValue(ContractDate) AS StartTime"
          sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "UPDATE"
        .Source = "tblTasks"
        .SetStatement = "tblTasks.DueDate = temp.DueDate," & _
            "tblTasks.DueTime = temp.DueTime," & _
            "tblTasks.TaskDescription = temp.TaskDescription," & _
            "tblTasks.Status = temp.Status," & _
            "tblTasks.StartDate = temp.StartDate," & _
            "tblTasks.StartTime = temp.StartTime"
        .Joins.Add GenerateJoinObj(sqlStr, "EventTimelineID", "temp")
        rowsAffected = .Run
    End With
    
    If ShowMessage Then MsgBox "Event timelines successfully synced with Tasks.."
    
    If IsFormOpen("mainTasks") Then
        Forms("mainTasks")("subform").Form.Requery
    End If
    
    If IsFormOpen("frmCustomDashboard") Then
        Forms("frmCustomDashboard")("subTasks").Form.Requery
    End If

End Function

Public Function ExportATaskToExcel(frm As Form, Optional TaskID)
    
    DoCmd.RunCommand acCmdSaveRecord
    
    If isFalse(TaskID) Then
        TaskID = frm("TaskID")
    End If
    
    If ExitIfTrue(isFalse(TaskID), "Task is empty..") Then Exit Function
    
    ''Get all the tasks to export. Make a query for that so that we can export easily (qryTasksToExport)
    Dim sqlStr: sqlStr = "SELECT DateValue(StartDate) As [Start Date], " & _
                                "Format(StartTime,""HH:mm"") As [Start Time], " & _
                                "DateValue(DueDate) As [End Date], " & _
                                "Format(DueTime,""HH:mm"") As [End Time], " & _
                                Esc("FALSE") & " As [All Day Event], " & _
                                "TaskDescription As Description, " & _
                                "Location, " & _
                                Esc("") & " As Private, " & _
                                "MemberName As [Attendee Name], " & _
                                "MemberPhoneNumber As [Attendee Phone Number], " & _
                                "MemberEmailAddress As [Attendee 1 Email Address], " & _
                                "MyPandaEmail As [My Panda Email], " & _
                                "Reminder, " & _
                                "MinutesBeforeReminder As [Minutes Before Reminder]" & _
                        "FROM qryTasks WHERE TaskID = " & TaskID & " ORDER BY TaskID"
    
    Dim db As Database: Set db = CurrentDb
    Dim qDef As QueryDef: Set qDef = db.QueryDefs("qryTasksToExport")
    qDef.sql = sqlStr: qDef.Close
    
    ''Title,Start Date,End Date,Location,Note (Leave the note blank for now)
    Dim fileName: fileName = GetExcelFileName & "Book1.xlsx"
'    If IsFileOpen(fileName) Then
'        MsgBox "The file: " & Esc(fileName) & " is open. Please close it first.", vbCritical + vbYesNo
'        Exit Function
'    End If
    
    ''Meaning the file is existing
    If Dir(fileName) <> "" Then
        Dim xl As Object
        Dim wb As Object
        Dim sht As Object
        
        Set xl = CreateObject("Excel.Application")
        Set wb = GetObject(fileName)
        
        xl.Visible = True
        wb.Activate
        wb.Windows(1).Visible = True
        
        Set sht = wb.Worksheets(1)
        
        Dim maxRow, maxCol, curRow
        maxRow = sht.UsedRange.Rows.Count
        maxCol = sht.UsedRange.Columns.Count
        curRow = maxRow + 1
        
        Dim rs As Recordset: Set rs = ReturnRecordset("qryTasksToExport")
        Do Until rs.EOF
            sht.cells(curRow, 1) = Format$(rs.fields(0), "Short Date")
            sht.cells(curRow, 2) = rs.fields(1)
            sht.cells(curRow, 3) = Format$(rs.fields(2), "Short Date")
            sht.cells(curRow, 4) = rs.fields(3)
            sht.cells(curRow, 5) = rs.fields(4)
            sht.cells(curRow, 6) = rs.fields(5)
            sht.cells(curRow, 7) = rs.fields(6)
            sht.cells(curRow, 8) = rs.fields(7)
            sht.cells(curRow, 9) = rs.fields(8)
            sht.cells(curRow, 10) = rs.fields(9)
            sht.cells(curRow, 11) = rs.fields(10)
            sht.cells(curRow, 12) = rs.fields(11)
            sht.cells(curRow, 13) = rs.fields(12)
            sht.cells(curRow, 14) = rs.fields(13)
            curRow = curRow + 1
            rs.MoveNext
        Loop
        
        wb.Save
        
    Else
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "qryTasksToExport", fileName
    End If
    
    ''FormatExcelFields fileName
    MsgBox "Task successfully exported at: " & Esc(fileName)
    
    ''CreateObject("Shell.Application").Open fileName
    
End Function

''FormatExcelFields(CurrentProject.Path & "\Files\Book1.xlsx")
Private Function FormatExcelFields(fileName)

    ''Autofit Pages, Format B & C Columns
    Dim xl As Object
    Dim Xw As Object
    Dim sht As Object
    
    Set xl = CreateObject("Excel.Application")
    Set Xw = GetObject(fileName)
    Set sht = Xw.Worksheets(1)
    
    xl.Visible = False
    Xw.Activate
    
    sht.Range("B:C").NumberFormat = "mm/dd/yyyy h:mm AM/PM"
    sht.Range("A:E").EntireColumn.AutoFit
    
    Xw.Save
    Xw.Close SaveChanges:=True
    xl.Quit
    Set Xw = Nothing
    Set xl = Nothing
    
End Function

Private Function GetExcelFileName()

    Dim xlDirectory
    xlDirectory = CurrentProject.Path & "\Files\"
    Select Case Environ("computername")
        Case "DESKTOP-CLLF13L":
            xlDirectory = "C:\Users\appli\OneDrive\MY PANDA REALTY\PANDA APP\GOOGLE RECORDS 2\EXCEL TO GOOGLE CALENDER\"
        Case "DESKTOP-3G3V8GO":
            xlDirectory = "C:\Users\Owner\OneDrive\GOOGLE RECORDS 2\EXCEL TO GOOGLE CALENDER\"
        Case Else:
            xlDirectory = CurrentProject.Path & "\Files\"
    End Select
    
    ''Check if the directory is existing, if not then create it.
    Dim strFolderExists
    strFolderExists = Dir(xlDirectory, vbDirectory)
    If strFolderExists = "" Then
        If Environ("computername") <> "DESKTOP-3G3V8GO" Then
            MsgBox EscapeString(xlDirectory) & " path does not exist."
            Exit Function
        Else
            MkDir xlDirectory
        End If
    End If
    
    GetExcelFileName = xlDirectory
    
End Function

Public Function TaskFormOnLoad(frm As Form)

    DefaultFormLoad frm, "TaskID", False
    frm("subNotes").Form("RecordID").ColumnHidden = True
    frm("subNotes").Form("TableName").ColumnHidden = True
    frmTasks_AttendeeID_SetRowSource frm
    
End Function

Public Function frmTasks_PropertListID_AfterUpdate(frm As Form)

    frmTasks_AttendeeID_SetRowSource frm
    
End Function

Public Function frmTasks_AttendeeID_SetRowSource(frm As Form)
    
    FixEntityMembers
    
    Dim sqlStr: sqlStr = "Select EntityMemberID,MemberName FROM qryPropertyEntityMembers WHERE EntityMemberID = 0"
    
    Dim PropertyListID: PropertyListID = frm("PropertyListID")
    
    If Not isFalse(PropertyListID) Then
        sqlStr = "Select EntityMemberID,MemberName FROM qryPropertyEntityMembers WHERE PropertyListID = " & PropertyListID & _
            " ORDER BY MemberName"
    End If
    
    frm("AttendeeID").rowSource = sqlStr
    
End Function

Public Function ImportNoteAsTask(frm As Form, EntityType)

    Dim PropertyListID
    PropertyListID = frm("PropertyListID")
    
    If ExitIfTrue(isFalse(PropertyListID), "Property is empty..") Then Exit Function
    
    Dim Note
    Note = frm("sub" & EntityType & "Notes").Form("Note")
    If ExitIfTrue(isFalse(Note), "Note is empty..") Then Exit Function
    
    Dim EntityMemberID
    If EntityType <> "Seller" Then EntityMemberID = frm("sub" & EntityType & "Members").Form("EntityMemberID")
    
    DoCmd.OpenForm "frmTasks", , , , acFormAdd
    
    Forms("frmTasks").PropertyListID = PropertyListID
    Forms("frmTasks").TaskDescription = Note
    Forms("frmTasks").AttendeeID = EntityMemberID
    Set frm = Forms("frmTasks")
    frmTasks_AttendeeID_SetRowSource frm
    
End Function

Public Function contUncompletedTasksStatus(frm As Form)
    
    Dim Status: Status = frm("Status")
    
    Dim EventTimelineID: EventTimelineID = frm("EventTimelineID")
    
    Dim IsChecked: IsChecked = "0"
    If Status = "Completed" Then
        IsChecked = "-1"
    End If
    
    If Not isFalse(EventTimelineID) Then
        RunSQL "UPDATE tblEventTimelines SET IsChecked = " & IsChecked & " WHERE EventTimelineID = " & EventTimelineID
    End If

    If Status = "Completed" Then frm.Requery
    
End Function

Public Function frmTasksUnload(frm As Form)
    
    UpdateTaskNote frm

    If IsFormOpen("frmCustomDashboard") Then
        Forms("frmCustomDashboard")("subTasks").Form.Requery
    End If
    
    If IsFormOpen("frmTaskCalendar") Then
        Set frm = Forms("frmTaskCalendar")
        SetCalendar frm
    End If

End Function

Public Function SendTaskToEmail(frm As Form)

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
    Set olMail = olApp.CreateItem(0)
    With olMail
        
        .Display
        .To = "admin@mypandarealty.com.au"
        .Subject = Subject
        .HTMLBody = TaskDescription
        '.Send
    End With
    
    
End Function

Public Function SyncScheduledViewingDateToCalendar(Optional RefreshForm As Boolean = False)
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "qryScheduledViewingDates"
            .fields = "TaskDescription,PropertyListID,ScheduledViewingDate As StartDate," & _
            "ScheduledViewingDate As DueDate," & Esc("Not Started") & " AS Status"
          sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = sqlStr
          .AddFilter "TaskID IS NULL"
          .fields = "temp.TaskDescription,temp.PropertyListID,temp.StartDate,temp.DueDate,temp.Status"
          .Joins.Add GenerateJoinObj("tblTasks", "TaskDescription", , , "LEFT")
          .SourceAlias = "temp"
          sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "tblTasks"
          .fields = "TaskDescription,PropertyListID,StartDate,DueDate,Status"
          .InsertSQL = sqlStr
          .InsertFilterField = "TaskDescription,PropertyListID,StartDate,DueDate,Status"
          rowsAffected = .Run
    End With
    
    If RefreshForm Then
        If IsFormOpen("frmCustomDashboard") Then
            Forms("frmCustomDashboard")("subTasks").Form.Requery
        End If
    End If
    
End Function
