Attribute VB_Name = "TaskCalendar Mod"
Option Compare Database
Option Explicit

Public Function TaskCalendarCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function OpenTaskForm(frm As Form)

    Dim TaskID: TaskID = frm("TaskID")
    If isFalse(TaskID) Then Exit Function
    
    DoCmd.OpenForm "frmTasks", , , "TaskID = " & TaskID
    
End Function

Public Function SetCalendar(frm As Form)
    
    Dim txtmonth: txtmonth = frm("txtmonth")
    Dim txtyear: txtyear = frm("txtyear")
    
    If ExitIfTrue(isFalse(txtmonth), "Month Field should not be blank") Then Exit Function
    If ExitIfTrue(isFalse(txtyear), "Year Field should not be blank") Then Exit Function
    
    frm("txtMonth").SetFocus
    
    GenerateCalendar frm, txtmonth, txtyear
    
End Function

Public Function TimerRun(frm As Form)
    MsgBox "I will run every " & frm.TimerInterval / 60 & " minutes."
End Function

Public Function txtTimerAfterUpdate(frm As Form)
    
    Dim txtTimer: txtTimer = frm("txtTimer")
    ''Update the settings
    RunSQL "UPDATE tblApplicationSettings SET ApplicationSettingValue = " & EscapeString(txtTimer) & " WHERE ApplicationSettingName = " & EscapeString("CalendarTimer")
    frm.TimerInterval = txtTimer * 60 * 1000

End Function

Public Function SetCalendarTimer(frm As Form)
    
    ''minutes in millisecond
    Dim CalendarTimer
    CalendarTimer = ELookup("tblApplicationSettings", "ApplicationSettingName = " & EscapeString("CalendarTimer"), "ApplicationSettingValue")
    frm("txtTimer") = CalendarTimer
    frm.TimerInterval = CalendarTimer * 60 * 1000
    
End Function

Private Function NB_DAYS(date_month, date_year)
    NB_DAYS = Day(DateSerial(date_year, date_month + 1, 1) - 1)
End Function

Private Function getWeekOfMonth(testDate) As Integer
    getWeekOfMonth = CInt(Format(testDate, "ww")) - CInt(Format(Format(testDate, "mm") & "/01/" & Format(testDate, "yyyy"), "ww")) + 1
End Function

Private Sub GenerateCalendar(frm As Form, txtmonth, txtyear)

    Dim Days As Integer
    Dim weekOfMonth As Integer
    Dim i As Integer
    Dim current_date As Date
    Dim v_weekday As Integer
    
    ClearLabels frm
    
    Days = NB_DAYS(txtmonth, txtyear)
    
    For i = 1 To Days
        current_date = DateSerial(txtyear, txtmonth, i)
        v_weekday = weekday(current_date)
        weekOfMonth = getWeekOfMonth(current_date)
        frm.Controls("lbl" & v_weekday & weekOfMonth).Caption = i
        frm.Controls("box" & v_weekday & weekOfMonth).Visible = True
    Next i
    
    FillBoxes frm, txtmonth, txtyear
    
End Sub

Private Sub FillBoxes(frm, txtmonth, txtyear)

    'Establish recordset
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblTasks WHERE Month(StartDate) = " & txtmonth & " AND Year(StartDate) = " & txtyear & " ORDER BY StartDate")
    
    Dim v_weekday As Integer
    Dim weekOfMonth As Integer
    Dim StartDate, sqlStr, boxName
    
    Do Until rs.EOF
    
        StartDate = DateValue(rs.fields("StartDate"))
        v_weekday = weekday(StartDate)
        weekOfMonth = getWeekOfMonth(StartDate)
        
        sqlStr = "SELECT * FROM tblTasks WHERE DateValue(StartDate) = #" & StartDate & "# ORDER BY StartDate"
        boxName = "box" & v_weekday & weekOfMonth
        
        SetQueryDef sqlStr, boxName, frm(boxName).Form
        ''frm.Controls("box" & v_weekday & weekOfMonth).RowSource = "SELECT TaskID, TaskDescription, Format$([StartDate],""hh:nn"") As StartTime FROM tblTasks WHERE StartDate = #" & StartDate & "# ORDER BY StartDate"
        ''Debug.Print frm.Controls("box" & v_weekday & weekOfMonth).ColumnWidths, frm.Controls("box" & v_weekday & weekOfMonth).ColumnCount
        rs.MoveNext
        
    Loop
    
    rs.Close
        
End Sub

Private Function SetQueryDef(sqlStr, boxName, frm As Form)

''Error 3048
    'On Error GoTo Err_Handler:
    Dim db As Database
    Set db = CurrentDb
    Dim qDef As QueryDef
    Set qDef = db.QueryDefs("qry" & boxName)
    qDef.sql = sqlStr
    qDef.Close
    db.Close
    frm.RecordSource = "qry" & boxName
'    frm.RecordSource = sqlStr
'    ''db.Close
'    Exit Function
'Err_Handler:
'    If Err.Number = 3048 Then
'        MsgBox "Please close the property form", vbCritical, "Resource exceeded."
'        DoCmd.CancelEvent
'        Exit Function
'    End If
    
End Function

Private Function ClearLabels(frm As Form)

    Dim col As Integer
    Dim row As Integer
    
    For col = 1 To 7
        For row = 1 To 6
            If col > 2 And row = 6 Then
                Exit For
            End If
            ''main labels+
            frm.Controls("lbl" & col & row).Caption = ""
            'sums'
            frm.Controls("sum_" & col & row).Caption = ""
            frm.Controls("sum_" & col & row).BackStyle = 0
            ''Main control
            With frm.Controls("box" & col & row)
            
                SetQueryDef "SELECT * FROM tblTasks WHERE TaskID = 0", "box" & col & row, .Form
                ''.RowSource = ""
                ''.Locked = False
                '.ColumnCount = 3
                ''.ColumnWidths = "0;1.5"";0.2"""
                .Visible = False
                ''HideControl frm.Controls("box" & col & row)
            End With
        Next row
    Next col

End Function

Private Function HideControl(ctl As Object)

    On Error Resume Next
    ctl.Visible = False
    
End Function


Public Function CopyBoxReport()

    ''Base name of the report box
    Dim col, row

    For col = 1 To 7
        For row = 1 To 6
            If col > 2 And row = 6 Then
                Exit For
            End If
            DoCmd.CopyObject , "box" & col & row, acForm, "box"
        Next row
    Next col

End Function

'Public Function CopyBoxQuery()
'
'    ''Base name of the report box
'    Dim col, row
'
'    For col = 1 To 7
'        For row = 1 To 6
'            If col > 2 And row = 6 Then
'                Exit For
'            End If
'            DoCmd.CopyObject , "qryBox" & col & row, acQuery, "qryBox"
'        Next row
'    Next col
'
'End Function

''SetBoxRecordsource(Forms("frmTaskCalendar"))
Public Function SetBoxRecordsource(frm As Form)

    Dim ctl As Control, boxName
    
    For Each ctl In frm.Controls
        
        If ctl.ControlType = 112 Then
            boxName = ctl.Name
            ctl.Visible = False
            'ctl.SourceObject = boxName
            ctl.Form.RecordSource = "qry" & boxName
        End If
        
    Next ctl

End Function

''GetColorCodes(Forms("frmTaskCalendar_v2"))
Public Function GetColorCodes(frm As Form)
    
    Dim item, lblArr As New clsArray: lblArr.arr = "NotStarted,InProgress,Completed"
    
    ''ForeColor & BackColor
    For Each item In lblArr.arr
        Debug.Print item, frm("lbl" & item).ForeColor, frm("lbl" & item).BackColor
    Next item
     
End Function


