Attribute VB_Name = "Dashboard Mod"
Option Compare Database
Option Explicit
Public isImported As Boolean

Public Function QuitApp()
    DoCmd.Quit
End Function

Public Function LogOutUser()
    If MsgBox("Are you sure you want to log off?", vbCritical + vbYesNo, "Log-off Prompt") = vbYes Then
    
        Dim frm As Form
        For Each frm In Application.Forms
            If frm.Name <> "frmCustomDashboard" And CurrentProject.AllForms(frm.Name).IsLoaded Then
                DoCmd.Close acForm, frm.Name, acSaveNo
            End If
        Next frm
        
        g_UserID = Null
        
        BackupBackendFile
        DoCmd.Quit
        ''DoCmd.OpenForm "frmLogin"
        ''DoCmd.Close acForm, "frmCustomDashboard", acSaveNo
    Else
        DoCmd.CancelEvent
    End If
End Function

Public Function LinkBackendTables(Optional forceLink As Boolean = False)
    
    If Environ("computername") <> "DESKTOP-3G3V8GO" Or forceLink Then
        RunOneTimeFixes True
        DeleteLocalTableIfLinkedTable
        LinkTheTables
        RunOneTimeFixes
    End If
    
    DoCmd.OpenForm "frmCustomDashboard"
    
    If IsFormOpen("frmLinkedTables") Then
        DoCmd.Close acForm, "frmLinkedTables", acSaveNo
    End If
    
End Function

Public Function DashboardLoad(frm As Form)
    
    g_UserID = 1
    
    If isFalse(g_UserID) Then Exit Function
    
    ''Remove the uncompleted tasks
    frm("subTasks").SourceObject = ""
    
    ''RunOneTimeFixes
    '''Link the tables
'    If Environ("computername") <> "DESKTOP-3G3V8GO" Then DeleteLocalTableIfLinkedTable
'    If Environ("computername") <> "DESKTOP-3G3V8GO" Then RunOneTimeFixes
'    If Environ("computername") <> "DESKTOP-3G3V8GO" Then LinkTheTables

    frm("subTasks").SourceObject = "contUncompletedTasks"
    
    UpdateEntitiesIsSeller
    FixEntityMembers
    InsertAnyMissingPropertyEntitiesFromOwn
    SyncScheduledViewingDateToCalendar

    Dim UserID, sqlObj As New clsSQL
    UserID = g_UserID
    
    Dim rs As Recordset, rs2 As Recordset
    ''SELECT STATEMENT
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblUsers"
        .AddFilter "UserID = " & g_UserID
        Set rs = .Recordset
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "qryUserUserGroups"
        .AddFilter "UserID = " & g_UserID
        Set rs2 = .Recordset
    End With
    
    Dim lblString As String
    If Not rs.EOF Then
        lblString = rs.fields("UserName") & " | "
        Do Until rs2.EOF
            lblString = lblString & rs2.fields("UserGroup") & " | "
            rs2.MoveNext
        Loop
    End If
    
    frm.lblLoginInfo.Caption = lblString & " Active Since: " & Now()
    
    FilterDashboardMenu frm
    
    If Environ("computername") <> "DESKTOP-3G3V8GO" Then Shell ("OUTLOOK")
    
    ''FilterDashboardReports frm

End Function

Private Sub UpdateEntitiesIsSeller()

    RunSQL "UPDATE tblEntities SET IsSeller = -1 WHERE EntityCategoryID = 2"
    
End Sub

Private Sub DeleteTable(LinkedTableName)

On Error GoTo ErrHandler:

    DoCmd.DeleteObject acTable, LinkedTableName
    
ErrHandler:
    Exit Sub
    
End Sub

Public Function DeleteLocalTableIfLinkedTable()
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblLinkedTables ORDER BY Timestamp ASC")
    
    Do Until rs.EOF
    
        On Error Resume Next
        DeleteTable rs.fields("LinkedTableName")
        rs.MoveNext
        
    Loop
    
End Function

Public Function LinkTheTables()
    
    Dim ProjectPath, filePath
    
    ProjectPath = CurrentProject.Path
    
    If Environ("computername") <> "DESKTOP-3G3V8GO" Then
        ProjectPath = "Z:\MY PANDA APP"
        If Not DirectoryExists(ProjectPath) Then
            ProjectPath = "\\TRUENAS\database\MY PANDA APP"
            If Not DirectoryExists(ProjectPath) Then
                MsgBox "The database tables can't be linked to the backend file. The app will exit.", vbCritical
                DoCmd.Quit
                Exit Function
            End If
        End If
    End If
    filePath = ProjectPath & "\PTS Backend.accdb"
    
    ''AlterBackendTable filePath
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblLinkedTables ORDER BY Timestamp ASC")
    
    Do Until rs.EOF
        
        DeleteTable rs.fields("LinkedTableName")
            
        DoCmd.TransferDatabase TransferType:=acLink, _
            DatabaseType:="Microsoft Access", _
            DatabaseName:=filePath, _
            ObjectType:=acTable, _
            Source:=rs.fields("LinkedTableName"), _
            Destination:=rs.fields("LinkedTableName")
        
        rs.MoveNext
        
    Loop
    
    isImported = True
    
End Function

Private Sub FilterDashboardMenu(frm As Form)

    Dim rs As Recordset
    
    If isPresent("qryUserUserGroups", "UserID = " & g_UserID & " And UserGroup = " & EscapeString("Administrator")) Then
        Set rs = ReturnRecordset("SELECT * FROM tblMainMenus ORDER BY MenuOrder")
    Else
        Set rs = ReturnRecordset("select * from tblMainMenus where MainMenuID IN(select MainMenuID from qryUserGroupButtons where UserID = " & g_UserID & " GROUP BY MainMenuID) ORDER BY MenuOrder")
    End If
    
    Dim i
    For i = 0 To 17
    
        If rs.EOF Then
            frm("cmd" & i).Visible = False
        Else
            frm("cmd" & i).Visible = True
            frm("cmd" & i).Caption = rs.fields("MenuCaption")
            frm("cmd" & i).OnClick = "=DoOpenForm(" & EscapeString(rs.fields("FormName")) & ",Null,False)"
            
            ''Override the "Web Search" Menu caption
            If frm("cmd" & i).Caption = "Web Search" Then
                frm("cmd" & i).OnClick = "=OpenGoogle()"
            End If
            
            If frm("cmd" & i).Caption = "Bulk Email" Then
                frm("cmd" & i).OnClick = "=OpenMainPropertyBulkEmail([Form])"
            End If
            
            rs.MoveNext
        End If
    
    Next i

End Sub

Private Sub FilterDashboardReports(frm As Form)

    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("Select * From tblCustomReports ORDER BY ReportName ASC")

    Dim i
    For i = 0 To 9
    
        If rs.EOF Then
            frm("rpt" & i).Visible = False
        Else
            frm("rpt" & i).Visible = True
            frm("rpt" & i).Caption = rs.fields("ReportName")
            frm("rpt" & i).OnClick = "=OpenCustomReportForm(" & EscapeString(rs.fields("FilterFormName")) & ")"
            rs.MoveNext
        End If
    
    Next i

End Sub


'Public Sub CreateCustomDashboard()
'
'    ''Initialize x and y axis
'    x = 200: y = 200
'    Set frm = CreateForm
'    SetFormProperties
'    InsertLogo
'    InsertDashboardHeader
'    InsertDashboardSubHeader
'    InsertDashboardMenu
'
'End Sub
'
'
'Private Sub SetFormProperties()
'
'    With frm
'        .RecordSelectors = 0
'        .CloseButton = 0
'        .NavigationButtons = 0
'        .ScrollBars = 0
'        .Caption = "Main Menu"
'        .Picture = "BackgroundCropped"
'        .PictureSizeMode = 1
'    End With
'
'End Sub
'
'Private Sub InsertLogo()
'
'    Dim ctl As Control
'    Set ctl = CreateControl(frm.Name, acImage, , , , x, y, 5000, 2500)
'    ctl.Name = "imgLogo"
'    ctl.Picture = "Logo"
'
'End Sub
'
'
'Private Sub InsertDashboardHeader()
'
'    x = frm("imgLogo").Left + frm("imgLogo").Width + 200
'    Dim ctl As Control
'    Set ctl = CreateControl(frm.Name, acLabel, , , , x, y, 9000, 660)
'    ctl.Name = "lblHeader"
'    ctl.Caption = "EPC INVENTORY AND ASSET TRACKING"
'    ctl.FontName = "Segoe UI Black"
'    ctl.ForeColor = RGB(254, 254, 254)
'    ctl.FontSize = 22
'
'End Sub
'
'Private Sub InsertDashboardSubHeader()
'
'    x = frm("imgLogo").Left + frm("imgLogo").Width + 200
'    y = frm("lblHeader").Top + frm("lblHeader").height
'    Dim ctl As Control
'    Set ctl = CreateControl(frm.Name, acLabel, , , , x, y, frm("lblHeader").Width, 450)
'    ctl.Caption = "Welcome. To begin, select an option below."
'    ctl.FontName = "Segoe UI Black"
'    ctl.ForeColor = RGB(254, 254, 254)
'    ctl.FontSize = 14
'
'End Sub
'
'Private Sub InsertDashboardMenu()
'
'    Dim proportionArr As New clsArray, controlArr As New clsArray, proportionTotal, totalWidth, colSpaceWidth, i, proportion
'
'    y = frm("imgLogo").Top + frm("imgLogo").height + 500
'    x = frm("imgLogo").Left
'
'    colSpaceWidth = 200
'    totalWidth = 7000
'
'    Dim ctl As Control
'    For i = 0 To 11
'         Set ctl = CreateControl(frm.Name, acCommandButton, , , , 0, 0) ''Button Portion
'         ctl.Name = "cmd" & i
'         SetControlProperties ctl
'    Next i
'
'    ''Render the Filter buttons
'    ''Filter and Clear
'    proportionArr.Arr = "4,4,4"
'    controlArr.Arr = "cmd0,cmd1,cmd2"
'    proportionTotal = GetProportionTotal(proportionArr)
'
'    For i = 0 To proportionArr.Count - 1
'
'        proportion = CDbl(proportionArr.Arr(i)) / proportionTotal
'        frm(controlArr.Arr(i)).Left = x
'        frm(controlArr.Arr(i)).Top = y
'        frm(controlArr.Arr(i)).Width = (totalWidth - ((proportionArr.Count - 1) * colSpaceWidth * 2)) * proportion
'        frm(controlArr.Arr(i)).HorizontalAnchor = acHorizontalAnchorLeft
'
'        x = x + (colSpaceWidth * 2) + frm(controlArr.Arr(i)).Width
'
'    Next i
'
'    y = y + frm("cmd0").height + 200
'    x = frm("imgLogo").Left
'
'    controlArr.Arr = "cmd3,cmd4,cmd5"
'
'    For i = 0 To proportionArr.Count - 1
'
'        proportion = CDbl(proportionArr.Arr(i)) / proportionTotal
'        frm(controlArr.Arr(i)).Left = x
'        frm(controlArr.Arr(i)).Top = y
'        frm(controlArr.Arr(i)).Width = (totalWidth - ((proportionArr.Count - 1) * colSpaceWidth * 2)) * proportion
'        frm(controlArr.Arr(i)).HorizontalAnchor = acHorizontalAnchorLeft
'
'        x = x + (colSpaceWidth * 2) + frm(controlArr.Arr(i)).Width
'
'    Next i
'
'    y = y + frm("cmd3").height + 200
'    x = frm("imgLogo").Left
'
'    controlArr.Arr = "cmd6,cmd7,cmd8"
'
'    For i = 0 To proportionArr.Count - 1
'
'        proportion = CDbl(proportionArr.Arr(i)) / proportionTotal
'        frm(controlArr.Arr(i)).Left = x
'        frm(controlArr.Arr(i)).Top = y
'        frm(controlArr.Arr(i)).Width = (totalWidth - ((proportionArr.Count - 1) * colSpaceWidth * 2)) * proportion
'        frm(controlArr.Arr(i)).HorizontalAnchor = acHorizontalAnchorLeft
'
'        x = x + (colSpaceWidth * 2) + frm(controlArr.Arr(i)).Width
'
'    Next i
'
'    y = y + frm("cmd6").height + 200
'    x = frm("imgLogo").Left
'
'    controlArr.Arr = "cmd9,cmd10,cmd11"
'
'    For i = 0 To proportionArr.Count - 1
'
'        proportion = CDbl(proportionArr.Arr(i)) / proportionTotal
'        frm(controlArr.Arr(i)).Left = x
'        frm(controlArr.Arr(i)).Top = y
'        frm(controlArr.Arr(i)).Width = (totalWidth - ((proportionArr.Count - 1) * colSpaceWidth * 2)) * proportion
'        frm(controlArr.Arr(i)).HorizontalAnchor = acHorizontalAnchorLeft
'
'        x = x + (colSpaceWidth * 2) + frm(controlArr.Arr(i)).Width
'
'    Next i
'
'
'End Sub








    
