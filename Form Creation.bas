Attribute VB_Name = "Form Creation"
Option Compare Database
Option Explicit

Public Function RunFunctionFromSubform(frm As Form, subformName, FunctionName)
    
    Run FunctionName, frm(subformName).Form

End Function

'Public Function CreateSetFromMain(frm As Form)
'
'    Dim TableName As String, ViewName As String, DataEntryCaption As String, RsPrefix As String, VerbosePlural As String
'    Dim MainFormCaption As String, SetFocus As String, RequeryOnClose As String
'
'    TableName = "tbl" & frm("TableRecord") & "s"
'    If Not IsNull(frm("PluralForm")) Then
'        TableName = "tbl" & frm("PluralForm")
'    End If
'
'    ''Creation of View Name
'    If frm("isQuery") Then
'        RsPrefix = "qry"
'    Else
'        RsPrefix = "tbl"
'    End If
'
'    ViewName = RsPrefix & frm("TableRecord") & "s"
'
'    If Not IsNull(frm("PluralForm")) Then
'        ViewName = RsPrefix & frm("PluralForm")
'    End If
'
'    ''DataEntryCaption
'    If Not IsNull(frm("ReadableCaption")) Then
'        DataEntryCaption = frm("ReadableCaption") & " Form"
'        MainFormCaption = frm("ReadableCaption") & " List"
'    Else
'        DataEntryCaption = frm("TableRecord") & " Form"
'        MainFormCaption = frm("TableRecord") & " List"
'    End If
'
'    ''Set Focus
'    If IsNull(frm("SetFocus")) Then
'        SetFocus = frm("TableRecord")
'    Else
'        SetFocus = frm("SetFocus")
'    End If
'
'    RequeryOnClose = "main" & frm("TableRecord") & "s.subform"
'
'    If Not IsNull(frm("PluralForm")) Then
'        RequeryOnClose = "main" & frm("PluralForm") & ".subform"
'    End If
'
'    ''Primary Key
'    Dim PrimaryKey As String
'    PrimaryKey = frm("TableRecord") & "ID"
'
'    ''FormName
'    Dim FormName As String
'    FormName = frm("TableRecord") & "s"
'
'    If Not IsNull(frm("PluralForm")) Then
'        FormName = frm("PluralForm")
'    End If
'
'    Dim Fields(8) As String, fieldValues(8) As String
'
'    If IsNull(frm("VerbosePlural")) Then
'        If Not IsNull(frm("ReadableCaption")) Then
'            VerbosePlural = frm("ReadableCaption") & "s"
'        Else
'            VerbosePlural = frm("TableRecord") & "s"
'        End If
'    Else
'        VerbosePlural = frm("VerbosePlural")
'    End If
'
'    Fields(0) = "TableName"
'    Fields(1) = "ViewName"
'    Fields(2) = "DataEntryCaption"
'    Fields(3) = "MainFormCaption"
'    Fields(4) = "SetFocus"
'    Fields(5) = "RequeryOnClose"
'    Fields(6) = "PrimaryKey"
'    Fields(7) = "FormName"
'    Fields(8) = "VerbosePlural"
'
'    fieldValues(0) = "'" & TableName & "'"
'    fieldValues(1) = "'" & ViewName & "'"
'    fieldValues(2) = "'" & DataEntryCaption & "'"
'    fieldValues(3) = "'" & MainFormCaption & "'"
'    fieldValues(4) = "'" & SetFocus & "'"
'    fieldValues(5) = "'" & RequeryOnClose & "'"
'    fieldValues(6) = "'" & PrimaryKey & "'"
'    fieldValues(7) = "'" & FormName & "'"
'    fieldValues(8) = EscapeString(VerbosePlural)
'
'    InsertAndLog "tblTables", Fields, fieldValues
'
'    CreateSet TableName
'
'End Function

'Public Function CreateSet(RecordSource As String)
'    CreateDataEntryForm RecordSource
'    CreateDataSheetForm RecordSource
'    CreateMainForm RecordSource
'End Function

'Public Function CreateNewForm(RecordSource As String, Optional FormType As Integer)
'
'    ''RecordSource => Must be the recordsource in which the form will be based
'    ''Form Type => Can either be 0 = "DataEntry", 1 = "DataSheet", 2 = "MainForm"
'
'    Select Case FormType
'        Case 0:
'            CreateDataEntryForm RecordSource
'        Case 1:
'            CreateDataSheetForm RecordSource
'        Case 2:
'            CreateMainForm RecordSource
'    End Select
'
'End Function

Private Function SetFormProperties(frmType As String, frm As Form)
    ''Set the Form Properties
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblFormProps WHERE FrmType = '" & frmType & "'")
    Do Until rs.EOF
        frm.Properties(rs.fields("FrmPropName")) = rs.fields("FrmPropVal")
        rs.MoveNext
    Loop
End Function

Private Function SetCaption(fld As Field)

    Dim ctl As Control
    ''Add a filed caption when it does not exist
    Dim fldCaption
    If DoesPropertyExists(fld.Properties, "Caption") Then
        fldCaption = fld.Properties("Caption")
    Else
        fldCaption = AddSpaces(fld.Name)
        Dim prop As property
        Set prop = fld.CreateProperty("Caption", dbText, fldCaption)
        fld.Properties.Append prop
    End If
    
    SetCaption = fldCaption
    
End Function

Private Function CreateDataSheetForm(RecordSource As String)

    ''Get the tblTables properties of the Matched TableName RecordSource
    Dim props As Recordset
    Set props = CurrentDb.OpenRecordset("SELECT * FROM tblTables WHERE TableName = '" & RecordSource & "'")
    Dim ViewName As String, PrimaryKey As String
    ViewName = props.fields("viewName")
    PrimaryKey = props.fields("PrimaryKey")
    
    Dim frm As Form
    Set frm = CreateForm
    frm.RecordSource = ViewName
    
    frm.Caption = props.fields("MainFormCaption")
    frm.BeforeUpdate = "=SaveFormData([Form],""" & RecordSource & """,""" & PrimaryKey & """)"
    frm.OnLoad = "=SetDefaultUserID([Form])"
    
    ''Set the Form Properties
    SetFormProperties "Datasheet", frm
    
    ''Establish the Recordset
    Dim rsDef As Object
    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    If DoesPropertyExists(CurrentDb.TableDefs, ViewName) Then
        Set rsDef = dbs.TableDefs(ViewName)
    Else
        Set rsDef = dbs.QueryDefs(ViewName)
    End If
    
    Dim x, y
    'x is the starting left, y is the starting top
    x = 800: y = 600
    
    Dim fld As Field
    
    For Each fld In rsDef.fields
    
        If fld.Name = props.fields("PrimaryKey") Then
            GoTo NextField
        End If
        
        ''Create the label first
        Dim ctl As Control
        ''Add a filed caption when it does not exist
        Dim fldCaption
        fldCaption = SetCaption(fld)
        
        Dim ControlTypeID
        If Not DoesPropertyExists(fld.Properties, "DisplayControl") Then
            ControlTypeID = acTextBox
        Else
            ControlTypeID = fld.Properties("DisplayControl")
        End If
        
        Set ctl = CreateControl(frm.Name, ControlTypeID, , "", fld.Name, x, y, 3000)
        ctl.Name = fld.Name: ctl.Properties("DatasheetCaption") = fldCaption
        
        ''Set Control Caption
        Dim rs As Recordset
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblCtlProps WHERE isNull(CtlPropType)")
        rs.MoveFirst
        
        Do Until rs.EOF
            If DoesPropertyExists(ctl.Properties, rs.fields("CtlPropName")) Then
                ctl.Properties(rs.fields("CtlPropName")) = rs.fields("CtlPropVal")
            End If
            rs.MoveNext
        Loop
        
        y = y + 400
        
        If y > 15000 Then
            y = 600
            x = 2500
        End If
        
        InsertToFields fld, RecordSource, fldCaption
        
NextField:
        
    Next fld
    
    frm("Timestamp").ColumnHidden = True
    frm("CreatedBy").ColumnHidden = True
    
    Dim frmName As String, customFrmName As String, i As Integer
    frmName = frm.Name: customFrmName = "dsht" & props.fields("FormName")
    DoCmd.Close acForm, frm.Name, acSaveYes
    
    Do Until Not FrmExist(customFrmName)
        customFrmName = customFrmName & "_1"
    Loop
    
    DoCmd.Rename customFrmName, acForm, frmName
    
End Function

Public Function RenderButton(x As Long, y As Long, Caption As String, QuickStyle As Integer, frm As Form, cmdName As String, Optional parentName = "")
    
    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acCommandButton, , parentName, , x, y, 1250)
    With ctl
        .Name = "cmd" & cmdName
        .Properties("Caption") = Caption
        
        Dim rs As Recordset
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblCtlProps WHERE CtlPropType = '104'")
        rs.MoveFirst
        
        Do Until rs.EOF
            If DoesPropertyExists(.Properties, rs.fields("CtlPropName")) Then
                .Properties(rs.fields("CtlPropName")) = rs.fields("CtlPropVal")
            End If
            rs.MoveNext
        Loop
        
        .Properties("QuickStyle") = QuickStyle
        .Properties("UseTheme") = False
        .Properties("CursorOnHover") = 1
    End With
    
End Function


Private Function CreateDataEntryForm(RecordSource As String)

    ''Get the tblTables properties of the Matched TableName RecordSource
    Dim props As Recordset
    Set props = CurrentDb.OpenRecordset("SELECT * FROM tblTables WHERE TableName = '" & RecordSource & "'")
    Dim ViewName As String, PrimaryKey As String
    ViewName = props.fields("ViewName")
    PrimaryKey = props.fields("PrimaryKey")
    
    Dim frm As Form
    Set frm = CreateForm
    frm.RecordSource = ViewName
    frm.Caption = props.fields("DataEntryCaption")
    
    frm.OnCurrent = "=SetFocusOnForm([Form],""" & props.fields("SetFocus") & """)"
    frm.BeforeUpdate = "=SaveFormData([Form],""" & RecordSource & """,""" & PrimaryKey & """)"
    frm.OnLoad = "=SetDefaultUserID([Form])"
    
    ''Set the Form Properties
    SetFormProperties "DataEntry", frm
    
    ''Establish the Recordset
    Dim rsDef As Object
    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    If DoesPropertyExists(CurrentDb.TableDefs, ViewName) Then
        Set rsDef = dbs.TableDefs(ViewName)
    Else
        Set rsDef = dbs.QueryDefs(ViewName)
    End If
    
    Dim x As Long, y As Long
    'x is the starting left, y is the starting top
    x = 800: y = 600
    
    Dim fld As Field
    
    For Each fld In rsDef.fields
        
        
        Select Case fld.Name
            Case props.fields("PrimaryKey"), "Timestamp", "CreatedBy":
                GoTo NextField
        End Select
        
'        If fld.Name = props.Fields("PrimaryKey") Or fld.Name = "Timestamp" Or fld.Name Then
'            GoTo NextField
'        End If
        
        ''Create the label first
        Dim ctl As Control
        ''Add a filed caption when it does not exist
        Dim fldCaption
        fldCaption = SetCaption(fld)
        
        Set ctl = CreateControl(frm.Name, acLabel, , fld.Name, fldCaption, x, y)
        
        Select Case fld.Name
            Case "RecordTimestamp", "UserID":
                ctl.Visible = False
        End Select
        
        ''Label Properties
        Dim rs As Recordset
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblCtlProps WHERE CtlPropType = ""100""")
        
        Do Until rs.EOF
            ctl.Properties(rs.fields("CtlPropName")) = rs.fields("CtlPropVal")
            rs.MoveNext
        Loop
        
        ''Generate the control field
        y = y + 380
        
        Dim ControlTypeID
        If Not DoesPropertyExists(fld.Properties, "DisplayControl") Then
            ControlTypeID = acTextBox
        Else
            ControlTypeID = fld.Properties("DisplayControl")
        End If
        
'        If ControlTypeID = 106 Then
'            Set ctl = CreateControl(frm.Name, ControlTypeID, , "", fld.Name, x + 1750, y - 370, 3000)
'            ctl.Name = fld.Name
'        Else
'            Set ctl = CreateControl(frm.Name, ControlTypeID, , "", fld.Name, x, y, 3000)
'            ctl.Name = fld.Name
'        End If

        Set ctl = CreateControl(frm.Name, ControlTypeID, , "", fld.Name, x, y, 3000)
        ctl.Name = fld.Name
        
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblCtlProps WHERE isNull(CtlPropType)")
        rs.MoveFirst
        
        
        Do Until rs.EOF
            If DoesPropertyExists(ctl.Properties, rs.fields("CtlPropName")) Then
                ctl.Properties(rs.fields("CtlPropName")) = rs.fields("CtlPropVal")
            End If
            rs.MoveNext
        Loop
        
        y = y + 400
        
        If y > 15000 Then
            y = 600
            x = 2500
        End If
        
        
        InsertToFields fld, RecordSource, fldCaption

NextField:
        
        
    Next fld
    
    ''Create the Timestamp and CreatedBy field (Hidden Fields)
    Set ctl = CreateControl(frm.Name, acTextBox, , "", "Timestamp", 0, 0, 0)
    ctl.Name = "Timestamp"
    
    Set ctl = CreateControl(frm.Name, acComboBox, , "", "CreatedBy", 0, 0, 0)
    ctl.Name = "CreatedBy"
    
    y = y + 400
    ''New Record
    RenderButton x, y, "Cancel", 23, frm, "Cancel"
    x = x + 1300
    RenderButton x, y, "New", 23, frm, "New"
    x = x + 1300
    ''Save Record
    RenderButton x, y, "Save", 23, frm, "SaveClose"
    x = x + 1300
    ''Delete Record
    RenderButton x, y, "Delete", 24, frm, "Delete"
    
    frm.cmdCancel.OnClick = "=CancelEdit([Form])"
    frm.cmdNew.OnClick = "=Save([Form],'" & RecordSource & "',0)"
    frm.cmdSaveClose.OnClick = "=Save([Form],'" & RecordSource & "',1)"
    frm.cmdDelete.OnClick = "=DeleteRecord([Form], '" & props.fields("PrimaryKey") & "', '" & RecordSource & "')"
    
    ''Set background color
    frm.Detail.BackColor = RGB(81, 163, 36)
    
    frm("Timestamp").Visible = False
    frm("CreatedBy").Visible = False
    
    Dim frmName As String, customFrmName As String, i As Integer
    frmName = frm.Name: customFrmName = "frm" & props.fields("FormName")
    DoCmd.Close acForm, frm.Name, acSaveYes
    
    Do Until Not FrmExist(customFrmName)
        customFrmName = customFrmName & "_1"
    Loop
    
    DoCmd.Rename customFrmName, acForm, frmName
    
End Function

Public Function FrmExist(sFrmName As String) As Boolean
    On Error GoTo Error_Handler
    Dim frm                   As Access.AccessObject
 
    For Each frm In Application.CurrentProject.AllForms
        If sFrmName = frm.Name Then
            FrmExist = True
            Exit For    'We know it exist so let leave, no point continuing
        End If
    Next frm
 
Error_Handler_Exit:
    On Error Resume Next
    Set frm = Nothing
    Exit Function
 
Error_Handler:
    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.number & vbCrLf & _
           "Error Source: FrmExist" & vbCrLf & _
           "Error Description: " & Err.Description, _
           vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

Public Function RptExist(sFrmName As String) As Boolean
    On Error GoTo Error_Handler
    Dim frm                   As Access.AccessObject
 
    For Each frm In Application.CurrentProject.AllReports
        If sFrmName = frm.Name Then
            RptExist = True
            Exit For    'We know it exist so let leave, no point continuing
        End If
    Next frm
 
Error_Handler_Exit:
    On Error Resume Next
    Set frm = Nothing
    Exit Function
 
Error_Handler:
    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.number & vbCrLf & _
           "Error Source: FrmExist" & vbCrLf & _
           "Error Description: " & Err.Description, _
           vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

Public Function ImportFieldsToTable(RecordSource As String)
    
    ''Establish the Recordset
    Dim rsDef As Object
    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    If DoesPropertyExists(CurrentDb.TableDefs, RecordSource) Then
        Set rsDef = dbs.TableDefs(RecordSource)
    Else
        Set rsDef = dbs.QueryDefs(RecordSource)
    End If
    
    Dim fld As Field
    
    For Each fld In rsDef.fields
    
        Dim fldCaption
        fldCaption = SetCaption(fld)
    
        InsertToFields fld, RecordSource, fldCaption
        
    Next fld
    
End Function

Public Function InsertToFields(fld As Field, RecordSource As String, fldCaption As Variant)
    If Not isPresent("tblFormFields", "TableName = '" & RecordSource & "' And FieldName = '" & fld.Name & "'") Then
        Dim fields(4) As String: Dim fieldValues(4) As String
        fields(0) = "FieldName"
        fields(1) = "FieldCaption"
        fields(2) = "FieldTypeID"
        fields(3) = "ValidationString"
        fields(4) = "TableName"
        
        fieldValues(0) = """" & fld.Name & """"
        fieldValues(1) = """" & fldCaption & """"
        fieldValues(2) = fld.Type
        
        If fld.Required = True Then
            fieldValues(3) = """required"""
        Else
            fieldValues(3) = "Null"
        End If
        
        fieldValues(4) = """" & RecordSource & """"
        
        InsertData "tblFormFields", fields, fieldValues
    End If
End Function
