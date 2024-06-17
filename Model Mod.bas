Attribute VB_Name = "Model Mod"
Option Compare Database
Option Explicit

Public Function tblModelFields_ModelField_AfterUpdate(frm As Form)

    Dim ModelID: ModelID = frm("ModelID")
    Dim ModelFieldID: ModelFieldID = frm("ModelFieldID")
    Dim FieldOrder: FieldOrder = frm("FieldOrder")
    
    If FieldOrder <> 0 Then
        FieldOrder = ELookup("tblModelFields", "ModelID = " & ModelID & " ANd ModelFieldID < " & ModelFieldID, "FieldOrder", "FieldOrder DESC")
        frm("FieldOrder") = FieldOrder
    End If
    
End Function

Public Function CreateFormSet(frm2 As Form)

    CreateMainForm frm2
    CreateDEForm frm2
    CreateSimpleDEForm frm2
    
    
End Function

Public Function CreateCustomModule(frm2 As Form)
    
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim LineNum As Long: LineNum = 4
    
    Set VBProj = Application.VBE.VBProjects(Application.GetOption("Project Name"))
    
    Dim ModelID, Model, VerboseName, VerbosePlural, MainField, TableWideValidation, FormColumns, SetFocus, IsKeyVisible, QueryName, OnFormCreate, subformName, Timestamp, CreatedBy, PrimaryKey
    
    ModelID = frm2("ModelID")
    Model = frm2("Model")
    VerboseName = frm2("VerboseName")
    VerbosePlural = frm2("VerbosePlural")
    MainField = frm2("MainField")
    TableWideValidation = frm2("TableWideValidation")
    FormColumns = frm2("FormColumns")
    SetFocus = frm2("SetFocus")
    IsKeyVisible = frm2("IsKeyVisible")
    QueryName = frm2("QueryName")
    OnFormCreate = frm2("OnFormCreate")
    subformName = frm2("SubformName")
    Timestamp = frm2("Timestamp")
    CreatedBy = frm2("CreatedBy")
    PrimaryKey = frm2("PrimaryKey")
    
    Dim ModuleName: ModuleName = concat(Model, " Mod")
    
    
    InsertToModelRelatedObjects ModelID, acModule, ModuleName
    
    If DoesPropertyExists(VBProj.VBComponents, ModuleName) Then
        Set VBComp = VBProj.VBComponents(ModuleName)
    Else
    
        Set VBComp = VBProj.VBComponents.Add(vbext_ct_StdModule)
        VBComp.Name = ModuleName
        
        Set CodeMod = VBComp.CodeModule
        
        InsertLines CodeMod, LineNum, concat("Public Function ", Model, "Create(frm AS Form, FormTypeID)")
        InsertLines CodeMod, LineNum, ""
        
        InsertLines CodeMod, LineNum, vbTab & "Select Case FormTypeID"
        InsertLines CodeMod, LineNum, vbTab & vbTab & "Case 4: ''Data Entry Form"
        InsertLines CodeMod, LineNum, vbTab & vbTab & "Case 5: ''Datasheet Form"
        InsertLines CodeMod, LineNum, vbTab & vbTab & "Case 6: ''Main Form"
        InsertLines CodeMod, LineNum, vbTab & vbTab & "Case 7: ''Tabular Report"
        InsertLines CodeMod, LineNum, vbTab & "End Select"
        
        InsertLines CodeMod, LineNum, ""
        
        InsertLines CodeMod, LineNum, "End Function"
        
        
    End If
    
    frm2("OnFormCreate") = concat(Model, "Create")
    
    DoCmd.Save acModule, ModuleName
    
    
    
End Function

Private Sub InsertLines(CodeMod As Object, LineNum As Long, CodeStr)
    
    CodeMod.InsertLines LineNum, CodeStr
    LineNum = LineNum + 1
    
End Sub

Public Function GenerateDatasheetControls(frm As Form)
    
    Dim ModelID, Model, VerboseName, VerbosePlural, MainField, TableWideValidation, FormColumns, SetFocus, IsKeyVisible, QueryName, OnFormCreate, subformName, UserQueryFields, IsSystemTable
    ModelID = frm("ModelID")
    
    If ExitIfTrue(IsNull(ModelID), "Please select a record.") Then Exit Function
    
    Model = frm("Model")
    VerboseName = frm("VerboseName")
    VerbosePlural = frm("VerbosePlural")
    MainField = frm("MainField")
    TableWideValidation = frm("TableWideValidation")
    FormColumns = frm("FormColumns")
    SetFocus = frm("SetFocus")
    IsKeyVisible = frm("IsKeyVisible")
    QueryName = frm("QueryName")
    OnFormCreate = frm("OnFormCreate")
    subformName = frm("SubformName")
    UserQueryFields = frm("UserQueryFields")
    IsSystemTable = frm("IsSystemTable")
    
    Dim dshtName
    If Not IsNull(subformName) Then
        dshtName = concat("dsht", subformName)
    Else
        If Not IsNull(VerbosePlural) Then
            dshtName = concat("dsht", VerbosePlural)
        Else
            dshtName = concat("dsht", Model, "s")
        End If
    End If
    
    DoCmd.OpenForm dshtName, acDesign
    
    Dim dshtFrm As Form, ctl As Control, ControlName, ControlCaption, ControlOrder As Integer
    Set dshtFrm = Forms(dshtName)
    
    ControlOrder = 1
    For Each ctl In dshtFrm.Section(acFooter).Controls
        ControlName = ctl.Name
        ControlCaption = AddSpaces(replace(ControlName, "Sum", ""))
        If Not isPresent("tblDatasheetTotals", "ControlName = " & EscapeString(ControlName) & " AND ModelID = " & ModelID) Then
            RunSQL "INSERT INTO tblDatasheetTotals (ControlName,ControlCaption,ControlOrder,ModelID) VALUES (" & _
                    EscapeString(ControlName) & "," & _
                    EscapeString(ControlCaption) & "," & _
                    ControlOrder & "," & _
                    ModelID & ")"
        End If
        ControlOrder = ControlOrder + 1
    Next ctl
    
    DoCmd.Close acForm, dshtFrm.Name, acSaveNo
    
    If DoesPropertyExists(frm, "subDatasheetTotals") Then
        frm("subDatasheetTotals").Form.Requery
    End If
    
    MsgBox "Datasheet Control Successfully Imported.."
    
End Function

Private Function OverrideProperties(ModelID, FormTypeID, frm As Object)
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblPropertyOverrides WHERE ModelID = " & ModelID & " And FormTypeID = " & FormTypeID)
    
    Dim PropertyOverrideID, ControlName, PropertyName, PropertyValue
    
    Do Until rs.EOF
        PropertyOverrideID = rs.fields("PropertyOverrideID")
        ControlName = rs.fields("ControlName")
        PropertyName = rs.fields("PropertyName")
        PropertyValue = rs.fields("PropertyValue")
        If Not IsNull(ControlName) Then
            If DoesPropertyExists(frm, ControlName) Then
                If DoesPropertyExists(frm(ControlName).Properties, PropertyName) Then
                    frm(ControlName).Properties(PropertyName) = PropertyValue
                End If
            End If
        Else
            If DoesPropertyExists(frm.Properties, PropertyName) Then
                frm.Properties(PropertyName) = PropertyValue
            End If
        End If
        rs.MoveNext
    Loop
    
End Function

Public Function CreateTabularReport(frm2 As Form)
        
    Dim rpt As Report, rs As Recordset, rsName, frmCaption, PrimaryKey, fldName, fldWidth
    Dim CurrentCol, x, y, isMemo As Boolean, maxWidth
    Dim ctl As Control
    
    Dim ModelID, Model, VerboseName, VerbosePlural, MainField, TableWideValidation, FormColumns, SetFocus, IsKeyVisible
    Dim QueryName, OnFormCreate, subformName, UserQueryFields, Timestamp, CreatedBy
    
    ModelID = frm2("ModelID")
    Model = frm2("Model")
    VerboseName = frm2("VerboseName")
    VerbosePlural = frm2("VerbosePlural")
    MainField = frm2("MainField")
    TableWideValidation = frm2("TableWideValidation")
    FormColumns = frm2("FormColumns")
    SetFocus = frm2("SetFocus")
    IsKeyVisible = frm2("IsKeyVisible")
    QueryName = frm2("QueryName")
    OnFormCreate = frm2("OnFormCreate")
    subformName = frm2("SubformName")
    UserQueryFields = frm2("UserQueryFields")
    Timestamp = frm2("Timestamp")
    CreatedBy = frm2("CreatedBy")
    
    If ExitIfTrue(IsNull(ModelID), "Selection is empty..") Then Exit Function
    
    '''Create the controls
    '''Fields, default buttons, additional buttons and controls
    
    ''Create the form
    Set rpt = CreateReport
    rsName = GetTableName(Model, VerbosePlural)
    
    Dim tblName
    tblName = rsName
    
    If Not IsNull(QueryName) Then
        rsName = QueryName
    End If
    
    rpt.RecordSource = rsName
    frmCaption = GetFieldCaption(VerboseName, Model)
    rpt.Caption = concat(frmCaption, " List")
    rpt.PopUp = True
    rpt.AutoCenter = True


    DoCmd.RunCommand acCmdReportHdrFtr
    
    Dim rsObj As Object, db As DAO.Database, fld As DAO.Field
    Set db = CurrentDb
    If DoesPropertyExists(db.TableDefs, rsName) Then
        Set rsObj = db.TableDefs(rsName)
    Else
        Set rsObj = db.QueryDefs(rsName)
    End If
    
    x = 0
    y = 0
    
    CurrentCol = 0: isMemo = False
    
    Dim sqlStr As String
    sqlStr = "SELECT * FROM tblModelFields WHERE ReportFieldOrder IS NOT NULL AND ModelID = " & ModelID
    
    If Not UserQueryFields Then
        sqlStr = sqlStr & " AND FieldSource = " & EscapeString(rsName)
    End If
    
    Set rs = ReturnRecordset(sqlStr & " ORDER BY ReportFieldOrder ASC")
    
    Do Until rs.EOF
    
        ''Look First if the ControlSource is not Null
        If Not IsNull(rs.fields("ControlSource")) Then
           If isPresent("qryModelFieldProperties", "Property = ""onDatasheetOnly"" And ModelFieldID = " & rs.fields("ModelFieldID")) Then
               
                GoTo NextField
            Else
                ''Create a control at the footer of the form
                fldWidth = 3000 * (rs.fields("Columns")) + (200 * (rs.fields("Columns") - 1))
                Set ctl = CreateReportControl(rpt.Name, acTextBox, , "", "", x, y, fldWidth)
                SetControlProperties ctl
                ctl.Name = rs.fields("ModelField")
                ctl.ControlSource = rs.fields("ControlSource")
                
                Select Case rs.fields("FieldTypeID")
                    Case dbMemo:
                        ctl.height = 900
                        isMemo = True
                    Case dbDouble
                        ctl.Format = "Standard"
                End Select
            
                ''Generate the label just above the control
                Dim CustomVerboseName
                If IsNull(rs.fields("VerboseName")) Then
                    CustomVerboseName = AddSpaces(rs.fields("ModelField"))
                Else
                    CustomVerboseName = rs.fields("VerboseName")
                End If
                Set ctl = CreateReportControl(rpt.Name, acLabel, , rs.fields("ModelField"), CustomVerboseName, x, y - 300)
                SetControlProperties ctl
                ctl.width = fldWidth
                
                GoTo SetVariables
            End If
            
        End If
        
        fldName = GetFieldName(rs.fields("ForeignKey"), rs.fields("ModelField"))
        
        If Not DoesPropertyExists(rsObj.fields, fldName) Then
            GoTo NextField
        End If
        
        Set fld = rsObj.fields(fldName)
        
        If Not IsKeyVisible And fld.Name = PrimaryKey Then
            GoTo NextField
        End If
        
        Select Case fld.Name
            Case "Timestamp", "CreatedBy":
                GoTo NextField
        End Select
        
        Dim ControlTypeValue
        If Not DoesPropertyExists(fld.Properties, "DisplayControl") Then
            ControlTypeValue = acTextBox
        Else
            ControlTypeValue = fld.Properties("DisplayControl")
        End If
        
        ''Generate the control first before the label
        ''Get the width depending on the number of columns but make sure that the CurrentCol + the columns will not exceed
        ''the FormColumn
        If fld.Type = dbMemo Then
            fldWidth = 2500
        Else
            fldWidth = 1200
        End If
        
        ''Setting the FieldOrder to 0 here will make the fld hidden..
        If rs.fields("FieldOrder") = 0 Then
'            Set ctl = CreateReportControl(rpt.Name, ControlTypeValue, , "", fld.Name, 0, 0, 0)
'            ctl.Name = fld.Name
'            SetControlProperties ctl
            GoTo NextField
        End If
        
        If fld.Type = dbBoolean Then
            Set ctl = CreateReportControl(rpt.Name, acTextBox, , "", fld.Name, x, y, fldWidth)
            SetControlProperties ctl
            ctl.FontName = "Wingdings"
            ctl.Format = "ü;\û"
            ctl.fontSize = 12
            ctl.TextAlign = 2
        Else
            Set ctl = CreateReportControl(rpt.Name, ControlTypeValue, , "", fld.Name, x, y, fldWidth)
            SetControlProperties ctl
        End If
        
        ctl.BottomMargin = 50
        ctl.LeftMargin = 50
        ctl.RightMargin = 50
        ctl.TopMargin = 50
        ctl.Name = fld.Name
        ctl.BackStyle = 0
        
        ''Set control property based on ControlTypeValue
        
        'ctl.BorderStyle = 0
    
        Select Case fld.Type
             Case dbMemo:
                 'ctl.height = 900
                 isMemo = True
             Case dbDouble:
                 ctl.Format = "Standard"
         End Select
         
         ctl.CanGrow = True
         ctl.InSelection = True
        
        ''Generate the label just above the control
        Dim ControlCaption
        If Not DoesPropertyExists(fld.Properties, "Caption") Then
            If IsNull(rs.fields("VerboseName")) Then
                ControlCaption = AddSpaces(rs.fields("ModelField"))
            Else
                ControlCaption = rs.fields("VerboseName")
            End If
        Else
            ControlCaption = fld.Properties("Caption")
        End If
        
        Set ctl = CreateReportControl(rpt.Name, acLabel, acPageHeader, , ControlCaption, x, y)

        SetControlProperties ctl
        ctl.width = fldWidth
        ctl.TextAlign = 2
        'ctl.Height = 200
        ctl.InSelection = True
        ctl.BackStyle = 1
        ctl.BackColor = 49407
        ctl.ForeColor = RGB(81, 163, 36)
        ctl.top = 2000
        ctl.BorderStyle = 1

SetVariables:
    
        CurrentCol = CurrentCol + rs.fields("Columns")
        
        x = x + fldWidth
                 
NextField:
        
        rs.MoveNext
        
    Loop
    
    DoCmd.RunCommand acCmdTabularLayout
    DoCmd.RunCommand acCmdControlPaddingNone
    
    For Each ctl In rpt.Controls
        If ctl.InSelection Then
            ctl.left = 0
            Exit For
        End If
    Next ctl
    
    Dim rptWidth
    rptWidth = (8.5 - (0.25 * 2)) * 1440
    
    ''Write the pageheader (The Report Title & Current Date Time)
    Set ctl = CreateReportControl(rpt.Name, acLabel, acPageHeader, , , 0, 0, 4800, 400)
    SetControlProperties ctl
    ctl.fontSize = 12
    ctl.height = 340
    ctl.Caption = AddSpaces(GetModelPlural(Model, VerbosePlural, ""))
    ctl.ForeColor = RGB(81, 163, 36)
    
    ''Current Date Time
    fldWidth = 3000
    Set ctl = CreateReportControl(rpt.Name, acTextBox, acPageHeader, , , rptWidth - fldWidth, 0, fldWidth, 400)
    SetControlProperties ctl
    ctl.ControlSource = "=Now()"
    ctl.BorderStyle = 0
    
    Set ctl = CreateReportControl(rpt.Name, acTextBox, acPageHeader, , , 0, ctl.top + ctl.height + 200, rptWidth, 570)
    SetControlProperties ctl
    ctl.fontSize = 12
    ctl.height = 570
    ctl.ForeColor = RGB(254, 254, 254)
    ctl.BackColor = RGB(81, 163, 36)
    ctl.ControlSource = "=" & EscapeString("SOME CAPTION")
    ctl.TextAlign = 2
    ctl.FontBold = True
    ctl.TopMargin = 100
    
    ''At pagefooter (The Page x of y)
    Set ctl = CreateReportControl(rpt.Name, acTextBox, acPageFooter, , , 0, 0, fldWidth, 400)
    SetControlProperties ctl
    ctl.ControlSource = "=""Page "" & [Page] & "" of "" & [Pages]"
    ctl.BorderStyle = 0
    
    rpt.Section(acHeader).height = 0
    rpt.Section(acHeader).BackColor = RGB(254, 254, 254)
    rpt.Section(acFooter).height = 0
    rpt.Section(acPageFooter).height = 0
    rpt.Section(acPageHeader).height = 0
    rpt.Section(acDetail).height = 0
    rpt.Section(acDetail).AlternateBackColor = RGB(250, 243, 232)
    rpt.Section(acDetail).BackColor = RGB(254, 254, 254)
    
    Dim frmName As String, customFrmName As String, baseFormName As String, i As Integer
    frmName = rpt.Name
    If Not IsNull(VerbosePlural) Then
        baseFormName = concat("rptTab", replace(VerbosePlural, " ", ""))
    Else
        baseFormName = concat("rptTab", Model, "s")
    End If
    
    If Not IsNull(subformName) Then
        baseFormName = concat("rptTab", subformName)
    End If
    
    ''Special Function to run on form creation
    If Not IsNull(OnFormCreate) Then
        Run OnFormCreate, rpt, 7
    End If
    
    ''Override
    OverrideProperties ModelID, 7, rpt
    
    DoCmd.Close acReport, rpt.Name, acSaveYes
    
    customFrmName = baseFormName
    
'    Do Until Not RptExist(customFrmName)
'
'        If MsgBox(customFrmName & " already exists. Would you like to replace it?", vbYesNo) = vbYes Then
'            Exit Do
'        End If
'        i = i + 1
'        customFrmName = baseFormName & "_" & i
'
'    Loop
    
    
    ''Insert the newly created form to the InsertToModelRelatedObjects
    InsertToModelRelatedObjects ModelID, acReport, customFrmName
    
    DoCmd.Rename customFrmName, acReport, frmName
    DoCmd.OpenReport customFrmName, acViewPreview
    
End Function

Public Function CreateColumnReport(frm2 As Form)
    
    Dim rpt As Report, rs As Recordset, rsName, frmCaption, PrimaryKey, fldName, fldWidth
    Dim CurrentCol, x, y, isMemo As Boolean, maxWidth
    Dim ctl As Control
    
    Dim ModelID, Model, VerboseName, VerbosePlural, MainField, TableWideValidation, FormColumns, SetFocus, IsKeyVisible
    Dim QueryName, OnFormCreate, subformName, UserQueryFields, Timestamp, CreatedBy
    
    ModelID = frm2("ModelID")
    Model = frm2("Model")
    VerboseName = frm2("VerboseName")
    VerbosePlural = frm2("VerbosePlural")
    MainField = frm2("MainField")
    TableWideValidation = frm2("TableWideValidation")
    FormColumns = frm2("FormColumns")
    SetFocus = frm2("SetFocus")
    IsKeyVisible = frm2("IsKeyVisible")
    QueryName = frm2("QueryName")
    OnFormCreate = frm2("OnFormCreate")
    subformName = frm2("SubformName")
    UserQueryFields = frm2("UserQueryFields")
    Timestamp = frm2("Timestamp")
    CreatedBy = frm2("CreatedBy")
    
    If ExitIfTrue(IsNull(ModelID), "Selection is empty..") Then Exit Function
    
    '''Create the controls
    '''Fields, default buttons, additional buttons and controls
    
    ''Create the form
    Set rpt = CreateReport
    rsName = GetTableName(Model, VerbosePlural)
    
    Dim tblName
    tblName = rsName
    
    If Not IsNull(QueryName) Then
        rsName = QueryName
    End If
    
    rpt.RecordSource = rsName
    frmCaption = GetFieldCaption(VerboseName, Model)
    rpt.Caption = concat(frmCaption, " Report")
    rpt.PopUp = True

    DoCmd.RunCommand acCmdReportHdrFtr
    
    Dim rsObj As Object, db As DAO.Database, fld As DAO.Field
    Set db = CurrentDb
    If DoesPropertyExists(db.TableDefs, rsName) Then
        Set rsObj = db.TableDefs(rsName)
    Else
        Set rsObj = db.QueryDefs(rsName)
    End If
    
    x = 0
    y = 0
    
    CurrentCol = 0: isMemo = False
    
    Dim sqlStr As String
    sqlStr = "SELECT * FROM tblModelFields WHERE ReportFieldOrder IS NOT NULL AND ModelID = " & ModelID
    
    If Not UserQueryFields Then
        sqlStr = sqlStr & " AND FieldSource = " & EscapeString(rsName)
    End If
    
    Set rs = ReturnRecordset(sqlStr & " ORDER BY ReportFieldOrder ASC")
    
    Dim maxY
    
    Do Until rs.EOF
        
        maxY = GetMaxY(rpt) + 200
        ''Look First if the ControlSource is not Null
        If Not IsNull(rs.fields("ControlSource")) Then
           If isPresent("qryModelFieldProperties", "Property = ""onDatasheetOnly"" And ModelFieldID = " & rs.fields("ModelFieldID")) Then
               
                GoTo NextField
            Else
                ''Create a control at the footer of the form
                fldWidth = 3000 * (rs.fields("Columns")) + (200 * (rs.fields("Columns") - 1))
                Set ctl = CreateReportControl(rpt.Name, acTextBox, , "", "", x, y, fldWidth)
                SetControlProperties ctl
                ctl.Name = rs.fields("ModelField")
                ctl.ControlSource = rs.fields("ControlSource")
                
                Select Case rs.fields("FieldTypeID")
                    Case dbMemo:
                        ctl.height = 900
                        isMemo = True
                    Case dbDouble
                        ctl.Format = "Standard"
                End Select
            
                ''Generate the label just above the control
                Dim CustomVerboseName
                If IsNull(rs.fields("VerboseName")) Then
                    CustomVerboseName = AddSpaces(rs.fields("ModelField"))
                Else
                    CustomVerboseName = rs.fields("VerboseName")
                End If
                Set ctl = CreateReportControl(rpt.Name, acLabel, , rs.fields("ModelField"), CustomVerboseName, x, y - 300)
                SetControlProperties ctl
                ctl.width = fldWidth
                
                GoTo SetVariables
            End If
            
        End If
        
        fldName = GetFieldName(rs.fields("ForeignKey"), rs.fields("ModelField"))
        
        If Not DoesPropertyExists(rsObj.fields, fldName) Then
            GoTo NextField
        End If
        
        Set fld = rsObj.fields(fldName)
        
        If Not IsKeyVisible And fld.Name = PrimaryKey Then
            GoTo NextField
        End If
        
        Select Case fld.Name
            Case "Timestamp", "CreatedBy":
                GoTo NextField
        End Select
        
        Dim ControlTypeValue
        If Not DoesPropertyExists(fld.Properties, "DisplayControl") Then
            ControlTypeValue = acTextBox
        Else
            ControlTypeValue = fld.Properties("DisplayControl")
        End If
        
        ''Generate the control first before the label
        ''Get the width depending on the number of columns but make sure that the CurrentCol + the columns will not exceed
        ''the FormColumn
        If fld.Type = dbMemo Then
            fldWidth = 2500
        Else
            fldWidth = 1200
        End If
        
        ''Setting the FieldOrder to 0 here will make the fld hidden..
        If rs.fields("FieldOrder") = 0 Then
'            Set ctl = CreateReportControl(rpt.Name, ControlTypeValue, , "", fld.Name, 0, 0, 0)
'            ctl.Name = fld.Name
'            SetControlProperties ctl
            GoTo NextField
        End If
        
        If fld.Type = dbBoolean Then
            Set ctl = CreateReportControl(rpt.Name, ControlTypeValue, , "", fld.Name, x + 550, y, 200)
        Else
            Set ctl = CreateReportControl(rpt.Name, ControlTypeValue, , "", fld.Name, x + fldWidth + 50, maxY, fldWidth)
        End If
        ctl.Name = fld.Name
        
        ''Set control property based on ControlTypeValue
        SetControlProperties ctl
        ctl.BorderStyle = 0
    
        Select Case fld.Type
             Case dbMemo:
                 ctl.height = 900
                 isMemo = True
             Case dbDouble:
                 ctl.Format = "Standard"
         End Select
         
        ctl.CanGrow = True
        ctl.InSelection = True
        
        ''Generate the label just above the control
        Dim ControlCaption
        If Not DoesPropertyExists(fld.Properties, "Caption") Then
            If IsNull(rs.fields("VerboseName")) Then
                ControlCaption = AddSpaces(rs.fields("ModelField"))
            Else
                ControlCaption = rs.fields("VerboseName")
            End If
        Else
            ControlCaption = fld.Properties("Caption")
        End If
        
        Set ctl = CreateReportControl(rpt.Name, acLabel, , ctl.Name, ControlCaption & ":", x, maxY)

        SetControlProperties ctl
        ctl.width = fldWidth
        ctl.TextAlign = 2
        ctl.InSelection = True
        ctl.BackStyle = 1


SetVariables:
    
        CurrentCol = CurrentCol + rs.fields("Columns")
        
        x = x + fldWidth
                 
NextField:
        
        rs.MoveNext
        
    Loop
    
    DoCmd.RunCommand acCmdStackedLayout
    DoCmd.RunCommand acCmdControlPaddingNone
    
    For Each ctl In rpt.Controls
        If ctl.InSelection Then
            ctl.left = 0
            Exit For
        End If
    Next ctl
    
    Dim rptWidth
    rptWidth = (8.5 - (0.25 * 2)) * 1440
    
    ''Write the pageheader (The Report Title & Current Date Time)
    Set ctl = CreateReportControl(rpt.Name, acLabel, acPageHeader, , , 0, 0, 4800, 400)
    SetControlProperties ctl
    ctl.fontSize = 12
    ctl.height = 1000
    ctl.Caption = AddSpaces(GetModelPlural(Model, VerbosePlural, ""))
    
    ''Current Date Time
    fldWidth = 3000
    Set ctl = CreateReportControl(rpt.Name, acTextBox, acPageHeader, , , rptWidth - fldWidth, 0, fldWidth, 400)
    SetControlProperties ctl
    ctl.ControlSource = "=Now()"
    ctl.BorderStyle = 0
    
    ''At pagefooter (The Page x of y)
    Set ctl = CreateReportControl(rpt.Name, acTextBox, acPageFooter, , , 0, 0, fldWidth, 400)
    SetControlProperties ctl
    ctl.ControlSource = "=""Page "" & [Page] & "" of "" & [Pages]"
    ctl.BorderStyle = 0
    
    rpt.Section(acHeader).height = 0
    rpt.Section(acHeader).BackColor = RGB(254, 254, 254)
    rpt.Section(acFooter).height = 0
    rpt.Section(acPageFooter).height = 0
    rpt.Section(acPageHeader).height = 0
    rpt.Section(acDetail).height = 0
    rpt.Section(acDetail).AlternateBackColor = RGB(254, 254, 254)
    rpt.Section(acDetail).BackColor = RGB(254, 254, 254)
    
    Dim frmName As String, customFrmName As String, baseFormName As String, i As Integer
    frmName = rpt.Name
    If Not IsNull(VerbosePlural) Then
        baseFormName = concat("rptCol", replace(VerbosePlural, " ", ""))
    Else
        baseFormName = concat("rptCol", Model, "s")
    End If
    
    If Not IsNull(subformName) Then
        baseFormName = concat("rptCol", subformName)
    End If
    
    ''Special Function to run on form creation
    If Not IsNull(OnFormCreate) Then
        Run OnFormCreate, rpt, 7
    End If
    
    ''Override
    OverrideProperties ModelID, 7, rpt
    
    DoCmd.Close acReport, rpt.Name, acSaveYes
    
    customFrmName = baseFormName
    
'    Do Until Not RptExist(customFrmName)
'
'        If MsgBox(customFrmName & " already exists. Would you like to replace it?", vbYesNo) = vbYes Then
'            Exit Do
'        End If
'        i = i + 1
'        customFrmName = baseFormName & "_" & i
'
'    Loop
    
    
    ''Insert the newly created form to the InsertToModelRelatedObjects
    InsertToModelRelatedObjects ModelID, acReport, customFrmName
    
    DoCmd.Rename customFrmName, acReport, frmName
    DoCmd.OpenReport customFrmName, acViewPreview
    
End Function
 
Public Function GetTableName(Model, VerbosePlural, Optional QueryName = Null) As String
    
    If Not IsNull(QueryName) Then
        GetTableName = QueryName
        Exit Function
    End If
    
    If Not IsNull(VerbosePlural) And Not VerbosePlural = "" Then
        GetTableName = concat("tbl", replace(VerbosePlural, " ", ""))
    Else
        GetTableName = concat("tbl", Model, "s")
    End If
    
End Function

Public Function GetTableNameFromModelID(ModelID)
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblModels WHERE ModelID = " & ModelID)
    
    Dim QueryName, Model, VerbosePlural
    QueryName = rs.fields("QueryName")
    Model = rs.fields("Model")
    VerbosePlural = rs.fields("VerbosePlural")
    
    If Not rs.EOF Then
        If Not IsNull(QueryName) Then
            GetTableNameFromModelID = QueryName
            Exit Function
        End If
        
        If Not IsNull(VerbosePlural) And Not VerbosePlural = "" Then
            GetTableNameFromModelID = concat("tbl", replace(VerbosePlural, " ", ""))
        Else
            GetTableNameFromModelID = concat("tbl", Model, "s")
        End If
    End If
    
    rs.Close
    Set rs = Nothing
    
End Function

Public Function GetModelPlural(Model, VerbosePlural, Optional prefix = "tbl") As String

    If Not IsNull(VerbosePlural) And Not VerbosePlural = "" Then
        GetModelPlural = concat(prefix, replace(VerbosePlural, " ", ""))
    Else
        GetModelPlural = concat(prefix, Model, "s")
    End If
    
End Function


Public Sub CreatePrimaryKey(Model, tblDef As TableDef)
    
    Dim pkName, fld As DAO.Field, idx As DAO.index
    pkName = concat(Model, "ID")
    Set fld = AddField(tblDef, pkName, dbLong, dbAutoIncrField)
    
    If Not DoesPropertyExists(tblDef.Indexes, pkName) Then
        Set idx = tblDef.CreateIndex(pkName)
    
        With idx
            .fields.Append .CreateField(pkName)
            .Primary = True
        End With
        
        tblDef.Indexes.Append idx
    End If

End Sub

Public Function GetFieldName(ForeignKey, ModelField, Optional ByPassFK As Boolean = False) As String
    
    If IsNull(ForeignKey) Then
        GetFieldName = ModelField
    Else
        If ByPassFK Then
            GetFieldName = ModelField
        Else
            GetFieldName = concat(ForeignKey, "ID")
        End If
    End If

End Function

Public Function GetFieldCaption(VerboseName, fldName) As String

    If IsNull(VerboseName) Then
        GetFieldCaption = AddSpaces(fldName)
    Else
        GetFieldCaption = VerboseName
    End If
    
End Function

Public Function AddField(tblDef As TableDef, fldName, fldType, Optional fldAttr As Variant) As DAO.Field
    
    Dim fld As DAO.Field
    
    If Not DoesPropertyExists(tblDef.fields, fldName) Then
    
        Set fld = tblDef.CreateField(fldName, fldType)
        
        If Not IsMissing(fldAttr) Then
            fld.Attributes = fld.Attributes Or fldAttr
        End If
        
        With tblDef.fields
            .Append fld
            .refresh
        End With
    Else
        
        Set fld = tblDef.fields(fldName)
        
    End If
    
    Set AddField = fld
    
End Function

Public Function SaveFormData2(frm As Form, Model As String) As Boolean

    If areDataValid2(frm, Model) Then
    
'        If validationSuccessCB <> "" Then
'            Run validationSuccessCB, frm
'        End If
        
        If Not frm.NewRecord Then
            UpdateFormData2 frm, Model
        End If
        
    End If
    
End Function

Public Function GetModelByPrimaryKey(PrimaryKey) As Recordset
    
    Dim Model, ModelID
    Model = left(PrimaryKey, Len(PrimaryKey) - 2)
    ModelID = ELookup("tblModels", "Model = " & EscapeString(Model), "ModelID")
    
    Set GetModelByPrimaryKey = ReturnRecordset("SELECT * FROM tblModels WHERE ModelID = " & ModelID)
    
End Function

Public Function GetTableNameByPrimaryKey(PrimaryKey)

    Dim Model, ModelID
    Model = left(PrimaryKey, Len(PrimaryKey) - 2)
    ModelID = ELookup("tblModels", "Model = " & EscapeString(Model), "ModelID")
    
    GetTableNameByPrimaryKey = GetTableNameFromModelID(ModelID)
    
End Function


Public Function UpdateFormData2(frm As Form, Model As String)
    
    Dim rs As Recordset, ModelID, tblName As String
    Set rs = ReturnRecordset("SELECT * FROM tblModels WHERE Model = """ & Model & """")
    tblName = GetTableName(Model, rs.fields("VerbosePlural"))
    ModelID = rs.fields("ModelID")
    
    Set rs = ReturnRecordset("SELECT * FROM tblModelFields WHERE ModelID = " & ModelID)
    
    Dim ctl As Control, recordID As Variant
    
    Dim FieldName As String, FieldTypeID As Integer, currentValue, oldValue, PrimaryKey
    Dim updateStatement() As String, i As Integer
    
    Do Until rs.EOF
    
        FieldName = rs.fields("ModelField")
        FieldTypeID = rs.fields("FieldTypeID")
        PrimaryKey = concat(Model, "ID")
        
        If Not DoesPropertyExists(CurrentDb.TableDefs, tblName) Then
            GoTo NextField:
        End If
        
        If ControlExists(FieldName, frm) And DoesPropertyExists(CurrentDb.TableDefs(tblName), FieldName) Then
            ''Get the oldvalue from the table
            recordID = frm(PrimaryKey)
            currentValue = frm(FieldName)
            oldValue = GetOldValue(tblName, FieldName, PrimaryKey, recordID)
            If oldValue <> currentValue Or (Not IsNull(oldValue) Xor Not IsNull(currentValue)) Then
                ReDim Preserve updateStatement(i)
                updateStatement(i) = FieldName & " = " & ReturnStringBasedOnType(currentValue, FieldTypeID)
                Update_Log tblName, oldValue, currentValue, recordID, FieldName
                i = i + 1
            End If
        End If
NextField:
        rs.MoveNext
    Loop
    
End Function

Public Function QueryProperty()

On Error Resume Next
    Dim db As DAO.Database, qDef As DAO.QueryDef, fld As DAO.Field, prop As DAO.property
    
    Set db = CurrentDb
    Set qDef = db.QueryDefs("qryEmployees")
    
    For Each fld In qDef.fields
        'Debug.Print fld.SourceTable
    Next fld
    

End Function

Public Function areDataValid2(frm As Form, Optional Model As String) As Boolean
    
    ''Fetch the tblModelFields from the specific form
    Dim rs As Recordset, ModelID
    Dim ModelField, VerboseName, ValidationString, ControlCaption, PrimaryKey
    ModelID = ELookup("tblModels", "Model = " & EscapeString(Model), "ModelID")
    PrimaryKey = ELookup("tblModels", "ModelID = " & ModelID, "PrimaryKey")
    If PrimaryKey = "" Then PrimaryKey = concat(Model, "ID")
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblModelFields where ModelID = " & ModelID & " ORDER BY FieldOrder")
    
    Dim ValidationArr As New clsArray, ValidationRule As Variant
    Dim ctl As Control
    
    Do Until rs.EOF
        ModelField = rs.fields("ModelField")
        VerboseName = rs.fields("VerboseName")
        ValidationString = rs.fields("ValidationString")
        
        ControlCaption = GetFieldCaption(VerboseName, ModelField)
    
        If ControlExists(ModelField, frm) Then
            Set ctl = frm.Controls(ModelField)
            
            If Not IsNull(ValidationString) Then
                ValidationArr.arr = Split(ValidationString, " ")
                For Each ValidationRule In ValidationArr.arr
                    Select Case Trim(ValidationRule)
                        Case "required":
                            If IsNull(ctl) Or ctl = "" Then
                                ShowError ControlCaption & " is a required field."
                                If ControlExists(ctl.Name, frm) And ctl.ColumnHidden = False Then
                                    ctl.SetFocus
                                End If
                                areDataValid2 = False
                                DoCmd.CancelEvent
                                Exit Function
                            End If
                        Case "+":
                            If ctl < 0 Then
                                ShowError ControlCaption & " must be not be less than 0."
                                If ControlExists(ctl.Name, frm) And ctl.ColumnHidden = False Then
                                    ctl.SetFocus
                                End If
                                areDataValid2 = False
                                DoCmd.CancelEvent
                                Exit Function
                            End If
                    End Select
                Next ValidationRule
            End If
            
        End If
        rs.MoveNext
    Loop
    
    ''Validate Uniqueness of records
    Set rs = CurrentDb.OpenRecordset("SELECT * from tblModelFields where ModelID = " & ModelID & " And ValidationString Like '*unique*'")
    Dim i As Integer
    Dim filterStr() As String
    Dim fieldCaptions() As String
    Dim fieldValue As String
    If Not rs.EOF Then
        Do Until rs.EOF
            fieldValue = frm(rs.fields("ModelField"))
            ReDim Preserve filterStr(i)
            Select Case rs.fields("FieldTypeID")
                Case 10:
                    fieldValue = EscapeString(fieldValue)
            End Select
            filterStr(i) = rs.fields("ModelField") & " = " & fieldValue
            ReDim Preserve fieldCaptions(i)
            fieldCaptions(i) = GetFieldCaption(rs.fields("VerboseName"), rs.fields("ModelField"))
            i = i + 1
            rs.MoveNext
        Loop
        
        Dim filterStmt As String
        filterStmt = Join(filterStr, " And ")
        
        ''If not a new record then disregard this record from the filter
        If Not frm.NewRecord Then
            filterStmt = filterStmt & " And " & PrimaryKey & " <> " & frm(PrimaryKey)
        End If
        
        Dim errorMsg As String
        errorMsg = Join(fieldCaptions, " | ") & " is already present from the record list"
        
        Dim rsName As String
        rsName = GetTableName(Model, ELookup("tblModels", "ModelID = " & ModelID, "VerbosePlural"))
        
        If isPresent(rsName, filterStmt) Then
            ShowError errorMsg
            DoCmd.CancelEvent
            areDataValid2 = False
            Exit Function
        End If
        
    End If
    
    ''Look for AdditionalValidation from tblTables
    Dim TableWideValidation
    TableWideValidation = ELookup("tblModels", "ModelID = " & ModelID, "TableWideValidation")
    If TableWideValidation <> "" Then
        If Not Application.Run(TableWideValidation, frm) Then
            areDataValid2 = False
            DoCmd.CancelEvent
            Exit Function
        End If
    End If
    
    If frm.NewRecord Then
        frm.OnClose = "=RequeryOnClose2('" & Model & "',True)"
    End If
    
    areDataValid2 = True
    
End Function

Public Function RequeryOnClose2(Model As String, Optional shouldRequery As Boolean)

    If Model = "Entity" Then
        
        Dim EntityArr As New clsArray, Entity
        EntityArr.arr = "Buyers,Sellers,Tenants,Contacts"
        For Each Entity In EntityArr.arr
            If IsFormOpen("main" & Entity) Then
                Forms("main" & Entity)("subform").Form.Requery
            End If
        Next Entity
        
    End If

    If shouldRequery Then
        'On Error Resume Next
        Dim requeryForms As Variant, requeryFormArray() As String, requeryForm As Variant, PrimaryKey
        Dim rs As Recordset, frm As Form, rsClone As Recordset, VerbosePlural, PluralizedName, ModelID
        Set rs = ReturnRecordset("SELECT * FROM tblModels WHERE Model = '" & Model & "'")
        
        VerbosePlural = rs.fields("VerbosePlural")
        PrimaryKey = concat(Model, "ID")
        ModelID = rs.fields("ModelID")
        
        If Not IsNull(VerbosePlural) Then
            PluralizedName = concat(replace(VerbosePlural, " ", ""))
        Else
            PluralizedName = concat(Model, "s")
        End If
        
        If IsFormOpen(concat("main", PluralizedName)) Then
            Forms(concat("main", PluralizedName)).subform.Requery
        
            Set frm = Forms(concat("main", PluralizedName)).subform.Form
                
            'ReturnMainForm(requeryForm).SetFocus
            Set rsClone = frm.RecordsetClone
            rsClone.FindFirst PrimaryKey & " = " & frm(PrimaryKey)
            If Not rsClone.NoMatch Then
                frm.Bookmark = rsClone.Bookmark
            End If
        End If
        
        'Eval "Forms!Main" & PluralizedName & ".Requery"
        
        ''Get all the foreign key models of this model
        Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected
        Set sqlObj = New clsSQL
        With sqlObj
            .Source = "tblModelFields"
            .AddFilter "ModelID = " & ModelID & " AND " & _
                       "ParentModelID IS NOT NULL"
            .fields = "ParentModelID"
            .GroupBy = "ParentModelID"
            sqlStr = .sql
        End With
        
        ''SELECT STATEMENT
        Set sqlObj = New clsSQL
        With sqlObj
            .Source = "tblModels"
            .fields = "Model,VerbosePlural"
            .Joins.Add GenerateJoinObj(sqlStr, "ModelID", "temp", "ParentModelID")
            Set rs = .Recordset
        End With
        
        Dim PluralizedName2 As String
        
        Do Until rs.EOF
        
            VerbosePlural = rs.fields("VerbosePlural")
            Model = rs.fields("Model")
            
            If Not IsNull(VerbosePlural) Then
                PluralizedName2 = concat(replace(VerbosePlural, " ", ""))
            Else
                PluralizedName2 = concat(Model, "s")
            End If
            
            Dim subformName
            subformName = concat("sub", PluralizedName)
On Error GoTo SkipForm:

            ''Check here if the form is existing and opened..
            If DoesPropertyExists(Forms, concat("frm", PluralizedName2)) Then
                
                Forms(concat("frm", PluralizedName2))(subformName).Requery
                Set frm = Forms(concat("frm", PluralizedName2))(subformName).Form
            
                'ReturnMainForm(requeryForm).SetFocus
                Set rsClone = frm.RecordsetClone
                rsClone.FindFirst PrimaryKey & " = " & frm(PrimaryKey)
                If Not rsClone.NoMatch Then
                    frm.Bookmark = rsClone.Bookmark
                End If
            End If
            
            
SkipForm:
            rs.MoveNext
        Loop
        
    End If
    
    Exit Function
    
ErrHandler:
    
    If Err.number = 2450 Then
        GoTo SkipForm
    Else
        ShowError concat(Err.number, vbCrLf, Err.Description)
    End If
    
End Function

Public Function SetFormProperties(FormTypeID, frm As Form)

    ''Set the Form Properties
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblFrmProps WHERE FormTypeID = " & FormTypeID)
    Do Until rs.EOF
        frm.Properties(rs.fields("FormProp")) = rs.fields("FormPropValue")
        rs.MoveNext
    Loop
    
End Function


Private Function AddTableDef(db As DAO.Database, tblName) As TableDef
    
    If Not DoesPropertyExists(db.TableDefs, tblName) Then
        Set AddTableDef = db.CreateTableDef(tblName)
    Else
        Set AddTableDef = db.TableDefs(tblName)
    End If
    
End Function

Private Function AddIndex(tblDef As TableDef, idxName, fldName)
    
    Dim idx As DAO.index
    
    If Not DoesPropertyExists(tblDef.Indexes, idxName) Then
        
        Set idx = tblDef.CreateIndex(idxName)
        With idx
            .fields.Append .CreateField(fldName)
        End With
        
        tblDef.Indexes.Append idx
        
    End If
        
End Function

Private Function CreateProperty(fld As DAO.Field, PropertyName, PropertyType, PropertyValue)
    
On Error GoTo Err_Handler:
    
    If Not DoesPropertyExists(fld.Properties, PropertyName) Then
        fld.Properties.Append fld.CreateProperty(PropertyName, PropertyType, PropertyValue)
    Else
        fld.Properties(PropertyName) = PropertyValue
    End If
Err_Handler:
    Exit Function

End Function



'Private Sub cmdCreateTableDef_Click()
'
'    Dim tblName As String
'    Dim db As DAO.Database
'    Dim tblDef As DAO.TableDef, fld As DAO.Field, idx As DAO.Index
'    Dim rel As DAO.Relation, relName, primaryField, foreignField, primaryTable, foreignTable
'    Dim pkName, fldName, idxName, MainField
'    Dim rs As Recordset, fldCaption, rowSourceSQL, ForeignKey, rs2 As Recordset
'
'    ''Get the name of the TableDef
'    tblName = GetTableName(Model, VerbosePlural)
'
'    Set db = CurrentDb
'    Set tblDef = AddTableDef(db, tblName)
'
'    ''Create the table fields using
'    ''tblDef.fields.Append .CreateField("FirstName", dbText)
'    ''If the field already exists then skip if not then append
'    ''First create the primary key (this is an autonumber field)
'    CreatePrimaryKey Model, tblDef
'
'    ''Add Custom Fields here via loop
'    Set rs = ReturnRecordset("SELECT * FROM tblModelFields WHERE ModelID = " & ModelID & " ORDER BY FieldOrder ASC")
'
'    Do Until rs.EOF
'
'        fldName = GetFieldName(rs.Fields("ForeignKey"), rs.Fields("ModelField"))
'
'
'        Set fld = AddField(tblDef, fldName, rs.Fields("FieldTypeID"))
'
'        rs.MoveNext
'    Loop
'
'    ''Also add the Timestamp and CreatedBy field
'    ''Timestamp Field
'    ''Created by will be set into a combo box looked up into tblUsers with Username as its field
'    ''And also create index for this field..
'    fldName = "Timestamp"
'    Set fld = AddField(tblDef, fldName, dbDate)
'    AddIndex tblDef, fldName, fldName
'
'    ''CreatedBy Field
'    fldName = "CreatedBy"
'    Set fld = AddField(tblDef, fldName, dbLong)
'    AddIndex tblDef, fldName, fldName
'
'    ''RecordImportID Field
'    fldName = "RecordImportID"
'    Set fld = AddField(tblDef, fldName, dbLong)
'    AddIndex tblDef, fldName, fldName
'
'    If Not DoesPropertyExists(db.TableDefs, tblName) Then
'        db.TableDefs.Append tblDef
'    End If
'
'    ''Set field properties here
'    rs.MoveFirst
'    Do Until rs.EOF
'
'
'        fldName = GetFieldName(rs.Fields("ForeignKey"), rs.Fields("ModelField"))
'        fldCaption = GetFieldCaption(rs.Fields("VerboseName"), fldName)
'
'        Set fld = tblDef.Fields(fldName)
'
'        ''Set the Caption
'        CreateProperty fld, "Caption", dbText, fldCaption
'
'        ''Set the index
'        If rs.Fields("IsIndexed") Then
'            AddIndex tblDef, fldName, fldName
'        End If
'
'        ''Set the foreign key
'        If Not IsNull(rs.Fields("ForeignKey")) Then
'
'            CreateProperty fld, "DisplayControl", dbInteger, acComboBox
'
'            ForeignKey = rs.Fields("ForeignKey")
'            MainField = ELookup("tblModels", Concat("Model = ", EscapeString(ForeignKey)), "MainField")
'            foreignTable = tblName: foreignField = Concat(ForeignKey, "ID")
'
'            Set rs2 = ReturnRecordset("SELECT * FROM tblModels WHERE Model = " & EscapeString(ForeignKey))
'
'            ''Get the name of the TableDef
'            primaryTable = GetTableName(rs2.Fields("Model"), rs2.Fields("VerbosePlural"))
'
'            primaryField = Concat(ForeignKey, "ID")
'
'            rowSourceSQL = Concat("SELECT ", primaryField, ",", MainField, " FROM ", primaryTable, " ORDER BY ", MainField)
'
'            CreateProperty fld, "RowSource", dbText, rowSourceSQL
'            CreateProperty fld, "ColumnCount", dbInteger, 2
'            CreateProperty fld, "ColumnWidths", dbText, "0;1"
'
'            ''Create relationship with the primaryTable
'            relName = Concat(primaryTable, primaryField, "_", foreignTable, foreignField)
'            If DoesPropertyExists(db.Relations, relName) Then
'                db.Relations.Delete relName
'            End If
'
'            Set rel = db.CreateRelation(relName, primaryTable, foreignTable, &H2000 + dbRelationUpdateCascade)
'            Set fld = rel.CreateField(primaryField)
'            fld.foreignName = foreignField
'
'            rel.Fields.Append fld
'            db.Relations.Append rel
'
'        End If
'
'        ''Set the default value property
'        If Not IsNull(rs.Fields("DefaultValue")) Then
'
'            fld.Properties("DefaultValue") = rs.Fields("DefaultValue")
'
'        End If
'
'        ''Set the value list
'        If Not IsNull(rs.Fields("PossibleValues")) Then
'
'            CreateProperty fld, "DisplayControl", dbInteger, acComboBox
'            CreateProperty fld, "RowSourceType", dbText, "Value List"
'            CreateProperty fld, "RowSource", dbText, rs.Fields("PossibleValues")
'            CreateProperty fld, "LimitToList", dbBoolean, True
'
'        End If
'
'        ''Default Format
'        If rs.Fields("FieldTypeID") = dbDouble Then
'
'             CreateProperty fld, "Format", dbText, "Standard"
'
'        End If
'
'        rs.MoveNext
'    Loop
'
'    fldName = "Timestamp"
'    Set fld = tblDef.Fields(fldName)
'    fld.Properties("DefaultValue") = "=Now()"
'
'    fldName = "CreatedBy"
'    Set fld = tblDef.Fields(fldName)
'    CreateProperty fld, "Caption", dbText, "Created By"
'    CreateProperty fld, "DisplayControl", dbInteger, acComboBox
'    CreateProperty fld, "RowSource", dbText, "SELECT UserID, UserName FROM tblUsers ORDER BY UserName"
'    CreateProperty fld, "ColumnCount", dbInteger, 2
'    CreateProperty fld, "ColumnWidths", dbText, "0;1"
'
'    idxName = fldName
'    AddIndex tblDef, idxName, fldName
'
'    ''Create relationship with tblUsers
'    foreignTable = tblName: foreignField = "CreatedBy": primaryTable = "tblUsers": primaryField = "UserID"
'    relName = Concat(primaryTable, primaryField, "_", foreignTable, foreignField)
'    If DoesPropertyExists(db.Relations, relName) Then
'        db.Relations.Delete relName
'    End If
'    Set rel = db.CreateRelation(relName, primaryTable, foreignTable, &H2000 + dbRelationUpdateCascade)
'    Set fld = rel.CreateField(primaryField)
'    fld.foreignName = foreignField
'
'    rel.Fields.Append fld
'    db.Relations.Append rel
'
'    MsgBox "Table Def successfully created.."
'
'End Sub

Public Function DeclareVariables(rsName, Optional encloser As String = "frm")
    
    Dim tblDef As Object, db As DAO.Database, fld As DAO.Field
    Set db = CurrentDb
    
    If DoesPropertyExists(db.TableDefs, rsName) Then
        Set tblDef = db.TableDefs(rsName)
    Else
        Set tblDef = db.QueryDefs(rsName)
    End If
    
    ''The Line where the DIM variables are declared
    Dim fieldArr As New clsArray
    For Each fld In tblDef.fields
        Select Case fld.Name
            Case "Timestamp", "CreatedBy", "RecordImportID":
                
            Case Else:
                fieldArr.Add fld.Name
        End Select
    Next fld
    
    Debug.Print "Dim " & fieldArr.JoinArr
    
    Dim fieldItem As Variant
    
    For Each fieldItem In fieldArr.arr
        Select Case fieldItem
            Case "Timestamp", "CreatedBy", "RecordImportID":
                
            Case Else:
                Debug.Print fieldItem & " = " & encloser & "(" & EscapeString(fieldItem) & ")"
        End Select
        
    Next fieldItem
    
End Function

Public Function MakeProductionCopy()
    
    Dim fso As Object
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    ''Make copy of the current database
    ''Make sure that the name is different from the current one
    Dim currentPath, dbName, copyDbName, dbNameArr As New clsArray
    currentPath = CurrentProject.Path
    dbName = CurrentProject.Name
    dbNameArr.arr = Split(dbName, ".")
    copyDbName = concat(dbNameArr.arr(0), "-Prod.", dbNameArr.arr(1))
    
    fso.CopyFile concat(currentPath, "\", dbName), concat(currentPath, "\", copyDbName), True
    
End Function

Public Function RemoveNonSystemTables()
    

    If MsgBox("Are you sure you want to remove all the non-system related objects?", vbYesNo) = vbNo Then
        Exit Function
    End If
    
    DoCmd.Close acForm, "mainModels", acSaveNo
    
    Dim rs As Recordset, rel As DAO.Relation, db As DAO.Database, relName
    Set rs = ReturnRecordset("SELECT * FROM qryModelRelatedObjects WHERE IsSystemTable = 0")
    Set db = CurrentDb
    
On Error GoTo ErrHandler:
    Dim relationArr As New clsArray, modelArr As New clsArray, modelArrItem
    Do Until rs.EOF
        
        If rs.fields("ObjectTypeID") = acTable Then
            For Each rel In db.Relations
                If rel.foreignTable = rs.fields("ObjectName") Then
                    relationArr.Add rel.Name
                End If
                If rel.TABLE = rs.fields("ObjectName") Then
                    relationArr.Add rel.Name
                End If
            Next rel
            
      
        End If
        
        'Debug.Print rs.Fields("ObjectName")
        modelArr.Add rs.fields("ModelID")
        'DoCmd.DeleteObject rs.Fields("ObjectTypeID"), rs.Fields("ObjectName")
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    If relationArr.Count > 0 Then
        For Each relName In relationArr.arr
            If DoesPropertyExists(db.Relations, relName) Then db.Relations.Delete relName
        Next relName
    End If
    
    Set rs = ReturnRecordset("SELECT * FROM qryModelRelatedObjects WHERE IsSystemTable = 0")
    Do Until rs.EOF
        DoCmd.DeleteObject rs.fields("ObjectTypeID"), rs.fields("ObjectName")
        rs.MoveNext
    Loop

    For Each modelArrItem In modelArr.arr
        RunSQL "DELETE FROM tblModels WHERE ModelID = " & modelArrItem
    Next modelArrItem
    
    DoCmd.OpenForm "mainModels"
    
    Exit Function

ErrHandler:
    If Err.number = 7874 Then
       Resume Next
    Else
        MsgBox Err.number & vbCrLf & Err.Description
    End If

End Function

Public Function CreateTableDef(frm As Form)
    
    Dim tblName As String
    Dim db As DAO.Database
    Dim tblDef As DAO.TableDef, fld As DAO.Field, idx As DAO.index
    Dim rel As DAO.Relation, relName, primaryField, foreignField, primaryTable, foreignTable
    Dim pkName, fldName, idxName
    Dim rs As Recordset, fldCaption, rowSourceSQL, ForeignKey, rs2 As Recordset
    
    ''Initiate the form data
    Dim ModelID, Model, VerboseName, VerbosePlural, MainField, TableWideValidation, FormColumns, SetFocus, IsKeyVisible, QueryName, OnFormCreate, Timestamp, CreatedBy
    
    ModelID = frm("ModelID")
    Model = frm("Model")
    VerboseName = frm("VerboseName")
    VerbosePlural = frm("VerbosePlural")
    MainField = frm("MainField")
    TableWideValidation = frm("TableWideValidation")
    FormColumns = frm("FormColumns")
    SetFocus = frm("SetFocus")
    IsKeyVisible = frm("IsKeyVisible")
    QueryName = frm("QueryName")
    OnFormCreate = frm("OnFormCreate")
    Timestamp = frm("Timestamp")
    CreatedBy = frm("CreatedBy")
        
    ''Get the name of the TableDef
    tblName = GetTableName(Model, VerbosePlural)
    
    Set db = CurrentDb
    Set tblDef = AddTableDef(db, tblName)
    
    ''Create the table fields using
    ''tblDef.fields.Append .CreateField("FirstName", dbText)
    ''If the field already exists then skip if not then append
    ''First create the primary key (this is an autonumber field)
    CreatePrimaryKey Model, tblDef
    
    ''Add Custom Fields here via loop
    Set rs = ReturnRecordset("SELECT * FROM tblModelFields WHERE ModelID = " & ModelID & _
                             " AND (FieldSource IS NULL OR FieldSource = " & EscapeString(tblName) & ") " & _
                             "AND IsAnExpression = 0 ORDER BY FieldOrder ASC")
    
    Do Until rs.EOF
        fldName = GetFieldName(rs.fields("ForeignKey"), rs.fields("ModelField"), True)
        Set fld = AddField(tblDef, fldName, rs.fields("FieldTypeID"))
        rs.MoveNext
    Loop
    
    ''Also add the Timestamp and CreatedBy field
    ''Timestamp Field
    ''Created by will be set into a combo box looked up into tblUsers with Username as its field
    ''And also create index for this field..
    fldName = "Timestamp"
    Set fld = AddField(tblDef, fldName, dbDate)
    AddIndex tblDef, fldName, fldName
    
    ''CreatedBy Field
    fldName = "CreatedBy"
    Set fld = AddField(tblDef, fldName, dbLong)
    AddIndex tblDef, fldName, fldName
    
    ''Create the RecordImportID
    fldName = "RecordImportID"
    Set fld = AddField(tblDef, fldName, dbLong)
    AddIndex tblDef, fldName, fldName
    
    If Not DoesPropertyExists(db.TableDefs, tblName) Then
        db.TableDefs.Append tblDef
    End If
    
    ''Set field properties here
    rs.MoveFirst
    Do Until rs.EOF
        
        fldName = rs.fields("ModelField")
        fldCaption = GetFieldCaption(rs.fields("VerboseName"), fldName)
        
        Set fld = tblDef.fields(fldName)
    
        ''Set the Caption
        CreateProperty fld, "Caption", dbText, fldCaption
        
        ''Set the index
        If rs.fields("IsIndexed") Then
            AddIndex tblDef, fldName, fldName
        End If
        
        ''Set the foreign key
        If Not IsNull(rs.fields("ParentModelID")) Then
            
            CreateProperty fld, "DisplayControl", dbInteger, acComboBox
            
            ForeignKey = ELookup("tblModels", "ModelID = " & rs.fields("ParentModelID"), "Model")
            MainField = ELookup("tblModels", "ModelID = " & rs.fields("ParentModelID"), "MainField")
            
            If MainField Like "=*" Then
                MainField = replace(MainField, "=", "")
            End If
            
            foreignTable = tblName: foreignField = fldName
            
            Set rs2 = ReturnRecordset("SELECT * FROM tblModels WHERE ModelID = " & rs.fields("ParentModelID"))
            
            ''Get the name of the TableDef
            primaryTable = GetTableName(rs2.fields("Model"), rs2.fields("VerbosePlural"))
            
            primaryField = concat(ForeignKey, "ID")
            
            rowSourceSQL = concat("SELECT ", primaryField, ",", MainField, " As MainField FROM ", primaryTable, " ORDER BY ", MainField)
            
            CreateProperty fld, "RowSource", dbText, rowSourceSQL
            CreateProperty fld, "ColumnCount", dbInteger, 2
            CreateProperty fld, "ColumnWidths", dbText, "0;1"
            CreateProperty fld, "ListItemsEditForm", dbText, GetModelPlural(rs2.fields("Model"), rs2.fields("VerbosePlural"), "frmSimple")
            
            ''Create relationship with the primaryTable
            relName = left(concat(primaryTable, primaryField, "_", foreignTable, foreignField), 30)
            If DoesPropertyExists(db.Relations, relName) Then
                db.Relations.Delete relName
            End If
            
            Set rel = db.CreateRelation(relName, primaryTable, foreignTable, dbRelationUpdateCascade + dbRelationDeleteCascade)
            Set fld = rel.CreateField(primaryField)
            fld.foreignName = foreignField
 On Error Resume Next
            rel.fields.Append fld
            db.Relations.Append rel
            
        End If
        
        ''If type is Boolean set the display control to checkbox
        If rs.fields("FieldTypeID") = dbBoolean Then
            CreateProperty fld, "DisplayControl", dbInteger, acCheckBox
        End If
        
        If rs.fields("FieldTypeID") = dbText Then
            CreateProperty fld, "AllowZeroLength", dbBoolean, True
        End If
        
        ''Set the default value property
        If Not IsNull(rs.fields("DefaultValue")) Then
            
            fld.Properties("DefaultValue") = rs.fields("DefaultValue")

        End If
        
        ''Set the value list
        If Not IsNull(rs.fields("PossibleValues")) Then
            
            CreateProperty fld, "DisplayControl", dbInteger, acComboBox
            CreateProperty fld, "RowSourceType", dbText, "Value List"
            CreateProperty fld, "RowSource", dbText, rs.fields("PossibleValues")
            CreateProperty fld, "LimitToList", dbBoolean, True
        
        End If
        
        ''Default Format
        If rs.fields("FieldTypeID") = dbDouble Then
        
             CreateProperty fld, "Format", dbText, "Standard"
             
        End If
        
        rs.MoveNext
    Loop
    
    fldName = "Timestamp"
    Set fld = tblDef.fields(fldName)
    fld.Properties("DefaultValue") = "=Now()"
    
    fldName = "CreatedBy"
    Set fld = tblDef.fields(fldName)
    CreateProperty fld, "Caption", dbText, "Created By"
    CreateProperty fld, "DisplayControl", dbInteger, acComboBox
    CreateProperty fld, "RowSource", dbText, "SELECT UserID, UserName FROM tblUsers ORDER BY UserName"
    CreateProperty fld, "ColumnCount", dbInteger, 2
    CreateProperty fld, "ColumnWidths", dbText, "0;1"

    idxName = fldName
    AddIndex tblDef, idxName, fldName
    
    ''Create relationship with tblUsers
On Error Resume Next
    foreignTable = tblName: foreignField = "CreatedBy": primaryTable = "tblUsers": primaryField = "UserID"
    relName = concat(primaryTable, primaryField, "_", foreignTable, foreignField)
    If DoesPropertyExists(db.Relations, relName) Then
        db.Relations.Delete relName
    End If
    Set rel = db.CreateRelation(relName, primaryTable, foreignTable, &H2000 + dbRelationUpdateCascade)
    Set fld = rel.CreateField(primaryField)
    fld.foreignName = foreignField
    rel.fields.Append fld
    db.Relations.Append rel
    
    ''Insert the newly created table to the InsertToModelRelatedObjects
    InsertToModelRelatedObjects ModelID, acTable, tblName
    
    MsgBox "Table Def successfully created.."

End Function


Public Function OpenMainForm(frm As Form)
        
    Dim ModelID, Model, VerboseName, VerbosePlural, MainField, TableWideValidation, FormColumns
    Dim SetFocus, IsKeyVisible, QueryName, OnFormCreate, subformName, UserQueryFields, Timestamp, CreatedBy
    
    ModelID = frm("ModelID")
    If ExitIfTrue(IsNull(ModelID), "Please select a record..") Then Exit Function
    
    Model = frm("Model")
    VerboseName = frm("VerboseName")
    VerbosePlural = frm("VerbosePlural")
    MainField = frm("MainField")
    TableWideValidation = frm("TableWideValidation")
    FormColumns = frm("FormColumns")
    SetFocus = frm("SetFocus")
    IsKeyVisible = frm("IsKeyVisible")
    QueryName = frm("QueryName")
    OnFormCreate = frm("OnFormCreate")
    subformName = frm("SubformName")
    UserQueryFields = frm("UserQueryFields")
    Timestamp = frm("Timestamp")
    CreatedBy = frm("CreatedBy")
    
    ''Open the mainForm of the record selected
    Dim MainFormName
    ''Check first if SubformName is not Null
    If Not IsNull(subformName) Then
        MainFormName = concat("main", subformName)
GoTo OpenMainForm:
    End If
    
    If Not IsNull(VerbosePlural) Then
        MainFormName = concat("main", RemoveSpaces(VerbosePlural))
    Else
        MainFormName = concat("main", Model, "s")
    End If
    
OpenMainForm:
    DoCmd.OpenForm MainFormName
    
End Function

Public Function CreateDEUploadForm(frm As Form, ModelID)
    
    Dim ModelFieldID
    ModelFieldID = ELookup("qryModelFieldProperties", "ModelID = " & ModelID & " And Property = " & EscapeString("imageType"), "ModelFieldID")
    
    If ModelFieldID = "" Then Exit Function
    
    Dim modelFieldRs As Recordset
    Set modelFieldRs = ReturnRecordset("SELECT * FROM tblModelFields WHERE ModelFieldID = " & ModelFieldID)
    
    Dim x, y: x = GetMaxX(frm): y = 600
    
    ''Render the image control
    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acImage, , , , x + 200, y, 3000, 3000)
    ctl.Name = concat(modelFieldRs.fields("ModelField"), "Img")
    SetControlProperties ctl
    ctl.Picture = "placeholder"
    
    Dim ControlCaption
    If IsNull(modelFieldRs.fields("VerboseName")) Then
        ControlCaption = AddSpaces(modelFieldRs.fields("ModelField"))
    Else
        ControlCaption = modelFieldRs.fields("VerboseName")
    End If
    
    Set ctl = CreateControl(frm.Name, acCommandButton, , , , x + 200, y - 300)
    SetControlProperties ctl
    ctl.Name = concat("cmd", modelFieldRs.fields("ModelField"))
    ctl.Caption = concat("Upload ", ControlCaption)
    ctl.OnClick = "=UploadImage([Form]," & EscapeString(modelFieldRs.fields("ModelField")) & ")"
    ctl.height = 300
    ctl.width = 3000
    
    ''Also render the textbox control of the ModelField
    Set ctl = CreateControl(frm.Name, acTextBox, , , modelFieldRs.fields("ModelField"), 0, 0, 3000, 3000)
    SetControlProperties ctl
    ctl.Name = modelFieldRs.fields("ModelField")
    ctl.Visible = False
    
End Function

Public Function FollowFormHyperlink(frm, FieldName, Optional WithStreetAddress As Boolean = False, Optional AbsoluteLink = Null)
    
    Dim fileName, PropertyListID
    If isFalse(AbsoluteLink) Then fileName = frm(FieldName)
    If WithStreetAddress Then PropertyListID = frm("PropertyListID")
    
    If IsNull(fileName) Then
        ShowError "The hyperlink is empty..."
        Exit Function
    End If
    
    Dim assetDir, fs As Object, uploadDirectory
    
    uploadDirectory = GetAttachmentsDirectory
    
    If WithStreetAddress Then
        Dim StreetAddress
        StreetAddress = ELookup("tblPropertyList", "PropertyListID = " & PropertyListID, "StreetAddress")
        uploadDirectory = GetAttachmentsDirectory(StreetAddress)
    End If
    
    assetDir = uploadDirectory
    
    ''Override assetDir value with AbsoluteLink if it's not Null
    If Not isFalse(AbsoluteLink) Then
        assetDir = AbsoluteLink
    End If
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim filePath
    filePath = concat(assetDir, fileName)
    
    If Not fs.fileExists(filePath) Then
        MsgBox "File does not exist at: " & EscapeString(filePath)
        Exit Function
    End If
    
    On Error Resume Next
    FollowHyperlink filePath
    
End Function

Public Function SelectDirectory(frm As Form, ModelField)
    
    Dim fs As Object

    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim fullPath
    With FileDialog(msoFileDialogFolderPicker)
        'Makes sure the user can select only one file
        .AllowMultiSelect = False
        'Show the dialog box
        If ExitIfTrue(Not .Show, "Directory Selection cancelled..") Then Exit Function
        
        'Store in fullpath variable
        fullPath = .SelectedItems.item(1)
        
    End With
    
    frm(ModelField) = fullPath
    
End Function

Public Function UploadMultiFile(frm As Form, ModelField, Optional WithStreetAddress As Boolean = False)

    Dim FileType, EntityID, PropertyListID
    FileType = frm("FileType")
    EntityID = frm("EntityID")
    PropertyListID = frm("PropertyListID")
    
    If ExitIfTrue(isFalse(FileType), "Select a valid file type..") Then Exit Function
    If ExitIfTrue(isFalse(EntityID), "One of the required fields is empty..") Then Exit Function
    
    Dim uploadDirectory, strFolderExists
    
    uploadDirectory = GetAttachmentsDirectory
    
    If WithStreetAddress Then
        Dim StreetAddress
        StreetAddress = ELookup("tblPropertyList", "PropertyListID = " & PropertyListID, "StreetAddress")
        uploadDirectory = GetAttachmentsDirectory(StreetAddress)
    End If
    
   
    strFolderExists = Dir(uploadDirectory, vbDirectory)
    
    ''Create the directory if it doesn't exist
    If strFolderExists = "" Then
        MkDir uploadDirectory
    End If
    
    Dim fs As Object, filePath, fileName
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim FileSelected
    With FileDialog(msoFileDialogFilePicker)
        'This will allow multi file selection
        .AllowMultiSelect = True
        .InitialFileName = uploadDirectory
        .filters.Clear
        If ExitIfTrue(Not .Show, "File upload cancelled..") Then Exit Function
        
        For Each FileSelected In .SelectedItems
            
           filePath = FilenameToBeUploaded(FileSelected, uploadDirectory, fileName)
           
           fs.CopyFile FileSelected, filePath, True
        
           InsertToEntityFiles fileName, FileType, EntityID, PropertyListID
           
           frm(ModelField) = filePath
            
        Next
        
    End With

End Function

Private Function InsertToEntityFiles(filePath, FileType, EntityID, PropertyListID)
    
    RunSQL "DELETE FROM tblEntityFiles WHERE EntityFileLink = " & EscapeString(filePath)
    RunSQL "INSERT INTO tblEntityFiles (EntityID,FileType,EntityFileLink,PropertyListID) VALUES (" & EntityID & "," & EscapeString(FileType) & "," & EscapeString(filePath) & "," & PropertyListID & ")"
    
End Function

Private Function FilenameToBeUploaded(FileSelected, uploadDirectory, fileName)
    
    With CreateObject("Scripting.FileSystemObject")
    
        Dim extName, baseName, fileExists
        fileName = .GetFileName(FileSelected)
        extName = .GetExtensionName(FileSelected)
        baseName = .GetBaseName(FileSelected)
        
        FilenameToBeUploaded = uploadDirectory & fileName
        
'        '''Check if the file already exists to the upload directory
'        fileExists = Dir(FilenameToBeUploaded)
'
'        If fileExists <> "" Then
'            ''Change the file name
'            fileName = baseName & Format(Now(), "_yyyy_mm_dd_hh_MM_ss") & "." & extName
'            FilenameToBeUploaded = uploadDirectory & fileName
'        End If
 
    End With
    
    Debug.Print FilenameToBeUploaded
    
End Function

Public Function GetAttachmentsDirectory(Optional StreetAddress = Null)
    
    GetAttachmentsDirectory = CurrentProject.Path & "\Files\"
    If Environ("computername") <> "DESKTOP-3G3V8GO" Then GetAttachmentsDirectory = "Z:\MY PANDA APP\Attachments\"
    
    If Not isFalse(StreetAddress) Then
        StreetAddress = replace(StreetAddress, "\", " ")
        StreetAddress = replace(StreetAddress, "/", " ")
        GetAttachmentsDirectory = GetAttachmentsDirectory & StreetAddress & "\"
    End If
    
End Function

Public Function UploadFile(frm As Form, ModelField)
    
    Dim assetDir, fs As Object
    assetDir = GetApplicationSetting("Asset Directory")
    
    If assetDir = "" Then assetDir = GetAttachmentsDirectory
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FolderExists(assetDir) Then assetDir = CurrentProject.Path
    
    Dim fullPath, fileName
    With FileDialog(msoFileDialogFilePicker)
        'Makes sure the user can select only one file
        .AllowMultiSelect = False
        .filters.Clear
        'Show the dialog box
        If ExitIfTrue(Not .Show, "File upload cancelled..") Then Exit Function
        
        'Store in fullpath variable
        fullPath = .SelectedItems.item(1)
        fileName = fs.GetFileName(fullPath)
        
    End With
    
    fs.CopyFile fullPath, concat(assetDir, fileName), True
    
    frm(ModelField) = fileName
    
End Function

Public Function UploadImage(frm As Form, ModelField)
    
    Dim assetDir, fs As Object
    assetDir = GetApplicationSetting("Asset Directory")
    
    If assetDir = "" Then assetDir = CurrentProject.Path
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FolderExists(assetDir) Then assetDir = CurrentProject.Path
    
    Dim fullPath, fileName
    With FileDialog(msoFileDialogFilePicker)
        'Makes sure the user can select only one file
        .AllowMultiSelect = False
        'Filter to just the following types of files to narrow down selection options
        .filters.Add "Image Files", "*.jpg; *.png", 1
        'Show the dialog box
        If ExitIfTrue(Not .Show, "File upload cancelled..") Then Exit Function
        
        'Store in fullpath variable
        fullPath = .SelectedItems.item(1)
        fileName = fs.GetFileName(fullPath)
        
    End With
    
    fs.CopyFile fullPath, concat(assetDir, "\", fileName), True
    
    frm(ModelField) = fileName
    frm(concat(ModelField, "Img")).Picture = concat(assetDir, "\", fileName)
    
End Function

Public Function CreateDEFileUploadControl(frm As Form, fldName, ByVal x, ByVal y, fldWidth)
    
    Dim proportionArr As New clsArray, controlArr As New clsArray, proportionTotal, totalWidth, colSpaceWidth, i, proportion
    
    colSpaceWidth = 50
    totalWidth = fldWidth
    
    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acLabel, , fldName, "Select " & AddSpaces(fldName), x, y - 300)
    SetControlProperties ctl
    ctl.width = totalWidth
    
    Set ctl = CreateControl(frm.Name, acTextBox, , , fldName, 0, 0, 0, 300) ''Texbox Portion
    SetControlProperties ctl
    ctl.Name = fldName
    ctl.OnClick = "=FollowFormHyperlink([Form]," & EscapeString(fldName) & ")"
    
    Set ctl = CreateControl(frm.Name, acCommandButton, , , , 0, 0) ''Button Portion
    SetControlProperties ctl
    ctl.Name = concat("cmd", fldName)
    ctl.Caption = "Browse..."
    ctl.OnClick = "=UploadFile([Form]," & EscapeString(fldName) & ")"
    ctl.height = 300
    
    ''Render the Filter buttons
    ''Filter and Clear
    proportionArr.arr = "10,2"
    controlArr.arr = fldName & "," & concat("cmd", fldName)
    proportionTotal = GetProportionTotal(proportionArr)
    
    For i = 0 To proportionArr.Count - 1
    
        proportion = CDbl(proportionArr.arr(i)) / proportionTotal
        frm(controlArr.arr(i)).left = x
        frm(controlArr.arr(i)).top = y
        frm(controlArr.arr(i)).width = (totalWidth - ((proportionArr.Count - 1) * colSpaceWidth * 2)) * proportion
        frm(controlArr.arr(i)).HorizontalAnchor = acHorizontalAnchorRight
       
        x = x + (colSpaceWidth * 2) + frm(controlArr.arr(i)).width
       
    Next i
    
    
End Function


Public Function CreateDEForm(frm2 As Form)
    
    Dim frm As Form, rs As Recordset, rsName, frmCaption, fldName, fldWidth
    Dim CurrentCol, x, y, isMemo As Boolean, maxWidth As Double
    Dim ctl As Control
    
    Dim ModelID, Model, VerboseName, VerbosePlural, MainField, TableWideValidation, FormColumns, SetFocus, IsKeyVisible
    Dim QueryName, OnFormCreate, SubformName2, UserQueryFields, Timestamp, CreatedBy, PrimaryKey
    
    ModelID = frm2("ModelID")
    Model = frm2("Model")
    VerboseName = frm2("VerboseName")
    VerbosePlural = frm2("VerbosePlural")
    MainField = frm2("MainField")
    TableWideValidation = frm2("TableWideValidation")
    FormColumns = frm2("FormColumns")
    SetFocus = frm2("SetFocus")
    IsKeyVisible = frm2("IsKeyVisible")
    QueryName = frm2("QueryName")
    OnFormCreate = frm2("OnFormCreate")
    SubformName2 = frm2("SubformName")
    UserQueryFields = frm2("UserQueryFields")
    Timestamp = frm2("Timestamp")
    CreatedBy = frm2("CreatedBy")
    PrimaryKey = frm2("PrimaryKey")
    
    If ExitIfTrue(IsNull(ModelID), "Selection is empty..") Then Exit Function
    
    GenerateFields frm2
    
    '''Create the controls
    '''Fields, default buttons, additional buttons and controls
    
    ''Create the form
    Set frm = CreateForm
    rsName = GetTableName(Model, VerbosePlural)
    
    Dim tblName
    tblName = rsName
    
    If Not IsNull(QueryName) Then
        rsName = QueryName
    End If
    
    frm.Detail.BackColor = RGB(81, 163, 36)
    frm.RecordSource = rsName
    frmCaption = GetFieldCaption(VerboseName, Model)
    frm.Caption = concat(frmCaption, " Form")
    
    frm.OnCurrent = "=SetFocusOnForm([Form],""" & SetFocus & """)"
    
    If IsNull(PrimaryKey) Then PrimaryKey = concat(Model, "ID")
    frm.BeforeUpdate = "=SaveFormData2([Form],""" & Model & """)"
    frm.OnLoad = "=DefaultFormLoad([Form]," & EscapeString(PrimaryKey) & ")"
    
    SetFormProperties 4, frm
    CurrentCol = 1
    
    Dim rsObj As Object, db As DAO.Database, fld As DAO.Field
    Set db = CurrentDb
    If DoesPropertyExists(db.TableDefs, rsName) Then
        Set rsObj = db.TableDefs(rsName)
    Else
        Set rsObj = db.QueryDefs(rsName)
    End If
    
    x = 400
    y = 600
    
    CurrentCol = 0: isMemo = False
    
    Dim sqlStr As String
    sqlStr = "SELECT * FROM tblModelFields WHERE FieldOrder IS NOT NULL AND ModelID = " & ModelID
    
    If Not UserQueryFields Then
        sqlStr = sqlStr & " AND FieldSource = " & EscapeString(rsName)
    End If
    
    Set rs = ReturnRecordset(sqlStr & " ORDER BY FieldOrder ASC")
    
    Do Until rs.EOF
    
        ''Skip the field if there is an imageType proprety on qryModelFieldProperties
        If isPresent("qryModelFieldProperties", "Property = " & EscapeString("imageType") & " And ModelFieldID = " & rs.fields("ModelFieldID")) Then
            GoTo NextField
        End If
    
        ''Look First if the ControlSource is not Null
        If Not IsNull(rs.fields("ControlSource")) Then
           If isPresent("qryModelFieldProperties", "Property = ""onDatasheetOnly"" And ModelFieldID = " & rs.fields("ModelFieldID")) Then
               
                GoTo NextField
            Else
                ''Create a control at the footer of the form
                fldWidth = 3000 * (rs.fields("Columns")) + (200 * (rs.fields("Columns") - 1))
                Set ctl = CreateControl(frm.Name, acTextBox, , "", "", x, y, fldWidth)
                SetControlProperties ctl
                ctl.Name = rs.fields("ModelField")
                ctl.ControlSource = rs.fields("ControlSource")
                
                Select Case rs.fields("FieldTypeID")
                    Case dbMemo:
                        ctl.height = 900
                        isMemo = True
                    Case dbDouble
                        ctl.Format = "Standard"
                End Select
            
                ''Generate the label just above the control
                Dim CustomVerboseName
                If IsNull(rs.fields("VerboseName")) Then
                    CustomVerboseName = AddSpaces(rs.fields("ModelField"))
                Else
                    CustomVerboseName = rs.fields("VerboseName")
                End If
                Set ctl = CreateControl(frm.Name, acLabel, , rs.fields("ModelField"), CustomVerboseName, x, y - 300)
                SetControlProperties ctl
                ctl.width = fldWidth
                
                
                
                GoTo SetVariables
            End If
            
        End If
    
        fldName = GetFieldName(rs.fields("ForeignKey"), rs.fields("ModelField"), True)
        fldWidth = 3000
        
        If Not DoesPropertyExists(rsObj.fields, fldName) Then
            GoTo NextField
        End If
        
        Set fld = rsObj.fields(fldName)
        
        If Not IsKeyVisible And fld.Name = PrimaryKey Then
            GoTo NextField
        End If
        
        Select Case fld.Name
            Case "Timestamp", "CreatedBy":
                GoTo NextField
        End Select
        
        Dim ControlTypeValue
        If Not DoesPropertyExists(fld.Properties, "DisplayControl") Then
            ControlTypeValue = acTextBox
        Else
            ControlTypeValue = fld.Properties("DisplayControl")
        End If
        
        ''Generate the control first before the label
        ''Get the width depending on the number of columns but make sure that the CurrentCol + the columns will not exceed
        ''the FormColumn
        fldWidth = 3000 * (rs.fields("Columns")) + (200 * (rs.fields("Columns") - 1))
        
        ''Setting the FieldOrder to 0 here will make the fld hidden..
        If rs.fields("FieldOrder") = 0 Then
            Set ctl = CreateControl(frm.Name, ControlTypeValue, , "", fld.Name, 0, 0, 0)
            ctl.Name = fld.Name
            SetControlProperties ctl
            GoTo NextField
        End If
        
        ''Check if the current field has fileType property. This will tell us that we need to use an upload form rather than a memo field
        ''Skip the field if there is an imageType proprety on qryModelFieldProperties
        If isPresent("qryModelFieldProperties", "Property = " & EscapeString("fileType") & " And ModelFieldID = " & rs.fields("ModelFieldID")) Then
            CreateDEFileUploadControl frm, fld.Name, x, y, fldWidth
            GoTo SetVariables
        End If
        
        ''Path Only Here
        If isPresent("qryModelFieldProperties", "Property = " & EscapeString("folderType") & " And ModelFieldID = " & rs.fields("ModelFieldID")) Then
            CreateDEFolderControl frm, fld.Name, x, y, fldWidth
            GoTo SetVariables
        End If
        
        Set ctl = CreateControl(frm.Name, ControlTypeValue, , "", fld.Name, x, y, fldWidth)
        ctl.Name = fld.Name
        
        ''Set control property based on ControlTypeValue
        SetControlProperties ctl
        
        Select Case fld.Type
            Case dbMemo:
                ctl.height = 900
                isMemo = True
            Case dbDouble:
                ctl.Format = "Standard"
        End Select
        
        ''Generate the label just above the control
        Dim ControlCaption
        If Not DoesPropertyExists(fld.Properties, "Caption") Then
            If IsNull(rs.fields("VerboseName")) Then
                ControlCaption = AddSpaces(rs.fields("ModelField"))
            Else
                ControlCaption = rs.fields("VerboseName")
            End If
        Else
            ControlCaption = fld.Properties("Caption")
        End If
        
        Set ctl = CreateControl(frm.Name, acLabel, , fld.Name, ControlCaption, x, y - 300)
        SetControlProperties ctl
        ctl.Name = concat("lbl", fld.Name)
        ctl.width = fldWidth

SetVariables:
    
        CurrentCol = CurrentCol + rs.fields("Columns")
        
        If CurrentCol = FormColumns Then
            
            If x + 3000 + 400 > maxWidth Then
                maxWidth = x + 3000 + 400
            End If
            
            CurrentCol = 0
            x = 400
            If Not isMemo Then
                y = y + 700
            Else
                isMemo = False
                y = y + 700 + 600
            End If


        Else
        
            x = x + (3200 * rs.fields("Columns"))
            
            
        End If
NextField:
        
        rs.MoveNext
    Loop
    
    ''Create the Timestamp and CreatedBy field (Hidden Fields)
    Set ctl = CreateControl(frm.Name, acTextBox, , "", "Timestamp", 0, 0, 0)
    ctl.Name = "Timestamp"
    SetControlProperties ctl
    
    Set ctl = CreateControl(frm.Name, acComboBox, , "", "CreatedBy", 0, 0, 0)
    ctl.Name = "CreatedBy"
    SetControlProperties ctl
    
    frm("Timestamp").Visible = False
    frm("CreatedBy").Visible = False
    
    ''Create Upload Form if there's any (Simple PictureBox) + Button at the bottom
    CreateDEUploadForm frm, ModelID
    
    ''Create the child forms here
    ''Child sub + Plural name of the Child Model
    x = 400
    y = y + 600
    
    
    fldWidth = GetMaxX(frm) - 400
    Dim minimumWidth: minimumWidth = (FormColumns * 3000)
    If minimumWidth > fldWidth Then fldWidth = minimumWidth
    
    Dim childModels
    Set rs = ReturnRecordset("SELECT * " & _
        "FROM qryModelFields WHERE ParentModelID = " & ModelID & " And HideSubformFromParent = 0 " & _
        "ORDER BY SubPageorder ASC")
        
    If rs.EOF Then
        childModels = 0
    Else
        rs.MoveLast
        rs.MoveFirst
        childModels = rs.recordCount
    End If
    
    Dim pg As page, tbCtl As TabControl, pgCaption, x1, y1 As Long, subformName, subModel, maxY, pgName, subTblName, ModelFieldID, maxX As Long
    
    Do Until rs.EOF
        
        If Not DoesPropertyExists(frm, "tabCtl") Then
        
            For Each ctl In frm.Controls
                If ctl.top + ctl.height > maxY Then
                    maxY = ctl.top + ctl.height
                End If
            Next ctl
            
            Set ctl = CreateControl(frm.Name, acTabCtl, , , , x, maxY + 400, fldWidth, 7000)
            ctl.Name = "tabCtl"
            SetControlProperties ctl
            
            'frm.Width = (FormColumns * 3000) + (FormColumns * 400) - 200
        End If
        
        Set tbCtl = frm.tabCtl
        
        Do Until tbCtl.Pages.Count > childModels
            tbCtl.Pages.Add
        Loop
        
        For Each pg In tbCtl.Pages
                
            If pg.PageIndex > childModels - 1 Then
            
                tbCtl.Pages.Remove pg.PageIndex
                
            Else

                If Not IsNull(rs.fields("VerbosePlural")) Then
                    pgCaption = AddSpaces(rs.fields("VerbosePlural"))
                    subModel = concat(replace(rs.fields("VerbosePlural"), " ", ""))
                    pgName = concat("pg", subModel)
                    subTblName = concat("tbl", subModel)
                Else
                    pgCaption = AddSpaces(concat(rs.fields("Model"), "s"))
                    subModel = concat(rs.fields("Model"), "s")
                    pgName = concat("pg", subModel)
                    subTblName = concat("tbl", subModel)
                End If
                
                If Not IsNull(rs.fields("VerboseChildName")) Then
                    pgCaption = rs.fields("VerboseChildName")
                    subModel = RemoveSpaces(pgCaption)
                    pgName = concat("pg", subModel)
                End If
                
                pg.Caption = pgCaption
                pg.Name = pgName
                
                ''Add the Buttons
                maxY = frm.Controls("tabCtl").top - 400
                x1 = 600
                y1 = maxY + 400 + 500
                
                ModelFieldID = rs.fields("ModelFieldID")
                
                Dim frmToOpen
                frmToOpen = GetModelPlural(rs.fields("Model"), rs.fields("VerbosePlural"), "frm")
                
                ''New Button
                If Not isPresent("qryModelFieldProperties", "Property = ""pgNewHidden"" And ModelFieldID = " & ModelFieldID) Then
                    RenderButton x1, y1, "New", frm, concat("Add", subModel), pg.Name
                    frm(concat("cmdAdd", subModel)).OnClick = "=OpenFormFromMain(" & EscapeString(frmToOpen) & ","""","""",[Form]," & EscapeString(PrimaryKey) & ")"
                    x1 = x1 + (3200 * 0.46)
                End If
                
                ''Edit/View Button
                If Not isPresent("qryModelFieldProperties", "Property = ""pgEditHidden"" And ModelFieldID = " & ModelFieldID) Then
                    RenderButton x1, y1, "Edit/View", frm, concat("Edit", subModel), pg.Name
                    frm(concat("cmdEdit", subModel)).OnClick = "=OpenFormFromMain(" & EscapeString(frmToOpen) & ", " & EscapeString(concat("sub", subModel)) & ", " & _
                            EscapeString(concat(rs.fields("Model"), "ID")) & ",[Form])"
                    x1 = x1 + (3200 * 0.46)
                End If
                
                ''Delete Button
                If Not isPresent("qryModelFieldProperties", "Property = ""pgDeleteHidden"" And ModelFieldID = " & ModelFieldID) Then
                    RenderButton x1, y1, "Delete", frm, concat("Delete", subModel), pg.Name
                    frm(concat("cmdDelete", subModel)).OnClick = "=DeleteRecord([Form], " & EscapeString(concat(rs.fields("Model"), "ID")) & ", " & _
                            EscapeString(subTblName) & "," & EscapeString(concat("sub", subModel)) & ")"
                    x1 = x1 + (3200 * 0.46)
                End If
                
                
                ''Get the x + length of the leftmost button
                maxY = frm.Controls("tabCtl").top - 400
                For Each ctl In frm.Controls
                    If ctl.ControlType = acCommandButton Then
                        If ctl.top + ctl.height > maxY Then
                            maxY = ctl.top + ctl.height
                        End If
                    End If
                Next ctl
                
                If x1 > 600 Then
                    y1 = y1 + 400 + 100
                End If
                
                x1 = 600
                
                Dim pgCtlHeight, pgCtlTop
                pgCtlHeight = frm.Controls("tabCtl").height
                pgCtlTop = frm.Controls("tabCtl").top
                pgCtlHeight = pgCtlTop + pgCtlHeight - y1 - 200
                
                Set ctl = CreateControl(frm.Name, acSubform, , concat("pg", subModel), , x1, y1, fldWidth - 400, pgCtlHeight)
                ctl.Name = concat("sub", subModel)
                ctl.Properties("RightPadding") = 100
            
        
                If Not IsNull(rs.fields("SubformSource")) Then
                    ctl.SourceObject = rs.fields("SubformSource")
                Else
                    ctl.SourceObject = "dsht" & subModel
                End If
                
                ctl.HorizontalAnchor = acHorizontalAnchorBoth
                ctl.VerticalAnchor = acVerticalAnchorBoth
                
                ''Join the subform using the PrimaryKey
                ctl.LinkMasterFields = PrimaryKey
                ctl.LinkChildFields = rs.fields("ModelField")
                
                ''Option button goes here ===>

                ''GenerateAdditionalOptionButton frm, ModelFieldID, Concat("sub", subModel), pg.Name
                
                            
            End If
            
            If Not rs.EOF Then
            
                rs.MoveNext
            
            End If
            
        Next pg
        
        
    Loop
    
    If childModels > 0 Then
        y = y + 7000
    End If
    
    ''Any subform totals will be placed after the subform
    Set rs = ReturnRecordset("SELECT * FROM tblSubformControls WHERE IsVisible = -1 And ModelID = " & ModelID & " ORDER BY FieldOrder ASC")
    
    If Not rs.EOF Then
        maxY = 0
        CurrentCol = 0
        For Each ctl In frm.Controls
            If (ctl.top + ctl.height) > maxY Then
                maxY = ctl.top + ctl.height
            End If
        Next ctl
        
        x = 400
        y = maxY + 800
        
        Dim ctlName
        
        Do Until rs.EOF
            
            ctlName = concat(rs.fields("SubformName"), rs.fields("ControlName"))
            
            fldWidth = 3000 * (0.5) + (200 * (0.5 - 1))
            Set ctl = CreateControl(frm.Name, acTextBox, , "", "", x, y, fldWidth)
            ctl.Name = ctlName
            ctl.ControlSource = concat("=IfError(", rs.fields("SubformName"), "!SUM", rs.fields("ControlName"), ")")
            ctl.Format = "Standard"
            ctl.HorizontalAnchor = acHorizontalAnchorLeft
            ctl.VerticalAnchor = acVerticalAnchorBottom
            
            ''Set control property based on ControlTypeValue
            SetControlProperties ctl
            
            ''Generate the label just above the control
            Set ctl = CreateControl(frm.Name, acLabel, , ctl.Name, rs.fields("ControlCaption"), x, y - 500)
            
            SetControlProperties ctl
            
            ctl.height = 400
            ctl.width = fldWidth
            ctl.HorizontalAnchor = acHorizontalAnchorLeft
            ctl.VerticalAnchor = acVerticalAnchorBottom
            CurrentCol = CurrentCol + 0.5
        
            If CurrentCol = FormColumns Then
                
                CurrentCol = 0
                x = 400
                y = y + 900
            Else
            
                x = x + (3200 * 0.5)
                
            End If
            
            
            rs.MoveNext
        Loop
    End If
    
    ''Buttons
    maxY = 0
    For Each ctl In frm.Controls
        If (ctl.top + ctl.height) > maxY Then
            maxY = ctl.top + ctl.height
        End If
    Next ctl
    
    x = 400
    y = maxY + 400
    
    Dim buttonMultiplier
    buttonMultiplier = 0.46
    CurrentCol = 0
    ''Cancel Button
    If Not isPresent("qryModelProperties", "Property = ""frmCancelHidden"" And ModelID = " & ModelID) Then
        RenderButton x, y, "Cancel", frm, "Cancel"
        frm.cmdCancel.OnClick = "=CancelEdit([Form])"
        frm.cmdCancel.HorizontalAnchor = 0
        frm.cmdCancel.VerticalAnchor = 1
        x = x + (3200 * buttonMultiplier)
        CurrentCol = CurrentCol + 0.5
    End If
    
    ''New Button
    If Not isPresent("qryModelProperties", "Property = ""frmNewHidden"" And ModelID = " & ModelID) Then
        RenderButton x, y, "New", frm, "New"
        frm.cmdNew.OnClick = "=Save2([Form],'" & Model & "',0)"
        frm.cmdNew.HorizontalAnchor = 0
        frm.cmdNew.VerticalAnchor = 1
        x = x + (3200 * buttonMultiplier)
        CurrentCol = CurrentCol + 0.5
    End If
    
    ''Save Button
    If Not isPresent("qryModelProperties", "Property = ""frmSaveHidden"" And ModelID = " & ModelID) Then
        RenderButton x, y, "Save", frm, "SaveClose"
        frm.cmdSaveClose.OnClick = "=Save2([Form],'" & Model & "',1)"
        frm.cmdSaveClose.HorizontalAnchor = 0
        frm.cmdSaveClose.VerticalAnchor = 1
        x = x + (3200 * buttonMultiplier)
        CurrentCol = CurrentCol + 0.5
    End If
    
    ''Delete Button
    If Not isPresent("qryModelProperties", "Property = ""frmDeleteHidden"" And ModelID = " & ModelID) Then
        RenderButton x, y, "Delete", frm, "Delete"
        frm.cmdDelete.OnClick = "=DeleteRecord([Form], '" & PrimaryKey & "', '" & rsName & "')"
        frm.cmdDelete.HorizontalAnchor = 0
        frm.cmdDelete.VerticalAnchor = 1
        x = x + (3200 * buttonMultiplier)
        CurrentCol = CurrentCol + 0.5
    End If
    
    maxX = x
    
    RenderAdditionalButtonOnDEForm ModelID, frm, y, maxX, CurrentCol, buttonMultiplier, FormColumns
    
    ''Align the collapsed controls if there's any
    If DoesPropertyExists(frm.Controls, "cboFormActions") Then
        maxX = GetMaxX(frm)
        frm("cmdRunFormActions").left = maxX - frm("cmdRunFormActions").width
        frm("cboFormActions").left = frm("cmdRunFormActions").left - 55 - frm("cboFormActions").width
        frm("lblFormActions").left = frm("cboFormActions").left - 55 - frm("lblFormActions").width
    End If
    
'    If Not rs.EOF Then
'
'        ''Get the x + length of the leftmost button
''        For Each ctl In frm.Controls
''            If ctl.ControlType = acCommandButton Then
''                If ctl.Left + ctl.Width > maxX And ctl.Top = y Then
''                    maxX = ctl.Left + ctl.Width
''                End If
''            End If
''        Next ctl
'
'        Do Until rs.EOF
'            ModelButton = rs.Fields("ModelButton")
'            modelButtonName = RemoveSpaces(ModelButton)
'            cmdButtonName = Concat("cmd", modelButtonName)
'            functionName = rs.Fields("FunctionName")
'            TableWideFunction = rs.Fields("TableWideFunction")
'
'            RenderButton maxX, y, ModelButton, frm, modelButtonName
'            If Not IsNull(functionName) Then
'                If TableWideFunction Then
'                    frm(cmdButtonName).OnClick = Concat("=", functionName, "()")
'                Else
'                    frm(cmdButtonName).OnClick = Concat("=", functionName, "([Form])")
'                End If
'            End If
'
'            CurrentCol = CurrentCol + 0.5
'
'            If CurrentCol = FormColumns Then
'
'                CurrentCol = 0
'                maxX = 400
'                y = y + 600
'
'            Else
'
'                maxX = maxX + (3200 * buttonMultiplier)
'
'            End If
'
'            rs.MoveNext
'        Loop
'
'    End If
    
    frm.Section("Detail").height = y + 800
    
    ''Set the Form Width
    maxX = 0
    For Each ctl In frm.Controls
        If ctl.left + ctl.width > maxX Then
            maxX = ctl.left + ctl.width
        End If
    Next ctl
    
    frm.width = maxX + 400
    
    ''Special Function to run on form creation
    If Not IsNull(OnFormCreate) Then
        Run OnFormCreate, frm, 4
    End If
    
    ''Override Properties Here
    OverrideProperties ModelID, 4, frm
    
    Dim frmName As String, customFrmName As String, baseFormName As String, i As Integer
    
    frmName = frm.Name
    If Not IsNull(VerbosePlural) Then
        baseFormName = concat("frm", replace(VerbosePlural, " ", ""))
    Else
        baseFormName = concat("frm", Model, "s")
    End If
    
    If Not IsNull(SubformName2) Then
        baseFormName = concat("frm", SubformName2)
    End If
    
    DoCmd.Close acForm, frm.Name, acSaveYes
    
    customFrmName = baseFormName
    
'    Do Until Not FrmExist(customFrmName)
'
'        If MsgBox(customFrmName & " already exists. Would you like to replace it?", vbYesNo) = vbYes Then
'            Exit Do
'        End If
'        i = i + 1
'        customFrmName = baseFormName & "_" & i
'
'    Loop
    
    InsertToModelRelatedObjects ModelID, acForm, customFrmName
    
    DoCmd.Rename customFrmName, acForm, frmName
    
    ''Load Save Form Layout
    LoadSavedFormLayout customFrmName
 
    DoCmd.OpenForm customFrmName


    InsertFormInFormForRights customFrmName, Model
    

End Function

Private Function GetFormType(customFrmName) As String
    
    Dim frmName
    frmName = customFrmName
    
    GetFormType = "DataEntry"
    
    If frmName Like "main*" Then
    
        GetFormType = "MainForm"
        Exit Function
        
    ElseIf frmName Like "dsht*" Then
            
        GetFormType = "DataSheet"
        Exit Function
            
    End If
    
End Function

Private Function InsertFormInFormForRights(customFrmName, Model)
    
    Dim frmType
    frmType = GetFormType(customFrmName)
    
    
    RunSQL "INSERT INTO tblFormForRights (ModelName, FormName, FormType) VALUES ('" & Model & "','" & customFrmName & "','" & frmType & "')"
    
End Function

Public Function CreateDEFolderControl(frm As Form, fldName, ByVal x, ByVal y, fldWidth)
    
    Dim proportionArr As New clsArray, controlArr As New clsArray, proportionTotal, totalWidth, colSpaceWidth, i, proportion
    
    colSpaceWidth = 50
    totalWidth = fldWidth
    
    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acLabel, , fldName, "Select " & AddSpaces(fldName), x, y - 300)
    SetControlProperties ctl
    ctl.width = totalWidth
    
    Set ctl = CreateControl(frm.Name, acTextBox, , , fldName, 0, 0, 0, 300) ''Texbox Portion
    SetControlProperties ctl
    ctl.Name = fldName
    ctl.OnClick = "=FollowFormHyperlink([Form]," & EscapeString(fldName) & ")"
    
    Set ctl = CreateControl(frm.Name, acCommandButton, , , , 0, 0) ''Button Portion
    SetControlProperties ctl
    ctl.Name = concat("cmd", fldName)
    ctl.Caption = "Browse..."
    ctl.OnClick = "=SelectDEDirectory([Form]," & EscapeString(fldName) & ")"
    ctl.height = 300
    
    ''Render the Filter buttons
    ''Filter and Clear
    proportionArr.arr = "10,2"
    controlArr.arr = fldName & "," & concat("cmd", fldName)
    proportionTotal = GetProportionTotal(proportionArr)
    
    For i = 0 To proportionArr.Count - 1
    
        proportion = CDbl(proportionArr.arr(i)) / proportionTotal
        frm(controlArr.arr(i)).left = x
        frm(controlArr.arr(i)).top = y
        frm(controlArr.arr(i)).width = (totalWidth - ((proportionArr.Count - 1) * colSpaceWidth * 2)) * proportion
        frm(controlArr.arr(i)).HorizontalAnchor = acHorizontalAnchorRight
       
        x = x + (colSpaceWidth * 2) + frm(controlArr.arr(i)).width
       
    Next i
    
End Function

Public Function CreateSimpleDEForm(frm2 As Form)
    
    Dim frm As Form, rs As Recordset, rsName, frmCaption, fldName, fldWidth
    Dim CurrentCol, x, y, isMemo As Boolean, maxWidth As Double
    Dim ctl As Control
    
    Dim ModelID, Model, VerboseName, VerbosePlural, MainField, TableWideValidation, FormColumns, SetFocus, IsKeyVisible
    Dim QueryName, OnFormCreate, SubformName2, UserQueryFields, Timestamp, CreatedBy, PrimaryKey
    
    ModelID = frm2("ModelID")
    Model = frm2("Model")
    VerboseName = frm2("VerboseName")
    VerbosePlural = frm2("VerbosePlural")
    MainField = frm2("MainField")
    TableWideValidation = frm2("TableWideValidation")
    FormColumns = frm2("FormColumns")
    SetFocus = frm2("SetFocus")
    IsKeyVisible = frm2("IsKeyVisible")
    QueryName = frm2("QueryName")
    OnFormCreate = frm2("OnFormCreate")
    SubformName2 = frm2("SubformName")
    UserQueryFields = frm2("UserQueryFields")
    Timestamp = frm2("Timestamp")
    CreatedBy = frm2("CreatedBy")
    PrimaryKey = frm2("PrimaryKey")
    
    If ExitIfTrue(IsNull(ModelID), "Selection is empty..") Then Exit Function
    
    GenerateFields frm2
    
    '''Create the controls
    '''Fields, default buttons, additional buttons and controls
    
    ''Create the form
    Set frm = CreateForm
    frm.DataEntry = True
    rsName = GetTableName(Model, VerbosePlural)
    
    Dim tblName
    tblName = rsName
    
    If Not IsNull(QueryName) Then
        rsName = QueryName
    End If
    
    frm.Detail.BackColor = RGB(81, 163, 36)
    frm.RecordSource = rsName
    frmCaption = GetFieldCaption(VerboseName, Model)
    frm.Caption = concat(frmCaption, " Form")
    
    frm.OnCurrent = "=SetFocusOnForm([Form],""" & SetFocus & """)"
    
    If IsNull(PrimaryKey) Then PrimaryKey = concat(Model, "ID")
    frm.BeforeUpdate = "=SaveFormData2([Form],""" & Model & """)"
    frm.OnLoad = "=DefaultFormLoad([Form]," & EscapeString(PrimaryKey) & ")"
    
    SetFormProperties 4, frm
    CurrentCol = 1
    
    Dim rsObj As Object, db As DAO.Database, fld As DAO.Field
    Set db = CurrentDb
    If DoesPropertyExists(db.TableDefs, rsName) Then
        Set rsObj = db.TableDefs(rsName)
    Else
        Set rsObj = db.QueryDefs(rsName)
    End If
    
    x = 400
    y = 600
    
    CurrentCol = 0: isMemo = False
    
    Dim sqlStr As String
    sqlStr = "SELECT * FROM tblModelFields WHERE FieldOrder IS NOT NULL AND ModelID = " & ModelID
    
    If Not UserQueryFields Then
        sqlStr = sqlStr & " AND FieldSource = " & EscapeString(rsName)
    End If
    
    Set rs = ReturnRecordset(sqlStr & " ORDER BY FieldOrder ASC")
    
    Do Until rs.EOF
    
        ''Skip the field if there is an imageType proprety on qryModelFieldProperties
        If isPresent("qryModelFieldProperties", "Property = " & EscapeString("imageType") & " And ModelFieldID = " & rs.fields("ModelFieldID")) Then
            GoTo NextField
        End If
    
        ''Look First if the ControlSource is not Null
        If Not IsNull(rs.fields("ControlSource")) Then
           If isPresent("qryModelFieldProperties", "Property = ""onDatasheetOnly"" And ModelFieldID = " & rs.fields("ModelFieldID")) Then
               
                GoTo NextField
            Else
                ''Create a control at the footer of the form
                fldWidth = 3000 * (rs.fields("Columns")) + (200 * (rs.fields("Columns") - 1))
                Set ctl = CreateControl(frm.Name, acTextBox, , "", "", x, y, fldWidth)
                SetControlProperties ctl
                ctl.Name = rs.fields("ModelField")
                ctl.ControlSource = rs.fields("ControlSource")
                
                Select Case rs.fields("FieldTypeID")
                    Case dbMemo:
                        ctl.height = 900
                        isMemo = True
                    Case dbDouble
                        ctl.Format = "Standard"
                End Select
            
                ''Generate the label just above the control
                Dim CustomVerboseName
                If IsNull(rs.fields("VerboseName")) Then
                    CustomVerboseName = AddSpaces(rs.fields("ModelField"))
                Else
                    CustomVerboseName = rs.fields("VerboseName")
                End If
                Set ctl = CreateControl(frm.Name, acLabel, , rs.fields("ModelField"), CustomVerboseName, x, y - 300)
                SetControlProperties ctl
                ctl.width = fldWidth
                
                GoTo SetVariables
            End If
            
        End If
    
        fldName = GetFieldName(rs.fields("ForeignKey"), rs.fields("ModelField"), True)
        fldWidth = 3000
        
        If Not DoesPropertyExists(rsObj.fields, fldName) Then
            GoTo NextField
        End If
        
        Set fld = rsObj.fields(fldName)
        
        If Not IsKeyVisible And fld.Name = PrimaryKey Then
            GoTo NextField
        End If
        
        Select Case fld.Name
            Case "Timestamp", "CreatedBy":
                GoTo NextField
        End Select
        
        Dim ControlTypeValue
        If Not DoesPropertyExists(fld.Properties, "DisplayControl") Then
            ControlTypeValue = acTextBox
        Else
            ControlTypeValue = fld.Properties("DisplayControl")
        End If
        
        ''Generate the control first before the label
        ''Get the width depending on the number of columns but make sure that the CurrentCol + the columns will not exceed
        ''the FormColumn
        fldWidth = 3000 * (rs.fields("Columns")) + (200 * (rs.fields("Columns") - 1))
        
        ''Setting the FieldOrder to 0 here will make the fld hidden..
        If rs.fields("FieldOrder") = 0 Then
            Set ctl = CreateControl(frm.Name, ControlTypeValue, , "", fld.Name, 0, 0, 0)
            ctl.Name = fld.Name
            SetControlProperties ctl
            GoTo NextField
        End If
        
        ''Check if the current field has fileType property. This will tell us that we need to use an upload form rather than a memo field
        ''Skip the field if there is an imageType proprety on qryModelFieldProperties
        If isPresent("qryModelFieldProperties", "Property = " & EscapeString("fileType") & " And ModelFieldID = " & rs.fields("ModelFieldID")) Then
            CreateDEFileUploadControl frm, fld.Name, x, y, fldWidth
            GoTo SetVariables
        End If
        
        Set ctl = CreateControl(frm.Name, ControlTypeValue, , "", fld.Name, x, y, fldWidth)
        ctl.Name = fld.Name
        
        ''Set control property based on ControlTypeValue
        SetControlProperties ctl
        
        Select Case fld.Type
            Case dbMemo:
                ctl.height = 900
                isMemo = True
            Case dbDouble:
                ctl.Format = "Standard"
        End Select
        
        ''Generate the label just above the control
        Dim ControlCaption
        If Not DoesPropertyExists(fld.Properties, "Caption") Then
            If IsNull(rs.fields("VerboseName")) Then
                ControlCaption = AddSpaces(rs.fields("ModelField"))
            Else
                ControlCaption = rs.fields("VerboseName")
            End If
        Else
            ControlCaption = fld.Properties("Caption")
        End If
        
        Set ctl = CreateControl(frm.Name, acLabel, , fld.Name, ControlCaption, x, y - 300)
        SetControlProperties ctl
        ctl.Name = concat("lbl", fld.Name)
        ctl.width = fldWidth

SetVariables:
    
        CurrentCol = CurrentCol + rs.fields("Columns")
        
        If CurrentCol = FormColumns Then
            
            If x + 3000 + 400 > maxWidth Then
                maxWidth = x + 3000 + 400
            End If
            
            CurrentCol = 0
            x = 400
            If Not isMemo Then
                y = y + 700
            Else
                isMemo = False
                y = y + 700 + 600
            End If


        Else
        
            x = x + (3200 * rs.fields("Columns"))
            
            
        End If
NextField:
        
        rs.MoveNext
    Loop
    
    ''Create the Timestamp and CreatedBy field (Hidden Fields)
    Set ctl = CreateControl(frm.Name, acTextBox, , "", "Timestamp", 0, 0, 0)
    ctl.Name = "Timestamp"
    SetControlProperties ctl
    
    Set ctl = CreateControl(frm.Name, acComboBox, , "", "CreatedBy", 0, 0, 0)
    ctl.Name = "CreatedBy"
    SetControlProperties ctl
    
    frm("Timestamp").Visible = False
    frm("CreatedBy").Visible = False
    
    ''Create Upload Form if there's any (Simple PictureBox) + Button at the bottom
    CreateDEUploadForm frm, ModelID
    
    ''Create the child forms here
    ''Child sub + Plural name of the Child Model
    x = 400
    y = y + 600
    
    
    fldWidth = GetMaxX(frm) - 400
    Dim minimumWidth: minimumWidth = (FormColumns * 3000)
    If minimumWidth > fldWidth Then fldWidth = minimumWidth
    
    ''This will make the child to be zero
    Dim childModels
    Set rs = ReturnRecordset("SELECT * " & _
        "FROM qryModelFields WHERE ParentModelID = " & ModelID & " And HideSubformFromParent = 0 AND ModelFieldID = 0 " & _
        "ORDER BY SubPageorder ASC")
        
    If rs.EOF Then
        childModels = 0
    Else
        rs.MoveLast
        rs.MoveFirst
        childModels = rs.recordCount
    End If
    
    Dim pg As page, tbCtl As TabControl, pgCaption, x1, y1, subformName, subModel, maxY, pgName, subTblName, ModelFieldID, maxX As Long
    
    Do Until rs.EOF
        
        If Not DoesPropertyExists(frm, "tabCtl") Then
        
            For Each ctl In frm.Controls
                If ctl.top + ctl.height > maxY Then
                    maxY = ctl.top + ctl.height
                End If
            Next ctl
            
            Set ctl = CreateControl(frm.Name, acTabCtl, , , , x, maxY + 400, fldWidth, 7000)
            ctl.Name = "tabCtl"
            SetControlProperties ctl
            
            'frm.Width = (FormColumns * 3000) + (FormColumns * 400) - 200
        End If
        
        Set tbCtl = frm.tabCtl
        
        Do Until tbCtl.Pages.Count > childModels
            tbCtl.Pages.Add
        Loop
        
        For Each pg In tbCtl.Pages
                
            If pg.PageIndex > childModels - 1 Then
            
                tbCtl.Pages.Remove pg.PageIndex
                
            Else

                If Not IsNull(rs.fields("VerbosePlural")) Then
                    pgCaption = AddSpaces(rs.fields("VerbosePlural"))
                    subModel = concat(replace(rs.fields("VerbosePlural"), " ", ""))
                    pgName = concat("pg", subModel)
                    subTblName = concat("tbl", subModel)
                Else
                    pgCaption = AddSpaces(concat(rs.fields("Model"), "s"))
                    subModel = concat(rs.fields("Model"), "s")
                    pgName = concat("pg", subModel)
                    subTblName = concat("tbl", subModel)
                End If
                
                If Not IsNull(rs.fields("VerboseChildName")) Then
                    pgCaption = rs.fields("VerboseChildName")
                    subModel = RemoveSpaces(pgCaption)
                    pgName = concat("pg", subModel)
                End If
                
                pg.Caption = pgCaption
                pg.Name = pgName
                
                ''Add the Buttons
                maxY = frm.Controls("tabCtl").top - 400
                x1 = 600
                y1 = maxY + 400 + 500
                
                ModelFieldID = rs.fields("ModelFieldID")
                
                Dim frmToOpen
                frmToOpen = GetModelPlural(rs.fields("Model"), rs.fields("VerbosePlural"), "frm")
                
                ''New Button
                If Not isPresent("qryModelFieldProperties", "Property = ""pgNewHidden"" And ModelFieldID = " & ModelFieldID) Then
                    RenderButton x1, y1, "New", frm, concat("Add", subModel), pg.Name
                    frm(concat("cmdAdd", subModel)).OnClick = "=OpenFormFromMain(" & EscapeString(frmToOpen) & ","""","""",[Form]," & EscapeString(PrimaryKey) & ")"
                    x1 = x1 + (3200 * 0.46)
                End If
                
                ''Edit/View Button
                If Not isPresent("qryModelFieldProperties", "Property = ""pgEditHidden"" And ModelFieldID = " & ModelFieldID) Then
                    RenderButton x1, y1, "Edit/View", frm, concat("Edit", subModel), pg.Name
                    frm(concat("cmdEdit", subModel)).OnClick = "=OpenFormFromMain(" & EscapeString(frmToOpen) & ", " & EscapeString(concat("sub", subModel)) & ", " & _
                            EscapeString(concat(rs.fields("Model"), "ID")) & ",[Form])"
                    x1 = x1 + (3200 * 0.46)
                End If
                
                ''Delete Button
                If Not isPresent("qryModelFieldProperties", "Property = ""pgDeleteHidden"" And ModelFieldID = " & ModelFieldID) Then
                    RenderButton x1, y1, "Delete", frm, concat("Delete", subModel), pg.Name
                    frm(concat("cmdDelete", subModel)).OnClick = "=DeleteRecord([Form], " & EscapeString(concat(rs.fields("Model"), "ID")) & ", " & _
                            EscapeString(subTblName) & "," & EscapeString(concat("sub", subModel)) & ")"
                    x1 = x1 + (3200 * 0.46)
                End If
                
                ''Get the x + length of the leftmost button
                maxY = frm.Controls("tabCtl").top - 400
                For Each ctl In frm.Controls
                    If ctl.ControlType = acCommandButton Then
                        If ctl.top + ctl.height > maxY Then
                            maxY = ctl.top + ctl.height
                        End If
                    End If
                Next ctl
                
                If x1 > 600 Then
                    y1 = y1 + 400 + 100
                End If
                
                x1 = 600
                
                Dim pgCtlHeight, pgCtlTop
                pgCtlHeight = frm.Controls("tabCtl").height
                pgCtlTop = frm.Controls("tabCtl").top
                pgCtlHeight = pgCtlTop + pgCtlHeight - y1 - 200
                
                Set ctl = CreateControl(frm.Name, acSubform, , concat("pg", subModel), , x1, y1, fldWidth - 400, pgCtlHeight)
                ctl.Name = concat("sub", subModel)
                ctl.Properties("RightPadding") = 100
                
                If Not IsNull(rs.fields("SubformSource")) Then
                    ctl.SourceObject = rs.fields("SubformSource")
                Else
                    ctl.SourceObject = "dsht" & subModel
                End If
                
                ctl.HorizontalAnchor = acHorizontalAnchorBoth
                ctl.VerticalAnchor = acVerticalAnchorBoth
                
                ''Join the subform using the PrimaryKey
                ctl.LinkMasterFields = PrimaryKey
                ctl.LinkChildFields = rs.fields("ModelField")
                
                            
            End If
            
            If Not rs.EOF Then
            
                rs.MoveNext
            
            End If
            
        Next pg
        
        
    Loop
    
    If childModels > 0 Then
        y = y + 7000
    End If
    
    ''Any subform totals will be placed after the subform
    Set rs = ReturnRecordset("SELECT * FROM tblSubformControls WHERE IsVisible = -1 And ModelID = " & ModelID & " ORDER BY FieldOrder ASC")
    
    If Not rs.EOF Then
        maxY = 0
        CurrentCol = 0
        For Each ctl In frm.Controls
            If (ctl.top + ctl.height) > maxY Then
                maxY = ctl.top + ctl.height
            End If
        Next ctl
        
        x = 400
        y = maxY + 800
        
        Dim ctlName
        
        Do Until rs.EOF
            
            ctlName = concat(rs.fields("SubformName"), rs.fields("ControlName"))
            
            fldWidth = 3000 * (0.5) + (200 * (0.5 - 1))
            Set ctl = CreateControl(frm.Name, acTextBox, , "", "", x, y, fldWidth)
            ctl.Name = ctlName
            ctl.ControlSource = concat("=IfError(", rs.fields("SubformName"), "!SUM", rs.fields("ControlName"), ")")
            ctl.Format = "Standard"
            ctl.HorizontalAnchor = acHorizontalAnchorLeft
            ctl.VerticalAnchor = acVerticalAnchorBottom
            
            ''Set control property based on ControlTypeValue
            SetControlProperties ctl
            
            ''Generate the label just above the control
            Set ctl = CreateControl(frm.Name, acLabel, , ctl.Name, rs.fields("ControlCaption"), x, y - 500)
            
            SetControlProperties ctl
            
            ctl.height = 400
            ctl.width = fldWidth
            ctl.HorizontalAnchor = acHorizontalAnchorLeft
            ctl.VerticalAnchor = acVerticalAnchorBottom
            CurrentCol = CurrentCol + 0.5
        
            If CurrentCol = FormColumns Then
                
                CurrentCol = 0
                x = 400
                y = y + 900
            Else
            
                x = x + (3200 * 0.5)
                
            End If
            
            
            rs.MoveNext
        Loop
    End If
    
    ''Buttons
    maxY = 0
    For Each ctl In frm.Controls
        If (ctl.top + ctl.height) > maxY Then
            maxY = ctl.top + ctl.height
        End If
    Next ctl
    
    x = 400
    y = maxY + 400
    
    Dim buttonMultiplier
    buttonMultiplier = 0.46
    CurrentCol = 0
    ''Cancel Button
    If Not isPresent("qryModelProperties", "Property = ""frmCancelHidden"" And ModelID = " & ModelID) Then
        RenderButton x, y, "Cancel", frm, "Cancel"
        frm.cmdCancel.OnClick = "=CancelEdit([Form])"
        frm.cmdCancel.HorizontalAnchor = 0
        frm.cmdCancel.VerticalAnchor = 1
        x = x + (3200 * buttonMultiplier)
        CurrentCol = CurrentCol + 0.5
    End If
    
    ''New Button
    If Not isPresent("qryModelProperties", "Property = ""frmNewHidden"" And ModelID = " & ModelID) Then
        RenderButton x, y, "New", frm, "New"
        frm.cmdNew.OnClick = "=Save2([Form],'" & Model & "',0)"
        frm.cmdNew.HorizontalAnchor = 0
        frm.cmdNew.VerticalAnchor = 1
        x = x + (3200 * buttonMultiplier)
        CurrentCol = CurrentCol + 0.5
    End If
    
    ''Save Button
    If Not isPresent("qryModelProperties", "Property = ""frmSaveHidden"" And ModelID = " & ModelID) Then
        RenderButton x, y, "Save", frm, "SaveClose"
        frm.cmdSaveClose.OnClick = "=Save2([Form],'" & Model & "',1)"
        frm.cmdSaveClose.HorizontalAnchor = 0
        frm.cmdSaveClose.VerticalAnchor = 1
        x = x + (3200 * buttonMultiplier)
        CurrentCol = CurrentCol + 0.5
    End If
    
    ''Delete Button
    If Not isPresent("qryModelProperties", "Property = ""frmDeleteHidden"" And ModelID = " & ModelID) Then
        RenderButton x, y, "Delete", frm, "Delete"
        frm.cmdDelete.OnClick = "=DeleteRecord([Form], '" & PrimaryKey & "', '" & rsName & "')"
        frm.cmdDelete.HorizontalAnchor = 0
        frm.cmdDelete.VerticalAnchor = 1
        x = x + (3200 * buttonMultiplier)
        CurrentCol = CurrentCol + 0.5
    End If
    
    maxX = x
    
    RenderAdditionalButtonOnDEForm ModelID, frm, y, maxX, CurrentCol, buttonMultiplier, FormColumns
    
    ''Align the collapsed controls if there's any
    If DoesPropertyExists(frm.Controls, "cboFormActions") Then
        maxX = GetMaxX(frm)
        frm("cmdRunFormActions").left = maxX - frm("cmdRunFormActions").width
        frm("cboFormActions").left = frm("cmdRunFormActions").left - 55 - frm("cboFormActions").width
        frm("lblFormActions").left = frm("cboFormActions").left - 55 - frm("lblFormActions").width
    End If
    
'    If Not rs.EOF Then
'
'        ''Get the x + length of the leftmost button
''        For Each ctl In frm.Controls
''            If ctl.ControlType = acCommandButton Then
''                If ctl.Left + ctl.Width > maxX And ctl.Top = y Then
''                    maxX = ctl.Left + ctl.Width
''                End If
''            End If
''        Next ctl
'
'        Do Until rs.EOF
'            ModelButton = rs.Fields("ModelButton")
'            modelButtonName = RemoveSpaces(ModelButton)
'            cmdButtonName = Concat("cmd", modelButtonName)
'            functionName = rs.Fields("FunctionName")
'            TableWideFunction = rs.Fields("TableWideFunction")
'
'            RenderButton maxX, y, ModelButton, frm, modelButtonName
'            If Not IsNull(functionName) Then
'                If TableWideFunction Then
'                    frm(cmdButtonName).OnClick = Concat("=", functionName, "()")
'                Else
'                    frm(cmdButtonName).OnClick = Concat("=", functionName, "([Form])")
'                End If
'            End If
'
'            CurrentCol = CurrentCol + 0.5
'
'            If CurrentCol = FormColumns Then
'
'                CurrentCol = 0
'                maxX = 400
'                y = y + 600
'
'            Else
'
'                maxX = maxX + (3200 * buttonMultiplier)
'
'            End If
'
'            rs.MoveNext
'        Loop
'
'    End If
    
    frm.Section("Detail").height = y + 800
    
    ''Set the Form Width
    maxX = 0
    For Each ctl In frm.Controls
        If ctl.left + ctl.width > maxX Then
            maxX = ctl.left + ctl.width
        End If
    Next ctl
    
    frm.width = maxX + 400
    
    ''Special Function to run on form creation
    If Not IsNull(OnFormCreate) Then
        Run OnFormCreate, frm, 4
    End If
    
    ''Override Properties Here
    OverrideProperties ModelID, 4, frm
    
    Dim frmName As String, customFrmName As String, baseFormName As String, i As Integer
    
    frmName = frm.Name
    If Not IsNull(VerbosePlural) Then
        baseFormName = concat("frmSimple", replace(VerbosePlural, " ", ""))
    Else
        baseFormName = concat("frmSimple", Model, "s")
    End If
    
    If Not IsNull(SubformName2) Then
        baseFormName = concat("frm", SubformName2)
    End If
    
    DoCmd.Close acForm, frm.Name, acSaveYes
    
    customFrmName = baseFormName
    
'    Do Until Not FrmExist(customFrmName)
'
'        If MsgBox(customFrmName & " already exists. Would you like to replace it?", vbYesNo) = vbYes Then
'            Exit Do
'        End If
'        i = i + 1
'        customFrmName = baseFormName & "_" & i
'
'    Loop
    
    InsertToModelRelatedObjects ModelID, acForm, customFrmName
    
    DoCmd.Rename customFrmName, acForm, frmName
    
    ''Load Save Form Layout
    LoadSavedFormLayout customFrmName
 
    DoCmd.OpenForm customFrmName
    
    InsertFormInFormForRights customFrmName, Model
    
End Function

Private Sub InsertToModelRelatedObjects(ModelID, ObjectTypeID, objectName)
    
    If Not isPresent("tblModelRelatedObjects", "ObjectName = " & EscapeString(objectName)) Then
        RunSQL "INSERT INTO tblModelRelatedObjects (ModelID, ObjectTypeID, ObjectName) VALUES (" & _
               ModelID & "," & _
               ObjectTypeID & "," & _
               EscapeString(objectName) & ")"
    End If

End Sub

Public Sub SetControlProperties(ctl As Control)
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM qryControlTypes WHERE ControlTypeValue = " & ctl.ControlType)
    
    Do Until rs.EOF
        ctl.Properties(rs.fields("ControlPropValue")) = rs.fields("ControlProp")
        rs.MoveNext
    Loop
    
End Sub

Private Function RenderButton(x, y, Caption, frm As Form, cmdName, Optional Parent As String)
        
    
    Dim ctl As Control, fldWidth
    
    fldWidth = 3000 * (0.5) + (200 * (0.5 - 1))
    
    Set ctl = CreateControl(frm.Name, acCommandButton, , Parent, , x, y, fldWidth)
    With ctl
    
        .Name = "cmd" & cmdName
        .Properties("Caption") = Caption
        
        Dim rs As Recordset
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM qryControlTypes WHERE ControlTypeValue = 104")
        
        Do Until rs.EOF
            If DoesPropertyExists(.Properties, rs.fields("ControlPropValue")) Then
                .Properties(rs.fields("ControlPropValue")) = rs.fields("ControlProp")
            End If
            rs.MoveNext
        Loop
        
        '.Properties("UseTheme") = False
        .Properties("CursorOnHover") = 1
        
    End With
    
End Function

Public Function GetMaxX(frm As Object, Optional atYPosition = Null) As Double
    
    ''MAX X doesn't take into the allowance or margin
    Dim ctl As Control, x As Double
    For Each ctl In frm.Controls
        If ctl.left + ctl.width > x Then
            If IsNull(atYPosition) Then
                x = ctl.left + ctl.width
            Else
                If ctl.top = atYPosition Then
                    x = ctl.left + ctl.width
                End If
            End If
        End If
    Next ctl
    
    
    GetMaxX = x
    
End Function

Public Function GetMaxY(frm As Object, Optional objectSection = Null, Optional x = Null, Optional totalWidth = Null) As Double
    
    Dim ctl As Control, y As Double
    Dim frmControls As Object
    If Not IsNull(objectSection) Then
        Set frmControls = frm.Section(objectSection).Controls
    Else
        Set frmControls = frm.Controls
    End If
    
    For Each ctl In frmControls
    
        If Not IsNull(x) Then
            If ctl.left + totalWidth >= x And ctl.left + totalWidth <= x + totalWidth Then
                If ctl.top + ctl.height > y Then
                    y = ctl.top + ctl.height
                End If
            End If
        Else
            If ctl.top + ctl.height > y Then
                y = ctl.top + ctl.height
            End If
        End If
        
    Next ctl
    
    GetMaxY = y
    
End Function

Private Sub RenderDatasheetTotals(frm As Form, ModelID, maxWidth)
    
    Dim y, formMargin, x
    formMargin = 100
    y = GetMaxY(frm) + 700
    x = formMargin
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblDatasheetTotals WHERE ModelID = " & ModelID & " AND " & _
                             "IsHidden <> -1 ORDER BY ControlOrder ASC")
                             
    Dim ctlName, fldWidth, ctl As Control
    fldWidth = 3000 * (0.5) + (200 * (0.5 - 1))
    
    If Not rs.EOF Then
        
        Do Until rs.EOF
            
            ctlName = rs.fields("ControlName")
            
            If x + fldWidth + 100 > maxWidth Then
                x = formMargin
                y = GetMaxY(frm) + 800
            End If
            
            Set ctl = CreateControl(frm.Name, acTextBox, , "", "", x, y, fldWidth)
            ctl.Name = ctlName
            ctl.ControlSource = concat("=IfError(subform!", ctlName, ")")
            ctl.Format = "Standard"
            ctl.HorizontalAnchor = acHorizontalAnchorLeft
            ctl.VerticalAnchor = acVerticalAnchorBottom
            
            ''Set control property based on ControlTypeValue
            SetControlProperties ctl
            
            ''Generate the label just above the control
            Set ctl = CreateControl(frm.Name, acLabel, , ctl.Name, rs.fields("ControlCaption"), x, y - 500)
            
            SetControlProperties ctl
            ctl.height = 500
            ctl.width = fldWidth
            ctl.HorizontalAnchor = acHorizontalAnchorLeft
            ctl.VerticalAnchor = acVerticalAnchorBottom
           
            x = x + fldWidth + 100
            
            rs.MoveNext
        Loop
        
    End If
    
End Sub

Private Sub RenderAdditionalButtonOnMainForm(ModelID, frm As Form, y)
    
    Dim rs As Recordset, sqlStr As String, ctl As Control
    Dim maxX As Long, ModelButton, modelButtonName, FunctionName, TableWideFunction, cmdButtonName
    ''Check first if there's atleast one additional button
    sqlStr = "SELECT * FROM tblModelButtons WHERE ModelID = " & ModelID & _
             " AND HideOnMain <> -1 ORDER BY ModelButtonOrder ASC, ModelButtonID"
    Set rs = ReturnRecordset(sqlStr)

    If Not rs.EOF Then
        ''Check on what needs to be rendered, individual buttons or combo boxes.
        If isPresent("qryModelProperties", "ModelID = " & ModelID & " And Property = " & EscapeString("collapsedButtonOnMain")) Then
            ''Collapsed button so this is a combo box
            ''Create a combo box
            ''Left position should account for the label "Action:"
            ''55 is the space between controls,
            Dim lblWidth: lblWidth = 1000
            maxX = GetMaxX(frm) + 55 + lblWidth + 55
            Set ctl = CreateControl(frm.Name, acComboBox, , , , maxX, y, 3000, 400)
            ''Set the Default Control Properties Here
            SetControlProperties ctl
            ''Additional Property make the RowSource to be the SQLStr, ColumnCount to 2, ColumnWidths to 0;1
            ''Set the Height to be the same height as the buttons
            sqlStr = "SELECT ModelButtonID,ModelButton FROM tblModelButtons WHERE ModelID = " & ModelID & _
                     " AND HideOnMain <> -1 ORDER BY ModelButtonOrder ASC, ModelButtonID"
            ctl.Name = "cboFormActions"
            ctl.rowSource = sqlStr
            ctl.ColumnCount = 2
            ctl.ColumnWidths = "0;1"
            ctl.height = 400
            ctl.TopMargin = 75
            ctl.LeftMargin = 75
            ctl.FontBold = True
            ctl.HorizontalAnchor = acHorizontalAnchorRight
            
            ''Render the label here
            Set ctl = CreateControl(frm.Name, acLabel, , "cboFormActions", , maxX - 55 - lblWidth, y, lblWidth, 400)
            ''Set the Default Control Properties Here
            SetControlProperties ctl
            ctl.Name = "lblFormActions"
            ctl.Caption = "Actions: "
            ctl.TextAlign = 3
            ctl.height = 400
            ctl.TopMargin = 75
            ctl.LeftMargin = 75
            ctl.FontBold = True
            ctl.HorizontalAnchor = acHorizontalAnchorRight
            
            maxX = GetMaxX(frm) + 55
            RenderButton maxX, y, "Run", frm, "RunFormActions"
            frm("cmdRunFormActions").width = frm("cmdRunFormActions").width / 2
            frm("cmdRunFormActions").HorizontalAnchor = acHorizontalAnchorRight
            
            frm("cmdRunFormActions").OnClick = "=RunFormActions([Form],[cboFormActions])"
            
        Else
            Do Until rs.EOF
                ModelButton = rs.fields("ModelButton")
                modelButtonName = RemoveSpaces(ModelButton)
                cmdButtonName = concat("cmd", modelButtonName)
                FunctionName = rs.fields("FunctionName")
                TableWideFunction = rs.fields("TableWideFunction")
                
                maxX = GetMaxX(frm) + 55
                RenderButton maxX, y, ModelButton, frm, modelButtonName
                If Not IsNull(FunctionName) Then
                    If TableWideFunction Then
                        frm(cmdButtonName).OnClick = concat("=", FunctionName, "()")
                    Else
                        frm(cmdButtonName).OnClick = concat("=RunFunctionFromSubform([Form],""subform"",", EscapeString(FunctionName), ")")
                    End If
                End If
                
                maxX = maxX + (3200 * 0.45)
                rs.MoveNext
            Loop
        End If
        
        
    End If
    
End Sub

Private Sub RenderAdditionalButtonOnDEForm(ModelID, frm As Form, y, x, CurrentCol, buttonMultiplier, FormColumns)
    
    Dim rs As Recordset, sqlStr As String, ctl As Control
    Dim maxX As Long, ModelButton, modelButtonName, FunctionName, TableWideFunction, cmdButtonName
    ''Check first if there's atleast one additional button
    sqlStr = "SELECT * FROM tblModelButtons WHERE ModelID = " & ModelID & _
             " AND HideOnForm <> -1 ORDER BY ModelButtonOrder ASC, ModelButtonID"
    Set rs = ReturnRecordset(sqlStr)

    If Not rs.EOF Then
        ''Check on what needs to be rendered, individual buttons or combo boxes.
        If isPresent("qryModelProperties", "ModelID = " & ModelID & " And Property = " & EscapeString("collapsedButtonOnForm")) Then
            ''Collapsed button so this is a combo box
            ''Create a combo box
            ''Left position should account for the label "Action:"
            ''55 is the space between controls,
            Dim lblWidth: lblWidth = 1000
            maxX = GetMaxX(frm, y) + 55 + lblWidth + 55
            Set ctl = CreateControl(frm.Name, acComboBox, , , , maxX, y, 3000, 400)
            ''Set the Default Control Properties Here
            SetControlProperties ctl
            ''Additional Property make the RowSource to be the SQLStr, ColumnCount to 2, ColumnWidths to 0;1
            ''Set the Height to be the same height as the buttons
            sqlStr = "SELECT ModelButtonID,ModelButton FROM tblModelButtons WHERE ModelID = " & ModelID & _
                     " AND HideOnMain <> -1 ORDER BY ModelButtonOrder ASC, ModelButtonID"
            ctl.Name = "cboFormActions"
            ctl.rowSource = sqlStr
            ctl.ColumnCount = 2
            ctl.ColumnWidths = "0;1"
            ctl.height = 400
            ctl.TopMargin = 75
            ctl.LeftMargin = 75
            ctl.FontBold = True
            ctl.HorizontalAnchor = acHorizontalAnchorRight
            ctl.VerticalAnchor = acVerticalAnchorBottom
            
            ''Render the label here
            Set ctl = CreateControl(frm.Name, acLabel, , "cboFormActions", , maxX - 55 - lblWidth, y, lblWidth, 400)
            ''Set the Default Control Properties Here
            SetControlProperties ctl
            ctl.Name = "lblFormActions"
            ctl.Caption = "Actions: "
            ctl.TextAlign = 3
            ctl.height = 400
            ctl.TopMargin = 75
            ctl.LeftMargin = 75
            ctl.FontBold = True
            ctl.HorizontalAnchor = acHorizontalAnchorRight
            ctl.VerticalAnchor = acVerticalAnchorBottom
            
            maxX = GetMaxX(frm, y) + 55
            RenderButton maxX, y, "Run", frm, "RunFormActions"
            frm("cmdRunFormActions").width = frm("cmdRunFormActions").width / 2
            frm("cmdRunFormActions").HorizontalAnchor = acHorizontalAnchorRight
            frm("cmdRunFormActions").VerticalAnchor = acVerticalAnchorBottom
            frm("cmdRunFormActions").OnClick = "=RunFormActionFromDE([Form],[cboFormActions])"
            
        Else
            
            maxX = x
            
            Do Until rs.EOF
            
                ModelButton = rs.fields("ModelButton")
                modelButtonName = RemoveSpaces(ModelButton)
                cmdButtonName = concat("cmd", modelButtonName)
                FunctionName = rs.fields("FunctionName")
                TableWideFunction = rs.fields("TableWideFunction")
                
                RenderButton maxX, y, ModelButton, frm, modelButtonName
                frm(cmdButtonName).HorizontalAnchor = 0
                frm(cmdButtonName).VerticalAnchor = 1
                If Not IsNull(FunctionName) Then
                    If TableWideFunction Then
                        frm(cmdButtonName).OnClick = concat("=", FunctionName, "()")
                    Else
                        frm(cmdButtonName).OnClick = concat("=", FunctionName, "([Form])")
                    End If
                End If
                
                CurrentCol = CurrentCol + 0.5
            
                If CurrentCol = FormColumns Then
                    
                    CurrentCol = 0
                    maxX = 400
                    y = y + 600
                    
                Else
                
                    maxX = maxX + (3200 * buttonMultiplier)
                    
                End If
                
                rs.MoveNext
            Loop
            
        End If
        
        
    End If
    
End Sub

Public Function RunFormActionFromDE(frm, ModelButtonID)
    
    Dim rs As Recordset, sqlStr As String
    Dim ModelButton, modelButtonName, cmdButtonName, FunctionName, TableWideFunction
    sqlStr = "SELECT * FROM tblModelButtons WHERE ModelButtonID = " & ModelButtonID
    
    Set rs = ReturnRecordset(sqlStr)
    
    ModelButton = rs.fields("ModelButton")
    modelButtonName = RemoveSpaces(ModelButton)
    cmdButtonName = concat("cmd", modelButtonName)
    FunctionName = rs.fields("FunctionName")
    TableWideFunction = rs.fields("TableWideFunction")
    
    If TableWideFunction Then
        Run FunctionName
    Else
        Run FunctionName, frm
    End If

End Function

Public Function RunFormActions(frm, ModelButtonID, Optional subformName = "subform")
    
    Dim rs As Recordset, sqlStr As String
    Dim ModelButton, modelButtonName, cmdButtonName, FunctionName, TableWideFunction
    sqlStr = "SELECT * FROM tblModelButtons WHERE ModelButtonID = " & ModelButtonID
    
    Set rs = ReturnRecordset(sqlStr)
    
    ModelButton = rs.fields("ModelButton")
    modelButtonName = RemoveSpaces(ModelButton)
    cmdButtonName = concat("cmd", modelButtonName)
    FunctionName = rs.fields("FunctionName")
    TableWideFunction = rs.fields("TableWideFunction")
    
    If TableWideFunction Then
        Run FunctionName
    Else
        Run "RunFunctionFromSubform", frm, subformName, FunctionName
    End If

    
End Function

Public Function CreateMainForm(frm2 As Form)
    
    Dim frm As Form, rs As Recordset, rsName, frmCaption, fldName, fldWidth
    Dim CurrentCol, x, y, isMemo As Boolean, maxWidth
    Dim tblDef As DAO.TableDef, db As DAO.Database, fld As DAO.Field
    Dim ctl As Control, baseName
    
    Dim ModelID, Model, VerboseName, VerbosePlural, MainField, TableWideValidation, FormColumns, SetFocus, IsKeyVisible, QueryName, OnFormCreate, subformName, Timestamp, CreatedBy, PrimaryKey
    
    ModelID = frm2("ModelID")
    Model = frm2("Model")
    VerboseName = frm2("VerboseName")
    VerbosePlural = frm2("VerbosePlural")
    MainField = frm2("MainField")
    TableWideValidation = frm2("TableWideValidation")
    FormColumns = frm2("FormColumns")
    SetFocus = frm2("SetFocus")
    IsKeyVisible = frm2("IsKeyVisible")
    QueryName = frm2("QueryName")
    OnFormCreate = frm2("OnFormCreate")
    subformName = frm2("SubformName")
    Timestamp = frm2("Timestamp")
    CreatedBy = frm2("CreatedBy")
    PrimaryKey = frm2("PrimaryKey")
    
    If ExitIfTrue(IsNull(ModelID), "Selection is empty..") Then Exit Function
    
    CreateDSForm frm2, True
    ''Create the form
    Set frm = CreateForm
    rsName = GetTableName(Model, VerbosePlural)
    frmCaption = GetFieldCaption(VerboseName, Model)
    frm.Caption = concat(frmCaption, " List")
    
    If IsNull(PrimaryKey) Then PrimaryKey = concat(Model, "ID")
    frm.OnLoad = "=DefaultMainFormLoad([Form])"
    
    SetFormProperties 6, frm
    
    x = 100
    y = 100
    
    If Not IsNull(VerbosePlural) Then
        baseName = concat(replace(VerbosePlural, " ", ""))
    Else
        baseName = concat(Model, "s")
    End If
    
    If Not IsNull(subformName) Then
        baseName = subformName
    End If
    
    ''Add Button
    If Not isPresent("qryModelProperties", "Property = ""mainAddHidden"" And ModelID = " & ModelID) Then
        RenderButton x, y, "Add New", frm, "Add"
        frm.cmdAdd.OnClick = "=OpenFormFromMain(""frm" & baseName & """)"
        x = x + (3200 * 0.45)
    End If
    
    ''New Button
    If Not isPresent("qryModelProperties", "Property = ""mainEditHidden"" And ModelID = " & ModelID) Then
        RenderButton x, y, "View/Edit", frm, "View"
        frm.cmdView.OnClick = "=OpenFormFromMain(""frm" & baseName & """, ""subform"", """ & PrimaryKey & """,[Form])"
        x = x + (3200 * 0.45)
    End If
    
    ''Save Button
    If Not isPresent("qryModelProperties", "Property = ""mainDeleteHidden"" And ModelID = " & ModelID) Then
        RenderButton x, y, "Delete", frm, "Delete"
        frm.cmdDelete.OnClick = "=DeleteRecord([Form], """ & PrimaryKey & """, """ & rsName & """, ""subform"")"
        x = x + (3200 * 0.45)
    End If
    
    ''Additional Buttons Here (tblModelButtons)
    RenderAdditionalButtonOnMainForm ModelID, frm, y
    
    Dim maxX As Long
    
    y = y + 500: x = 100
    
    ''Get the Max x to set the width of the subform
    ''must not be less than 10000
    ''Get the x + length of the leftmost button
    maxX = GetMaxX(frm) - 100
    If maxX < 10000 Then maxX = 10000
    
    Set ctl = CreateControl(frm.Name, acSubform, , , "subform", x, y, maxX, 7000)
    ctl.Name = "subform"
    ctl.Properties("RightPadding") = 100
    ctl.SourceObject = "dsht" & baseName
    ctl.HorizontalAnchor = acHorizontalAnchorBoth
    ctl.VerticalAnchor = acVerticalAnchorBoth
    
    RenderDatasheetTotals frm, ModelID, maxX
    
    ''Set background color
    frm.Detail.BackColor = RGB(81, 163, 36)
    
    ''Align the collapsed controls if there's any
    If DoesPropertyExists(frm.Controls, "cboFormActions") Then
        maxX = GetMaxX(frm)
        frm("cmdRunFormActions").left = maxX - frm("cmdRunFormActions").width
        frm("cboFormActions").left = frm("cmdRunFormActions").left - 55 - frm("cboFormActions").width
        frm("lblFormActions").left = frm("cboFormActions").left - 55 - frm("lblFormActions").width
    End If
    
    ''Render the filterForm here
    RenderFilterForm frm, ModelID
    
    frm.Section("Detail").height = GetMaxY(frm) + 200
    frm.width = GetMaxX(frm) + 100
    
    Dim frmName As String, customFrmName As String, baseFormName As String, i As Integer
    frmName = frm.Name
    If Not IsNull(VerbosePlural) Then
        baseFormName = concat("main", replace(VerbosePlural, " ", ""))
    Else
        baseFormName = concat("main", Model, "s")
    End If
    
    If Not IsNull(subformName) Then
        baseFormName = concat("main", subformName)
    End If
    
    ''Special Function to run on form creation
    If Not IsNull(OnFormCreate) Then
        Run OnFormCreate, frm, 6
    End If
    
    ''Override Properties Here
    OverrideProperties ModelID, 6, frm
    
    DoCmd.Close acForm, frm.Name, acSaveYes
    
    customFrmName = baseFormName
    
'    Do Until Not FrmExist(customFrmName)
'
'        If MsgBox(customFrmName & " already exists. Would you like to replace it?", vbYesNo) = vbYes Then
'            Exit Do
'        End If
'        i = i + 1
'        customFrmName = baseFormName & "_" & i
'
'    Loop
    
    ''Insert the newly created form to the InsertToModelRelatedObjects
    InsertToModelRelatedObjects ModelID, acForm, customFrmName
    
    DoCmd.Rename customFrmName, acForm, frmName
    
    ''Load Save Form Layout
    LoadSavedFormLayout customFrmName
    
    DoCmd.OpenForm customFrmName
    
    InsertFormInFormForRights customFrmName, Model

End Function

Private Sub NewRow(frm As Form, x, ByRef y, originalY, totalWidth)

    y = GetMaxY(frm, , x, totalWidth) + 100
    If y = 100 Then
        y = originalY
    End If
    
End Sub

Private Sub SetComboBoxSQLForFilter(ctl As Control, ModelFieldID)
    
    Dim parentModelRs As Recordset, modelFieldRs As Recordset
    Set modelFieldRs = ReturnRecordset("SELECT * FROM tblModelFields WHERE ModelFieldID = " & ModelFieldID)
    Set parentModelRs = ReturnRecordset("SELECT * FROM tblModels WHERE ModelID = " & modelFieldRs.fields("ParentModelID"))
    
    Dim MainField
    MainField = parentModelRs.fields("MainField")
    If MainField Like "=*" Then MainField = replace(MainField, "=", "")
    
    Dim primaryTable
    primaryTable = GetTableName(parentModelRs.fields("Model"), parentModelRs.fields("VerbosePlural"))
    
    Dim primaryField
    primaryField = concat(parentModelRs.fields("Model"), "ID")
        
    Dim rowSourceSQL
    rowSourceSQL = concat("SELECT ", primaryField, ",", MainField, " As MainField FROM ", primaryTable, " ORDER BY ", MainField)
    ctl.rowSource = rowSourceSQL
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "0;1"
    
End Sub

Private Function GenerateWildSearchSQL(rs As Recordset, ctlValue) As String

    Dim modelFieldRs As Recordset, fltrArr As New clsArray
    rs.MoveFirst
    Do Until rs.EOF
        Set modelFieldRs = ReturnRecordset("SELECT * FROM tblModelFields WHERE ModelFieldID = " & rs.fields("ModelFieldID"))
        fltrArr.Add modelFieldRs.fields("ModelField") & " Like " & EscapeString("*" & ctlValue & "*")
        rs.MoveNext
    Loop
    
    GenerateWildSearchSQL = "(" & fltrArr.JoinArr(" OR ") & ")"
    
End Function

Private Function GenerateNumericSearch(FieldName, fromValue, toValue) As String
    
    If IsNull(fromValue) And IsNull(toValue) Then Exit Function
    
    If Not IsNull(fromValue) And Not IsNull(toValue) Then
        GenerateNumericSearch = FieldName & " Between " & fromValue & " And " & toValue
    ElseIf Not IsNull(toValue) Then
        GenerateNumericSearch = FieldName & " <= " & toValue
    ElseIf Not IsNull(fromValue) Then
        GenerateNumericSearch = FieldName & " >= " & fromValue
    End If
    
End Function

Private Function GenerateDateSearch(FieldName, fromValue, toValue) As String
    
    If IsNull(fromValue) And IsNull(toValue) Then Exit Function
    
    If Not IsNull(fromValue) And Not IsNull(toValue) Then
        GenerateDateSearch = FieldName & " Between #" & fromValue & "# And #" & toValue & "#"
    ElseIf Not IsNull(toValue) Then
        GenerateDateSearch = FieldName & " <= #" & toValue & "#"
    ElseIf Not IsNull(fromValue) Then
        GenerateDateSearch = FieldName & " >= #" & fromValue & "#"
    End If
    
End Function


Private Function GenerateMonthYearSearch(FieldName, monthValue, yearValue) As String
    
    Dim fltrArr As New clsArray
    If IsNull(monthValue) And IsNull(yearValue) Then Exit Function
    
    If Not IsNull(monthValue) Then
        fltrArr.Add "Month(" & FieldName & ") = " & monthValue
    End If
    
    If Not IsNull(yearValue) Then
        fltrArr.Add "Year(" & FieldName & ") = " & yearValue
    End If
    
    If fltrArr.Count > 0 Then
        GenerateMonthYearSearch = fltrArr.JoinArr(" AND ")
    End If
    
End Function

Public Function ClearFilterSubform(frm, ModelID)
    
    Dim rsName
    rsName = GetTableNameFromModelID(ModelID)
    
    Dim ctl As Control
    For Each ctl In frm.Controls
        If ctl.Name Like "fltr*" Then
            If ctl.ControlType <> acCheckBox Then
                ctl = Null
            End If
        End If
        
        
        If ctl.ControlType = acOptionGroup Then
            Debug.Print ""
            ctl = 2
        End If
        
    Next ctl
    
    Dim filterStr, OrderBy, sqlStr
    sqlStr = "SELECT * FROM " & rsName
    
    If Not frm.subform.SourceObject Like "Report.*" Then
'        orderBy = frm.subform.Form.orderBy
'        If Not IsNull(orderBy) And orderBy <> "" Then sqlStr = sqlStr & " ORDER BY " & orderBy
'        frm.subform.Form.RecordSource = sqlStr
'        frm.subform.Form.orderBy = orderBy
'        frm.subform.Requery
        frm.subform.Form.FilterOn = False
    Else
'        orderBy = frm.subform.Form.orderBy
'        If Not IsNull(orderBy) And orderBy <> "" Then sqlStr = sqlStr & " ORDER BY " & orderBy
'        frm.subform.Report.RecordSource = sqlStr
'        frm.subform.Form.orderBy = orderBy
'        frm.subform.Requery
        frm.subform.Report.FilterOn = False
    End If
 
End Function

Public Function FilterSubform(frm, ModelID)
    
    ''Fetch the FilterFields
    Dim rs As Recordset, modelRs As Recordset, modelFieldRs As Recordset, ctl As Control
    Set rs = ReturnRecordset("SELECT * FROM tblFilterFields WHERE ModelID = " & ModelID)

    ''Check first if the recordset is empty
    If rs.EOF Then Exit Function
    ''Open the wildcards filterFields
    Set rs = ReturnRecordset("SELECT * FROM tblFilterFields WHERE ModelID = " & ModelID & " ANd IsWildSearch = -1")
    
    Dim ctlName, fltrArr As New clsArray, FieldName
    
    If Not rs.EOF Then
        ctlName = "fltrWildSearch"
        ''Check if the control is not null
        If Not IsNull(frm(ctlName)) Then
            fltrArr.Add GenerateWildSearchSQL(rs, frm(ctlName))
        End If
    End If
    
    ''User filter controls which is not a wildsearch
    Set rs = ReturnRecordset("SELECT * FROM tblFilterFields WHERE ModelID = " & ModelID & " ANd IsWildSearch = 0 Order By FilterOrder ASC")
    Do Until rs.EOF
    
        Set modelFieldRs = ReturnRecordset("SELECT * FROM tblModelFields WHERE ModelFieldID = " & rs.fields("ModelFieldID"))

        ''Boolean Filter
        If modelFieldRs.fields("FieldTypeID") = dbBoolean Then
            
            FieldName = modelFieldRs.fields("ModelField")
            ctlName = "ogFltr" & FieldName
            
            If Not IsNull(frm(ctlName)) Then
                Select Case frm(ctlName)
                    Case 2:
                        
                    Case Else:
                        fltrArr.Add FieldName & " = " & frm(ctlName)
                        
                End Select
            End If
            GoTo NextFilter:
        End If
        
        If Not IsNull(rs.fields("FilterOperator")) Then
            
            FieldName = modelFieldRs.fields("ModelField")
            ctlName = "fltr" & FieldName
               
            If Not IsNull(frm(ctlName)) Then fltrArr.Add FieldName & " Like " & EscapeString("*" & frm(ctlName) & "*")
            
            GoTo NextFilter:
            
        End If

        If rs.fields("IsList") Then
            
            FieldName = modelFieldRs.fields("ModelField")
            ctlName = "fltr" & FieldName
            
            If IsNull(modelFieldRs.fields("PossibleValues")) Then
                If Not IsNull(frm(ctlName)) Then fltrArr.Add FieldName & " = " & frm(ctlName)
            Else
                If Not IsNull(frm(ctlName)) Then fltrArr.Add FieldName & " = " & EscapeString(frm(ctlName))
            End If
            
            GoTo NextFilter:
        End If

        ''Double Filter
        Dim resultingSQL
        If modelFieldRs.fields("FieldTypeID") = dbDouble Then
            
            FieldName = modelFieldRs.fields("ModelField")
            ctlName = "fltr" & FieldName
            
            resultingSQL = GenerateNumericSearch(FieldName, frm(ctlName & "From"), frm(ctlName & "To"))

            If resultingSQL <> "" Then
                fltrArr.Add resultingSQL
            End If
            GoTo NextFilter:
        End If

        If rs.fields("IsMonthYear") Then
        
            FieldName = modelFieldRs.fields("ModelField")
            ctlName = "fltr" & FieldName
            
            resultingSQL = GenerateMonthYearSearch(FieldName, frm(ctlName & "Month"), frm(ctlName & "Year"))
            
            If resultingSQL <> "" Then
                fltrArr.Add resultingSQL
            End If
            GoTo NextFilter:
        End If

        If rs.fields("IsBetween") Then
        
            FieldName = modelFieldRs.fields("ModelField")
            ctlName = "fltr" & FieldName
            
            resultingSQL = GenerateDateSearch(FieldName, frm(ctlName & "From"), frm(ctlName & "To"))

            If resultingSQL <> "" Then
                fltrArr.Add resultingSQL
            End If
            GoTo NextFilter:

        End If
NextFilter:
        rs.MoveNext
    Loop
    
    Dim rsName
    rsName = GetTableNameFromModelID(ModelID)
    
    Dim filterStr, OrderBy, sqlStr
    
'    If fltrArr.Count > 0 Then
'        filterStr = fltrArr.JoinArr(" AND ")
'        sqlStr = "SELECT * FROM " & rsName & " WHERE " & filterStr
'        If Not frm.subform.SourceObject Like "Report.*" Then
'            orderBy = frm.subform.Form.orderBy
'            If Not IsNull(orderBy) And orderBy <> "" Then sqlStr = sqlStr & " ORDER BY " & orderBy
'            frm.subform.Form.RecordSource = sqlStr
'            frm.subform.Form.orderBy = orderBy
'            frm.subform.Requery
'        Else
'            orderBy = frm.subform.Report.orderBy
'            If Not IsNull(orderBy) And orderBy <> "" Then sqlStr = sqlStr & " ORDER BY " & orderBy
'            frm.subform.Report.RecordSource = sqlStr
'            frm.subform.Report.orderBy = orderBy
'            frm.subform.Requery
'        End If
'    End If

    If fltrArr.Count > 0 Then
        filterStr = fltrArr.JoinArr(" AND ")
        If Not frm.subform.SourceObject Like "Report.*" Then
            frm.subform.Form.Filter = filterStr
            frm.subform.Form.FilterOn = True
'            orderBy = frm.subform.Form.orderBy
'            If Not IsNull(orderBy) And orderBy <> "" Then sqlStr = sqlStr & " ORDER BY " & orderBy
'            frm.subform.Form.RecordSource = sqlStr
'            frm.subform.Form.orderBy = orderBy
'            frm.subform.Requery
        Else
            frm.subform.Report.Filter = filterStr
            frm.subform.Report.FilterOn = True
'            orderBy = frm.subform.Report.orderBy
'            If Not IsNull(orderBy) And orderBy <> "" Then sqlStr = sqlStr & " ORDER BY " & orderBy
'            frm.subform.Report.RecordSource = sqlStr
'            frm.subform.Report.orderBy = orderBy
'            frm.subform.Requery
        End If
    End If
    
End Function


Private Sub FilterControlSetCommonProperties(ctl As Control)

    SetControlProperties ctl
    ctl.HorizontalAnchor = acHorizontalAnchorRight
    If ctl.ControlType = acLabel Then
        Select Case ctl.Caption
            Case "TO", "Month", "Year":
            Case "ALL", "YES", "NO":
                ctl.TextAlign = 2
            Case Else:
                ctl.BackColor = RGB(255, 226, 73)
                ctl.ForeColor = RGB(81, 163, 36)
                ctl.BackStyle = 1
        End Select
        
    End If
    
End Sub

Public Function CopyFromToToDate(frm As Form, ctlName)
    
    Dim ctlTo As Control, ctlFrom As Control
    Set ctlTo = frm(ctlName & "To")
    Set ctlFrom = frm(ctlName & "From")
    
'    If IsNull(ctlTo) Then
'        ctlTo = ctlFrom
'    Else
'        If ctlTo < ctlFrom Then
'            ctlTo = ctlFrom
'        End If
'    End If
    
    If ctlTo < ctlFrom Then
        ctlTo = ctlFrom
    End If
    
End Function

Public Function CopyFromToIfEarlier(frm As Form, ctlName)

    Dim ctlTo As Control, ctlFrom As Control
    Set ctlTo = frm(ctlName & "To")
    Set ctlFrom = frm(ctlName & "From")
    
'    If IsNull(ctlFrom) Then
'        ctlFrom = ctlTo
'    Else
'        If ctlFrom > ctlTo Then
'            ctlFrom = ctlTo
'        End If
'    End If
    
    If ctlFrom > ctlTo Then
        ctlFrom = ctlTo
    End If
    
End Function

Private Sub RenderFilterForm(frm As Form, ModelID)

    ''Filter Form Property
    Dim totalWidth, colSpaceWidth, x, y, originalY, ctlHeight
    totalWidth = 3500 ''2 inches => 1 inc => 1440 twips
    colSpaceWidth = 50
    ctlHeight = 300 ''300 twips
    y = frm("subform").top
    originalY = y
    x = frm("subform").left + frm("subform").width + 100 ''200 twips is the form margin
    
    ''Fetch the FilterFields
    Dim rs As Recordset, modelRs As Recordset, modelFieldRs As Recordset, ctl As Control
    Set rs = ReturnRecordset("SELECT * FROM tblFilterFields WHERE ModelID = " & ModelID)
    Set modelRs = ReturnRecordset("SELECT * FROM tblModels WHERE ModelID = " & ModelID)
    
    
    ''Check first if the recordset is empty
    If rs.EOF Then Exit Sub
    ''Open the wildcards filterFields
    Set rs = ReturnRecordset("SELECT * FROM tblFilterFields WHERE ModelID = " & ModelID & " ANd IsWildSearch = -1")
    If Not rs.EOF Then
        ''Render the wildcard seach box here the text box first
        Set ctl = CreateControl(frm.Name, acTextBox, , , , x, y + ctlHeight + 100, totalWidth)
        FilterControlSetCommonProperties ctl
        ctl.Name = "fltrWildSearch"
        ''Then the label
        Set ctl = CreateControl(frm.Name, acLabel, , "fltrWildSearch", , x, y, totalWidth)
        FilterControlSetCommonProperties ctl
        ctl.Caption = "Search " & GetModelPlural(modelRs.fields("Model"), modelRs.fields("VerbosePlural"), "")
        
    End If
    
    ''Render other filter controls Not isWildSearch
    Dim ctlName
    Dim proportionArr As New clsArray, proportionTotal, controlArr As New clsArray, i As Integer, proportion As Double, startX
    
    Set rs = ReturnRecordset("SELECT * FROM tblFilterFields WHERE ModelID = " & ModelID & " ANd IsWildSearch = 0 Order By FilterOrder ASC")
    Do Until rs.EOF
    
        NewRow frm, x, y, originalY, totalWidth
        
        Set modelFieldRs = ReturnRecordset("SELECT * FROM tblModelFields WHERE ModelFieldID = " & rs.fields("ModelFieldID"))
        
        ''Boolean Filter
        If modelFieldRs.fields("FieldTypeID") = dbBoolean Then
            
            startX = x
            ctlName = "fltr" & modelFieldRs.fields("ModelField")
  
            proportionArr.arr = "6,2,6,2,6,2"
            controlArr.arr = "lbl" & ctlName & "ALL" & "," & _
                             ctlName & "ALL," & _
                             "lbl" & ctlName & "YES" & "," & _
                             ctlName & "YES," & _
                             "lbl" & ctlName & "NO" & "," & _
                             ctlName & "NO"
                             
            proportionTotal = GetProportionTotal(proportionArr)
            
            ''Render the optionGroup control
            Set ctl = CreateControl(frm.Name, acOptionGroup, , , , 0, 0, totalWidth)
            ctl.Name = "og" & ctlName
            ctl.top = y
            ctl.left = x
            ctl.DefaultValue = 2
            ctl.BorderStyle = 0
            ctl.SpecialEffect = 0
            ctl.HorizontalAnchor = acHorizontalAnchorRight
            ''Render the ALL
            Set ctl = CreateControl(frm.Name, acCheckBox, , "og" & ctlName, , 0, 0, totalWidth)
            ctl.Name = ctlName & "ALL"
            ctl.optionValue = 2
            FilterControlSetCommonProperties ctl
            ''Render the ALL Label
            Set ctl = CreateControl(frm.Name, acLabel, , ctlName & "ALL", , x, y, totalWidth)
            ctl.Name = "lbl" & ctlName & "ALL"
            ctl.Caption = "ALL"
            FilterControlSetCommonProperties ctl
            
            ''Render the YES
            Set ctl = CreateControl(frm.Name, acCheckBox, , "og" & ctlName, , 0, 0, totalWidth)
            ctl.Name = ctlName & "YES"
            ctl.optionValue = -1
            FilterControlSetCommonProperties ctl
            ''Render the YES Label
            Set ctl = CreateControl(frm.Name, acLabel, , ctlName & "YES", , x, y, totalWidth)
            ctl.Name = "lbl" & ctlName & "YES"
            ctl.Caption = "YES"
            FilterControlSetCommonProperties ctl
            
            ''Render the NO
            Set ctl = CreateControl(frm.Name, acCheckBox, , "og" & ctlName, , 0, 0, totalWidth)
            ctl.Name = ctlName & "NO"
            ctl.optionValue = 0
            FilterControlSetCommonProperties ctl
            ''Render the NO Label
            Set ctl = CreateControl(frm.Name, acLabel, , ctlName & "NO", , x, y, totalWidth)
            ctl.Name = "lbl" & ctlName & "NO"
            ctl.Caption = "NO"
            FilterControlSetCommonProperties ctl
            
            For i = 0 To proportionArr.Count - 1
    
                proportion = CDbl(proportionArr.arr(i)) / proportionTotal
                frm(controlArr.arr(i)).left = startX
                frm(controlArr.arr(i)).top = y + ctlHeight + 100
                frm(controlArr.arr(i)).width = (totalWidth - ((proportionArr.Count - 1) * colSpaceWidth * 2)) * proportion
                frm(controlArr.arr(i)).HorizontalAnchor = acHorizontalAnchorRight
               
                startX = startX + (colSpaceWidth * 2) + frm(controlArr.arr(i)).width
                
            Next i
            
            frm("og" & ctlName).top = y
            frm("og" & ctlName).left = x
            frm("og" & ctlName).width = 0
            frm("og" & ctlName).height = 0
            ''Then the label
            Set ctl = CreateControl(frm.Name, acLabel, , , , x, y, totalWidth)
            FilterControlSetCommonProperties ctl
            ctl.Caption = GetFieldCaption(modelFieldRs.fields("VerboseName"), modelFieldRs.fields("ModelField"))
            
        End If
        
        If rs.fields("IsList") Then
            ''Render the combo box here first
            Set ctl = CreateControl(frm.Name, acComboBox, , , , x, y + ctlHeight + 100, totalWidth)
            FilterControlSetCommonProperties ctl
            ctlName = "fltr" & modelFieldRs.fields("ModelField")
            ctl.Name = ctlName
            
            If IsNull(modelFieldRs.fields("PossibleValues")) Then
                SetComboBoxSQLForFilter ctl, rs.fields("ModelFieldID")
                
            Else
                ctl.rowSource = Join(Split(modelFieldRs.fields("PossibleValues"), ","), ";")
                ctl.ColumnCount = 1
                ctl.ColumnWidths = "1"
                ctl.RowSourceType = "Value List"
                ctl.LimitToList = -1
                ctl.AllowValueListEdits = 0
            End If
            
            ''Then the label
            Set ctl = CreateControl(frm.Name, acLabel, , ctlName, , x, y, totalWidth)
            FilterControlSetCommonProperties ctl
            ctl.Caption = GetFieldCaption(modelFieldRs.fields("VerboseName"), modelFieldRs.fields("ModelField"))
            
            GoTo NextFilter:
        End If
        
        ''Double Filter
        If modelFieldRs.fields("FieldTypeID") = dbDouble Then
            
            startX = x
            ctlName = "fltr" & modelFieldRs.fields("ModelField")
            proportionArr.arr = "10,2,10"
            controlArr.arr = ctlName & "From,lbl" & ctlName & "To," & ctlName & "To"
            proportionTotal = GetProportionTotal(proportionArr)
            
            ''Render the From
            Set ctl = CreateControl(frm.Name, acTextBox, , , , 0, 0, totalWidth)
            ctl.Name = ctlName & "From"
            FilterControlSetCommonProperties ctl
            ctl.Format = "Standard"

            ''Render the To
            Set ctl = CreateControl(frm.Name, acTextBox, , , , 0, 0, totalWidth)
            ctl.Name = ctlName & "To"
            FilterControlSetCommonProperties ctl
            ctl.Format = "Standard"
            ''Render the label
            Set ctl = CreateControl(frm.Name, acLabel, , , , 0, 0, totalWidth)
            ctl.Name = "lbl" & ctlName & "To"
            ctl.Caption = "TO"
            ctl.TextAlign = 2
            FilterControlSetCommonProperties ctl
            
            For i = 0 To proportionArr.Count - 1
    
                proportion = CDbl(proportionArr.arr(i)) / proportionTotal
                frm(controlArr.arr(i)).left = startX
                frm(controlArr.arr(i)).top = y + ctlHeight + 100
                frm(controlArr.arr(i)).width = (totalWidth - ((proportionArr.Count - 1) * colSpaceWidth * 2)) * proportion
                frm(controlArr.arr(i)).HorizontalAnchor = acHorizontalAnchorRight
               
                startX = startX + (colSpaceWidth * 2) + frm(controlArr.arr(i)).width
               
            Next i
            
            ''Then the label
            Set ctl = CreateControl(frm.Name, acLabel, , ctlName & "From", , x, y, totalWidth)
            FilterControlSetCommonProperties ctl
            ctl.Caption = GetFieldCaption(modelFieldRs.fields("VerboseName"), modelFieldRs.fields("ModelField"))
            
        End If
        
        If rs.fields("IsMonthYear") Then
            
            startX = x
            ctlName = "fltr" & modelFieldRs.fields("ModelField")
            proportionArr.arr = "6,10,4,8"
            controlArr.arr = "lbl" & ctlName & "Month" & "," & _
                             ctlName & "Month," & _
                             "lbl" & ctlName & "Year," & _
                             ctlName & "Year"
            proportionTotal = GetProportionTotal(proportionArr)
            
            ''Render the Month
            Set ctl = CreateControl(frm.Name, acComboBox, , , , 0, 0, totalWidth)
            ctl.Name = ctlName & "Month"
            FilterControlSetCommonProperties ctl
            ctl.ColumnCount = 2
            ctl.ColumnWidths = "0;1"
            ctl.rowSource = "SELECT MonthID, MonthName FROM tblMonths ORDER BY MonthID"
            ''Render the month label
            Set ctl = CreateControl(frm.Name, acLabel, , ctlName & "Month", , x, y, totalWidth)
            ctl.Name = "lbl" & ctlName & "Month"
            ctl.Caption = "Month"
            FilterControlSetCommonProperties ctl
            ''Render the year
            Set ctl = CreateControl(frm.Name, acTextBox, , , , 0, 0, totalWidth)
            ctl.Name = ctlName & "Year"
            FilterControlSetCommonProperties ctl
             ''Render the year label
            Set ctl = CreateControl(frm.Name, acLabel, , ctlName & "Year", , x, y, totalWidth)
            ctl.Name = "lbl" & ctlName & "Year"
            ctl.Caption = "Year"
            FilterControlSetCommonProperties ctl
            
            For i = 0 To proportionArr.Count - 1
    
                proportion = CDbl(proportionArr.arr(i)) / proportionTotal
                frm(controlArr.arr(i)).left = startX
                frm(controlArr.arr(i)).top = y + ctlHeight + 100
                frm(controlArr.arr(i)).width = (totalWidth - ((proportionArr.Count - 1) * colSpaceWidth * 2)) * proportion
                frm(controlArr.arr(i)).HorizontalAnchor = acHorizontalAnchorRight
               
                startX = startX + (colSpaceWidth * 2) + frm(controlArr.arr(i)).width
                
            Next i
            
            ''Then the label
            Set ctl = CreateControl(frm.Name, acLabel, , , , x, y, totalWidth)
            FilterControlSetCommonProperties ctl
            ctl.Caption = GetFieldCaption(modelFieldRs.fields("VerboseName"), modelFieldRs.fields("ModelField"))
            
        End If
        
        If rs.fields("IsBetween") Then
            
            startX = x
            ctlName = "fltr" & modelFieldRs.fields("ModelField")
            proportionArr.arr = "10,2,10"
            controlArr.arr = ctlName & "From,lbl" & ctlName & "To," & ctlName & "To"
            proportionTotal = GetProportionTotal(proportionArr)
            
            ''Render the From
            Set ctl = CreateControl(frm.Name, acTextBox, , , , 0, 0, totalWidth)
            ctl.Name = ctlName & "From"
            FilterControlSetCommonProperties ctl
            ctl.Format = "Short Date"
            ctl.AfterUpdate = "=CopyFromToToDate([Form], " & EscapeString(ctlName) & ")"
            ''Render the To
            Set ctl = CreateControl(frm.Name, acTextBox, , , , 0, 0, totalWidth)
            ctl.Name = ctlName & "To"
            FilterControlSetCommonProperties ctl
            ctl.Format = "Short Date"
            ctl.AfterUpdate = "=CopyFromToIfEarlier([Form], " & EscapeString(ctlName) & ")"
            ''Render the label
            Set ctl = CreateControl(frm.Name, acLabel, , , , 0, 0, totalWidth)
            ctl.Name = "lbl" & ctlName & "To"
            ctl.Caption = "TO"
            ctl.TextAlign = 2
            FilterControlSetCommonProperties ctl
            
            For i = 0 To proportionArr.Count - 1
    
                proportion = CDbl(proportionArr.arr(i)) / proportionTotal
                frm(controlArr.arr(i)).left = startX
                frm(controlArr.arr(i)).top = y + ctlHeight + 100
                frm(controlArr.arr(i)).width = (totalWidth - ((proportionArr.Count - 1) * colSpaceWidth * 2)) * proportion
                frm(controlArr.arr(i)).HorizontalAnchor = acHorizontalAnchorRight
               
                startX = startX + (colSpaceWidth * 2) + frm(controlArr.arr(i)).width
               
            Next i
            
            ''Then the label
            Set ctl = CreateControl(frm.Name, acLabel, , ctlName & "From", , x, y, totalWidth)
            FilterControlSetCommonProperties ctl
            ctl.Caption = GetFieldCaption(modelFieldRs.fields("VerboseName"), modelFieldRs.fields("ModelField"))
            
        End If
NextFilter:
        rs.MoveNext
    Loop
    
    ''Render the Filter buttons
    ''Filter and Clear
    proportionArr.arr = "6,6"
    controlArr.arr = "cmdFilter,cmdClear"
    proportionTotal = GetProportionTotal(proportionArr)
    NewRow frm, x, y, originalY, totalWidth
    
    RenderButton 0, 0, "Filter", frm, "Filter"
    frm("cmdFilter").OnClick = "=FilterSubform([Form]," & ModelID & ")"
    RenderButton 0, 0, "Clear Filter", frm, "Clear"
    frm("cmdClear").OnClick = "=ClearFilterSubform([Form]," & ModelID & ")"
    
    For i = 0 To proportionArr.Count - 1
    
        proportion = CDbl(proportionArr.arr(i)) / proportionTotal
        frm(controlArr.arr(i)).left = x
        frm(controlArr.arr(i)).top = y
        frm(controlArr.arr(i)).width = (totalWidth - ((proportionArr.Count - 1) * colSpaceWidth * 2)) * proportion
        frm(controlArr.arr(i)).HorizontalAnchor = acHorizontalAnchorRight
        frm(controlArr.arr(i)).height = frm(controlArr.arr(i)).height * 0.8
       
        x = x + (colSpaceWidth * 2) + frm(controlArr.arr(i)).width
       
    Next i
    
    
End Sub

Public Function GetProportionTotal(proportionArr As clsArray) As Double

    Dim proportion
    For Each proportion In proportionArr.arr
        GetProportionTotal = GetProportionTotal + CDbl(proportion)
    Next proportion
    
End Function


Public Function CreateDSForm(frm2 As Form, Optional DontOpen As Boolean = False)

    Dim frm As Form, rs As Recordset, rsName, frmCaption, fldName, fldWidth
    Dim CurrentCol, x, y, isMemo As Boolean, maxWidth
    Dim ctl As Control
    
    Dim ModelID, Model, VerboseName, VerbosePlural, MainField, TableWideValidation, FormColumns, SetFocus, IsKeyVisible
    Dim QueryName, OnFormCreate, subformName, UserQueryFields, Timestamp, CreatedBy, PrimaryKey
    
    ModelID = frm2("ModelID")
    Model = frm2("Model")
    VerboseName = frm2("VerboseName")
    VerbosePlural = frm2("VerbosePlural")
    MainField = frm2("MainField")
    TableWideValidation = frm2("TableWideValidation")
    FormColumns = frm2("FormColumns")
    SetFocus = frm2("SetFocus")
    IsKeyVisible = frm2("IsKeyVisible")
    QueryName = frm2("QueryName")
    OnFormCreate = frm2("OnFormCreate")
    subformName = frm2("SubformName")
    UserQueryFields = frm2("UserQueryFields")
    Timestamp = frm2("Timestamp")
    CreatedBy = frm2("CreatedBy")
    PrimaryKey = frm2("PrimaryKey")
    
    If ExitIfTrue(IsNull(ModelID), "Selection is empty..") Then Exit Function
    
    GenerateFields frm2
    '''Create the controls
    '''Fields, default buttons, additional buttons and controls
    
    ''Create the form
    Set frm = CreateForm
    rsName = GetTableName(Model, VerbosePlural)
    
    If Not IsNull(QueryName) Then
        rsName = QueryName
    End If
    
    frm.RecordSource = rsName
    frmCaption = GetFieldCaption(VerboseName, Model)
    frm.Caption = concat(frmCaption, " Datasheet")
    DoCmd.RunCommand acCmdFormHdrFtr
    frm.Section(acHeader).height = 0
    frm.Section(acFooter).height = 0
    
    'frm.OnCurrent = "=SetFocusOnForm([Form],""" & SetFocus & """)"
    
    If IsNull(PrimaryKey) Then PrimaryKey = concat(Model, "ID")
    frm.BeforeUpdate = "=SaveFormData2([Form],""" & Model & """)"
    frm.OnLoad = "=SetDefaultUserID([Form])"
    
    SetFormProperties 5, frm
    CurrentCol = 1
    
    Dim rsObj As Object, db As DAO.Database, fld As DAO.Field
    Set db = CurrentDb
    If DoesPropertyExists(db.TableDefs, rsName) Then
        Set rsObj = db.TableDefs(rsName)
    Else
        Set rsObj = db.QueryDefs(rsName)
    End If
    
    x = 400
    y = 600
    
    CurrentCol = 0: isMemo = False
    
    Dim sqlStr As String
    sqlStr = "SELECT * FROM tblModelFields WHERE FieldOrder <> 0 AND ModelID = " & ModelID
    
    If Not UserQueryFields Then
        sqlStr = sqlStr & " AND FieldSource = " & EscapeString(rsName)
    End If
    
    'sqlStr = sqlStr & ")"
    
    Set rs = ReturnRecordset(sqlStr & " ORDER BY FieldOrder ASC")
    
    Do Until rs.EOF
        
        fldName = GetFieldName(rs.fields("ForeignKey"), rs.fields("ModelField"), Not IsNull(rs.fields("ParentModelID")))
        fldWidth = 3000
        Set fld = rsObj.fields(fldName)
        
        If (Not IsKeyVisible And fld.Name = PrimaryKey) Or Not IsNull(rs.fields("ControlSource")) Then
            GoTo NextField
        End If
        
        ''Skip the field if there is an imageType proprety on qryModelFieldProperties
        If isPresent("qryModelFieldProperties", "Property = " & EscapeString("imageType") & " And ModelFieldID = " & rs.fields("ModelFieldID")) Then
            GoTo NextField
        End If
        
        Select Case fld.Name
            Case "Timestamp", "CreatedBy":
                GoTo NextField
        End Select
        
        Dim ControlTypeValue
        If Not DoesPropertyExists(fld.Properties, "DisplayControl") Then
            ControlTypeValue = acTextBox
        Else
            ControlTypeValue = fld.Properties("DisplayControl")
        End If
        
        ''Generate the control first before the label
        ''Get the width depending on the number of columns but make sure that the CurrentCol + the columns will not exceed
        ''the FormColumn
        fldWidth = 3000 * (rs.fields("Columns")) + (200 * (rs.fields("Columns") - 1))
        
        Set ctl = CreateControl(frm.Name, ControlTypeValue, , "", fld.Name, x, y, fldWidth)
        ctl.Name = fld.Name
        Dim ColumnWidth: ColumnWidth = rs.fields("ColumnWidth")
        If Not isFalse(ColumnWidth) Then
            ctl.Tag = ctl.Tag & " DontAutoWidth"
            ctl.ColumnWidth = ColumnWidth + 3600
        End If
        
        ''Set control property based on ControlTypeValue
        SetControlProperties ctl
        
        If Not IsNull(rs.fields("ColumnWidth")) Then
            ctl.ColumnWidth = 2000 + rs.fields("ColumnWidth")
        End If
        
        If isPresent("qryModelFieldProperties", "Property = " & EscapeString("alwaysHideOnDatasheet") & " And ModelFieldID = " & rs.fields("ModelFieldID")) Then
            ctl.ColumnHidden = True
            ctl.Tag = ctl.Tag & "alwaysHideOnDatasheet"
        End If
        
        Select Case fld.Type
            Case dbMemo:
                ctl.height = 900
                isMemo = True
            Case dbDouble:
                ctl.Format = "Standard"
            
        End Select
        
        Select Case fld.Type
        
            Case dbDouble, dbInteger:
                ''Create a control at the footer of the form
                Dim footerCtl As Control, footerControlCaption
                Set footerCtl = CreateControl(frm.Name, acTextBox, acFooter, "", , 400, 600, 3000)
                SetControlProperties footerCtl
                footerCtl.Name = concat("Sum", rs.fields("ModelField"))
                footerCtl.ControlSource = "=CdblNz(Sum([" & fld.Name & "]))"
                
                If IsNull(rs.fields("VerboseName")) Then
                    footerControlCaption = AddSpaces(rs.fields("ModelField"))
                Else
                    footerControlCaption = rs.fields("VerboseName")
                End If
                
                footerCtl.Properties("DatasheetCaption") = footerControlCaption
                
                
        End Select
        
        ''Also set the DataSheetCaption
        If Not IsNull(rs.fields("VerboseName")) Then
            ctl.Properties("DatasheetCaption") = rs.fields("VerboseName")
        Else
            If Not DoesPropertyExists(fld.Properties, "Caption") Then
                ctl.Properties("DatasheetCaption") = AddSpaces(fld.Name)
            Else
                ctl.Properties("DatasheetCaption") = fld.Properties("Caption")
            End If
        End If
    
'        ''Generate the label just above the control
'        Set ctl = CreateControl(frm.Name, acLabel, , fld.Name, fld.Properties("Caption"), x, y - 300)
'        SetControlProperties ctl
'        ctl.Width = fldWidth
        
        CurrentCol = CurrentCol + rs.fields("Columns")
        
        If CurrentCol >= FormColumns Then
            
            If x + 3000 + 400 > maxWidth Then
                maxWidth = x + 3000 + 400
            End If
            
            CurrentCol = 0
            x = 400
            If Not isMemo Then
                y = y + 350
            Else
                isMemo = False
                y = y + 350 + 600
            End If


        Else
        
            x = x + (3200 * rs.fields("Columns"))
            
            
        End If
NextField:
        
        rs.MoveNext
    Loop
    
    ''Cancel Button
    If Not isPresent("qryModelProperties", "Property = ""dontRenderTimestampCreatedBy"" And ModelID = " & ModelID) Then
        ''Create the Timestamp and CreatedBy field (Hidden Fields)
        Set ctl = CreateControl(frm.Name, acTextBox, , "", "Timestamp", 0, 0, 0)
        ctl.Name = "Timestamp"
        SetControlProperties ctl
        ctl.ColumnWidth = 2000
        
        Set ctl = CreateControl(frm.Name, acComboBox, , "", "CreatedBy", 0, 0, 0)
        ctl.Name = "CreatedBy"
        ctl.Properties("DatasheetCaption") = "Created By"
        SetControlProperties ctl
        
        frm("Timestamp").Locked = True
        frm("CreatedBy").Locked = True
    End If
    
    frm.width = (FormColumns * 3000) + (FormColumns * 400) - 200

    ''Attach the form validation
    
    ''Buttons
    frm.Section("Detail").height = y + 800
    
    ''Special Function to run on form creation
    If Not IsNull(OnFormCreate) Then
        Run OnFormCreate, frm, 5
    End If
    
    ''Override Properties Here
    OverrideProperties ModelID, 5, frm
    
    Dim frmName As String, customFrmName As String, baseFormName As String, i As Integer
    frmName = frm.Name
    
    If Not IsNull(VerbosePlural) Then
        baseFormName = concat("dsht", replace(VerbosePlural, " ", ""))
    Else
        baseFormName = concat("dsht", Model, "s")
    End If
    
    If Not IsNull(subformName) Then
        baseFormName = concat("dsht", subformName)
    End If
    
    DoCmd.Close acForm, frm.Name, acSaveYes
    
    customFrmName = baseFormName
    
'    Do Until Not FrmExist(customFrmName)
'
'        If MsgBox(customFrmName & " already exists. Would you like to replace it?", vbYesNo) = vbYes Then
'            Exit Do
'        End If
'        i = i + 1
'        customFrmName = baseFormName & "_" & i
'
'    Loop
    
    ''Insert the newly created form to the InsertToModelRelatedObjects
    InsertToModelRelatedObjects ModelID, acForm, customFrmName
    
    DoCmd.Rename customFrmName, acForm, frmName
    
    If Not DontOpen Then DoCmd.OpenForm customFrmName, acFormDS
    
    InsertFormInFormForRights customFrmName, Model
    
End Function


Public Function GenerateFields(frm As Form)

    Dim rsName
    
    Dim ModelID, Model, VerboseName, VerbosePlural, MainField, TableWideValidation, FormColumns, SetFocus, IsKeyVisible, QueryName, OnFormCreate, Timestamp, CreatedBy
    
    ModelID = frm("ModelID")
    Model = frm("Model")
    VerboseName = frm("VerboseName")
    VerbosePlural = frm("VerbosePlural")
    MainField = frm("MainField")
    TableWideValidation = frm("TableWideValidation")
    FormColumns = frm("FormColumns")
    SetFocus = frm("SetFocus")
    IsKeyVisible = frm("IsKeyVisible")
    QueryName = frm("QueryName")
    OnFormCreate = frm("OnFormCreate")
    Timestamp = frm("Timestamp")
    CreatedBy = frm("CreatedBy")

    rsName = GetTableName(Model, VerbosePlural)
    
    If Not IsNull(QueryName) Then
        rsName = QueryName
    End If
    
    Dim rsObj As Object, db As DAO.Database, fld As DAO.Field
    Set db = CurrentDb
    If DoesPropertyExists(db.TableDefs, rsName) Then
        Set rsObj = db.TableDefs(rsName)
    Else
        Set rsObj = db.QueryDefs(rsName)
    End If
    
    Dim fieldArr As New clsArray, valueArr As New clsArray, PrimaryKey
    
    PrimaryKey = concat(Model, "ID")
    
    fieldArr.arr = "ModelID,ModelField,FieldTypeID,FieldOrder,ColumnWidth,FieldSource"
    
    Dim ModelField, i As Integer, maxFieldOrder
    maxFieldOrder = ELookup("tblModelFields", "ModelID = " & ModelID & " AND FieldOrder IS NOT NULL", "FieldOrder", "FieldOrder DESC")
    If maxFieldOrder = "" Then
        i = 1
    Else
        i = CInt(maxFieldOrder) + 1
    End If
    
    
    For Each fld In rsObj.fields
            
        ModelField = fld.Name
        If Not isPresent("tblModelFields", "ModelField = " & EscapeString(ModelField) & _
                                          " And ModelID = " & ModelID) Then
            Select Case ModelField
                Case PrimaryKey, "Timestamp", "CreatedBy", "RecordImportID":
                    ''Empty
                Case Else:
                    Set valueArr = New clsArray
                    valueArr.Add ModelID
                    valueArr.Add EscapeString(fld.Name)
                    valueArr.Add fld.Type
                    valueArr.Add i
                    valueArr.Add "Null"
                    valueArr.Add EscapeString(fld.SourceTable)
                    
                    RunSQL "INSERT INTO tblModelFields (" & fieldArr.JoinArr & ") VALUES (" & valueArr.JoinArr & ")"
                    
                    i = i + 1
            End Select
            
        Else
        
            RunSQL "UPDATE tblModelFields SET FieldSource = " & EscapeString(fld.SourceTable) & " WHERE ModelField = " & EscapeString(ModelField) & _
                                          " And ModelID = " & ModelID
            
        End If
        
    Next fld
    
    'DoCmd.OpenForm "frmModels", , , "ModelID = " & ModelID

End Function

Public Function EnumerateSubformFields(frm As Form)
    
    ''Loop at all the models in which the current model is a ParentModelID
    ''Look at each fields and enumerate all the dbInteger and dbDouble FieldTypeIDs
    Dim ModelID
    
    ModelID = frm("ModelID")
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, sqlStr2, rowsAffected, rs As Recordset
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblModelFields"
        .fields = "tblModelFields.*, Model, VerbosePlural"
        .AddFilter "ParentModelID = " & ModelID
        .Joins.Add GenerateJoinObj("tblModels", "ModelID")
        Set rs = .Recordset
    End With
    
    Do Until rs.EOF
    
        Dim subformName
        
        If IsNull(rs("VerboseChildName")) Then
            If IsNull(rs.fields("VerbosePlural")) Then
                subformName = concat("sub", rs.fields("Model"), "s")
            Else
                subformName = concat("sub", rs.fields("VerbosePlural"))
            End If
        Else
            subformName = RemoveSpaces(concat("sub", rs("VerboseChildName")))
        End If
        
        Dim ParentModelID
        ParentModelID = rs.fields("ModelID")
        
        ''SELECT All the tblSubformControls of this ModelID
        Set sqlObj = New clsSQL
        With sqlObj
            .Source = "tblSubformControls"
            .AddFilter "ModelID = " & ModelID
            sqlStr2 = .sql
        End With
        
        ''SELECT all the fields from the ParentModelID's ModelID
        Set sqlObj = New clsSQL
        With sqlObj
            .Source = "tblModelFields"
            .AddFilter "tblModelFields.FieldTypeID In (" & dbInteger & "," & dbDouble & ") AND ModelID = " & ParentModelID
            .fields = "tblModelFields.ModelField As ControlName, " & EscapeString(subformName) & " AS SubformName, " & ModelID & " As ModelID" & _
                      ",AddSpaces([ModelField]) AS ControlCaption, FieldTypeID"
            sqlStr = .sql
        End With
        
        Set sqlObj = New clsSQL
        With sqlObj
            .Source = sqlStr
            .AddFilter "SubformControlID IS NULL"
            .fields = "temp2.*"
            .Joins.Add GenerateJoinObj(sqlStr2, "ControlName,SubformName", "temp", "ControlName,SubformName", "LEFT")
            .SourceAlias = "temp2"
            sqlStr = .sql
        End With
        
        'INSERT STATEMENT
        Set sqlObj = New clsSQL
        With sqlObj
            .SQLType = "INSERT"
            .Source = "tblSubformControls"
            .fields = "ControlName, SubformName, ModelID, ControlCaption, FieldTypeID"
            .InsertSQL = sqlStr
            .InsertFilterField = "ControlName, SubformName, ModelID, ControlCaption, FieldTypeID"
            rowsAffected = .Run
        End With
    
        rs.MoveNext
        
    Loop
    
    If DoesPropertyExists(frm, "subSubformControls") Then
        frm("subSubformControls").Form.Requery
    End If
    
End Function

