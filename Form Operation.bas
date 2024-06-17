Attribute VB_Name = "Form Operation"
Option Compare Database
Option Explicit

Public Function SaveFormData(frm As Form, tblName As String, PrimaryKey As String, Optional validationSuccessCB As String) As Boolean

    If areDataValid(frm, tblName) Then
    
        If validationSuccessCB <> "" Then
            Run validationSuccessCB, frm
        End If
        
        If Not frm.NewRecord Then
            UpdateFormData frm, tblName, PrimaryKey
        End If
        
    End If
    
End Function

Public Function GetOldValue(tblName, FieldName, PrimaryKey, recordID As Variant) As Variant
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM " & tblName & " WHERE " & PrimaryKey & " = " & recordID)
    
    GetOldValue = rs(FieldName)
    
End Function

Private Function UpdateFormData(frm As Form, tblName As String, PrimaryKey As String)
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblFormFields WHERE TableName = """ & tblName & """")
    
    Dim ctl As Control, recordID As Variant
    
    Dim FieldName As String, FieldTypeID As Integer, currentValue, oldValue
    Dim updateStatement() As String, i As Integer
    
    Do Until rs.EOF
        FieldName = rs.fields("FieldName")
        FieldTypeID = rs.fields("FieldTypeID")
        If ControlExists(FieldName, frm) Then
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
        rs.MoveNext
    Loop
    
End Function

Private Function InsertFormData(frm As Form, tblName As String)
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblFormFields WHERE TableName = """ & tblName & """")
    
    Dim ctl As Control
    
    Dim FieldName As String, FieldTypeID As Integer
    Dim fields() As String, fieldValues() As String, i As Integer
    
    Do Until rs.EOF
        FieldName = rs.fields("FieldName")
        FieldTypeID = rs.fields("FieldTypeID")
        If ControlExists(FieldName, frm) Then
            ReDim Preserve fields(i)
            ReDim Preserve fieldValues(i)
            
            fields(i) = FieldName
            fieldValues(i) = ReturnStringBasedOnType(frm(FieldName), FieldTypeID)
            
            i = i + 1
            
        End If
        rs.MoveNext
    Loop
    
    InsertAndLog tblName, fields(), fieldValues()
    
End Function

Public Function CancelEdit(frm As Form, Optional isChild As Boolean = False)
    frm.Undo
    If isChild Then
        DoCmd.Close acForm, frm.Parent.Form.Name, acSaveNo
    Else
        DoCmd.Close acForm, frm.Name, acSaveNo
    End If
End Function

Public Function Save(frm As Form, tblName As String, operation As Integer, Optional isChild As Boolean = False)
    
    ''Operation: 0 is save and new, 1 is save and close
    If areDataValid(frm, tblName) Then
    
'        If validationSuccessCB <> "" Then
'            Run validationSuccessCB, frm
'        End If
        
        Select Case operation
            Case 0:
                DoCmd.GoToRecord acDataForm, frm.Name, acNewRec
            Case 1:
                If isChild Then
                    DoCmd.Close acForm, frm.Parent.Form.Name, acSaveNo
                Else
                    DoCmd.Close acForm, frm.Name, acSaveNo
                End If
        End Select
    End If
    
End Function

Public Function IsFormOpen(frmName) As Boolean
    On Error GoTo Err_Handler:
    
    IsFormOpen = CurrentProject.AllForms(frmName).IsLoaded
    Exit Function
Err_Handler:
    
    IsFormOpen = False
End Function

Public Function Save2(frm As Form, Model As String, operation As Integer, Optional isChild As Boolean = False)

    Dim BeforeUpdate
    BeforeUpdate = frm.BeforeUpdate
    
    frm.BeforeUpdate = ""
    
    ''Operation: 0 is save and new, 1 is save and close
    If areDataValid2(frm, Model) Then
    
'        If validationSuccessCB <> "" Then
'            Run validationSuccessCB, frm
'        End If
        
        Select Case operation
            Case 0:
                DoCmd.GoToRecord acDataForm, frm.Name, acNewRec
                frm.BeforeUpdate = BeforeUpdate
            Case 1:
                If isChild Then
                    DoCmd.Close acForm, frm.Parent.Form.Name, acSaveNo
                Else
                    DoCmd.Close acForm, frm.Name, acSaveNo
                End If
            Case Else:
            
                DoCmd.RunCommand acCmdSaveRecord
                frm.BeforeUpdate = BeforeUpdate
                
        End Select
        
    Else
    
        frm.BeforeUpdate = BeforeUpdate
        
    End If
    
End Function

Public Function DefaultFormLoad(frm As Form, PrimaryKey, Optional AutoWidth As Boolean = True)
    
    SetDefaultUserID frm
    
    '''Also hide the subforms if this is a main form
    Dim ctl As Control, subCtl As Control, LinkChildFields
    
    For Each ctl In frm.Controls
        If ctl.ControlType = acSubform Then
            
            If Not ctl.SourceObject Like "Report.*" Then
            
                ctl.Form.DatasheetAlternateBackColor = RGB(254, 254, 254)
                ''Hide the related field
                LinkChildFields = ctl.LinkChildFields
                If frm(ctl.Name).Form.DefaultView = 2 Then
                    For Each subCtl In frm(ctl.Name).Form.Controls
                        
                        If AutoWidth And Not subCtl.Tag Like "*DontAutoWidth*" Then
                            subCtl.ColumnWidth = -2
                        End If
                        
                        subCtl.ColumnHidden = subCtl.Name = LinkChildFields
                        
                        Select Case subCtl.Name
                            Case "Timestamp", "CreatedBy", "VerboseName", "Model":
                            Case Else:
                                SetColumnHidden subCtl, frm
                        End Select
                        
                        If subCtl.Tag Like "*alwaysHideOnDatasheet*" Then
                            subCtl.ColumnHidden = True
                        End If
                        
                    Next subCtl
                End If
            End If
            
        End If
    Next ctl
   
End Function

Private Sub SetColumnHidden(subCtl As Control, frm As Form)

    On Error GoTo ErrHandler:
        
        subCtl.ColumnHidden = DoesPropertyExists(frm, subCtl.Name)
        Exit Sub
    
ErrHandler:
    
    If Err.number = 2101 Then Exit Sub
    
    ShowError "Error # " & Err.number & vbCrLf & Err.Description

End Sub

Public Function DefaultMainFormLoad(frm As Form)
    
    '''Reveal all the hidden fields from the subform
    Dim ctl As Control
    
    If Not frm.subform.SourceObject Like "Report.*" Then
    
        frm.subform.Form.DatasheetAlternateBackColor = RGB(254, 254, 254)
        
        For Each ctl In frm.subform.Form.Controls
            
            ctl.ColumnWidth = -2
            ''Hide the related field
            If ctl.ColumnHidden = True Then
            
                ctl.ColumnHidden = False
                
            End If
            
            If ctl.Tag Like "*alwaysHideOnDatasheet*" Then
                ctl.ColumnHidden = True
            End If
            
        Next ctl
   End If
   
End Function


Public Function OpenFormFromRecord(frm As Form, FieldName, frmName)
    
    Dim fieldValue
    fieldValue = frm(FieldName)
    
    If IsNull(fieldValue) Then Exit Function
    
    DoCmd.OpenForm frmName, , , FieldName & "=" & fieldValue
   
End Function

Public Function CustomMainFormLoad(frm As Form)

    DefaultMainFormLoad frm
    
    ''Disable base on their rights
    CheckUserRights frm

End Function

Public Function CustomFormLoad(frm As Form, ModelID)
    
    DefaultFormLoad frm, ModelID
    CheckUserRights frm
    
End Function

Public Function CustomDatasheetFormLoad(frm As Form)

    SetDefaultUserID frm
    
    ''Disable base on their rights
    CheckUserRights frm

End Function

Private Function CheckUserRights(frm As Form)
    
    
    Dim frmName, frmType, modelName
    frmName = frm.Name
    frmType = GetFormType(frm) ''DataEntry,DataSheet,MainForm
    modelName = ELookup("tblFormForRights", "FormName = '" & frmName & "'", "ModelName")
    
    ResetFormToDefault frm, frmType
    
    If isPresent("tblUsers", "UserID = " & g_UserID & " AND isAdmin") Then Exit Function
    
    ''Check if can be added
    Dim CanAdd, CanEdit, CanDelete
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblUserRights WHERE User = " & g_UserID & " AND ModelName = '" & modelName & "'")
    
    If rs.EOF Then Exit Function
 
    CanAdd = rs.fields("canEdit")
    CanEdit = rs.fields("canEdit")
    CanDelete = rs.fields("canDelete")
    
    If Not CanAdd Then handleCantAdd frm, frmType
    If Not CanEdit Then handleCantEdit frm, frmType
    If Not CanDelete Then handleCantDelete frm, frmType
    
End Function

Private Function handleCantAdd(frm As Form, frmType)

    If frmType = "MainForm" Then
    
        ''Reset to Default
        frm("cmdAdd").Enabled = False
        
    ElseIf frmType = "DataSheet" Then
        
        frm.AllowAdditions = False
        
    ElseIf frmType = "DataEntry" Then
        
        frm("cmdNew").Enabled = False
        frm("cmdSaveClose").Enabled = True
        
    End If
    
End Function

Private Function handleCantEdit(frm As Form, frmType)

    If frmType = "MainForm" Then
    
        ''Reset to Default
        frm("cmdView").Caption = "View"
        
    ElseIf frmType = "DataSheet" Then
        
        frm.AllowEdits = False
        
    ElseIf frmType = "DataEntry" Then
        
        frm("cmdSaveClose").Enabled = False

    End If
    
End Function

Private Function handleCantDelete(frm As Form, frmType)

    If frmType = "MainForm" Then
    
        ''Reset to Default
        frm("cmdDelete").Enabled = False
        
    ElseIf frmType = "DataSheet" Then
        
        frm.AllowDeletions = False
    
    ElseIf frmType = "DataEntry" Then
        
        frm("cmdDelete").Enabled = False

    End If
    
End Function

Private Function ResetFormToDefault(frm As Form, frmType)
    
    On Error Resume Next
    If frmType = "MainForm" Then
    
        ''Reset to Default
        frm("cmdAdd").Enabled = True
        frm("cmdView").Enabled = True
        frm("cmdDelete").Enabled = True
        
    ElseIf frmType = "DataSheet" Then
        
        frm.AllowAdditions = True
        frm.AllowEdits = True
        frm.AllowDeletions = True
        
    ElseIf frmType = "DataEntry" Then
        
        frm("cmdCancel").Enabled = True
        frm("cmdNew").Enabled = True
        frm("cmdSaveClose").Enabled = True
        frm("cmdDelete").Enabled = True
        
    End If

End Function

Private Function GetFormType(frm As Form) As String
    
    Dim frmName
    frmName = frm.Name
    
    GetFormType = "DataEntry"
    
    If frmName Like "main*" Then
    
        GetFormType = "MainForm"
        Exit Function
        
    ElseIf frmName Like "dsht*" Then
            
        GetFormType = "DataSheet"
        Exit Function
            
    End If
    
End Function

Public Function SetDefaultUserID(frm As Form)

    If isFalse(g_UserID) Then
        LogIn
    End If
    
    If DoesPropertyExists(frm, "CreatedBy") Then
        On Error Resume Next
        frm("CreatedBy").DefaultValue = "=" & g_UserID
    End If
    
End Function

Public Function SetFocusOnForm(frm As Form, ctlName As String)
    If ctlName <> "" Then frm(ctlName).SetFocus
End Function

Public Function DeleteRecord(frm As Form, pkName As String, tblName As String, Optional subformName As String = "")
        
    Dim frm2 As Form
    If subformName = "" Then Set frm2 = frm Else Set frm2 = frm(subformName).Form
    
    If frm2.NewRecord Then
        ShowError "You can't delete an unsaved new record."
        Exit Function
    End If
    
    If IsNull(frm2(pkName)) Then
        Exit Function
    End If
    
    If MsgBox("Are you sure you want to delete this record?", vbYesNo) = vbYes Then
    
        ''frmType 0 => DataEntry | 1 => Datasheet
        Dim recordID As Variant
        recordID = frm2(pkName)
        
        RunSQL "DELETE FROM " & tblName & " WHERE " & pkName & " = " & recordID
        
        Insert_Delete_Log tblName, "DELETE", recordID
        
        frm.OnClose = "=RequeryOnClose('" & tblName & "',True)"
         
        If subformName = "" Then frm.Requery Else frm(subformName).Requery
        
    End If
    
End Function

Public Function OpenFormFromMain(frmName As String, Optional subformName As String, Optional PrimaryKey As String, Optional frm As Form, Optional DefaultField As String)

    If DefaultField <> "" Then
        If Not DoesPropertyExists(frm, DefaultField) Then
            ShowError "The parent form is empty.."
            Exit Function
        ElseIf IsNull(frm(DefaultField)) Then
            ShowError "The parent form is empty.."
            Exit Function
        End If
    End If
    
On Error GoTo Err_Handler:
    If PrimaryKey = "" Then
        
        ''2501
        DoCmd.OpenForm frmName, , , , acFormAdd
        If DefaultField <> "" Then
            If DoesPropertyExists(Forms(frmName), DefaultField) Then
                Forms(frmName)(DefaultField).DefaultValue = frm(DefaultField)
            End If
        End If
    Else
        Dim pkVal As Variant
        pkVal = frm(subformName).Form(PrimaryKey)
        If IsNull(pkVal) Then
            ShowError "Please select a record from the list.."
            Exit Function
        End If
        
        DoCmd.OpenForm frmName, , , PrimaryKey & " = " & pkVal
    End If
    
    Exit Function
    
Err_Handler:

    If Err.number = 2501 Then
        Exit Function
    Else
        MsgBox Err.Description
    End If
    
End Function


Public Function RequeryOnClose(tblName As String, Optional shouldRequery As Boolean)
    
    If shouldRequery Then
        On Error Resume Next
        Dim requeryForms As Variant, requeryFormArray() As String, requeryForm As Variant, PrimaryKey
        Dim rs As Recordset, frm As Form, rsClone As Recordset
        Set rs = ReturnRecordset("SELECT * FROM tblTables WHERE TableName = '" & tblName & "'")
        requeryForms = rs.fields("RequeryOnClose")
        PrimaryKey = rs.fields("PrimaryKey")
    
        requeryFormArray = Split(requeryForms)
        For Each requeryForm In requeryFormArray
            Eval "Forms!" & requeryForm & ".Requery"
            Set frm = ReturnFormObject(requeryForm)
            'ReturnMainForm(requeryForm).SetFocus
            Set rsClone = frm.RecordsetClone
            rsClone.FindFirst PrimaryKey & " = " & frm(PrimaryKey)
            If Not rsClone.NoMatch Then
                frm.Bookmark = rsClone.Bookmark
            End If
        Next requeryForm
    End If
    
    
End Function

Public Function ReturnFormObject(requeryForm As Variant) As Form

    Dim formParts() As String
    Dim formPart As Variant
    
    formParts = Split(requeryForm, ".")
    Dim obj As Object
    Set obj = Forms
    
    For Each formPart In formParts
        Set obj = obj(formPart)
    Next formPart
    
    Set ReturnFormObject = obj.Form
    
End Function

Public Function ReturnMainForm(requeryForm As Variant) As Form

    Dim formParts() As String
    Dim formPart As Variant
    
    formParts = Split(requeryForm, ".")
    Dim obj As Object
    Set obj = Forms
    
    Set ReturnMainForm = obj(formParts(0))
    
End Function

Public Function SelectSubformRecords(frm As Form, Optional mode As Boolean = False)
    
    Dim tblName As String, sqlObj As New clsSQL
    tblName = frm.subform.Form.RecordSource
    
    'UPDATE STATEMENT
    With sqlObj
        .SQLType = "UPDATE"
        .Source = tblName
        .SetStatement = "Selected = " & mode
        .Run
    End With
    
    frm.subform.Requery
    
End Function

'Public Function RefreshSubformData(frm As Form, MainFormUtilityID)
'
'    Dim sqlObj As New clsSQL
'    Dim rs As Recordset
'
'    ''FETCH the variables from MainFormUtilities
'    With sqlObj
'        .Source = "tblMainFormUtilities"
'        .AddFilter "MainFormUtilityID = " & MainFormUtilityID
'        Set rs = .Recordset
'    End With
'
'    Dim QueryName, TempTableName, IgnoreFields, AdditionalFields
'    QueryName = rs.Fields("QueryName")
'    TempTableName = rs.Fields("TempTableName")
'    IgnoreFields = rs.Fields("IgnoreFields")
'    AdditionalFields = rs.Fields("AdditionalFields")
'
'    ''Delete the content of the subform
'    ''DELETE STATEMENT
'    Set sqlObj = New clsSQL
'    With sqlObj
'        .SQLType = "DELETE"
'        .Source = TempTableName
'        .Run
'    End With
'
'    ''Insert the query to the content subform
'    ''SELECT STATEMENT
'    Dim fieldNames, sqlStr
'    fieldNames = GenerateFieldNamesString(TempTableName, IgnoreFields) & AdditionalFields
'    Set sqlObj = New clsSQL
'    With sqlObj
'        .Source = QueryName
'        .Fields = fieldNames
'        sqlStr = .SQL
'    End With
'
'    Set sqlObj = New clsSQL
'    With sqlObj
'        .SQLType = "INSERT"
'        .Source = TempTableName
'        .Fields = fieldNames
'        .InsertSQL = sqlStr
'        .Run
'    End With
'
'    frm.subform.Form.Requery
'
'End Function

Public Function DoOpenForm(frmName, Optional whereCondition, Optional addNew As Boolean = True)
    
    If isFalse(whereCondition) Then
         
         If addNew Then
            DoCmd.OpenForm frmName, , , , acFormAdd
        Else
            DoCmd.OpenForm frmName
        End If
    
    Else
         
         DoCmd.OpenForm frmName, , , whereCondition
    
    End If
   
End Function




