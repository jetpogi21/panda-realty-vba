Attribute VB_Name = "FieldsToToggle Mod"
Option Compare Database
Option Explicit

Public Function FieldsToToggleCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function


Public Function OpenHideUnhideForm(frm As Form, subformName)
    
    
    Dim MainFormName
    MainFormName = frm.Name
'
'    Dim ctl As Control
'
'    For Each ctl In frm(SubformName).Form.Controls
'
'        Dim FieldCaption, FieldName
'        FieldCaption = ctl.Properties("DatasheetCaption")
'        FieldName = ctl.Name
'
'        RunSQL "INSERT INTO tblFieldsToToggles (MainFormName,SubformName,FieldCaption,FieldName) VALUES (" & _
'                EscapeString(MainFormName) & "," & EscapeString(SubformName) & "," & EscapeString(FieldCaption) & "," & _
'                EscapeString(FieldName) & ")"
'
'    Next ctl
    DoCmd.OpenForm "mainFieldsToToggles"
    
    Dim frm2 As Form
    Set frm2 = Forms("mainFieldsToToggles")
    frm2("cmdShowALL").OnClick = "=GenericShowAllPropertyListFields([Form]," & EscapeString(MainFormName) & "," & EscapeString(subformName) & ")"
    frm2("subform").Form.Filter = "MainFormName = " & EscapeString(MainFormName) & " AND SubformName = " & EscapeString(subformName)
    frm2("subform").Form.FilterOn = True
    
End Function

Public Function GenericShowAllPropertyListFields(frm As Form, MainFormName, subformName)
    
    Dim rs As Recordset
    Set rs = frm("subform").Form.RecordsetClone
    
    Dim frm2 As Form
    Set frm2 = Forms(MainFormName)
    
    rs.MoveFirst
    Do Until rs.EOF
    
        Dim isShown, FieldName
        isShown = rs.fields("isShown")
        FieldName = rs.fields("FieldName")
        
        rs.Edit
        rs.fields("isShown") = True
        
        frm2(subformName).Form.Controls(FieldName).ColumnHidden = False
        
        rs.Update
        
        rs.MoveNext
    Loop
    
    
End Function

Public Function GenericIsShownAfterUpdate(frm As Form)
    
    
    
    Dim frm2 As Form
    Dim FieldName, isShown As Boolean, MainFormName, subformName
    FieldName = frm("FieldName")
    isShown = frm("IsShown")
    MainFormName = frm("MainFormName")
    subformName = frm("SubformName")
    
    DoCmd.RunCommand acCmdSaveRecord
    
    If IsFormOpen(MainFormName) Then
        Set frm2 = Forms(MainFormName)
        GenericHideOrShowColumn frm2, MainFormName, subformName
    End If
    
End Function

Public Function GenericHideOrShowColumn(frm As Form, MainFormName, subformName)

    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblFieldsToToggles WHERE MainFormName = " & EscapeString(MainFormName) & " AND SubformName = " & EscapeString(subformName) & " ORDER BY FieldOrder")
    
    Dim frmName, isShownField
    frmName = frm.Name
    
    Do Until rs.EOF
    
        Dim isShown, FieldName
        FieldName = rs.fields("FieldName")
        isShown = rs.fields("isShown")
        
        frm(subformName).Form.Controls(FieldName).ColumnHidden = Not isShown
        rs.MoveNext
        
    Loop

End Function
