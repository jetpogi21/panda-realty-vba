Attribute VB_Name = "PropertyListField Mod"
Option Compare Database
Option Explicit

Public Function PropertyListFieldCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function HideOrShowColumn(frm As Form, Optional CalledFrom = "mainPropertyList")

    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblPropertyListFields ORDER BY FieldOrder")
    
    Dim frmName, isShownField
    frmName = frm.Name
    isShownField = "IsShownFavorite"
    If frmName = "mainPropertyList" Or CalledFrom = "mainPropertyList" Then isShownField = "isShown"
    
    Do Until rs.EOF
    
        Dim isShown, FieldName
        FieldName = rs.fields("FieldName")
        isShown = rs.fields(isShownField)
        
        If ControlExists(FieldName, frm("subform").Form) Then
            frm("subform").Form.Controls(FieldName).ColumnHidden = Not isShown
        End If
        rs.MoveNext
        
    Loop

End Function

Public Function IsShownAfterUpdate(frm As Form, Optional frmName As String = "mainPropertyList")

    Dim frm2 As Form
    Dim FieldName, isShown As Boolean
    FieldName = frm("FieldName")
    isShown = frm("IsShown")
    
    DoCmd.RunCommand acCmdSaveRecord
    
    If IsFormOpen(frmName) Then
        Set frm2 = Forms(frmName)
        HideOrShowColumn frm2, frmName
    End If
    
    ''Also sync the fields of mainPropertyReports
    If IsFormOpen("mainPropertyReports") Then
        Set frm2 = Forms("mainPropertyReports")
        HideOrShowColumn frm2
    End If
    
End Function

Public Function ShowAllPropertyListFields(frm As Form, Optional frmName = "mainPropertyList", Optional notReversed = True)
    
    Dim rs As Recordset
    Set rs = frm("subform").Form.RecordsetClone
    
    Dim frm2 As Form, isShownField
    isShownField = "IsShown"
    If frmName = "mainFavoriteProperties" Then isShownField = "IsShownFavorite"
    Set frm2 = Forms(frmName)
    
    rs.MoveFirst
    Do Until rs.EOF
        Dim isShown, FieldName
        isShown = rs.fields(isShownField)
        FieldName = rs.fields("FieldName")
        
        rs.Edit
        rs.fields(isShownField) = notReversed
        
        frm2("subform").Form.Controls(FieldName).ColumnHidden = False
        
        rs.Update
        
        rs.MoveNext
    Loop
    
    If IsFormOpen(frmName) Then
        HideOrShowColumn Forms(frmName)
    End If
    
    If IsFormOpen("mainPropertyReports") Then
        Set frm2 = Forms("mainPropertyReports")
        HideOrShowColumn frm2, frmName
    End If
    
End Function
