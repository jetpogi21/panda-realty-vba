Attribute VB_Name = "PropertyReportField Mod"
Option Compare Database
Option Explicit

Public Function PropertyReportFieldCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function ShowAllPropertyReportFields(frm As Form, Optional ShowAll As Boolean = True)

    RunSQL "UPDATE tblPropertyReportFields SET Shown = " & ShowAll
    frm("subform").Form.Requery
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblPropertyReportFields WHERE Shown = " & ShowAll)
    Do Until rs.EOF
        Dim DatabaseField: DatabaseField = rs.fields("DatabaseField")
        ToggleColumnPropertyReportFieldsVisibility DatabaseField, ShowAll
        rs.MoveNext
    Loop
    
End Function


Public Function ToggleColumnPropertyReportFieldsVisibility(ctlName, Shown As Boolean)
    
    Dim frm As Form
    If Not CurrentProject.AllForms("mainPropertyReports").IsLoaded Then
        Exit Function
    End If
    
    Set frm = Forms("mainPropertyReports")("subform").Form
    frm(ctlName).ColumnHidden = Not Shown
    
End Function
