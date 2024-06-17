Attribute VB_Name = "SaveFormLayout Mod"
Option Compare Database
Option Explicit

Public Function SaveFormLayout(frm As Form)

    If Not areDataValid2(frm, "SaveFormLayout") Then Exit Function
    
    Dim FormName
    FormName = frm("FormName")
           
    Dim PropertyArr As New clsArray
    PropertyArr.arr = "Top,Left,Height,Width"
    
    ''Open the form in design view
    Dim FormID
    FormID = ELookup("tblForms", "FormName = " & EscapeString("FormName"), "FormID")
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    If FormID = "" Then
        Set sqlObj = New clsSQL
        With sqlObj
            .SQLType = "INSERT"
            .Source = "tblForms"
            .fields = "FormName"
            .insertValues = EscapeString(FormName)
            rowsAffected = .Run
            FormID = .LastInsertID
        End With
    End If
    
    ''Delete all the data from tblFormControls using FormID
    RunSQL "DELETE FROM tblFormControls WHERE FormID = " & FormID
    
    Dim ctl As Control, frm2 As Form
    DoCmd.OpenForm FormName, acDesign
    Set frm2 = Forms(FormName)
    
    Dim property
    For Each ctl In frm2.Controls
        For Each property In PropertyArr.arr
            Set sqlObj = New clsSQL
            With sqlObj
                .SQLType = "INSERT"
                .Source = "tblFormControls"
                .fields = "FormControlName,FormID,FormControlProperty,FormControlProperyValue"
                .insertValues = EscapeString(ctl.Name) & "," & _
                                FormID & "," & _
                                EscapeString(property) & "," & _
                                EscapeString(ctl.Properties(property))
                rowsAffected = .Run
            End With
        Next property
    Next ctl
    
    DoCmd.Close acForm, FormName
    MsgBox "Layout Saved..."
    
End Function

Public Function LoadSavedFormLayout(FormName)

    ''This will load the form layout saved
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM qryFormControls WHERE FormName = " & EscapeString(FormName))
    
    If rs.EOF Then Exit Function
    
    DoCmd.OpenForm FormName, acDesign
    
    Dim frm As Form
    Set frm = Forms(FormName)
    
    Dim FormControlID, FormControlName, FormControlProperty, FormControlProperyValue
    
    Do Until rs.EOF
        
        FormControlName = rs.fields("FormControlName")
        FormControlProperty = rs.fields("FormControlProperty")
        FormControlProperyValue = rs.fields("FormControlProperyValue")
        
        frm(FormControlName).Properties(FormControlProperty) = FormControlProperyValue
        rs.MoveNext
        
    Loop
    
    DoCmd.Close acForm, FormName, acSaveYes

End Function
