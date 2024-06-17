Attribute VB_Name = "EmailTemplate Mod"
Option Compare Database
Option Explicit

Public Function EmailTemplateCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function UsableFieldsOnDoubleClick(frm As Form)

    Dim Subject, Body, FieldCaption
    Dim UserResponse As VbMsgBoxResult
     
    FieldCaption = frm("listUsableFields").Column(1)
    
    Dim CompatibleWith: CompatibleWith = frm("listUsableFields").Column(3)
    
    If CompatibleWith = "INDIVIDUAL" Then
        If Not ValidateFieldCaption(frm, "BULK") Then
            UserResponse = MsgBox("Adding individual together with bulk will show the other recipients to the recipient of the email.", vbCritical)
            Exit Function
        End If
    ElseIf CompatibleWith = "BULK" Then
        If Not ValidateFieldCaption(frm, "INDIVIDUAL") Then
            UserResponse = MsgBox("Adding individual together with bulk will show the other recipients to the recipient of the email.", vbCritical)
            Exit Function
        End If
    End If

    
    Dim ActiveField
    ActiveField = frm("txtActiveField")
    
    Subject = frm(ActiveField)
    
    frm(ActiveField) = Subject & " [" & FieldCaption & "]"
    
End Function

Private Function ValidateFieldCaption(frm As Form, CompatibleWith) As Boolean

    Dim Subject: Subject = frm("Subject")
    Dim Body: Body = frm("Body")
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblEmailUsableFields WHERE CompatibleWith = " & EscapeString(CompatibleWith))
    Dim items As New clsArray
    Dim FieldCaption
    Do Until rs.EOF
        FieldCaption = rs.fields("FieldCaption")
        ''Check on both subject and body if the FieldCaption is present on it. If yes then ValidateFieldCaption should be false
        If InStr(Subject, FieldCaption) > 0 Or InStr(Body, FieldCaption) > 0 Then
            ValidateFieldCaption = False
            Exit Function
        End If
        rs.MoveNext
    Loop

    ValidateFieldCaption = True
    
End Function

Public Function SetFieldFocusEmailTemplate(frm As Form, fieldName)

    frm("txtActiveField") = fieldName
    
End Function
