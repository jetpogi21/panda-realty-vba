Attribute VB_Name = "PropertyEntity Mod"
Option Compare Database
Option Explicit

Public Function PropertyEntityCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function RefreshPropertyEntityNotes()

    Dim frmName
    frmName = "frmPropertyList"
    
    If IsFormOpen(frmName) Then
        Forms(frmName)("subPropertyEnityNotes").Form.Requery
    End If
    
End Function


'Public Function PropertyEntityValidation(frm As Form) As Boolean
'
'    PropertyEntityValidation = True
'    If frm.Tag <> "Seller" Then Exit Function
'
'    Dim EntityName
'    EntityName = frm("EntityName")
'
'    If isFalse(EntityName) Then
'        PropertyEntityValidation = False
'        MsgBox "Please provide a valid seller name.", vbCritical
'        Exit Function
'    End If
'
'    Dim rs As Recordset
'    Set rs = ReturnRecordset("SELECT * FROM tblEntities WHERE EntityName = " & EscapeString(EntityName))
'
'    Dim EntityID
'    If rs.EOF Then
'        EntityID = InsertEntity(frm)
'        Set rs = ReturnRecordset("SELECT * FROM tblEntities WHERE EntityName = " & EscapeString(EntityName))
'        EntityID = rs.Fields("EntityID")
'    Else
'        EntityID = rs.Fields("EntityID")
'    End If
'
'    frm("EntityName") = Null
'
'    frm("EntityID").Requery
'    frm("EntityID") = EntityID
'
'End Function
'
'Private Function InsertEntity(frm As Form)
'
'    Dim EntityName, Address, PhoneNumber, EmailAddress, EntityCategoryID
'
'    EntityCategoryID = 2
'    EntityName = frm("EntityName")
'    Address = frm("Address")
'    PhoneNumber = frm("PhoneNumber")
'    EmailAddress = frm("EmailAddress")
'
'    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
'    Set sqlObj = New clsSQL
'    With sqlObj
'        .SQLType = "INSERT"
'        .Source = "tblEntities"
'        .Fields = "EntityName, Address, PhoneNumber, EmailAddress, EntityCategoryID"
'        ''.InsertSQL
'        ''.InsertFilterField
'        .insertValues = EscapeString(EntityName) & "," & _
'                        EscapeString(Address) & "," & _
'                        EscapeString(PhoneNumber) & "," & _
'                        EscapeString(EmailAddress) & "," & _
'                        EntityCategoryID
'        ''.InsertUseAsPlain
'        rowsAffected = .Run
'        InsertEntity = .LastInsertID
'        ''.SQL
'        ''makeQuery .SQL
'
'    End With
'
'End Function
