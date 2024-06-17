Attribute VB_Name = "Entity Mod"
Option Compare Database
Option Explicit

Public Function FixEntityTable()
    ''Remove all Entities with blank id
    RunSQL "DELETE FROM tblEntities WHERE EntityID IS NULL"
End Function

Public Function EntityCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            'frm.OnLoad = "=BuyerFormOnLoad([Form])"
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function


'Public Function BuyerFormOnLoad(frm As Form)
'
'     DefaultFormLoad frm, "BuyerID"
'     frm("subPropertyBuyers").Form.AllowAdditions = False
'     frm("subPropertyBuyers").Form.AllowEdits = False
'     frm("subPropertyBuyers").Form.AllowDeletions = False
'
'End Function

Public Function FixEntityMembers()

    ''Select entities without any members
    Dim sqlStr
    sqlStr = SelectEntitiesWithoutMembers
    InsertEntitiesWithoutMembers sqlStr
    
End Function

Private Function SelectEntitiesWithoutMembers()

    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblEntityMembers"
        .AddFilter "EntityMemberID IS NULL AND EntityCategoryID <> 2"
        .fields = "tblEntities.EntityID, EntityName As MemberName, Address AS MemberAddress, PhoneNumber As MemberPhoneNumber, EmailAddress As MemberEmailAddress"
        .Joins.Add GenerateJoinObj("tblEntities", "EntityID", , , "RIGHT")
        SelectEntitiesWithoutMembers = .sql
    End With
    
End Function

Private Function InsertEntitiesWithoutMembers(sqlStr)

    ''makeQuery sqlStr
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "INSERT"
        .Source = "tblEntityMembers"
        .fields = "EntityID,MemberName,MemberAddress,MemberPhoneNumber,MemberEmailAddress"
        .InsertSQL = sqlStr
        ''.InsertFilterField
        ''.InsertValues
        .InsertUseAsPlain = True
        ''.LastInsertID
        ''.SQL
        ''makeQuery .SQL
        rowsAffected = .Run
    End With
  
End Function

Private Function SaveRecord()

    On Error Resume Next
    DoCmd.RunCommand acCmdSaveRecord
    
End Function

Public Function OnEntityInsert(frm As Form)
    
    ''SaveRecord
    
    Dim EntityID
    EntityID = frm("EntityID")
    
    ''Check if the entity member for this entity is empty..
    Dim MemberCount
    MemberCount = ECount("tblEntityMembers", "EntityID = " & EntityID)
    
    ''Run Insert -> copy the infos from the entity
    Dim EntityName, Address, PhoneNumber, EmailAddress, EntityCategoryName
    EntityName = frm("EntityName")
    Address = frm("Address")
    PhoneNumber = frm("PhoneNumber")
    EmailAddress = frm("EmailAddress")
    EntityCategoryName = frm("EntityCategoryName")
    
    Dim EntityMemberID
    
    If MemberCount = 0 Then
    
        RunSQL "INSERT INTO tblEntityMembers (EntityID,MemberName,MemberAddress,MemberPhoneNumber,MemberEmailAddress) VALUES (" & _
            EntityID & "," & EscapeString(EntityName) & "," & EscapeString(Address) & "," & EscapeString(PhoneNumber) & _
            "," & EscapeString(EmailAddress) & ")"
            
        EntityMemberID = ELookup("tblEntityMembers", "EntityID = " & EntityID, "EntityMemberID", "[Timestamp] ASC")
        
    Else
        
        EntityMemberID = ELookup("tblEntityMembers", "EntityID = " & EntityID, "EntityMemberID", "[Timestamp] ASC")
        RunSQL "UPDATE tblEntityMembers SET EntityID = " & EntityID & _
                                            ",MemberName = " & EscapeString(EntityName) & _
                                            ",MemberAddress = " & EscapeString(Address) & _
                                            ",MemberPhoneNumber = " & EscapeString(PhoneNumber) & _
                                            ",MemberEmailAddress = " & EscapeString(EmailAddress) & _
                " WHERE EntityMemberID = " & EntityMemberID
        ''ExportEntityToExcel EntityMemberID

    End If
    
    ''Update the LastViewedProperty here to be the PropertyListID
    Dim PropertyListID: PropertyListID = frm("PropertyListID")
    If IsNull(PropertyListID) Then
        ''Get the parent
        PropertyListID = frm.Parent.Form("PropertyListID")
    End If
    frm("LastViewedProperty") = PropertyListID
    
    If IsFormOpen("frmPropertyList") Then
        Dim subformNameArr As New clsArray, subformItem
        subformNameArr.arr = "Buyer,Contact,Tenant"
        For Each subformItem In subformNameArr.arr
            Forms("frmPropertyList")("sub" & subformItem & "Members").Form.Requery
        Next subformItem
    End If
    
    SyncBuyerRequirements Forms("frmCustomDashboard"), PropertyListID
    
    If IsFormOpen("frmTasks") Then
        frm("frmTasks")("AttendeeID").Requery
    End If
    
End Function

Public Function FixLastViewedProperty()

    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblEntities WHERE EntityCategoryID = 1 ORDER BY EntityID")
    Do Until rs.EOF
        Dim EntityID: EntityID = rs.fields("EntityID")
        Dim PropertyListID: PropertyListID = ELookup("qryPropertyBuyers", "EntityID = " & EntityID, "PropertyListID", "PropertyEntityID DESC")
        
        If Not isFalse(PropertyListID) Then
            rs.Edit
            rs.fields("LastViewedProperty") = PropertyListID
            rs.Update
            ''Update the LastViewedProperty of the tblEntities to PropertyListID
        End If

        rs.MoveNext
    Loop
    
End Function

Public Function RemovePropertyEntityDuplicates()
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "MAKE"
        .Source = "qryPropertyEntitiesToBeDeleted"
        .MakeTable = "tempPropertyEntitiesToBeDelete"
        .fields = "*"
        rowsAffected = .Run
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "DELETE"
        .Source = "tblPropertyEntities"
        .fields = "tblPropertyEntities.*"
        .Joins.Add GenerateJoinObj("tempPropertyEntitiesToBeDelete", "PropertyEntityID")
        rowsAffected = .Run
    End With
    
    MsgBox "Property Entity Fixed"
    
    
End Function


