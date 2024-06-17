Attribute VB_Name = "PropertySellerImport Mod"
Option Compare Database
Option Explicit

Public Function PropertySellerImportCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

''ImportPropertySeller(Forms("mainFavoriteProperties").subform.Form)
Public Function ImportPropertySeller(Optional frm As Form, Optional NullProperty As Boolean = False)
    
    Dim PropertyListID
    If NullProperty Then
        PropertyListID = Null
    Else
        PropertyListID = frm("PropertyListID")
    End If
    
    Dim PropertyOwnerSQL: PropertyOwnerSQL = GetPropertyOwnerSQL(PropertyListID)
    
    InsertUniqueOwners PropertyOwnerSQL
    InsertOwnersToProperties PropertyOwnerSQL
    
End Function

Public Function GetPropertyOwnerSQL(Optional PropertyListID = Null)
    
    Dim i As Integer, sqlArr As New clsArray, sqlStr
    
    For i = 1 To 3
        sqlStr = "SELECT StreetAddress, CombinedOwner, Owner" & i & _
                 "Name As OwnerName, Owner" & i & _
                 "Address As Address FROM tblPropertyList WHERE Not Owner" & i & _
                 "Name IS NULL And Owner" & i & "Name <> """""
                 
        If Not IsNull(PropertyListID) Then
            sqlStr = sqlStr & " AND PropertyListID = " & PropertyListID
        Else
            sqlStr = sqlStr & " AND IsFavorite"
        End If
        
        sqlArr.Add sqlStr
        
    Next i
    
    GetPropertyOwnerSQL = sqlArr.JoinArr(" UNION ")
    
End Function

Private Function InsertOwnersToProperties(PropertyOwnerSQL)
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    ''Seller Entities only
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblEntities"
        .AddFilter "EntityCategoryID = 2"
        .fields = "*"
        sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "MAKE"
        .MakeTable = "tempEntityPropertyList"
        .Source = PropertyOwnerSQL
        .fields = "Cdbl(EntityID) As vEntityID,Cdbl(PropertyListID) As vPropertyListID"
        .Joins.Add GenerateJoinObj(sqlStr, "OwnerName", "tempSellers", "EntityName")
        .Joins.Add GenerateJoinObj("tblPropertyList", "StreetAddress,CombinedOwner")
        .SourceAlias = "temp"
        'makeQuery .sql
        rowsAffected = .Run
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "INSERT"
        .Source = "tblPropertyEntities"
        .fields = "EntityID,PropertyListID"
        .InsertSQL = "tempEntityPropertyList"
        .InsertFilterField = "vEntityID,vPropertyListID"
        ''.InsertValues
        .InsertUseAsPlain = True
        ''.LastInsertID
        ''.SQL
        ''makeQuery .SQL
        rowsAffected = .Run
    End With
    
End Function

Private Function InsertUniqueOwners(PropertyOwnerSQL)
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    ''unionPropertyOwners
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "MAKE"
        .Source = PropertyOwnerSQL
        .MakeTable = "tempSellers"
        .fields = "2 As EntityCategoryID,OwnerName As EntityName,Address,-1 As IsSeller"
        .OrderBy = "OwnerName,Address"
        .GroupBy = "OwnerName,Address"
        .SourceAlias = "temp"
        rowsAffected = .Run
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tempSellers"
        .AddFilter "EntityID IS NULL"
        .fields = "tempSellers.EntityCategoryID, tempSellers.EntityName,tempSellers.Address,tempSellers.IsSeller,EntityID"
        .Joins.Add GenerateJoinObj("tblEntities", "EntityName,IsSeller", , , "LEFT")
        .OrderBy = "tempSellers.EntityCategoryID, tempSellers.EntityName,tempSellers.Address,tempSellers.IsSeller"
        sqlStr = .sql
        ''makeQuery .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "INSERT"
        .Source = "tblEntities"
        .fields = "EntityCategoryID,EntityName,Address,IsSeller"
        .InsertSQL = sqlStr
        .InsertFilterField = "EntityCategoryID,EntityName,Address,IsSeller"
        ''.InsertValues
        ''.InsertUseAsPlain = True
        rowsAffected = .Run
    End With
    
    
End Function
