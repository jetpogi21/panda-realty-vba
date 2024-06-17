Attribute VB_Name = "PropertyRemoveDuplicate Mod"
Option Compare Database
Option Explicit

Public Function PropertyRemoveDuplicateCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function PropertyRemoveDuplicate()
    
    
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblPropertyList"
        .fields = "StreetAddress,JoinOwners([Owner1Name],[Owner2Name],[Owner3Name]) as CombinedOwner, Count(PropertyListID) As RecordCount"
        .OrderBy = "StreetAddress,JoinOwners([Owner1Name],[Owner2Name],[Owner3Name])"
        .GroupBy = "StreetAddress,JoinOwners([Owner1Name],[Owner2Name],[Owner3Name])"
        .Having = "Count(PropertyListID) > 1"
        sqlStr = .sql
        ''makeQuery .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "MAKE"
        .Source = sqlStr
        .fields = "StreetAddress,CombinedOwner, RecordCount"
        ''.Joins.Add GenerateJoinObj(sqlStr1, "BuyerStatusID", "temp1", "MinBuyerStatusID", "LEFT")
        ''.OrderBy = "BuyerStatusID,MinBuyerStatusID"
        .SourceAlias = "temp"
        sqlStr = .sql
        .MakeTable = "tempDuplicateProperties"
        rowsAffected = .Run
        ''makeQuery .sql
    End With
    
    ''Make the real combined owner not the current one.
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "MAKE"
        .Source = "tblPropertyList"
        .fields = "PropertyListID,StreetAddress,JoinOwners([Owner1Name],[Owner2Name],[Owner3Name]) as CombinedOwner"
        .MakeTable = "tempPropertyListCombinedOwner"
        rowsAffected = .Run
        ''makeQuery .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tempPropertyListCombinedOwner"
        ''.AddFilter
        .fields = "PropertyListID, tempPropertyListCombinedOwner.StreetAddress, tempPropertyListCombinedOwner.CombinedOwner"
        .Joins.Add GenerateJoinObj("tempDuplicateProperties", "StreetAddress,CombinedOwner")
        .OrderBy = "tempPropertyListCombinedOwner.StreetAddress, tempPropertyListCombinedOwner.CombinedOwner"
        sqlStr = .sql
        ''makeQuery .sql
    End With
    
    Dim sqlStr1
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = sqlStr
        .fields = "Min(PropertyListID) As MinPropertyListID, StreetAddress, CombinedOwner"
        .OrderBy = "StreetAddress, CombinedOwner"
        .GroupBy = "StreetAddress, CombinedOwner"
        .SourceAlias = "temp"
        sqlStr1 = .sql
        ''makeQuery .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "MAKE"
        .Source = sqlStr
        .AddFilter "MinPropertyListID IS NULL"
        .fields = "PropertyListID"
        .Joins.Add GenerateJoinObj(sqlStr1, "PropertyListID", "temp1", "MinPropertyListID", "LEFT")
        .OrderBy = "PropertyListID,MinPropertyListID"
        .SourceAlias = "temp"
        sqlStr = .sql
        .MakeTable = "tempPropertyListID"
        rowsAffected = .Run
        ''makeQuery .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "DELETE"
        .Source = "tblPropertyList"
        .fields = "tblPropertyList.*"
        .Joins.Add GenerateJoinObj("tempPropertyListID", "PropertyListID")
        rowsAffected = .Run
    End With
      
End Function


