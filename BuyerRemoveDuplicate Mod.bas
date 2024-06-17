Attribute VB_Name = "BuyerRemoveDuplicate Mod"
Option Compare Database
Option Explicit

Public Function BuyerRemoveDuplicateCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function BuyerStatusRemoveDuplicate()

    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblBuyerStatus"
        .fields = "BuyerStatus, Count(BuyerStatusID) As RecordCount"
        .OrderBy = "BuyerStatus"
        .GroupBy = "BuyerStatus"
        .Having = "Count(BuyerStatusID) > 1"
        sqlStr = .sql
        ''makeQuery .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblBuyerStatus"
        ''.AddFilter
        .fields = "BuyerStatusID, tblBuyerStatus.BuyerStatus"
        .Joins.Add GenerateJoinObj(sqlStr, "BuyerStatus", "temp")
        .OrderBy = "tblBuyerStatus.BuyerStatus"
        sqlStr = .sql
        ''makeQuery .sql
    End With
    
    Dim sqlStr1
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = sqlStr
        .fields = "Min(BuyerStatusID) As MinBuyerStatusID, BuyerStatus"
        .OrderBy = "BuyerStatus"
        .GroupBy = "BuyerStatus"
        .SourceAlias = "temp"
        sqlStr1 = .sql
        ''makeQuery .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "MAKE"
        .Source = sqlStr
        .AddFilter "MinBuyerStatusID IS NULL"
        .fields = "BuyerStatusID"
        .Joins.Add GenerateJoinObj(sqlStr1, "BuyerStatusID", "temp1", "MinBuyerStatusID", "LEFT")
        .OrderBy = "BuyerStatusID,MinBuyerStatusID"
        .SourceAlias = "temp"
        sqlStr = .sql
        .MakeTable = "tempBuyerStatusID"
        rowsAffected = .Run
        ''makeQuery .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "DELETE"
        .Source = "tblBuyerStatus"
        .fields = "tblBuyerStatus.*"
        .Joins.Add GenerateJoinObj("tempBuyerStatusID", "BuyerStatusID")
        rowsAffected = .Run
    End With
    
       
End Function

Public Function SellerRemoveDuplicate()
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblEntities"
        .AddFilter "IsSeller = -1"
        .fields = "EntityName,IsSeller,Count(EntityID) As RecordCount"
        .OrderBy = "EntityName,IsSeller"
        .GroupBy = "EntityName,IsSeller"
        .Having = "Count(EntityID) > 1"
        sqlStr = .sql
        ''makeQuery .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblEntities"
        ''.AddFilter
        .fields = "EntityID, tblEntities.EntityName,tblEntities.IsSeller"
        .Joins.Add GenerateJoinObj(sqlStr, "EntityName,IsSeller", "temp")
        .OrderBy = "tblEntities.EntityName,tblEntities.IsSeller"
        sqlStr = .sql
        ''makeQuery .sql
    End With
    
    Dim sqlStr1
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = sqlStr
        .fields = "Min(EntityID) As MinEntityID, EntityName, IsSeller"
        .OrderBy = "EntityName, IsSeller"
        .GroupBy = "EntityName, IsSeller"
        .SourceAlias = "temp"
        sqlStr1 = .sql
        ''makeQuery .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "MAKE"
        .Source = sqlStr
        .AddFilter "MinEntityID IS NULL"
        .fields = "EntityID"
        .Joins.Add GenerateJoinObj(sqlStr1, "EntityID", "temp1", "MinEntityID", "LEFT")
        .OrderBy = "EntityID,MinEntityID"
        .SourceAlias = "temp"
        sqlStr = .sql
        .MakeTable = "tempEntityID"
        rowsAffected = .Run
        ''makeQuery .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "DELETE"
        .Source = "tblEntities"
        .fields = "tblEntities.*"
        .Joins.Add GenerateJoinObj("tempEntityID", "EntityID")
        rowsAffected = .Run
    End With
    
End Function

Public Function BuyerRemoveDuplicate()
    
    ''Group records and get all those that have more than one count e.g. duplicate
    SetIsSellerTo0
    GetEntitiesWithDuplicates
    
End Function

Private Function SetIsSellerTo0()

    RunSQL "UPDATE tblEntities SET isSeller = 0 WHERE isSeller IS NULL"
    
End Function

Private Function GetEntitiesWithDuplicates()

    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblEntities"
        .AddFilter "IsSeller = 0"
        .fields = "EntityName,PhoneNumber,IsSeller,Count(EntityID) As RecordCount"
        .OrderBy = "EntityName,PhoneNumber,IsSeller"
        .GroupBy = "EntityName,PhoneNumber,IsSeller"
        .Having = "Count(EntityID) > 1"
        sqlStr = .sql
        ''makeQuery .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblEntities"
        ''.AddFilter
        .fields = "EntityID, tblEntities.EntityName,tblEntities.PhoneNumber,tblEntities.IsSeller"
        .Joins.Add GenerateJoinObj(sqlStr, "EntityName,PhoneNumber,IsSeller", "temp")
        .OrderBy = "tblEntities.EntityName,tblEntities.PhoneNumber,tblEntities.IsSeller"
        sqlStr = .sql
        ''makeQuery .sql
    End With
    
    Dim sqlStr1
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = sqlStr
        .fields = "Min(EntityID) As MinEntityID, EntityName, PhoneNumber, IsSeller"
        .OrderBy = "EntityName, PhoneNumber, IsSeller"
        .GroupBy = "EntityName, PhoneNumber, IsSeller"
        .SourceAlias = "temp"
        sqlStr1 = .sql
        ''makeQuery .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "MAKE"
        .Source = sqlStr
        .AddFilter "MinEntityID IS NULL"
        .fields = "EntityID"
        .Joins.Add GenerateJoinObj(sqlStr1, "EntityID", "temp1", "MinEntityID", "LEFT")
        .OrderBy = "EntityID,MinEntityID"
        .SourceAlias = "temp"
        sqlStr = .sql
        .MakeTable = "tempEntityID"
        rowsAffected = .Run
        ''makeQuery .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "DELETE"
        .Source = "tblEntities"
        .fields = "tblEntities.*"
        .Joins.Add GenerateJoinObj("tempEntityID", "EntityID")
        rowsAffected = .Run
    End With
    
End Function

