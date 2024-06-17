Attribute VB_Name = "FixPropertyEntity Mod"
Option Compare Database
Option Explicit

Public Function FixPropertyEntityCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function InsertAnyMissingPropertyContactAndFiles()

    Dim sourceTableName: sourceTableName = "qryPropertyContacts"
    Dim SourceDatabaseFile: SourceDatabaseFile = "C:\Users\User\Desktop\Client Files\Richard F\PTS.accdb"
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "MAKE"
          .Source = "[" & sourceTableName & "] IN " & EscapeString(SourceDatabaseFile)
          .MakeTable = "temp_tblPropertyEntities_original"
          .fields = "StreetAddress,SaleDate,EntityID,PropertyEntityNote"
          rowsAffected = .Run
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "temp_tblPropertyEntities_original"
          .fields = "tblPropertyList.PropertyListID,temp_tblPropertyEntities_original.EntityID,temp_tblPropertyEntities_original.PropertyEntityNote"
          .Joins.Add GenerateJoinObj("tblPropertyList", "StreetAddress")
          sqlStr = .sql
''          makeQuery sqlStr
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = sqlStr
          .SourceAlias = "temp"
          .AddFilter "tblPropertyEntities.PropertyEntityID IS NULL"
          .fields = "temp.EntityID,temp.PropertyListID,temp.PropertyEntityNote"
          .Joins.Add GenerateJoinObj("tblPropertyEntities", "EntityID,PropertyListID", , , "LEFT")
          sqlStr = .sql
          ''makeQuery sqlStr
    End With
'
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "tblPropertyEntities"
          .fields = "EntityID,PropertyListID,PropertyEntityNote"
          .InsertSQL = sqlStr
          .InsertFilterField = "EntityID,PropertyListID,PropertyEntityNote"
          .SourceAlias = "temp"
          ''makeQuery .sql
          rowsAffected = .Run
    End With
'
End Function

Public Function InsertAnyMissingPropertyBuyers()

    Dim sourceTableName: sourceTableName = "qryPropertyBuyers"
    Dim SourceDatabaseFile: SourceDatabaseFile = "C:\Users\User\Desktop\Client Files\Richard F\PTS.accdb"
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
'    Set sqlObj = New clsSQL
'    With sqlObj
'          .SQLType = "MAKE"
'          .Source = "[" & sourceTableName & "] IN " & EscapeString(SourceDatabaseFile)
'          .MakeTable = "temp_tblPropertyEntities_original"
'          .Fields = "StreetAddress,SaleDate,EntityID,PropertyEntityNote"
'          rowsAffected = .Run
'    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "temp_tblPropertyEntities_original"
          .fields = "tblPropertyList.PropertyListID,temp_tblPropertyEntities_original.EntityID,temp_tblPropertyEntities_original.PropertyEntityNote"
          .Joins.Add GenerateJoinObj("tblPropertyList", "StreetAddress")
          sqlStr = .sql
''          makeQuery sqlStr
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = sqlStr
          .SourceAlias = "temp"
          .AddFilter "tblPropertyEntities.PropertyEntityID IS NULL"
          .fields = "temp.EntityID,temp.PropertyListID,temp.PropertyEntityNote"
          .Joins.Add GenerateJoinObj("tblPropertyEntities", "EntityID,PropertyListID", , , "LEFT")
          sqlStr = .sql
          ''makeQuery sqlStr
    End With
'
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "tblPropertyEntities"
          .fields = "EntityID,PropertyListID,PropertyEntityNote"
          .InsertSQL = sqlStr
          .InsertFilterField = "EntityID,PropertyListID,PropertyEntityNote"
          .SourceAlias = "temp"
          ''makeQuery .sql
          rowsAffected = .Run
    End With
'
End Function

Public Function InsertAnyMissingEntityFiles()

    Dim sourceTableName: sourceTableName = "qryPropertyEntityFiles"
    Dim SourceDatabaseFile: SourceDatabaseFile = "C:\Users\User\Desktop\Client Files\Richard F\PTS.accdb"
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "MAKE"
          .Source = "[" & sourceTableName & "] IN " & EscapeString(SourceDatabaseFile)
          .MakeTable = "temp_PropertyEntityFiles"
          .fields = "StreetAddress,EntityID,FileType,EntityFileLink"
          rowsAffected = .Run
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "temp_PropertyEntityFiles"
          .fields = "tblPropertyList.PropertyListID,EntityID,FileType,EntityFileLink"
          .Joins.Add GenerateJoinObj("tblPropertyList", "StreetAddress")
          sqlStr = .sql
          'makeQuery sqlStr
    End With
    
'    Set sqlObj = New clsSQL
'    With sqlObj
'          .Source = sqlStr
'          .SourceAlias = "temp"
'          .AddFilter "tblEntityFiles.EntityFileID IS NULL"
'          .fields = "temp.PropertyListID,temp.EntityID,temp.FileType,temp.EntityFileLink"
'          .Joins.Add GenerateJoinObj("tblEntityFiles", "EntityID,EntityFileLink", , , "LEFT")
'          sqlStr = .sql
'          'makeQuery sqlStr
'    End With
'

    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "tblEntityFiles"
          .fields = "EntityID,FileType,EntityFileLink,PropertyListID"
          .InsertSQL = sqlStr
          .InsertFilterField = "EntityID,FileType,EntityFileLink,PropertyListID"
          .SourceAlias = "temp"
          ''makeQuery .sql
          rowsAffected = .Run
    End With
    
End Function

Public Function InsertAnyMissingPropertyEntitiesFromOwn()

    Dim sourceTableName: sourceTableName = "qryPropertyEntities"
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "MAKE"
          .Source = "[" & sourceTableName & "]"
          .MakeTable = "temp_tblPropertyEntities_original"
          .fields = "StreetAddress,SaleDate,EntityID,PropertyEntityNote"
          rowsAffected = .Run
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "temp_tblPropertyEntities_original"
          .fields = "tblPropertyList.PropertyListID,temp_tblPropertyEntities_original.EntityID,temp_tblPropertyEntities_original.PropertyEntityNote"
          .Joins.Add GenerateJoinObj("tblPropertyList", "StreetAddress")
          sqlStr = .sql
''          makeQuery sqlStr
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = sqlStr
          .SourceAlias = "temp"
          .AddFilter "tblPropertyEntities.PropertyEntityID IS NULL"
          .fields = "temp.EntityID,temp.PropertyListID,temp.PropertyEntityNote"
          .Joins.Add GenerateJoinObj("tblPropertyEntities", "EntityID,PropertyListID", , , "LEFT")
          sqlStr = .sql
          ''makeQuery sqlStr
    End With
'
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "tblPropertyEntities"
          .fields = "EntityID,PropertyListID,PropertyEntityNote"
          .InsertSQL = sqlStr
          .InsertFilterField = "EntityID,PropertyListID,PropertyEntityNote"
          .SourceAlias = "temp"
          ''makeQuery .sql
          rowsAffected = .Run
    End With
'
End Function
 
Public Function FixPropertyEntities()
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "MAKE"
        .Source = "qryPropertyEntities"
        .MakeTable = "tempPropertyEntities"
        .AddFilter "EntityCategoryName  = 'Seller'"
        .fields = "PropertyEntityID, PropertyListID,EntityID,Format([Timestamp],""mm/dd/yyyy hh"") As FormattedTS"
        .OrderBy = "PropertyEntityID"
        rowsAffected = .Run
    End With
    
    GeneratePropertyEntititiesToBeDeleted
    CreatePropertiesBasedOnToBeDeleted
    ReimportToBeDeletedOwners
    FixInlineOwners
    DeletePropertyEntities
    
End Function

Public Function DeletePropertyEntities()
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "DELETE"
        .Source = "tblPropertyEntities"
        .Joins.Add GenerateJoinObj("tempPropertyEntitiesToBeDelete", "PropertyEntityID")
        .fields = "tblPropertyEntities.*"
        ''makeQuery .sql
        rowsAffected = .Run
    End With
  
End Function

Private Function FixInlineOwners()
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tempPropertyEntitiesToBeDelete"
        .fields = "PropertyEntityID,EntityID,StreetAddress,tempPropertyEntitiesToBeDelete.FormattedTS"
        .Joins.Add GenerateJoinObj("tblPropertyList", "PropertyListID")
        sqlStr = .sql
        'makeQuery .sql
    End With

    Set sqlObj = New clsSQL
    With sqlObj
        .Source = sqlStr
        .fields = "tblPropertyList.PropertyListID,tblEntities.EntityID,EntityName,Address"
        .Joins.Add GenerateJoinObj("tblPropertyList", "StreetAddress,FormattedTS")
        .Joins.Add GenerateJoinObj("tblEntities", "EntityID")
        .OrderBy = "tblPropertyList.PropertyListID,PropertyEntityID"
        .SourceAlias = "temp"
        ''makeQuery .sql
        Set rs = .Recordset
    End With
    
    Dim currentProperty As Long, i As Integer, OwnerName, ownerAddress, isFirst, setArr As New clsArray
    isFirst = True
    Do Until rs.EOF
        If currentProperty <> rs.fields("PropertyListID") Then
            If isFirst Then
                isFirst = False
            Else
                ''Debug.Print "UPDATE tblPropertyList SET " & setArr.JoinArr(",") & " WHERE PropertyListID = " & currentProperty
                RunSQL "UPDATE tblPropertyList SET " & setArr.JoinArr(",") & " WHERE PropertyListID = " & currentProperty
            End If
            currentProperty = rs.fields("PropertyListID")
            Set setArr = New clsArray
            i = 1
        Else
            i = i + 1
        End If
        OwnerName = "Owner" & i & "Name"
        ownerAddress = "Owner" & i & "Address"
        setArr.Add OwnerName & " = " & EscapeString(rs.fields("EntityName"))
        setArr.Add ownerAddress & " = " & EscapeString(rs.fields("Address"))
        rs.MoveNext
    Loop
    
    If currentProperty <> 0 Then
        ''Debug.Print "UPDATE tblPropertyList SET " & setArr.JoinArr(",") & " WHERE PropertyListID = " & currentProperty
        RunSQL "UPDATE tblPropertyList SET " & setArr.JoinArr(",") & " WHERE PropertyListID = " & currentProperty
    End If
    
    
End Function

Private Function ReimportToBeDeletedOwners()
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tempPropertyEntitiesToBeDelete"
        .fields = "PropertyEntityID,EntityID,StreetAddress,tempPropertyEntitiesToBeDelete.FormattedTS"
        .Joins.Add GenerateJoinObj("tblPropertyList", "PropertyListID")
        sqlStr = .sql
        'makeQuery .sql
    End With

    Set sqlObj = New clsSQL
    With sqlObj
        .Source = sqlStr
        .fields = "tblPropertyList.PropertyListID,EntityID"
        .Joins.Add GenerateJoinObj("tblPropertyList", "StreetAddress,FormattedTS")
        .OrderBy = "PropertyEntityID"
        .SourceAlias = "temp"
        sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "INSERT"
        .Source = "tblPropertyEntities"
        .fields = "EntityID,PropertyListID"
        .InsertSQL = sqlStr
        .InsertFilterField = "EntityID,PropertyListID"
        rowsAffected = .Run
    End With
    
End Function

Private Function GeneratePropertyEntititiesToBeDeleted()
     
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tempPropertyEntities"
        .fields = "PropertyListID,Min(FormattedTS) As vFormattedTS"
        .GroupBy = "PropertyListID"
        .OrderBy = "PropertyListID"
        sqlStr = .sql
        ''makeQuery .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "MAKE"
        .Source = "tempPropertyEntities"
        .MakeTable = "tempPropertyEntitiesToBeDelete"
        .AddFilter "vFormattedTS IS NULL"
        .fields = "PropertyEntityID,tempPropertyEntities.PropertyListID,EntityID,FormattedTS"
        .Joins.Add GenerateJoinObj(sqlStr, "PropertyListID,FormattedTS", "temp", "PropertyListID,vFormattedTS", "LEFT")
        .OrderBy = "tempPropertyEntities.PropertyListID,EntityID,FormattedTS"
        rowsAffected = .Run
    End With
    
End Function

Private Function CreatePropertiesBasedOnToBeDeleted()
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tempPropertyEntitiesToBeDelete"
        .fields = "PropertyListID,FormattedTS"
        .OrderBy = "PropertyListID,FormattedTS"
        .GroupBy = "PropertyListID,FormattedTS"
        sqlStr = .sql
        ''makeQuery .sql
    End With
    
    ''Select Properties with Blank Owner names and addresses , Blank PropertyListID and a Custom CombinedOwner
    Dim PropertyListFields
    PropertyListFields = GetPropertyListFields
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblPropertyList"
        .fields = PropertyListFields & ",temp.FormattedTS, GetCombinedOwners(temp.PropertyListID) As CombinedOwner"
        .Joins.Add GenerateJoinObj(sqlStr, "PropertyListID", "temp")
        .OrderBy = "temp.PropertyListID"
        sqlStr = .sql
        'makeQuery .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "INSERT"
        .Source = "tblPropertyList"
        .fields = PropertyListFields & ",FormattedTS,CombinedOwner"
        .InsertSQL = sqlStr
        .InsertFilterField = PropertyListFields & ",FormattedTS,CombinedOwner"
        rowsAffected = .Run
    End With
    
End Function

Public Function GetCombinedOwners(PropertyListID) As String
        
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tempPropertyEntitiesToBeDelete"
        .AddFilter "PropertyListID = " & PropertyListID
        .fields = "EntityName"
        .Joins.Add GenerateJoinObj("tblEntities", "EntityID")
        .OrderBy = "PropertyEntityID"
        Set rs = .Recordset
    End With
    
    Dim OwnerArr As New clsArray
    Do Until rs.EOF
        OwnerArr.Add rs.fields("EntityName")
        rs.MoveNext
    Loop
    
    If OwnerArr.Count > 0 Then GetCombinedOwners = OwnerArr.JoinArr(" ")
        
End Function

Public Function GetPropertyListFields() As String
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblPropertyList WHERE PropertyListID = 0")
    Dim fld As Field, fieldArr As New clsArray
    Dim exemptedFields As New clsArray
    exemptedFields.arr = "PropertyListID,CombinedOwner,Timestamp,CreatedBy,FormattedTS"
    
    For Each fld In rs.fields
        If Not fld.Name Like "Owner*" And Not exemptedFields.InArray(fld.Name) Then
            fieldArr.Add fld.Name
        End If
    Next fld
    
    GetPropertyListFields = fieldArr.JoinArr(",")
    
End Function
