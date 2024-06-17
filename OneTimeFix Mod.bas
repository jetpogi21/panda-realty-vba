Attribute VB_Name = "OneTimeFix Mod"
Option Compare Database
Option Explicit

''https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/create-table-statement-microsoft-access-sql

Public Function OneTimeFixCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function UpdateLandSizeToDouble()
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    RunSQLOnBackend BEPath, "ALTER TABLE [tblPropertyList] Add COLUMN [LandSize2] FLOAT"
    RunSQLOnBackend BEPath, "UPDATE [tblPropertyList] SET [LandSize2] = [LandSize]"
    RunSQLOnBackend BEPath, "ALTER TABLE [tblPropertyList] DROP COLUMN [LandSize]"
    RunSQLOnBackend BEPath, "ALTER TABLE [tblPropertyList] Add COLUMN [LandSize] FLOAT"
    RunSQLOnBackend BEPath, "UPDATE [tblPropertyList] SET [LandSize] = [LandSize2]"
    RunSQLOnBackend BEPath, "ALTER TABLE [tblPropertyList] DROP COLUMN [LandSize2]"
    
End Function

Public Function AddTaskNoteTotblTasks()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    RunSQLOnBackend BEPath, "ALTER TABLE tblTasks ADD COLUMN [TaskNote] MEMO"
    
    ''Fix the TaskNotes for each task using the GetTaskNotes(TaskID)
    RunSQL "UPDATE tblTasks SET TaskNote = GetTaskNotes(TaskID)"
    
    DoCmd.CopyObject BEPath, "Tasks", acQuery, "qryTasksToExport"
    
End Function

Public Function AddYearBuiltTotblPropertyListTemp()
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    RunSQLOnBackend BEPath, "ALTER TABLE tblPropertyListTemp ADD COLUMN [Parcel Details] TEXT(255)"
    RunSQLOnBackend BEPath, "ALTER TABLE tblPropertyListTemp ADD COLUMN [Year Built] TEXT(255)"
    RunSQLOnBackend BEPath, "ALTER TABLE tblPropertyListTemp ADD COLUMN [Owner Type] TEXT(255)"
    
End Function

Public Function AddEventOrderTo_tblEventList()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    RunSQLOnBackend BEPath, "ALTER TABLE tblEventList ADD COLUMN [EventOrder] DOUBLE"
    
    ''Fix the TaskNotes for each task using the GetTaskNotes(TaskID)
    ''RunSQL "UPDATE tblEventList SET EventOrder = EventListID"

End Function

Public Function AddNewFieldsTo_tblTasks()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    Dim sqlStr: sqlStr = "ALTER TABLE tblTasks " & _
        "ADD AttendeeID LONG, " & _
        "MyPandaEmail VARCHAR(255), " & _
        "Reminder INT, " & _
        "MinutesBeforeReminder INT"
    RunSQLOnBackend BEPath, sqlStr
    
    ''Fix the TaskNotes for each task using the GetTaskNotes(TaskID)
    ''RunSQL "UPDATE tblEventList SET EventOrder = EventListID"

End Function

Public Function Add_EventTimelineIDTo_tblTasks()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    Dim sqlStr: sqlStr = "ALTER TABLE tblTasks " & _
        "ADD EventTimelineID LONG"
    RunSQLOnBackend BEPath, sqlStr
    
    ''Fix the TaskNotes for each task using the GetTaskNotes(TaskID)
    ''RunSQL "UPDATE tblEventList SET EventOrder = EventListID"

End Function

Public Function Add_LeadSourceIDTo_tblEntities()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    Dim sqlStr: sqlStr = "ALTER TABLE tblEntities " & _
        "ADD LeadSourceID LONG"
    RunSQLOnBackend BEPath, sqlStr
    
    ''Fix the TaskNotes for each task using the GetTaskNotes(TaskID)
    ''RunSQL "UPDATE tblEventList SET EventOrder = EventListID"

End Function

Public Function RemoveAllEventTimelineTasks()

    RunSQL "DELETE FROM tblTasks WHERE NOT EventTimelineID IS NULL"
    Sync_tblEventTimelines_with_tblTasks False
    
End Function

Public Function AddTimeFieldsTo_tblTasks()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    Dim sqlStr: sqlStr = "ALTER TABLE tblTasks " & _
        "ADD StartTime DATETIME, " & _
        "DueTime DATETIME"
    RunSQLOnBackend BEPath, sqlStr
    
    ''Update the current task's start time and due time
    sqlStr = "UPDATE tblTasks SET StartTime = iif(isNull(StartDate),Null,TimeValue(StartDate)), " & _
        " DueTime = iif(isNull(DueDate),Null,TimeValue(DueDate))"
    RunSQLOnBackend BEPath, sqlStr
    
    ''Fix the TaskNotes for each task using the GetTaskNotes(TaskID)
    ''RunSQL "UPDATE tblEventList SET EventOrder = EventListID"

End Function

Public Function AddNotesTo_tblEventTimelines()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    Dim sqlStr: sqlStr = "ALTER TABLE tblEventTimelines " & _
        "ADD Notes MEMO"
        
    RunSQLOnBackend BEPath, sqlStr
    
End Function

Public Function AddMissingFieldsTo_tblEntities()
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    ''ESuburb, EState, EPostcode, ABN, Mobile, Fax
    Dim sqlStr: sqlStr = "ALTER TABLE tblEntities " & _
         "ADD COLUMN ESuburb TEXT, " & _
         "EState TEXT, " & _
         "EPostcode TEXT, " & _
         "ABN TEXT, " & _
         "Mobile TEXT, " & _
         "Fax TEXT;"

    RunSQLOnBackend BEPath, sqlStr
    
    sqlStr = "ALTER TABLE tblEntities " & _
         "ADD COLUMN Ref TEXT, " & _
         "Contact TEXT"
    
    RunSQLOnBackend BEPath, sqlStr
    
End Function

Public Function AddMissingFieldsTo_tblEntitiesPart2()
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    ''ESuburb, EState, EPostcode, ABN, Mobile, Fax
    Dim sqlStr: sqlStr = "ALTER TABLE tblEntities " & _
         "ADD COLUMN Ref TEXT, " & _
         "Contact TEXT"
    
    RunSQLOnBackend BEPath, sqlStr
    
End Function

Public Function AddMissingFieldsTo_tblEntityMembers()
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    ''ESuburb, EState, EPostcode, ABN, Mobile, Fax
    Dim sqlStr: sqlStr = "ALTER TABLE tblEntityMembers " & _
         "ADD COLUMN ESuburb TEXT, " & _
         "EState TEXT, " & _
         "EPostcode TEXT, " & _
         "ABN TEXT, " & _
         "Mobile TEXT, " & _
         "Fax TEXT, " & _
         "Ref TEXT, " & _
         "Contact TEXT, " & _
         "LicenseNo TEXT"
    
    RunSQLOnBackend BEPath, sqlStr
    
End Function

Public Function AddMissingFieldsTo_tblEntitiesPart3()
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    ''ESuburb, EState, EPostcode, ABN, Mobile, Fax
    Dim sqlStr: sqlStr = "ALTER TABLE tblEntities " & _
         "ADD COLUMN LicenseNo TEXT"
    
    RunSQLOnBackend BEPath, sqlStr
    
End Function


Public Function Add_EventFileTo_tblEventList()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    Dim sqlStr: sqlStr = "ALTER TABLE tblEventList " & _
        "ADD EventFile MEMO"
        
    RunSQLOnBackend BEPath, sqlStr
    
End Function

Public Function Add_DiscountTo_tblPropertyExpenses()
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function

    RunSQLOnBackend BEPath, "ALTER TABLE tblPropertyExpenses ADD COLUMN [DiscountRate] Double"
    
End Function

Public Function Add_AmountReceivedTo_tblPropertyExpenses()
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function

    RunSQLOnBackend BEPath, "ALTER TABLE tblPropertyExpenses ADD COLUMN [AmountReceived] Double"

End Function

Public Function Add_EventTimelineAmountTo_tblEventTimelines()
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function

    RunSQLOnBackend BEPath, "ALTER TABLE tblEventTimelines ADD COLUMN [EventTimelineAmount] Double"

End Function

Public Function CopyEntityQueries()
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located..") Then Exit Function
    
    Dim sqlStr
    sqlStr = GetEntityMemberSQL
    
    ''Buyer,Seller,Tenant,Contact
    Dim catArr As New clsArray, i, qDef As QueryDef, sqlStr1
    catArr.arr = "Buyer,Seller,Tenant,Contact"
    
    For Each i In catArr.arr
        If i = "Contact" Then
            
            sqlStr1 = "SELECT [Company Name],[Member],[Phone Number],[Email Address],Website,[Contact Category],[Timestamp] FROM (" & sqlStr & ") temp WHERE EntityCategoryName = " & EscapeString(i)
            sqlStr1 = sqlStr1 & " ORDER BY [Timestamp]"
            sqlStr1 = "SELECT [Company Name],[Member],[Phone Number],[Email Address],Website,[Contact Category] FROM (" & sqlStr1 & ") temp2 GROUP BY [Company Name],[Member],[Phone Number],[Email Address],Website,[Contact Category] ORDER BY Min([Timestamp])"
            
        ElseIf i = "Buyer" Then
        
            sqlStr1 = "SELECT TOP 20 [Company Name],[Member],[Phone Number],[Email Address],Website,[Contact Category],[Timestamp],[Property Address] FROM (" & sqlStr & ") temp WHERE EntityCategoryName = " & EscapeString(i)
            sqlStr1 = "SELECT [Company Name],[Member],[Phone Number],[Email Address],[Property Address],Website,[Timestamp] FROM (" & sqlStr1 & ") temp"
            sqlStr1 = sqlStr1 & " ORDER BY [Timestamp] ASC"
            
            sqlStr1 = "SELECT [Company Name],[Member],[Phone Number],[Email Address],[Property Address],Website FROM (" & sqlStr1 & ") temp GROUP BY [Company Name],[Member],[Phone Number],[Email Address],[Property Address],Website ORDER BY Min([Timestamp])"
           
        Else
            
            sqlStr1 = "SELECT [Company Name],[Member],[Phone Number],[Email Address],Website,[Contact Category],[Timestamp],[Property Address] FROM (" & sqlStr & ") temp WHERE EntityCategoryName = " & EscapeString(i)
            sqlStr1 = sqlStr1 & " ORDER BY [Timestamp]"
            sqlStr1 = "SELECT [Company Name],[Member],[Phone Number],[Email Address],[Property Address],Website FROM (" & sqlStr1 & ") temp ORDER BY [Timestamp]"
            ''makeQuery sqlStr1
            
        End If
        
        Set qDef = CurrentDb.QueryDefs("qryEntityQueries")
        qDef.sql = sqlStr1
        DoCmd.SetWarnings False
        DoCmd.CopyObject BEPath, i, acQuery, "qryEntityQueries"
        DoCmd.SetWarnings True
    Next i
    
End Function

Public Function Copy_tblBuyerSubcategories()
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located..") Then Exit Function
    
    DoCmd.SetWarnings False
    DoCmd.CopyObject BEPath, "tblBuyerSubcategories", acTable, "tblBuyerSubcategories"
    DoCmd.SetWarnings True
    
End Function

Public Function Copy_tblAddressVariants()
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located..") Then Exit Function
    
    DoCmd.SetWarnings False
    DoCmd.CopyObject BEPath, "tblAddressVariants", acTable, "tblAddressVariants"
    DoCmd.SetWarnings True
    
End Function

Public Sub CopyAllLinkedTablesToBackend()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located..") Then Exit Sub
    
    DoCmd.SetWarnings False
    Dim rs As Recordset: Set rs = ReturnRecordset("Select * from tblLinkedTables Order by LinkedTableID")
    Do Until rs.EOF
        Dim LinkedTableName: LinkedTableName = rs.fields("LinkedTableName")
        DoCmd.CopyObject BEPath, LinkedTableName, acTable, LinkedTableName
        rs.MoveNext
    Loop
    DoCmd.SetWarnings True
    
End Sub

Public Function Copy_tblLeadSourcesAnd_tblContracts()
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located..") Then Exit Function
    
    If Not TableExists(BEPath, "tblLeadSources") Then
        DoCmd.CopyObject BEPath, "tblLeadSources", acTable, "tblLeadSources1"
    End If
    
    If Not TableExists(BEPath, "tblContracts") Then
        DoCmd.CopyObject BEPath, "tblContracts", acTable, "tblContracts1"
    End If
    
    If Not TableExists(BEPath, "tblSuburbs") Then
        DoCmd.CopyObject BEPath, "tblSuburbs", acTable, "tblSuburbs1"
    End If
    
End Function

Public Function Copy_tblExpenseTypesAnd_tblPropertyExpenses()
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located..") Then Exit Function
    
    If Not TableExists(BEPath, "tblExpenseTypes") Then
        DoCmd.CopyObject BEPath, "tblExpenseTypes", acTable, "tblExpenseTypes1"
    End If
    
    If Not TableExists(BEPath, "tblPropertyExpenses") Then
        DoCmd.CopyObject BEPath, "tblPropertyExpenses", acTable, "tblPropertyExpenses1"
    End If
    
End Function

Public Function Copy_tblForm6()
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located..") Then Exit Function
    
    If Not TableExists(BEPath, "tblForm6") Then
        DoCmd.CopyObject BEPath, "tblForm6", acTable, "tblForm61"
    End If
    
End Function

Public Function Copy_tblContractManualChanges()
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located..") Then Exit Function
    
    If Not TableExists(BEPath, "tblContractManualChanges") Then
        DoCmd.CopyObject BEPath, "tblContractManualChanges", acTable, "tblContractManualChanges1"
    End If
    
End Function

Public Function Copy_TrustReconciliationTables()
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located..") Then Exit Function
    
    DoCmd.SetWarnings False
    DoCmd.CopyObject BEPath, "tblTrustReconciliations", acTable, "tblTrustReconciliations"
    DoCmd.CopyObject BEPath, "tblTrustReconciliationBanks", acTable, "tblTrustReconciliationBanks"
    DoCmd.CopyObject BEPath, "tblTrustReconciliationLedgers", acTable, "tblTrustReconciliationLedgers"
    DoCmd.SetWarnings True
    
End Function

Public Function Migrate_tblEventList_tblEventTimelines()
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located..") Then Exit Function
    
    DoCmd.SetWarnings False
    DoCmd.CopyObject BEPath, "tblEventList", acTable, "tblEventList"
    DoCmd.CopyObject BEPath, "tblEventTimelines", acTable, "tblEventTimelines"
    DoCmd.SetWarnings True
    
End Function

Public Function Modify_tblPaymentDetails()
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located..") Then Exit Function
    
    DoCmd.SetWarnings False
    RunSQLOnBackend BEPath, "ALTER TABLE tblPaymentDetails " & _
             "ADD COLUMN ForPanda YesNo"
             
    RunSQLOnBackend BEPath, "CREATE UNIQUE INDEX idx_PaymentDetail " & _
     "ON tblPaymentDetails (PaymentDetail)"
    DoCmd.SetWarnings True
    
End Function

Public Function MigratePropertyLedgerTables()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located..") Then Exit Function
    
    RunSQL "DELETE FROM tblTrustReceipts"
    RunSQL "DELETE FROM tblTrustPayments"
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblLinkedTables WHERE LinkedTableID Between 26 And 37")
    DoCmd.SetWarnings False
    Do Until rs.EOF
        Dim LinkedTableName: LinkedTableName = rs.fields("LinkedTableName")
        DoCmd.CopyObject BEPath, LinkedTableName, acTable, LinkedTableName
        rs.MoveNext
    Loop
    DoCmd.SetWarnings True
    
End Function

Public Function GetEntityMemberSQL()
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblEntities"
        .fields = "EntityID,EntityName,PhoneNumber,EmailAddress,Website,EntityCategoryName,ContactCategoryName,isSeller"
        .Joins.Add GenerateJoinObj("tblEntityCategories", "EntityCategoryID")
        .Joins.Add GenerateJoinObj("tblContactCategories", "ContactCategoryID", , , "LEFT")
        sqlStr = .sql
    End With
    
    Dim sqlStr1
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblPropertyEntities"
        .fields = "EntityName As [Company Name],MemberName As [Member],MemberPhoneNumber As [Phone Number] ,MemberEmailAddress As [Email Address]," & _
                  "StreetAddress As [Property Address],Website,ContactCategoryName As [Contact Category],EntityCategoryName,tblPropertyEntities.[Timestamp]"
        .Joins.Add GenerateJoinObj("tblEntityMembers", "EntityID")
        .Joins.Add GenerateJoinObj(sqlStr, "EntityID", "temp")
        .Joins.Add GenerateJoinObj("tblPropertyList", "PropertyListID")
        .OrderBy = "tblPropertyEntities.[Timestamp] DESC"
        sqlStr1 = .sql
    End With
    
    Dim sqlStr2
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblPropertyEntities"
        .AddFilter "isFavorite And isSeller"
        .fields = "EntityName As [Company Name],EntityName As [Member Name],PhoneNumber As [Phone Number],EmailAddress As [Email Address]," & _
                  "StreetAddress As [Property Address],Website,ContactCategoryName As [Contact Category],EntityCategoryName,tblPropertyEntities.[Timestamp]"
        .Joins.Add GenerateJoinObj(sqlStr, "EntityID", "temp")
        .Joins.Add GenerateJoinObj("tblPropertyList", "PropertyListID")
        .OrderBy = "tblPropertyEntities.[Timestamp] DESC"
        sqlStr2 = .sql
    End With
    
    GetEntityMemberSQL = sqlStr1 & " UNION " & sqlStr2
    
End Function



Public Function AddCommissionRateTo_tblPropertyList()
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function

    RunSQLOnBackend BEPath, "ALTER TABLE tblPropertyList ADD COLUMN [CommissionRate] Double"
    
End Function

Public Function AddContractDateTo_tblPropertyList()
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function

    RunSQLOnBackend BEPath, "ALTER TABLE tblPropertyList ADD COLUMN [ContractDate] Date"
    
End Function


Public Function AddCombinedOwnerToPropertyList()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function

    RunSQLOnBackend BEPath, "ALTER TABLE tblPropertyList ADD COLUMN [CombinedOwner] CHAR"
    RunSQL "UPDATE tblPropertyList SET CombinedOwner = JoinOwners(Owner1Name, Owner2Name, Owner3Name)"
    
    RunSQLOnBackend BEPath, "ALTER TABLE tblPropertyList ADD COLUMN [FormattedTS] CHAR"
    
End Function

Public Function AddAdvertisementIDTotblPropertyList()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function

    RunSQLOnBackend BEPath, "ALTER TABLE tblPropertyList ADD COLUMN [AdvertisementID] CHAR"
    
End Function

Public Function ModifyPropertyTempIndex()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    RunSQLOnBackend BEPath, "DROP INDEX UniqueProperty ON tblPropertyListTemp"
    RunSQLOnBackend BEPath, "CREATE UNIQUE INDEX UniqueProperty ON tblPropertyListTemp ([Street Address],[Owner 1 Name],[Owner 2 Name],[Owner 3 Name])"
    
End Function

Public Function DropPropertyUniqueIndex()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    RunSQLOnBackend BEPath, "DROP INDEX UniqueProperties ON tblPropertyList"
    
End Function

Public Function RunOneTimeFixes(Optional PreLinkOnly As Boolean = False)

    Dim rs As Recordset
    Set rs = ReturnRecordset("select * from tblOneTimeFixes WHERE Not [Run] AND PreLink = " & PreLinkOnly & " ORDER BY FunctionOrder,OneTimeFixID")
    
    Do Until rs.EOF
        Run rs.fields("FunctionName")
        RunSQL "UPDATE tblOneTimeFixes SET [Run] = -1 WHERE OneTimeFixID = " & rs.fields("OneTimeFixID")
        rs.MoveNext
    Loop
    
    Call DoCmd.NavigateTo("acNavigationCategoryObjectType")
    Call DoCmd.RunCommand(acCmdWindowHide)
    
End Function

Public Function MakeBuyerStatusUnique()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    RunSQLOnBackend BEPath, "CREATE UNIQUE INDEX idxBuyerStatus ON tblBuyerStatus (BuyerStatus) WITH DISALLOW NULL"
    RunSQLOnBackend BEPath, "CREATE UNIQUE INDEX idxEntityNameIsSeller ON tblEntities (EntityName,PhoneNumber,EmailAddress,isSeller) WITH IGNORE NULL"

End Function

Public Function EditblEntitiesField()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    ''CustomerType,ToBeContacted & ToBeContactedDate
    RunSQLOnBackend BEPath, "ALTER TABLE tblEntities ADD COLUMN [CustomerType] CHAR, [ToBeContacted] BIT, [ToBeContactedDate] DATETIME"
    
End Function

Public Function AddPrimaryKeyIndexTotblPropertyEntities()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    ' Define the SQL statement to create the primary key index
    Dim strSql As String
    strSql = "CREATE INDEX PK_PropertyEntityID ON tblPropertyEntities ([PropertyEntityID])"
    
    ' Execute the SQL statement on the backend database
    RunSQLOnBackend BEPath, strSql
    
End Function

Public Function AddIndexTotblEntityExtraFeatures()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    ' Define the SQL statement to create the primary key index
    Dim strSql As String
    strSql = "ALTER TABLE tblEntityExtraFeatures ADD CONSTRAINT idx_tblEntityExtraFeatures UNIQUE (EntityID,BuyerRequirementID,[Value])"
    
    ' Execute the SQL statement on the backend database
    RunSQLOnBackend BEPath, strSql
    
    SyncAllBuyerRequirements
    
End Function

Public Function RemoveDashFromCombinedOwner()
    
    RunSQL "UPDATE tblPropertyList SET CombinedOwner = Replace([CombinedOwner], ' -', '')"
End Function


Public Function AddPropertyAltLinkTotblPropertyList()

    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    RunSQLOnBackend BEPath, "ALTER TABLE tblPropertyList ADD COLUMN [PropertyAltLinks] MEMO"
        
End Function

Public Function AddExcludeFromReportTotblPropertyList()
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located.") Then Exit Function
    
    RunSQLOnBackend BEPath, "ALTER TABLE tblPropertyList ADD COLUMN [ExcludeFromReport] BIT"
    RunSQLOnBackend BEPath, "UPDATE tblPropertyList SET [ExcludeFromReport] = 0"
        
End Function

Public Function CopyActivityStudentsToBE()
    
    ''Change the tblName here"
    Dim tblName
    tblName = "tblActivityStudents"
    
    Dim FEPath, BEPath As String
    GetFEAndBE FEPath, BEPath
    
    If ExitIfTrue(Not fileExists(BEPath), EscapeString(BEPath) & " can't be located..") Then Exit Function
    
    DoCmd.CopyObject BEPath, tblName, acTable, tblName
    RunSQLOnBackend BEPath, "DELETE FROM " & tblName
    
End Function

Private Function GetFEAndBE(FEPath, BEPath)
    
    Dim ProjectPath, FEName, BEName
    ProjectPath = CurrentProject.Path & "\"
    FEName = "PTS.accdb"
    BEName = "PTS Backend.accdb"
    
    FEPath = ProjectPath & FEName
    BEPath = ProjectPath & BEName
    
    If Environ("computername") <> "DESKTOP-3G3V8GO" Then
        ProjectPath = "Z:\MY PANDA APP"
        If Not DirectoryExists(ProjectPath) Then
            ProjectPath = "\\TRUENAS\database\MY PANDA APP"
            If Not DirectoryExists(ProjectPath) Then
                MsgBox "The database tables can't be linked to the backend file. The app will exit.", vbCritical
                DoCmd.Quit
                Exit Function
            End If
        End If
        BEPath = ProjectPath & "\PTS Backend.accdb"
    End If
        
End Function

Public Function FixBackendTable()

    ''Available Field types on tblFieldTypes, Use COUNTER for autonumber fields, FLOAT for DOUBLE, INTEGER (Long) AND SMALLINT
    ''CONSTRAINT MyTableConstraint UNIQUE & (FirstName, LastName, DateOfBirth));"
    ''SSN INTEGER CONSTRAINT MyFieldConstraint PRIMARY KEY"

End Function

Private Function RunSQLOnBackend(BEPath, sqlStr)
    
    Dim db As Database
    Set db = OpenDatabase(BEPath)

On Error GoTo Err_Handler:

    db.Execute sqlStr
    db.Close
    Exit Function
    
Err_Handler:
    
    'MsgBox Err.Description
    db.Close
    Exit Function
   
End Function

