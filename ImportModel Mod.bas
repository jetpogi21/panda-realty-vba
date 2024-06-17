Attribute VB_Name = "ImportModel Mod"
Option Compare Database
Option Explicit

Public Function ImportModelCurrent(frm As Form)

    SetFocusOnForm frm, ""
    
    ''Set the ModelID Rowsource
    frm("ModelID").rowSource = "SELECT ModelID, Model FROM tblModels WHERE ModelID = 0"
    frm("ModelID").ListItemsEditForm = ""
    
End Function

Public Function BrowserSourceDatabaseFile(frm As Form)
    
    Dim fs As Object
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim fullPath, fileName
    With FileDialog(msoFileDialogFilePicker)
        'Makes sure the user can select only one file
        .AllowMultiSelect = False
        'Filter to just the following types of files to narrow down selection options
        .filters.Add "MS Access Database Files", "*.accdb; *.mdb", 1
        'Show the dialog box
        If ExitIfTrue(Not .Show, "Database selection cancelled..") Then Exit Function
        
        'Store in fullpath variable
        fullPath = .SelectedItems.item(1)
        fileName = fs.GetFileName(fullPath)
        
    End With
    
    frm("SourceDatabaseFile") = fullPath
    ''Select the Models From the database path
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblModels IN " & EscapeString(fullPath)
        .AddFilter "IsSystemTable = 0"
        .OrderBy = "ModelID ASC"
        sqlStr = .sql
        Set rs = .Recordset
    End With
    
    frm("ModelID").rowSource = sqlStr

    
End Function

Public Function BuildSetStatement(rsToBeUpdated, rsToBeUsedToUpdate, FieldNames, Optional IncludeRecordImportID As Boolean = False) As String
    
    Dim fieldArr As New clsArray, setStatementArr As New clsArray, fieldArrItem
    fieldArr.arr = FieldNames
    
    If IncludeRecordImportID Then fieldArr.Add "RecordImportID"
    
    For Each fieldArrItem In fieldArr.arr
        setStatementArr.Add concat(rsToBeUpdated, "!", fieldArrItem, " = ", rsToBeUsedToUpdate, "!", fieldArrItem)
    Next fieldArrItem
    
    BuildSetStatement = setStatementArr.JoinArr(",")
    
End Function

Public Function ImportRelatedModelRecords(ModelID, SourceDatabaseFile)
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    ''Get all the models with ParentModelID of Model
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblModelFields"
        .AddFilter concat("ParentModelID = ", ELookup("tblModels", "Model = " & EscapeString("Model"), "ModelID"))
        .fields = "ModelID"
        .GroupBy = "ModelID"
        sqlStr = .sql
        Set rs = .Recordset
    End With
    
    Dim tblName, PrimaryKey, db As DAO.Database
    Set db = OpenDatabase(SourceDatabaseFile)
    Do Until rs.EOF
        tblName = GetTableNameFromModelID(rs.fields("ModelID"))
        PrimaryKey = concat(ELookup("tblModels", "ModelID = " & rs.fields("ModelID"), "Model"), "ID")
        If DoesPropertyExists(db.TableDefs, tblName) Then
            ImportTableRecords tblName, ModelID, PrimaryKey, SourceDatabaseFile
        End If
        rs.MoveNext
    Loop
    
End Function

Private Function ImportTableRecords(tblName, ModelID, PrimaryKey, SourceDatabaseFile)
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = tblName & " IN " & EscapeString(SourceDatabaseFile)
        .AddFilter "ModelID = " & ModelID
        .OrderBy = PrimaryKey & " ASC"
        sqlStr = .sql
        Set rs = .Recordset
    End With
    
    ''Get the fieldnames to be used for the insert statement
    Dim fieldArr As New clsArray, fieldArrItem, fld As DAO.Field, FieldNames, isParentModel
    
    isParentModel = False
    For Each fld In rs.fields
        Select Case fld.Name
            Case "Timestamp", "CreatedBy", "RecordImportID", PrimaryKey, "ModelID":
            
            Case Else:
                If fld.Name = "ParentModelID" Then isParentModel = True
                fieldArr.Add fld.Name
        End Select
    Next fld
    
    If isParentModel Then
    
        Dim allFieldExceptParentModelIDArr As New clsArray
        For Each fld In rs.fields
            Select Case fld.Name
                Case "ParentModelID":
                    isParentModel = True
                Case Else:
                    allFieldExceptParentModelIDArr.Add "temp." & fld.Name
            End Select
        Next fld
        
        If isParentModel Then
            ''SELECT STATEMENT
            Set sqlObj = New clsSQL
            With sqlObj
                .Source = sqlStr
                .fields = allFieldExceptParentModelIDArr.JoinArr & ", Clng(tblModels.ModelID) As ParentModelID"
                .Joins.Add GenerateJoinObj("tblModels", "ParentModelID", , "RecordImportID", "LEFT")
                .SourceAlias = "temp"
                sqlStr = .sql
            End With
        End If
        
    End If
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "MAKE"
        .Source = sqlStr
        .MakeTable = "temp"
        .SourceAlias = "temp"
        rowsAffected = .Run
    End With
    
    FieldNames = fieldArr.JoinArr
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "temp"
        .fields = FieldNames & "," & _
                  PrimaryKey & " As RecordImportID," & _
                  ELookup("tblModels", "RecordImportID = " & ModelID, "ModelID") & " AS ModelID"
        sqlStr = .sql
    End With
       
    ''Make sure that the chosen model is not yet present from the current file's own models
    ''or else do an update
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "UPDATE"
        .Source = tblName
        .SetStatement = BuildSetStatement(tblName, "temp", FieldNames & ",ModelID", True)
        .Joins.Add GenerateJoinObj(sqlStr, PrimaryKey, "temp", "RecordImportID")
        rowsAffected = .Run
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = sqlStr
        .AddFilter tblName & "." & PrimaryKey & " IS NULL"
        .fields = "temp.*"
        .SourceAlias = "temp"
        .Joins.Add GenerateJoinObj(tblName, "RecordImportID", , , "LEFT")
        sqlStr = .sql
    End With
    
     ''Insert the actual table model
     Set sqlObj = New clsSQL
     With sqlObj
         .SQLType = "INSERT"
         .Source = tblName
         .fields = concat(FieldNames, ",RecordImportID,ModelID")
         .InsertSQL = sqlStr
         .InsertFilterField = concat(FieldNames, ",RecordImportID,ModelID")
         rowsAffected = .Run
     End With
    
     
End Function

Public Function ImportModelClick(frm As Form)
    
    ''Get the recordset of the ModelID
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set rs = ReturnRecordset(frm("ModelID").rowSource)
    
    ''Get the fieldnames to be used for the insert statement
    Dim fieldArr As New clsArray, fieldArrItem, fld As DAO.Field, FieldNames
    For Each fld In rs.fields
        Select Case fld.Name
            Case "Timestamp", "CreatedBy", "RecordImportID", "ModelID":

            Case Else:
                fieldArr.Add fld.Name
        End Select
    Next fld

    FieldNames = fieldArr.JoinArr
    
    Dim SourceDatabaseFile, ModelID
    SourceDatabaseFile = frm("SourceDatabaseFile")
    ModelID = frm("ModelID")
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblModels IN " & EscapeString(SourceDatabaseFile)
        .AddFilter "IsSystemTable = 0 AND ModelID = " & ModelID
        .fields = concat(FieldNames, ",ModelID As RecordImportID")
        sqlStr = .sql
        Set rs = .Recordset
        
    End With
    
    ''Make sure that the chosen model is not yet present from the current file's own models
    ''or else do an update
    
    If Not isPresent("tblModels", concat("Model = ", EscapeString(rs.fields("Model")))) Then
        ''Insert the actual table model
        Set sqlObj = New clsSQL
        With sqlObj
            .SQLType = "INSERT"
            .Source = "tblModels"
            .fields = concat(FieldNames, ",RecordImportID")
            .InsertSQL = sqlStr
            .InsertFilterField = concat(FieldNames, ",RecordImportID")
            rowsAffected = .Run
        End With
        
    Else
    
        Set sqlObj = New clsSQL
        With sqlObj
            .SQLType = "UPDATE"
            .Source = "tblModels"
            .SetStatement = BuildSetStatement("tblModels", "temp", FieldNames, True)
            .Joins.Add GenerateJoinObj(sqlStr, "Model", "temp")
            rowsAffected = .Run
        End With

    End If
    
    ''Also import the related records
    ImportRelatedModelRecords ModelID, SourceDatabaseFile
    
    MsgBox "Model succesfully imported..."

    
End Function
