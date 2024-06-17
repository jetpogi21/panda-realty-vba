Attribute VB_Name = "ImportExcelFile Mod"
Option Compare Database
Option Explicit

Public Function ImportExcelFileCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Sub UpdateExcelCellValue(sht As Object, find As Variant, replace As Variant, Optional startRow = 5)
    Dim targetCell As Object
    'On Error GoTo ErrorHandler
    ' Replace Excel constants with their corresponding values
    ' xlValues = -4163, xlPart = 2, xlByRows = 1, xlNext = 1
    Set targetCell = sht.Rows(startRow).find(What:=find, LookIn:=-4163, LookAt:=2, SearchOrder:=1, SearchDirection:=1, MatchCase:=False)
    If Not targetCell Is Nothing Then
        targetCell.Value = replace
    End If
    Exit Sub
'ErrorHandler:
    'MsgBox "An error occurred. Please ensure you are using a valid Excel sheet."
End Sub

'Public Sub UpdateExcelCellValue(sht As Object, find, replace)
'
'    ' Find the cell with the value "Property Photo"
'    Dim targetCell As Object
'    Set targetCell = sht.Cells.find(find)
'
'    ' Check if the target cell was found
'    If Not targetCell Is Nothing Then
'        targetCell.Value = replace
'    End If
'
'End Sub

Public Sub RemoveExcelColumnByHeader(sht As Object, find)
    
    ' Find the cell with the value "Property Photo"
    Dim targetCell As Object
    Set targetCell = sht.cells.find(find)
    
    ' Check if the target cell was found
    If Not targetCell Is Nothing Then
        ' Get the column index of the target cell
        Dim targetColumnIndex As Integer
        targetColumnIndex = targetCell.Column
        
        ' Delete the entire column
        sht.Columns(targetColumnIndex).Delete
    End If
    
End Sub

Public Sub AddExcelColumnToRightOfHeader(sht As Object, find, newHeader)
    
    '' Check if the cell with the new header can be found. If it is then don't add this
    Dim newHeaderCell As Object
    Set newHeaderCell = sht.cells.find(What:=newHeader, LookIn:=-4163, LookAt:=1)
    
    If newHeader = "RPD" Then
        Debug.Print newHeader
    End If
    
    If Not newHeaderCell Is Nothing Then Exit Sub
    
    ' Find the cell with the specified header value
    Dim targetCell As Object
    Set targetCell = sht.cells.find(find)
    
    ' Check if the target cell was found
    If Not targetCell Is Nothing Then
        ' Get the column index of the target cell
        Dim targetColumnIndex As Integer
        targetColumnIndex = targetCell.Column
        
        ' Insert a new column to the right of the target column
        'sht.Columns.Add targetColumnIndex + 1
        ' Insert a new column to the right of the target column
        sht.Columns(targetColumnIndex + 1).Insert Shift:=2, CopyOrigin:=1
        
        ' Set the new header value in the cell next to the target cell
        sht.cells(targetCell.row, targetColumnIndex + 1).Value = newHeader
    End If
    
End Sub

Public Function GetColumnLetter(ByVal columnNumber As Long) As String
    Dim dividend As Long
    Dim columnLetter As String
    Dim modulo As Long
    
    dividend = columnNumber
    columnLetter = ""
    
    Do While dividend > 0
        modulo = (dividend - 1) Mod 26
        columnLetter = Chr(65 + modulo) & columnLetter
        dividend = (dividend - modulo) \ 26
    Loop
    
    GetColumnLetter = columnLetter
End Function

Public Sub FindAndValidateData(xlSheet As Object, findStr)

    Dim startRow As Long
    Dim endRow As Long
    Dim cellCol As Long
    Dim rowNum As Long

    ' Find the first occurrence of findStr in the range
    On Error Resume Next
    startRow = xlSheet.cells.find(findStr).row
    On Error GoTo 0
    
     ' Check if "Bed" was found
    If startRow = 0 Then
        Exit Sub
    End If
    
    ' Get the last used row in the worksheet
    endRow = xlSheet.cells(xlSheet.Rows.Count, 2).End(-4162).row ' -4162 represents xlUp
    ' Get the column of the "Bed" cell
    cellCol = xlSheet.cells.find(findStr).Column
    
   ' Validate each cell in the column between startRow and endRow
    For rowNum = startRow + 1 To endRow
        ' Check if the cell is numeric
        If Not IsNumeric(xlSheet.cells(rowNum, cellCol).Value) Then
            xlSheet.cells(rowNum, cellCol).Value = Null
        End If
    Next rowNum
    
    ' Get the column of the "Bed" cell
    If startRow > 0 Then
        cellCol = xlSheet.cells.find("Bed").Column
    Else
        cellCol = 0 ' Value not found
    End If
    
End Sub


Private Function ImportTo_tblPropertyListTemp(filePath) As Boolean
    
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim excelWorksheet As Object
    Dim startCell As Object
    Dim lastCol As Long
    Dim lastRow As Long
    Dim startRow As Long
    
    ' Create an instance of Excel and open the workbook
    
     ' Check if Excel application is already running
    On Error Resume Next
    Set excelApp = GetObject(, "Excel.Application")
    On Error GoTo 0
    
     ' If Excel application is not running, create a new instance
    If excelApp Is Nothing Then
        Set excelApp = CreateObject("Excel.Application")
    End If
    
    Set excelWorkbook = excelApp.Workbooks.Open(filePath)
    
    ' Set the worksheet to be used
    Set excelWorksheet = excelWorkbook.Worksheets(1) ' Assuming you want to work with the first worksheet
    
    ' Find the first occurrence of "Property Photo" in the worksheet
    Set startCell = excelWorksheet.cells.find("Property Photo")
    
    startRow = startCell.row
    
    UpdateExcelCellValue excelWorksheet, "Property Photo", "Property Image", startRow
    UpdateExcelCellValue excelWorksheet, "Property Type", "Property Type/Category", startRow
    UpdateExcelCellValue excelWorksheet, "Land Size (m²)", "Land Size", startRow
    UpdateExcelCellValue excelWorksheet, "Floor Size (m²)", "Building Area", startRow
    UpdateExcelCellValue excelWorksheet, "Land Use", "Land Use 1", startRow
    UpdateExcelCellValue excelWorksheet, "Open in RPData", "Open in RPP", startRow
    
    FindAndValidateData excelWorksheet, "Bed"
    FindAndValidateData excelWorksheet, "Bath"
    FindAndValidateData excelWorksheet, "Car"
    FindAndValidateData excelWorksheet, "Land Size"
    FindAndValidateData excelWorksheet, "Building Area"
    FindAndValidateData excelWorksheet, "Valuation Amount"
    FindAndValidateData excelWorksheet, "Valuation Date"
    
    ' Check if the word was found
    If Not startCell Is Nothing Then
        ' Get the row number of the startCell
        
        
        ' Get the last used column in the worksheet
        lastCol = excelWorksheet.cells(startRow, excelWorksheet.Columns.Count).End(-4159).Column ' -4159 represents xlToRight
        
        ' Get the last used row in the worksheet
        lastRow = excelWorksheet.cells(excelWorksheet.Rows.Count, 2).End(-4162).row ' -4162 represents xlUp
        
        ' Adjust the range to import by offsetting the startRow and lastCol by -1
        Dim importRange As String
        'importRange = excelWorksheet.Name & "!" & excelWorksheet.Cells(startRow, 1).Address & ":" & excelWorksheet.Cells(lastRow, lastCol - 1).Address
        importRange = "A" & startRow & ":" & GetColumnLetter(lastCol) & lastRow
        
        ' Use DoCmd.TransferSpreadsheet to import the range starting from column A and startRow, and ending at lastCol + lastRow
        DoCmd.SetWarnings False
        RunSQL "DELETE FROM tblPropertyListTemp"
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "tblPropertyListTemp", filePath, True, importRange
        DoCmd.SetWarnings True
        
        ' Close the Excel workbook
        excelWorkbook.Close False
        excelApp.Quit
        
        ' Clean up the objects
        Set excelWorksheet = Nothing
        Set excelWorkbook = Nothing
        Set excelApp = Nothing
        
        ImportTo_tblPropertyListTemp = True
    
    Else
        MsgBox "Invalid Data Structure."
        ' Close the Excel workbook
        excelWorkbook.Close False
        excelApp.Quit
        
        ' Clean up the objects
        Set excelWorksheet = Nothing
        Set excelWorkbook = Nothing
        Set excelApp = Nothing
        Exit Function
    End If
      
End Function

Public Function ImportExcelFile()
    
    Dim filePath
    filePath = GetFilePath
    If ExitIfTrue(filePath = "", "Please select a valid file..") Then Exit Function
    
    'Dim startT, endT
    'startT = Now
    
    If Not ImportTo_tblPropertyListTemp(filePath) Then
        Exit Function
    End If
    'endT = Now
    'MsgBox DateDiff("s", startT, endT)
    
    InsertDataToPropertyList
    
    Dim frm As Form
    If IsFormOpen("mainPropertyList") Then
        Set frm = Forms("mainPropertyList")
        frm("subform").Form.Requery
        frm("fltrSuburb").Requery
    End If
    
    If IsFormOpen("mainFavoriteProperties") Then
        Set frm = Forms("mainFavoriteProperties")
        frm("subform").Form.Requery
        frm("fltrSuburb").Requery
    End If
    
    
End Function

Public Function JoinOwners(Owner1Name, Owner2Name, Owner3Name) As String
     
    Dim OwnerArr As New clsArray
    If Not IsNull(Owner1Name) Then
        If Trim(Owner1Name) <> "-" Then OwnerArr.Add Owner1Name
    End If
    If Not IsNull(Owner2Name) Then
        If Trim(Owner2Name) <> "-" Then OwnerArr.Add Owner2Name
    End If
    If Not IsNull(Owner3Name) Then
        If Trim(Owner3Name) <> "-" Then OwnerArr.Add Owner3Name
    End If
    
    If OwnerArr.Count > 0 Then JoinOwners = OwnerArr.JoinArr(" ")
    
End Function

Public Function ParsedOwnerName(OwnerName) As String
    
    If isFalse(OwnerName) Then Exit Function
    Dim trimmedOwnerName: trimmedOwnerName = Trim(OwnerName)
    
    If trimmedOwnerName = "-" Then
        ParsedOwnerName = ""
    Else
        ParsedOwnerName = trimmedOwnerName
    End If
    
End Function


Private Function InsertDataToPropertyList()
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Dim PropertyListFields, propertyListFieldsPlain
    
    GetPropertyListTempFields PropertyListFields, propertyListFieldsPlain

    ''Get the Properties from qryPropertyListTemp in which the StreetAddress,Owner1Name,Owner2Name and Owner3Name isn't present from the list
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "qryPropertyListTemp"
        .AddFilter "PropertyListID IS NULL"
        .fields = PropertyListFields
        .Joins.Add GenerateJoinObj("tblPropertyList", "StreetAddress,CombinedOwner", , , "LEFT")
        sqlStr = .sql
        ''makeQuery .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj

        .SQLType = "INSERT"
        .Source = "tblPropertyList"
        .fields = propertyListFieldsPlain
        .InsertSQL = sqlStr
        .InsertFilterField = propertyListFieldsPlain
        ''.InsertUseAsPlain = True
        'makeQuery .sql
        rowsAffected = .Run

    End With

    InsertUniqueOwners
    InsertOwnersToProperties
    
    MsgBox "Data import finished..."

End Function

Public Function InsertOwnersToProperties()
    
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
        .Source = "unionPropertyOwners"
        .fields = "Cdbl(EntityID) As vEntityID,Cdbl(PropertyListID) As vPropertyListID"
        .Joins.Add GenerateJoinObj(sqlStr, "OwnerName", "tempSellers", "EntityName")
        .Joins.Add GenerateJoinObj("tblPropertyList", "StreetAddress,CombinedOwner")
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

Private Function InsertUniqueOwners()
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    ''unionPropertyOwners
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "MAKE"
        .Source = "unionPropertyOwners"
        .MakeTable = "tempSellers"
        .fields = "2 As EntityCategoryID,OwnerName As EntityName,Address,-1 As IsSeller"
        .OrderBy = "OwnerName,Address"
        .GroupBy = "OwnerName,Address"
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

Private Function GetPropertyListTempFields(PropertyListFields, propertyListFieldsPlain) As String
        
    'Dim exArr As New clsArray
    'exArr.arr = "PropertyImage"
    
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("qryPropertyListTemp")
        
    Dim fld As Field
    Dim fields() As String, FieldsPlain() As String
    Dim i As Integer
    
    For Each fld In rs.fields
        'If Not exArr.InArray(fld.Name) Then
        ReDim Preserve fields(i)
        ReDim Preserve FieldsPlain(i)
        fields(i) = "qryPropertyListTemp." & fld.Name
        FieldsPlain(i) = fld.Name
        i = i + 1
        'End If
    Next fld
    
    PropertyListFields = Join(fields, ",")
    propertyListFieldsPlain = Join(FieldsPlain, ",")

End Function

Private Function GetFilePath() As String

    Dim fd As Office.FileDialog
    Dim strFile As String
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
     
    With fd
     
        .filters.Clear
        .filters.Add "Excel Files", "*.xlsx?", 1
        .title = "Choose an Excel file"
        .AllowMultiSelect = False
     
        If .Show = True Then
     
            GetFilePath = .SelectedItems(1)
            
            Dim xl As Object
            Dim Xw As Object
            Dim sht As Object
            
            On Error Resume Next
            Set xl = GetObject(, "Excel.Application")
            On Error GoTo 0
            
             ' If Excel application is not running, create a new instance
            If xl Is Nothing Then
                Set xl = CreateObject("Excel.Application")
            End If
            
            ' Check if the workbook is already open
            For Each Xw In xl.Workbooks
                If Xw.fullName = GetFilePath Then
                    ' Close the workbook without saving changes
                    Xw.Close False
                    Exit For
                End If
            Next
            
            Set Xw = GetObject(GetFilePath)
            Set sht = Xw.Worksheets(1)
            
            xl.Visible = False
            Xw.Windows(1).Visible = True
            Xw.Activate
            
            sht.Range("L:L").NumberFormat = "@"
            sht.Range("R:R").NumberFormat = "@"
            
            Xw.Save
            Xw.Close SaveChanges:=True
            xl.Quit
            Set Xw = Nothing
            Set xl = Nothing
     
        End If
     
    End With
    
End Function
