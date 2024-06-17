Attribute VB_Name = "LeadImportData Mod"
Option Compare Database
Option Explicit

Public Function LeadImportDataCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Private Sub ExtractInformation(inputString, PropertyAddress, vName, Email, Phone, WouldLikeTo, Comments)
    
    ' Extract Property Address
    PropertyAddress = ExtractInfo(inputString, "Property id: (.+)")
    If Not isFalse(PropertyAddress) Then PropertyAddress = Trim(PropertyAddress)
    ' Extract Name
    vName = ExtractInfo(inputString, "Name: (.+)")
    
    ' Extract Email
    Email = ExtractInfo(inputString, "Email: (.+)")
    
    ' Extract Phone
    Phone = ExtractInfo(inputString, "Phone: (.+)")
    
    ' Extract "I would like to"
    WouldLikeTo = ExtractInfo(inputString, "I would like to: (.+)")
    
    ' Extract Comments
    Comments = ExtractInfo(inputString, "Comments: (.+)")
    
End Sub

Private Function ExtractInfo(inputString, regexPattern) As String
    Dim regexMatches As Object
    Dim regex As Object
    Dim match As Object
    
    ' Create a RegExp object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Set the pattern
    regex.pattern = regexPattern
    
    ' Execute the regex on the input string
    Set regexMatches = regex.Execute(inputString)
    
    ' Check if any matches were found
    If regexMatches.Count > 0 Then
        ' Extracted information
        ExtractInfo = regexMatches(0).SubMatches(0)
    Else
        ExtractInfo = ""
    End If
End Function

Private Sub UpdateExcelFileHeaders(xlApp As Object, wb As Object, ws As Object)
    
    ' Set values for cells A1 and B1
    ws.cells(1, 1).Value = "LeadImportData"
    ws.cells(1, 2).Value = "Processed"
    
    ' Save the changes
    wb.Save
    
    ' Close the workbook without saving changes (since we already saved above)
    wb.Close False
    
    ' Quit Excel application
    xlApp.Quit
    
    ' Release the objects
    Set ws = Nothing
    Set wb = Nothing
    Set xlApp = Nothing
End Sub

Public Function ImportExcelToTable(filePath, tableName) As Boolean
    On Error GoTo ErrorHandler
    
    ' Specify the import range (entire sheet)
    Dim importRange As String
    importRange = "Sheet1$" ' Change "Sheet1" to the actual sheet name if necessary
    
    ' Perform the import
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, tableName, filePath, True, importRange
    
    ' If we reach this point without an error, the import was successful
    ImportExcelToTable = True
    Exit Function
    
ErrorHandler:
    Debug.Print Err.Description
    ' An error occurred during the import process
    ImportExcelToTable = False
End Function

Public Function SelectExcelLeadTemplateFile(frm As Form)
    
    Dim xl As Object, wb As Object, sht As Object
    Dim filePath: filePath = GetFilePath(xl, wb, sht)
    
    If ExitIfTrue(isFalse(filePath), "There is no file selected.") Then Exit Function
    
    UpdateExcelFileHeaders xl, wb, sht
'    Dim maxRow, maxCol, curRow
'    maxRow = sht.UsedRange.Rows.Count
'    maxCol = sht.UsedRange.Columns.Count
'
    Dim PropertyAddress, vName, Email, Phone, WouldLikeTo, Comments
    
    RunSQL "DELETE FROM tblLeadImportData"
    
    ImportExcelToTable filePath, "tblLeadImportData"
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblLeadImportData WHERE Not Processed AND Not isFalse(LeadImportData)")
    
    If rs.EOF Then
        MsgBox "There's no record to import.", vbOKOnly
        Exit Function
    End If
    rs.MoveLast: rs.MoveFirst
    Dim recordCount: recordCount = rs.recordCount
    
    Dim cls_ProgressBar As New cls_ProgressBar
    cls_ProgressBar.ProgressBar_Show
    cls_ProgressBar.ProgressBar_ProgressOverlay True
    cls_ProgressBar.ProgressBar_Caption "Importing Lead Data"
    ''cls_ProgressBar.ProgressBar_Resize 3000, 6000
    cls_ProgressBar.ProgressBar_Message1_FontParam 1, "Calibri", 400, 10, vbBlack
    cls_ProgressBar.ProgressBar_ProgressValue_Align 2
    Dim i As Integer: i = 1
    Do Until rs.EOF
        Dim LeadImportData: LeadImportData = rs.fields("LeadImportData")
        Dim LeadImportDataID: LeadImportDataID = rs.fields("LeadImportDataID")
        ExtractInformation LeadImportData, PropertyAddress, vName, Email, Phone, WouldLikeTo, Comments
        cls_ProgressBar.ProgressBar_Progress (i / recordCount)
        cls_ProgressBar.ProgressBar_Message1 "Importing Lead Data"
        If Not isFalse(PropertyAddress) Then
            RunSQL "UPDATE tblLeadImportData SET " & _
                "PropertyAddress = " & Esc(PropertyAddress) & "," & _
                "[Name] = " & Esc(vName) & "," & _
                "[Email] = " & Esc(Email) & "," & _
                "[Phone] = " & Esc(Phone) & "," & _
                "[WouldLikeTo] = " & Esc(WouldLikeTo) & "," & _
                "[Comments] = " & Esc(Comments) & _
                " WHERE LeadImportDataID  = " & LeadImportDataID
        End If
        i = i + 1
        rs.MoveNext
    Loop
    
    cls_ProgressBar.ProgressBar_Hide
    Insert_tblLeadImportDataToDatabase
    
    SetColumnBTrue filePath
    
    MsgBox "Excel Lead finished importing.."
    
End Function

Sub SetColumnBTrue(filePath)
    Dim xs As Object
    Dim wb As Object
    Dim ws As Object
    Dim lastRow As Long
    
    ' Create a new Excel Application object
    Set xs = CreateObject("Excel.Application")
    
    ' Open the workbook
    Set wb = xs.Workbooks.Open(filePath)
    
    ' Assuming the first sheet is the one we want to modify
    Set ws = wb.Worksheets(1)
    
    ' Find the last used row in the worksheet
    lastRow = ws.UsedRange.Rows.Count
    
    ' Set all cells in column B from row   2 to the last used row to TRUE
    ws.Range("B2:B" & lastRow).Value = True
    
    ' Save the changes
    wb.Save
    
    ' Close the workbook without saving
    wb.Close False
    
    ' Quit the Excel Application
    xs.Quit
    
    ' Release the objects
    Set ws = Nothing
    Set wb = Nothing
    Set xs = Nothing
End Sub

'Public Function SelectExcelLeadTemplateFile(frm As Form)
'
'    Dim xl As Object, wb As Object, sht As Object
'    Dim filePath: filePath = GetFilePath(xl, wb, sht)
'
'    If ExitIfTrue(isFalse(filePath), "There is no file selected.") Then Exit Function
'
'    Dim maxRow, maxCol, curRow
'    maxRow = sht.UsedRange.Rows.Count
'    maxCol = sht.UsedRange.Columns.Count
'
'    Dim PropertyAddress, vName, Email, Phone, WouldLikeTo, Comments
'
'    RunSQL "DELETE FROM tblLeadImportData"
'    Dim cls_ProgressBar As New cls_ProgressBar
'
'    cls_ProgressBar.ProgressBar_Show
'    cls_ProgressBar.ProgressBar_ProgressOverlay True
'    cls_ProgressBar.ProgressBar_Caption "Importing Lead Data"
'    ''cls_ProgressBar.ProgressBar_Resize 3000, 6000
'    cls_ProgressBar.ProgressBar_Message1_FontParam 1, "Calibri", 400, 10, vbBlack
'    cls_ProgressBar.ProgressBar_ProgressValue_Align 2
'    Dim isImported As Boolean
'    For curRow = 1 To maxRow + 1
'
'        cls_ProgressBar.ProgressBar_Progress (curRow / maxRow)
'        cls_ProgressBar.ProgressBar_Message1 "Importing Lead Data"
'
'        Dim cellValue: cellValue = sht.Cells(curRow, 1).Value
'        cellValue = replace(cellValue, Chr(10), vbCrLf)
'        If isFalse(cellValue) Then GoTo NextRow:
'
'        isImported = (sht.Cells(curRow, 2).Value = True)
'        If isImported Then GoTo NextRow:
'        ExtractInformation cellValue, PropertyAddress, vName, Email, Phone, WouldLikeTo, Comments
'
'        If isFalse(PropertyAddress) Then GoTo NextRow:
'        RunSQL "INSERT INTO tblLeadImportData (LeadImportData,PropertyAddress, Name, Email, Phone, WouldLikeTo, Comments) VALUES (" & _
'            Esc(cellValue) & "," & _
'            Esc(PropertyAddress) & "," & _
'            Esc(vName) & "," & _
'            Esc(Email) & "," & _
'            Esc(Phone) & "," & _
'            Esc(WouldLikeTo) & "," & _
'            Esc(Comments) & _
'            ")"
'
'        sht.Cells(curRow, 2).Value = True
'NextRow:
'    Next curRow
'
'    wb.Save
'    wb.Close SaveChanges:=True
'    xl.Quit
'    Set wb = Nothing
'    Set xl = Nothing
'
'    cls_ProgressBar.ProgressBar_Hide
'
'    Insert_tblLeadImportDataToDatabase
'
'    MsgBox "Excel Lead finished importing.."
'
'End Function

Public Function Insert_tblLeadImportDataToDatabase()

    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblLeadImportData WHERE Not isFalse(PropertyAddress) ORDER BY LeadImportDataID")
    Dim EntityCategoryID: EntityCategoryID = 1
    Dim BuyerStatusID: BuyerStatusID = ELookup("tblBuyerStatus", "BuyerStatus = " & EscapeString("BUYER"), "BuyerStatusID")
    Dim EntityID
    
    Dim cls_ProgressBar As New cls_ProgressBar
    cls_ProgressBar.ProgressBar_Show
    cls_ProgressBar.ProgressBar_ProgressOverlay True
    cls_ProgressBar.ProgressBar_Caption "Inserting Data"
    ''cls_ProgressBar.ProgressBar_Resize 3000, 6000
    cls_ProgressBar.ProgressBar_Message1_FontParam 1, "Calibri", 400, 10, vbBlack
    cls_ProgressBar.ProgressBar_ProgressValue_Align 2
    
    Dim totalRecords: totalRecords = ECount("tblLeadImportData", "LeadImportDataID > 0")
    Dim i As Integer: i = 1
    Dim Notes As New clsArray
    Do Until rs.EOF
        Dim EntityName, PhoneNumber, EmailAddress, WouldLikeTo, Comments, PropertyAddress
        EntityName = rs.fields("Name")
        PhoneNumber = rs.fields("Phone")
        EmailAddress = rs.fields("Email")
        PropertyAddress = rs.fields("PropertyAddress")
        WouldLikeTo = rs.fields("WouldLikeTo")
        Comments = rs.fields("Comments")
        
        cls_ProgressBar.ProgressBar_Progress (i / totalRecords)
        cls_ProgressBar.ProgressBar_Message1 "Importing: " & Esc(EntityName)
        
        Dim filterStr As String: filterStr = "EntityName = " & Esc(EntityName) & _
            " AND PhoneNumber = " & Esc(PhoneNumber) & _
            " AND EmailAddress = " & Esc(EmailAddress) & _
            " AND EntityCategoryID = " & EntityCategoryID
        Dim PropertyListID
        Dim isEntityPresentFomPropertList
        If Not isPresent("tblEntities", filterStr) Then
            RunSQL "INSERT INTO tblEntities (EntityName,PhoneNumber,EmailAddress,EntityCategoryID,BuyerStatusID) VALUES (" & Esc(EntityName) & _
                "," & Esc(PhoneNumber) & "," & Esc(EmailAddress) & "," & EntityCategoryID & "," & BuyerStatusID & ")"
            ''Get the last inserted id
            EntityID = ELookup("tblEntities", "EntityID > 0", "EntityID", "EntityID DESC")
            
            PropertyListID = ELookup("tblPropertyList", "Trim(AdvertisementID) = " & Esc(PropertyAddress), "PropertyListID")
            
            If isFalse(PropertyListID) Then GoTo NextRecord:
            
            Set Notes = New clsArray
            If Not isFalse(WouldLikeTo) Then Notes.Add "Would Like To: " & WouldLikeTo
            If Not isFalse(Comments) Then Notes.Add "Comments: " & Comments
            Dim vNote As String
            If Notes.Count > 0 Then
                vNote = Notes.JoinArr(vbNewLine)
            End If
            
            vNote = replace(vNote, Chr(10), vbCrLf)
            
'            isEntityPresentFomPropertList = isPresent("tblPropertyEntities", "EntityID = " & EntityID & _
'                " AND PropertyListID = " & PropertyListID)
'
'            If Not isEntityPresentFomPropertList Then
            RunSQL "INSERT INTO tblPropertyEntities (EntityID,PropertyListID) VALUES (" & EntityID & "," & PropertyListID & ")"
            RunSQL "INSERT INTO tblEntityNotes (EntityID,PropertyListID,[Note]) VALUES (" & EntityID & "," & PropertyListID & _
                    "," & Esc(vNote) & ")"
'            End If
            RunSQL "UPDATE tblEntities SET LastViewedProperty = " & PropertyListID & " WHERE EntityID = " & EntityID
        Else
            EntityID = ELookup("tblEntities", filterStr, "EntityID", "EntityID DESC")
            
            PropertyListID = ELookup("tblPropertyList", "Trim(AdvertisementID) = " & Esc(PropertyAddress), "PropertyListID")
            
            If isFalse(PropertyListID) Then GoTo NextRecord:
            
            isEntityPresentFomPropertList = isPresent("tblPropertyEntities", "EntityID = " & EntityID & _
                " AND PropertyListID = " & PropertyListID)
            If Not isEntityPresentFomPropertList Then
                RunSQL "INSERT INTO tblPropertyEntities (EntityID,PropertyListID) VALUES (" & EntityID & "," & PropertyListID & ")"
            End If
            RunSQL "UPDATE tblEntities SET LastViewedProperty = " & PropertyListID & " WHERE EntityID = " & EntityID
        End If
NextRecord:
        i = i + 1
        rs.MoveNext
    Loop
    
    
    cls_ProgressBar.ProgressBar_Hide
    
    
End Function

Private Function GetFilePath(xl As Object, Xw As Object, sht As Object) As String

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
            
'            xw.Save
'            xw.Close SaveChanges:=True
'            xl.Quit
'            Set xw = Nothing
'            Set xl = Nothing
     
        End If
     
    End With
    
End Function

