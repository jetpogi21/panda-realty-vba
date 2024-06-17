Attribute VB_Name = "ExcelLegacyConverter Mod"
Option Compare Database
Option Explicit

Public Function ExcelLegacyConverterCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function SelectExcelFile(frm As Form)

    Dim ExcelFilePath
    ExcelFilePath = PromptFile("Excel Files")
    
    If isFalse(ExcelFilePath) Then Exit Function
    
    frm("ExcelFilePath") = ExcelFilePath
    
End Function

Public Function SelectExcelDirectory(frm As Form)
    
    Dim OutputDirectory
    
    OutputDirectory = PromptDirectory("output.xlsx", "Excel Files")
    
    If isFalse(OutputDirectory) Then Exit Function
    
    frm("OutputDirectory") = OutputDirectory
    
End Function

Public Function ConvertExcelToOldFormat(frm As Form)

    Dim filePath: filePath = frm("ExcelFilePath")
    
    If ExitIfTrue(isFalse(filePath), "Please select an excel file to be converted.") Then Exit Function
    
    Dim OutputDirectory: OutputDirectory = frm("OutputDirectory")
    
    If ExitIfTrue(isFalse(OutputDirectory), "Please select a destination for the converted file") Then Exit Function
    
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim excelWorksheet As Object
    Dim startCell As Object
    Dim lastCol As Long
    Dim lastRow As Long
    Dim startRow As Long
    
    ' Check if Excel application is already running
    On Error Resume Next
    Set excelApp = GetObject(, "Excel.Application")
    On Error GoTo 0
    
    ' If Excel application is not running, create a new instance
    If excelApp Is Nothing Then
        Set excelApp = CreateObject("Excel.Application")
    End If
    
    excelApp.Visible = True
    
    Set excelWorkbook = excelApp.Workbooks.Open(filePath)
    
    ' Set the worksheet to be used
    Set excelWorksheet = excelWorkbook.Worksheets(1) ' Assuming you want to work with the first worksheet
    
    'excelWorksheet.Rows("1:2").Delete

    ' Find the first occurrence of "Property Photo" in the worksheet
    Set startCell = excelWorksheet.cells.find("Property Photo")
    
    If ExitIfTrue(startCell Is Nothing, "Invalid template file. A header with ""Property Photo"" can't be found.") Then Exit Function
    startRow = startCell.row
    
    ' Get the last row in column B
    lastRow = excelWorksheet.cells(excelWorksheet.Rows.Count, "B").End(-4162).row
    
    ' Loop through the rows from startRow to lastRow
    Dim currentRow
    For currentRow = lastRow To startRow Step -1
        ' Check if column B in the current row is blank and if it's not the startRow
        If currentRow <> startRow And excelWorksheet.cells(currentRow, "B").Value = "" Then
            ' Delete the entire row if column B is blank
            excelWorksheet.Rows(currentRow).Delete
        End If
    Next currentRow
    
    UpdateExcelCellValue excelWorksheet, "Property Photo", "Property Image"
    UpdateExcelCellValue excelWorksheet, "Property Type", "Property Type/Category"
    UpdateExcelCellValue excelWorksheet, "Land Size (m²)", "Land Size"
    UpdateExcelCellValue excelWorksheet, "Floor Size (m²)", "Building Area"
    UpdateExcelCellValue excelWorksheet, "Land Use", "Land Use 1"
    UpdateExcelCellValue excelWorksheet, "Open in RPData", "Open in RPP"
    
    RemoveExcelColumnByHeader excelWorksheet, "Year Built"
    RemoveExcelColumnByHeader excelWorksheet, "Parcel Details"
    RemoveExcelColumnByHeader excelWorksheet, "Owner Type"
    RemoveExcelColumnByHeader excelWorksheet, "RPD Valuation No"
    
    ''Primary plan/Lot Plan   RPD Valuation No    Valuation Type  Valuation Amount    Valuation Date
    AddExcelColumnToRightOfHeader excelWorksheet, "Agent", "Improve/Improvement"
    AddExcelColumnToRightOfHeader excelWorksheet, "Improve/Improvement", "Construction"
    AddExcelColumnToRightOfHeader excelWorksheet, "Land Use 1", "Zoning"
    AddExcelColumnToRightOfHeader excelWorksheet, "Development Zone", "Primary plan/Lot Plan"
    AddExcelColumnToRightOfHeader excelWorksheet, "Primary plan/Lot Plan", "RPD"
    AddExcelColumnToRightOfHeader excelWorksheet, "RPD", "Valuation No"
    AddExcelColumnToRightOfHeader excelWorksheet, "Valuation No", "Valuation Type"
    AddExcelColumnToRightOfHeader excelWorksheet, "Valuation Type", "Valuation Amount"
    AddExcelColumnToRightOfHeader excelWorksheet, "Valuation Amount", "Valuation Date"
    ''Owner 1 Address Owner 2 Name    Owner 2 Address Owner 3 Name    Owner 3 Address Vendor 1 Name   Vendor 1 Address    Vendor 2 Name   Vendor 2 Address    Vendor 3 Name   Vendor 3 Address    Relationship
    AddExcelColumnToRightOfHeader excelWorksheet, "Owner 1 Name", "Owner 1 Address"
    AddExcelColumnToRightOfHeader excelWorksheet, "Owner 2 Name", "Owner 2 Address"
    AddExcelColumnToRightOfHeader excelWorksheet, "Owner 3 Name", "Owner 3 Address"
    AddExcelColumnToRightOfHeader excelWorksheet, "Vendor 1 Name", "Vendor 1 Address"
    AddExcelColumnToRightOfHeader excelWorksheet, "Vendor 2 Name", "Vendor 2 Address"
    AddExcelColumnToRightOfHeader excelWorksheet, "Vendor 3 Name", "Vendor 3 Address"
    AddExcelColumnToRightOfHeader excelWorksheet, "Vendor 3 Address", "Relationship"
    
    'FindAndValidateData excelWorksheet, "Bed"
    'FindAndValidateData excelWorksheet, "Bath"
    'FindAndValidateData excelWorksheet, "Car"
    'FindAndValidateData excelWorksheet, "Land Size"
    'FindAndValidateData excelWorksheet, "Building Area"
    'FindAndValidateData excelWorksheet, "Valuation Amount"
    'FindAndValidateData excelWorksheet, "Valuation Date"
    
    ''get the row of startCell then use the column "B"
    Dim startCellRow: startCellRow = startCell
    ''move down starting from the row then delete any rows that has blank "B" + row
    
    ' Close the Excel workbook
    excelWorkbook.SaveAs OutputDirectory
    excelWorkbook.Close False
    
    excelApp.WindowState = -4137
    'excelApp.Top = 0
    '.Left = 0
    excelApp.Workbooks.Open OutputDirectory
    
    ''excelApp.Quit
    
    ' Clean up the objects
    Set excelWorksheet = Nothing
    Set excelWorkbook = Nothing
    Set excelApp = Nothing
    
End Function
