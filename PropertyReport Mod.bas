Attribute VB_Name = "PropertyReport Mod"
Option Compare Database
Option Explicit

Public Function PropertyReportCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function MainPropertyReports_OnLoad(frm As Form)

    DefaultMainFormLoad frm
    
    'frm("subform").Controls("txtOpenInRPP").columnWidth = 1500
    'frm("subform").Controls("PropertyStatusCode").ColumnHidden = False
    'frm("subform").Controls("PropertyStatusCode").columnWidth = 1500
    'frm("subform").Controls("PropertyStatusCode").ColumnOrder = 2
    
End Function

Public Function OpenPropertyReportForm(frm As Form)
    
    ''Open the PropertyReportForm
    
    ''This code will overwrite the fieldsToExportArr (should list the fields to be included to the form)
    Dim fieldsToExportArr As New clsArray
    GetfieldsToExport frm, fieldsToExportArr
    
    ''''Get cmdExportToExcel OnClick Property
    Dim exportToExcelClick, frmName
    frmName = frm.Name
    
    ''ExportToExcelSpreadsheet Attach this function to the cmdExportToExcel of mainPropertyReports
    exportToExcelClick = frm("cmdExportToExcel").OnClick
    exportToExcelClick = replace(exportToExcelClick, "[Form]", "[Forms].[" & frmName & "]")
    
    Dim showHideFieldsClick: showHideFieldsClick = frm("cmdShow/HideFields").OnClick
    showHideFieldsClick = replace(showHideFieldsClick, "[Form]", "[Forms].[" & frmName & "]")
    
    DoCmd.OpenForm "mainPropertyReports"
    
    ''This will hide all the fields not found on the main form
    HidePropertyReportFields fieldsToExportArr
    Forms("mainPropertyReports")("cmdExportToExcel").OnClick = exportToExcelClick
    Forms("mainPropertyReports")("cmdPreviewReport").OnClick = replace(exportToExcelClick, "ExportToExcelSpreadsheet", "PreviewPropertyReport")
    Forms("mainPropertyReports")("cmdShowHideFields").OnClick = showHideFieldsClick
    
    Dim printFunction
    printFunction = replace(exportToExcelClick, "ExportToExcelSpreadsheet", "PreviewPropertyReport")
    printFunction = replace(printFunction, ")", ",True)")
    Forms("mainPropertyReports")("cmdPrintReport").OnClick = printFunction
    
End Function

Public Function AddToPropertyReportTable(frm As Form, Optional IsFavorite As Boolean = False)
    
    Dim sqlStr
    sqlStr = "SELECT PropertyListID FROM qryPropertyList"
    
    Dim fltrArr As New clsArray, filterStr As String
    If frm("subform").Form.FilterOn Then
        fltrArr.Add frm("subform").Form.Filter
    End If
    
    fltrArr.Add "Not ExcludeFromReport"
    If IsFavorite Then fltrArr.Add "isFavorite"
    filterStr = fltrArr.JoinArr(" AND ")
    
    If Not isFalse(filterStr) Then sqlStr = sqlStr & " WHERE " & filterStr
    
    Dim recordCount
    If Not isFalse(filterStr) Then
        recordCount = ECount("qryPropertyList", filterStr)
    Else
        recordCount = ECount("qryPropertyList", "PropertyListID > 0")
    End If
     
    If ExitIfTrue(recordCount = 0, "There is no record to import.") Then Exit Function
    
    InsertToPropertyReport sqlStr
    
    ''Open the PropertyReportForm
    Dim fieldsToExportArr As New clsArray
    GetfieldsToExport frm, fieldsToExportArr
    
    ''''Get cmdExportToExcel OnClick Property
    Dim exportToExcelClick, frmName
    frmName = frm.Name
    exportToExcelClick = frm("cmdExportToExcel").OnClick
    exportToExcelClick = replace(exportToExcelClick, "[Form]", "[Forms].[" & frmName & "]")
    
    DoCmd.OpenForm "mainPropertyReports"
    HidePropertyReportFields fieldsToExportArr
    Forms("mainPropertyReports")("subform").Form.Requery
    Forms("mainPropertyReports")("cmdExportToExcel").OnClick = exportToExcelClick
    
    ''PreviewPropertyReport
    Forms("mainPropertyReports")("cmdPreviewReport").OnClick = replace(exportToExcelClick, "ExportToExcelSpreadsheet", "PreviewPropertyReport")
    
    Dim printFunction
    printFunction = replace(exportToExcelClick, "ExportToExcelSpreadsheet", "PreviewPropertyReport")
    printFunction = replace(printFunction, ")", ",True)")
    Forms("mainPropertyReports")("cmdPrintReport").OnClick = printFunction
    
End Function

Public Function Open_mainPropertyReportFields(frm As Form)

    Set frm = frm("subform").Form
    
    ''Delete the records from tblPropertyReportFields
    RunSQL "DELETE FROM tblPropertyReportFields"
    
    ''OpenInRPP,Favorite,AppraisedAmount,CompleteAddress
    Dim exceptedArr As New clsArray: exceptedArr.arr = "txtOpenInRPP,SumAppraisedAmount,txtPropertyStatusID,PropertyImage"
    
    Dim ctl As Control, ctlName, ctlCaption, IsVisible, ColumnOrder
    
    For Each ctl In frm.Controls
        ctlName = ctl.Name
        ctlCaption = ctl.Properties("DatasheetCaption")
        ColumnOrder = ctl.ColumnOrder
       
        If isFalse(ctlCaption) Then
            ctlCaption = ctlName
        End If
        
        IsVisible = Not ctl.ColumnHidden
        
        If Not ctlName Like "Sum*" And Not exceptedArr.InArray(ctlName) Then
            RunSQL "INSERT INTO tblPropertyReportFields (PropertyReportField,DatabaseField,Shown,FieldOrder) " & _
                " VALUES (" & EscapeString(ctlCaption) & "," & EscapeString(ctlName) & "," & IsVisible & "," & ColumnOrder & ")"
            Debug.Print ctlName, ctl.ColumnHidden, ctlCaption, ctl.ColumnOrder
        End If
          
    Next ctl
    
    
    DoCmd.OpenForm "mainPropertyReportFields"
    
End Function

Private Function HidePropertyReportFields(fieldsToExportArr As clsArray)
    
    If ExitIfTrue(Not IsFormOpen("mainPropertyReports"), "Property Report List isn't visible.") Then Exit Function
    Dim frm As Form: Set frm = Forms("mainPropertyReports")("subform").Form
    
    Dim ctl As Control
    For Each ctl In frm.Controls
        
        HideColumn ctl, fieldsToExportArr
        
    Next ctl
    
    ''Reorder Columns
    Dim item, i As Integer
    i = 0
    For Each item In fieldsToExportArr.arr
        i = i + 1
        SetColumnOrder frm, item, i
        ''frm(item).ColumnOrder = i
    Next item
    
End Function

Private Function SetColumnOrder(frm As Form, item, i)

    On Error Resume Next
    frm(item).ColumnOrder = i
    
End Function

Private Function HideColumn(ctl As Control, fieldsToExportArr As clsArray)
    
    On Error GoTo ErrHandler:
    
        ctl.ColumnHidden = Not fieldsToExportArr.InArray(ctl.Name)
    Exit Function
    
ErrHandler:
    
    Exit Function
    
End Function

''additionalException will remove the fields stated as default
Private Function GetfieldsToExport(frm As Form, Optional fieldsToExportArr As clsArray, Optional additionalException As String)
    
    ''Make a reference to the subform of the form
    Dim frm2
    Set frm2 = frm("subform").Form
    
    ''This array will remove the fields from the returned list
    Dim exceptedArr As New clsArray, fieldArr As New clsArray
    Dim exceptedStr: exceptedStr = "IsFavorite,txtPropertyAltlink,ExcludeFromReport"
    If additionalException <> "" Then exceptedStr = exceptedStr & "," & additionalException
    ''exceptedArr.arr = "IsFavorite,txtOpenInRPP,txtPropertyAltlink,ExcludeFromReport"
    exceptedArr.arr = exceptedStr
    ''exceptedArr.arr = "IsFavorite,txtOpenInRPP,ExcludeFromReport"
    Dim ctl As Control
    For Each ctl In frm2.Controls
        ''Exlude Hidden columns, In excepted array and begins with sum
        If Not ctl.ColumnHidden And Not exceptedArr.InArray(ctl.Name) And Not ctl.Name Like "Sum*" Then
            fieldArr.Add ctl.Name
        End If
    Next ctl
    
    If Not fieldsToExportArr Is Nothing Then
        Set fieldsToExportArr = fieldArr
    End If
    
    ''Get the field list as string of the subform
    GetfieldsToExport = fieldArr.JoinArr(",")
    
End Function


Private Function InsertToPropertyReport(sqlStr)

    Dim sqlObj As clsSQL, joinObj As clsJoin, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "INSERT"
        .Source = "tblPropertyReports"
        .fields = "PropertyListID"
        .InsertSQL = sqlStr
        .InsertUseAsPlain = True
        rowsAffected = .Run
    End With
    
End Function
'
Public Function PreviewPropertyReport(frm As Form, Optional rsName = "", Optional PrintOut As Boolean = False)

    Dim fileName
    CustomPropertyListFilter frm, False, fileName
    
    Dim savepath: savepath = GetSavePath
    
    If DoesFileExist(GetSavePath) Then
        If ExitIfTrue(IsFileOpen(savepath), "Please close the " & EscapeString("preview.xlsx") & " file first...") Then Exit Function
    End If
    
    Dim fieldsToExport, fieldsToExportArr As New clsArray
    fieldsToExport = GetfieldsToExport(frm, fieldsToExportArr, "txtOpenInRPP")
    Debug.Print fieldsToExport
    
    Dim sqlStr
    sqlStr = "SELECT " & fieldsToExport & " FROM qryPropertyReports"
    
    Dim qDef As DAO.QueryDef
    Set qDef = CurrentDb.QueryDefs("qryToExport")
    qDef.sql = sqlStr
    qDef.Close
    
    DoCmd.OutputTo acOutputQuery, "qryToExport", acFormatXLSX, savepath
    
    ''Process the excel file here
    ProcessExcelFile savepath, PrintOut
    
End Function

Private Function ProcessExcelFile(savepath, Optional PrintOut As Boolean)

    ''Change the font size to Calibri 8
    ''Fix the number format (mimic the one shown on the subform)
    Dim xl As Object, xb As Object, sht As Object
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.ScreenUpdating = True
    
    Set xb = xl.Workbooks.Open(savepath)
    xb.Activate
    Set sht = xb.ActiveSheet

    Dim maxRow, maxCol
    maxRow = sht.UsedRange.Rows.Count
    maxCol = sht.UsedRange.Columns.Count

    sht.Range(sht.cells(1, 1), sht.cells(maxRow, maxCol)).Font.size = 8
    sht.Range(sht.cells(1, 1), sht.cells(maxRow, maxCol)).WrapText = False
    sht.Range(sht.cells(1, 1), sht.cells(maxRow, maxCol)).EntireColumn.AutoFit
    sht.Range(sht.cells(1, 1), sht.cells(maxRow, maxCol)).EntireRow.AutoFit
    
    'xl.Visible = True
    xl.PrintCommunication = False
    With sht.PageSetup
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Orientation = 2
        .paperSize = 5
    End With
    
    If PrintOut Then
        Dim Response
        Response = MsgBox("This will send the file to your default printer. Do you want to proceed?", vbYesNo)
        If Response = vbYes Then
            xl.PrintCommunication = True
            xl.Visible = True
            xl.CommandBars.ExecuteMso "PrintPreviewAndPrint"
        End If
        ''If response = vbYes Then sht.PrintOut
'        xb.Close SaveChanges:=False
'        xl.Quit
    Else
        xl.PrintCommunication = True
        xl.Visible = True
        sht.PrintPreview
    End If
    
End Function

Private Function GetSavePath()

    ''Default directory is on backend directory + Property Reports Folder
    Dim xlDirectory
    xlDirectory = CurrentProject.Path & "\Files\"
    If Environ("computername") <> "DESKTOP-3G3V8GO" Then xlDirectory = "Z:\MY PANDA APP\Property Reports\"
    
    ''Check if the directory is existing, if not then create it.
    Dim strFolderExists
    strFolderExists = Dir(xlDirectory, vbDirectory)
    If strFolderExists = "" Then MkDir xlDirectory
    
    ''Build the search criteria
    ''Change File Name - Computer Specific Name so that multiple excels can be opened by different PCS.
    Dim computerName
    computerName = Environ("ComputerName")
    GetSavePath = xlDirectory & "preview" & computerName & ".xlsx"
    
End Function


Public Function ClearListPropertyReport(frm As Form)

    RunSQL "DELETE FROM tblPropertyReports"
    
    frm("subform").Form.Requery
    
End Function
