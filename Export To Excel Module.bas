Attribute VB_Name = "Export To Excel Module"
Option Compare Database
Option Explicit

Private Function BuildFileName(fileName) As String

    ''Default directory is on backend directory + Property Reports Folder
    Dim xlDirectory
    xlDirectory = CurrentProject.Path & "\Files\"
    If Environ("computername") <> "DESKTOP-3G3V8GO" Then xlDirectory = "Z:\MY PANDA APP\Property Reports\"
    
    ''Check if the directory is existing, if not then create it.
    Dim strFolderExists
    strFolderExists = Dir(xlDirectory, vbDirectory)
    If strFolderExists = "" Then MkDir xlDirectory
    
    ''Build the search criteria
    ''Change File Name
    BuildFileName = xlDirectory & fileName & ".xlsx"
    
End Function

Public Function ExportToExcelSpreadsheet(frm As Form, Optional rsName = "tblPropertyList")
    
    Dim fileName
    CustomPropertyListFilter frm, False, fileName
    
    Dim savepath: savepath = GetSavePath(fileName)
    If ExitIfTrue(isFalse(savepath), "Please select a valid save path...") Then Exit Function
    
    Dim fltrStr As String
    fltrStr = frm("subform").Form.Filter
    
    Dim fieldsToExport
    fieldsToExport = GetfieldsToExport(frm)
    
    Dim sqlStr
    sqlStr = "SELECT " & fieldsToExport & " FROM " & rsName
    If Not isFalse(fltrStr) Then sqlStr = sqlStr & " WHERE " & fltrStr
    
    Dim qDef As DAO.QueryDef
    Set qDef = CurrentDb.QueryDefs("qryToExport")
    qDef.sql = sqlStr
    qDef.Close
    
    DoCmd.OutputTo acOutputQuery, "qryToExport", acFormatXLSX, savepath
    
    CreateObject("Shell.Application").Open savepath
    
End Function

Private Function GetfieldsToExport(frm As Form)
    
    Dim frm2
    Set frm2 = frm("subform").Form
    
    Dim exceptedArr As New clsArray, fieldArr As New clsArray
    exceptedArr.arr = "IsFavorite,txtOpenInRPP,ExcludeFromReport"
    Dim ctl As Control
    For Each ctl In frm2.Controls
        ''Exlude Hidden columns, In excepted array and begins with sum
        If Not ctl.ColumnHidden And Not exceptedArr.InArray(ctl.Name) And Not ctl.Name Like "Sum*" Then
            fieldArr.Add ctl.Name
        End If
    Next ctl
    
    GetfieldsToExport = fieldArr.JoinArr(",")
    
End Function

Private Function GetSavePath(fileName)
    
    fileName = BuildFileName(fileName)
    Dim Path: Path = fileName
    
    Dim fd As FileDialog
    Set fd = FileDialog(msoFileDialogSaveAs)
    With fd
        .title = "Choose a Location and Name of the File to Save This File"
        .buttonName = "Click to Save"
        .InitialFileName = Path
        If .Show <> 0 Then
            GetSavePath = .SelectedItems(1)
        End If
    End With
    
End Function

Public Function ExportToExcelFromRecordset(frm As Form)
    
    ''Filter the form as usual
    CustomPropertyListFilter frm
    
    ''Excel Title
    Dim title: title = frm.Caption
    
    ''Get the recordset of the subform chosen
    Dim rs As Recordset, sbFrm As Form
    Set sbFrm = frm("subform").Form
    Set rs = sbFrm.RecordsetClone
    
    ''If there is no record then exit
    If ExitIfTrue(rs.recordCount = 0, "There is no record to be imported..") Then Exit Function
    rs.MoveFirst
    
    ''Filter the field to only show the one shown on the subform
    Dim FieldNames As New clsArray, fieldCaptions As New clsArray, fieldTypes As New clsArray
    Dim fld As Field, fldName, ctl As Control
'    For Each fld In rs.Fields
'        fldName = fld.Name
'        If ControlExists(fldName, sbFrm) And fldName <> "Timestamp" And fldName <> "CreatedBy" Then
'            Set ctl = sbFrm.Controls(fldName)
'            fieldNames.Add fldName
'            fieldCaptions.Add ctl.Properties("DatasheetCaption")
'            fieldTypes.Add fld.Type
'        End If
'    Next fld
    
    For Each ctl In sbFrm.Controls
        If DoesPropertyExists(rs.fields, ctl.Name) Then
            fldName = ctl.Name
            If fldName <> "Timestamp" And fldName <> "CreatedBy" Then
                FieldNames.Add fldName
                fieldCaptions.Add ctl.Properties("DatasheetCaption")
                fieldTypes.Add rs.fields(fldName).Type
            End If
        End If
    Next ctl
    
    ''Open the excel and ready for export
    Dim xl As Object
    Dim sht As Object
    Dim xb As Object
    
    Set xl = CreateObject("Excel.Application")
    xl.Visible = True
    
    Set xb = xl.Workbooks.Add
    xb.Activate
    
    Set sht = xb.ActiveSheet
    
    sht.cells(1, 1) = title
    
    ''Format the columns
    Dim i As Integer, maxI As Integer, currentRow As Integer
    currentRow = 3
    maxI = FieldNames.Count - 1
    
    For i = 0 To maxI
        sht.cells(currentRow, i + 1) = fieldCaptions.arr(i)
        
        If fieldTypes.arr(i) = 8 Then
            sht.Columns(i + 1).NumberFormat = "m/d/yyyy"
        ElseIf fieldTypes.arr(i) = 7 Then
            sht.Columns(i + 1).NumberFormat = "#,##0.00"
        End If
        
    Next i
    
    currentRow = currentRow + 1
        
    Do Until rs.EOF
        For i = 0 To maxI
            sht.cells(currentRow, i + 1) = rs.fields(FieldNames.arr(i))
        Next i
        currentRow = currentRow + 1
        rs.MoveNext
    Loop
    
    For i = 0 To maxI
    
        sht.Columns(i + 1).AutoFit
    
    Next i

End Function


Public Function ExportToExcel(CustomReportID, frm As Form, dateCaption, dateField)

    Dim StartDate, endDate
    StartDate = frm("startDate")
    endDate = frm("endDate")
    
    If ExitIfTrue(IsNull(StartDate), "Please supply a start date..") Then Exit Function
    If ExitIfTrue(IsNull(endDate), "Please supply an end date..") Then Exit Function
    
    FilterCaption = dateCaption & " From " & StartDate & " and " & endDate

    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblCustomReports WHERE CustomReportID = " & CustomReportID)
    
    Dim reportName, ReportObjectName, FilterFormName, recordsetName, PreAppliedFilter, OrderBy, ReportOrientation, paperSize
    reportName = rs.fields("ReportName")
    recordsetName = rs.fields("RecordsetName")
    OrderBy = rs.fields("OrderBy")

    rs.Close
    
    If ExitIfTrue(ECount(recordsetName, dateField & " BETWEEN #" & StartDate & "# And #" & endDate & "#") = 0, "There is no record to show..") Then Exit Function
    
    Set rs = ReturnRecordset("SELECT * FROM tblCustomReportFields WHERE CustomReportID = " & CustomReportID & " And FieldOrder <> 0 ORDER BY FieldOrder ASC")
    
    Dim FieldNames As New clsArray, fieldCaptions As New clsArray, fieldTypes As New clsArray
    Dim CustomReportField, FieldTypeID, VerboseName
    
    Do Until rs.EOF
        
        CustomReportField = rs.fields("CustomReportField")
        FieldTypeID = rs.fields("FieldTypeID")
        VerboseName = rs.fields("VerboseName")
   
        FieldNames.Add CustomReportField
        fieldCaptions.Add VerboseName
        fieldTypes.Add FieldTypeID
        
        rs.MoveNext
        
    Loop
    
    rs.Close
    
    Dim xl As Object
    Dim sht As Object
    Dim xb As Object
    
    Set xl = CreateObject("Excel.Application")
    xl.Visible = True
    
    Set xb = xl.Workbooks.Add
    xb.Activate
    
    Set sht = xb.ActiveSheet
    
    sht.cells(1, 1) = reportName
    sht.cells(2, 1) = FilterCaption
    
    Dim sqlStr
    sqlStr = "SELECT * FROM " & recordsetName & " WHERE " & dateField & " BETWEEN #" & StartDate & "# And #" & endDate & "#"
    If Not IsNull(OrderBy) Then
        sqlStr = sqlStr & " ORDER BY " & OrderBy
    End If
    
    Set rs = ReturnRecordset(sqlStr)
    rs.MoveFirst
    
    Dim i As Integer, maxI As Integer, currentRow As Integer
    currentRow = 4
    maxI = FieldNames.Count - 1
    
    For i = 0 To maxI
        sht.cells(currentRow, i + 1) = fieldCaptions.arr(i)
        
        If fieldTypes.arr(i) = 8 Then
            sht.Columns(i + 1).NumberFormat = "m/d/yyyy"
        ElseIf fieldTypes.arr(i) = 7 Then
            sht.Columns(i + 1).NumberFormat = "#,##0.00"
        End If
        
    Next i
    currentRow = currentRow + 1
        
    Do Until rs.EOF
        For i = 0 To maxI
            sht.cells(currentRow, i + 1) = rs.fields(FieldNames.arr(i))
        Next i
        currentRow = currentRow + 1
        rs.MoveNext
    Loop
    
    For i = 0 To maxI
    
        sht.Columns(i + 1).AutoFit
    
    Next i
    
End Function
