Attribute VB_Name = "ExportEntityToExcel Mod"
Option Compare Database
Option Explicit

''Test ExportEntitiesToExcel(6)
Public Function ExportEntitiesToExcel(PropertyListID)
    
    If Environ("computername") = "DESKTOP-3G3V8GO" Then Exit Function
    
    If isFalse(PropertyListID) Then Exit Function
    
    ''Get all the qryPropertyEntityMembers related to PropertyListID OrderBy EntityCategoryName,EntityName,MemberName
    Dim sqlStr: sqlStr = GetUnionedSQL(PropertyListID)
    Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
    
    Dim EntityCategoryName: EntityCategoryName = ""
    
    Dim xl As Object, sht As Object, xb As Object, fileName As String
    
    Do Until rs.EOF
    
        If EntityCategoryName <> rs.fields("EntityCategoryName") Then
            
            If EntityCategoryName <> "" Then SaveExcelFile xb, sht, fileName
            
            EntityCategoryName = rs.fields("EntityCategoryName")
            Dim EntityDirectory: EntityDirectory = GetEntityDirectory(EntityCategoryName)
            If isFalse(EntityDirectory) Then Exit Function
            
            fileName = EntityDirectory & EntityCategoryName & ".xlsx"
            
            SetUpExcelFile xl, sht, xb, fileName
            
        End If
        
        ''Names,Phone Number,Email Address,Property Address,Website
        Dim i: i = sht.UsedRange.Rows.Count
        sht.Columns(3).NumberFormat = "@"
        If IsEmpty(sht.Range("A1")) Then
            i = 1
            If EntityCategoryName = "Contact" Then
                sht.cells(i, 1) = "Company name"
                sht.cells(i, 2) = "MEMBER"
                sht.cells(i, 3) = "Phone Number"
                sht.cells(i, 4) = "Email Address"
                sht.cells(i, 5) = "Website"
            Else
                sht.cells(i, 1) = "Company name"
                sht.cells(i, 2) = "MEMBER"
                sht.cells(i, 3) = "Phone Number"
                sht.cells(i, 4) = "Email Address"
                sht.cells(i, 5) = "Property Address"
                sht.cells(i, 6) = "Website"
            End If
        End If
        
        Dim EntityName: EntityName = rs.fields("EntityName")
        Dim MemberName: MemberName = rs.fields("MemberName")
        Dim MemberPhoneNumber: MemberPhoneNumber = rs.fields("MemberPhoneNumber")
        Dim MemberEmailAddress: MemberEmailAddress = rs.fields("MemberEmailAddress")
        Dim Website: Website = rs.fields("Website")
        Dim StreetAddress: StreetAddress = rs.fields("StreetAddress")
        
        Dim rowValue: rowValue = GetRow(MemberName, xl, sht, i)
        i = i + 1
        If rowValue <> 0 Then i = rowValue
        
        If EntityCategoryName = "Contact" Then
            sht.cells(i, 1) = EntityName
            sht.cells(i, 2) = MemberName
            sht.cells(i, 3) = MemberPhoneNumber
            sht.cells(i, 4) = MemberEmailAddress
            sht.cells(i, 5) = Website
        Else
            sht.cells(i, 1) = EntityName
            sht.cells(i, 2) = MemberName
            sht.cells(i, 3) = MemberPhoneNumber
            sht.cells(i, 4) = MemberEmailAddress
            sht.cells(i, 5) = StreetAddress
            sht.cells(i, 6) = Website
        End If
        
        ''Continue working on the current excel file
        rs.MoveNext
    Loop
    
    SaveExcelFile xb, sht, fileName
    
    MsgBox "Entities successfully exported.."
    
End Function

Private Function GetUnionedSQL(PropertyListID)

    Dim sqlStr: sqlStr = "SELECT EntityName,MemberName,MemberPhoneNumber,MemberEmailAddress,Website,StreetAddress,EntityCategoryName FROM qryPropertyEntityMembers WHERE PropertyListID = " & PropertyListID
    sqlStr = sqlStr & " ORDER BY EntityCategoryName,EntityName,MemberName UNION ALL "
    
    ''2nd Query
    sqlStr = sqlStr & "SELECT EntityName,EntityName As MemberName,PhoneNumber,EmailAddress,Website,StreetAddress,EntityCategoryName FROM qryPropertyEntities WHERE PropertyListID = " & PropertyListID & " AND EntityCategoryName = 'Seller'"
    sqlStr = sqlStr & " ORDER BY EntityCategoryName,EntityName"
    
    GetUnionedSQL = sqlStr
    
End Function

Private Function SaveExcelFile(xb, sht, fileName)

    sht.Columns("A:F").AutoFit
    If fileExists(fileName) Then
        xb.Save
    Else
        xb.SaveAs fileName
    End If
    xb.Close
    
End Function


Public Function SetUpExcelFile(xl, sht, xb, fileName) As Object
    
    If fileExists(fileName) Then
        If IsFileOpen(fileName) Then
            Set xl = GetObject(, "Excel.Application")
            Set xb = GetObject(fileName)
        Else
            Set xl = CreateObject("Excel.Application")
            Set xb = xl.Workbooks.Open(fileName)
        End If
    Else
        Set xl = CreateObject("Excel.Application")
        Set xb = xl.Workbooks.Add
    End If
    
    xl.Visible = False
    xb.Activate
    Set sht = xb.ActiveSheet
     
End Function



'Public Function ExportEntityToExcel(EntityMemberID, Optional EntityID = Null)
'
'    Dim rs As Recordset
'    If IsNull(EntityID) Then
'        Set rs = ReturnRecordset("SELECT * FROM qryPropertyEntityMembers WHERE EntityMemberID = " & EntityMemberID)
'    Else
'        Set rs = ReturnRecordset("SELECT * FROM qryPropertyEntities WHERE EntityID = " & EntityID)
'    End If
'
'    If rs.EOF Then Exit Function
'
'    Dim EntityCategoryName: EntityCategoryName = rs.Fields("EntityCategoryName")
'    Dim EntityDirectory, fileName
'    EntityDirectory = GetEntityDirectory(EntityCategoryName)
'    fileName = EntityDirectory & EntityCategoryName & ".xlsx"
'
'    If isFalse(EntityDirectory) Then Exit Function
'
'    Dim MemberName, MemberPhoneNumber, MemberEmailAddress, StreetAddress, Website
'    If IsNull(EntityID) Then
'        MemberName = rs.Fields("MemberName")
'        MemberPhoneNumber = rs.Fields("MemberPhoneNumber")
'        MemberEmailAddress = rs.Fields("MemberEmailAddress")
'        StreetAddress = rs.Fields("StreetAddress")
'        Website = rs.Fields("Website")
'    Else
'        MemberName = rs.Fields("EntityName")
'        MemberPhoneNumber = rs.Fields("PhoneNumber")
'        MemberEmailAddress = rs.Fields("EmailAddress")
'        StreetAddress = rs.Fields("StreetAddress")
'        Website = rs.Fields("Website")
'        EntityMemberID = EntityID
'    End If
'
'    SetUpExcelFile MemberName, MemberPhoneNumber, MemberEmailAddress, StreetAddress, Website, EntityMemberID, fileName, EntityCategoryName
'    ''EntityCategoryName for directory naming
'    ''Name, PhoneNumber, Email Address, Property Address,Website
'    ''MemberName,MemberPhoneNumber,MemberEmailAddress,StreetAddress,Website
'
'End Function

Private Function GetEntityDirectory(EntityCategoryName)
    
    Dim defaultDirectory
    GetEntityDirectory = CurrentProject.Path & "\Files\" & UCase(EntityCategoryName) & " FILES\"
    If Environ("computername") <> "DESKTOP-3G3V8GO" Then
        Dim OneDriveDir: OneDriveDir = "C:\Users\Owner\OneDrive\GOOGLE RECORDS 2\"
        If Not DirectoryExists(OneDriveDir) Then OneDriveDir = "C:\Users\appli\OneDrive\MY PANDA REALTY\PANDA APP\GOOGLE RECORDS 2\"
        GetEntityDirectory = OneDriveDir & UCase(EntityCategoryName & "S") & " LIST FOR GOOGLE\"
    End If
    
    Dim strFolderExists
    strFolderExists = Dir(GetEntityDirectory, vbDirectory)
    
    ''Create the directory if it doesn't exist
    If strFolderExists = "" Then
        MsgBox EscapeString(GetEntityDirectory) & " does not exist. Please create the directory first!"
        GetEntityDirectory = ""
    End If
    
End Function

Private Function GetRow(MemberName, xl, sht, i)

    On Error GoTo Err_Handler:
    GetRow = xl.WorksheetFunction.match(MemberName, sht.Columns(2), 0)
    Exit Function
    
Err_Handler:
    GetRow = 0
    Exit Function
    
End Function




'Set xl = CreateObject("Excel.Application")
'xl.Visible = True
'
'Set Xb = xl.Workbooks.Add
'Xb.Activate
'
'Set sht = Xb.ActiveSheet
