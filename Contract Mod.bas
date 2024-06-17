Attribute VB_Name = "Contract Mod"
Option Compare Database
Option Explicit

Public Function ContractCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function Save_rptContractsAsPDF(frm As Form)

    Dim PropertyDirectory: PropertyDirectory = GetPropertyDirectory(frm, "PropertyListID")
    Dim fileName: fileName = "Sales Contract.pdf"
    Dim filePath: filePath = PropertyDirectory & fileName
    
    Dim ContractID: ContractID = frm("ContractID")
    
    DoCmd.OpenReport "rptContracts", acViewDesign, , , acHidden
    Dim rpt As Report: Set rpt = Reports("rptContracts")
    
    SetReportMargin rpt
    DoCmd.Close acReport, "rptContacts", acSaveYes
    
    DoCmd.OpenReport "rptContracts", acViewReport, , "ContractID = " & ContractID, acHidden
    
    
    SaveReportAsPDF "rptContracts", filePath
    
    MsgBox "Sales Contract generated at: " & Esc(filePath)
    
    CreateObject("Shell.Application").Open filePath
    
End Function

Public Function Save_rptForm6AsPDF(frm As Form)

    Dim PropertyDirectory: PropertyDirectory = GetPropertyDirectory(frm, "PropertyListID")
    Dim fileName: fileName = "Form-6.pdf"
    Dim filePath: filePath = PropertyDirectory & fileName
    
    Dim Form6ID: Form6ID = frm("Form6ID")
    
    DoCmd.OpenReport "rptForm6", acViewDesign, , , acHidden
    Dim rpt As Report: Set rpt = Reports("rptForm6")
    
    SetReportMargin rpt
    DoCmd.Close acReport, "rptForm6", acSaveYes
    
    DoCmd.OpenReport "rptForm6", acViewReport, , "Form6ID = " & Form6ID, acHidden
    
    
    SaveReportAsPDF "rptForm6", filePath
    
    MsgBox "Form 6 generated at: " & Esc(filePath)
    
    CreateObject("Shell.Application").Open filePath
    
End Function


Private Function UpdateBaseOnTimeline(EventName, FieldToUpdate, ContractID, PropertyListID)
    
    Dim Value: Value = ELookup("qryEventTimelines", "PropertyListID = " & PropertyListID & _
        " AND Description = " & Esc(EventName), "EventTimelineDate")
        
'    Dim SettlementDate: SettlementDate = ELookup("qryEventTimelines", "PropertyListID = " & PropertyListID & _
'        " AND Description = ""SETTLEMENT DATE""", "EventTimelineDate")
        
    If Not isFalse(Value) Then
        If FieldToUpdate = "SettlementDate" Then
            Value = Esc(Format(Value, "DD/MM/YYYY"))
        Else
            Value = "#" & DateValue(Value) & "#"
        End If
        RunSQL "UPDATE tblContracts SET " & FieldToUpdate & " = " & Value & " WHERE ContractID = " & ContractID
    Else
        RunSQL "UPDATE tblContracts SET " & FieldToUpdate & " = Null WHERE ContractID = " & ContractID
    End If
    
End Function

Public Function CreateSaleContract(frm As Form)
    
    ''RunSQL "DELETE FROM tblContracts"
    
    Dim PropertyListID: PropertyListID = frm("PropertyListID")
    If isFalse(PropertyListID) Then Exit Function
    
    Dim txtSuccessfulBuyer
    Dim SuccessfulBuyerName, SuccessfulBuyerID, SuccessfulDate, ContractDate, txtSuccessfulOffer
    SuccessfulBuyerID = frm("txtSuccessfulBuyer").Column(0)
    SuccessfulBuyerName = frm("txtSuccessfulBuyer").Column(1)
    SuccessfulDate = GetSuccessfulOfferDate(PropertyListID)
    ContractDate = frm("ContractDate")
    txtSuccessfulOffer = frm("txtSuccessfulOffer")

    Dim ContractID: ContractID = ELookup("tblContracts", "PropertyListID = " & PropertyListID, "ContractID", "ContractID DESC")
    If isFalse(ContractID) Then
        RunSQL "INSERT INTO tblContracts (PropertyListID) VALUES (" & PropertyListID & ")"
        ContractID = ELookup("tblContracts", "PropertyListID = " & PropertyListID, "ContractID", "ContractID DESC")
    End If
    
    ''RunSQL "UPDATE tblContracts SET TenantID = 0 WHERE ContractID = " & ContractID & " WHERE Tenand
    ''RunSQL "UPDATE tblContracts SET PropertyMangerID = 0 WHERE ContractID = " & ContractID
    UpdateBaseOnTimeline "SETTLEMENT DATE", "SettlementDate", ContractID, PropertyListID
    UpdateBaseOnTimeline "Finance Due Date", "FinanceDate", ContractID, PropertyListID
    UpdateBaseOnTimeline "BPI Due", "InspectionDate", ContractID, PropertyListID
    UpdateBaseOnTimeline "Balance Deposit Due Date", "BalancePayableOn", ContractID, PropertyListID
    
    If Not isFalse(ContractDate) Then
        SuccessfulDate = DateAdd("d", 1, ContractDate)
        RunSQL "UPDATE tblContracts SET InitialDepositPayableOn = #" & DateValue(SuccessfulDate) & "# WHERE ContractID = " & ContractID
    End If
    
    If Not isFalse(ContractDate) Then
        ContractDate = DateAdd("d", 14, ContractDate)
        RunSQL "UPDATE tblContracts SET BalancePayableOn = #" & DateValue(ContractDate) & "# WHERE ContractID = " & ContractID & " AND BalancePayableOn IS NULL"
    End If
    
    If Not isFalse(txtSuccessfulOffer) Then
        RunSQL "UPDATE tblContracts SET PurchasePrice = " & txtSuccessfulOffer & "  WHERE ContractID = " & ContractID
    End If
    
    ''Update the SellerID and Seller2ID
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT TOP 2 * FROM qryPropertyEntities WHERE PropertyListID = " & PropertyListID & _
        " AND EntityCategoryName = ""Seller""")
    Dim i As Integer: i = 1
    Do Until rs.EOF
        If i = 1 Then
            RunSQL "UPDATE tblContracts SET SellerID = " & rs.fields("EntityID") & " WHERE ContractID = " & ContractID
        Else
            RunSQL "UPDATE tblContracts SET Seller2ID = " & rs.fields("EntityID") & " WHERE ContractID = " & ContractID
        End If
        i = i + 1
        rs.MoveNext
    Loop
    
    ''Update the BuyerID and Buyer2ID dependning on the member
    If Not isFalse(SuccessfulBuyerID) Then
        Set rs = ReturnRecordset("SELECT TOP 2 * FROM qryPropertyEntityMembers WHERE MemberName <> " & Esc(SuccessfulBuyerName) & _
            " AND PropertyListID = " & PropertyListID & " AND EntityID = " & SuccessfulBuyerID)
        
        If rs.EOF Then
            Set rs = ReturnRecordset("SELECT Top 1 * FROM qryPropertyEntityMembers WHERE PropertyListID = " & PropertyListID & " AND EntityID = " & SuccessfulBuyerID)
            If Not rs.EOF Then
                RunSQL "UPDATE tblContracts SET BuyerID = " & rs.fields("EntityMemberID") & " WHERE ContractID = " & ContractID
            End If
        Else
            i = 1
            Do Until rs.EOF
                If i = 1 Then
                    RunSQL "UPDATE tblContracts SET BuyerID = " & rs.fields("EntityMemberID") & " WHERE ContractID = " & ContractID
                Else
                    RunSQL "UPDATE tblContracts SET Buyer2ID = " & rs.fields("EntityMemberID") & " WHERE ContractID = " & ContractID
                End If
                i = i + 1
                rs.MoveNext
            Loop
        End If
    End If
    
    ''Update the PropertyMangerID
    Set rs = ReturnRecordset("SELECT TOP 1 * FROM qryPropertyEntities WHERE PropertyListID = " & PropertyListID & _
        " AND EntityCategoryName = ""Contact"" AND ContactCategoryName = ""Management"" AND Association = ""Seller""")
    Dim EntityID, EntityName, EntityMemberID
    Do Until rs.EOF
        EntityID = rs.fields("EntityID")
        EntityName = rs.fields("EntityName")
        
        EntityMemberID = ELookup("qryPropertyEntityMembers", "EntityID = " & EntityID, "EntityMemberID", "[Timestamp] DESC")
        RunSQL "UPDATE tblContracts SET PropertyMangerID = " & EntityMemberID & _
            ",AgencyName = " & Esc(EntityName) & " WHERE ContractID = " & ContractID
        rs.MoveNext
    Loop
    
    ''Update the BuyerAgentID
    Dim BuyerAgentID: BuyerAgentID = ELookup("qryPropertyEntities", "PropertyListID = " & PropertyListID & _
        " AND EntityCategoryName = ""Contact"" AND ContactCategoryName = ""Buyers Agent""", "EntityID")
    If Not isFalse(BuyerAgentID) Then
        RunSQL "UPDATE tblContracts SET BuyerAgentID = " & BuyerAgentID & " WHERE ContractID = " & ContractID
    End If
        
        
    ''BuyerSolicitorID,SellerSolicitorID
    Dim Association: Association = "Seller"
    Dim SellerSolicitorID:  SellerSolicitorID = ELookup("qryPropertyEntities", "EntityCategoryName = ""Contact"" AND " & _
        "ContactCategoryName = ""Solicitor"" AND PropertyListID = " & PropertyListID & " AND Association = " & Esc(Association), "EntityID", "PropertyEntityID")
    If Not isFalse(SellerSolicitorID) Then
        RunSQL "UPDATE tblContracts SET SellerSolicitorID = " & SellerSolicitorID & " WHERE ContractID = " & ContractID
    End If
    
    Association = "Buyer"
    Dim BuyerSolicitorID:  BuyerSolicitorID = ELookup("qryPropertyEntities", "EntityCategoryName = ""Contact"" AND " & _
        "ContactCategoryName = ""Solicitor"" AND PropertyListID = " & PropertyListID & " AND Association = " & Esc(Association), "EntityID", "PropertyEntityID")
    If Not isFalse(BuyerSolicitorID) Then
        RunSQL "UPDATE tblContracts SET BuyerSolicitorID = " & BuyerSolicitorID & " WHERE ContractID = " & ContractID
    End If
    
    DoCmd.OpenForm "frmContracts", , , "ContractID = " & ContractID
       
End Function

Public Function CreateForm6(frm As Form)
    
    ''RunSQL "DELETE FROM tblContracts"
    
    Dim PropertyListID: PropertyListID = frm("PropertyListID")
    If isFalse(PropertyListID) Then Exit Function


    Dim Form6ID: Form6ID = ELookup("tblForm6", "PropertyListID = " & PropertyListID, "Form6ID", "Form6ID DESC")
    
    If isFalse(Form6ID) Then
        RunSQL "INSERT INTO tblForm6 (PropertyListID) VALUES (" & PropertyListID & ")"
        Form6ID = ELookup("tblForm6", "PropertyListID = " & PropertyListID, "Form6ID", "Form6ID DESC")
    End If
    
    ''Update the SellerID and Seller2ID
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT TOP 2 * FROM qryPropertyEntities WHERE PropertyListID = " & PropertyListID & _
        " AND EntityCategoryName = ""Seller""")
    Dim i As Integer: i = 1
    Do Until rs.EOF
        If i = 1 Then
            RunSQL "UPDATE tblForm6 SET SellerID = " & rs.fields("EntityID") & " WHERE Form6ID = " & Form6ID
        Else
            RunSQL "UPDATE tblForm6 SET Seller2ID = " & rs.fields("EntityID") & " WHERE Form6ID = " & Form6ID
        End If
        i = i + 1
        rs.MoveNext
    Loop
    
    DoCmd.OpenForm "frmForm6", , , "Form6ID = " & Form6ID
       
End Function


Public Sub MakeReportCopies(rptName As String, copies As Integer)
    Dim i As Integer
    Dim newRptName As String

    For i = 1 To copies
        ' Generate a new report name by appending the copy number
        newRptName = rptName & i
        ' Copy the report
        DoCmd.CopyObject , newRptName, acReport, rptName
    Next i
End Sub


Public Sub CopyFormWithImage(number, FormName, croppedFilesDirectory)
    
    ''C:\Users\User\Desktop\Client Files\Richard F\5-9-2024\Document files\Form 6\cropped\
    Dim originalFormName: originalFormName = "frm" & FormName & "Page"
    ' Define the new form name with the number appended
    Dim newFormName As String
    newFormName = originalFormName & number
    
    ' Check if a form with the new name already exists and delete it if it does
    If FormExists(newFormName) Then
        DoCmd.DeleteObject acForm, newFormName
    End If
    
    ' Copy the original form to the new form name
    DoCmd.CopyObject , newFormName, acForm, originalFormName
    
    ' Open the new form in design view
    DoCmd.OpenForm newFormName, acDesign
    
    ' Set the picture path for the imgBackground control
    Dim imgBackground As Control
    Set imgBackground = Forms(newFormName).Controls("imgBackground")
    imgBackground.Picture = croppedFilesDirectory & FormName & "_page-" & Format(number, "0000") & "_cropped.jpg"
    
    ' Save and close the new form
    DoCmd.Close acForm, newFormName, acSaveYes
    
End Sub

Public Sub CompareControlPositions(page, ctlName)
    
    Dim frmName: frmName = "frmContractsPage" & page
    Dim rptName: rptName = "rptContractsPage" & page
    DoCmd.OpenForm frmName, acDesign
    DoCmd.OpenReport rptName, acDesign
    
    Dim frm As Form, rpt As Report
    Set frm = Forms(frmName)
    Set rpt = Reports(rptName)
    Dim frmControl As Control, rptControl As Control
    
    Set frmControl = frm(ctlName)
    Set rptControl = rpt(ctlName)
    
    Dim topDiff: topDiff = frmControl.top - rptControl.top
    Dim leftDiff: leftDiff = frmControl.left - rptControl.left
    Dim widthDiff: widthDiff = frmControl.width - rptControl.width
    
    Debug.Print "Top: " & frmControl.top & " - " & rptControl.top & " = " & topDiff & "(" & Divide(topDiff, frmControl.top) & ")"
    Debug.Print "Left: " & frmControl.left & " - " & rptControl.left & " = " & leftDiff & "(" & Divide(leftDiff, frmControl.left) & ")"
    Debug.Print "Width: " & frmControl.width & " - " & rptControl.width & " = " & widthDiff & "(" & Divide(widthDiff, frmControl.width) & ")"
    
End Sub

Public Sub SetUp_rptContracts()
    
    Dim rptName: rptName = "rptContracts"
    DoCmd.OpenReport rptName, acViewDesign
    
    Dim rpt As Report: Set rpt = Reports(rptName)
    
    rpt.RecordSource = "tblContracts"
    
    ''Create subreports 1-14
    Dim i As Integer, ctl As Control
    For i = 1 To 14
        Dim subreportName: subreportName = "subContractsPage" & i
        If ControlExists(subreportName, rpt) Then
            Set ctl = rpt(subreportName)
        Else
            Set ctl = CreateReportControl(rptName, acSubform, acDetail, , , 0, 0 + (300 * (i - 1)), InchToTwip(8), 300)
            ctl.Name = subreportName
        End If
        
        ctl.BorderStyle = 0
        ctl.SourceObject = "rptContractsPage" & i
        ctl.LinkMasterFields = "ContractID"
        ctl.LinkChildFields = "ContractID"
        ctl.CanShrink = False
        ctl.CanGrow = True
    Next i
    
    SetReportMargin rpt
    
    DoCmd.Close acReport, rptName, acSaveYes
    DoCmd.OpenReport rptName, acViewDesign
    
End Sub

Public Sub SetUp_rptForm6()
    
    Dim rptName: rptName = "rptForm6"
    DoCmd.OpenReport rptName, acViewDesign
    
    Dim rpt As Report: Set rpt = Reports(rptName)
    
    rpt.RecordSource = "tblForm6"
    
    Dim i As Integer, ctl As Control
    For i = 1 To CountImageFiles("C:\Users\User\Desktop\Client Files\Richard F\5-9-2024\Document files\Form 6\cropped\")
        Dim subreportName: subreportName = "subForm6" & i
        If ControlExists(subreportName, rpt) Then
            Set ctl = rpt(subreportName)
        Else
            Set ctl = CreateReportControl(rptName, acSubform, acDetail, , , 0, 0 + (300 * (i - 1)), InchToTwip(8), 300)
            ctl.Name = subreportName
        End If
        
        ctl.BorderStyle = 0
        ctl.SourceObject = "rptForm6Page" & i
        ctl.LinkMasterFields = "Form6ID"
        ctl.LinkChildFields = "Form6ID"
        ctl.CanShrink = False
        ctl.CanGrow = True
    Next i
    
    SetReportMargin rpt
    
    DoCmd.Close acReport, rptName, acSaveYes
    DoCmd.OpenReport rptName, acViewDesign
    
End Sub

Private Function SetReportMargin(rpt As Report)
    
    rpt.Printer.BottomMargin = InchToTwip(0)
    rpt.Printer.TopMargin = InchToTwip(0.25)
    rpt.Printer.LeftMargin = InchToTwip(0.25)
    rpt.Printer.RightMargin = InchToTwip(0.25)
    
    
End Function

Public Sub CopyReportWithImage(number, frmName, croppedFilesDirectory)

    Dim originalReportName: originalReportName = "rpt" & frmName & "Page"
    ' Define the new form name with the number appended
    Dim newReportName As String
    newReportName = originalReportName & number
    
    ' Check if a form with the new name already exists and delete it if it does
    If ReportExists(newReportName) Then
        DoCmd.DeleteObject acReport, newReportName
    End If
    
    ' Copy the original form to the new form name
    DoCmd.CopyObject , newReportName, acReport, originalReportName
    
    ' Open the new form in design view
    DoCmd.OpenReport newReportName, acDesign
    
    ' Set the picture path for the imgBackground control
    Dim imgBackground As Control
    Set imgBackground = Reports(newReportName).Controls("imgBackground")
    
    Reports(newReportName).Printer.BottomMargin = InchToTwip(0)
    Reports(newReportName).Printer.TopMargin = InchToTwip(0.25)
    Reports(newReportName).Printer.LeftMargin = InchToTwip(0.25)
    Reports(newReportName).Printer.RightMargin = InchToTwip(0.25)
    
    Dim width, height
    width = InchToTwip(8)
    height = InchToTwip(10.75)
    imgBackground.Picture = croppedFilesDirectory & frmName & "_page-" & Format(number, "0000") & "_cropped.jpg"
    imgBackground.width = width
    imgBackground.height = height
    ' Save and close the new form
    DoCmd.Close acReport, newReportName, acSaveYes
    
End Sub

Public Sub CopyFormWithAllImage(frmName, croppedFilesDirectory As String, Optional reportMode As Boolean = False)
    
    Dim i As Integer
    Dim fileCount As Integer
    
    ' Count the number of image files in the directory
    fileCount = CountImageFiles(croppedFilesDirectory)
    
    For i = 1 To fileCount
        If reportMode Then
            CopyReportWithImage i, frmName, croppedFilesDirectory
        Else
            CopyFormWithImage i, frmName, croppedFilesDirectory
        End If
    Next i
    
End Sub

Private Function CountImageFiles(directoryPath As String) As Integer
    Dim fileSys As Object
    Dim objFile As Object
    Dim fileCount As Integer
    
    ' Create a File System Object
    Set fileSys = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the directory exists
    If fileSys.FolderExists(directoryPath) Then
        ' Loop through each file in the directory
        For Each objFile In fileSys.GetFolder(directoryPath).Files
            ' Check if the file is an image file (you may need to adjust this condition based on your image file extensions)
            If IsImageFile(objFile.Path) Then
                ' Increment file count
                fileCount = fileCount + 1
            End If
        Next objFile
    End If
    
    ' Return the count of image files
    CountImageFiles = fileCount
End Function

Private Function IsImageFile(filePath As String) As Boolean
    ' Check if the file extension is an image format you want to include
    Select Case LCase(Right(filePath, 4))
        Case ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff"
            IsImageFile = True
        Case Else
            IsImageFile = False
    End Select
End Function


Public Sub ModifyTabPage(pageNumber, customFormName, croppedImagesDirectory As String, LinkFields)
    
    Dim left, top, width, height
    ''subform positions
    Dim subformName: subformName = "subform1"
    
    Dim frmName: frmName = "frm" & customFormName
    DoCmd.OpenForm frmName, acDesign
    Dim frm As Form: Set frm = Forms(frmName)
    
    Dim ctl As Control
    Set ctl = frm(subformName)
    left = ctl.left
    top = ctl.top
    width = ctl.width
    height = ctl.height
    
    ' Get the tab control and the form
    Dim tabPages As TabControl: Set tabPages = frm.Controls("tabPages")
    Dim tabCount: tabCount = CountImageFiles(croppedImagesDirectory)
    If tabCount < tabPages.Pages.Count Then
        Dim i As Integer
        For i = tabPages.Pages.Count To tabCount + 1 Step -1
            tabPages.Pages.Remove i - 1
        Next i
    End If

    If tabCount < pageNumber Then
        Do While tabCount < pageNumber
            ' Create tabs until the desired pageNumber is reached.
            Set ctl = tabPages.Pages.Add
            ctl.Name = "page" & pageNumber
            ctl.Caption = "Page " & pageNumber
            tabCount = tabPages.Pages.Count
        Loop
    End If
    
    subformName = "subform" & pageNumber
    If ControlExists(subformName, frm) Then
        Set ctl = frm(subformName)
        ctl.left = left
        ctl.top = top
        ctl.width = width
        ctl.height = height
    Else
        Set ctl = CreateControl(frmName, acSubform, acDetail, "page" & pageNumber, , left, top, width, height)
        ctl.Name = subformName
        
    End If
    
    ctl.SourceObject = frmName & "Page" & pageNumber
    ctl.LinkMasterFields = LinkFields
    ctl.LinkChildFields = LinkFields

End Sub

Public Sub ModifyAllTabPages(frmName, croppedImagesDirectory As String, LinkFields)
    
    Dim i As Integer
    Dim tabCount: tabCount = CountImageFiles(croppedImagesDirectory)
    For i = 1 To tabCount
        ModifyTabPage i, frmName, croppedImagesDirectory, LinkFields
    Next i
    
End Sub

Public Function frmContracts_vPropertyListID_AfterUpdate(frm As Form)
    
'    Dim shouldRequery: shouldRequery = Not frm.NewRecord
    DoCmd.RunCommand acCmdSaveRecord
    
    Dim ctl As Control
    For Each ctl In frm.Controls
        If ctl.ControlType = acSubform Then
            RunOnCurrent ctl, frm
        End If
    Next ctl
    
'    If shouldRequery Then frm.Requery
    
End Function

Private Sub RunOnCurrent(subformControl As Control, mainForm As Form)

    Dim frm As Form: Set frm = subformControl.Form
    Dim ctl As Control
    
    For Each ctl In frm.Controls
        If ctl.ControlType = acComboBox And ctl.Name <> "CreatedBy" Then
            mainForm(ctl.Name) = Null
        End If
    Next ctl
    
    If Not isFalse(frm.OnCurrent) Then
        frm.Requery
        ''Run frm.Name & "_OnCurrent", frm
    End If
    
End Sub

Private Function SetComboboxRowSource(frm As Form, whereClause, comboboxName, Optional useMember As Boolean = False, Optional addNIL As Boolean = False)
    
    Dim PropertyListID: PropertyListID = frm("PropertyListID")
    Dim sqlStr
    ''Set the seller id
    If isFalse(PropertyListID) Then
        sqlStr = "SELECT EntityID,EntityName,1 As RecordOrder FROM qryPropertyEntities WHERE PropertyEntityID = 0"
    Else
        If useMember Then
            sqlStr = "SELECT EntityMemberID,MemberName,1 As RecordOrder FROM qryPropertyEntityMembers WHERE PropertyListID = " & PropertyListID
            If Not isFalse(whereClause) Then
                sqlStr = sqlStr & " AND " & whereClause
            End If
            sqlStr = sqlStr & IIf(addNIL, " ORDER BY RecordOrder", " ORDER BY MemberName")
        Else
            sqlStr = "SELECT EntityID,EntityName,1 As RecordOrder FROM qryPropertyEntities WHERE PropertyListID = " & PropertyListID & _
                " AND EntityName <> ""My Panda Realty"""
            If Not isFalse(whereClause) Then
                sqlStr = sqlStr & " AND " & whereClause
            End If
            sqlStr = sqlStr & IIf(addNIL, " ORDER BY RecordOrder", " ORDER BY EntityName")
        End If
        
    End If
    
    If addNIL Then
        sqlStr = "SELECT All_Number,[All],0 As RecordOrder FROM tblAlls WHERE [All] = ""NIL""" & " UNION ALL " & sqlStr
    End If
    
    frm(comboboxName).rowSource = sqlStr
    
End Function

Private Function PrependDatabaseFields(csvString, tblName) As String
    Dim items() As String
    Dim resultString As String
    Dim i As Integer
    
    ' Split the comma-separated string into an array
    items = Split(csvString, ",")
    
    ' Loop through each item and prepend the provided string
    For i = LBound(items) To UBound(items)
        ' Trim any leading/trailing spaces from the item
        items(i) = Trim(items(i))
        
        ' Prepend the string
        If Len(items(i)) > 0 Then
            items(i) = tblName & "." & items(i)
        End If
    Next i
    
    ' Join the items back into a comma-separated string
    resultString = Join(items, ",")
    
    ' Return the processed string
    PrependDatabaseFields = resultString
End Function

Public Function ExtractAfterAS(inputString As String) As String
    Dim position As Integer
    position = InStr(1, inputString, "AS ", vbTextCompare)
    
    If position > 0 Then
        ' "AS " was found in the string, return the part after it
        ExtractAfterAS = Mid(inputString, position + 3)
    Else
        ' "AS " was not found, return the whole string
        ExtractAfterAS = inputString
    End If
End Function

Private Function ProcessDatabaseFields(csvString, tblName, commonAlias) As String

    Dim items() As String
    Dim resultString As String
    Dim i As Integer
    
    ' Split the comma-separated string into an array
    items = Split(csvString, ",")
    
    ' Loop through each item and prepend the provided string
    For i = LBound(items) To UBound(items)
        ' Trim any leading/trailing spaces from the item
        items(i) = Trim(items(i))
        
        ' Prepend the string
        If Len(items(i)) > 0 Then
            
            Dim fldName: fldName = ExtractAfterAS(items(i))
            items(i) = tblName & "." & fldName & " AS " & commonAlias & fldName
        
        End If
    Next i
    
    ' Join the items back into a comma-separated string
    resultString = Join(items, ",")
    
    ' Return the processed string
    ProcessDatabaseFields = resultString
End Function

Public Function ExtractAddressPart(completeAddress, Optional mode As String = "StreetAddress") As String

    If isFalse(completeAddress) Then Exit Function
    
    Dim regex As Object
    Dim matches As Object
    Dim extractedPart As String
    
    ' Create a regex object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Configure the regex pattern to match "L" followed by a number and a space
    Dim pattern
    Select Case mode
        Case "StreetAddress", "Suburb":
            pattern = "(.*)\s\w+\s\d+"
        Case "State":
            pattern = ".+\s(\w+)\s\d+"
        Case "Postcode":
            pattern = ".+\s\w+\s(\d+)"
    End Select

    With regex
        .Global = False
        .IgnoreCase = True
        .pattern = pattern
    End With
    
    ' Execute the regex search
    Set matches = regex.Execute(completeAddress)
    
    ' If a match is found, extract the number
    If matches.Count > 0 Then
        extractedPart = matches.item(0).SubMatches(0)
        If mode = "StreetAddress" Or mode = "Suburb" Then
            extractedPart = ExtractSuburb(regex, extractedPart, mode)
        End If
    Else
        extractedPart = ""
    End If
    
    ' Return the extracted number
    ExtractAddressPart = extractedPart
End Function

Private Function ExtractSuburb(regex As Object, streetAddressAndSuburb As String, mode As String) As String
    Dim Suburb, originalSuburb
    Dim matches As Object
    
    ' Additional processing (try to find match for suburb
    regex.pattern = ".+\s(\w+)"
    Set matches = regex.Execute(streetAddressAndSuburb)
    If matches.Count > 0 Then
        Suburb = matches.item(0).SubMatches(0)
        originalSuburb = Suburb
        If isPresent("tblSuburbs", "Suburb = " & Esc(Suburb)) Then
            ExtractSuburb = SetExtractSuburb(mode, Suburb, streetAddressAndSuburb)
        Else
            ' Try another pattern
            regex.pattern = ".+\s(\w+\s\w+)"
            Set matches = regex.Execute(streetAddressAndSuburb)
            If matches.Count > 0 Then
                Suburb = matches.item(0).SubMatches(0)
                If isPresent("tblSuburbs", "Suburb = " & Esc(Suburb)) Then
                    ExtractSuburb = SetExtractSuburb(mode, Suburb, streetAddressAndSuburb)
                Else
                    ' Try another pattern
                    regex.pattern = ".+\s(\w+\s\w+\s\w+)"
                    Set matches = regex.Execute(streetAddressAndSuburb)
                    If matches.Count > 0 Then
                        Suburb = matches.item(0).SubMatches(0)
                        If isPresent("tblSuburbs", "Suburb = " & Esc(Suburb)) Then
                            ExtractSuburb = SetExtractSuburb(mode, Suburb, streetAddressAndSuburb)
                        Else
                            ExtractSuburb = SetExtractSuburb(mode, originalSuburb, streetAddressAndSuburb)
                        End If
                    Else
                        ExtractSuburb = ""
                    End If
                End If
            Else
                ExtractSuburb = ""
            End If
        End If
    Else
        ExtractSuburb = ""
    End If
End Function

Private Function SetExtractSuburb(mode As String, Suburb, streetAddressAndSuburb As String) As String
    If mode = "Suburb" Then
        SetExtractSuburb = Suburb
    Else
        SetExtractSuburb = replace(streetAddressAndSuburb, Suburb, "")
    End If
End Function

Sub TestExtractAddressPart()

    Dim address1 As String
    Dim address2 As String
    Dim address3 As String
    Dim address4 As String

    address1 = "18 PERCH CRCT SANDSTONE POINT QLD 4511"
    address2 = "12 Kookaburra Way Woodgate Qld 4660"
    address3 = "290/176-208 TORRENS ROAD (invalid)"
    address4 = "PO BOX 1010 BURPENGARY QLD 4505"

    ' Test Street Address
    MsgBoxResult address1, ExtractAddressPart(address1, "StreetAddress")
    MsgBoxResult address2, ExtractAddressPart(address2, "StreetAddress")
    MsgBoxResult address3, ExtractAddressPart(address3, "StreetAddress")
    MsgBoxResult address4, ExtractAddressPart(address4, "StreetAddress")
    
    ' Test State
    MsgBoxResult address1, ExtractAddressPart(address1, "State")
    MsgBoxResult address2, ExtractAddressPart(address2, "State")
    MsgBoxResult address3, ExtractAddressPart(address3, "State")
    MsgBoxResult address4, ExtractAddressPart(address4, "State")

    ' Test Postcode
    MsgBoxResult address1, ExtractAddressPart(address1, "Postcode")
    MsgBoxResult address2, ExtractAddressPart(address2, "Postcode")
    MsgBoxResult address3, ExtractAddressPart(address3, "Postcode")
    MsgBoxResult address4, ExtractAddressPart(address4, "Postcode")
    
End Sub

Public Sub MakeControlsTransparent(rpt As Report)

    Dim ctl As Control
    For Each ctl In rpt.Controls
        Select Case ctl.ControlType
            Case acTextBox, acComboBox:
                ctl.BackStyle = 0
        End Select
    Next ctl
    
End Sub


Public Function Setup_frmContractsPage1()
    
    Dim entityFieldStr: entityFieldStr = GetEntityFields().arr(0)
    Dim entitySqlStr: entitySqlStr = GetEntityFields().arr(1)
    Dim memberFieldStr: memberFieldStr = GetEntityFields("Member").arr(0)
    Dim memberSqlStr: memberSqlStr = GetEntityFields("Member").arr(1)
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "tblContracts"
          .fields = PrependDatabaseFields("ContractID,SellerID,Seller2ID,SellerSolicitorID,PropertyListID", "tblContracts") & "," & _
            PrependDatabaseFields("ContractDate", "tblPropertyList") & "," & _
            ProcessDatabaseFields(entityFieldStr, "tempSellers", "Seller") & "," & _
            ProcessDatabaseFields(entityFieldStr, "tempSellers2", "Seller2") & "," & _
            ProcessDatabaseFields(entityFieldStr, "tempSellerSolicitors", "SellerSolicitor")
          .Joins.Add GenerateJoinObj(entitySqlStr, "SellerID", "tempSellers", "EntityID", "LEFT")
          .Joins.Add GenerateJoinObj(entitySqlStr, "Seller2ID", "tempSellers2", "EntityID", "LEFT")
          .Joins.Add GenerateJoinObj(entitySqlStr, "SellerSolicitorID", "tempSellerSolicitors", "EntityID", "LEFT")
          .Joins.Add GenerateJoinObj("tblPropertyList", "PropertyListID")
'          .OrderBy =
'          .GroupBy =
'          .SourceAlias =
'          .Having =
          sqlStr = .sql
          ''.SQL
          ''Set rs = .Recordset
    End With
    
    Dim frmName: frmName = "frmContractsPage1"
    DoCmd.OpenForm frmName, acDesign
    Dim frm As Form: Set frm = Forms(frmName)
    frm.RecordSource = sqlStr
    
    ''Also set the report here
    SetUpReport frmName, sqlStr
    
    ''Fix the controls -> Colicitors -> should be Solicitor
    Dim ctl As Control
    For Each ctl In frm.Controls
        If ctl.Name Like "*Colicitors*" Then
            ctl.Name = replace(ctl.Name, "Colicitors", "Solicitor")
            ctl.ControlSource = replace(ctl.ControlSource, "Colicitors", "Solicitor")
            
        End If
    Next ctl
    
    SetAddressPartControlSource frm, "Seller"
    SetAddressPartControlSource frm, "Seller2"
    SetAddressPartControlSource frm, "SellerSolicitor"
    
    ''Fix the controls -> use Seller2 fields -> Uncomment since this is already fixed.
'    Dim i As Integer, ControlSource, NewControlSource
'    For i = 24 To 32
'        Set ctl = frm("Text" & i)
'        ControlSource = ctl.ControlSource
'        If Not ControlSource Like "Seller2*" Then
'            NewControlSource = replace(ctl.ControlSource, "Seller", "Seller2")
'            ctl.ControlSource = NewControlSource
'            ctl.Name = NewControlSource
'        End If
'
'    Next i
    
    frm.OnCurrent = "=frmContractsPage1_OnCurrent([Form])"
    
End Function

Public Function Setup_frmForm6Page6()
    
    Dim entityFieldStr: entityFieldStr = GetEntityFields().arr(0)
    Dim entitySqlStr: entitySqlStr = GetEntityFields().arr(1)
    Dim memberFieldStr: memberFieldStr = GetEntityFields("Member").arr(0)
    Dim memberSqlStr: memberSqlStr = GetEntityFields("Member").arr(1)
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "tblForm6"
          .fields = PrependDatabaseFields("Form6ID,SellerID,Seller2ID,Agent,SellerSignatureDate,Seller2SignatureDate," & _
            "AgentSignatureDate,SchedulesAndAttachment,AgentReappointed,AgentReappointSignatureDate,SellerReappointSignatureDate,Seller2ReappointSignatureDate,PropertyListID", "tblForm6") & "," & _
            ProcessDatabaseFields(entityFieldStr, "tempSellers", "Seller") & "," & _
            ProcessDatabaseFields(entityFieldStr, "tempSellers2", "Seller2")
          .Joins.Add GenerateJoinObj(entitySqlStr, "SellerID", "tempSellers", "EntityID", "LEFT")
          .Joins.Add GenerateJoinObj(entitySqlStr, "Seller2ID", "tempSellers2", "EntityID", "LEFT")
          .Joins.Add GenerateJoinObj("tblPropertyList", "PropertyListID")
'          .OrderBy =
'          .GroupBy =
'          .SourceAlias =
'          .Having =
          sqlStr = .sql
          ''.SQL
          ''Set rs = .Recordset
    End With
    

    Dim frmName: frmName = "frmForm6Page6"
    DoCmd.OpenForm frmName, acDesign
    Dim frm As Form: Set frm = Forms(frmName)
    frm.RecordSource = sqlStr
    
    RenameBaseControlSource frm.Name
    
    SetUpForm6Report frmName, sqlStr, True
    
    frm.OnCurrent = "=" & frmName & "_OnCurrent([Form])"
    
End Function

Public Sub RenameBaseControlSource(frmName)
    
    DoCmd.OpenForm frmName, acDesign
    Dim frm As Form: Set frm = Forms(frmName)

    Dim ctl As Control, ControlSource
    For Each ctl In frm.Controls
        If ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Then
            ControlSource = ctl.ControlSource
            If Not ControlSource Like "=*" Then
                ctl.Name = ControlSource
            End If
        End If
    Next ctl
    
End Sub

Public Function Setup_frmForm6Page1()
    
    Dim entityFieldStr: entityFieldStr = GetEntityFields().arr(0)
    Dim entitySqlStr: entitySqlStr = GetEntityFields().arr(1)
    Dim memberFieldStr: memberFieldStr = GetEntityFields("Member").arr(0)
    Dim memberSqlStr: memberSqlStr = GetEntityFields("Member").arr(1)
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "tblForm6"
          .fields = PrependDatabaseFields("Form6ID,SellerID,Seller2ID,PropertyListID", "tblForm6") & "," & _
            PrependDatabaseFields("ContractDate", "tblPropertyList") & "," & _
            ProcessDatabaseFields(entityFieldStr, "tempSellers", "Seller") & "," & _
            ProcessDatabaseFields(entityFieldStr, "tempSellers2", "Seller2")
          .Joins.Add GenerateJoinObj(entitySqlStr, "SellerID", "tempSellers", "EntityID", "LEFT")
          .Joins.Add GenerateJoinObj(entitySqlStr, "Seller2ID", "tempSellers2", "EntityID", "LEFT")
          .Joins.Add GenerateJoinObj("tblPropertyList", "PropertyListID")
'          .OrderBy =
'          .GroupBy =
'          .SourceAlias =
'          .Having =
          sqlStr = .sql
          ''.SQL
          ''Set rs = .Recordset
    End With
    
    Dim frmName: frmName = "frmForm6Page1"
    DoCmd.OpenForm frmName, acDesign
    Dim frm As Form: Set frm = Forms(frmName)
    frm.RecordSource = sqlStr
    
    ''Also set the report here
    SetUpForm6Report frmName, sqlStr
    
    ''Fix the controls -> Colicitors -> should be Solicitor
    Dim ctl As Control, ControlSource
    For Each ctl In frm.Controls
        If ctl.ControlType = acTextBox Then
            ControlSource = ctl.ControlSource
            If Not ControlSource Like "=*" Then
                ctl.Name = ControlSource
            End If
        End If
'        If ctl.Name Like "*Colicitors*" Then
'            ctl.Name = replace(ctl.Name, "Colicitors", "Solicitor")
'            ctl.ControlSource = replace(ctl.ControlSource, "Colicitors", "Solicitor")
'
'        End If
    Next ctl
    
    SetAddressPartControlSource frm, "Seller"
    SetAddressPartControlSource frm, "Seller2"
'   SetAddressPartControlSource frm, "SellerSolicitor"
    RenameBaseControlSource frm.Name
    
    ''Fix the controls -> use Seller2 fields -> Uncomment since this is already fixed.
'    Dim i As Integer, ControlSource, NewControlSource
'    For i = 24 To 32
'        Set ctl = frm("Text" & i)
'        ControlSource = ctl.ControlSource
'        If Not ControlSource Like "Seller2*" Then
'            NewControlSource = replace(ctl.ControlSource, "Seller", "Seller2")
'            ctl.ControlSource = NewControlSource
'            ctl.Name = NewControlSource
'        End If
'
'    Next i
    
    frm.OnCurrent = "=" & frmName & "_OnCurrent([Form])"
    
End Function

Public Function Setup_frmForm6Page4()
    
    Dim entityFieldStr: entityFieldStr = GetEntityFields().arr(0)
    Dim entitySqlStr: entitySqlStr = GetEntityFields().arr(1)
    Dim memberFieldStr: memberFieldStr = GetEntityFields("Member").arr(0)
    Dim memberSqlStr: memberSqlStr = GetEntityFields("Member").arr(1)
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "tblForm6"
          .fields = PrependDatabaseFields("Form6ID,SellerID,PropertyListID,ClientAcknowledgementDate,Agent," & _
            "AgentAcknowledgementDate,Commission,WhenCommissionIsPayable,OtherWhenCommissionIsPayable", "tblForm6") & "," & _
            PrependDatabaseFields("ContractDate", "tblPropertyList") & "," & _
            ProcessDatabaseFields(entityFieldStr, "tempSellers", "Seller")
          .Joins.Add GenerateJoinObj(entitySqlStr, "SellerID", "tempSellers", "EntityID", "LEFT")
          .Joins.Add GenerateJoinObj("tblPropertyList", "PropertyListID")
'          .OrderBy =
'          .GroupBy =
'          .SourceAlias =
'          .Having =
          sqlStr = .sql
          ''.SQL
          ''Set rs = .Recordset
    End With
    
    Dim frmName: frmName = "frmForm6Page4"
    DoCmd.OpenForm frmName, acDesign
    Dim frm As Form: Set frm = Forms(frmName)
    frm.RecordSource = sqlStr
    
    ''Also set the report here
    SetUpForm6Report frmName, sqlStr, True
    
    ''Fix the controls -> Colicitors -> should be Solicitor
    Dim ctl As Control, ControlSource
    For Each ctl In frm.Controls
        If ctl.ControlType = acTextBox Then
            ControlSource = ctl.ControlSource
            If Not ControlSource Like "=*" Then
                ctl.Name = ControlSource
            End If
        End If
'        If ctl.Name Like "*Colicitors*" Then
'            ctl.Name = replace(ctl.Name, "Colicitors", "Solicitor")
'            ctl.ControlSource = replace(ctl.ControlSource, "Colicitors", "Solicitor")
'
'        End If
    Next ctl
    
'    SetAddressPartControlSource frm, "Seller"
'    SetAddressPartControlSource frm, "Seller2"
'   SetAddressPartControlSource frm, "SellerSolicitor"
    
    ''Fix the controls -> use Seller2 fields -> Uncomment since this is already fixed.
'    Dim i As Integer, ControlSource, NewControlSource
'    For i = 24 To 32
'        Set ctl = frm("Text" & i)
'        ControlSource = ctl.ControlSource
'        If Not ControlSource Like "Seller2*" Then
'            NewControlSource = replace(ctl.ControlSource, "Seller", "Seller2")
'            ctl.ControlSource = NewControlSource
'            ctl.Name = NewControlSource
'        End If
'
'    Next i
    
    CreateFormOptionGroups frm, "WhenCommissionIsPayable"
    SetCustomFormDimensionControls frm
    RenameBaseControlSource frm.Name
    
    frm.OnCurrent = "=" & frmName & "_OnCurrent([Form])"
    
End Function

Public Function Setup_frmForm6Page5()
    
    Dim entityFieldStr: entityFieldStr = GetEntityFields().arr(0)
    Dim entitySqlStr: entitySqlStr = GetEntityFields().arr(1)
    Dim memberFieldStr: memberFieldStr = GetEntityFields("Member").arr(0)
    Dim memberSqlStr: memberSqlStr = GetEntityFields("Member").arr(1)
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "tblForm6"
          .fields = PrependDatabaseFields("Form6ID,Premium,Platinum,AuthorisedAmount,WhenPayable,MaxRepair,MaxRoutineRepair,OtherFeeDescription1,OtherFeeAmount1," & _
            "OtherFeeWhenPayable1,OtherFeeDescription2,OtherFeeAmount2,OtherFeeWhenPayable2,OtherFeeDescription3,OtherFeeAmount3,OtherFeeWhenPayable3," & _
            "AgentRebateService1,AgentRebateSource1,AgentRebateEstimatedAmount1,AgentRebateService2,AgentRebateSource2,AgentRebateEstimatedAmount2,AgentRebateService3," & _
            "AgentRebateSource3,AgentRebateEstimatedAmount3,PropertyListID", "tblForm6") & "," & _
            PrependDatabaseFields("ContractDate", "tblPropertyList")
          .Joins.Add GenerateJoinObj("tblPropertyList", "PropertyListID")
'          .OrderBy =
'          .GroupBy =
'          .SourceAlias =
'          .Having =
          sqlStr = .sql
          ''.SQL
          ''Set rs = .Recordset
    End With
    
    Dim frmName: frmName = "frmForm6Page5"
    DoCmd.OpenForm frmName, acDesign
    Dim frm As Form: Set frm = Forms(frmName)
    frm.RecordSource = sqlStr
    
    ''Also set the report here
    ''SetUpReport frmName, sqlStr
    SetUpForm6Report frmName, sqlStr, True
    
'    SetAddressPartControlSource frm, "Seller"
'    SetAddressPartControlSource frm, "Seller2"
'   SetAddressPartControlSource frm, "SellerSolicitor"
    
    ''Fix the controls -> use Seller2 fields -> Uncomment since this is already fixed.
'    Dim i As Integer, ControlSource, NewControlSource
'    For i = 24 To 32
'        Set ctl = frm("Text" & i)
'        ControlSource = ctl.ControlSource
'        If Not ControlSource Like "Seller2*" Then
'            NewControlSource = replace(ctl.ControlSource, "Seller", "Seller2")
'            ctl.ControlSource = NewControlSource
'            ctl.Name = NewControlSource
'        End If
'
'    Next i
    
'    CreateFormOptionGroups frm, "WhenCommissionIsPayable"
    SetCustomFormDimensionControls frm
    RenameBaseControlSource frm.Name
    
    frm.OnCurrent = "=" & frmName & "_OnCurrent([Form])"
    
End Function

Private Function RemoveLastTwoChars(ByVal str As String) As String
    If Len(str) > 2 Then
        RemoveLastTwoChars = left(str, Len(str) - 2)
    Else
        RemoveLastTwoChars = ""
    End If
End Function

Public Function GetControlDimension(rptName, ctlName)

    Dim ctl As Control: Set ctl = Reports(rptName).Controls(ctlName)
    Dim top, left, width
    top = ctl.top
    left = ctl.left
    width = ctl.width
    
    Debug.Print top & "," & left & "," & width
    
End Function

Private Function SetUpForm6Report(frmName, sqlStr, Optional bigCheckBoxUsed As Boolean = False)
    
    Dim rptName: rptName = replace(frmName, "frm", "rpt")
    DoCmd.OpenReport rptName, acDesign
    Dim rpt As Report: Set rpt = Reports(rptName)
    
    ''8.27 x 11.69
    rpt("imgBackground").height = InchToTwip(11.69 - 0.5)
    rpt("imgBackground").width = InchToTwip(8.2 - 0.5)
    
    
    If sqlStr <> "tblContracts" Then
        Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
        Dim fld As Field, FieldNames As New clsArray
        For Each fld In rs.fields
            FieldNames.Add fld.Name, True
        Next fld
        
        sqlStr = "SELECT " & FieldNames.JoinArr(",") & " FROM (" & sqlStr & ") temp GROUP BY  " & FieldNames.JoinArr(",")
    End If
    
    rpt.RecordSource = sqlStr
    
    MakeControlsTransparent rpt
    
    DoCmd.OpenForm frmName, acDesign
    Dim ctl As Control
    Dim frm As Form
    Set frm = Forms(frmName)
    
    For Each ctl In frm.Controls
        ''Also not picture type
        If ctl.Name <> "CreatedBy" And ctl.Name <> "Timestamp" And ctl.ControlType <> 103 Then
            Dim top, left, ctlName, width
            top = ctl.top
            left = ctl.left
            width = ctl.width
            ctlName = ctl.Name
            
            Dim newtop, newLeft, newwidth
            newtop = top * 0.935779816513762
            newLeft = left * 0.899328859060403
            newwidth = width * 0.897637795275591
            
            Select Case ctl.ControlType
                Case acTextBox:
                    
                    rpt(ctlName).top = newtop
                    rpt(ctlName).left = newLeft
                    rpt(ctlName).width = newwidth
                
                Case acComboBox:
                    
                    rpt(ctlName).top = newtop
                    rpt(ctlName).left = newLeft
                    rpt(ctlName).width = newwidth
                    
                    Dim rowSource
                    
                    If ctlName = "BuyerID" Or ctlName = "Buyer2ID" Or ctlName = "PropertyMangerID" Then
                        ''BuyerID Buyer2ID, PropertyMangerID
                        rowSource = "SELECT EntityMemberID,MemberName AS MainField,1 As RecordOrder FROM qryPropertyEntityMembers GROUP BY EntityMemberID,MemberName"
                    Else
                        rowSource = "SELECT EntityID,EntityName As MainField, 1 As RecordOrder FROM tblEntities"
                    End If
            
                    rowSource = "SELECT All_Number,[All],0 As RecordOrder FROM tblAlls WHERE [All] = ""NIL""" & " UNION ALL " & rowSource

                    rpt(ctlName).rowSource = rowSource
                    
                Case acCommandButton:
                    
                    ctl.Visible = True
                    
                    Dim newCtl As Control, newCtlName, fieldName
                    newCtlName = ctlName & "1"
                    fieldName = replace(ctlName, "cmd", "")
                    
                    newtop = newtop - InchToTwip(0.0175)
                    
                    If bigCheckBoxUsed Then
                        newtop = newtop + InchToTwip(0.0175)
                        newLeft = newLeft + InchToTwip(0.0175)
                    End If
                    
                    If Not ControlExists(newCtlName, rpt) Then
                        Set newCtl = CreateReportControl(rptName, acTextBox, , , , newLeft, newtop, newwidth, ctl.height)
                        newCtl.Name = newCtlName
                    Else
                        Set newCtl = rpt(newCtlName)
                        newCtl.left = newLeft
                        newCtl.top = newtop
                        newCtl.width = newwidth
                    End If
                    
                    ''IIf(fieldValue And Not IsNull(fieldValue), ChrW$(10004), " ")
                    If fieldName Like "*_0" Then
                        newCtl.ControlSource = "=iif(Not [" & RemoveLastTwoChars(fieldName) & "],ChrW$(10004),"""")"
                    ElseIf fieldName Like "*_1" Then
                        newCtl.ControlSource = "=iif([" & RemoveLastTwoChars(fieldName) & "],ChrW$(10004),"""")"
                    Else
                        newCtl.ControlSource = "=iif([" & fieldName & "],ChrW$(10004),"""")"
                    End If
                    newCtl.BackStyle = 0
                    newCtl.fontSize = 9
                    newCtl.FontBold = True
                    newCtl.BorderStyle = 0
                    
                    
            End Select
            
            If ctlName Like "*ABN" Then
                CreateBoxedControls rpt, rpt(ctlName), 11
            End If
            
            If ctlName Like "*State" Then
                CreateBoxedControls rpt, rpt(ctlName), 3
            End If
            
            If ctlName Like "*PostCode" Then
                CreateBoxedControls rpt, rpt(ctlName), 4
            End If
            
            Dim dateFields As New clsArray: dateFields.arr = "AuctionDate,SingleAppointmentStart,SingleAppointmentEnd,ContinuingAppointmentStart" & _
                ",ClientAcknowledgementDate,AgentAcknowledgementDate,WhenPayable,SellerSignatureDate,Seller2SignatureDate,AgentSignatureDate,AgentReappointSignatureDate" & _
                ",SellerReappointSignatureDate,Seller2ReappointSignatureDate"
            
            If dateFields.InArray(ctlName) Then
                CreateBoxedDateControls rpt, rpt(ctlName)
            End If
            
        End If
    Next ctl
    
    rpt("Timestamp").ControlSource = ""
    rpt("CreatedBy").ControlSource = ""
    rpt.width = 0
    rpt.Section(acDetail).height = 0
    
    'DoCmd.Close acReport, rptName, acSaveYes
    'DoCmd.Close acForm, frmName, acSaveNo
    
End Function

Public Function GetNthCharacter(Value As Variant, character As Integer, Optional uppercase As Boolean = True, Optional AsDate As Boolean = False) As Variant
    ' Check if the value is empty or not a string
    If isFalse(Value) Then
        GetNthCharacter = Null
        Exit Function
    End If

    ' Convert the value to a string
    Dim strValue As String
    If AsDate Then
        strValue = Format$(Value, "DDMMYYYY")
    Else
        strValue = CStr(Value)
    End If
    

    ' Check if the character position is within the bounds of the string
    If character > Len(strValue) Then
        GetNthCharacter = Null
        Exit Function
    End If

    ' Return the nth character
    GetNthCharacter = Mid(strValue, character, 1)
    
    If uppercase Then GetNthCharacter = UCase(GetNthCharacter)
    
End Function



Private Sub CreateBoxedControls(rpt As Report, ctl As Control, boxes As Integer)
    
    Dim rptName: rptName = rpt.Name
    Dim left: left = ctl.left
    Dim height: height = ctl.height
    Dim top: top = ctl.top
    Dim width: width = ctl.width
    Dim eachBoxWidth: eachBoxWidth = width / boxes
    Dim originalControlSource: originalControlSource = ctl.ControlSource
    
    If originalControlSource Like "=*" Then
        originalControlSource = Right(originalControlSource, Len(originalControlSource) - 1)
    Else
        originalControlSource = "[" & originalControlSource & "]"
    End If
    
    Dim i, newCtl, newLeft, newCtlName As String
    newLeft = left - InchToTwip(0.02)
    For i = 1 To boxes
        newCtlName = ctl.Name & "_" & i
        
        If Not DoesPropertyExists(rpt, newCtlName) Then
            Set newCtl = CreateReportControl(rptName, acTextBox, acDetail, , , newLeft, top, eachBoxWidth, height)
            newCtl.Name = newCtlName
        Else
            Set newCtl = rpt(newCtlName)
            newCtl.left = newLeft
            newCtl.top = top
            newCtl.width = eachBoxWidth
            newCtl.height = height
        End If
        CopyControlFormatting rpt, ctl.Name, newCtlName
        newCtl.ControlSource = "=GetNthCharacter(" & originalControlSource & "," & i & ")"
        newLeft = newLeft + eachBoxWidth
        newCtl.TextAlign = 2
        
    Next i
    
    ctl.Visible = False
    
End Sub

Private Sub CreateBoxedDateControls(rpt As Report, ctl As Control)
    
    'hardcoded variables
    Dim totalSpace: totalSpace = InchToTwip(0.28)
    Dim spaceWidth: spaceWidth = totalSpace / 2
    Dim boxes: boxes = 8
    
    Dim rptName: rptName = rpt.Name
    Dim left: left = ctl.left
    Dim height: height = ctl.height
    Dim top: top = ctl.top
    Dim width: width = ctl.width
    Dim eachBoxWidth: eachBoxWidth = (width - totalSpace) / boxes
    Dim originalControlSource: originalControlSource = ctl.ControlSource
    
    ''For when the control source starts with =
    If originalControlSource Like "=*" Then
        originalControlSource = Right(originalControlSource, Len(originalControlSource) - 1)
    Else
        originalControlSource = "[" & originalControlSource & "]"
    End If
    
    Dim i, newCtl, newLeft, newCtlName As String
    newLeft = left - InchToTwip(0.02)
    For i = 1 To boxes
        newCtlName = ctl.Name & "_" & i
        
        If Not DoesPropertyExists(rpt, newCtlName) Then
            Set newCtl = CreateReportControl(rptName, acTextBox, acDetail, , , newLeft, top, eachBoxWidth, height)
            newCtl.Name = newCtlName
        Else
            Set newCtl = rpt(newCtlName)
            newCtl.left = newLeft
            newCtl.top = top
            newCtl.width = eachBoxWidth
            newCtl.height = height
        End If
        CopyControlFormatting rpt, ctl.Name, newCtlName
        newCtl.ControlSource = "=GetNthCharacter(" & originalControlSource & "," & i & ",False,True)"
        newLeft = newLeft + eachBoxWidth + IIf(i = 2 Or i = 4, spaceWidth, 0)
        newCtl.TextAlign = 2
        
    Next i
    
    ctl.Visible = False
    
End Sub

Private Sub CopyControlFormatting(frmOrRpt As Object, sourceCtlName As String, targetCtlName As String)
    Dim sourceCtl As Control
    Dim targetCtl As Control

    ' Set references to the source and target controls
    Set sourceCtl = frmOrRpt.Controls(sourceCtlName)
    Set targetCtl = frmOrRpt.Controls(targetCtlName)

    ' Copy common formatting properties
    On Error Resume Next ' Skip any properties that are not available for both controls
    
    targetCtl.BackColor = sourceCtl.BackColor
    targetCtl.ForeColor = sourceCtl.ForeColor
    targetCtl.FontName = sourceCtl.FontName
    targetCtl.fontSize = sourceCtl.fontSize
    targetCtl.FontBold = sourceCtl.FontBold
    targetCtl.FontItalic = sourceCtl.FontItalic
    targetCtl.FontUnderline = sourceCtl.FontUnderline
    targetCtl.TextAlign = sourceCtl.TextAlign
    targetCtl.BorderStyle = sourceCtl.BorderStyle
    targetCtl.BorderColor = sourceCtl.BorderColor
    targetCtl.BorderWidth = sourceCtl.BorderWidth

    On Error GoTo 0 ' Turn error handling back to default

    ' Clean up
    Set sourceCtl = Nothing
    Set targetCtl = Nothing
   
End Sub


Private Function SetUpReport(frmName, sqlStr, Optional paperSize = "Letter")
    
    Dim rptName: rptName = replace(frmName, "frm", "rpt")
    DoCmd.OpenReport rptName, acDesign
    Dim rpt As Report: Set rpt = Reports(rptName)
    
    ''8.27 x 11.69
    If paperSize = "A4" Then
        rpt("imgBackground").height = InchToTwip(11.69 - 0.5)
        rpt("imgBackground").width = InchToTwip(8.27 - 0.5)
    End If
    
    If sqlStr <> "tblContracts" Then
        Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)
        Dim fld As Field, FieldNames As New clsArray
        For Each fld In rs.fields
            FieldNames.Add fld.Name, True
        Next fld
        
        sqlStr = "SELECT " & FieldNames.JoinArr(",") & " FROM (" & sqlStr & ") temp GROUP BY  " & FieldNames.JoinArr(",")
    End If
    
    rpt.RecordSource = sqlStr
    
    MakeControlsTransparent rpt
    
    DoCmd.OpenForm frmName, acDesign
    Dim ctl As Control
    Dim frm As Form
    Set frm = Forms(frmName)
    
    For Each ctl In frm.Controls
        If ctl.Name <> "CreatedBy" And ctl.Name <> "Timestamp" Then
            Dim top, left, ctlName, width
            top = ctl.top
            left = ctl.left
            width = ctl.width
            ctlName = ctl.Name
            
            Dim newtop, newLeft, newwidth
            newtop = top * (1 - 0.099)
            newLeft = left * (1 - 0.072)
            newwidth = width * (1 - 0.057)
            
            Select Case ctl.ControlType
                Case acTextBox:
                    
                    rpt(ctlName).top = newtop
                    rpt(ctlName).left = newLeft
                    rpt(ctlName).width = newwidth
                
                Case acComboBox:
                    
                    rpt(ctlName).top = newtop
                    rpt(ctlName).left = newLeft
                    rpt(ctlName).width = newwidth
                    
                    Dim rowSource
                    
                    If ctlName = "BuyerID" Or ctlName = "Buyer2ID" Or ctlName = "PropertyMangerID" Then
                        ''BuyerID Buyer2ID, PropertyMangerID
                        rowSource = "SELECT EntityMemberID,MemberName AS MainField,1 As RecordOrder FROM qryPropertyEntityMembers GROUP BY EntityMemberID,MemberName"
                    Else
                        rowSource = "SELECT EntityID,EntityName As MainField, 1 As RecordOrder FROM tblEntities"
                    End If
            
                    rowSource = "SELECT All_Number,[All],0 As RecordOrder FROM tblAlls WHERE [All] = ""NIL""" & " UNION ALL " & rowSource

                    rpt(ctlName).rowSource = rowSource
                    
                Case acCommandButton:
                    
                    ctl.Visible = True
                    
                    Dim newCtl As Control, newCtlName, fieldName
                    newCtlName = ctlName & "1"
                    fieldName = replace(ctlName, "cmd", "")
                    If Not ControlExists(newCtlName, rpt) Then
                        Set newCtl = CreateReportControl(rptName, acTextBox, , , , newLeft, newtop - 50, newwidth, ctl.height)
                        newCtl.Name = newCtlName
                    Else
                        Set newCtl = rpt(newCtlName)
                    End If
                    
                    ''IIf(fieldValue And Not IsNull(fieldValue), ChrW$(10004), " ")
                    If fieldName Like "*_0" Then
                        newCtl.ControlSource = "=iif(Not [" & RemoveLastTwoChars(fieldName) & "],ChrW$(10004),"""")"
                    ElseIf fieldName Like "*_1" Then
                        newCtl.ControlSource = "=iif([" & RemoveLastTwoChars(fieldName) & "],ChrW$(10004),"""")"
                    Else
                        newCtl.ControlSource = "=iif([" & fieldName & "],ChrW$(10004),"""")"
                    End If
                    newCtl.BackStyle = 0
                    newCtl.fontSize = 9
                    newCtl.FontBold = True
                    newCtl.BorderStyle = 0
                    
                    
            End Select
        End If
    Next ctl
    
    rpt("Timestamp").ControlSource = ""
    rpt("CreatedBy").ControlSource = ""
    rpt.width = 0
    rpt.Section(acDetail).height = 0
    
    'DoCmd.Close acReport, rptName, acSaveYes
    'DoCmd.Close acForm, frmName, acSaveNo
    
End Function

Public Function SetAddressPartControlSource(frm As Form, EntityType)
    
    Dim mode, modes As New clsArray: modes.arr = "StreetAddress,Suburb,State,Postcode"
    Dim targetField, targetFields As New clsArray: targetFields.arr = "Address,ESuburb,EState,EPostcode"
    Dim ctlName
    Dim i As Integer: i = 0
    For Each targetField In targetFields.arr
        ctlName = EntityType & targetField
        If ControlExists(ctlName, frm) Then
            frm(ctlName).Name = "txt" & ctlName
        End If
        
        frm("txt" & ctlName).ControlSource = "=ExtractAddressPart([" & EntityType & "Address]," & Esc(modes.arr(i)) & ")"
        i = i + 1
    Next targetField
    
End Function

Public Function frmContractsPage1_OnCurrent(frm As Form, Optional reportMode As Boolean = False)
    
    SetComboboxRowSource frm, "EntityCategoryName = ""Seller""", "SellerID"
    SetComboboxRowSource frm, "EntityCategoryName = ""Seller""", "Seller2ID"
    SetComboboxRowSource frm, "EntityCategoryName = ""Contact"" AND ContactCategoryName = ""Solicitor""", "SellerSolicitorID"
'    Dim PropertyListID: PropertyListID = frm("PropertyListID")
'    Dim sqlStr
'    Dim whereClause: whereClause = "EntityCategoryName = ""Seller"""
'    ''Set the seller id
'    If isFalse(PropertyListID) Then
'        sqlStr = "SELECT EntityID,EntityName FROM qryPropertyEntities WHERE PropertyEntityID = 0"
'    Else
'        sqlStr = "SELECT EntityID,EntityName FROM qryPropertyEntities WHERE PropertyListID = " & PropertyListID
'        If Not isFalse(whereClause) Then
'            sqlStr = sqlStr & " AND " & whereClause
'        End If
'        sqlStr = sqlStr & " ORDER BY EntityName"
'    End If
'
'    frm("SellerID").RowSource = sqlStr
    
End Function

Public Sub CreateFormCheckboxes(frm As Form, FieldNames As String)
    
    ''Create a checkbox control
    ''Set the name
    ''Set the tag
    ''Set the onClick
    ''Set the size (height and width)
    Dim fieldName, fieldNamesArr As New clsArray: fieldNamesArr.arr = FieldNames
    Dim optionValue
    
    For Each fieldName In fieldNamesArr.arr
        CreateCheckboxControl frm, fieldName
        CreateCheckboxControl frm, fieldName
        
    Next fieldName
    
End Sub

Private Sub CreateCheckboxControl(frm As Form, fieldName)
    
    Dim ctlName: ctlName = "cmd" & fieldName
    
    Dim ctl As Control
    If DoesPropertyExists(frm, ctlName) Then
        Set ctl = frm(ctlName)
    Else
        Set ctl = CreateControl(frm.Name, acCommandButton, , , , 0, 0, 300, 300)
    End If
    
    ctl.BackStyle = 0
    ctl.BorderStyle = 0
    ctl.Name = ctlName
    ctl.OnClick = "=Toggle_frmContractsPageButtons([Form]," & Esc(fieldName) & ")"
    
End Sub

Public Sub CreateFormOptionGroups(frm As Form, FieldNames As String)
    
    ''Create a checkbox control
    ''Set the name
    ''Set the tag
    ''Set the onClick
    ''Set the size (height and width)
    Dim fieldName, fieldNamesArr As New clsArray: fieldNamesArr.arr = FieldNames
    Dim optionValue
    
    For Each fieldName In fieldNamesArr.arr
        CreateOptionControl frm, "0", fieldName
        CreateOptionControl frm, "1", fieldName
        
    Next fieldName
    
End Sub

Private Sub CreateOptionControl(frm As Form, optionValue, fieldName)
    
    Dim ctlName: ctlName = "cmd" & fieldName & "_" & optionValue
    
    Dim ctl As Control
    If DoesPropertyExists(frm, ctlName) Then
        Set ctl = frm(ctlName)
    Else
        Set ctl = CreateControl(frm.Name, acCommandButton, , , , 0, 0, 300, 300)
    End If
    
    ctl.BackStyle = 0
    ctl.BorderStyle = 0
    ctl.Name = ctlName
    ctl.Tag = optionValue
    ctl.OnClick = "=Toggle_frmContractsOptionGroup([Form]," & Esc(fieldName) & "," & Esc(ctlName) & ")"
    
End Sub

Public Function Toggle_frmContractsPageButtons(frm As Form, fieldName)
    
    Dim fieldValue: fieldValue = Not frm(fieldName)
    If IsNull(fieldValue) Then
        fieldValue = True
    End If
'    If Not IsNull(fieldValue) Then
'        fieldValue = False
'    End If
    SetToggleButtonCaption frm, fieldName, fieldValue
    frm(fieldName) = fieldValue
    
End Function

Private Function SetToggleButtonCaption(frm As Form, fieldName, fieldValue)
    
    ''Ctrl + F: Check Symbol CheckSymbol
    frm("cmd" & fieldName).Caption = IIf(fieldValue And Not IsNull(fieldValue), ChrW$(10004), " ")
    ''frm("cmd" & FieldName).Caption = IIf(fieldValue And Not IsNull(fieldValue), " ", ChrW$(10004))
    
End Function

Public Function Toggle_frmContractsOptionGroup(frm As Form, optionGroupName, buttonName)
    
    Dim btn As Control: Set btn = frm(buttonName)
    Dim buttonValue: buttonValue = btn.Caption <> " "
    
    Dim newValue: newValue = Not buttonValue
    frm(buttonName).Caption = IIf(newValue, ChrW$(10004), " ")
    
    Dim btnTag: btnTag = IIf(btn.Tag = "-1" Or btn.Tag = "1", True, False)
    frm(optionGroupName) = btnTag
    
    ''Untick other buttons
    Dim ctl As Control
    For Each ctl In frm.Controls
        If ctl.Name Like "cmd" & optionGroupName & "_*" And ctl.Name <> buttonName Then
            ctl.Caption = " "
        End If
    Next ctl
    
End Function

Public Function ExtractLotAndOn(RPD, Optional mode As String = "Lot") As String
    
    If isFalse(RPD) Then Exit Function
    
    Dim regex As Object
    Dim matches As Object
    Dim extractedNumber As String
    
    ' Create a regex object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Configure the regex pattern to match "L" followed by a number and a space
    With regex
        .Global = False
        .IgnoreCase = True
        .pattern = IIf(mode = "Lot", "L(\d+)\s.*", "L\d+\s(.*)")
    End With
    
    ' Execute the regex search
    Set matches = regex.Execute(RPD)
    
    ' If a match is found, extract the number
    If matches.Count > 0 Then
        extractedNumber = matches.item(0).SubMatches(0)
    Else
        extractedNumber = ""
    End If
    
    ' Return the extracted number
    ExtractLotAndOn = extractedNumber
End Function

Public Function GetEntityFields(Optional mode = "Entity") As clsArray
    
    Dim returnArr As New clsArray
    Dim fieldStr
    If mode = "Entity" Then
        fieldStr = "EntityID,EntityName,Address,ESuburb,EState,EPostcode,ABN,PhoneNumber,Mobile,Fax,EmailAddress,Ref,Contact,LicenseNo"
    Else
        fieldStr = "EntityMemberID As EntityID,MemberName AS EntityName, " & _
        "MemberAddress As Address,ABN,MemberPhoneNumber As PhoneNumber,Mobile,Fax,MemberEmailAddress AS EmailAddress,Ref,Contact,LicenseNo"
    End If
    
    Dim fields As New clsArray
    fields.arr = fieldStr
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = IIf(mode = "Entity", "tblEntities", "qryPropertyEntityMembers")
          .fields = fields.JoinArr
          sqlStr = .sql
    End With
    
    
    returnArr.Add fieldStr
    returnArr.Add sqlStr
    
    Set GetEntityFields = returnArr
    
End Function


Public Function Setup_frmContractsPage2()
    
    Dim entityFieldStr: entityFieldStr = GetEntityFields().arr(0)
    Dim entitySqlStr: entitySqlStr = GetEntityFields().arr(1)
    Dim memberFieldStr: memberFieldStr = GetEntityFields("Member").arr(0)
    Dim memberSqlStr: memberSqlStr = GetEntityFields("Member").arr(1)
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "tblContracts"
          .fields = PrependDatabaseFields("ContractID,BuyerID,Buyer2ID,BuyerAgentID,BuyerSolicitorID,BuiltOn,Vacant," & _
            "TitleReference,LandSoldFreeHold,LandSoldAsLeasehold,PresentUse,LocalGovernment,PropertyListID", "tblContracts") & "," & _
            "StreetAddress,Suburb,State,Postcode,LandSize,RPD," & _
            ProcessDatabaseFields(memberFieldStr, "tempBuyer", "Buyer") & "," & _
            ProcessDatabaseFields(memberFieldStr, "tempBuyer2", "Buyer2") & "," & _
            ProcessDatabaseFields(entityFieldStr, "tempBuyerAgent", "BuyerAgent") & "," & _
            ProcessDatabaseFields(entityFieldStr, "tempBuyerSolicitors", "BuyerSolicitor")
          .Joins.Add GenerateJoinObj("tblPropertyList", "PropertyListID")
          .Joins.Add GenerateJoinObj(memberSqlStr, "BuyerID", "tempBuyer", "EntityID", "LEFT")
          .Joins.Add GenerateJoinObj(memberSqlStr, "Buyer2ID", "tempBuyer2", "EntityID", "LEFT")
          .Joins.Add GenerateJoinObj(entitySqlStr, "BuyerAgentID", "tempBuyerAgent", "EntityID", "LEFT")
          .Joins.Add GenerateJoinObj(entitySqlStr, "BuyerSolicitorID", "tempBuyerSolicitors", "EntityID", "LEFT")
'          .OrderBy =
'          .GroupBy =
'          .SourceAlias =
'          .Having =
          sqlStr = .sql
          ''.SQL
          ''Set rs = .Recordset
    End With
    
    Dim frmName: frmName = "frmContractsPage2"
    DoCmd.OpenForm frmName, acDesign
    Dim frm As Form: Set frm = Forms(frmName)
    frm.RecordSource = sqlStr
    
    SetUpReport frmName, sqlStr
    
    ''Fix the controls -> Colicitors -> should be Solicitor
    Dim ctl As Control
    For Each ctl In frm.Controls
        If ctl.Name Like "*SellerSolicitor*" Then
            ctl.Name = replace(ctl.Name, "SellerSolicitor", "BuyerSolicitor")
            ctl.ControlSource = replace(ctl.ControlSource, "SellerSolicitor", "BuyerSolicitor")
        End If
    Next ctl
    
    ''Fix the controls -> use Seller2 fields -> Uncomment since this is already fixed.
    Dim i As Integer, ControlSource, NewControlSource
'    For i = 24 To 32
'        Set ctl = frm("Text" & i)
'        ControlSource = ctl.ControlSource
'        If Not ControlSource Like "Buyer2*" Then
'            NewControlSource = replace(ctl.ControlSource, "Buyer", "Buyer2")
'            ctl.ControlSource = NewControlSource
'            ctl.Name = NewControlSource
'        End If
'
'    Next i

'    Dim ctlName
'    For i = 34 To 43
'        ctlName = "Text" & i
'        If ControlExists(ctlName, frm) Then
'            Set ctl = frm(ctlName)
'            ControlSource = ctl.ControlSource
'            If ControlSource Like "Buyer2*" Then
'                NewControlSource = replace(ctl.ControlSource, "Buyer2", "BuyerAgent")
'                ctl.ControlSource = NewControlSource
'                ctl.Name = NewControlSource
'            End If
'        End If
'    Next i

'    SetComboboxRowSource frm, "EntityCategoryName = ""Buyer""", "BuyerID"
'    SetComboboxRowSource frm, "EntityCategoryName = ""Buyer""", "Buyer2ID"
'    SetComboboxRowSource frm, "EntityCategoryName = ""Contact""", "BuyerSolicitorID"
'    SetComboboxRowSource frm, "EntityCategoryName = ""Contact""", "BuyerAgentID"
    
    SetAddressPartControlSource frm, "Buyer"
    SetAddressPartControlSource frm, "Buyer2"
    SetAddressPartControlSource frm, "BuyerSolicitor"
    SetAddressPartControlSource frm, "BuyerAgent"
    
    frm("LotDescription").ControlSource = "=ExtractLotAndOn([RPD],""Lot"")"
    frm("OnDescription").ControlSource = "=ExtractLotAndOn([RPD],""On"")"
    
    frm.OnCurrent = "=" & frmName & "_OnCurrent([Form])"
    
End Function


Public Function Setup_frmForm6Page2()
    
    Dim entityFieldStr: entityFieldStr = GetEntityFields().arr(0)
    Dim entitySqlStr: entitySqlStr = GetEntityFields().arr(1)
    Dim memberFieldStr: memberFieldStr = GetEntityFields("Member").arr(0)
    Dim memberSqlStr: memberSqlStr = GetEntityFields("Member").arr(1)
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "tblForm6"
          .fields = PrependDatabaseFields("Form6ID,Description,Sale,Purchase,Letting,Auction,AuctionDate,Other,SingleAppointment," & _
            "SingleAppointmentStart,SingleAppointmentEnd,ContinuingAppointment,ContinuingAppointmentStart,OtherCheck," & _
            "PriceReserve,PriceList,PriceLetting,Price1,Price2,Instructions,PropertyListID,TitleReference", "tblForm6") & "," & _
            "StreetAddress,Suburb,State,Postcode,LandSize,RPD"
          .Joins.Add GenerateJoinObj("tblPropertyList", "PropertyListID")
'          .OrderBy =
'          .GroupBy =
'          .SourceAlias =
'          .Having =
          sqlStr = .sql
          ''.SQL
          ''Set rs = .Recordset
    End With
    
    Dim frmName: frmName = "frmForm6Page2"
    DoCmd.OpenForm frmName, acDesign
    Dim frm As Form: Set frm = Forms(frmName)
    frm.RecordSource = sqlStr
    
    SetUpForm6Report frmName, sqlStr
    
    ''Fix the controls -> Colicitors -> should be Solicitor
'    Dim ctl As Control
'    For Each ctl In frm.Controls
'        If ctl.Name Like "*SellerSolicitor*" Then
'            ctl.Name = replace(ctl.Name, "SellerSolicitor", "BuyerSolicitor")
'            ctl.ControlSource = replace(ctl.ControlSource, "SellerSolicitor", "BuyerSolicitor")
'        End If
'    Next ctl
    
    ''Fix the controls -> use Seller2 fields -> Uncomment since this is already fixed.
    Dim i As Integer, ControlSource, NewControlSource
'    For i = 24 To 32
'        Set ctl = frm("Text" & i)
'        ControlSource = ctl.ControlSource
'        If Not ControlSource Like "Buyer2*" Then
'            NewControlSource = replace(ctl.ControlSource, "Buyer", "Buyer2")
'            ctl.ControlSource = NewControlSource
'            ctl.Name = NewControlSource
'        End If
'
'    Next i

'    Dim ctlName
'    For i = 34 To 43
'        ctlName = "Text" & i
'        If ControlExists(ctlName, frm) Then
'            Set ctl = frm(ctlName)
'            ControlSource = ctl.ControlSource
'            If ControlSource Like "Buyer2*" Then
'                NewControlSource = replace(ctl.ControlSource, "Buyer2", "BuyerAgent")
'                ctl.ControlSource = NewControlSource
'                ctl.Name = NewControlSource
'            End If
'        End If
'    Next i

'    SetComboboxRowSource frm, "EntityCategoryName = ""Buyer""", "BuyerID"
'    SetComboboxRowSource frm, "EntityCategoryName = ""Buyer""", "Buyer2ID"
'    SetComboboxRowSource frm, "EntityCategoryName = ""Contact""", "BuyerSolicitorID"
'    SetComboboxRowSource frm, "EntityCategoryName = ""Contact""", "BuyerAgentID"
    
'    SetAddressPartControlSource frm, "Buyer"
'    SetAddressPartControlSource frm, "Buyer2"
'    SetAddressPartControlSource frm, "BuyerSolicitor"
'    SetAddressPartControlSource frm, "BuyerAgent"
    
'    frm("LotDescription").ControlSource = "=ExtractLotAndOn([RPD],""Lot"")"
'    frm("OnDescription").ControlSource = "=ExtractLotAndOn([RPD],""On"")"
    
    Dim ctl As Control
    For Each ctl In frm.Controls
        If ctl.Name Like "cmd*" Then
            ''=Toggle_frmContractsPageButtons([Form],"Sale")
            Dim fieldName: fieldName = GetCmdFieldName(ctl.Name)
            ctl.OnClick = "=Toggle_frmContractsPageButtons([Form]," & Esc(fieldName) & ")"
        End If
    Next ctl
    
    frm.OnCurrent = "=" & frmName & "_OnCurrent([Form])"
    
End Function

Public Function Setup_frmForm6Page3()
    
    Dim entityFieldStr: entityFieldStr = GetEntityFields().arr(0)
    Dim entitySqlStr: entitySqlStr = GetEntityFields().arr(1)
    Dim memberFieldStr: memberFieldStr = GetEntityFields("Member").arr(0)
    Dim memberSqlStr: memberSqlStr = GetEntityFields("Member").arr(1)
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "tblForm6"
          .fields = PrependDatabaseFields("Form6ID,OpenListing,SoleAgency,ExclusiveAgency,ExclusiveAgencyAgree,PropertyListID", "tblForm6") & "," & _
            "StreetAddress,Suburb,State,Postcode,LandSize,RPD"
          .Joins.Add GenerateJoinObj("tblPropertyList", "PropertyListID")
'          .OrderBy =
'          .GroupBy =
'          .SourceAlias =
'          .Having =
          sqlStr = .sql
          ''.SQL
          ''Set rs = .Recordset
    End With
    
    Dim frmName: frmName = "frmForm6Page3"
    DoCmd.OpenForm frmName, acDesign
    Dim frm As Form: Set frm = Forms(frmName)
    frm.RecordSource = sqlStr
    
    SetUpForm6Report frmName, sqlStr, True
    
    ''Fix the controls -> Colicitors -> should be Solicitor
'    Dim ctl As Control
'    For Each ctl In frm.Controls
'        If ctl.Name Like "*SellerSolicitor*" Then
'            ctl.Name = replace(ctl.Name, "SellerSolicitor", "BuyerSolicitor")
'            ctl.ControlSource = replace(ctl.ControlSource, "SellerSolicitor", "BuyerSolicitor")
'        End If
'    Next ctl
    
    ''Fix the controls -> use Seller2 fields -> Uncomment since this is already fixed.
    Dim i As Integer, ControlSource, NewControlSource
'    For i = 24 To 32
'        Set ctl = frm("Text" & i)
'        ControlSource = ctl.ControlSource
'        If Not ControlSource Like "Buyer2*" Then
'            NewControlSource = replace(ctl.ControlSource, "Buyer", "Buyer2")
'            ctl.ControlSource = NewControlSource
'            ctl.Name = NewControlSource
'        End If
'
'    Next i

'    Dim ctlName
'    For i = 34 To 43
'        ctlName = "Text" & i
'        If ControlExists(ctlName, frm) Then
'            Set ctl = frm(ctlName)
'            ControlSource = ctl.ControlSource
'            If ControlSource Like "Buyer2*" Then
'                NewControlSource = replace(ctl.ControlSource, "Buyer2", "BuyerAgent")
'                ctl.ControlSource = NewControlSource
'                ctl.Name = NewControlSource
'            End If
'        End If
'    Next i

'    SetComboboxRowSource frm, "EntityCategoryName = ""Buyer""", "BuyerID"
'    SetComboboxRowSource frm, "EntityCategoryName = ""Buyer""", "Buyer2ID"
'    SetComboboxRowSource frm, "EntityCategoryName = ""Contact""", "BuyerSolicitorID"
'    SetComboboxRowSource frm, "EntityCategoryName = ""Contact""", "BuyerAgentID"
    
'    SetAddressPartControlSource frm, "Buyer"
'    SetAddressPartControlSource frm, "Buyer2"
'    SetAddressPartControlSource frm, "BuyerSolicitor"
'    SetAddressPartControlSource frm, "BuyerAgent"
    
'    frm("LotDescription").ControlSource = "=ExtractLotAndOn([RPD],""Lot"")"
'    frm("OnDescription").ControlSource = "=ExtractLotAndOn([RPD],""On"")"
    
    Dim ctl As Control
    For Each ctl In frm.Controls
        If ctl.Name Like "cmd*" Then
            ''=Toggle_frmContractsPageButtons([Form],"Sale")
            Dim fieldName: fieldName = GetCmdFieldName(ctl.Name)
            ctl.OnClick = "=Toggle_frmContractsPageButtons([Form]," & Esc(fieldName) & ")"
        End If
    Next ctl
    
    CreateFormOptionGroups frm, "ExclusiveAgencyAgree"
    CreateFormCheckboxes frm, "OpenListing,SoleAgency,ExclusiveAgency"
    SetCustomFormDimensionControls frm
    
    frm.OnCurrent = "=" & frmName & "_OnCurrent([Form])"
    
End Function

Private Function SetCustomFormDimensionControls(frm As Form)
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM qryCustomFormDimensionControls WHERE FormName = " & Esc(frm.Name) & " ORDER BY ControlOrder,CustomFormDimensionControlID")
    Do Until rs.EOF
        Dim ControlName: ControlName = rs.fields("ControlName")
        frm(ControlName).top = rs.fields("Top")
        frm(ControlName).left = rs.fields("Left")
        frm(ControlName).width = rs.fields("Width")
        frm(ControlName).height = rs.fields("Height")
        rs.MoveNext
    Loop
    
End Function

Private Function GetCmdFieldName(ctlName As String) As String
    Dim fieldName As String
    Dim startIndex As Integer
    
    ' Find the index of "cmd" in the control name
    startIndex = InStr(1, ctlName, "cmd", vbTextCompare)
    
    ' If "cmd" is found, extract the substring after it
    If startIndex > 0 Then
        fieldName = Mid(ctlName, startIndex + 3) ' Adding 3 to skip "cmd"
    Else
        ' If "cmd" is not found, return empty string or handle it as required
        fieldName = ""
    End If
    
    ' Return the extracted field name
    GetCmdFieldName = fieldName
End Function


Public Function frmContractsPage2_OnCurrent(frm As Object)
    
    Dim isReportMode: isReportMode = Not frm.Name Like "frm*"
    
    If Not isReportMode Then
        SetComboboxRowSource frm, "EntityCategoryName = ""Buyer""", "BuyerID", True
        SetComboboxRowSource frm, "EntityCategoryName = ""Buyer""", "Buyer2ID", True
        SetComboboxRowSource frm, "EntityCategoryName = ""Contact"" AND ContactCategoryName = ""Solicitor""", "BuyerSolicitorID"
        SetComboboxRowSource frm, "EntityCategoryName = ""Contact""", "BuyerAgentID"
    End If
    
    Dim ctl As Control
    For Each ctl In frm.Controls
        If ctl.ControlType = acCommandButton Then
            Dim fieldName: fieldName = replace(ctl.Name, "cmd", "")
            SetToggleButtonCaption frm, fieldName, frm(fieldName)
        End If
    Next ctl
'    SetToggleButtonCaption frm, "BuiltOn", frm("BuiltOn")
'    SetToggleButtonCaption frm, "Vacant", frm("Vacant")
'    SetToggleButtonCaption frm, "LandSoldFreeL", frm("BuiltOn")
'    SetToggleButtonCaption frm, "BuiltOn", frm("BuiltOn")
    
End Function

Public Function frmForm6Page3_OnCurrent(frm As Object)
    
    Dim isReportMode: isReportMode = Not frm.Name Like "frm*"
    
    Dim ctl As Control
    For Each ctl In frm.Controls
        If ctl.ControlType = acCommandButton Then
            Dim fieldName: fieldName = replace(ctl.Name, "cmd", "")
            
            Dim rawName: rawName = fieldName
            If fieldName Like "*_0" Or fieldName Like "*_1" Then
                rawName = left(fieldName, Len(fieldName) - 2)
            End If
            
            SetToggleButtonCaption frm, fieldName, frm(rawName)
        End If
    Next ctl
'    SetToggleButtonCaption frm, "BuiltOn", frm("BuiltOn")
'    SetToggleButtonCaption frm, "Vacant", frm("Vacant")
'    SetToggleButtonCaption frm, "LandSoldFreeL", frm("BuiltOn")
'    SetToggleButtonCaption frm, "BuiltOn", frm("BuiltOn")
    
End Function

Public Function frmForm6Page6_OnCurrent(frm As Object)
    
    Dim isReportMode: isReportMode = Not frm.Name Like "frm*"
    
    SetComboboxRowSource frm, "EntityCategoryName = ""Seller""", "SellerID"
    SetComboboxRowSource frm, "EntityCategoryName = ""Seller""", "Seller2ID"
    
    SetComboboxRowSource frm, "EntityCategoryName = ""Seller""", "txtSellerID"
    SetComboboxRowSource frm, "EntityCategoryName = ""Seller""", "txtSeller2ID"
    
'    Dim ctl As Control
'    For Each ctl In frm.Controls
'        If ctl.ControlType = acCommandButton Then
'            Dim fieldName: fieldName = replace(ctl.Name, "cmd", "")
'            Dim rawName
'            If fieldName Like "*_0" Or fieldName Like "*_1" Then
'                rawName = left(fieldName, Len(fieldName) - 2)
'            End If
'
'            Dim fieldValue: fieldValue = frm(rawName)
'            If fieldName Like "*_0" And IsNull(fieldValue) Then
'                fieldValue = True
'            End If
'            SetToggleButtonCaption frm, fieldName, fieldValue
'        End If
'    Next ctl
'    SetToggleButtonCaption frm, "BuiltOn", frm("BuiltOn")
'    SetToggleButtonCaption frm, "Vacant", frm("Vacant")
'    SetToggleButtonCaption frm, "LandSoldFreeL", frm("BuiltOn")
'    SetToggleButtonCaption frm, "BuiltOn", frm("BuiltOn")
    
End Function

Public Function frmForm6Page1_OnCurrent(frm As Object)
    
    Dim isReportMode: isReportMode = Not frm.Name Like "frm*"
    
    SetComboboxRowSource frm, "EntityCategoryName = ""Seller""", "SellerID"
    SetComboboxRowSource frm, "EntityCategoryName = ""Seller""", "Seller2ID"
    
'    Dim ctl As Control
'    For Each ctl In frm.Controls
'        If ctl.ControlType = acCommandButton Then
'            Dim fieldName: fieldName = replace(ctl.Name, "cmd", "")
'            Dim rawName
'            If fieldName Like "*_0" Or fieldName Like "*_1" Then
'                rawName = left(fieldName, Len(fieldName) - 2)
'            End If
'
'            Dim fieldValue: fieldValue = frm(rawName)
'            If fieldName Like "*_0" And IsNull(fieldValue) Then
'                fieldValue = True
'            End If
'            SetToggleButtonCaption frm, fieldName, fieldValue
'        End If
'    Next ctl
'    SetToggleButtonCaption frm, "BuiltOn", frm("BuiltOn")
'    SetToggleButtonCaption frm, "Vacant", frm("Vacant")
'    SetToggleButtonCaption frm, "LandSoldFreeL", frm("BuiltOn")
'    SetToggleButtonCaption frm, "BuiltOn", frm("BuiltOn")
    
End Function

Public Function frmForm6Page5_OnCurrent(frm As Object)
    
    Dim isReportMode: isReportMode = Not frm.Name Like "frm*"
    
'    Dim ctl As Control
'    For Each ctl In frm.Controls
'        If ctl.ControlType = acCommandButton Then
'            Dim fieldName: fieldName = replace(ctl.Name, "cmd", "")
'            Dim rawName
'            If fieldName Like "*_0" Or fieldName Like "*_1" Then
'                rawName = left(fieldName, Len(fieldName) - 2)
'            End If
'
'            Dim fieldValue: fieldValue = frm(rawName)
'            If fieldName Like "*_0" And IsNull(fieldValue) Then
'                fieldValue = True
'            End If
'            SetToggleButtonCaption frm, fieldName, fieldValue
'        End If
'    Next ctl
'    SetToggleButtonCaption frm, "BuiltOn", frm("BuiltOn")
'    SetToggleButtonCaption frm, "Vacant", frm("Vacant")
'    SetToggleButtonCaption frm, "LandSoldFreeL", frm("BuiltOn")
'    SetToggleButtonCaption frm, "BuiltOn", frm("BuiltOn")
    
End Function

Public Function frmForm6Page4_OnCurrent(frm As Object)
    
    Dim isReportMode: isReportMode = Not frm.Name Like "frm*"
    
    Dim ctl As Control
    For Each ctl In frm.Controls
        If ctl.ControlType = acCommandButton Then
            Dim fieldName: fieldName = replace(ctl.Name, "cmd", "")
            Dim rawName
            If fieldName Like "*_0" Or fieldName Like "*_1" Then
                rawName = left(fieldName, Len(fieldName) - 2)
            End If
            
            Dim fieldValue: fieldValue = frm(rawName)
            If fieldName Like "*_0" And IsNull(fieldValue) Then
                fieldValue = True
            End If
            SetToggleButtonCaption frm, fieldName, fieldValue
        End If
    Next ctl
'    SetToggleButtonCaption frm, "BuiltOn", frm("BuiltOn")
'    SetToggleButtonCaption frm, "Vacant", frm("Vacant")
'    SetToggleButtonCaption frm, "LandSoldFreeL", frm("BuiltOn")
'    SetToggleButtonCaption frm, "BuiltOn", frm("BuiltOn")
    
End Function

Public Function frmForm6Page2_OnCurrent(frm As Object)
    
    Dim isReportMode: isReportMode = Not frm.Name Like "frm*"
    
'    If Not isReportMode Then
'        SetComboboxRowSource frm, "EntityCategoryName = ""Buyer""", "BuyerID", True
'        SetComboboxRowSource frm, "EntityCategoryName = ""Buyer""", "Buyer2ID", True
'        SetComboboxRowSource frm, "EntityCategoryName = ""Contact"" AND ContactCategoryName = ""Solicitor""", "BuyerSolicitorID"
'        SetComboboxRowSource frm, "EntityCategoryName = ""Contact""", "BuyerAgentID"
'    End If
    
    Dim ctl As Control
    For Each ctl In frm.Controls
        If ctl.ControlType = acCommandButton Then
            Dim fieldName: fieldName = replace(ctl.Name, "cmd", "")
            SetToggleButtonCaption frm, fieldName, frm(fieldName)
        End If
    Next ctl
'    SetToggleButtonCaption frm, "BuiltOn", frm("BuiltOn")
'    SetToggleButtonCaption frm, "Vacant", frm("Vacant")
'    SetToggleButtonCaption frm, "LandSoldFreeL", frm("BuiltOn")
'    SetToggleButtonCaption frm, "BuiltOn", frm("BuiltOn")
    
End Function

Public Function Setup_frmContractsPage3()
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "tblContracts"
          .fields = PrependDatabaseFields("ContractID,ExcludedFixtures,IncludedChattels,PurchasePrice,InitialDeposit,BalanceDeposit," & _
            "DefaultInterestRate,FinanceAmount,Financier,FinanceDate,InspectionDate,HasEncumbrance,Encumbrances,InitialDepositPayableOn," & _
            "BalancePayableOn", "tblContracts")
          sqlStr = .sql
    End With
    
    Dim frmName: frmName = "frmContractsPage3"
    DoCmd.OpenForm frmName, acDesign
    Dim frm As Form: Set frm = Forms(frmName)
    frm.RecordSource = sqlStr
    
    SetUpReport frmName, sqlStr
    
    ''Fix the controls -> Colicitors -> should be Solicitor
    Dim ctl As Control
    For Each ctl In frm.Controls
        If ctl.Name Like "*SellerSolicitor*" Then
            ctl.Name = replace(ctl.Name, "SellerSolicitor", "BuyerSolicitor")
            ctl.ControlSource = replace(ctl.ControlSource, "SellerSolicitor", "BuyerSolicitor")
        End If
    Next ctl
    
    ''Fix the controls -> use Seller2 fields -> Uncomment since this is already fixed.
    Dim i As Integer, ControlSource, NewControlSource
'    For i = 24 To 32
'        Set ctl = frm("Text" & i)
'        ControlSource = ctl.ControlSource
'        If Not ControlSource Like "Buyer2*" Then
'            NewControlSource = replace(ctl.ControlSource, "Buyer", "Buyer2")
'            ctl.ControlSource = NewControlSource
'            ctl.Name = NewControlSource
'        End If
'
'    Next i

'    Dim ctlName
'    For i = 34 To 43
'        ctlName = "Text" & i
'        If ControlExists(ctlName, frm) Then
'            Set ctl = frm(ctlName)
'            ControlSource = ctl.ControlSource
'            If ControlSource Like "Buyer2*" Then
'                NewControlSource = replace(ctl.ControlSource, "Buyer2", "BuyerAgent")
'                ctl.ControlSource = NewControlSource
'                ctl.Name = NewControlSource
'            End If
'        End If
'    Next i
    
    frm.OnCurrent = "=" & frmName & "_OnCurrent([Form])"
    
End Function

Public Function SetOptionGroupOnCurrent(frm As Form, fieldName)

    Dim fieldValue: fieldValue = frm(fieldName)
    Dim ctl As Control, rightValue
    
    If fieldValue Then
        rightValue = "-1"
    Else
        rightValue = "0"
    End If
    
    For Each ctl In frm.Controls
        If ctl.Name Like "cmd" & fieldName & "_*" Then
            If ctl.Tag <> rightValue Then
                ctl.Caption = " "
            Else
                ctl.Caption = ChrW$(10004)
            End If
        End If
    Next ctl
    
End Function

Public Function frmContractsPage3_OnCurrent(frm As Form, Optional reportMode As Boolean = False)
    
    SetOptionGroupOnCurrent frm, "HasEncumbrance"

End Function

Public Function Setup_frmContractsPage4()
    
    Dim entityFieldStr: entityFieldStr = GetEntityFields().arr(0)
    Dim entitySqlStr: entitySqlStr = GetEntityFields().arr(1)
    Dim memberFieldStr: memberFieldStr = GetEntityFields("Member").arr(0)
    Dim memberSqlStr: memberSqlStr = GetEntityFields("Member").arr(1)
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "tblContracts"
          .fields = PrependDatabaseFields("ContractID,TenantID,TermAndOptions,TermStartingDate,TermStartingEnd,Rent,Bond,AgencyName," & _
            "PropertyMangerID,HasPool,HasPoolComplianceCertificate,SafetySwitchNotified,SmokeAlaramNotified,PropertyListID", "tblContracts") & "," & _
            ProcessDatabaseFields(memberFieldStr, "tempPropertyManager", "PropertyManager")
          .Joins.Add GenerateJoinObj(memberSqlStr, "PropertyMangerID", "tempPropertyManager", "EntityID", "LEFT")
          sqlStr = .sql
    End With
    
    Dim frmName: frmName = "frmContractsPage4"
    DoCmd.OpenForm frmName, acDesign
    Dim frm As Form: Set frm = Forms(frmName)
    frm.RecordSource = sqlStr
    
    SetUpReport frmName, sqlStr
    
    ''Fix the controls -> Colicitors -> should be Solicitor
    Dim ctl As Control
    For Each ctl In frm.Controls
        If ctl.Name Like "*Buyer*" Then
            ctl.Name = replace(ctl.Name, "Buyer", "PropertyManager")
            ctl.ControlSource = replace(ctl.ControlSource, "Buyer", "PropertyManager")
        End If
    Next ctl
    
    ''Fix the controls -> use Seller2 fields -> Uncomment since this is already fixed.
    Dim i As Integer, ControlSource, NewControlSource
'    For i = 24 To 32
'        Set ctl = frm("Text" & i)
'        ControlSource = ctl.ControlSource
'        If Not ControlSource Like "Buyer2*" Then
'            NewControlSource = replace(ctl.ControlSource, "Buyer", "Buyer2")
'            ctl.ControlSource = NewControlSource
'            ctl.Name = NewControlSource
'        End If
'
'    Next i

'    Dim ctlName
'    For i = 34 To 43
'        ctlName = "Text" & i
'        If ControlExists(ctlName, frm) Then
'            Set ctl = frm(ctlName)
'            ControlSource = ctl.ControlSource
'            If ControlSource Like "Buyer2*" Then
'                NewControlSource = replace(ctl.ControlSource, "Buyer2", "BuyerAgent")
'                ctl.ControlSource = NewControlSource
'                ctl.Name = NewControlSource
'            End If
'        End If
'    Next i

   ''SetComboboxRowSource frm, "EntityCategoryName = ""Contact""", "PropertyMangerID"
    SetAddressPartControlSource frm, "PropertyManager"
    
    frm.OnCurrent = "=" & frmName & "_OnCurrent([Form])"
    
End Function

Public Function frmContractsPage4_OnCurrent(frm As Form, Optional reportMode As Boolean = False)
        
    SetComboboxRowSource frm, "EntityCategoryName = ""Contact""", "PropertyMangerID", True, True
    SetComboboxRowSource frm, "EntityCategoryName = ""Tenant""", "TenantID", , True
    
    SetOptionGroupOnCurrent frm, "HasPool"
    SetOptionGroupOnCurrent frm, "HasPoolComplianceCertificate"
    SetOptionGroupOnCurrent frm, "SafetySwitchNotified"
    SetOptionGroupOnCurrent frm, "SmokeAlaramNotified"
    
End Function

Public Function Setup_frmContractsPage5()
    
    Dim frmName: frmName = "frmContractsPage5"
    DoCmd.OpenForm frmName, acDesign
    Dim frm As Form: Set frm = Forms(frmName)
    frm.OnCurrent = "=" & frmName & "_OnCurrent([Form])"
    
    SetUpReport frmName, "tblContracts"
    
    Dim ctl As Control
    For Each ctl In frm.Controls
        Dim ctlName: ctlName = ctl.Name
        If ctlName Like "*_*" And ctl.ControlType = acCommandButton Then
            Dim fldName: fldName = RemoveChars(ctlName, 3, 2)
            
            ctl.Tag = IIf(ctlName Like "*_1", "-1", "0")
            ctl.OnClick = "=Toggle_frmContractsOptionGroup([Form]," & Esc(fldName) & "," & Esc("cmd" & fldName & "_" & IIf(ctl.Tag = "0", "0", "1")) & ")"
        End If
    Next ctl
    
End Function

Public Function frmContractsPage5_OnCurrent(frm As Form, Optional reportMode As Boolean = False)
    
    SetOptionGroupOnCurrent frm, "NeighborhoodDisputeNotified"
    SetOptionGroupOnCurrent frm, "BuyerGSTRegistered"
    SetOptionGroupOnCurrent frm, "WithholdingLawNotified"
    
End Function

Public Function Setup_frmContractsPage6()
    
'    Dim entityFieldStr: entityFieldStr = GetEntityFields().arr(0)
'    Dim entitySqlStr: entitySqlStr = GetEntityFields().arr(1)
'    Dim memberFieldStr: memberFieldStr = GetEntityFields("Member").arr(0)
'    Dim memberSqlStr: memberSqlStr = GetEntityFields("Member").arr(1)
'
'    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
'    Set sqlObj = New clsSQL
'    With sqlObj
'          .Source = "tblContracts"
'          .fields = PrependDatabaseFields("ContractID,SpecialConditions,SettlementDate,SettlementPlace,BuyerWitness,Buyer2Witness,SellerWitness,Seller2Witness", "tblContracts") & "," & _
'            ProcessDatabaseFields(memberFieldStr, "tempBuyer", "Buyer") & "," & _
'            ProcessDatabaseFields(memberFieldStr, "tempBuyer2", "Buyer2") & "," & _
'            ProcessDatabaseFields(entityFieldStr, "tempSeller", "Seller") & "," & _
'            ProcessDatabaseFields(entityFieldStr, "tempSeller2", "Seller2")
'          .Joins.Add GenerateJoinObj(memberSqlStr, "BuyerID", "tempBuyer", "EntityID", "LEFT")
'          .Joins.Add GenerateJoinObj(memberSqlStr, "Buyer2ID", "tempBuyer2", "EntityID", "LEFT")
'          .Joins.Add GenerateJoinObj(entitySqlStr, "SellerID", "tempSeller", "EntityID", "LEFT")
'          .Joins.Add GenerateJoinObj(entitySqlStr, "Seller2ID", "tempSeller2", "EntityID", "LEFT")
          Dim sqlStr: sqlStr = "tblContracts"
'    End With
    
    Dim frmName: frmName = "frmContractsPage6"
    DoCmd.OpenForm frmName, acDesign
    Dim frm As Form: Set frm = Forms(frmName)
    frm.RecordSource = sqlStr
    
    SetUpReport frmName, sqlStr
    
    ''Fix the controls -> Colicitors -> should be Solicitor
'    Dim ctl As Control
'    For Each ctl In frm.Controls
'        If ctl.Name Like "*Buyer*" Then
'            ctl.Name = replace(ctl.Name, "Buyer", "PropertyManager")
'            ctl.ControlSource = replace(ctl.ControlSource, "Buyer", "PropertyManager")
'        End If
'    Next ctl
    
    ''Fix the controls -> use Seller2 fields -> Uncomment since this is already fixed.
    Dim i As Integer, ControlSource, NewControlSource
'    For i = 24 To 32
'        Set ctl = frm("Text" & i)
'        ControlSource = ctl.ControlSource
'        If Not ControlSource Like "Buyer2*" Then
'            NewControlSource = replace(ctl.ControlSource, "Buyer", "Buyer2")
'            ctl.ControlSource = NewControlSource
'            ctl.Name = NewControlSource
'        End If
'
'    Next i

'    Dim ctlName
'    For i = 34 To 43
'        ctlName = "Text" & i
'        If ControlExists(ctlName, frm) Then
'            Set ctl = frm(ctlName)
'            ControlSource = ctl.ControlSource
'            If ControlSource Like "Buyer2*" Then
'                NewControlSource = replace(ctl.ControlSource, "Buyer2", "BuyerAgent")
'                ctl.ControlSource = NewControlSource
'                ctl.Name = NewControlSource
'            End If
'        End If
'    Next i
    
'    frm.OnCurrent = "=" & frmName & "_OnCurrent([Form])"
    
End Function



