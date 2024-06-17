Attribute VB_Name = "PropertyEmailRecipient Mod"
Option Compare Database
Option Explicit

Public Function PropertyEmailRecipientCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function FilterEmailRecipients(frm As Form)
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblEmailRecipientFieldFilters WHERE IsChecked")
    
    Dim filterArr As New clsArray
    
    Dim fltrWildSearch: fltrWildSearch = frm("fltrWildSearch")
    
    If Not isFalse(fltrWildSearch) Then
        filterArr.Add "EntityName Like " & Esc("*" & fltrWildSearch & "*") & " OR StreetAddress Like " & Esc("*" & fltrWildSearch & "*")
    End If
    
    Do Until rs.EOF
        Dim FilterField, FilterCaption
        FilterField = rs.fields("FilterField")
        FilterCaption = rs.fields("FilterCaption")
        filterArr.Add FilterField & " = " & EscapeString(FilterCaption)
        rs.MoveNext
    Loop
    
    If filterArr.Count = 0 Then
        frm("subform").Form.FilterOn = False
        Exit Function
    End If
    
    frm("subform").Form.Filter = filterArr.JoinArr(" AND ")
    frm("subform").Form.FilterOn = True
    
End Function

Public Function ClearFilterEmailRecipients()
    
    Dim frm As Form, ctl As Control, fieldArr As New clsArray, fieldItem
    If IsFormOpen("mainPropertyEmailRecipients") Then
        Set frm = Forms("mainPropertyEmailRecipients")
        
        frm("fltrWildSearch") = Null
        RunSQL "UPDATE tblEmailRecipientFieldFilters SET IsChecked = 0"
        
        fieldArr.arr = "EntityCategoryName,ContactCategoryName,Association,BuyerStatus"
        
        For Each fieldItem In fieldArr.arr
            frm("fltr" & fieldItem).Form.Requery
        Next fieldItem
        
        frm("subform").Form.FilterOn = False
    End If
    
    If IsFormOpen("mainBulkEmailRefineFilter") Then
        Set frm = Forms("mainBulkEmailRefineFilter")
        For Each ctl In frm.Controls
            If ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Then ctl.Value = Null
            If ctl.ControlType = acCheckBox Then ctl.Value = 0
        Next ctl
    End If
    
    If IsFormOpen("mainPropertyBulkEmail") Then
        Set frm = Forms("mainPropertyBulkEmail")
        
        frm("fltrWildSearch") = Null
        RunSQL "UPDATE tblEmailRecipientFieldFilters SET IsChecked = 0"
        
         
        fieldArr.arr = "EntityCategoryName,ContactCategoryName,Association,BuyerStatus"
        
        For Each fieldItem In fieldArr.arr
            frm("fltr" & fieldItem).Form.Requery
        Next fieldItem
        
        frm("subform").Form.FilterOn = False
    End If
    
End Function

Public Function BuildFromSuccessfulBuyer(PropertyListID, fieldName, Caption, PartOnly) As String
    
    Dim txtSuccessfulBuyer, frm As Form
    Set frm = Forms("frmPropertyList")
    txtSuccessfulBuyer = frm("txtSuccessfulBuyer").Column(1)
    
    If isFalse(txtSuccessfulBuyer) Then
        txtSuccessfulBuyer = "No successful buyer."
        Exit Function
    End If
    
    If PartOnly = "Firstname" Then txtSuccessfulBuyer = GetSurname(txtSuccessfulBuyer, True)
    If PartOnly = "Surname" Then txtSuccessfulBuyer = GetSurname(txtSuccessfulBuyer)
    
    BuildFromSuccessfulBuyer = txtSuccessfulBuyer

End Function

Public Function BuildFromPropertyField(PropertyListID, fieldName, Caption, PartOnly, Optional PropertyEmailRecipientID = Null) As String
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblPropertyList WHERE PropertyListID = " & PropertyListID)
    
    If rs.EOF Then Exit Function
    
    Dim PropertyArr As New clsArray
    PropertyArr.Add rs.fields(fieldName)
    
    If PartOnly = "Type" Then
        PropertyArr.Add rs.fields("Suburb")
    End If
    
    Dim StreetAddress: StreetAddress = GetPropertyAddress(PropertyListID)
    
    BuildFromPropertyField = StreetAddress
'    BuildFromPropertyField = PropertyArr.JoinArr(" ")
    
End Function

Public Function BuildEntityMembers(PropertyListID, EntityCategoryName, Caption, PartOnly, Optional PropertyEmailRecipientID = Null) As String
    
    If PropertyListID = 0 Then Exit Function
    BuildEntityMembers = "No " & EntityCategoryName
    Dim SuccessfulBuyerMode As Boolean
    SuccessfulBuyerMode = False
    
    If EntityCategoryName = "Successful" Then
        SuccessfulBuyerMode = True
        EntityCategoryName = "Buyer"
    End If
    
    Dim rs As Recordset, sqlStr
    sqlStr = "SELECT * FROM qryPropertyEntityMembers WHERE PropertyListID = " & PropertyListID & " AND EntityCategoryName = " & EscapeString(EntityCategoryName)
    If SuccessfulBuyerMode Then
    
        Dim frm As Form, txtSuccessfulBuyer
        Set frm = Forms("frmPropertyList")
        txtSuccessfulBuyer = frm("txtSuccessfulBuyer")
        
        If isFalse(txtSuccessfulBuyer) Then Exit Function
        
        sqlStr = sqlStr & " AND EntityID = " & txtSuccessfulBuyer
        
    End If
    
    Set rs = ReturnRecordset(sqlStr)
    
    Dim MemberName, MemberArr As New clsArray
    
    If rs.EOF Then Exit Function
    
    Dim isFirst As Boolean: isFirst = False
    
    Do Until rs.EOF
        
        ''Questionable Why this is needed...
'        If isFirst Then
'            MemberName = rs.fields("EntityName")
'            isFirst = False
'        Else
'            MemberName = rs.fields("MemberName")
'        End If
        MemberName = rs.fields("MemberName")
        
        Dim MemberNameToBeAdded
        If PartOnly = "Firstname" Then MemberNameToBeAdded = GetSurname(MemberName, True)
        If PartOnly = "Surname" Then MemberNameToBeAdded = GetSurname(MemberName)
        If isFalse(PartOnly) Then MemberNameToBeAdded = MemberName
        
        Dim ConvertedName: ConvertedName = ELookup("tblSurnameConversions", "Surname = " & Esc(MemberNameToBeAdded), "ConvertedSurname")
        If Not isFalse(ConvertedName) Then MemberNameToBeAdded = ConvertedName
        MemberArr.Add MemberNameToBeAdded
'        ''Add first name, full name, surname logic here
'        If PartOnly = "Firstname" Then MemberArr.Add GetSurname(MemberName, True)
'        If PartOnly = "Surname" Then MemberArr.Add GetSurname(MemberName)
'        If isFalse(PartOnly) Then MemberArr.Add MemberName
        rs.MoveNext
        
    Loop
    
    BuildEntityMembers = JoinArrayProperly(MemberArr)
    
End Function

Public Function BuildEmailSellers(PropertyListID, fieldName, Caption, PartOnly, Optional PropertyEmailRecipientID = Null)
    
    BuildEmailSellers = "No Seller"
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM qryPropertySellers WHERE PropertyListID = " & PropertyListID)
    
    
    If rs.EOF Then Exit Function
    Dim EntityName, EntityArr As New clsArray
    
    Do Until rs.EOF
        EntityName = rs.fields("EntityName")
        
        If PartOnly = "Firstname" Then EntityArr.Add GetSurname(EntityName, True)
        If PartOnly = "Surname" Then EntityArr.Add GetSurname(EntityName)
        If isFalse(PartOnly) Then EntityArr.Add EntityName
        
        rs.MoveNext
    Loop
    
    BuildEmailSellers = JoinArrayProperly(EntityArr)
    
End Function

Public Function BuildAdvertisementLink(PropertyListID)
    
    BuildAdvertisementLink = "Advertisement ID Not Existing."
    
    Dim AdvertisementID: AdvertisementID = ELookup("tblPropertyList", "PropertyListID = " & PropertyListID, "AdvertisementID")
    
    If isFalse(AdvertisementID) Then Exit Function

    BuildAdvertisementLink = "https://www.realestate.com.au/" & Trim(AdvertisementID)
    
End Function

Public Function BuildEmailRecipients(PropertyListID, fieldName, Caption, PartOnly, Optional EmailRecipientID = Null)
    
    BuildEmailRecipients = "No Recipient"
    
    Dim rs As Recordset
    Dim sqlStr: sqlStr = "SELECT * FROM tblPropertyEmailRecipients WHERE "
    If isFalse(EmailRecipientID) Then
        sqlStr = sqlStr & "Recipient"
    Else
        sqlStr = sqlStr & "PropertyEmailRecipientID = " & EmailRecipientID
    End If
    Set rs = ReturnRecordset(sqlStr & " ORDER BY EntityName")
    
    
    If rs.EOF Then Exit Function
    Dim EntityName, EntityArr As New clsArray
    
    Do Until rs.EOF
        EntityName = rs.fields("EntityName")
        
        If PartOnly = "Firstname" Then EntityArr.Add GetSurname(EntityName, True)
        If PartOnly = "Surname" Then EntityArr.Add GetSurname(EntityName)
        If isFalse(PartOnly) Then EntityArr.Add EntityName
        
        rs.MoveNext
    Loop
    
    BuildEmailRecipients = JoinArrayProperly(EntityArr)
    
End Function

Private Function JoinArrayProperly(arr As clsArray)

    Dim arrCount, i, arrItem, subArr As New clsArray
    arrCount = arr.Count
    
    If arrCount = 1 Then
        JoinArrayProperly = arr.JoinArr("")
        Exit Function
    End If
    
    i = 0
    
    For Each arrItem In arr.arr
        
        arrItem = StrConv(arrItem, vbProperCase)
        If i = arrCount - 1 Then
            JoinArrayProperly = subArr.JoinArr(",") & " and " & arrItem
            Exit Function
        Else
            subArr.Add arrItem
        End If
        
        i = i + 1
        
    Next arrItem
    
End Function

Public Function GetSurname(fullName, Optional FirstNameOnly As Boolean = False) As String
    
    Dim NonHumanArr As New clsArray, NonHuman
    NonHumanArr.arr = "Corporate,LTD,PTY,Law"
    
    ' Convert the full name to uppercase for case-insensitive comparison
    fullName = UCase(fullName)
    
    For Each NonHuman In NonHumanArr.arr
        If InStr(1, fullName, " " & NonHuman & " ") > 0 Then
            GetSurname = fullName
            Exit Function
        End If
    Next NonHuman
    
    Dim str As Variant, strItem
    str = Split(fullName, " ")
    
    If FirstNameOnly Then
        If UBound(str) = 0 Then
            GetSurname = fullName
        Else
            GetSurname = str(0)
        End If
    Else
        GetSurname = str(UBound(str))
    End If
    
    GetSurname = StrConv(GetSurname, vbProperCase)
    
End Function


Public Function OpenMainPropertyBulkEmail(frm As Form)
    
    DoCmd.RunCommand acCmdSaveRecord
    
    Dim PropertyListID
    
    If DoesPropertyExists(frm, "PropertyListID") Then PropertyListID = frm("PropertyListID")
    
    ''If ExitIfTrue(isFalse(PropertyListID), "Property is empty.") Then Exit Function
    
    ''tblPropertyEmailRecipients -> remove all data first

    RunSQL "DELETE * from tblPropertyEmailRecipients"
    
    ''Create a sublist of the property entities where PropertyListID = PropertyListID and
    ''EmailAddress is not null
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    ''Also do the buyer members here
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "qryPropertyEntityMembers"
        .AddFilter "NOT MemberEmailAddress IS NULL AND EntityCategoryName = ""Buyer"""
        .fields = "PropertyEntityID,MemberName As EntityName,EntityCategoryName,MemberEmailAddress AS EmailAddress,Association,ContactCategoryName,BuyerStatus,StreetAddress"
        .OrderBy = "EntityCategoryID,MemberName"
        sqlStr = .sql
        ''makeQuery sqlStr
    End With
    
'    Set sqlObj = New clsSQL
'    With sqlObj
'        .Source = "qryPropertyEntities"
'        .AddFilter "NOT MemberEmailAddress IS NULL AND EntityCategoryName = ""Buyer"""
'        .fields = "PropertyEntityID,MemberName As EntityName,EntityCategoryName,MemberEmailAddress AS EmailAddress,Association,ContactCategoryName,BuyerStatus"
'        .OrderBy = "EntityCategoryID, MemberName"
'        sqlStr = .sql
'        ''makeQuery sqlStr
'    End With
    
    ''Then insert the sublist
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "INSERT"
        .Source = "tblPropertyEmailRecipients"
        .fields = "PropertyEntityID,EntityName,EntityCategoryName,EmailAddress,Association,ContactCategoryName,BuyerStatus,StreetAddress"
        .InsertSQL = sqlStr
        .InsertUseAsPlain = True
        rowsAffected = .Run
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "qryPropertyEntityMembers"
        .AddFilter "NOT MemberEmailAddress IS NULL AND EntityCategoryName = ""Contact"""
        .fields = "PropertyEntityID,MemberName As EntityName,EntityCategoryName,MemberEmailAddress AS EmailAddress,Association,ContactCategoryName,BuyerStatus,StreetAddress"
        .OrderBy = "EntityCategoryID, MemberName"
        sqlStr = .sql
        ''makeQuery sqlStr
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "INSERT"
        .Source = "tblPropertyEmailRecipients"
        .fields = "PropertyEntityID,EntityName,EntityCategoryName,EmailAddress,Association,ContactCategoryName,BuyerStatus,StreetAddress"
        .InsertSQL = sqlStr
        .InsertUseAsPlain = True
        rowsAffected = .Run
    End With
    
    ''Select the sellers
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "qryPropertyEntities"
        .AddFilter "NOT EmailAddress IS NULL AND EntityCategoryName = ""Seller"""
        .fields = "PropertyEntityID,EntityName,EntityCategoryName,EmailAddress,Association,ContactCategoryName,BuyerStatus,StreetAddress"
        .OrderBy = "EntityCategoryID, EntityName"
        sqlStr = .sql
        ''makeQuery sqlStr
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "INSERT"
        .Source = "tblPropertyEmailRecipients"
        .fields = "PropertyEntityID,EntityName,EntityCategoryName,EmailAddress,Association,ContactCategoryName,BuyerStatus,StreetAddress"
        .InsertSQL = sqlStr
        .InsertUseAsPlain = True
        rowsAffected = .Run
    End With

    Dim recordCount
    recordCount = ECount("tblPropertyEmailRecipients", "PropertyEmailRecipientID > 0")
    If ExitIfTrue(recordCount = 0, "There's no recipients found..") Then Exit Function
    
    ''Build the files here (collect all the files related to the property)
    RunSQL "DELETE * from tblEmailAttachments"
    If Not isFalse(PropertyListID) Then BuildProperyListFiles PropertyListID
    
    BuildEmailRecipientFieldFilters
    
    DoCmd.OpenForm "mainPropertyBulkEmail"
    ''ResizeFilterSubforms "mainPropertyBulkEmail"
    
    Forms("mainPropertyBulkEmail")("subform").Form.Requery
    Forms("mainPropertyBulkEmail")("cmdSendBulk").OnClick = "=ComposeMailIndividualRecipients([Form]" & IIf(Not isFalse(PropertyListID), "," & PropertyListID, ",0") & ")"
    Forms("mainPropertyBulkEmail")("cmdPreviewEmail").OnClick = "=PreviewIndividualRecipients([Form]" & IIf(Not isFalse(PropertyListID), "," & PropertyListID, ",0") & ")"

    Forms("mainPropertyBulkEmail")("subEmailAttachments").Form.Requery
    
End Function

Public Function OpenMainPropertyEmailRecipient(frm As Form)
    
    Dim PropertyListID
    PropertyListID = frm("PropertyListID")
    
    If ExitIfTrue(isFalse(PropertyListID), "Property is empty.") Then Exit Function
    
    ''tblPropertyEmailRecipients -> remove all data first

    RunSQL "DELETE * from tblPropertyEmailRecipients"
    
    ''Create a sublist of the property entities where PropertyListID = PropertyListID and
    ''EmailAddress is not null
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    ''Also do the buyer members here
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "qryPropertyEntityMembers"
        .AddFilter "PropertyListID = " & PropertyListID & " AND NOT MemberEmailAddress IS NULL"
        .fields = "MemberName As EntityName,EntityCategoryName,MemberEmailAddress AS EmailAddress,Association,ContactCategoryName,BuyerStatus,StreetAddress"
        .OrderBy = "EntityCategoryID, MemberName"
        sqlStr = .sql
        ''makeQuery sqlStr
    End With
    
    ''Then insert the sublist
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "INSERT"
        .Source = "tblPropertyEmailRecipients"
        .fields = "EntityName,EntityCategoryName,EmailAddress,Association,ContactCategoryName,BuyerStatus,StreetAddress"
        .InsertSQL = sqlStr
        .InsertUseAsPlain = True
        rowsAffected = .Run
    End With
    
    ''Property Entities First
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "qryPropertyEntities"
        .AddFilter "PropertyListID = " & PropertyListID & " AND NOT EmailAddress IS NULL AND EntityCategoryName = 'Seller'"
        .fields = "EntityName,EntityCategoryName,EmailAddress,Association,ContactCategoryName,BuyerStatus,StreetAddress"
        .OrderBy = "EntityCategoryID, EntityName"
        sqlStr = .sql
    End With

    ''Then insert the sublist
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "INSERT"
        .Source = "tblPropertyEmailRecipients"
        .fields = "EntityName,EntityCategoryName,EmailAddress,Association,ContactCategoryName,BuyerStatus,StreetAddress"
        .InsertSQL = sqlStr
        .InsertUseAsPlain = True
        rowsAffected = .Run
    End With

    Dim recordCount
    recordCount = ECount("tblPropertyEmailRecipients", "PropertyEmailRecipientID > 0")
    If ExitIfTrue(recordCount = 0, "There's no recipients found..") Then Exit Function
    
    ''Build the files here (collect all the files related to the property)
    RunSQL "DELETE FROM tblEmailAttachments"
    BuildProperyListFiles PropertyListID
    
    BuildEmailRecipientFieldFilters
    
    DoCmd.OpenForm "mainPropertyEmailRecipients"
    ResizeFilterSubforms
    
    Forms("mainPropertyEmailRecipients")("subform").Form.Requery
    Forms("mainPropertyEmailRecipients")("cmdComposeMail").OnClick = "=ComposeMailMultipleRecipients([Form]," & PropertyListID & ")"
    Forms("mainPropertyEmailRecipients")("subEmailAttachments").Form.Requery
    
End Function

Private Function ResizeFilterSubforms(Optional frmName As String = "mainPropertyEmailRecipients")

    ''GetTotalRecordCount
    
    Dim frm As Form
    Set frm = Forms(frmName)
    
    Dim fieldArr As New clsArray, fieldItem
    fieldArr.arr = "Association,EntityCategoryName,ContactCategoryName,BuyerStatus"
    
    Dim lblHeight, lblTop, lblBottom, fltrHeight
    
    For Each fieldItem In fieldArr.arr
        
        Dim rs As Recordset, SizeInInch
        Set rs = frm("fltr" & fieldItem).Form.RecordsetClone
        SizeInInch = GetTotalRecordCount(rs) * 0.25
'        frm("fltr" & fieldItem).height = InchToTwip(SizeInInch)
        If SizeInInch > 1 Then
            SizeInInch = 1
            frm("fltr" & fieldItem).Form.ScrollBars = 2
        Else
            frm("fltr" & fieldItem).Form.ScrollBars = 0
        End If
        
        
        frm("fltr" & fieldItem).height = InchToTwip(SizeInInch)
        
    Next fieldItem
    
    For Each fieldItem In fieldArr.arr
        
        fltrHeight = frm("fltr" & fieldItem).height
        If fltrHeight = 0 Then
            frm("lbl" & fieldItem).height = 0
        End If
        
        lblHeight = frm("lbl" & fieldItem).height
        
        lblTop = frm("lbl" & fieldItem).top
        lblBottom = lblHeight + lblTop
        
        frm("fltr" & fieldItem).top = lblBottom
        
    Next fieldItem
    
    fieldArr.arr = "Cc,Bcc,To"
    For Each fieldItem In fieldArr.arr
        
        frm("subform").Form(fieldItem).ColumnHidden = frmName <> "mainPropertyEmailRecipients"
        
    Next fieldItem
    
End Function

Private Function BuildEmailRecipientFieldFilters()
    
    RunSQL "DELETE FROM tblEmailRecipientFieldFilters"
    
    Dim fieldArr As New clsArray, fieldItem
    fieldArr.arr = "EntityCategoryName,ContactCategoryName,Association,BuyerStatus"
    
    For Each fieldItem In fieldArr.arr
        BuildEmailRecipientFieldFilter fieldItem
    Next fieldItem
    
End Function


Private Function BuildEmailRecipientFieldFilter(fieldName)
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblPropertyEmailRecipients"
        .AddFilter "NOT " & fieldName & " IS NULL"
        .fields = fieldName & " AS FilterCaption," & EscapeString(fieldName) & " AS FilterField"
        .OrderBy = fieldName
        sqlStr = .sql
    End With
    
    InsertTotblEmailRecipientFieldFilters sqlStr
    
End Function

Private Function InsertTotblEmailRecipientFieldFilters(sqlStr)
    
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, rowsAffected, rs As Recordset
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "INSERT"
        .Source = "tblEmailRecipientFieldFilters"
        .fields = "FilterCaption,FilterField"
        .InsertSQL = sqlStr
        .InsertUseAsPlain = True
        rowsAffected = .Run
    End With
    
End Function

Private Function BuildProperyListFiles(PropertyListID)
    
    ''Delete the files first
    ''qryPropertyEntityFiles, tblEmailAttachments
    Dim StreetAddress, uploadDirectory
    StreetAddress = ELookup("tblPropertyList", "PropertyListID = " & PropertyListID, "StreetAddress")
    uploadDirectory = GetAttachmentsDirectory(StreetAddress)
            
    ''Do a loop here instead of a query
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM qryPropertyEntityFiles WHERE PropertyListID = " & PropertyListID & " ORDER BY EntityCategoryName, EntityName")
    
    Do Until rs.EOF
        
        Dim EntityFileLink, AbsoluteLink, EntityName, EntityCategoryName, vTimestamp
        EntityFileLink = rs.fields("EntityFileLink")
        EntityName = rs.fields("EntityName")
        EntityCategoryName = rs.fields("EntityCategoryName")
        vTimestamp = rs.fields("Timestamp")
        
        ''AbsoluteLink is the StreetAddress use this function ---> GetAttachmentsDirectory(StreetAddress) + FileLink
        AbsoluteLink = uploadDirectory & EntityFileLink
        
        RunSQL "INSERT INTO tblEmailAttachments (FileLink,AbsoluteLink,EntityName,EntityCategoryName,[Timestamp]) VALUES (" & EscapeString(EntityFileLink) & "," & EscapeString(AbsoluteLink) & "," & EscapeString(EntityName) & "," & _
            EscapeString(EntityCategoryName) & ",#" & vTimestamp & "#)"
        rs.MoveNext
    Loop
End Function

Public Function SelectAllEmailRecepient(frm As Form, Optional EmailPart = "To")
    
    Dim rs As Recordset
    Set rs = frm("subform").Form.RecordsetClone
    
    rs.MoveFirst
    Do Until rs.EOF
        Dim AddRecepient
        AddRecepient = rs.fields(EmailPart)
        
        rs.Edit
        rs.fields(EmailPart) = True
        
        rs.Update
        
        rs.MoveNext
    Loop
    
    
End Function

Public Function ComposeMailMultipleRecipients(frm As Form, PropertyListID)
    
    Dim rs As Recordset
    frm("subform").Form.Requery
    
    Set rs = ReturnRecordset("SELECT * FROM tblPropertyEmailRecipients ORDER BY PropertyEmailRecipientID")
    
    If ExitIfTrue(rs.EOF, "There's no recipients found..") Then Exit Function
    
    rs.MoveFirst
    
    Dim ToArr As New clsArray, BccArr As New clsArray, ccArr As New clsArray
    Dim EmailAddress, vTo, Bcc, Cc
    
    Do Until rs.EOF
        EmailAddress = rs.fields("EmailAddress")
        vTo = rs.fields("To")
        Bcc = rs.fields("Bcc")
        Cc = rs.fields("Cc")
        
        If vTo Then ToArr.Add EmailAddress, True
        If Bcc Then BccArr.Add EmailAddress, True
        If Cc Then ccArr.Add EmailAddress, True
        
        rs.MoveNext
    Loop
    
    If ExitIfTrue(ToArr.Count = 0, "Email address is empty...") Then Exit Function
    
    Dim olApp As Object, olMail As Object, olRecipient, Recipient
    
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)
    With olMail
        
        .Display
        
'        For Each Recipient In ToArr.arr
'
'            Set olRecipient = .Recipients.Add(Recipient)
'            olRecipient.Type = 1
'
'        Next Recipient
        
        .To = ToArr.JoinArr(";")
        
        If ccArr.Count > 0 Then
            .Cc = ccArr.JoinArr(";")
        End If
        
        If BccArr.Count > 0 Then
            .Bcc = BccArr.JoinArr(";")
        End If
        
        ''.Attachments.Add "C:\Users\user\Desktop\Databases\Access Files\Property Management\tblCards.csv"
        AttachFilesToEmail olMail
        SetEmailSubjectAndBody frm, olMail, PropertyListID
        
        '.SentOnBehalfOfName = "jet_pradas@yahoo.com"
        '.Subject = Subject
        '.HTMLBody = htmlTxt
        '.Send
    End With
    
End Function

Public Function PreviewIndividualRecipients(frm As Form, PropertyListID)
    
    Dim rs As Recordset
    DoCmd.RunCommand acCmdSaveRecord
    ''frm("subform").Form.Requery
    
    Dim Recipient: Recipient = frm("subform").Form("Recipient")
    If ExitIfTrue(Not Recipient, "Please toggle the recipient checkbox of the row to preview.") Then Exit Function
    
    Dim PropertyEmailRecipientID: PropertyEmailRecipientID = frm("subform").Form("PropertyEmailRecipientID")
    
    Set rs = ReturnRecordset("SELECT * FROM tblPropertyEmailRecipients WHERE PropertyEmailRecipientID = " & PropertyEmailRecipientID)
    
    If ExitIfTrue(rs.EOF, "There's no recipients found..") Then Exit Function
    
    rs.MoveFirst
    
    Dim ToArr As New clsArray, BccArr As New clsArray, ccArr As New clsArray
    Dim EmailAddress, vTo, Bcc, Cc
    
    Dim olApp As Object: Set olApp = CreateObject("Outlook.Application")
    
    
    
    Do Until rs.EOF
        EmailAddress = rs.fields("EmailAddress")
        PropertyEmailRecipientID = rs.fields("PropertyEmailRecipientID")
        
        If vTo Then ToArr.Add EmailAddress, True
        Dim olMail As Object, olRecipient
        
        Set olMail = olApp.CreateItem(0)
        With olMail
            
            .Display
            .To = EmailAddress
            
            ''.Attachments.Add "C:\Users\user\Desktop\Databases\Access Files\Property Management\tblCards.csv"
            AttachFilesToEmail olMail
            SetEmailSubjectAndBody frm, olMail, PropertyListID, PropertyEmailRecipientID
            
            '.SentOnBehalfOfName = "jet_pradas@yahoo.com"
            '.Subject = Subject
            '.HTMLBody = htmlTxt
            '.Send
        End With
        
        rs.MoveNext
    Loop
    
    
End Function


Public Function ComposeMailIndividualRecipients(frm As Form, Optional PropertyListID = "")
    
    Dim rs As Recordset
    DoCmd.RunCommand acCmdSaveRecord
    
    Dim resp: resp = MsgBox("This will send the email to all recipient directly. Do you want to proceed?", vbYesNo)
    If resp = vbNo Then Exit Function
    
    Set rs = ReturnRecordset("SELECT * FROM tblPropertyEmailRecipients WHERE Recipient ORDER BY PropertyEmailRecipientID")
    
    If ExitIfTrue(rs.EOF, "There's no recipients found..") Then Exit Function
    
    rs.MoveFirst
    
    Dim ToArr As New clsArray, BccArr As New clsArray, ccArr As New clsArray
    Dim EmailAddress, vTo, Bcc, Cc, PropertyEmailRecipientID
    
    Dim olApp As Object: Set olApp = CreateObject("Outlook.Application")
    
    
    Do Until rs.EOF
        EmailAddress = rs.fields("EmailAddress")
        PropertyEmailRecipientID = rs.fields("PropertyEmailRecipientID")
        
        If vTo Then ToArr.Add EmailAddress, True
        Dim olMail As Object, olRecipient, Recipient
        
        Set olMail = olApp.CreateItem(0)
        With olMail
            
            ''.Display
            .To = EmailAddress
            
            ''.Attachments.Add "C:\Users\user\Desktop\Databases\Access Files\Property Management\tblCards.csv"
            AttachFilesToEmail olMail
            SetEmailSubjectAndBody frm, olMail, PropertyListID, PropertyEmailRecipientID
            
            '.SentOnBehalfOfName = "jet_pradas@yahoo.com"
            '.Subject = Subject
            '.HTMLBody = htmlTxt
            .send
        End With
        
        rs.MoveNext
    Loop
    
    
End Function

Private Function SetEmailSubjectAndBody(frm As Form, olMail As Object, PropertyListID, Optional PropertyEmailRecipientID = Null)
    
    Dim txtEmailTemplate
    txtEmailTemplate = frm("txtEmailTemplate")
   
    
    If isFalse(txtEmailTemplate) Then Exit Function
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblEmailTemplates WHERE EmailTemplateID = " & txtEmailTemplate)
    
    Dim EmailTemplate, Subject, Body
    
    EmailTemplate = rs.fields("EmailTemplate")
    Subject = rs.fields("Subject")
    Body = rs.fields("Body")
    
    If Not isFalse(Subject) Then
        GetEmailPart Subject, PropertyListID, PropertyEmailRecipientID
        olMail.Subject = Subject
    End If
    
    If Not isFalse(Body) Then
        GetEmailPart Body, PropertyListID, PropertyEmailRecipientID
        GetEmailSignature Body, frm
        olMail.Body = Body
    End If
    
    
End Function

Private Function GetEmailRecipientList(Subject)

   
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblPropertyEmailRecipients WHERE Recipient ORDER BY EntityName")
    
    If rs.EOF Then Exit Function
    
    Dim EntityName, EntityArr As New clsArray
    
    Do Until rs.EOF
        EntityName = rs.fields("EntityName")
        EntityArr.Add GetSurname(EntityName, True)
        rs.MoveNext
    Loop
    
    If EntityArr.Count = 0 Then Exit Function
    
    Subject = Subject & " " & JoinArrayProperly(EntityArr)
    
    
End Function

Private Function GetEmailSignature(Body, frm As Form)
    
    Dim txtEmailSignature
    txtEmailSignature = frm("txtEmailSignature")
    
    If isFalse(txtEmailSignature) Then Exit Function
    
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblEmailSignatures WHERE EmailSignatureID = " & txtEmailSignature)
    
    If rs.EOF Then Exit Function
    
    Dim EmailSignatureBody
    EmailSignatureBody = rs.fields("EmailSignatureBody")
    
    If isFalse(EmailSignatureBody) Then Exit Function
    
    Body = Body & vbCrLf & vbCrLf & replace(EmailSignatureBody, vbCrLf, "<br/>")
    
End Function

Private Function GetEmailPart(EmailPart, PropertyListID, Optional PropertyEmailRecipientID = Null)
    
    EmailPart = Trim(EmailPart)
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblEmailUsableFields ORDER BY FieldOrder")
    
    Dim EmailUsableFieldID, FieldCaption, fieldName, FunctionToCall, FieldOrder, fieldValue, PartOnly
    Do Until rs.EOF
        
        EmailUsableFieldID = rs.fields("EmailUsableFieldID")
        FieldCaption = rs.fields("FieldCaption")
        fieldName = rs.fields("FieldName")
        FunctionToCall = rs.fields("FunctionToCall")
        FieldOrder = rs.fields("FieldOrder")
        PartOnly = rs.fields("PartOnly")
        
        If EmailPart Like "*[" & FieldCaption & "]*" Then
            If FunctionToCall = "BuildAdvertisementLink" Then
                fieldValue = Run("BuildAdvertisementLink", PropertyListID)
            Else
                fieldValue = Run(FunctionToCall, PropertyListID, fieldName, FieldCaption, PartOnly, PropertyEmailRecipientID)
            End If
            
            EmailPart = replace(EmailPart, "[" & FieldCaption & "]", fieldValue)
            
        End If
        
        rs.MoveNext
    Loop
    
End Function

Private Function AttachFilesToEmail(olMail As Object)

    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblEmailAttachments WHERE IsAttached")
    
    If rs.EOF Then Exit Function
    
    Dim fileName As String
    
    Do Until rs.EOF
        fileName = rs.fields("AbsoluteLink")
        ''mailAttachments.AddAttachment rs.Fields("AbsoluteLink")
        olMail.Attachments.Add fileName
        rs.MoveNext
    Loop
       
End Function
