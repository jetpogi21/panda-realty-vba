Attribute VB_Name = "TrustReceipt Mod"
Option Compare Database
Option Explicit

Public Function TrustReceiptCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function TrustPaymentValidate(frm As Form) As Boolean

    Dim Amount: Amount = frm("Amount")
    
    If Amount = 0 Then
        MsgBox "Please enter an amount."
        frm("Amount").SetFocus
        Exit Function
    End If
    
    TrustPaymentValidate = True
    
    
End Function

Public Function TrustReceiptValidate(frm As Form) As Boolean

    Dim Amount: Amount = frm("Amount")
    
    If Amount = 0 Then
        MsgBox "Please enter an amount."
        frm("Amount").SetFocus
        Exit Function
    End If
    
    TrustReceiptValidate = True
    
    
End Function

Public Function UpdateTrustReceiptSubject()

    RunSQL "UPDATE tblTrusReceiptSettings SET EmailSubject = " & Esc("RE SALE: [StreetAddress] [Suburb] [State] [Postcode]")
    
End Function

Public Function GetPropertyDirectory(frm As Form, Optional propertyFieldName = "Property")

    Dim uploadDirectory, strFolderExists
    
    uploadDirectory = GetAttachmentsDirectory
    Dim PropertyListID: PropertyListID = frm(propertyFieldName)
   
    Dim StreetAddress: StreetAddress = ELookup("tblPropertyList", "PropertyListID = " & PropertyListID, "StreetAddress")
    uploadDirectory = GetAttachmentsDirectory(StreetAddress)
    
    strFolderExists = Dir(uploadDirectory, vbDirectory)
    
    ''Create the directory if it doesn't exist
    If strFolderExists = "" Then
        CreateDirectoryStructure uploadDirectory
    End If
    
    Dim fs As Object, filePath, fileName
    Set fs = CreateObject("Scripting.FileSystemObject")

    filePath = uploadDirectory & fileName
    
    GetPropertyDirectory = uploadDirectory
    
End Function

''For lawrence pass 163360
Public Function GetPropertyReceiptReceivedFrom(EntityID, PropertyListID) As String
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM qryEntities WHERE EntityID = " & EntityID)
    Dim EntityCategoryName: EntityCategoryName = rs.fields("EntityCategoryName")
    Dim EntityName: EntityName = rs.fields("EntityName")
    Dim members As New clsArray
    
    Select Case EntityCategoryName
        Case "Buyer"
            Set rs = ReturnRecordset("SELECT * FROM tblEntityMembers WHERE EntityID = " & EntityID)
            
            Set members = New clsArray
            If rs.EOF Then
                members.Add EntityName
                GetPropertyReceiptReceivedFrom = GetConciseName(members)
                Exit Function
            End If
            
            Do Until rs.EOF
                members.Add rs.fields("MemberName")
                rs.MoveNext
            Loop
            
            GetPropertyReceiptReceivedFrom = GetConciseName(members)
        Case "Seller"
            
            Set rs = ReturnRecordset("SELECT * FROM qryPropertyEntities WHERE PropertyListID = " & PropertyListID & " AND IsOriginalOwner " & _
                " AND EntityName <> ""-"" ORDER BY PropertyEntityID")
            Set members = New clsArray
            Do Until rs.EOF
                ''Debug.Print rs.fields("EntityName")
                members.Add rs.fields("EntityName")
                rs.MoveNext
            Loop
            GetPropertyReceiptReceivedFrom = GetConciseName(members)
        Case "Tenant"
            GetPropertyReceiptReceivedFrom = EntityName
        Case "Contact"
            GetPropertyReceiptReceivedFrom = EntityName
        Case Else
            GetPropertyReceiptReceivedFrom = ""

    End Select
    
End Function

Public Function GetEmailRecipientSalutation(EntityID) As String
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM qryEntities WHERE EntityID = " & EntityID)
    Dim EntityCategoryName: EntityCategoryName = rs.fields("EntityCategoryName")
    Dim EntityName: EntityName = rs.fields("EntityName")
    
    Select Case EntityCategoryName
        Case "Buyer", "Seller"
            Set rs = ReturnRecordset("SELECT * FROM tblEntityMembers WHERE EntityID = " & EntityID)
            
            Dim members As New clsArray
            If rs.EOF Then
                members.Add EntityName
                GetEmailRecipientSalutation = GenerateSalutation(members)
                Exit Function
            End If
            
            
            Do Until rs.EOF
                members.Add rs.fields("MemberName")
                rs.MoveNext
            Loop
            
            GetEmailRecipientSalutation = GenerateSalutation(members)
        Case "Tenant"
            GetEmailRecipientSalutation = EntityName
        Case "Contact"
            GetEmailRecipientSalutation = EntityName
        Case Else
            GetEmailRecipientSalutation = ""

    End Select
    
End Function

Private Function GenerateSalutation(fullNameArr As clsArray) As String
    Dim numNames As Integer
    Dim salutation As String

    
    ' Get the number of names in the array
    numNames = fullNameArr.Count
    
    If numNames = 1 Then
        ' Only one recipient
        GenerateSalutation = GetSurname(fullNameArr.arr(0), True)
    ElseIf numNames = 2 Then
        ' Two recipients
        GenerateSalutation = GetSurname(fullNameArr.arr(0), True) & " and " & GetSurname(fullNameArr.arr(1), True)
    ElseIf numNames > 2 Then
        ' More than two recipients
        Dim i As Integer
        For i = 0 To numNames - 1
            salutation = salutation & GetSurname(fullNameArr.arr(i), True) & ", "
        Next i
        salutation = left(salutation, Len(salutation) - 2) & " and " & GetSurname(fullNameArr.arr(numNames - 1), True)
        GenerateSalutation = salutation
    Else
        ' No recipients
        GenerateSalutation = "No recipients specified"
    End If
End Function

Private Function GetConciseName(namesArray As clsArray) As String
    Dim surnameDict As Object
    Set surnameDict = CreateObject("Scripting.Dictionary")
    
    Dim result As String
    result = ""
    
    Dim Name As Variant
    For Each Name In namesArray.arr
        Dim words As Variant
        words = Split(Name, " ")
        
        ' Extract last word as surname
        Dim surname As String
        surname = words(UBound(words))
        
        Dim convertedSurname As String
        convertedSurname = ELookup("tblSurnameConversions", "Surname = " & Esc(surname), "ConvertedSurname")
        
        If Not isFalse(convertedSurname) Then surname = convertedSurname
        
        ' Extract first character as initial with period
        Dim initial As String
        initial = left(words(0), 1)
        
        ' Convert surname to lowercase for case-insensitive comparison
        Dim lowercaseSurname As String
        lowercaseSurname = LCase(surname)
        
        ' Check if surname already encountered (case-insensitive)
        If surnameDict.Exists(lowercaseSurname) Then
            ' If yes, append initial to existing value with an ampersand
            surnameDict(lowercaseSurname) = surnameDict(lowercaseSurname) & " & " & initial
        Else
            ' If no, add to dictionary (using lowercase as the key)
            surnameDict.Add lowercaseSurname, initial
        End If
    Next Name
    
    ' Build the result string with uppercase surnames
    Dim key As Variant
    For Each key In surnameDict.Keys
        result = result & surnameDict(key) & " " & UCase(key) & ","
    Next key
    
    ' Remove trailing comma
    If Len(result) > 0 Then
        result = left(result, Len(result) - 1)
    End If
    
    GetConciseName = result
End Function

Public Function SendTrustReceiptToEmail(frm As Form)
    
    DoCmd.RunCommand acCmdSaveRecord
    Dim TrustReceiptID: TrustReceiptID = frm("TrustReceiptID")
    Dim TrustReceiptNo: TrustReceiptNo = frm("TrustReceiptNo")
    Dim EntityID: EntityID = frm("ReceivedFrom")
    Dim PropertyListID:  PropertyListID = frm("Property")
    
    If isFalse(TrustReceiptID) Then
        MsgBox "Trust Receipt is empty.", vbCritical + vbOKOnly
        Exit Function
    End If
    
    Dim EmailAddress: EmailAddress = ELookup("tblEntities", "EntityID = " & EntityID, "EmailAddress")
    
    If isFalse(EmailAddress) Then
        MsgBox "Entity has no email address.", vbCritical + vbOKOnly
        Exit Function
    End If
    
    Dim PropertyDirectory: PropertyDirectory = GetPropertyDirectory(frm)
    Dim fileName: fileName = "TR" & TrustReceiptNo & ".pdf"
    Dim filePath: filePath = PropertyDirectory & fileName
    
    DoCmd.OpenReport "rptTrustReceiptOriginal", acViewReport, , "TrustReceiptID = " & TrustReceiptID, acHidden
    SaveReportAsPDF "rptTrustReceiptOriginal", filePath
    
    Dim subjectArr As New clsArray
    Dim DefaultSubject: DefaultSubject = ELookup("tblTrusReceiptSettings", "TrusReceiptSettingID > 0", "EmailSubject")
    If Not isFalse(DefaultSubject) Then subjectArr.Add DefaultSubject
    subjectArr.Add "Trust Receipt No. " & TrustReceiptNo
    Dim Subject: Subject = subjectArr.JoinArr(": ")
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblPropertyList WHERE PropertyListID = " & PropertyListID)
    Dim fld As Field
    For Each fld In rs.fields
    If InStr(Subject, "[" & fld.Name & "]") > 0 Then
        Subject = replace(Subject, "[" & fld.Name & "]", fld.Value)
        End If
    Next fld
    
    ''Buyer Surname (special case)
    Dim txtToFind As String: txtToFind = "[Buyer Surname]"
    If InStr(Subject, txtToFind) > 0 Then
        Dim surname: surname = BuildEntityMembers(PropertyListID, "Successful", "", "Surname")
        Subject = replace(Subject, txtToFind, surname)
    End If

'    Dim EmailSalutation: EmailSalutation = GetEmailRecipientSalutation(EntityID)
'    Dim DefaultSignature: DefaultSignature = ELookup("tblTrusReceiptSettings", "TrusReceiptSettingID > 0", "EmailSignature")
'    DefaultSignature = replace(DefaultSignature, vbNewLine, "<br/>")
'    Dim HTMLBody: HTMLBody = "Hi " & EmailSalutation & ", <br/><br/>Please see the attached file. <br/><br/>" & DefaultSignature
    
    Dim DefaultSignature: DefaultSignature = ELookup("tblTrusReceiptSettings", "TrusReceiptSettingID > 0", "EmailSignature")
    DefaultSignature = replace(DefaultSignature, vbNewLine, "<br/>")
    Dim HTMLBody: HTMLBody = DefaultSignature
    
    txtToFind = "[Salutation]"
    If InStr(HTMLBody, txtToFind) > 0 Then
        Dim EmailSalutation: EmailSalutation = GetEmailRecipientSalutation(EntityID)
        HTMLBody = replace(HTMLBody, txtToFind, EmailSalutation)
    End If
    
    SendFileToEmail frm, Subject, filePath, EmailAddress, HTMLBody
    
End Function

Public Function frmTrustReceipts_ReceivedFrom_AfterUpdate(frm As Form, Optional TransactionType = "Receipt")
    
    frm("Details") = Null
    SetDetails_RowSource frm, TransactionType
    
End Function

Public Function frmTrustReceipts_OnCurrent(frm As Form)
    
    Property_frmTrustReceipts_AfterUpdate frm, True
    SetDetails_RowSource frm
    
End Function

Public Function frmTrustPayments_OnCurrent(frm As Form)
    
    Property_frmTrustReceipts_AfterUpdate frm, True, "Payment"
    SetDetails_RowSource frm, "Payment"
    
End Function

Public Function InsertSuccessfulBuyerTo_tblTrustEntities(vProperty)
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "qryBuyerOffers"
          .AddFilter "qryBuyerOffers.PropertyListID = " & vProperty
          .Joins.Add GenerateJoinObj("tblPropertyEntityCaptions", "PropertyEntityID")
          .fields = "qryBuyerOffers.EntityID,PropertyEntitycaption,""Buyer"" AS EntityCategoryName"
          .GroupBy = "qryBuyerOffers.EntityID,PropertyEntitycaption"
          sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "tblTrustEntities"
          .fields = "EntityID,PropertyEntityCaption,EntityCategoryName"
          .InsertSQL = sqlStr
          .InsertFilterField = "EntityID,PropertyEntityCaption,EntityCategoryName"
          rowsAffected = .Run
    End With
    
End Function

Public Function InsertFirstSeller_tblTrustEntities(vProperty)
    
    ''Get the first owner of this property
    Dim PropertyEntityID:  PropertyEntityID = ELookup("qryPropertyEntities", "EntityCategoryName = ""Seller"" AND EntityName <> ""-"" AND PropertyListID = " & vProperty, "PropertyEntityID")
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "qryPropertyEntities"
          .AddFilter "PropertyEntityID = " & PropertyEntityID
          .fields = "EntityID,PropertyEntitycaption,EntityCategoryName"
          sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "tblTrustEntities"
          .fields = "EntityID,PropertyEntityCaption,EntityCategoryName"
          .InsertSQL = sqlStr
          .InsertFilterField = "EntityID,PropertyEntityCaption,EntityCategoryName"
          rowsAffected = .Run
    End With
    
End Function

Public Function InsertSolicitors_tblTrustEntities(vProperty)
    
    ''Get the first owner of this property
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "qryPropertyEntities"
          .AddFilter "PropertyListID = " & vProperty & " AND EntityCategoryName = ""Contact"" and ContactCategoryName = ""Solicitor"""
          .fields = "EntityID,PropertyEntitycaption,EntityCategoryName"
          sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "tblTrustEntities"
          .fields = "EntityID,PropertyEntityCaption,EntityCategoryName"
          .InsertSQL = sqlStr
          .InsertFilterField = "EntityID,PropertyEntityCaption,EntityCategoryName"
          rowsAffected = .Run
    End With
    
End Function

Public Function InsertMyPandaGeneralAccountTo_tblTrustEntities(vProperty)
    
    ''SELECT EntityID," & Esc("My Panda Realty General Account") & " AS MainField,EntityCategoryName FROM qryEntities WHERE EntityName = " & Esc("My Panda Realty")
    ''Get the first owner of this property
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "qryEntities"
          .AddFilter "EntityName = " & Esc("My Panda Realty")
          .fields = "EntityID," & Esc("My Panda Realty General Account") & " AS PropertyEntitycaption,EntityCategoryName"
          sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "tblTrustEntities"
          .fields = "EntityID,PropertyEntityCaption,EntityCategoryName"
          .InsertSQL = sqlStr
          .InsertFilterField = "EntityID,PropertyEntityCaption,EntityCategoryName"
          rowsAffected = .Run
    End With
    
End Function

Public Function Property_frmTrustReceipts_AfterUpdate(frm As Form, Optional onCurrentEvent As Boolean = False, Optional TransactionType = "Receipt")
    
    Dim vProperty: vProperty = frm("Property")
    Dim sqlStr
    
    If isFalse(vProperty) Then
        sqlStr = "SELECT EntityID,PropertyEntityCaption As MainField,EntityCategoryName FROM tblTrustEntities WHERE TrustEntityID = 0 ORDER BY PropertyEntityCaption"
    Else
        RunSQL "DELETE FROM tblTrustEntities"
        If TransactionType = "Receipt" Then
            ''get the succesfullbuyer sql
            InsertSuccessfulBuyerTo_tblTrustEntities vProperty
            ''get the first seller
            InsertFirstSeller_tblTrustEntities vProperty
            ''get solicitors
            InsertSolicitors_tblTrustEntities vProperty
            
            sqlStr = "SELECT EntityID,PropertyEntityCaption As MainField,EntityCategoryName FROM tblTrustEntities ORDER BY PropertyEntityCaption"
        Else
            
            ''Insert my panda realty general account
            InsertMyPandaGeneralAccountTo_tblTrustEntities vProperty
            ''get the first seller
            InsertFirstSeller_tblTrustEntities vProperty
            ''get solicitors
            InsertSolicitors_tblTrustEntities vProperty
            InsertSuccessfulBuyerTo_tblTrustEntities vProperty
            
            sqlStr = "SELECT EntityID,PropertyEntityCaption As MainField,EntityCategoryName FROM tblTrustEntities ORDER BY PropertyEntityCaption"
            
        End If
    End If
    
    Dim EntityFieldName: EntityFieldName = IIf(TransactionType = "Receipt", "ReceivedFrom", "PaidTo")
    
    If Not onCurrentEvent Then frm(EntityFieldName) = Null
    
    frm(EntityFieldName).rowSource = sqlStr
    frm(EntityFieldName).Requery
    
    ''SELECT EntityID,EntityName As MainField FROM tblEntities ORDER BY EntityName
End Function

Public Function PropertyTrustReceiptRefresh(frm As Form)
    
    Dim resp: resp = MsgBox("This will rebuild the property owners. Do you want to proceed?", vbYesNo)
    
    If resp = vbNo Then Exit Function
    Build_tblPropertyOwnerSurnames True
    frm("Property").Requery
    
    MsgBox "Property owners rebuilt.", vbOKOnly
    
End Function

Public Function ReceivedFromRefresh(frm As Form, Optional TransactionType = "Receipt")
    
    Dim resp: resp = MsgBox("This will rebuild the property entities. Do you want to proceed?", vbYesNo)
    
    If resp = vbNo Then Exit Function
    
    Dim EntityFieldName: EntityFieldName = IIf(TransactionType = "Receipt", "ReceivedFrom", "PaidTo")
    Build_tblPropertyEntityCaptions True
    frm(EntityFieldName).Requery
    
    MsgBox "Property entities rebuilt.", vbOKOnly
    
End Function

Private Function Build_tblPropertyEntityCaptions(Optional refresh As Boolean = False)
    
    If ECount("tblPropertyEntityCaptions", "PropertyEntityCaptionID > 0") And Not refresh Then Exit Function
    RunSQL "DELETE FROM tblPropertyEntityCaptions"
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "qryPropertyEntities"
          .AddFilter "IsFavorite"
          .fields = "PropertyEntityID,EntityID,PropertyListID"
          .OrderBy = "PropertyEntityID"
          sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "MAKE"
          .Source = sqlStr
          .SourceAlias = "temp"
          .MakeTable = "madePropertyEntities"
          rowsAffected = .Run
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "madePropertyEntities"
          .fields = "PropertyEntityID,GetPropertyReceiptReceivedFrom(EntityID,PropertyListID) As PropertyEntityCaption"
          .OrderBy = "PropertyEntityID"
          sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "tblPropertyEntityCaptions"
          .fields = "PropertyEntityID,PropertyEntityCaption"
          .InsertSQL = sqlStr
          .InsertFilterField = "PropertyEntityID,PropertyEntityCaption"
          rowsAffected = .Run
    End With
    
End Function


Private Function Build_tblPropertyOwnerSurnames(Optional refresh As Boolean = False)
    
    If ECount("tblPropertyOwnerSurnames", "PropertyListID > 0") And Not refresh Then Exit Function
    RunSQL "DELETE FROM tblPropertyOwnerSurnames"
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "tblPropertyList"
          .fields = "PropertyListID, GetPropertyIdentifier(StreetAddress, Suburb, Owner1Name, Owner2Name, Owner3Name) As PropertyOwnerSurname"
          .OrderBy = "PropertyListID"
          sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "tblPropertyOwnerSurnames"
          .fields = "PropertyListID,PropertyOwnerSurname"
          .InsertSQL = sqlStr
          .InsertFilterField = "PropertyListID,PropertyOwnerSurname"
          rowsAffected = .Run
    End With
    
End Function

Public Function SetDetails_RowSource(frm As Form, Optional TransactionType = "Receipt")
    
    Dim EntityFieldName: EntityFieldName = IIf(TransactionType = "Receipt", "ReceivedFrom", "PaidTo")
    Dim EntityCategory: EntityCategory = frm(EntityFieldName).Column(2)
    Dim EntityName: EntityName = frm(EntityFieldName).Column(1) ''My Panda Realty General Account
    
    Dim sqlStr
    If isFalse(EntityCategory) Then
        sqlStr = "SELECT " & TransactionType & "DetailID," & TransactionType & "Detail FROM tbl" & TransactionType & "Details WHERE " & TransactionType & "DetailID = 0"
    Else
        If EntityName = "My Panda Realty General Account" Then
            sqlStr = "SELECT " & TransactionType & "DetailID," & TransactionType & "Detail FROM tbl" & TransactionType & "Details WHERE ForPanda ORDER BY " & TransactionType & "Detail"
        Else
            sqlStr = "SELECT " & TransactionType & "DetailID," & TransactionType & "Detail FROM tbl" & TransactionType & "Details WHERE For" & EntityCategory & " ORDER BY " & TransactionType & "Detail"
        End If
    End If
    frm("Details").rowSource = sqlStr
    frm("Details").Requery
    
End Function

Public Function AddPropertyTrustReceipt(frm As Form, Optional PropertyListID, Optional TransactionType = "Receipt")

    Build_tblPropertyOwnerSurnames
    
    Dim FormName: FormName = IIf(TransactionType = "Receipt", "frmTrustReceipts", "frmTrustPayments")
    DoCmd.OpenForm FormName, , , , acFormAdd
    
    If Not isFalse(PropertyListID) Then
        Forms(FormName)("Property") = PropertyListID
        Set frm = Forms(FormName)
        Property_frmTrustReceipts_AfterUpdate frm, False, TransactionType
    End If
    
End Function

Public Function Open_frmTrusReceiptSettings(frm As Form)
    
    DoCmd.OpenForm "frmTrusReceiptSettings"
    
End Function

Public Function EditPropertyTrustReceipt(frm As Form, Optional subformName As String = "subform")
    
    OpenFormFromMain "frmTrustReceipts", subformName, "TrustReceiptID", frm
   
End Function

Public Function EditPropertyTrustPayment(frm As Form, Optional subformName As String = "subform")
    
    OpenFormFromMain "frmTrustPayments", subformName, "TrustPaymentID", frm
   
End Function

Public Function EditPropertyTrustReceiptByNumber(TransactionNumber, TransactionType)
    
    If isFalse(TransactionNumber) Then Exit Function
    Dim FormName, transactionNumberField
    
    If TransactionType = "RJ" Then
        FormName = "frmTrustReceipts"
        transactionNumberField = "TrustReceiptNo"
    Else
        FormName = "frmTrustPayments"
        transactionNumberField = "TrustPaymentNo"
    End If
    
    DoCmd.OpenForm FormName, , , transactionNumberField & " = " & TransactionNumber
    
End Function

Public Function Open_rptTrustReceipts(frm As Form, Optional printMode As String = "Preview")
    On Error GoTo ErrorHandler
    
    DoCmd.RunCommand acCmdSaveRecord
    
    Dim TrustReceiptID As Variant, TrustReceiptNo
    TrustReceiptID = frm("TrustReceiptID")
    TrustReceiptNo = frm("TrustReceiptNo")
    
    If IsNull(TrustReceiptID) Then
        MsgBox "TrustReceiptID is null or undefined.", vbExclamation
        Exit Function
    End If
    
    Dim reportVersion: reportVersion = frm("cboReportVersion")
    
    If IsNull(reportVersion) Then
        MsgBox "Please select either ""Original"",""Duplicate"" Or ""Both""", vbExclamation
        Exit Function
    End If
    
    If printMode = "Print" Then
        Dim resp: resp = MsgBox("This will send the file to your default printer. Do you want to proceed?", vbYesNo)
        If resp = vbNo Then Exit Function
    End If
    
    ''Original;Duplicate
    Dim reportName
    If reportVersion = "Original" Then
        reportName = "rptTrustReceiptOriginal"
    ElseIf reportVersion = "Duplicate" Then
        reportName = "rptTrustReceipts"
    Else
        reportName = "rptTrustReceiptBoth"
    End If
    
    If printMode = "PDF" Then
        Dim PropertyDirectory: PropertyDirectory = GetPropertyDirectory(frm)
        Dim fileName: fileName = "TR" & TrustReceiptNo & "-" & reportVersion & ".pdf"
        Dim filePath: filePath = PropertyDirectory & fileName
        
        DoCmd.OpenReport reportName, acViewReport, , "TrustReceiptID = " & TrustReceiptID, acHidden
        SaveReportAsPDF reportName, filePath
        
        MsgBox "Trust Receipt Saved as: " & Esc(filePath)
    Else
        Dim printingMode: printingMode = acViewPreview
        If printMode = "Print" Then printingMode = acViewPreview
        
        DoCmd.OpenReport reportName, printingMode, , "TrustReceiptID = " & TrustReceiptID
        
        DoCmd.RunCommand acCmdPrint
    End If

    Exit Function
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Function


Public Function GetDefaultTrustReceiptNo(frm As Form, Optional TransactionType = "Receipt")

    Dim tblName As String, tblPk, tblMainField, StartingNumberField
    
    If TransactionType = "Receipt" Then
        tblName = "tblTrustReceipts"
        tblPk = "TrustReceiptID"
        tblMainField = "TrustReceiptNo"
        StartingNumberField = "StartingNumber"
    Else
        tblName = "tblTrustPayments"
        tblPk = "TrustPaymentID"
        tblMainField = "TrustPaymentNo"
        StartingNumberField = "StartNumberPayment"
    End If
    
    If frm.NewRecord Then
        ''Get the max from tblTrustReceipts
        Dim TrustReceiptNo: TrustReceiptNo = ELookup(tblName, tblPk & " > 0", tblMainField, tblMainField & " DESC")
        If isFalse(TrustReceiptNo) Then
            TrustReceiptNo = ELookup("tblTrusReceiptSettings", StartingNumberField & " > 0", StartingNumberField, StartingNumberField & " DESC")
        Else
            TrustReceiptNo = TrustReceiptNo + 1
        End If
        
        GetDefaultTrustReceiptNo = TrustReceiptNo
        Exit Function
        
    End If
    
    GetDefaultTrustReceiptNo = 0
    
End Function

Public Function rptTrustReceipts_OnLoad(rpt As Report)
    
    Dim digitalSignature: digitalSignature = ELookup("tblTrusReceiptSettings", "TrusReceiptSettingID > 0", "DigitalSignature")
    
    If isFalse(digitalSignature) Then Exit Function
    
    Dim assetDir, fs As Object
    assetDir = GetApplicationSetting("Asset Directory")
    
    If assetDir = "" Then assetDir = CurrentProject.Path
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FolderExists(assetDir) Then assetDir = CurrentProject.Path
    
    Dim filePath: filePath = concat(assetDir, "\", digitalSignature)
    
    On Error Resume Next
    rpt("imgSignature").Picture = filePath
    
    If DoesPropertyExists(rpt, "imgSignature1") Then
        rpt("imgSignature1").Picture = filePath
    End If
    ''rpt("imgSignature1").Picture = filePath
    
End Function
