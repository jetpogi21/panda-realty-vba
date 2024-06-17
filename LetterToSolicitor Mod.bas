Attribute VB_Name = "LetterToSolicitor Mod"
Option Compare Database
Option Explicit

Public Function LetterToSolicitorCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function SendLetterToSolicitor(frm As Form)
    
    Dim filePath: filePath = GenerateLetterToSolicitor(frm, False)
    If isFalse(filePath) Then Exit Function
    
    Dim PropertyListID: PropertyListID = frm("PropertyListID")
    If isFalse(PropertyListID) Then Exit Function
    
    Dim Subject: Subject = "RE: SALE : " & GetPropertyAddress(PropertyListID)
    
    Dim BuyerSolicitorEmail: BuyerSolicitorEmail = GetSolicitorEmail(PropertyListID, "Buyer")
    If ExitIfTrue(isFalse(PropertyListID), "Buyer solicitor email is empty.") Then Exit Function
    
    Dim SellerSolicitorEmail: SellerSolicitorEmail = GetSolicitorEmail(PropertyListID, "Seller")
    If ExitIfTrue(isFalse(PropertyListID), "Seller solicitor email is empty.") Then Exit Function
    Dim recipientEmail: recipientEmail = BuyerSolicitorEmail & "; " & SellerSolicitorEmail
    
    Dim EmailSignatureBody: EmailSignatureBody = ELookup("tblEmailSignatures", "EmailSignature = " & Esc("Richard Falzon"), "EmailSignatureBody")
    
    Dim HTMLBody: HTMLBody = "Dear Colleagues,<br/><br/>Please find the confirmation letter for deposit monies received in trust pending Settlement advice, attached." & _
        "<br/><br/>" & replace(EmailSignatureBody, vbCrLf, "<br/>")

    SendFileToEmail frm, Subject, filePath, recipientEmail, HTMLBody
     
End Function

Private Function GetSolicitorEmail(PropertyListID, Association)

    GetSolicitorEmail = ELookup("qryPropertyEntities", "EntityCategoryName = ""Contact"" AND " & _
        "ContactCategoryName = ""Solicitor"" AND PropertyListID = " & PropertyListID & " AND Association = " & Esc(Association), "EmailAddress", "PropertyEntityID")
        
End Function

Public Function GenerateLetterToSolicitor(frm As Form, Optional OpenFile As Boolean = True) As String
    
    Dim txtLetterType: txtLetterType = frm("txtLetterType")
    
    If ExitIfTrue(isFalse(txtLetterType), "Please select a letter type.") Then Exit Function
    
    Dim PropertyListID: PropertyListID = frm("PropertyListID")
    If isFalse(PropertyListID) Then Exit Function
    
    ''Validate -> There should be an amount for the LetterType being generated
    Dim Description: Description = IIf(txtLetterType = "Initial Deposit", "INITIAL DEPOSIT RECEIPT", "BALANCE DEPOSIT RECEIPT")
    Dim LetterTypeAmount: LetterTypeAmount = GetEventTimelineAmount(PropertyListID, Description, True)
    
    If ExitIfTrue(LetterTypeAmount = 0, "There's no checked amount for the " & Esc(Description) & ".") Then Exit Function
    
    RunSQL "DELETE FROM tblLetterToSolicitors"
    RunSQL "INSERT INTO tblLetterToSolicitors (LetterType, PropertyListID) VALUES (" & Esc(txtLetterType) & "," & PropertyListID & ")"
    
    Dim fileName:  fileName = "Letter to solicitors - " & txtLetterType
    Dim filePath: filePath = GeneratePropertyRelatedPDF(frm, PropertyListID, fileName, "rptLetterToSolicitors", fileName, OpenFile)
    If OpenFile Then CreateObject("Shell.Application").Open filePath
    ''DoCmd.OpenReport "rptLetterToSolicitors", acViewPreview
    GenerateLetterToSolicitor = filePath
    
End Function

Public Function GetTotalDeposits(PropertyListID, LetterType)
    
    GetTotalDeposits = GetEventTimelineAmount(PropertyListID, "INITIAL DEPOSIT RECEIPT")
    
    If LetterType <> "Initial Deposit" Then
        GetTotalDeposits = GetTotalDeposits + GetEventTimelineAmount(PropertyListID, "BALANCE DEPOSIT RECEIPT")
    End If
    
End Function
