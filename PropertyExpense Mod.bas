Attribute VB_Name = "PropertyExpense Mod"
Option Compare Database
Option Explicit

Public Function PropertyExpenseCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function Open_rptTaxInvoice(frm As Form)
    
    DoCmd.RunCommand acCmdSaveRecord
    Dim PropertyListID: PropertyListID = frm("PropertyListID")
    
    InsertTo_tblExpensesIncurred PropertyListID
    
    Dim PropertyDirectory: PropertyDirectory = GetPropertyDirectory(frm, "PropertyListID")
    Dim fileName: fileName = "Tax Invoice.pdf"
    Dim filePath: filePath = PropertyDirectory & fileName
    
    Dim rptName: rptName = "rptTaxInvoice"
    DoCmd.OpenReport rptName, acViewDesign, , , acHidden
    Dim rpt As Report: Set rpt = Reports(rptName)
    
    SetReportMargin rpt
    DoCmd.Close acReport, rptName, acSaveYes
    
    DoCmd.OpenReport rptName, acViewReport, , "PropertyListID = " & PropertyListID, acHidden
    
    SaveReportAsPDF rptName, filePath
    
    MsgBox "Tax Invoice generated at: " & Esc(filePath)
    
    CreateObject("Shell.Application").Open filePath
    
End Function

Private Function SetReportMargin(rpt As Report)
    
    rpt.Printer.BottomMargin = InchToTwip(1)
    rpt.Printer.TopMargin = InchToTwip(1)
    rpt.Printer.LeftMargin = InchToTwip(0.75)
    rpt.Printer.RightMargin = InchToTwip(0.75)

End Function

Private Function InsertTo_tblExpensesIncurred(PropertyListID)
    
    RunSQL "DELETE FROM tblExpensesIncurred"
    
    Dim TotalAmount: TotalAmount = 0
    Dim rs: Set rs = ReturnRecordset("SELECT * FROM tblPropertyExpenses WHERE PropertyListID = " & PropertyListID & " ORDER BY PropertyExpenseID")
    
    Do Until rs.EOF
        Dim PropertyExpenseID: PropertyExpenseID = rs.fields("PropertyExpenseID")
        Dim ExpenseAmount: ExpenseAmount = rs.fields("ExpenseAmount")
        TotalAmount = TotalAmount + ExpenseAmount
        RunSQL "INSERT INTO tblExpensesIncurred (PropertyExpenseID) VALUES (" & PropertyExpenseID & ")"
        rs.MoveNext
    Loop
    
    Dim LastID: LastID = ELookup("tblExpensesIncurred", "ExpensesIncurredID > 0", "ExpensesIncurredID", "ExpensesIncurredID DESC")
    If Not isFalse(LastID) Then
        RunSQL "UPDATE tblExpensesIncurred set TotalAmount = " & TotalAmount & " WHERE ExpensesIncurredID = " & LastID
    End If
    
End Function

Public Function GetSettlementDate(PropertyListID) As Variant
    
    If isFalse(PropertyListID) Then Exit Function
    Dim SettlementDate: SettlementDate = ELookup("qryEventTimelines", "Description = ""SETTLEMENT DATE"" AND " & _
        "PropertyListID = " & PropertyListID, "EventTimelineDate", "EventTimelineID")
    
    If isFalse(SettlementDate) Then Exit Function
        
    GetSettlementDate = CDate(SettlementDate)
    
End Function

Public Function GetEventTimelineAmount(PropertyListID, Description, Optional CheckedOnly As Boolean = False) As Double
    
    If isFalse(PropertyListID) Then Exit Function
    
    Dim TrustReceiptDescription As String
    If Description = "INITIAL DEPOSIT RECEIPT" Then
        TrustReceiptDescription = "Initial Deposit Trust Receipt"
    ElseIf Description = "BALANCE DEPOSIT RECEIPT" Then
        TrustReceiptDescription = "Balance Deposit Trust Receipt"
    End If
    
    Dim filterStr As String: filterStr = "(Description = " & Esc(Description) & " AND NOT EventTimelineAmount IS NULL)"
    
    If TrustReceiptDescription <> "" Then
        filterStr = "(" & filterStr & " OR (Description Like " & Esc(TrustReceiptDescription & "*") & " AND NOT EventTimelineAmount IS NULL))"
    End If
    
    filterStr = filterStr & " AND PropertyListID = " & PropertyListID
    
    If CheckedOnly Then filterStr = filterStr & " AND IsChecked"
    
    Dim EventTimelineAmount: EventTimelineAmount = ELookup("qryEventTimelines", filterStr, "EventTimelineAmount", "EventTimelineID")
        
    If isFalse(EventTimelineAmount) Then Exit Function
    
    GetEventTimelineAmount = EventTimelineAmount
    
End Function

Public Function GetSellerSolicitorName(PropertyListID, Optional Association = "Seller") As String
    
    If isFalse(PropertyListID) Then Exit Function
    Dim EntityID:  EntityID = ELookup("qryPropertyEntities", "EntityCategoryName = ""Contact"" AND " & _
        "ContactCategoryName = ""Solicitor"" AND PropertyListID = " & PropertyListID & " AND Association = " & Esc(Association), "EntityID", "PropertyEntityID")
    If isFalse(EntityID) Then Exit Function
    GetSellerSolicitorName = GetPropertyReceiptReceivedFrom(EntityID, PropertyListID)
    
End Function

Public Function GetTaxInvoiceSellerName(PropertyListID) As String
    
    If isFalse(PropertyListID) Then Exit Function
    Dim EntityID:  EntityID = ELookup("qryPropertyEntities", "EntityCategoryName = ""Seller"" AND EntityName <> ""-"" AND PropertyListID = " & PropertyListID, "EntityID", "PropertyEntityID")
    
    If isFalse(EntityID) Then Exit Function
    
    GetTaxInvoiceSellerName = GetPropertyReceiptReceivedFrom(EntityID, PropertyListID)
    
End Function

Public Function GetExpensesIncurredTotalAmount() As Double
    
    GetExpensesIncurredTotalAmount = ESum2("tblExpensesIncurred", "ExpensesIncurredID > 0", "TotalAmount")
    
End Function

Public Function GetPaymentReceivedTotalAmountReceived(PropertyListID) As Double
    
    If isFalse(PropertyListID) Then Exit Function
    GetPaymentReceivedTotalAmountReceived = ESum2("qryPaymentsReceived", "PropertyListID = " & PropertyListID, "AmountReceived")
    
End Function

Public Function GetAgreedReimbursementTotal(PropertyListID) As Double
    
    If isFalse(PropertyListID) Then Exit Function
    GetAgreedReimbursementTotal = ESum2("qryAgreedReimbursements", "PropertyListID = " & PropertyListID, "AgreedReimbursement")
    
End Function
