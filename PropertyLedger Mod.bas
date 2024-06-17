Attribute VB_Name = "PropertyLedger Mod"
Option Compare Database
Option Explicit

Public Function PropertyLedgerCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function RefreshPropertyLedger()
    
    If Not IsFormOpen("frmPropertyList") Then Exit Function
    
    Dim frm As Form: Set frm = Forms("frmPropertyList")
    Dim PropertyListID: PropertyListID = frm("PropertyListID")
    
    If isFalse(PropertyListID) Then Exit Function
    
    RunSQL "DELETE FROM tblPropertyLedgers WHERE PropertyListID = " & PropertyListID
    
    InsertReceiptsToLedger PropertyListID
    InsertPaymentsToLedger PropertyListID
    
    frm("subTrustReceipts").Report.Requery
    
End Function

Public Function InsertReceiptsToLedger(PropertyListID)
    InsertTransactionsToLedger PropertyListID, "qryTrustReceipts", "RJ", "Debit", "Credit", "DateReceived", "ReceiptDetail", "TrustReceiptNo"
End Function

Public Function InsertPaymentsToLedger(PropertyListID)
    InsertTransactionsToLedger PropertyListID, "qryTrustPayments", "PJ", "Credit", "Debit", "DatePaid", "PaymentDetail", "TrustPaymentNo"
End Function

Private Sub InsertTransactionsToLedger(PropertyListID, sourceQuery, TransactionType, debitField, creditField, dateField, detailField, transactionNoField)
    Dim sqlObj As clsSQL
    Dim sqlStr As String
    Dim rowsAffected As Integer
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = sourceQuery
        .AddFilter "Property = " & PropertyListID
        .fields = "Property AS PropertyListID, " & _
                  dateField & " AS TransactionDate, " & _
                  transactionNoField & " AS TransactionNumber, " & _
                  "PropertyEntityCaption AS Name, " & _
                  detailField & " AS Purpose, " & _
                  "0 AS " & debitField & ", " & _
                  "Amount AS " & creditField & ", " & _
                  Esc(TransactionType) & " AS TransactionType"
        sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "INSERT"
        .Source = "tblPropertyLedgers"
        .fields = "PropertyListID,TransactionDate,TransactionNumber,Name,Purpose," & debitField & "," & creditField & ",TransactionType"
        .InsertSQL = sqlStr
        .InsertFilterField = "PropertyListID,TransactionDate,TransactionNumber,Name,Purpose," & debitField & "," & creditField & ",TransactionType"
        rowsAffected = .Run
    End With
End Sub

