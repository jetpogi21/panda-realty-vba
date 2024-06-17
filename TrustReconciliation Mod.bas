Attribute VB_Name = "TrustReconciliation Mod"
Option Compare Database
Option Explicit

Public Function TrustReconciliationCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function Open_mainFormFromAmount(frm As Form, Optional transaction As String = "Receipt")
    
    Dim TrustReconciliationYear, monthID
    monthID = frm("MonthID"): TrustReconciliationYear = frm("TrustReconciliationYear")
    
    Dim frmName: frmName = "mainTrust" & transaction & "Modal"
    DoCmd.OpenForm frmName, , , , , acHidden
    
    Dim dateField
    If transaction = "Receipt" Then
        dateField = "DateReceived"
    Else
        dateField = "DatePaid"
    End If
    Set frm = Forms(frmName).Controls("subform").Form
    frm.Filter = "Month(" & dateField & ") = " & monthID & " AND Year(" & dateField & ") = " & TrustReconciliationYear
    frm.FilterOn = True
    
    DoCmd.OpenForm frmName
    
    
End Function

Public Function frmTrustReconciliations_OnCurrent(frm As Form)
    
    SetTotalAmounts frm, "Receipt"
    SetTotalAmounts frm, "Payment"
    InsertClientLedgers frm
    Set_frmTrustReconciliations_Balance frm
    
End Function

Public Function frmTrustReconciliations_OnLoad(frm As Form)
    
    Dim width: width = 3
    ''subTrustDeposit,subTrustCheque
    With frm("subTrustDeposit").Form
        .Controls("OutstandingType").ColumnHidden = True
        .Controls("Description").ColumnWidth = 1000 * width
        .Controls("OutstandingAmount").ColumnWidth = 1000 * width
        If IsFormOpen("frmTrustReconciliations") Then
            .AfterUpdate = "=Set_frmTrustReconciliations_Balance([Forms]![frmTrustReconciliations])"
        End If
    End With
    
    With frm("subTrustCheque").Form
        .Controls("OutstandingType").ColumnHidden = True
        .Controls("Description").ColumnWidth = 1000 * width
        .Controls("OutstandingAmount").ColumnWidth = 1000 * width
        If IsFormOpen("frmTrustReconciliations") Then
            .AfterUpdate = "=Set_frmTrustReconciliations_Balance([Forms]![frmTrustReconciliations])"
        End If
    End With
    
    With frm("subClientLedger").Form
        .Controls("PropertyOwnerSurname").ColumnWidth = 1500 * width
        .Controls("Balance").ColumnWidth = 500 * width
        .Controls("PropertyListID").ColumnHidden = True
    End With
    
    
End Function

Public Function AreValuesEqual(value1 As Variant, value2 As Variant, value3 As Variant) As Boolean
    If value1 = value2 And value2 = value3 Then
        AreValuesEqual = True
    Else
        AreValuesEqual = False
    End If
End Function

Public Function SetTotalAmounts(frm As Form, Optional mode = "Receipt")
    
    Dim monthID, TrustReconciliationYear
    
    monthID = frm("MonthID")
    TrustReconciliationYear = frm("TrustReconciliationYear")
    
    Dim tblName As String, FieldName, dateField
    If mode = "Receipt" Then
        tblName = "tblTrustReceipts"
        dateField = "DateReceived"
        FieldName = "txtReceipts"
    Else
        tblName = "tblTrustPayments"
        dateField = "DatePaid"
        FieldName = "txtPayments"
    End If
    Dim sum
    If isFalse(monthID) Or isFalse(TrustReconciliationYear) Then
        sum = 0
    Else
        sum = ESum2(tblName, "Month(" & dateField & ") = " & monthID & " AND Year(" & dateField & ") = " & TrustReconciliationYear, "Amount")
    End If
    
    frm(FieldName).ControlSource = "=" & sum
    
End Function

Public Function InsertClientLedgers(frm As Form)
    
    
    Dim monthID, TrustReconciliationYear, TrustReconciliationID, EndOfMonth
    
    monthID = frm("MonthID")
    TrustReconciliationYear = frm("TrustReconciliationYear")
    TrustReconciliationID = frm("TrustReconciliationID")
    EndOfMonth = GetEndOfMonth(monthID, TrustReconciliationYear)
    
    If isFalse(TrustReconciliationID) Then Exit Function
    
    If isFalse(EndOfMonth) Then Exit Function
    
    DoCmd.RunCommand acCmdSaveRecord
    
    Dim propertySQL, sumReceiptSQL, sumPaymentSQL
    propertySQL = "Select Distinct Property FROM tblTrustReceipts WHERE DateReceived <= #" & EndOfMonth & "#"
    propertySQL = propertySQL & " UNION Select Distinct Property FROM tblTrustPayments WHERE DatePaid <= #" & EndOfMonth & "#"
    
    sumReceiptSQL = "Select Property, Sum(Amount) As SumReceipt FROM  tblTrustReceipts WHERE DateReceived <= #" & EndOfMonth & "# GROUP BY Property"
    sumPaymentSQL = "Select Property, Sum(Amount) As SumPayment FROM  tblTrustPayments WHERE DatePaid <= #" & EndOfMonth & "# GROUP BY Property"
    
    RunSQL "DELETE FROM tblTrustReconciliationLedgers WHERE TrustReconciliationID = " & TrustReconciliationID
    ' Now you can use propertySQL, sumReceiptSQL, and sumPaymentSQL as needed in your code
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = propertySQL
          .fields = TrustReconciliationID & " AS TrustReconciliationID,tempProperties.Property,CdblNz(SumReceipt) - CdblNz(SumPayment) As Balance"
          .Joins.Add GenerateJoinObj(sumReceiptSQL, "Property", "tempReceipt", , "LEFT")
          .Joins.Add GenerateJoinObj(sumPaymentSQL, "Property", "tempPayment", , "LEFT")
          .SourceAlias = "tempProperties"
          sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "tblTrustReconciliationLedgers"
          .fields = "TrustReconciliationID,Property,Balance"
          .InsertSQL = sqlStr
          .InsertFilterField = "TrustReconciliationID,Property,Balance"
          rowsAffected = .Run
    End With
    
    frm("subClientLedger").Form.Requery
    
End Function


Public Function frmTrustReconciliations_PeriodControls_AfterUpdate(frm As Form)
    
    SetTotalAmounts frm, "Receipt"
    SetTotalAmounts frm, "Payment"
    InsertClientLedgers frm
    Set_frmTrustReconciliations_Balance frm
    
End Function

Public Function Set_frmTrustReconciliations_Balance(frm As Form)

    frm("BookBalance") = frm("txtClosingCashbookBalance")
    frm("ReconciledBankBalance") = frm("txtBankBalance")
    frm("LedgerBalance") = frm("txtLedgerBalance")
    
End Function

Public Function Save_frmTrustReconciliations(frm As Form)
    Set_frmTrustReconciliations_Balance frm
    Save2 frm, "TrustReconciliation", 1
End Function



