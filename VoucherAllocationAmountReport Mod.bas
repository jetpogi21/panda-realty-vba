Attribute VB_Name = "VoucherAllocationAmountReport Mod"
Option Compare Database
Option Explicit

Public Function VoucherAllocationAmountReportPreview(frm As Form)
        
    If Not areDataValid2(frm, "VoucherAllocationAmountReport") Then Exit Function
    
    Dim YearAllocated
    
    YearAllocated = frm("YearAllocated")
    
    ''Check First if there will be record to be printed
    If Not isPresent("qryVoucherAllocationAmounts", "YearAllocated = " & YearAllocated) Then
        ShowError "There is no record.."
        Exit Function
    End If
    
    DoCmd.OpenReport "rptTabVoucherAllocationAmounts", acViewPreview, , "YearAllocated = " & YearAllocated
    
End Function
