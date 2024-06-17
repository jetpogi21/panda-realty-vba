Attribute VB_Name = "Report Helper"
Option Compare Database
Option Explicit

Public Function RunReport(rptCaption)

    Select Case rptCaption
        Case "Recepients":
            DoCmd.OpenForm "frmChildRecipientReport", , , , acFormAdd
        Case "Donation Requests":
            DoCmd.OpenForm "frmDonationRequestReport", , , , acFormAdd
        Case "Books And Petty Fees":
            DoCmd.OpenForm "frmPettyFeeReport", , , , acFormAdd
        Case "Textbook Stationeries":
            DoCmd.OpenForm "frmTextbookStationeryReport", , , , acFormAdd
        Case "Voucher Allocation Amounts":
            DoCmd.OpenForm "frmVoucherAllocationAmountReports", , , , acFormAdd
        Case "Uniform List":
            DoCmd.OpenForm "frmChildUniformReport", , , , acFormAdd
    End Select
    
End Function

Public Function GetOfficeInfo(OfficeID) As String
 
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblOffices WHERE OfficeID = " & OfficeID)
    
    Dim officeArr As New clsArray
    officeArr.Add rs.fields("OfficeAddress")
    officeArr.Add rs.fields("ContactNumber")
    officeArr.Add rs.fields("EmailAddress")
    
    GetOfficeInfo = officeArr.JoinArr(" | ")
    
End Function

Public Function SaveReportAsPDF(reportName, outputPath) As Boolean
    On Error Resume Next
    
    ' Output the report as PDF
    DoCmd.OutputTo acOutputReport, reportName, acFormatPDF, outputPath
    
    ' Check for errors during the OutputTo operation
    If Err.number <> 0 Then
        SaveReportAsPDF = False
        MsgBox "Error saving report as PDF: " & Err.Description, vbExclamation
    Else
        SaveReportAsPDF = True
        'MsgBox "Report saved as PDF successfully!", vbInformation
    End If
    
    On Error GoTo 0
End Function

