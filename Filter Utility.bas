Attribute VB_Name = "Filter Utility"
Option Compare Database
Option Explicit
Public FilterCaption As String

Public Function FilterReportBetween(frm As Form, rptName, dateField, dateCaption, recordsetName)
    
    Dim StartDate, endDate
    StartDate = frm("startDate")
    endDate = frm("endDate")
    
    If ExitIfTrue(IsNull(StartDate), "Please supply a start date..") Then Exit Function
    If ExitIfTrue(IsNull(endDate), "Please supply an end date..") Then Exit Function
    
    FilterCaption = dateCaption & " From " & StartDate & " and " & endDate
    
    If ExitIfTrue(ECount(recordsetName, dateField & " BETWEEN #" & StartDate & "# And #" & endDate & "#") = 0, "There is no record to show..") Then Exit Function
    
    DoCmd.OpenReport rptName, acViewPreview, , dateField & " BETWEEN #" & StartDate & "# And #" & endDate & "#"

End Function
