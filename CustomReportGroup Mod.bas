Attribute VB_Name = "CustomReportGroup Mod"
Option Compare Database
Option Explicit

Public Function SetCustomReportFieldIDCbo(frm As Form)
    
On Error GoTo ErrHandler:
    Dim CustomReportID, sqlStr
    CustomReportID = frm.Parent.Form.CustomReportID
    
    If IsNull(CustomReportID) Then
        sqlStr = "SELECT CustomReportFieldID, CustomReportField FROM tblCustomReportFields ORDER BY CustomReportField"
    Else
        sqlStr = "SELECT CustomReportFieldID, CustomReportField FROM tblCustomReportFields WHERE CustomReportID = " & CustomReportID & _
                 " ORDER BY CustomReportField"
    End If
    
    
    frm("CustomReportFieldID").rowSource = sqlStr
    
    Exit Function
ErrHandler:

    CustomReportID = Null
    
End Function



