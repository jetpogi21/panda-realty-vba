Attribute VB_Name = "Ribbon Functions"
Option Compare Database
Option Explicit

Public Sub OpenFormFromRibbon(ctl As IRibbonControl)
    DoCmd.OpenForm ctl.Id
End Sub

Public Sub changeGlobal(ctl As IRibbonControl)

    'UPDATE STATEMENT
    Dim sqlObj As New clsSQL, fltrObj As New clsArray
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "UPDATE"
        .Source = "tblGlobalSettings"
        .SetStatement = "GlobalSettingValue = " & EscapeString(CurrentProject.Path, "tblGlobalSettings", "GlobalSettingValue")
        .AddFilter "GlobalSetting = ""systemProductImages_FilePath"""
        .Run
    End With
    
    fltrObj.arr = "Application_ImportCSV_FilePath,rptShelfLocationLabels,rptPackSheets_FilePath,rptPickSheets_FilePath," & _
        "rptIntermediateLabels_FilePath,rptPrintH_FilePath"
    
    Dim arrItem
    For Each arrItem In fltrObj.arr
    
        Set sqlObj = New clsSQL
        With sqlObj
            .SQLType = "UPDATE"
            .Source = "tblGlobalSettings"
            .SetStatement = "GlobalSettingValue = " & EscapeString("C:\Users\user\Desktop\Printables\")
            .AddFilter "GlobalSetting = " & EscapeString(arrItem)
            .Run
        End With
    
    Next arrItem
    
End Sub
