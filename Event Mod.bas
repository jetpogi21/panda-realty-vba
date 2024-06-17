Attribute VB_Name = "Event Mod"
Option Compare Database
Option Explicit

Public Function EventCheckoutCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 6:
            frm("subform").SourceObject = "Report.rptTabEventCheckouts"
            ''Create a new control
            Dim ctl As Control
            Set ctl = CreateControl(frm.Name, acCommandButton, , , , 0, 0, 0, 0)
            SetControlProperties ctl
            ctl.Name = "cmdPrintReport"
            ctl.Caption = "Print Report"
            ctl.height = frm("cmdFilter").height
            ctl.OnClick = "=PrintReport([Form])"
            ctl.HorizontalAnchor = acHorizontalAnchorRight
            
            Dim proportionArr As New clsArray, controlArr As New clsArray, x, y, totalWidth
            proportionArr.arr = "1,1,1"
            controlArr.arr = "cmdFilter,cmdClear,cmdPrintReport"
            x = frm("cmdFilter").left
            y = frm("cmdFilter").top
            totalWidth = frm("cmdClear").left + frm("cmdClear").width - frm("cmdFilter").left
            
            RepositionControls frm, proportionArr, controlArr, x, y, totalWidth
    End Select
    
End Function


Public Function EventCreate(frm As Form, FormTypeID)

    If FormTypeID = 4 Then
        ''Create a combo box and a button at the side
        ''parent is pgCheckouts
        ''reference cmdDeleteCheckouts Top and Height
        ''subCheckouts x + width
        ''width is two buttons
        CreateAssetGroupBulkCheckoutForm frm

        
    End If
    
End Function

Public Function DoBulkCheckout(frm As Form)
    
    Dim EventID, EventRunDay, AssetGroupID
    If Not areDataValid2(frm, "Event") Then Exit Function
    
    frm.Dirty = False
    EventID = frm("EventID")
    EventRunDay = frm("EventRunDay")
    AssetGroupID = frm("AssetGroupID")
    
    If ExitIfTrue(IsNull(AssetGroupID), "Please select an asset group..") Then Exit Function
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Dim UserID
    
    UserID = g_UserID
    If IsEmpty(UserID) Then UserID = 1
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblAssetAssetGroups"
        .AddFilter "AssetGroupID = " & AssetGroupID
        .fields = "AssetID, " & _
                  EventID & " As EventID, " & _
                  "#" & SQLDate(DateAdd("d", EventRunDay, Date)) & "# As CheckoutDueDate, " & _
                  1 & " As CheckoutQty, " & _
                  "#" & SQLDate(Date) & "# As CheckoutDate, " & _
                  UserID & " As CreatedBy"
        sqlStr = .sql
        If ExitIfTrue(.Count = 0, "The asset group you selected is empty..") Then Exit Function
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = sqlStr
        .AddFilter "CheckoutID Is NULL And temp.EventID"
        .fields = "temp.*"
        .Joins.Add GenerateJoinObj("tblCheckouts", "AssetID,EventID", , , "LEFT")
        .SourceAlias = "temp"
        sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "INSERT"
        .Source = "tblCheckOuts"
        .fields = "CheckoutDueDate, CheckoutQty, CheckoutDate, AssetID, EventID,CreatedBy"
        .InsertSQL = sqlStr
        .InsertFilterField = "CheckoutDueDate, CheckoutQty, CheckoutDate, AssetID, EventID,CreatedBy"
        rowsAffected = .Run
    End With
    
    frm("subCheckouts").Requery
    
    If rowsAffected > 0 Then MsgBox "Bulk checkout successfull.."
    
End Function

Public Function DoBulkCheckin(frm As Form)
        
    Dim EventID, EventRunDay, AssetGroupID
    If Not areDataValid2(frm, "Event") Then Exit Function
    
    frm.Dirty = False
    EventID = frm("EventID")
    EventRunDay = frm("EventRunDay")
    AssetGroupID = frm("AssetGroupID")
    
    Dim UserID
    UserID = g_UserID
    If IsEmpty(UserID) Then UserID = 1
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "UPDATE"
        .Source = "tblCheckouts"
        .SetStatement = "TotalCheckinQty = CheckoutQty - TotalCheckinQty"
        .AddFilter "EventID = " & EventID & " And CheckoutQty - TotalCheckinQty > 0"
        rowsAffected = .Run
    End With
    
    frm("subCheckouts").Requery
    
    If rowsAffected > 0 Then MsgBox "Bulk checkin successfull.."
        
End Function

Public Function CreateAssetGroupBulkCheckoutForm(frm As Form)
    
    Dim ctl As Control, y, x, totalWidth, colSpaceWidth, height
    y = frm("cmdDeleteCheckouts").top
    totalWidth = frm("cmdDeleteCheckouts").width * 3
    x = frm("subCheckouts").left + frm("subCheckouts").width - totalWidth
    colSpaceWidth = 50
    height = frm("cmdDeleteCheckouts").height
    
    Set ctl = CreateControl(frm.Name, acComboBox, , "pgCheckouts", , x, y, 0, height)
    SetControlProperties ctl
    ctl.Name = "AssetGroupID"
    'ctl.RowSource = sqlStr
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "0;1"
    ctl.TopMargin = 75
    ctl.LeftMargin = 75
    ctl.height = height
    ctl.FontBold = True
    
    Dim sqlStr
    sqlStr = "SELECT AssetGroupID, AssetGroup FROM tblAssetGroups ORDER BY AssetGroup"
    ctl.rowSource = sqlStr
    
    Set ctl = CreateControl(frm.Name, acCommandButton, , "pgCheckouts", , x, y, 0, height)
    ctl.Name = "cmdBulkCheckout"
    ctl.Caption = "Bulk Checkout"
    ctl.OnClick = "=DoBulkCheckout([Form])"
    SetControlProperties ctl
    
    Set ctl = CreateControl(frm.Name, acCommandButton, , "pgCheckouts", , x, y, 0, height)
    ctl.Name = "cmdBulkCheckin"
    ctl.Caption = "Bulk Checkin"
    ctl.OnClick = "=DoBulkCheckin([Form])"
    SetControlProperties ctl
    
    Dim proportionArr As New clsArray, controlArr As New clsArray, proportionTotal, i
    Dim proportion
    
    proportionArr.arr = "8,4,4"
    controlArr.arr = "AssetGroupID,cmdBulkCheckout,cmdBulkCheckin"
    proportionTotal = GetProportionTotal(proportionArr)
    
    For i = 0 To proportionArr.Count - 1
    
        proportion = CDbl(proportionArr.arr(i)) / proportionTotal
        frm(controlArr.arr(i)).left = x
        frm(controlArr.arr(i)).top = y
        frm(controlArr.arr(i)).width = (totalWidth - ((proportionArr.Count - 1) * colSpaceWidth * 2)) * proportion
        frm(controlArr.arr(i)).HorizontalAnchor = acHorizontalAnchorRight
       
        x = x + (colSpaceWidth * 2) + frm(controlArr.arr(i)).width
       
    Next i
    
End Function


'Set ctl = CreateControl(frm.Name, acComboBox, , , , maxX, y, 3000, 400)
'            ''Set the Default Control Properties Here
'            SetControlProperties ctl
'            ''Additional Property make the RowSource to be the SQLStr, ColumnCount to 2, ColumnWidths to 0;1
'            ''Set the Height to be the same height as the buttons
'            sqlStr = "SELECT ModelButtonID,ModelButton FROM tblModelButtons WHERE ModelID = " & ModelID & _
'                     " AND HideOnMain <> -1 ORDER BY ModelButtonOrder ASC, ModelButtonID"
'            ctl.Name = "cboFormActions"
            
'            ctl.HorizontalAnchor = acHorizontalAnchorRight
