Attribute VB_Name = "Model2 Mod"
Option Compare Database
Option Explicit

Public Function GenerateAdditionalOptionButton(frm As Form, ModelFieldID, subformName, pgName)
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset, y As Long
    Dim ModelID
    
    y = frm(pgName).top + 100
    
    ModelID = ELookup("tblModelFields", "ModelFieldID = " & ModelFieldID, "ModelID")
    
    ''Collapsed button so this is a combo box
    ''Create a combo box
    ''Left position should account for the label "Action:"
    ''55 is the space between controls,
    Dim lblWidth, maxX As Long: lblWidth = 1000
    maxX = frm(subformName).width - 3100
    
    Dim ctl As Control
    Set ctl = CreateControl(frm.Name, acComboBox, , pgName, , maxX, y, 3000, 400)
    ''Set the Default Control Properties Here
    SetControlProperties ctl
    ''Additional Property make the RowSource to be the SQLStr, ColumnCount to 2, ColumnWidths to 0;1
    ''Set the Height to be the same height as the buttons
    sqlStr = "SELECT ModelButtonID,ModelButton FROM tblModelButtons WHERE ModelID = " & ModelID & _
             " AND HideOnMain <> -1 ORDER BY ModelButtonOrder ASC, ModelButtonID"
    
    Dim cboName, lblName, btnName As String
    cboName = "cbo" & subformName & "FormActions": lblName = "lbl" & subformName & "FormActions"
    btnName = "cmdRun" & subformName & "FormActions"
    ctl.Name = cboName
    ctl.rowSource = sqlStr
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "0;1"
    ctl.height = 400
    ctl.TopMargin = 75
    ctl.LeftMargin = 75
    ctl.FontBold = True
    ctl.HorizontalAnchor = acHorizontalAnchorRight
    
    ''Render the label here
    Set ctl = CreateControl(frm.Name, acLabel, , cboName, , maxX - 55 - lblWidth, y, lblWidth, 400)
    ''Set the Default Control Properties Here
    SetControlProperties ctl
    ctl.Name = lblName
    ctl.Caption = "Actions: "
    ctl.TextAlign = 3
    ctl.height = 400
    ctl.TopMargin = 75
    ctl.LeftMargin = 75
    ctl.FontBold = True
    ctl.HorizontalAnchor = acHorizontalAnchorRight
    
    maxX = frm(cboName).left + frm(cboName).width + 55
    RenderButton maxX, y, "Run", 23, frm, btnName, pgName
    
    btnName = "cmd" & btnName
    frm(btnName).width = frm(btnName).width / 2
    frm(btnName).HorizontalAnchor = acHorizontalAnchorRight
    
    ''Resize the page width to be that of the subform but with a little but of margin
    frm(pgName).width = frm(subformName).width + 200
    
    frm(btnName).OnClick = "=RunFormActions([Form],[" & cboName & "], " & EscapeString(subformName) & ")"
    
    
End Function
