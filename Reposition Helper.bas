Attribute VB_Name = "Reposition Helper"
Option Compare Database
Option Explicit


Public Sub RepositionControls(frm As Form, proportionArr As clsArray, controlArr As clsArray, x, y, totalWidth, Optional colSpaceWidth = 50)

    Dim proportionTotal, i, proportion, controlWidth
    proportionTotal = GetProportionTotal(proportionArr)
    
    For i = 0 To proportionArr.Count - 1

        proportion = CDbl(proportionArr.arr(i)) / proportionTotal
        controlWidth = (totalWidth - ((proportionArr.Count - 1) * colSpaceWidth * 2)) * proportion
        
        If controlArr.arr(i) <> "empty" Then
            frm(controlArr.arr(i)).left = x
            frm(controlArr.arr(i)).top = y
            frm(controlArr.arr(i)).width = controlWidth
        End If
        
        x = x + (colSpaceWidth * 2) + controlWidth
       
    Next i
    
End Sub
