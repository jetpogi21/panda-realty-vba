Attribute VB_Name = "CustomFormDimension Mod"
Option Compare Database
Option Explicit

Public Function CustomFormDimensionCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function GetTheCurrentDimension(frm As Form)
    
    Dim FormName: FormName = frm("FormName")
    
    Dim subform As Form: Set subform = frm("subCustomFormDimensionControls").Form
    
    Dim CustomFormDimensionID: CustomFormDimensionID = subform("CustomFormDimensionID")
    Dim ControlName: ControlName = subform("ControlName")
    
    If ExitIfTrue(isFalse(ControlName), "There's no valid control name") Then Exit Function
      
    If ExitIfTrue(Not DoesFormExist(FormName), "Form " & Esc(FormName) & " doesn't exist..") Then Exit Function
    
    DoCmd.OpenForm FormName, acDesign
    Set frm = Forms(FormName)
    
    If ExitIfTrue(Not DoesPropertyExists(frm, ControlName), "Control " & Esc(ControlName) & "  doesn't exist..") Then Exit Function
    
    ''Update the dimensions
    Dim ctl As Control: Set ctl = frm(ControlName)
    subform("Top") = ctl.top
    subform("Left") = ctl.left
    subform("Width") = ctl.width
    subform("Height") = ctl.height
    
    MsgBox "Dimensions successfully updated.."
    
End Function

Private Function DoesFormExist(frmName) As Boolean

    On Error GoTo ErrHandler:
    DoCmd.OpenForm frmName, acDesign
    
    DoCmd.Close acForm, frmName, acSaveNo
    
    DoesFormExist = True
    Exit Function
    
ErrHandler:
    Exit Function
    
End Function
