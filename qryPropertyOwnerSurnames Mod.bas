Attribute VB_Name = "qryPropertyOwnerSurnames Mod"
Option Compare Database
Option Explicit

Public Function qryPropertyOwnerSurnamesCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function PropertySelector_SelectProperty(frm As Form)

    Dim PropertyListID: PropertyListID = frm("PropertyListID")
    
    If IsFormOpen("frmTrustReceipts") Then
        Forms("frmTrustReceipts").property = PropertyListID
        Property_frmTrustReceipts_AfterUpdate Forms("frmTrustReceipts")
    End If
    
    DoCmd.Close acForm, frm.Parent.Name, acSaveNo
    
End Function


Public Function FilterPropertySelector(frm As Form)
    
    Dim txtSearch: txtSearch = frm("txtSearch")
    If isFalse(txtSearch) Then
        ClearPropertySelectorFilter frm
        Exit Function
    End If
    
    Dim aText As Variant, splittedText As New clsArray:  splittedText.arr = txtSearch
    Dim filters As New clsArray
    
    For Each aText In splittedText.arr
        filters.Add "StreetAddress Like " & EscapeString("*" & aText & "*")
    Next aText
    
    frm("subform").Form.Filter = filters.JoinArr(" AND ")
    frm("subform").Form.FilterOn = True
    
End Function

Public Function ClearPropertySelectorFilter(frm As Form)
    
    frm("subform").Form.FilterOn = False
    
End Function
