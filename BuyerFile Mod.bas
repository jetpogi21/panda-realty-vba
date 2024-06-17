Attribute VB_Name = "BuyerFile Mod"
Option Compare Database
Option Explicit

Public Function BuyerFileCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            SetEntityFileForm frm, FormTypeID, "Buyer"
        Case 5: ''Datasheet Form
            SetEntityFileForm frm, FormTypeID, "Buyer"
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function SetEntityFileForm(frm As Form, FormTypeID, EntityCategory)
    
    frm.BeforeUpdate = "=SaveFormData2([Form],""EntityFile"")"
    
    frm("FileType").AllowValueListEdits = False
    'frm("EntityID").ListItemsEditForm = False
    
    If FormTypeID = 5 Then
        frm.AllowAdditions = False
        frm.AllowEdits = False
        frm("EntityFileLink").DisplayAsHyperlink = 1
        'frm("EntityFileLink").IsHyperlink = True
        frm("EntityFileLink").Properties("DatasheetCaption") = "File Link"
        frm("EntityFileLink").OnDblClick = "=FollowFormHyperlink([Form],""EntityFileLink"")"
    Else
        frm("lblEntityID").Caption = EntityCategory & " Name"
        frm("Label5").Caption = "File Link"
    End If
    
    Dim sqlStr As String
    sqlStr = "SELECT EntityID, EntityName FROM qryEntities WHERE EntityCategoryName = " & EscapeString(EntityCategory) & " ORDER BY EntityName"
    SetFormControlProperties frm, "EntityID", "RowSource", sqlStr
    SetFormControlProperties frm, "EntityID", "ListItemsEditForm", ""

    
End Function

Public Function SetFormControlProperties(frm As Form, ControlName, PropertyName, PropertyValue)

    On Error Resume Next
    frm(ControlName).Properties(PropertyName) = PropertyValue

End Function


