Attribute VB_Name = "Contact Mod"
Option Compare Database
Option Explicit

Public Function ContactCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            SetEntityCategoryDefault frm, "Contact", False
            frm.OnLoad = "=EntityDEFormLoad([Form],""Contact"")"
        Case 5: ''Datasheet Form
            SetEntityCategoryDefault frm, "Contact"
        Case 6: ''Main Form
            SetMainFormProperty frm, "Contact"
        Case 7: ''Tabular Report
    End Select

End Function
