Attribute VB_Name = "Tenant Mod"
Option Compare Database
Option Explicit

Public Function TenantCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            SetEntityCategoryDefault frm, "Tenant", False
            frm.OnLoad = "=EntityDEFormLoad([Form],""Tenant"")"
        Case 5: ''Datasheet Form
            SetEntityCategoryDefault frm, "Tenant"
        Case 6: ''Main Form
            SetMainFormProperty frm, "Tenant"
        Case 7: ''Tabular Report
    End Select

End Function


Public Function PropertyTenantEntityIDAfterUpdate(frm As Form)

    frm("Address") = frm("StreetAddress")
    
End Function
