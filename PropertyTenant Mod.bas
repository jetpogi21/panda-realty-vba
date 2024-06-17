Attribute VB_Name = "PropertyTenant Mod"
Option Compare Database
Option Explicit

Public Function PropertyTenantCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            
        Case 5: ''Datasheet Form
            ''frm("TenantID").OnNotInList = "=PropertyTenantTenantIDNotInList([Form])"
            SetPropertyEntityForm frm, "Tenant"
        Case 6: ''Main Form
            
        Case 7: ''Tabular Report
    End Select

End Function
