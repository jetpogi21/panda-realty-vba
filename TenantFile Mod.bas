Attribute VB_Name = "TenantFile Mod"
Option Compare Database
Option Explicit

Public Function TenantFileCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            SetEntityFileForm frm, FormTypeID, "Tenant"
        Case 5: ''Datasheet Form
            SetEntityFileForm frm, FormTypeID, "Tenant"
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function
