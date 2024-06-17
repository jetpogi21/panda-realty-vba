Attribute VB_Name = "MaintenanceMenu Mod"
Option Compare Database
Option Explicit

Public Function MaintenanceMenuCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function OpenMaintenanceForm(frm As Form)
    
    Dim frmName
    frmName = frm("FormName")
    
    DoCmd.OpenForm frmName
    
End Function
