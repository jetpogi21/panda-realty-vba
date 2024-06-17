Attribute VB_Name = "FixPropertSellers Mod"
Option Compare Database
Option Explicit

Public Function FixPropertSellersCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function FixPropertSellers()
    
    RunSQL "DELETE * FROM tblEntities WHERE IsSeller = -1"
    ImportPropertySeller , True
    
End Function


