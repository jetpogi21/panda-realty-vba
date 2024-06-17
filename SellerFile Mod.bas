Attribute VB_Name = "SellerFile Mod"
Option Compare Database
Option Explicit

Public Function SellerFileCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            SetEntityFileForm frm, FormTypeID, "Seller"
        Case 5: ''Datasheet Form
            SetEntityFileForm frm, FormTypeID, "Seller"
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function
