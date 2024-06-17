Attribute VB_Name = "PropertyListTemp Mod"
Option Compare Database
Option Explicit

Public Function PropertyListTempCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function


Public Function GetSalePrice(SalePrice) As Double
    
    GetSalePrice = 0
    
    If IsNumeric(SalePrice) Then GetSalePrice = SalePrice
    
End Function
