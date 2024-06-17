Attribute VB_Name = "Seller Mod"
Option Compare Database
Option Explicit

Public Function SellerCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            SetEntityCategoryDefault frm, "Seller", False
            frm.OnLoad = "=EntityDEFormOnLoad([Form],""Seller"")"
        Case 5: ''Datasheet Form
            SetEntityCategoryDefault frm, "Seller"
        Case 6: ''Main Form
            SetMainFormProperty frm, "Seller"
        Case 7: ''Tabular Report
    End Select

End Function

Public Function SellerDEFormOnload(frm As Form)

    EntityDEFormLoad frm, "Seller"
    frm("subPropertySellers").Form("OpenInRPP").ColumnHidden = True
    frm("subPropertySellers").Form("txtOpenInRPP").ColumnWidth = 1000
    
'    frm("subPropertyBuyers").Form.AllowAdditions = False
'    frm("subPropertyBuyers").Form.AllowEdits = False
'    frm("subPropertyBuyers").Form.AllowDeletions = False
'    frm("subPropertyBuyers").Form("txtOpenInRPP").columnWidth = 1000
'
'    frm("subBuyerOffers").Form.AllowAdditions = False
'    frm("subBuyerOffers").Form.AllowEdits = False
'    frm("subBuyerOffers").Form.AllowDeletions = False
    
End Function

'Public Function SellerAfterUpdate(frm As Form)
'
'    If IsFormOpen("frmPropertyList") Then
'        Forms("frmPropertyList")("subPropertySeller2").Form.Requery
'        Dim EntityID: EntityID = frm("EntityID")
'        ExportEntityToExcel Null, EntityID
'    End If
'
'End Function
