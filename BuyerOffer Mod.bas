Attribute VB_Name = "BuyerOffer Mod"
Option Compare Database
Option Explicit

Public Function BuyerOfferCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet FormtblE
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function BuyerOfferCreateAfterUpdate(frm As Form)
    
    Dim vTimestamp, IsAccepted
    vTimestamp = frm("Timestamp")
    IsAccepted = frm("IsAccepted")
    
    If IsFormOpen("frmPropertyList") Then
    
        Set frm = Forms("frmPropertyList")
        frm("subLatestBuyerOffers").Form.Requery
        
        If IsAccepted Then
            frm("txtSuccessfulDate") = Format(vTimestamp, "dd-mmm-yy")
        Else
            frm("txtSuccessfulDate") = Null
        End If
        
        frm.Controls("txtSuccessfulBuyer").Requery
   
    End If
    
End Function

Public Function BuyerOfferValidation(frm As Form) As Boolean
    
    Dim PropertyListID, IsAccepted, BuyerOfferID
    PropertyListID = frm("PropertyListID")
    IsAccepted = frm("IsAccepted")
    BuyerOfferID = frm("BuyerOfferID")
    
    BuyerOfferValidation = True
    
    If IsAccepted Then
        
        If frm.NewRecord Then
            If isPresent("qryBuyerOffers", "PropertyListID = " & PropertyListID & " AND IsAccepted") Then
                MsgBox "There should only be one offer accepted in each property."
                frm("IsAccepted") = False
                BuyerOfferValidation = False
                Exit Function
            End If
        Else
            If isPresent("qryBuyerOffers", "PropertyListID = " & PropertyListID & " AND IsAccepted AND BuyerOfferID <> " & BuyerOfferID) Then
                MsgBox "There should only be one offer accepted in each property."
                frm("IsAccepted") = False
                BuyerOfferValidation = False
                Exit Function
            End If
        End If
        
    End If
    
End Function
