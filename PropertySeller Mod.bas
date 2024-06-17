Attribute VB_Name = "PropertySeller Mod"
Option Compare Database
Option Explicit

Public Function PropertySellerCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            
        Case 5: ''Datasheet Form
            ''frm("SellerID").OnNotInList = "=PropertySellerSellerIDNotInList([Form])"
            SetPropertyEntityForm frm, "Seller"
        Case 6: ''Main Form
            
        Case 7: ''Tabular Report
    End Select

End Function

Public Function PropertySellerAfterUpdate()

    Dim frm As Form
    Set frm = Forms("frmPropertyList")
    
    Dim OwnerNoArr As New clsArray, OwnerNo
    Dim FieldToFetchArr As New clsArray, FieldToFetch
    
    OwnerNoArr.arr = "1,2,3"
    FieldToFetchArr.arr = "PhoneNumber,EmailAddress"
    
    Dim UpdatedField, OwnerTextField, OwnerName
    
    For Each OwnerNo In OwnerNoArr.arr
        
        For Each FieldToFetch In FieldToFetchArr.arr
        
            FetchSellerDetail frm, OwnerNo, FieldToFetch
            
        Next FieldToFetch
        
    Next OwnerNo

End Function
