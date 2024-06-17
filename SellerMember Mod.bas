Attribute VB_Name = "SellerMember Mod"
Option Compare Database
Option Explicit

Public Function SellerMemberCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
            SetEntityMemberForm frm
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function
