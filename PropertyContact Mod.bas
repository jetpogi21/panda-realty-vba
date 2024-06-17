Attribute VB_Name = "PropertyContact Mod"
Option Compare Database
Option Explicit

Public Function PropertyContactCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            
        Case 5: ''Datasheet Form
            ''frm("ContactID").OnNotInList = "=PropertyContactContactIDNotInList([Form])"
            SetPropertyEntityForm frm, "Contact"
        Case 6: ''Main Form
            
        Case 7: ''Tabular Report
    End Select

End Function
