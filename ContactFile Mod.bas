Attribute VB_Name = "ContactFile Mod"
Option Compare Database
Option Explicit

Public Function ContactFileCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            SetEntityFileForm frm, FormTypeID, "Contact"
        Case 5: ''Datasheet Form
            SetEntityFileForm frm, FormTypeID, "Contact"
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function
