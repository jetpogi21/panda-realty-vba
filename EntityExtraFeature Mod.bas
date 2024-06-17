Attribute VB_Name = "EntityExtraFeature Mod"
Option Compare Database
Option Explicit

Public Function EntityExtraFeatureCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function


Public Function EntityExtraFeatureFeatureAfterUpdate(frm As Form)
    
    Dim fieldType
    fieldType = frm("Feature").Column(1)
    
    If IsNull(fieldType) Then frm("Value") = ""
    If fieldType = "Number" Then frm("Value") = 0
    If fieldType = "Text" Then frm("Value") = ""
    If fieldType = "Date" Then frm("Value") = Format$(Date, "dd-mmm-yy")
    
End Function
