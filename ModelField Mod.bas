Attribute VB_Name = "ModelField Mod"
Option Compare Database
Option Explicit

Public Function ModelFieldCreate(frm As Form, FormTypeID)
    
    If FormTypeID = 5 Then ModelFieldDSCreate frm
    Select Case FormTypeID
        Case 4, 5:
            AttachFunctions frm
            'frm("ModelID").RowSource = "SELECT ModelID, Model FROM tblModels WHERE UserQueryFields = 0"
    End Select
    
End Function

Private Sub ModelFieldDSCreate(frm As Form)
    
    frm.OrderBy = "FieldOrder ASC, ModelFieldID ASC"
    frm.SubPageOrder.DefaultValue = ""
    
End Sub

Private Sub AttachFunctions(frm As Form)

    frm.ParentModelID.AfterUpdate = "=ModelFieldParentModelIDChange([Form])"

End Sub

Public Function ModelFieldParentModelIDChange(frm As Form)
    
    ''On change of ParentModelID VerboseName, ForeignKey, Indexed should be automatically filled-up if it is not null
    Dim ParentModelID, ctl As Control
    ParentModelID = frm("ParentModelID")
    Set ctl = frm("ParentModelID")
    If Not IsNull(ParentModelID) Then
        frm("VerboseName") = AddSpaces(ctl.Column(1))
        frm("ForeignKey") = ctl.Column(1)
        frm("IsIndexed") = -1
        frm("FieldTypeID") = dbLong
    End If

End Function
