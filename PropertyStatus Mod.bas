Attribute VB_Name = "PropertyStatus Mod"
Option Compare Database
Option Explicit

Public Function PropertyStatusCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function PropertyStatusCaption(PropertyStatusCode, PropertyStatus) As String
    
    Dim PropertyStatusArr As New clsArray
    
    If Not IsNull(PropertyStatusCode) Then PropertyStatusArr.Add PropertyStatusCode
    ''If Not IsNull(PropertyStatus) Then PropertyStatusArr.Add PropertyStatus
   
    If PropertyStatusArr.Count > 0 Then PropertyStatusCaption = PropertyStatusArr.JoinArr("-")

End Function
