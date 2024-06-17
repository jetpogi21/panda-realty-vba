Attribute VB_Name = "PropertyForm Mod"
Option Compare Database
Option Explicit

Public Function PropertyFormCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function


Public Function GetFormPath(frm As Form)

    Dim fs As Object
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim fullPath, fileName
    With FileDialog(msoFileDialogFilePicker)
        'Makes sure the user can select only one file
        .AllowMultiSelect = False
        .filters.Clear
        'Show the dialog box
        If ExitIfTrue(Not .Show, "Please select a valid file path..") Then Exit Function
        
        'Store in fullpath variable
        fullPath = .SelectedItems.item(1)
        
    End With
    
    
    frm("FormLink") = fullPath
    
    ''FormLink
    
End Function
