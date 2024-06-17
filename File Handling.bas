Attribute VB_Name = "File Handling"
Option Compare Database
Option Explicit

Public Function GeneratePropertyRelatedPDF(frm As Form, PropertyListID, fileName, rptName, reportCaption, Optional notifyFilePath As Boolean = True) As String
 
    Dim PropertyDirectory As String
    PropertyDirectory = GetPropertyDirectory(frm, "PropertyListID")
    
    Dim filePath As String
    filePath = PropertyDirectory & fileName & ".pdf"
    
    DoCmd.OpenReport rptName, acViewDesign, , , acHidden
    Dim rpt As Report
    Set rpt = Reports(rptName)
    
    ''SetReportMargin rpt
    DoCmd.Close acReport, rptName, acSaveYes
    
    DoCmd.OpenReport rptName, acViewReport, , "PropertyListID = " & PropertyListID, acHidden
    
    SaveReportAsPDF rptName, filePath
    
    If notifyFilePath Then MsgBox reportCaption & " generated at: " & Esc(filePath)
    
    GeneratePropertyRelatedPDF = filePath
    
End Function

Public Function SanitizeFileName(fileName)
    
    Dim bannedArr As New clsArray, i
    bannedArr.arr = "#,%,&,{,},\,<,>,*,?,/,$,!,',"",:,@,+,`,|,="
    
    For Each i In bannedArr.arr
        fileName = replace(fileName, i, "-")
    Next i
    
    SanitizeFileName = fileName
    
End Function

Public Function IsFileOpen(fileName)

    Dim fileNum As Integer
    Dim errNum As Integer
    
    'Allow all errors to happen
    On Error Resume Next
    fileNum = FreeFile()
    
    'Try to open and close the file for input.
    'Errors mean the file is already open
    Open fileName For Input Lock Read As #fileNum
    Close fileNum
    
    'Get the error number
    errNum = Err
    
    'Do not allow errors to happen
    On Error GoTo 0
    
    'Check the Error Number
    Select Case errNum
    
        'errNum = 0 means no errors, therefore file closed
        Case 0
        IsFileOpen = False
     
        'errNum = 70 means the file is already open
        Case 70
        IsFileOpen = True
    
        'Something else went wrong
        Case Else
        IsFileOpen = errNum
    
    End Select

End Function


Function fileExists(filePath) As Boolean
 
    '--------------------------------------------------
    'Checks if a file exists (using the Dir function).
    '--------------------------------------------------
 
    On Error Resume Next
    If Len(filePath) > 0 Then
        If Not Dir(filePath, vbDirectory) = vbNullString Then fileExists = True
    End If
    On Error GoTo 0
 
End Function

'Public Function SetImageSource(SKU As String) As String
'
'    SKU = ConvertImageName(SKU)
'
'    Dim imgPath As String
'
'    imgPath = GetGlobalSetting("systemProductImages_FilePath") & SKU & ".jpg"
'
'    If Not FileExists(imgPath) Then
'      SetImageSource = CurrentProject.Path & "\default.jpg"
'      Exit Function
'    End If
'
'    SetImageSource = imgPath
'
'End Function

'Public Function ConvertImageName(ByVal SKU As String) As String
'
'    Dim regexOne As Object
'    Set regexOne = New RegExp
'
'    regexOne.Pattern = "[/:""""*?<>|]"
'    regexOne.Global = True
'    ConvertImageName = regexOne.Replace(SKU, "-")
'
'    Set regexOne = Nothing
'
'End Function

Public Function DirectoryExists(Path) As Boolean
On Error GoTo DirError:
    DirectoryExists = Dir(Path, vbDirectory) <> ""
    Exit Function
DirError:
    If Err.number = 52 Then
        DirectoryExists = False
    End If
End Function

Public Function SetDefaultDirectory(ctl As Control) As String

    ''Open the fileDialog selecting a directory
    Dim fd As FileDialog
    Dim strPath As String
    ' Set up the File Dialog.
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fd
        ' Allow user to make multiple selections in dialog box
        .AllowMultiSelect = False
             
        ' Set the title of the dialog box.
        .title = "Please select a directory"
        
        If .Show = True Then
            If .SelectedItems.Count > 0 Then
                'get the file path selected by the user
                ctl = .SelectedItems(1)
                Exit Function
            End If
        End If
         
    End With
    
    
    ctl = Environ$("USERPROFILE") & "\Downloads"
   
End Function

Public Function PromptDirectory(Optional defaultFileName = "", Optional FileType As String = "") As String

    ''Open the fileDialog selecting a directory
    Dim fd As FileDialog
    Dim strPath As String
    ' Set up the File Dialog.
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    Dim filePath As String
    
    With fd
        ' Allow user to make multiple selections in dialog box
        .AllowMultiSelect = False
        
        ' Set the title of the dialog box.
        .title = "Please select a save directory"
        
        If Not isFalse(defaultFileName) Then
            .InitialFileName = defaultFileName
        End If
        
        If .Show = True Then
            If .SelectedItems.Count > 0 Then
                'get the file path selected by the user
                PromptDirectory = .SelectedItems(1)
                Exit Function
            End If
        End If
         
    End With
    
    
    PromptDirectory = ""
   
End Function

Public Function PromptFile(Optional FileType As String = "") As String
    
    ''Open the fileDialog selecting a directory
    Dim fd As FileDialog
    Dim strPath As String
    ' Set up the File Dialog.
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Dim filePath As String
    With fd
        Dim filterStr, title
        
        If FileType <> "" Then
            Dim rs As Recordset
            Set rs = ReturnRecordset("SELECT * FROM tblFileDialogs WHERE FileType = '" & FileType & "'")
            fd.filters.Clear
            fd.filters.Add rs.fields("FileType"), rs.fields("FileFilters"): fd.title = rs.fields("Title")
        End If

        ' Allow user to make multiple selections in dialog box
        .AllowMultiSelect = False
        
        If .Show = True Then
            If .SelectedItems.Count > 0 Then
                'get the file path selected by the user
                PromptFile = .SelectedItems(1)
                Exit Function
            End If
        End If
         
    End With
    
    
    PromptFile = ""
   
End Function


Public Function directoryPath(FieldName As String, Optional GlobalSettingName As String) As String
    
    Dim strPath As String
    If FieldName <> "" Then
        strPath = ELookup("tblUsers", "UserID = " & g_UserID, FieldName)
    End If
    
    ''If strPath is "", set the filepath to Default Download Path
    If strPath = "" Then
    
        If Not isFalse(GlobalSettingName) Then
            strPath = ELookup("tblGlobalSettings", "GlobalSetting = '" & GlobalSettingName & "'", "GlobalSettingValue")
            If strPath = "" Then
                ShowError "A fallback path was not defined under the global setting"
                directoryPath = ""
                Exit Function
            End If
        Else
            ShowError "A fallback path was not defined under the global setting"
            directoryPath = ""
            Exit Function
        End If

    End If
    
    ''Validate filePath if existing
    If Not DirectoryExists(strPath) Then
        MsgBox "The directory path: """ & strPath & """ is not a valid directory...", vbCritical + vbOKOnly
        directoryPath = ""
        Exit Function
    End If
    
    directoryPath = strPath
    
End Function

Function CreateDirectoryStructure(ByVal destinationPath As String) As Boolean
    Dim fso As Object
    Dim folderPath As String
    Dim folderArray As Variant
    Dim i As Integer
    
    ' Create a FileSystemObject
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo 0
    
    ' Check if the FileSystemObject was created successfully
    If fso Is Nothing Then
        MsgBox "Error creating FileSystemObject. Check if the 'Microsoft Scripting Runtime' reference is added.", vbExclamation
        CreateDirectoryStructure = False
        Exit Function
    End If
    
    ' Split the path into an array of folders
    folderArray = Split(destinationPath, "\")
    
    ' Initialize the folderPath variable
    folderPath = ""
    
    ' Loop through the array and create each folder
    For i = LBound(folderArray) To UBound(folderArray)
        ' Append the current folder to the path
        folderPath = folderPath & folderArray(i) & "\"
        
        ' Check if the folder exists, and create it if it doesn't
        If Not fso.FolderExists(folderPath) Then
            On Error Resume Next
            fso.CreateFolder folderPath
            On Error GoTo 0
            
            ' Check if an error occurred during folder creation
            If Err.number <> 0 Then
                MsgBox "Error creating folder: " & folderPath, vbExclamation
                CreateDirectoryStructure = False
                Exit Function
            End If
        End If
    Next i
    
    ' Return True if the directory structure was created successfully
    CreateDirectoryStructure = True
End Function



