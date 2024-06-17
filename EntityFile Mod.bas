Attribute VB_Name = "EntityFile Mod"
Option Compare Database
Option Explicit

Public Function EntityFileCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function SyncPropertyFiles(frm As Form)

    Dim PropertyListID
    PropertyListID = frm("PropertyListID")
    
    If isFalse(PropertyListID) Then Exit Function
    
    Dim StreetAddress
    StreetAddress = frm("StreetAddress")
    
    Dim uploadDirectory
    uploadDirectory = GetAttachmentsDirectory(StreetAddress)
    
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblEntityFiles WHERE PropertyListID = " & PropertyListID)
    
    Do Until rs.EOF
        
        Dim filePath, EntityFileLink, EntityFileID
        EntityFileID = rs.fields("EntityFileID")
        EntityFileLink = rs.fields("EntityFileLink")
        filePath = concat(uploadDirectory, EntityFileLink)
        
        If Not DoesFileExist(filePath) Then
            Debug.Print filePath
            DeleteEntityFile rs
        End If
        
        rs.MoveNext
        
    Loop
    
    RequerySubforms frm

    
End Function

Private Function RequerySubforms(frm As Form)

    Dim subformArr As New clsArray, subformItem
    
    subformArr.arr = "Buyer,Contact,Tenant,Seller,PropertyEnity"
    
    For Each subformItem In subformArr.arr
        Dim subformName
        subformName = "sub" & subformItem & "Files"
        
        frm(subformName).Form.Requery
        
    Next subformItem
    
End Function

Private Function DeleteEntityFile(rs As Recordset)
    
    rs.Delete
    'RunSQL "DELETE FROM tblEntityFiles WHERE EntityFileID = " & EntityFileID
    
End Function

Public Function DoesFileExist(filePath) As Boolean
    
    
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    DoesFileExist = fs.fileExists(filePath)

End Function
