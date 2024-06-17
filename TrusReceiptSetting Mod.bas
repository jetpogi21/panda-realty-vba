Attribute VB_Name = "TrusReceiptSetting Mod"
Option Compare Database
Option Explicit

Public Function TrusReceiptSettingCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function


Public Function GetTrustReceiptHeader()
    
    GetTrustReceiptHeader = ELookup("tblTrusReceiptSettings", "TrusReceiptSettingID > 0", "Header")
    
End Function

Public Function SetDigitalSignatureImg(frm As Form)

    Dim assetDir, fs As Object
    assetDir = GetApplicationSetting("Asset Directory")
    
    If assetDir = "" Then assetDir = CurrentProject.Path
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FolderExists(assetDir) Then assetDir = CurrentProject.Path
    Dim ModelField: ModelField = "DigitalSignature"
    Dim fileName: fileName = frm(ModelField)
    frm(concat(ModelField, "Img")).Picture = concat(assetDir, "\", fileName)

End Function
