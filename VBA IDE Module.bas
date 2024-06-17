Attribute VB_Name = "VBA IDE Module"
Option Compare Database
Option Explicit

Public Sub CreateOnListEvent()
    
    DoCmd.OpenForm "dshtBookDetails", acDesign
    Forms!dshtBookDetails.HasModule = True

    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim LineNum As Long
    
    Set VBProj = Application.VBE.VBProjects("Database")
    If DoesPropertyExists(VBProj.VBComponents, "Form_dshtBookDetails") Then
        Set VBComp = VBProj.VBComponents("Form_dshtBookDetails")
    Else
        Set VBComp = VBProj.VBComponents.Add(vbext_ct_MSForm)
        VBComp.Name = "Form_dshtBookDetails"
    End If
    
    Set CodeMod = VBComp.CodeModule
    
    With CodeMod
        LineNum = .CreateEventProc("NotInList", "StudentID")
        LineNum = LineNum + 1
        .InsertLines LineNum, "    MsgBox " & EscapeString("Hello World")
    End With
    
End Sub

