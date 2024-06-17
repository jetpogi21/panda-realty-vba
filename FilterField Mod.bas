Attribute VB_Name = "FilterField Mod"
Option Compare Database
Option Explicit

Public Function FilterFieldOnLoad(frm As Form)

    SetDefaultUserID frm
    
    Dim sqlStr
    sqlStr = "SELECT ModelFieldID, ModelField FROM tblModelFields"
    ''Check if there is a parent
    If DoesObjectExists(frm.Parent) Then
        If frm.Parent.Name = "frmModels" Then
            sqlStr = sqlStr & " WHERE ModelID = " & frm.Parent.ModelID
        End If
    End If
    
    sqlStr = sqlStr & " ORDER BY ModelField"
    
    frm.ModelFieldID.rowSource = sqlStr
    frm.ModelFieldID.Requery
    
End Function
