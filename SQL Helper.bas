Attribute VB_Name = "SQL Helper"
Option Compare Database
Option Explicit

Public Function EscapeString(Value, Optional tblName = "", Optional fieldName As Variant = "") As String

    If IsNull(Value) Then
        EscapeString = "Null"
        Exit Function
    End If
    
    If tblName <> "" Then
        Dim defType As Object, fieldType
        If DoesPropertyExists(CurrentDb.TableDefs, tblName) Then
            Set defType = CurrentDb.TableDefs
        Else
            Set defType = CurrentDb.QueryDefs
        End If
        
        fieldType = defType(tblName).fields(fieldName).Type
        
        Select Case fieldType
            Case 10, 12:
                EscapeString = Chr(34) & replace(Value, Chr(34), Chr(34) & Chr(34)) & Chr(34)
            Case 8:
                EscapeString = "#" & SQLDate(Value) & "#"
            Case Else:
                EscapeString = Value
        End Select
        
    Else
        EscapeString = Chr(34) & Value & Chr(34)
    End If
    
End Function
