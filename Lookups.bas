Attribute VB_Name = "Lookups"
Option Compare Database
Option Explicit

Public Function isPresent(tblName, filterStr) As Boolean

    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT *  FROM " & tblName & " WHERE " & filterStr)
    
    If rs.EOF Then
        isPresent = False
    Else
        isPresent = True
    End If
    
    Exit Function

End Function

Public Function ELookup(tblName As String, filterStr As String, fldName, Optional orderStr As String) As String
    
    Dim rs As Recordset
    Dim sqlStr As String
    sqlStr = "SELECT * FROM " & tblName & " WHERE " & filterStr
    
    If orderStr <> "" Then
        sqlStr = sqlStr & " ORDER BY " & orderStr
    End If
    
    Set rs = CurrentDb.OpenRecordset(sqlStr)
    
    If rs.EOF Then
        ELookup = ""
    Else
'On Error GoTo ErrHandler:
        If isFalse(rs.fields(fldName)) Then
            ELookup = ""
            Exit Function
        End If
        ELookup = rs.fields(fldName)
    End If
    
    Exit Function

'ErrHandler:
'    LogError Err.Number, Err.Description, "ELookup", , True
'    ELookup = ""

End Function

Public Function ELookupDate(tblName As String, filterStr As String, fldName As String, Optional orderStr As String) As Date

    Dim rs As Recordset
    Dim sqlStr As String
    sqlStr = "SELECT * FROM " & tblName & " WHERE " & filterStr
    
    If orderStr <> "" Then
        sqlStr = sqlStr & " ORDER BY " & orderStr
    End If
    
    Set rs = CurrentDb.OpenRecordset(sqlStr)
    
    If rs.EOF Then
        ELookupDate = #1/1/2100#
    Else
        If isFalse(rs.fields(fldName)) Then
            ELookupDate = #1/1/2100#
            Exit Function
        End If
        ELookupDate = SQLDate(rs.fields(fldName))
    End If

End Function

Public Function ReturnRecordset(sqlStr) As Recordset
    
    Set ReturnRecordset = CurrentDb.OpenRecordset(sqlStr)
    
End Function

Public Function ESum(sqlStr As String, FieldName As String) As Double
    
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset(sqlStr)
    
    If rs.EOF Then
        ESum = 0
        Exit Function
    End If
    
    If isFalse(rs.fields(FieldName)) Then
        ESum = 0
        Exit Function
    Else
        ESum = rs.fields(FieldName)
    End If
    
End Function

Public Function ESum2(tblName As String, filterStr As String, FieldName As String) As Double
    
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT SUM(" & FieldName & ") As SumOfRecord FROM " & tblName & " WHERE " & filterStr)
    
    If rs.EOF Then
        ESum2 = 0
        Exit Function
    End If
    
    If isFalse(rs.fields("SumOfRecord")) Then
        ESum2 = 0
        Exit Function
    Else
        ESum2 = rs.fields("SumOfRecord")
    End If
    
End Function

Public Function ECount(tblName, filterStr As String) As Double
    
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT COUNT(*) As CountOfRecord FROM " & tblName & " WHERE " & filterStr)
    
    If rs.EOF Then
        ECount = 0
        Exit Function
    End If
    
    If isFalse(rs.fields("CountOfRecord")) Then
        ECount = 0
        Exit Function
    Else
        ECount = rs.fields("CountOfRecord")
    End If
    
End Function

