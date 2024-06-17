Attribute VB_Name = "Database Procedures"
Option Compare Database
Option Explicit

Public Function InsertData(tblName As String, fields() As String, fieldValues() As String) As Variant
    
    RunSQL "INSERT INTO " & tblName & " (" & Join(fields, ",") & ") VALUES (" & Join(fieldValues, ",") & ")"
    InsertData = CurrentDb.OpenRecordset("SELECT @@IDENTITY")(0)
    
End Function

Public Function UpdateData(tblName As String, setStatements() As String, filterStr As String) As Variant
    
    RunSQL "UPDATE " & tblName & " SET " & Join(setStatements, ",") & " WHERE " & filterStr
    
End Function

Public Function RunSQL(sqlStr) As Long

    DoCmd.SetWarnings True
    Dim dbs As DAO.Database
    Set dbs = CurrentDb
    dbs.Execute sqlStr
    DoCmd.SetWarnings True
    
    RunSQL = dbs.RecordsAffected
    
End Function

