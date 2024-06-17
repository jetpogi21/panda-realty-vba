Attribute VB_Name = "PropertyStatusFilter Mod"
Option Compare Database
Option Explicit

Public Function PropertyStatusFilterCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function ShowAllFavoriteFields(frm As Form, Optional reversed As Boolean = True)
    
    RunSQL "UPDATE tblPropertyStatus SET IsShownOnFavorite = " & reversed
    IsShownOnFavoriteUpdate frm, False
    frm("subform").Form.Requery
    
End Function

Public Function IsShownOnFavoriteUpdate(frm As Form, Optional filterOutStatusID As Boolean = True)
    
    If filterOutStatusID Then frm.Requery
    Dim frm2 As Form
    Dim PropertyStatusArr As New clsArray
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblPropertyStatus WHERE IsShownOnFavorite = -1")
    
    If Not rs.EOF Then rs.MoveFirst
    
    If filterOutStatusID Then
        Dim PropertyStatusID, IsShownOnFavorite
        PropertyStatusID = frm("PropertyStatusID")
        IsShownOnFavorite = frm("IsShownOnFavorite")
    End If
    
    Do Until rs.EOF
        If filterOutStatusID Then
            If IsShownOnFavorite Then
                If PropertyStatusID = rs.fields("PropertyStatusID") Then
                    PropertyStatusArr.Add rs.fields("PropertyStatusID")
                Else
                    PropertyStatusArr.Add rs.fields("PropertyStatusID")
                End If
            Else
                PropertyStatusArr.Add rs.fields("PropertyStatusID")
            End If
        Else
            PropertyStatusArr.Add rs.fields("PropertyStatusID")
        End If
        
        rs.MoveNext
    Loop
    
    Dim sqlStr
    sqlStr = "SELECT * FROM qryFavoriteProperties"
    If PropertyStatusArr.Count > 0 Then
        sqlStr = "SELECT * FROM qryFavoriteProperties WHERE PropertyStatusID in(" & PropertyStatusArr.JoinArr(",") & ") OR PropertyStatusID IS NULL"
    End If
    
    Debug.Print sqlStr
     
    If IsFormOpen("mainFavoriteProperties") Then
        Set frm2 = Forms("mainFavoriteProperties")
        frm2("subform").Form.RecordSource = sqlStr
        frm2("subform").Form.Requery
    End If
    
End Function
