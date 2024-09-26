Attribute VB_Name = "SyncBuyerRequirements Mod"
Option Compare Database
Option Explicit

Public Function SyncBuyerRequirementsCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function SyncAllBuyerRequirements()
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT PropertyListID FROM qryPropertyBuyers GROUP BY PropertyListID")
    Do Until rs.EOF
        Dim PropertyListID: PropertyListID = rs.fields("PropertyListID")
        SyncBuyerRequirements Forms!frmCustomDashboard, PropertyListID
        rs.MoveNext
    Loop
    
    MsgBox "Done"
    
End Function


Public Function SyncBuyerRequirements(frm As Form, Optional PropertyListID = "")

On Error GoTo ErrHandler:
    If isFalse(PropertyListID) Then
        PropertyListID = frm("PropertyListID")
    End If
    
    If isFalse(PropertyListID) Then Exit Function

    ''Main goal insert to tblEntityExtraFeatures using join from a table
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    Set rs = ReturnRecordset("SELECT * FROM tblPropertyList WHERE PropertyListID = " & PropertyListID)
    If rs.EOF Then Exit Function
    
    ''Exit if not property buyer since there's no need to sync
    If Not isPresent("qryPropertyBuyers", "PropertyListID = " & PropertyListID) Then Exit Function
    
    RunSQL "DELETE FROM tblSyncBuyerRequirements"
    Dim rs2 As Recordset: Set rs2 = ReturnRecordset("SELECT * FROM tblBuyerRequirements WHERE Not PropertyListFieldName IS NULL")
    Do Until rs2.EOF
        Dim PropertyListFieldName: PropertyListFieldName = rs2.fields("PropertyListFieldName")
        Dim BuyerRequirementID: BuyerRequirementID = rs2.fields("BuyerRequirementID")
        Dim Value: Value = rs.fields(PropertyListFieldName)
        If Not isFalse(Value) Then
            Dim Operator: Operator = "null"
            If PropertyListFieldName = "LandSize" Or PropertyListFieldName = "AppraisedAmount" Then
                Operator = Esc("<=")
            End If
            RunSQL "INSERT INTO tblSyncBuyerRequirements (BuyerRequirementID,[Value],Operator) VALUES (" & BuyerRequirementID & _
                "," & Esc(Value) & "," & Operator & ")"
        End If
        ''LandSize,AppraisedAmount
        rs2.MoveNext
    Loop
    
    ''qryPropertyBuyers --> PropertyListID, EntityID
    sqlStr = "SELECT * FROM (SELECT EntityID FROM qryPropertyBuyers WHERE PropertyListID = " & PropertyListID & ") a, " & _
        "(SELECT BuyerRequirementID,Value,Operator FROM tblSyncBuyerRequirements) b"
        
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "tblEntityExtraFeatures"
          .fields = "EntityID,BuyerRequirementID,[Value],Operator"
          .InsertSQL = sqlStr
          .InsertFilterField = "EntityID,BuyerRequirementID,[Value],Operator"
          ''.InsertValues
          ''.InsertUseAsPlain
          ''.LastInsertID
          ''.SQL
          ''makeQuery .SQL
          rowsAffected = .Run
    End With
    
    Exit Function
ErrHandler:
    If Err.number = 2465 Then
        Exit Function
    End If

End Function
