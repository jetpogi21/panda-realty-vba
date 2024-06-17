Attribute VB_Name = "EventTimeline Mod"
Option Compare Database
Option Explicit

Public Function EventTimelineCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Private Function GetDepositType(Description) As String
    
    If IsNull(Description) Then Exit Function
    
    If Description = "INITIAL DEPOSIT RECEIPT" Or Description Like "Initial Deposit Trust Receipt*" Then
        GetDepositType = "INITIAL DEPOSIT RECEIPT"
        Exit Function
    End If
    
    If Description = "BALANCE DEPOSIT RECEIPT" Or Description Like "Balance Deposit Trust Receipt*" Then
        GetDepositType = "BALANCE DEPOSIT RECEIPT"
        Exit Function
    End If
    
End Function

Public Function dshtEventTimelines_IsChecked_AfterUpdate(frm As Form)

    Dim IsChecked: IsChecked = frm("IsChecked")
    Dim Description: Description = frm("Description")
    
    If Not IsChecked Then Exit Function
    
    Dim DepositType: DepositType = GetDepositType(Description)
    
    If Not (DepositType = "INITIAL DEPOSIT RECEIPT" Or DepositType = "BALANCE DEPOSIT RECEIPT") Then Exit Function
         
    ''Balance Deposit Trust Receipt, Initial Deposit Trust Receipt
    Dim resp: resp = MsgBox(Esc(Description) & " has been checked. Do you want to send a solicitor letter?", vbYesNo)
    If resp = vbNo Then Exit Function
    'INITIAL DEPOSIT RECEIPT,BALANCE DEPOSIT RECEIPT
    
    DoCmd.RunCommand acCmdSaveRecord
    Set frm = Forms("frmPropertyList")
    Dim LetterType As String: LetterType = IIf(DepositType = "INITIAL DEPOSIT RECEIPT", "Initial Deposit", "Balance Deposit")
    
    frm("txtLetterType") = LetterType
    
    SendLetterToSolicitor frm
    
End Function

Public Function GetEventTimelineDate(Days, ContractDate)
    
    If isFalse(ContractDate) Then
        GetEventTimelineDate = Null
        Exit Function
    End If
    
    GetEventTimelineDate = DateAdd("d", Days, ContractDate)
End Function

Public Function RebuildEventTimeline(frm As Form)
    
    Dim PropertyListID: PropertyListID = frm("PropertyListID")
    
    If isFalse(PropertyListID) Then Exit Function
    
    Dim resp: resp = MsgBox("This will delete the existing events and will be replaced by a new one. Would you like to proceed?")
    
    If resp = vbNo Then Exit Function
    
    RunSQL "DELETE FROM tblEventTimelines WHERE PropertyListID = " & PropertyListID
    InsertTo_tblEventTimelines frm
    
End Function

Public Function DeleteEventTimeline(frm As Form)
    
    Set frm = frm("dshtEventTimelines").Form
    
    Dim EventTimelineID: EventTimelineID = frm("EventTimelineID")
    If isFalse(EventTimelineID) Then Exit Function
    
    Dim resp: resp = MsgBox("This will delete the select event. Would you like to proceed?")
    
    If resp = vbNo Then Exit Function
    
    RunSQL "DELETE FROM tblEventTimelines WHERE EventTimelineID = " & EventTimelineID
    
    frm.Requery
    
End Function
