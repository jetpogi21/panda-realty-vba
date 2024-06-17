Attribute VB_Name = "EventInsert Mod"
Option Compare Database
Option Explicit

Public Function EventInsertCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function AddPropertyEvent(frm As Form, PropertyListID)

    Dim EventName: EventName = frm("EventName")
    Dim position: position = frm("Position")
    Dim EventOrder: EventOrder = frm("EventTimelineID").Column(2)
    
    
    If ExitIfTrue(isFalse(EventName), "Event Name can't be blank.") Then Exit Function
    If ExitIfTrue(isFalse(position), "Position can't be blank.") Then Exit Function
    If ExitIfTrue(isFalse(EventOrder), "Event Timeline can't be blank.") Then Exit Function
    
    If isPresent("tblEventTimelines", "Description = " & Esc(EventName) & " AND PropertyListID = " & PropertyListID) Then
        MsgBox Esc(EventName) & " is already present."
        Exit Function
    End If
    
    EventOrder = CDbl(EventOrder)
    If position = "Before" Then
        EventOrder = EventOrder - 0.001
    Else
        EventOrder = EventOrder + 0.001
    End If
    
    RunSQL "INSERT INTO tblEventTimelines (PropertyListID,EventOrder,Description) values (" & _
        PropertyListID & "," & EventOrder & "," & Esc(EventName) & ")"
    
    If IsFormOpen("frmPropertyList") Then
        Forms("frmPropertyList")("dshtEventTimelines").Form.Requery
    End If
    
    DoCmd.Close acForm, frm.Name, acSaveNo
    
End Function

Public Function AddEventList(frm As Form)

    Dim EventName: EventName = frm("EventName")
    Dim position: position = frm("Position")
    Dim EventOrder: EventOrder = frm("EventTimelineID").Column(2)
    
    If ExitIfTrue(isFalse(EventName), "Event Name can't be blank.") Then Exit Function
    If ExitIfTrue(isFalse(position), "Position can't be blank.") Then Exit Function
    If ExitIfTrue(isFalse(EventOrder), "Event Timeline can't be blank.") Then Exit Function
    
    If isPresent("tblEventList", "EventName = " & Esc(EventName)) Then
        MsgBox Esc(EventName) & " is already present."
        Exit Function
    End If
    
    EventOrder = CDbl(EventOrder)
    If position = "Before" Then
        EventOrder = EventOrder - 0.001
    Else
        EventOrder = EventOrder + 0.001
    End If
    
    RunSQL "INSERT INTO tblEventList (EventOrder,EventName) values (" & _
        EventOrder & "," & Esc(EventName) & ")"
    
    If IsFormOpen("mainEventList") Then
        Forms("mainEventList")("subform").Form.Requery
    End If
    
    DoCmd.Close acForm, frm.Name, acSaveNo
    
End Function

Public Function Open_frmEventInserts_FromMainForm(frm As Form)
    
    DoCmd.OpenForm "frmEventInserts", , , , acFormAdd
    
    Set frm = Forms("frmEventInserts")
    frm("lblEventTimelineID").Caption = "Target Event"
    
    ''EventTimelineID
    Dim sqlStr: sqlStr = "Select EventListID,EventName,EventOrder FROM tblEventList ORDER BY EventOrder"
    frm("EventTimelineID").rowSource = sqlStr
    
    ''Get the default EventTimeLineID
    Dim EventListID: EventListID = ELookup("tblEventList", "EventListID > 0", "EventListID", "EventOrder")
    frm("EventTimelineID") = EventListID
    frm("cmdSaveClose").OnClick = "=AddEventList([Form])"
    
End Function


Public Function Open_frmEventInserts(frm As Form)
    
    Dim PropertyListID: PropertyListID = frm("PropertyListID")
    
    If isFalse(PropertyListID) Then Exit Function
    
    DoCmd.OpenForm "frmEventInserts", , , , acFormAdd
    
    Set frm = Forms("frmEventInserts")
    
    ''EventTimelineID
    Dim sqlStr: sqlStr = "Select EventTimelineID,Description,EventOrder FROM tblEventTimelines WHERE PropertyListID = " & PropertyListID & " ORDER BY EventOrder"
    frm("EventTimelineID").rowSource = sqlStr
    
    ''Get the default EventTimeLineID
    Dim EventTimelineID: EventTimelineID = ELookup("tblEventTimelines", "PropertyListID = " & PropertyListID, "EventTimelineID", "EventOrder")
    frm("EventTimelineID") = EventTimelineID
    frm("cmdSaveClose").OnClick = "=AddPropertyEvent([Form]," & PropertyListID & ")"
    
End Function
