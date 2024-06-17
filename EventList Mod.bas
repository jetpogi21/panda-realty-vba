Attribute VB_Name = "EventList Mod"
Option Compare Database
Option Explicit

Public Function EventListCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function FollowHyperlinkOrOpenForm(frm As Form, Optional useDescription As Boolean = False)
    
    Dim EventFile: EventFile = frm("EventFile")
    
    Dim Identifier
    If useDescription Then
        Identifier = frm("Description")
    Else
        Identifier = frm("EventListID")
    End If
    
    Dim filterStr
    If useDescription Then
        filterStr = "EventName = " & Esc(Identifier)
    Else
        filterStr = "EventListID = " & Identifier
    End If
    
    If Not isPresent("tblEventList", filterStr) Then
        MsgBox "Event Name not found."
        Exit Function
    End If
    
    If isFalse(EventFile) Then
        DoCmd.OpenForm "frmEventList", , , filterStr
    Else
        FollowFormHyperlink frm, "EventFile"
    End If

End Function

Public Function RebuildAllPropertyEvents()

    Dim resp: resp = MsgBox("WARNING: This will replace all events attached to each property with new ones. Do you want to proceed?", vbCritical + vbYesNo)
    
    If resp = vbNo Then Exit Function
    
    RunSQL "DELETE FROM tblEventTimelines"
    
End Function
