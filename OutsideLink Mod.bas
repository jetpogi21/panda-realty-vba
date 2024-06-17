Attribute VB_Name = "OutsideLink Mod"
Option Compare Database
Option Explicit

Public Function OutsideLinkCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function OpenWebsiteLink(frm As Form)

    Dim Link
    Link = frm("Link")
    
    If ExitIfTrue(isFalse(Link), "There's no valid link...") Then Exit Function
    
    CreateObject("Shell.Application").Open Link
    
End Function

Public Function OpenGoogle()

    CreateObject("Shell.Application").Open "https://google.com"
    
End Function

Public Function OpenGoogleTask()

    CreateObject("Shell.Application").Open "https://calendar.google.com/calendar/u/0?cid=Y184MTk3ZDI2ZTAxNDViYmQ5Y2I1ODJhYWRmY2I1MTdmMmY4NjU5N2Y5NGY0MGQ3YzAyZDI1Yjg0MmU2YjFlODY0QGdyb3VwLmNhbGVuZGFyLmdvb2dsZS5jb20"
    
End Function


