Attribute VB_Name = "BuyerMember Mod"
Option Compare Database
Option Explicit

Public Function BuyerMemberCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
            SetEntityMemberForm frm
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function


Public Function SetEntityMemberForm(frm As Form)

    frm.BeforeUpdate = "=SaveFormData2([Form],""EntityMember"")"
    frm("MemberName").AfterUpdate = "=EntityMemberEntityIDAfterUpdate([Form])"
    

End Function

Public Function EntityMemberEntityIDAfterUpdate(frm As Form)
    
    Dim EntityID, EntityMemberID
    EntityID = frm("EntityID")
    
    If IsNull(EntityID) Then Exit Function
    
    EntityMemberID = frm("EntityMemberID")
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblEntities WHERE EntityID = " & EntityID)
    
    If Not rs.EOF Then
        frm("MemberAddress") = rs.fields("Address")
        
        If frm.NewRecord Then
            If ECount("tblEntityMembers", "EntityID = " & EntityID) = 0 Then
                frm("MemberPhoneNumber") = rs.fields("PhoneNumber")
                frm("MemberEmailAddress") = rs.fields("EmailAddress")
            End If
        Else
            If ECount("tblEntityMembers", "EntityID = " & EntityID & " AND EntityMemberID <> " & EntityMemberID) = 0 Then
                frm("MemberPhoneNumber") = rs.fields("PhoneNumber")
                frm("MemberEmailAddress") = rs.fields("EmailAddress")
            End If
        End If
        
    End If
    
End Function
