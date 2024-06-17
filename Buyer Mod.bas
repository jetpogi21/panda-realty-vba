Attribute VB_Name = "Buyer Mod"
Option Compare Database
Option Explicit

Public Function BuyerCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            SetEntityCategoryDefault frm, "Buyer", False
            frm.OnLoad = "=EntityDEFormLoad([Form],""Buyer"")"
        Case 5: ''Datasheet Form
            SetEntityCategoryDefault frm, "Buyer"
        Case 6: ''Main Form
            SetMainFormProperty frm, "Buyer"
        Case 7: ''Tabular Report
    End Select

End Function

Public Function InsertToSubCategories()

    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "tblEntities"
          .AddFilter "EntityCategoryID = 1 AND tblEntities.BuyerStatusID <> 0 AND NOT tblEntities.BuyerStatusID IS NULL"
          .fields = "tblEntities.*,BuyerStatus"
          .Joins.Add GenerateJoinObj("tblBuyerStatus", "BuyerStatusID")
          .OrderBy = "EntityID"
          Set rs = .Recordset
    End With
    
    Dim BuyerStatusArr As New clsArray, item
    Do Until rs.EOF
        Dim BuyerStatus: BuyerStatus = rs.fields("BuyerStatus")
        Dim EntityID: EntityID = rs.fields("EntityID")
        BuyerStatusArr.arr = replace(BuyerStatus, "/", ",")
        For Each item In BuyerStatusArr.arr
            Dim Trimmed: Trimmed = Trim(item)
            
            Dim BuyerStatusID:
            ''Lookup from the tblBuyerStatus if not existing insert if existing then get the id
            If isPresent("tblBuyerStatus", "BuyerStatus = " & EscapeString(Trimmed)) Then
                BuyerStatusID = ELookup("tblBuyerStatus", "BuyerStatus = " & EscapeString(Trimmed), "BuyerStatusID")
            Else
                RunSQL "INSERT INTO tblBuyerStatus (BuyerStatus) VALUES (" & EscapeString(Trimmed) & ")"
                BuyerStatusID = ELookup("tblBuyerStatus", "BuyerStatusID > 0", "BuyerStatusID", "BuyerStatusID DESC")
            End If
            
            RunSQL "INSERT INTO tblBuyerSubcategories (EntityID,BuyerStatusID) VALUES (" & EntityID & "," & BuyerStatusID & ")"
            
        Next item
        rs.MoveNext
    Loop

End Function

Public Function RemovePureSellers()
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    ''tblEntities, tblEntityCategories -> 1 tblBuyerStatus = "SELLER"
    
    Set sqlObj = New clsSQL
    With sqlObj
      .SQLType = "DELETE"
      .fields = "tblEntities.*"
      .Source = "tblEntities"
      .AddFilter "EntityCategoryID = 1 AND BuyerStatus = " & EscapeString("SELLER")
      .Joins.Add GenerateJoinObj("tblBuyerStatus", "BuyerStatusID")
      rowsAffected = .Run
    End With
    
End Function

Public Function ShowHideBuyerStatus(frm As Form)

    Dim ColumnHidden, ctl As Control
    
    Set ctl = frm("subform").Form("BuyerStatusID")
    
    ColumnHidden = ctl.ColumnHidden
    ctl.ColumnHidden = Not ColumnHidden
    
End Function

''Bed, Bath, Car, LandSize, SalePrice
Public Function BuyerFormFeatureEquality(frm As Form, Feature, Operator)

    Dim txtBoxVal, txtBoxName, BuyerRequirement, PropertyListFieldName, BuyerRequirementID, EntityID
    
    EntityID = frm("EntityID")
    
    If ExitIfTrue(isFalse(EntityID), "Buyer is empty...") Then Exit Function
    
    Dim gtOrlt
    gtOrlt = "Gt"
    If Operator = "<=" Then gtOrlt = "Lt"

    txtBoxName = "fltr" & Feature & gtOrlt
    txtBoxVal = frm(txtBoxName)
    
    PropertyListFieldName = Feature 'To be used as filter field in future (PropertyListFieldName)
    BuyerRequirement = Feature ''To be used in filtering the buyer requirement and to be used as caption
    
    ''See if it's already existing from the qryEntityExtraFeatures
    Dim EntityExtraFeatureID
    EntityExtraFeatureID = ELookup("qryEntityExtraFeatures", "BuyerRequirement = " & EscapeString(BuyerRequirement) & _
        " AND Operator = " & EscapeString(Operator), "EntityExtraFeatureID")

    If Not IsNull(txtBoxVal) Then
        
        If EntityExtraFeatureID = "" Then
        
            ''If not then add the buyerrequirement first and get the ID
            ''then do an update
            
            RunSQL "INSERT INTO tblBuyerRequirements (BuyerRequirement, PropertyListFieldName) VALUES (" & EscapeString(BuyerRequirement) & _
                "," & EscapeString(PropertyListFieldName) & ")"
            ''Get the BuyerRequirement ID
            BuyerRequirementID = ELookup("tblBuyerRequirements", "BuyerRequirement = " & EscapeString(BuyerRequirement), "BuyerRequirementID")
            
            RunSQL "INSERT INTO tblEntityExtraFeatures (EntityID, BuyerRequirementID, [Value], Operator) VALUES (" & EntityID & "," & _
                BuyerRequirementID & "," & txtBoxVal & "," & EscapeString(Operator) & ")"
        
        Else
            
            RunSQL "UPDATE tblEntityExtraFeatures SET [Value] = " & txtBoxVal & " WHERE EntityExtraFeatureID = " & EntityExtraFeatureID
                
        End If
        
    Else
        
        On Error Resume Next
        RunSQL "DELETE FROM tblEntityExtraFeatures WHERE EntityExtraFeatureID = " & EntityExtraFeatureID
    
    End If
    
    ''Refresh the subform
    frm("subBuyerRequirements").Form.Requery
    frm("subBuyerRequirements").Form("BuyerRequirementID").Requery
    
End Function

Public Function BuyerFormFeatureQtyUpdate(frm As Form, Feature, Qty, Optional Operator = Null)

    Dim chkBoxVal, chkBoxName, BuyerRequirement, PropertyListFieldName, BuyerRequirementID, EntityID, PropertyTypeVal
    
    EntityID = frm("EntityID")
    
    If ExitIfTrue(isFalse(EntityID), "Buyer is empty...") Then Exit Function
    
    chkBoxName = "fltr" & Feature & Qty
    chkBoxVal = frm(chkBoxName)
    
    PropertyListFieldName = Feature 'To be used as filter field in future (PropertyListFieldName)
    BuyerRequirement = Feature ''To be used in filtering the buyer requirement and to be used as caption
    PropertyTypeVal = frm(chkBoxName).Tag
    
    ''See if it's already existing from the qryEntityExtraFeatures
    Dim EntityExtraFeatureID
    EntityExtraFeatureID = ELookup("qryEntityExtraFeatures", "BuyerRequirement = " & EscapeString(BuyerRequirement) & _
        " AND Value = " & EscapeString(PropertyTypeVal), "EntityExtraFeatureID")

    If chkBoxVal Then
        
        If EntityExtraFeatureID = "" Then
        
            ''If not then add the buyerrequirement first and get the ID
            ''then do an update
            
            RunSQL "INSERT INTO tblBuyerRequirements (BuyerRequirement, PropertyListFieldName) VALUES (" & EscapeString(BuyerRequirement) & _
                "," & EscapeString(PropertyListFieldName) & ")"
            ''Get the BuyerRequirement ID
            BuyerRequirementID = ELookup("tblBuyerRequirements", "BuyerRequirement = " & EscapeString(BuyerRequirement), "BuyerRequirementID")
            
            RunSQL "INSERT INTO tblEntityExtraFeatures (EntityID, BuyerRequirementID, [Value], Operator) VALUES (" & EntityID & "," & _
                BuyerRequirementID & "," & EscapeString(PropertyTypeVal) & "," & EscapeString(Operator) & ")"
                
        End If
        
    Else
        
        On Error Resume Next
        RunSQL "DELETE FROM tblEntityExtraFeatures WHERE EntityExtraFeatureID = " & EntityExtraFeatureID
    
    End If
    
    ''Refresh the subform
    frm("subBuyerRequirements").Form.Requery
    frm("subBuyerRequirements").Form("BuyerRequirementID").Requery
    
End Function



Public Function BuyerFormPropertyTypeUpdate(frm As Form, PropertyType)

    Dim chkBoxVal, chkBoxName, BuyerRequirement, PropertyListFieldName, BuyerRequirementID, EntityID, PropertyTypeVal
    
    EntityID = frm("EntityID")
    
    If ExitIfTrue(isFalse(EntityID), "Buyer is empty...") Then Exit Function
    
    chkBoxName = "fltrPropertyType" & PropertyType
    chkBoxVal = frm(chkBoxName)
    PropertyListFieldName = "PropertyType" 'To be used as filter field in future (PropertyListFieldName)
    BuyerRequirement = "Property Type" ''To be used in filtering the buyer requirement and to be used as caption
    PropertyTypeVal = frm(chkBoxName).Tag
    
    ''See if it's already existing from the qryEntityExtraFeatures
    Dim EntityExtraFeatureID
    EntityExtraFeatureID = ELookup("qryEntityExtraFeatures", "BuyerRequirement = " & EscapeString(BuyerRequirement) & _
        " AND Value = " & EscapeString(PropertyTypeVal), "EntityExtraFeatureID")
    
    If chkBoxVal Then
        
        
        If EntityExtraFeatureID = "" Then
        
            ''If not then add the buyerrequirement first and get the ID
            ''then do an update
            RunSQL "INSERT INTO tblBuyerRequirements (BuyerRequirement, PropertyListFieldName) VALUES (" & EscapeString(BuyerRequirement) & _
                "," & EscapeString(PropertyListFieldName) & ")"
            ''Get the BuyerRequirement ID
            BuyerRequirementID = ELookup("tblBuyerRequirements", "BuyerRequirement = " & EscapeString(BuyerRequirement), "BuyerRequirementID")
            
            RunSQL "INSERT INTO tblEntityExtraFeatures (EntityID, BuyerRequirementID, [Value]) VALUES (" & EntityID & "," & _
                BuyerRequirementID & "," & EscapeString(PropertyTypeVal) & ")"
                
        End If
        
        
    Else
        
        RunSQL "DELETE FROM tblEntityExtraFeatures WHERE EntityExtraFeatureID = " & EntityExtraFeatureID
    
    End If
    
    
    ''Refresh the subform
    frm("subBuyerRequirements").Form.Requery
    frm("subBuyerRequirements").Form("BuyerRequirementID").Requery
    
End Function


Public Function OpenBuyerForm(frm As Form)

    Dim EntityID
    EntityID = frm("EntityID")
    
    If ExitIfTrue(isFalse(EntityID), "Buyer is empty..") Then Exit Function
    
    DoCmd.OpenForm "frmBuyers", , , "EntityID = " & EntityID
    
End Function

Public Function RefreshBuyerForm()

    Dim frm As Form
    Set frm = Forms("frmBuyers")
    
    frm("subBuyerFiles").Form.Requery
    
End Function

Public Function BuyerEnityNameAfterUpdate(frm As Form)

    Dim EntityName
    EntityName = frm("EntityName")
    
    Dim Address, PhoneNumber, EmailAddress, Notes
    Address = frm("EntityName").Column(1)
    PhoneNumber = frm("EntityName").Column(2)
    EmailAddress = frm("EntityName").Column(3)
    Notes = frm("EntityName").Column(4)
    
    frm("Address") = Address
    frm("PhoneNumber") = PhoneNumber
    frm("EmailAddress") = EmailAddress
    If Notes <> "" Then frm("Notes") = Notes
    
End Function

Public Function OpenBuyerSellerForm(frm As Form)
    
    Dim EntityCategoryID, EntityID
    EntityCategoryID = frm("subform").Form("EntityCategoryID")
    EntityID = frm("subform").Form("EntityID")
    
    Dim frmName As String
    frmName = "frmBuyers"
    If EntityCategoryID = 2 Then frmName = "frmSellers"
    OpenFormFromMain frmName, "subform", "EntityID", frm
    
End Function

Public Function BuyerDEFormLoad(frm As Form)

    EntityDEFormLoad frm, "Buyer"
    
    frm("subPropertyBuyers").Form("StreetAddress").ColumnHidden = False
    frm("subPropertyBuyers").Form("OpenInRPP").ColumnHidden = True
    frm("subPropertyBuyers").Form("txtOpenInRPP").ColumnWidth = 1000
    frm("subEntityNotes").Form("Note").ColumnWidth = 4000
    frm("subEntityNotes").Form("StreetAddress").ColumnHidden = False
    
    
    ''Override the Properties Viewed
    frm("subPropertyBuyers").Form.AllowEdits = True
    frm("subPropertyBuyers").Form("vPropertyListID").Locked = True
    frm("subPropertyBuyers").Form("StreetAddress").Locked = True
    
    
    
End Function

Public Function frmBuyers_OnCurrent(frm As Form)
    
    SetFocusOnForm frm, "EntityName"
    SetBuyerFilterForm frm
    
End Function

Private Function ResetFilterForm(frm As Form)
    
    Dim ctl As Control
    For Each ctl In frm.Controls
        If ctl.Name Like "fltr*" Then
            ''If checkbox
            If ctl.ControlType = acCheckBox Then
                ctl.Value = 0
            Else
                ctl.Value = Null
            End If
            ''If not
        End If
    Next ctl
    
End Function

Private Function SetBuyerFilterForm(frm As Form)

    Dim EntityID
    EntityID = frm("EntityID")
    ResetFilterForm frm
    
    If isFalse(EntityID) Then Exit Function
    
    ''Bed, Bath, Car, PropertyType, LandSize, SalePrice
    Dim rs As Recordset
    Set rs = ReturnRecordset("Select * from qryEntityExtraFeatures where BuyerRequirement In('Bed','Bath','Car') And EntityID = " & EntityID)
    
    Dim Value, BuyerRequirement, ControlName, PropertyListFieldName, Suffix
    If Not rs.EOF Then rs.MoveFirst
    Do Until rs.EOF
        PropertyListFieldName = rs.fields("PropertyListFieldName")
        BuyerRequirement = rs.fields("BuyerRequirement")
        Value = rs.fields("Value")
        
        If BuyerRequirement = "Bed" Then
            Suffix = IIf(CInt(Value) > 5, "6", Value)
        Else
            Suffix = IIf(CInt(Value) > 3, "4", Value)
        End If
        
        If Suffix <> 0 Then
            ControlName = "fltr" & PropertyListFieldName & Suffix
            frm(ControlName) = -1
        End If
        rs.MoveNext
    Loop
    
    Set rs = ReturnRecordset("Select * from qryEntityExtraFeatures where BuyerRequirement = 'Property Type' And EntityID = " & EntityID)
    
    If Not rs.EOF Then rs.MoveFirst
    Do Until rs.EOF
        PropertyListFieldName = rs.fields("PropertyListFieldName")
        BuyerRequirement = rs.fields("BuyerRequirement")
        Value = rs.fields("Value")
        Dim ctl As Control
        For Each ctl In frm.Controls
            If ctl.Name Like "*PropertyType*" And ctl.Tag = Value Then
                 ctl.Value = -1
                 Exit For
            End If
        Next ctl
        rs.MoveNext
    Loop
    
    Set rs = ReturnRecordset("Select * from qryEntityExtraFeatures where BuyerRequirement In('Appraised Amount','Land Size') And EntityID = " & EntityID)
    
    If Not rs.EOF Then rs.MoveFirst
    Dim maxPrice
    Dim maxSize
    
    Do Until rs.EOF
        Dim Operator, OperatorCaption
        BuyerRequirement = rs.fields("BuyerRequirement")
        PropertyListFieldName = rs.fields("PropertyListFieldName")
        Value = rs.fields("Value")
    
        If PropertyListFieldName = "AppraisedAmount" Then
            UpdateMaxValue maxPrice, Value
            Value = maxPrice
        Else
            UpdateMaxValue maxSize, Value
        End If
    
        Operator = rs.fields("Operator")
        OperatorCaption = "Gt"
        If Operator = "<=" Then OperatorCaption = "Lt"
    
        ControlName = "fltr" & PropertyListFieldName & OperatorCaption
        frm(ControlName) = IIf(PropertyListFieldName = "AppraisedAmount", maxPrice, maxSize)
        
        rs.MoveNext
    Loop
    
    
End Function

Private Sub UpdateMaxValue(ByRef maxValue As Variant, ByVal newValue As Variant)
    If IsNull(maxValue) Then
        maxValue = newValue
    Else
        If CDbl(newValue) > CDbl(maxValue) Then
            maxValue = newValue
        End If
    End If
End Sub

Public Function BuyerMainFormLoad(frm As Form)

    frm("subform").Form("txtOpenInRPP").ColumnWidth = 1000
    frm("subform").Form("Timestamp").ColumnWidth = 3000
    frm("subform").Form("PropertyListID").ColumnHidden = True
    frm("subform").Form("EntityCategoryID").ColumnHidden = True
    
End Function

Public Function EntityDEFormLoad(frm As Form, EntityCategory)
    
    DefaultFormLoad frm, EntityCategory & "ID"
    On Error Resume Next
    frm("subProperty" & EntityCategory & "s").Form.AllowAdditions = False
    frm("subProperty" & EntityCategory & "s").Form.AllowEdits = False
    frm("subProperty" & EntityCategory & "s").Form.AllowDeletions = False
    
    If EntityCategory = "Buyer" Then
        frm("subBuyerOffers").Form.AllowAdditions = False
        frm("subBuyerOffers").Form.AllowEdits = False
        frm("subBuyerOffers").Form.AllowDeletions = False
    End If
    
    frm("subEntityNotes").Form("Note").ColumnWidth = 4000
    
End Function

Public Function SetEntityCategoryDefault(frm As Form, EntityCategory, Optional createIDControl = True)

    Dim EntityCategoryID
    EntityCategoryID = ELookup("tblEntityCategories", "EntityCategoryName = '" & EntityCategory & "'", "EntityCategoryID")
    
    If createIDControl Then
        Dim ctl As Control
        Set ctl = CreateControl(frm.Name, acComboBox, , "", "EntityCategoryID", 0, 0, 0)
        ctl.Name = "EntityCategoryID"
        SetControlProperties ctl
        ctl.Tag = "alwaysHideOnDatasheet"
        ctl.Visible = False
    Else
        frm("lblEntityName").Caption = EntityCategory & " Name"
        frm("cmdNew").OnClick = "=Save2([Form]," & EscapeString("Entity") & ",0)"
        frm("cmdSaveClose").OnClick = "=Save2([Form]," & EscapeString("Entity") & ",1)"
        frm("cmdDelete").OnClick = "=DeleteRecord([Form]," & EscapeString("EntityID") & "," & EscapeString("tblEntities") & ")"
        If DoesPropertyExists(frm, "pgProperty" & EntityCategory & "s") Then frm("pgProperty" & EntityCategory & "s").Caption = "Properties Viewed"
    End If
    
    ''SetMainFormProperty frm, entityCategory
    frm.BeforeUpdate = "=SaveFormData2([Form]," & EscapeString("Entity") & ")"
    
    frm("EntityCategoryID").DefaultValue = EntityCategoryID
    
End Function


Public Function SetMainFormProperty(frm As Form, EntityCategory)

    'On Error GoTo Err_Handler:
    
    Dim EntityID, entityForm, entityTable
    EntityID = "EntityID"
    entityForm = "frm" & EntityCategory & "s"
    entityTable = "tblEntities"
    
    frm("cmdView").OnClick = "=OpenFormFromMain(" & EscapeString(entityForm) & ",""subform""," & EscapeString(EntityID) & ",[Form])"
    frm("cmdDelete").OnClick = "=DeleteRecord([Form]," & EscapeString(EntityID) & "," & EscapeString(entityTable) & ",""subform"")"
    
'Err_Handler:
'    Debug.Print Err.Number
    
End Function
