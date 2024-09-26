Attribute VB_Name = "PropertyList Mod"
Option Compare Database
Option Explicit

Public Function PropertyListCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
            SetUpForms frm
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

Public Function GetPropertyListCaption(StreetAddress, Suburb, State, Postcode)

    Dim addressArr As New clsArray
    
    If Not isFalse(StreetAddress) Then
        addressArr.Add StreetAddress
    End If
    
    If Not isFalse(Suburb) Then
        addressArr.Add Suburb
    End If
    
    If Not isFalse(State) Then
        addressArr.Add State
    End If
    
    If Not isFalse(Postcode) Then
        addressArr.Add Postcode
    End If
    
    GetPropertyListCaption = addressArr.JoinArr(" ")
    
End Function

Sub TestComparison()
    Dim result As String
    
    ' Test with null value
    result = GetPropertyOwners("B", Null, "B")
    MsgBox result
    
    ' Test with one valid item
    result = GetPropertyOwners("A")
    MsgBox result
    
    ' Test with no valid item
    result = GetPropertyOwners(Null, Null)
    MsgBox result
    
    ' Test with valid items
    result = GetPropertyOwners("A", "B", "A")
    MsgBox result
End Sub

Private Function GetPropertyOwners(ParamArray items() As Variant) As String
    Dim result As String
    Dim i As Integer
    
    ' Check for null values
    For i = LBound(items) To UBound(items)
        If isFalse(items(i)) Then
            items(i) = ""
        End If
    Next i
    
    ' Filter out duplicates and build the result
    result = ""
    For i = LBound(items) To UBound(items)
        If InStr(result, items(i)) = 0 Then
            If result <> "" Then
                result = result & ","
            End If
            result = result & items(i)
        End If
    Next i
    
    GetPropertyOwners = result
End Function

Public Function GetPropertyIdentifier(StreetAddress, SaleDate, Owner1Name, Owner2Name, Owner3Name) As String
    
    ''Check for the surname of each
    Dim Owner1Surname, Owner2Surname, Owner3Surname
    
    If Not isFalse(Owner1Name) And Owner1Name <> "-" Then Owner1Surname = GetSurname(Owner1Name)
    If Not isFalse(Owner2Name) And Owner2Name <> "-" Then Owner2Surname = GetSurname(Owner2Name)
    If Not isFalse(Owner3Name) And Owner3Name <> "-" Then Owner3Surname = GetSurname(Owner3Name)
    
    Dim PropertyOwners: PropertyOwners = GetPropertyOwners(Owner1Surname, Owner2Surname, Owner3Surname)
    
    Dim identifiers As New clsArray: identifiers.Add StreetAddress
    If Not isFalse(SaleDate) Then identifiers.Add SaleDate
    If Not isFalse(PropertyOwners) Then identifiers.Add PropertyOwners
    
    GetPropertyIdentifier = identifiers.JoinArr(" - ")
    
End Function

Public Function GetPropertyAlternateLink(PropertyListID) As String
    
    If isFalse(PropertyListID) Then Exit Function
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT * FROM tblPropertyList WHERE PropertyListID = " & PropertyListID)
    
    If rs.EOF Then Exit Function
    
    Dim StreetAddress: StreetAddress = rs.fields("StreetAddress")
    Dim Suburb: Suburb = rs.fields("Suburb")
    Dim State: State = rs.fields("State")
    Dim Postcode: Postcode = rs.fields("Postcode")
    Dim OpenInRPP: OpenInRPP = rs.fields("OpenInRPP")
    
    Dim concat: concat = StreetAddress & "-" & Suburb & "-" & State & "-" & Postcode
    concat = replace(concat, " ", "-")
    Dim altLink
    If IsNull(OpenInRPP) Then
        ''https://rpp.rpdata.com/rpp/property/detail.html?propertyId=16118042
        ''https://corelogic.com.au/property/"
        altLink = "https://corelogic.com.au/property/" & concat
    Else
        altLink = replace(OpenInRPP, "rpdata.com/rpp", "corelogic.com.au")
        altLink = replace(altLink, "detail.html?propertyId=", concat & "/")
    End If
    
    
    GetPropertyAlternateLink = altLink
    
End Function

Public Function GetPropertyAltLink(PropertyAltLink)

    If Not isFalse(PropertyAltLink) Then
        GetPropertyAltLink = "Open ALT Link"
    End If
    
End Function

Public Function ExcludeAllToReport(frm As Form, rsName, Optional reversed = False)
    
    
    Dim filterStr
    filterStr = frm("subform").Form.Filter
    
    Dim sqlStr
    sqlStr = "SELECT PropertyListID FROM " & rsName
    If Not isFalse(filterStr) Then
        sqlStr = sqlStr & " WHERE " & filterStr
    End If
    
    ''ExcludeFromReport
    frm("subform").Form.Requery
    Dim Value: Value = -1
    If reversed Then Value = 0
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, rowsAffected, rs As Recordset
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "UPDATE"
        .Source = "tblPropertyList"
        .SetStatement = "ExcludeFromReport = " & Value
        .Joins.Add GenerateJoinObj(sqlStr, "PropertyListID", "temp")
        rowsAffected = .Run
    End With
    
    frm("subform").Form.Requery
    
End Function

Public Function FavoriteClearFilter(frm As Form)
    
    CustomPropertyListClearFilter frm
    frm("fltrFavorites") = -1
    
End Function

Public Function GetPropertyAddress(PropertyListID)
    
    If IsNull(PropertyListID) Then Exit Function
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM tblPropertyList WHERE PropertyListID = " & PropertyListID)
    
    If rs.EOF Then Exit Function
    
    Dim StreetAddress, Suburb, State, Postcode
    StreetAddress = rs.fields("StreetAddress")
    Suburb = rs.fields("Suburb")
    State = rs.fields("State")
    Postcode = rs.fields("Postcode")
    
    Dim addressPart As New clsArray
    If Not IsNull(StreetAddress) Then addressPart.Add StreetAddress
    If Not IsNull(Suburb) Then addressPart.Add Suburb
    If Not IsNull(State) Then addressPart.Add State
    If Not IsNull(Postcode) Then addressPart.Add Postcode
    
    GetPropertyAddress = addressPart.JoinArr(" ")
    
End Function

Public Function IsFavoriteAfterUpdate(frm As Form)

    Dim IsFavorite, PropertyListID, StreetAddress
    IsFavorite = frm("isFavorite")
    PropertyListID = frm("PropertyListID")
    StreetAddress = frm("StreetAddress")
    
    If Not IsFavorite Then Exit Function
    
    ''If isFavorite then create a folder with streetAddress
    Dim uploadDirectory, strFolderExists
    uploadDirectory = GetAttachmentsDirectory(StreetAddress)
    strFolderExists = Dir(uploadDirectory, vbDirectory)
    
    ''Create the directory if it doesn't exist
    If strFolderExists = "" Then
        MkDir uploadDirectory
    End If
    
    ''Run the Owner Import if this is IsFavorite
    ImportPropertySeller frm
    
End Function

Public Function pgPropertyFilesOnClick(frm As Form)

    frm("subPropertyEnityFiles").Form.Requery
    
End Function

Public Function OwnerFieldOnUpdate(frm As Form, OwnerNo, UpdatedField)

    Dim OwnerName, TextField, PropertyListID
    PropertyListID = frm("PropertyListID")
    
    If IsNull(PropertyListID) Then Exit Function
    
    TextField = "Owner" & OwnerNo & "Name"
    OwnerName = frm(TextField)
    
    If IsNull(OwnerName) Or OwnerName = "" Then Exit Function
    
    Dim FieldToFetchValue, FieldToFetchTextField
    FieldToFetchTextField = "Owner" & OwnerNo & UpdatedField
    FieldToFetchValue = frm(FieldToFetchTextField)
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "UPDATE"
        .Source = "tblEntities"
        .SetStatement = UpdatedField & " = " & EscapeString(FieldToFetchValue)
        .AddFilter "EntityName = " & EscapeString(OwnerName) & " AND PropertyListID = " & PropertyListID & " AND EntityCategoryID = 2"
        .Joins.Add GenerateJoinObj("tblPropertyEntities", "EntityID")
        rowsAffected = .Run
    End With
    
    frm("subSellers").Form.Requery
    
End Function

Public Function FetchSellerDetail(frm As Form, OwnerNo, FieldToFetch) As String

    Dim OwnerName, TextField, PropertyListID
    PropertyListID = frm("PropertyListID")
    
    If IsNull(PropertyListID) Then Exit Function
    
    ''Fetch the corresponding owner value
    TextField = "Owner" & OwnerNo & "Name"
    OwnerName = frm(TextField)
    
    If IsNull(OwnerName) Or OwnerName = "" Then Exit Function
    
    Dim FieldToFetchValue, FieldToFetchTextField
    FieldToFetchTextField = "Owner" & OwnerNo & FieldToFetch
    FieldToFetchValue = ELookup("qryPropertySellers", "EntityName = " & EscapeString(OwnerName) & " AND PropertyListID = " & PropertyListID, FieldToFetch)

    frm(FieldToFetchTextField) = FieldToFetchValue
    
End Function

Public Function ReturnToMainDashboard()

    DoCmd.OpenForm "frmCustomDashboard"
    
End Function

Public Function OpenRPDataLink()

    CreateObject("Shell.Application").Open "https://rpp.corelogic.com.au/"
    
End Function

Public Function PropertyListMainFormLoad(frm As Form, Optional PropertyStatusFiltered As Boolean = False)

    CustomMainFormLoad frm
    
    If PropertyStatusFiltered Then
        IsShownOnFavoriteUpdate frm, False
    End If
    
    frm("subform").Form("txtOpenInRPP").ColumnWidth = 1000
    frm("subform").Form("txtPropertyAltlink").ColumnWidth = 1500
    
    HideOrShowColumn frm
    
End Function

Public Function frmPropertyListOnUnload(frm As Form)

    Dim PropertyListID: PropertyListID = frm("PropertyListID")
    If isFalse(PropertyListID) Then Exit Function
    
    SyncBuyerRequirements Forms("frmCustomDashboard"), PropertyListID
    
End Function

Public Function frmPropertyListOnCurrent(frm As Form)

     SetFocusOnForm frm, ""
     ''Fetch the Successful Offer Date
     ''If Successful Offer date is null then the field should be locked
     
     Dim SuccessfulDate, PropertyListID
     PropertyListID = frm("PropertyListID")
     
     SuccessfulDate = GetSuccessfulOfferDate(PropertyListID)
     
     If IsNull(SuccessfulDate) Then
        frm("txtSuccessfulDate").Locked = True
     Else
        frm("txtSuccessfulDate").Locked = False
        frm("txtSuccessfulDate") = Format(SuccessfulDate, "dd-mmm-yy")
     End If
     
     ''Run this only when there's no EventTimeline record for this property
    If Not isFalse(PropertyListID) Then
        If Not isPresent("tblEventTimelines", "PropertyListID = " & PropertyListID) Then
            InsertTo_tblEventTimelines frm
        End If
    End If

     RefreshPropertyLedger
     
    If Not isFalse(PropertyListID) Then
        ''Update the OriginalOwner of this Property
        Update_IsOriginalOwner_tblPropertyEntities PropertyListID
    End If
    
End Function

Private Sub Update_IsOriginalOwner_tblPropertyEntities(PropertyListID)
    
    Dim EarliestTimestamp: EarliestTimestamp = ELookupDate("qryPropertyEntities", "PropertyListID = " & PropertyListID & _
        " AND EntityCategoryName = ""Seller""", "[Timestamp]", "[Timestamp]")
    
    Dim PropertyEntityIDs: PropertyEntityIDs = Elookups("qryPropertyEntities", "[Timestamp] = #" & SQLDate(EarliestTimestamp) & _
        "# AND PropertyListID = " & PropertyListID & " AND EntityCategoryName = ""SELLER""", "PropertyEntityID")
    
    RunSQL "UPDATE tblPropertyEntities SET IsOriginalOwner = -1 WHERE PropertyEntityID In(" & PropertyEntityIDs & ")"
    
End Sub

Public Sub InsertTo_tblEventTimelines(frm As Form)
    
    Dim PropertyListID: PropertyListID = frm("PropertyListID")
    
    If isFalse(PropertyListID) Then Exit Sub
    
    Dim rs As Recordset: Set rs = ReturnRecordset("SELECT EventName FROM tblEventList order by EventOrder")
    Dim i As Integer: i = 1
    Do Until rs.EOF
        Dim EventName: EventName = rs.fields("EventName")
        RunSQL "INSERT INTO tblEventTimelines (PropertyListID,EventOrder,Description) VALUES (" & PropertyListID & _
            "," & i & "," & Esc(EventName) & ")"
            i = i + 1
        rs.MoveNext
    Loop
    
    frm("dshtEventTimelines").Form.Requery
    
End Sub


Public Function frmPropertyListtxtSuccessfulDateAfterUpdate(frm As Form)

    Dim SuccessfulDate, PropertyListID
    PropertyListID = frm("PropertyListID")
    SuccessfulDate = frm("txtSuccessfulDate")
    
    If IsNull(PropertyListID) Then Exit Function
    If IsNull(SuccessfulDate) Then Exit Function
    
    If ExitIfTrue(Not IsDate(SuccessfulDate), "Please enter a valid date/time value.") Then Exit Function
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM qryBuyerOffers WHERE PropertyListID = " & PropertyListID & " AND IsAccepted")
    
    If Not rs.EOF Then
        Dim BuyerOfferID
        BuyerOfferID = rs.fields("BuyerOfferID")
        RunSQL "UPDATE tblBuyerOffers SET [Timestamp] = #" & SQLDate(SuccessfulDate) & "# WHERE BuyerOfferID = " & BuyerOfferID
    End If
    
    frm("subBuyerOffers").Form.Requery
    
End Function

Public Function GetSuccessfulBuyer(PropertyListID)

    If IsNull(PropertyListID) Then Exit Function
    GetSuccessfulBuyer = ELookup("qryBuyerOffers", "PropertyListID = " & PropertyListID & " AND IsAccepted", "EntityID")

End Function

Public Function GetSuccessfulOffer(PropertyListID) As Double

    If IsNull(PropertyListID) Then Exit Function
    Dim offer
    offer = ELookup("qryBuyerOffers", "PropertyListID = " & PropertyListID & " AND IsAccepted", "Offer")
    If IsNumeric(offer) Then
        GetSuccessfulOffer = offer
    End If

End Function

Public Function GetSuccessfulOfferDate(PropertyListID)

    If IsNull(PropertyListID) Then Exit Function
    Dim offerDate
    offerDate = ELookup("qryBuyerOffers", "PropertyListID = " & PropertyListID & " AND IsAccepted", "Timestamp")
    If IsDate(offerDate) Then
        GetSuccessfulOfferDate = offerDate
    End If

End Function

Public Function GetCommissionAmount(AppraisedAmount, SuccessfulOffer, CommissionRate)
    
    Dim vAppraisedAmount, vSuccessfulOffer, vCommissionRate
    
    vAppraisedAmount = AppraisedAmount
    vSuccessfulOffer = SuccessfulOffer
    
    Dim BaseAmount: BaseAmount = vAppraisedAmount
    If Not isFalse(vSuccessfulOffer) Then
        BaseAmount = vSuccessfulOffer
    End If
    
    If isFalse(BaseAmount) Then BaseAmount = 0
    
    vCommissionRate = IIf(IsNull(CommissionRate), "0", CommissionRate)
    
    GetCommissionAmount = BaseAmount * vCommissionRate
    
End Function

Public Function CustomPropertyListClearFilter(frm As Form)
    
    Dim ctl As Control
    For Each ctl In frm.Controls
        If ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Then ctl.Value = Null
        If ctl.ControlType = acCheckBox Then ctl.Value = 0
    Next ctl
    
    ClearFilterSubform frm, 463
    
End Function

Public Function CustomRecipientListFilter(frm As Form)
    
    If IsFormOpen("mainPropertyBulkEmail") Then
        
        Dim fileNameArr As New clsArray
        Dim filterArr As New clsArray, filterStr
        ''Wildcard Search
        ''StreetAddress,Suburb,State,PropertyType,Postcode,Owner1Name,Owner2Name,Owner3Name
        ''fltrWildSearch
        GetWildCardFilter frm, filterArr, fileNameArr
        GetSuburbFilter frm, filterArr, fileNameArr
        GetPropertyStatusFilter frm, filterArr, fileNameArr
        GetFavoriteFilter frm, filterArr, fileNameArr
        GetPropertyTypeFilter frm, filterArr, fileNameArr
        GetBedFilter frm, filterArr, fileNameArr
        GetBathFilter frm, filterArr, fileNameArr
        GetCarFilter frm, filterArr, fileNameArr
        GetLandSizeFilter frm, filterArr, fileNameArr
        GetLandUseFilter frm, filterArr, fileNameArr
        GetSalePriceFilter frm, filterArr, fileNameArr
        GetSaleDateFilter frm, filterArr, fileNameArr
        ''Add fltrZoning & fltrDevelopmentZone
        GetZoningFilter frm, filterArr, fileNameArr
        GetDevelopmentZoneFilter frm, filterArr, fileNameArr
        
        GetNumberRangeFilter frm, filterArr, fileNameArr, "AppraisedAmount"
        
        ''fileNameArr.Add Format$(Now, "ddmmyyyyhhnnss")
        filterStr = filterArr.JoinArr(" AND ")
        
        ''fileName = fileNameArr.JoinArr("-")
        ''fileName = SanitizeFileName(fileName)
        Dim sqlStr
        sqlStr = "SELECT PropertyListID FROM tblPropertyList"
        
        If filterStr <> "" Then
            sqlStr = sqlStr & " WHERE " & filterStr
        End If
        
        Set frm = Forms("mainPropertyBulkEmail")
        
        frm.subform.Form.Filter = "PropertyListID In (" & sqlStr & ")"
        frm.subform.Form.FilterOn = True
    End If
    
End Function

Private Function InsertToFilteredPropertyBySuburb(fltrSuburb, IsFavorite)
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "tblPropertyList"
          .AddFilter "Suburb = " & EscapeString(fltrSuburb)
          If IsFavorite Then
            .AddFilter "IsFavorite"
          End If
          .fields = "PropertyListID,StreetAddress"
          sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "tblPropertyStreetPreFilteredBySuburbs"
          .fields = "PropertyListID,StreetAddress"
          .InsertSQL = sqlStr
          .InsertFilterField = "PropertyListID,StreetAddress"
          rowsAffected = .Run
    End With
End Function

Public Function ContainsSubstringWithCondition(inputString, substring) As Boolean
    Dim index As Integer
    
    ' Find the position of the substring in the input string
    index = InStr(inputString, substring)
    
    ' Check if the substring is found and if it is at the end of the input string
    If index > 0 And index + Len(substring) - 1 = Len(inputString) Then
        ContainsSubstringWithCondition = True
    Else
        ContainsSubstringWithCondition = False
    End If
End Function

Public Function SyncPropertyStreetPreFiltereds(frm As Form, filterArr As clsArray, fileNameArr As clsArray)

    RunSQL "DELETE FROM tblPropertyStreetPreFiltereds"
    
    Dim inputValue
    inputValue = frm("fltrWildSearch")
    
    ''Filter by suburb first
    Dim fltrSuburb: fltrSuburb = frm("fltrSuburb")
    
    If (IsNull(inputValue) Or inputValue = "") Then Exit Function
    If (isFalse(fltrSuburb)) Then Exit Function
    
    
    InsertToFilteredPropertyBySuburb fltrSuburb, frm.Name = "mainFavoriteProperties"

    Dim inputValueArr As New clsArray, item: inputValueArr.arr = inputValue
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    Dim totalRecords: totalRecords = inputValueArr.Count
    
    Dim cls_ProgressBar As New cls_ProgressBar
    
    cls_ProgressBar.ProgressBar_Show
    cls_ProgressBar.ProgressBar_ProgressOverlay True
    cls_ProgressBar.ProgressBar_Caption "Filtering Data"
    ''cls_ProgressBar.ProgressBar_Resize 3000, 6000
    cls_ProgressBar.ProgressBar_Message1_FontParam 1, "Calibri", 400, 10, vbBlack
    cls_ProgressBar.ProgressBar_ProgressValue_Align 2

    Dim i: i = 1
    For Each item In inputValueArr.arr
        Dim TrimmedItem: TrimmedItem = Trim(item)
        cls_ProgressBar.ProgressBar_Progress (i / totalRecords)
        cls_ProgressBar.ProgressBar_Message1 "Narrowing search by: " & TrimmedItem
        ''wildCardArrFilter.Add wildCardItem & " Like '*" & TrimmedItem & "*'"
        Set rs = ReturnRecordset("Select AddressText,Variants from tblAddressVariants ORDER BY Variants")
        Do Until rs.EOF
            Dim Variants: Variants = rs.fields("Variants")
            Dim AddressText: AddressText = rs.fields("AddressText")
            
            If ContainsSubstringWithCondition(TrimmedItem, Variants) > 0 Then
                TrimmedItem = replace(TrimmedItem, " " & Variants, " " & AddressText)
                InsertTo_tblPropertyStreetPreFiltereds TrimmedItem
                Exit Do
            End If
            
            rs.MoveNext
        Loop
        
        InsertTo_tblPropertyStreetPreFiltereds TrimmedItem
        
        i = i + 1
        
    Next item
    
    If Not isFalse(inputValue) Then
        filterArr.Add "StreetFilterPassed"
    End If
    
    cls_ProgressBar.ProgressBar_Hide
    
End Function

Private Function InsertTo_tblPropertyStreetPreFiltereds(TrimmedItem)
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    
    Set sqlObj = New clsSQL
    With sqlObj
          .Source = "tblPropertyStreetPreFilteredBySuburbs"
          .AddFilter "StreetAddress Like " & EscapeString("*" & TrimmedItem & "*")
          .fields = "PropertyListID"
          sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
          .SQLType = "INSERT"
          .Source = "tblPropertyStreetPreFiltereds"
          .fields = "PropertyListID"
          .InsertSQL = sqlStr
          .InsertFilterField = "PropertyListID"
          rowsAffected = .Run
    End With
    
End Function

Public Function CustomPropertyListFilter(frm As Form, Optional OnBuyerForm As Boolean = False, Optional fileName = "")

    Dim fileNameArr As New clsArray
    Dim filterArr As New clsArray, filterStr
    ''Wildcard Search
    ''StreetAddress,Suburb,State,PropertyType,Postcode,Owner1Name,Owner2Name,Owner3Name
    ''fltrWildSearch
    
    SyncPropertyStreetPreFiltereds frm, filterArr, fileNameArr
    ''GetWildCardFilter frm, filterArr, fileNameArr
    GetSuburbFilter frm, filterArr, fileNameArr
    GetPropertyStatusFilter frm, filterArr, fileNameArr
    GetFavoriteFilter frm, filterArr, fileNameArr
    GetPropertyTypeFilter frm, filterArr, fileNameArr
    GetBedFilter frm, filterArr, fileNameArr
    GetBathFilter frm, filterArr, fileNameArr
    GetCarFilter frm, filterArr, fileNameArr
    GetLandSizeFilter frm, filterArr, fileNameArr
    GetLandUseFilter frm, filterArr, fileNameArr
    GetSalePriceFilter frm, filterArr, fileNameArr
    GetSaleDateFilter frm, filterArr, fileNameArr
    ''Add fltrZoning & fltrDevelopmentZone
    GetZoningFilter frm, filterArr, fileNameArr
    GetDevelopmentZoneFilter frm, filterArr, fileNameArr
    
    ''Add filter for the Appraised Amount
    GetNumberRangeFilter frm, filterArr, fileNameArr, "AppraisedAmount"
    
    fileNameArr.Add Format$(Now, "ddmmyyyyhhnnss")
    filterStr = filterArr.JoinArr(" AND ")
    
    fileName = fileNameArr.JoinArr("-")
    fileName = SanitizeFileName(fileName)
    
    If OnBuyerForm Then
        Dim sqlStr
        sqlStr = "SELECT PropertyListID FROM tblPropertyList"
        
        If filterStr <> "" Then
            sqlStr = sqlStr & " WHERE " & filterStr
        End If
        
        frm.subform.Form.Filter = "LastViewedProperty In (" & sqlStr & ")"
        frm.subform.Form.FilterOn = True
    Else
        frm.subform.Form.Filter = filterStr
        frm.subform.Form.FilterOn = True
    End If
    
End Function

Private Function GetSuburbFilter(frm As Form, filterArr As clsArray, fileNameArr As clsArray)
    
    Dim fltrSuburb
    fltrSuburb = frm("fltrSuburb")
    
    If Not IsNull(fltrSuburb) Then
        fileNameArr.Add "Suburb" & frm("fltrSuburb")
        filterArr.Add "Suburb = " & EscapeString(fltrSuburb)
    End If

End Function


Private Function GetPropertyStatusFilter(frm As Form, filterArr As clsArray, fileNameArr As clsArray)
    
    Dim fltrPropertyStatusID
    fltrPropertyStatusID = frm("fltrPropertyStatusID")
    
    If Not IsNull(fltrPropertyStatusID) Then
        fileNameArr.Add "Status" & frm("fltrPropertyStatusID").Column(1)
        filterArr.Add "PropertyStatusID = " & fltrPropertyStatusID
    End If

End Function

Private Function GetSaleDateFilter(frm As Form, filterArr As clsArray, fileNameArr As clsArray)
    
    Dim fltrSaleDateGt, fltrSaleDateLt
    fltrSaleDateGt = frm("fltrSaleDateGt")
    fltrSaleDateLt = frm("fltrSaleDateLt")
    
    Dim tagArr As New clsArray
    
    If Not IsNull(fltrSaleDateGt) Then
        tagArr.Add "GT" & Format$(fltrSaleDateGt, "dd-mmm-yy")
        filterArr.Add "SaleDate >= #" & fltrSaleDateGt & "#"
    End If
    
    If Not IsNull(fltrSaleDateLt) Then
        tagArr.Add "LT" & Format$(fltrSaleDateLt, "dd-mmm-yy")
        filterArr.Add "SaleDate <= #" & fltrSaleDateLt & "#"
    End If
    
    If Not IsNull(fltrSaleDateGt) And Not IsNull(fltrSaleDateLt) Then
        fileNameArr.Add "SaleDate" & tagArr.JoinArr("-")
    End If

End Function

Private Function GetNumberRangeFilter(frm As Form, filterArr As clsArray, fileNameArr As clsArray, FieldName)

    Dim flrtGt, fltrLt
    flrtGt = frm("fltr" & FieldName & "Gt")
    fltrLt = frm("fltr" & FieldName & "Lt")
    
    Dim tagArr As New clsArray
    
    If Not IsNull(flrtGt) Then
        tagArr.Add "GT" & flrtGt
        filterArr.Add FieldName & " >= " & flrtGt
    End If
    
    If Not IsNull(fltrLt) Then
        tagArr.Add "LT" & fltrLt
        filterArr.Add FieldName & " <= " & fltrLt
    End If
    
    If Not IsNull(flrtGt) And Not IsNull(fltrLt) Then
        fileNameArr.Add FieldName & tagArr.JoinArr("-")
    End If
    
End Function

Private Function GetSalePriceFilter(frm As Form, filterArr As clsArray, fileNameArr As clsArray)
    
    Dim fltrSalePriceGt, fltrSalePriceLt
    fltrSalePriceGt = frm("fltrSalePriceGt")
    fltrSalePriceLt = frm("fltrSalePriceLt")
    
    Dim tagArr As New clsArray
    
    If Not IsNull(fltrSalePriceGt) Then
        tagArr.Add "GT" & fltrSalePriceGt
        filterArr.Add "SalePrice >= " & fltrSalePriceGt
    End If
    
    If Not IsNull(fltrSalePriceLt) Then
        tagArr.Add "LT" & fltrSalePriceLt
        filterArr.Add "SalePrice <= " & fltrSalePriceLt
    End If
    
    If Not IsNull(fltrSalePriceGt) And Not IsNull(fltrSalePriceLt) Then
        fileNameArr.Add "SalePrice" & tagArr.JoinArr("-")
    End If

End Function

Private Function GetZoningFilter(frm As Form, filterArr As clsArray, fileNameArr As clsArray)
    
    Dim fltrZoning
    fltrZoning = frm("fltrZoning")
    
    If Not IsNull(fltrZoning) Then
        fileNameArr.Add "Zoning" & fltrZoning
        filterArr.Add "Zoning = '" & fltrZoning & "'"
    End If

End Function

Private Function GetDevelopmentZoneFilter(frm As Form, filterArr As clsArray, fileNameArr As clsArray)
    ''DevelopmentZone
    Dim fltrDevelopmentZone
    fltrDevelopmentZone = frm("fltrDevelopmentZone")
    
    If Not IsNull(fltrDevelopmentZone) Then
        fileNameArr.Add "DevelopmentZone" & fltrDevelopmentZone
        filterArr.Add "DevelopmentZone = '" & fltrDevelopmentZone & "'"
    End If

End Function

Private Function GetLandUseFilter(frm As Form, filterArr As clsArray, fileNameArr As clsArray)
    
    Dim fltrLandUse
    fltrLandUse = frm("fltrLandUse")
    
    If Not IsNull(fltrLandUse) Then
        fileNameArr.Add "LandUse" & fltrLandUse
        filterArr.Add "LandUse1 = '" & fltrLandUse & "'"
    End If

End Function

Private Function GetLandSizeFilter(frm As Form, filterArr As clsArray, fileNameArr As clsArray)
    
    Dim fltrLandSizeGt, fltrLandSizeLt
    fltrLandSizeGt = frm("fltrLandSizeGt")
    fltrLandSizeLt = frm("fltrLandSizeLt")
    
    Dim tagArr As New clsArray
    
    If Not IsNull(fltrLandSizeGt) Then
        tagArr.Add "GT" & fltrLandSizeGt
        filterArr.Add "LandSize >= " & fltrLandSizeGt
    End If
    
    If Not IsNull(fltrLandSizeLt) Then
        tagArr.Add "LT" & fltrLandSizeLt
        filterArr.Add "LandSize <= " & fltrLandSizeLt
    End If
    
    If Not IsNull(fltrLandSizeGt) And Not IsNull(fltrLandSizeLt) Then
        fileNameArr.Add "LandSize" & tagArr.JoinArr("-")
    End If
    
    
End Function

Private Function GetCarFilter(frm As Form, filterArr As clsArray, fileNameArr As clsArray)
    
    ''Wildcard Search
    ''Car
    ''fltrCar1
    
    Dim wildCardArrFilter As New clsArray
    Dim tagArr As New clsArray
    
    Dim ctl
    For Each ctl In frm.Controls
        If ctl.Name Like "fltrCar*" Then
            If ctl.Value Then
                
                If ctl.Tag = "4" Then
                    tagArr.Add ">=4"
                    wildCardArrFilter.Add "Car >= 4"
                Else
                    tagArr.Add ctl.Tag
                    wildCardArrFilter.Add "Car = " & ctl.Tag
                End If
                
                
            End If
        End If
    Next ctl
    
    If wildCardArrFilter.Count = 0 Then Exit Function
    
    fileNameArr.Add "Car" & tagArr.JoinArr("-")
    filterArr.Add "(" & wildCardArrFilter.JoinArr(" OR ") & ")"
    
End Function


Private Function GetBedFilter(frm As Form, filterArr As clsArray, fileNameArr As clsArray)
    
    ''Wildcard Search
    ''Bed
    ''fltrBed1
    
    Dim wildCardArrFilter As New clsArray
    Dim tagArr As New clsArray
    
    Dim ctl
    For Each ctl In frm.Controls
        If ctl.Name Like "fltrBed*" Then
            If ctl.Value Then
                
                If ctl.Tag = "6" Then
                    tagArr.Add ">=6"
                    wildCardArrFilter.Add "Bed >= 6"
                Else
                    tagArr.Add ctl.Tag
                    wildCardArrFilter.Add "Bed = " & ctl.Tag
                End If
                
                
            End If
        End If
    Next ctl
    
    If wildCardArrFilter.Count = 0 Then Exit Function
    
    fileNameArr.Add "Bed" & tagArr.JoinArr("-")
    filterArr.Add "(" & wildCardArrFilter.JoinArr(" OR ") & ")"
    
End Function

Private Function GetBathFilter(frm As Form, filterArr As clsArray, fileNameArr As clsArray)
    
    ''Wildcard Search
    ''Bath
    ''fltrBathroom1
    
    Dim wildCardArrFilter As New clsArray
    Dim tagArr As New clsArray
    
    Dim ctl
    For Each ctl In frm.Controls
        If ctl.Name Like "fltrBathroom*" Then
            If ctl.Value Then
                
                If ctl.Tag = "4" Then
                    tagArr.Add ">=4"
                    wildCardArrFilter.Add "Bath >= 4"
                Else
                    tagArr.Add ctl.Tag
                    wildCardArrFilter.Add "Bath = " & ctl.Tag
                End If
                
                
            End If
        End If
    Next ctl
    
    If wildCardArrFilter.Count = 0 Then Exit Function
    
    fileNameArr.Add "Bathroom" & tagArr.JoinArr("-")
    filterArr.Add "(" & wildCardArrFilter.JoinArr(" OR ") & ")"
    
End Function

Private Function GetFavoriteFilter(frm As Form, filterArr As clsArray, fileNameArr As clsArray)

    ''IsFavorite
    ''fltrFavorites
    Dim IsFavorite
    IsFavorite = frm("fltrFavorites")
    
    fileNameArr.Add "Favorite" & IsFavorite
    If IsFavorite Then filterArr.Add "(IsFavorite)"
    
End Function


Private Function GetWildCardFilter(frm As Form, filterArr As clsArray, fileNameArr As clsArray)
    
    ''Wildcard Search
    ''StreetAddress,Suburb,State,PropertyType,Postcode,Owner1Name,Owner2Name,Owner3Name
    ''fltrWildSearch
    Dim wildCardArr As New clsArray, wildCardItem, wildCardArrFilter As New clsArray
    wildCardArr.arr = "StreetAddress"
    
    Dim inputValue
    inputValue = frm("fltrWildSearch")
    
    If (IsNull(inputValue) Or inputValue = "") Then Exit Function
    
    Dim inputValueArr As New clsArray, item: inputValueArr.arr = inputValue
    fileNameArr.Add "q" & inputValue
    
    For Each wildCardItem In wildCardArr.arr
        For Each item In inputValueArr.arr
            Dim TrimmedItem: TrimmedItem = Trim(item)
            ''wildCardArrFilter.Add wildCardItem & " Like '*" & TrimmedItem & "*'"
            Dim rs As Recordset: Set rs = ReturnRecordset("Select AddressText,Variants from tblAddressVariants ORDER BY Variants")
            Do Until rs.EOF
                Dim Variants: Variants = rs.fields("Variants")
                Dim AddressText: AddressText = rs.fields("AddressText")
                TrimmedItem = replace(TrimmedItem, " " & Variants, " " & AddressText)
                rs.MoveNext
            Loop
            wildCardArrFilter.Add wildCardItem & " Like '*" & TrimmedItem & "*'"
        Next item
    Next wildCardItem
    
    Debug.Print "(" & wildCardArrFilter.JoinArr(" OR ") & ")"
    filterArr.Add "(" & wildCardArrFilter.JoinArr(" OR ") & ")"
    
End Function

Private Function GetPropertyTypeFilter(frm As Form, filterArr As clsArray, fileNameArr As clsArray)
    
    ''Wildcard Search
    ''PropertyType
    ''fltrPropertyTypeHouses
    
    Dim wildCardArrFilter As New clsArray
    Dim tagArr As New clsArray
    
    Dim ctl
    For Each ctl In frm.Controls
        If ctl.Name Like "fltrPropertyType*" Then
            If ctl.Value Then
                tagArr.Add ctl.Tag
                wildCardArrFilter.Add "PropertyType Like '*" & ctl.Tag & "*'"
                
            End If
        End If
    Next ctl
    
    If wildCardArrFilter.Count = 0 Then Exit Function
    
    fileNameArr.Add "PropertyType" & tagArr.JoinArr("-")
    filterArr.Add "(" & wildCardArrFilter.JoinArr(" OR ") & ")"
    
End Function

Public Function ShowPropertyStatusFilterForm()

'    Dim frm As Form, subFrm As Form
'    Set frm = Forms("mainPropertyList")
'    Set subFrm = frm("subform").Form
'
'    Dim ctl As Control, ctlName, ctlCaption, i As Integer
'    i = 1
'    For Each ctl In subFrm.Controls
'
'        ctlName = ctl.Name
'        ctlCaption = ctl.Properties("DataSheetCaption")
'
'
'
'        If Not ctlName Like "Sum*" Then
'            RunSQL "INSERT INTO tblPropertyListFields (FieldName,FieldCaption,FieldOrder) VALUES ('" & ctlName & "','" & ctlCaption & "'," & i & ")"
'            i = i + 1
'        End If
'
'    Next ctl
    
    DoCmd.OpenForm "mainPropertyStatusFilter"
    
End Function

Public Function ShowHidePropListFields(Optional frmName = "mainPropertyListFields")

'    Dim frm As Form, subFrm As Form
'    Set frm = Forms("mainPropertyList")
'    Set subFrm = frm("subform").Form
'
'    Dim ctl As Control, ctlName, ctlCaption, i As Integer
'    i = 1
'    For Each ctl In subFrm.Controls
'
'        ctlName = ctl.Name
'        ctlCaption = ctl.Properties("DataSheetCaption")
'
'
'
'        If Not ctlName Like "Sum*" Then
'            RunSQL "INSERT INTO tblPropertyListFields (FieldName,FieldCaption,FieldOrder) VALUES ('" & ctlName & "','" & ctlCaption & "'," & i & ")"
'            i = i + 1
'        End If
'
'    Next ctl
    
    DoCmd.OpenForm frmName
    
End Function

Private Function SetUpForms(frm As Form)
    
    frm.AllowAdditions = False
    
End Function

Public Function OpenPropertyAltLink(PropertyListID)
    
    Dim PropertyAlternateLink: PropertyAlternateLink = GetPropertyAlternateLink(PropertyListID)
    If isFalse(PropertyAlternateLink) Then Exit Function
    CreateObject("Shell.Application").Open PropertyAlternateLink
    
End Function

Public Function OpenPropertyLink(frm As Form, Optional fldName = "OpenInRPP")

    Dim OpenInRPP
    OpenInRPP = frm(fldName)
    
    If IsNull(OpenInRPP) Or OpenInRPP = "" Then
        Exit Function
    End If
    
    CreateObject("Shell.Application").Open OpenInRPP
    
End Function


Public Function PropertyListFormLoad(frm As Form)
    
    If IsFormOpen("frmTaskCalendar") Then DoCmd.Close acForm, "frmTaskCalendar", acSaveNo
    
    DefaultFormLoad frm, "PropertyListID", True
    
    Dim EntityArr As New clsArray, Entity
    EntityArr.arr = "Buyer,Tenant,Contact,Seller"
    For Each Entity In EntityArr.arr
        frm("sub" & Entity & "s").Form.Controls("EntityName").ColumnHidden = True
        frm("sub" & Entity & "s").Form.Controls("StreetAddress").ColumnHidden = False
        If Entity <> "Seller" Then
            frm("sub" & Entity & "Members").Form("EntityID").ColumnHidden = True
        Else
            ''Cancel Out the Seller tab
            frm("sub" & Entity & "s").Form.Controls("EntityID").ColumnHidden = False
            frm("sub" & Entity & "s").Form.Controls("EntityName").ColumnHidden = True
            ''frm("sub" & entity & "s").Form.Controls("EntityName").columnWidth = -2
        End If
        frm("sub" & Entity & "Files").Form("EntityID").ColumnHidden = True
        frm("sub" & Entity & "Notes").Form("EntityID").ColumnHidden = True
        frm("sub" & Entity & "Notes").Form("Note").ColumnWidth = 14000
        
        frm("sub" & Entity & "Files").Form("EntityFileLink").ColumnWidth = 9000
        
    Next Entity
    
    frm("subBuyers").Form("LastViewedProperty").ColumnHidden = True
    frm("subBuyers").Form("txtOpenInRPP").ColumnHidden = True
    frm("subSellers").Form("txtOpenInRPP").ColumnHidden = True

    frm("subBuyerOffers").Form("CreatedBy").ColumnHidden = True
    frm("subBuyerOffers").Form("Timestamp").ColumnWidth = 3000
    frm("subBuyerOffers").Form("PropertyEntityID").ColumnHidden = True
    
    frm("subBuyers").Form("PropertyEntityNote").ColumnHidden = True
    
    frm("subPropertyEnityNotes").Form("Note").ColumnWidth = 10000
    frm("subPropertyEnityNotes").Form("CreatedBy").ColumnHidden = True
    frm("subPropertyEnityNotes").Form("Timestamp").ColumnWidth = 3000
    frm("subPropertyEnityNotes").Form("EntityName").ColumnWidth = 2000
    frm("subPropertyEnityNotes").Form("EntityCategoryName").ColumnWidth = 2000
    
    frm("subTenants").Form("StreetAddress").ColumnHidden = True
    
    frm("subContacts").Form("vPropertyListID").ColumnWidth = 0
    frm("subBuyers").Form("vPropertyListID").ColumnWidth = 0
    frm("subSellers").Form("vPropertyListID").ColumnWidth = 0
    frm("subTenants").Form("vPropertyListID").ColumnWidth = 0
    
End Function

Public Function OpenEntityFileForm(frm As Form, EntityCategory, Optional fromBuyerForm = False)

    Dim fieldValue, FieldName, objectName, PropertyListID
    PropertyListID = frm("PropertyListID")
    objectName = EntityCategory
    FieldName = "EntityID"
    If fromBuyerForm Then
        fieldValue = frm(FieldName)
    Else
        fieldValue = frm("sub" & objectName & "s")(FieldName)
    End If
    
    If ExitIfTrue(IsNull(PropertyListID), "Property is empty...") Then Exit Function
    If ExitIfTrue(IsNull(fieldValue), "Please select a valid " & objectName) Then Exit Function
    
    Dim frm2 As Form
    DoCmd.OpenForm "frmSimple" & objectName & "Files"
    Set frm2 = Forms("frmSimple" & objectName & "Files")
    
    frm2("PropertyListID") = PropertyListID
    frm2(FieldName) = fieldValue
    frm2(FieldName).AllowValueListEdits = False
    frm2(FieldName).Locked = True
    
    If fromBuyerForm Then
        frm2.OnUnload = "=RefreshBuyerForm()"
    Else
        frm2.OnUnload = "=PropertyListRefreshSubforms()"
    End If
    
End Function

Public Function OpenTenantFileForm(frm As Form)

    Dim TenantID
    TenantID = frm.subTenants("TenantID")
    
    If ExitIfTrue(IsNull(TenantID), "Please select a valid Tenant") Then Exit Function
    
    Dim frm2 As Form
    DoCmd.OpenForm "frmSimpleTenantFiles"
    Set frm2 = Forms("frmSimpleTenantFiles")
    
    frm2("TenantID") = TenantID
    frm2("TenantID").AllowValueListEdits = False
    frm2("TenantID").Locked = True
    
    frm2.OnUnload = "=PropertyListRefreshSubforms()"
    
End Function

Public Function OpenBuyerFileForm(frm As Form)

    Dim fieldValue, FieldName, objectName
    objectName = "Buyer"
    FieldName = "EntityID"
    fieldValue = frm.subBuyers(FieldName)
    
    If ExitIfTrue(IsNull(fieldValue), "Please select a valid " & objectName) Then Exit Function
    
    Dim frm2 As Form
    DoCmd.OpenForm "frmSimple" & objectName & "Files"
    Set frm2 = Forms("frmSimple" & objectName & "Files")
    
    frm2(FieldName) = fieldValue
    frm2(FieldName).AllowValueListEdits = False
    frm2(FieldName).Locked = True
    
    frm2.OnUnload = "=PropertyListRefreshSubforms()"
    
End Function

Public Function OpenCommunityFileForm(frm As Form)

    Dim PropertyCommunityID
    PropertyCommunityID = frm.subCommunities("PropertyCommunityID")
    
    If ExitIfTrue(IsNull(PropertyCommunityID), "Please select a valid community") Then Exit Function
    
    Dim frm2 As Form
    DoCmd.OpenForm "frmSimpleCommunityFiles"
    Set frm2 = Forms("frmSimpleCommunityFiles")
    
    frm2("PropertyCommunityID") = PropertyCommunityID
    frm2("PropertyCommunityID").AllowValueListEdits = False
    frm2("PropertyCommunityID").Locked = True
    
    frm2.OnUnload = "=PropertyListRefreshSubforms()"
    
End Function

Public Function PropertyListRefreshSubforms()
    
    Dim frm As Form
    Set frm = Forms("frmPropertyList")
    
    Dim subformArr As New clsArray, subformItem
    subformArr.arr = "subTenantFiles,subBuyerFiles,subContactFiles,subSellerFiles,subPropertyEnityFiles"
    For Each subformItem In subformArr.arr
        If DoesPropertyExists(frm, subformItem) Then
            frm(subformItem).Form.Requery
        End If
    Next subformItem
    
    
End Function
