Attribute VB_Name = "PropertyBuyer Mod"
Option Compare Database
Option Explicit

Public Function PropertyBuyerCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
            
        Case 5: ''Datasheet Form
            ''frm("BuyerID").OnNotInList = "=PropertyBuyerBuyerIDNotInList([Form])"
            SetPropertyEntityForm frm, "Buyer"
        Case 6: ''Main Form
            
        Case 7: ''Tabular Report
    End Select

End Function


Public Function PropertyBuyerEntityIDAfterUpdate(frm As Form)

    Dim EntityID, PropertyListID
    EntityID = frm("EntityID")
    PropertyListID = frm("PropertyListID")
    
    frm("LastViewedProperty") = PropertyListID
    
End Function

Public Function SetPropertyEntityForm(frm As Form, EntityCategory)
    
    frm("EntityID").Properties("DatasheetCaption") = EntityCategory & " Name"
    frm("EntityID").ListItemsEditForm = ""
    frm("EntityID").rowSource = "SELECT EntityID, EntityName FROM qryEntities WHERE EntityCategoryName = " & EscapeString(EntityCategory) & " ORDER BY EntityName"
    frm.BeforeUpdate = "=SaveFormData2([Form],""PropertyEntity"")"
    
End Function

''Test Run ImportBuyerFromExcel(Forms("mainBuyers"))
Public Function ImportBuyerFromExcel(frm As Form)

    ''Select the Excel file to import (the logic is probably same with Importing Property List from an excel file)
    Dim filePath: filePath = GetFilePath
    If ExitIfTrue(filePath = "", "Please select a valid file..") Then Exit Function
    
    ''Do the actual transfer here
    DoCmd.SetWarnings False
    If DoesPropertyExists(CurrentDb.TableDefs, "tblBuyersTemp") Then RunSQL "DELETE FROM tblBuyersTemp"
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "tblBuyersTemp", filePath, 1
    DoCmd.SetWarnings True

    ''Import all Distinct CustomerName + PhoneNumber to tblEntities
    MakeBuyersTemp2
    ImportToBuyerStatus
    ''ImportTotblEntities
    ImportTotblPropertyEntities
    
    RemovePureSellers
    InsertToSubCategories
    
End Function

Private Function MakeBuyersTemp2()

    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "MAKE"
        .Source = "qryBuyersTemp"
        .MakeTable = "tblBuyersTemp2"
        rowsAffected = .Run
    End With
    
End Function

Public Function ImportTotblPropertyEntities()

    ImportTotblEntities
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblEntities"
        .AddFilter "EntityCategoryID = 1"
        .fields = "EntityID,EntityName"
        sqlStr = .sql
    End With
    
   
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblBuyersTemp2"
        .fields = "EntityID,PropertyListID"
        .Joins.Add GenerateJoinObj(sqlStr, "CustomerName", "temp", "EntityName")
        .Joins.Add GenerateJoinObj("tblPropertyList", "StreetAddress")
        sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "INSERT"
        .Source = "tblPropertyEntities"
        .fields = "EntityID,PropertyListID"
        .InsertSQL = sqlStr
        .InsertUseAsPlain = True
        rowsAffected = .Run
    End With
    
    MsgBox "Buyers successfully imported..."
    
End Function

Private Function ImportToBuyerStatus()
    
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblBuyersTemp2"
        .AddFilter "[CUSTOMER TYPE] <> """""
        .fields = "[CUSTOMER TYPE] As BuyerStatus"
        .OrderBy = "[CUSTOMER TYPE]"
        .GroupBy = "[CUSTOMER TYPE]"
        sqlStr = .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "INSERT"
        .Source = "tblBuyerStatus"
        .fields = "BuyerStatus"
        .InsertSQL = sqlStr
        .InsertUseAsPlain = True
        rowsAffected = .Run
    End With
    
End Function

Private Function ImportTotblEntities()

    ''EntityCategoryID => 1
    Dim sqlObj As clsSQL, joinObj As clsJoin, sqlStr, rowsAffected, rs As Recordset
    Set sqlObj = New clsSQL
    With sqlObj
        .Source = "tblBuyersTemp2"
        .fields = "1 As EntityCategoryID,CustomerName As EntityName,[PHONE NUMBER] As PhoneNumber,[EMAIL ADDRESS] AS EmailAddress, BuyerStatusID, ToBeContacted, ToBeContactedDate, 0 As IsSeller"
        .Joins.Add GenerateJoinObj("tblBuyerStatus", "[CUSTOMER TYPE]", , "BuyerStatus", "LEFT")
        .OrderBy = "CustomerName"
        sqlStr = .sql
        ''makeQuery .sql
    End With
    
    Set sqlObj = New clsSQL
    With sqlObj
        .SQLType = "INSERT"
        .Source = "tblEntities"
        .fields = "EntityCategoryID,EntityName,PhoneNumber,EmailAddress,BuyerStatusID,ToBeContacted,ToBeContactedDate,IsSeller"
        .InsertSQL = sqlStr
        .InsertUseAsPlain = True
        rowsAffected = .Run
    End With
    
End Function

Public Function GetStreetAddress(UnitNumber, HouseNumber, StreetName)
    
    Dim numberArr As New clsArray
    If Not UnitNumber Like "*&*" Then If Not isFalse(UnitNumber) Then numberArr.Add UnitNumber
    If Not isFalse(HouseNumber) Then numberArr.Add HouseNumber
    
    GetStreetAddress = numberArr.JoinArr("/") & " " & StreetName
    
End Function



Private Function GetFilePath() As String

    Dim fd As Office.FileDialog
    Dim strFile As String
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
     
    With fd
     
        .filters.Clear
        .filters.Add "Excel File", "*.xls?", 1
        .title = "Choose an Excel file"
        .AllowMultiSelect = False
     
        If .Show = True Then
     
            GetFilePath = .SelectedItems(1)
            ''Disable this bits for now since I don't know yet if there will be formatting to be done
            Dim xl As Object
            Dim Xw As Object
            Dim sht As Object

            Set xl = CreateObject("Excel.Application")
            Set Xw = GetObject(GetFilePath)
            Set sht = Xw.Worksheets(1)

            xl.Visible = True
            Xw.Activate

            sht.Range("B:B").NumberFormat = "@"
            sht.Range("J:J").NumberFormat = "@"
        
            Xw.Save
            Xw.Close SaveChanges:=True
            xl.Quit
            Set Xw = Nothing
            Set xl = Nothing
     
        End If
     
    End With
    
End Function


