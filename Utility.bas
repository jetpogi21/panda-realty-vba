Attribute VB_Name = "Utility"
Option Compare Database
Option Explicit

Public ProgressPopup_StartTime As Long
Public ProgressPopup_LastProgress As Double

Public Sub D(str)

    Dim strArr As New clsArray: strArr.arr = str
    
    Dim lines As New clsArray
    
    Dim item, i As Integer: i = 0
    For Each item In strArr.arr
        lines.Add IIf(i > 0, vbTab, "") & replace("Dim [item]: [item] = rs.Fields(""[item]"")", "[item]", item)
        i = i + 1
    Next item
    
    CopyToClipboard lines.JoinArr(vbNewLine)
    
End Sub

Public Sub rs()
    Dim str: str = "Dim rs As Recordset: Set rs = ReturnRecordset(sqlStr)"
    ''Debug.Print str
    CopyToClipboard str
End Sub

Public Sub rsLoop()

    Dim strs As New clsArray
    strs.Add "Do until rs.EOF"
    strs.Add vbTab & vbTab & "rs.Movenext"
    strs.Add vbTab & "Loop"
    
    Dim str: str = strs.JoinArr(vbNewLine)
    CopyToClipboard str
    
End Sub

Public Function TableExists(filePath, tableName) As Boolean
    Dim db As DAO.Database
    Dim tbl As DAO.TableDef
    
    ' Initialize the database object
    On Error Resume Next ' Handle the case where the database file does not exist
    Set db = OpenDatabase(filePath)
    On Error GoTo 0 ' Revert to default error handling
    
    If Not db Is Nothing Then
        ' Check if the table already exists in the destination
        For Each tbl In db.TableDefs
            If tbl.Name = tableName Then
                ' Clean up
                Set tbl = Nothing
                Set db = Nothing
                TableExists = True
                Exit Function
            End If
        Next tbl
        ' Clean up
        Set db = Nothing
    End If
    
    ' If the table doesn't exist
    TableExists = False
End Function

Sub MsgBoxResult(inputText As String, result As Variant)
    MsgBox "Input Text: " & vbNewLine & inputText & vbNewLine & vbNewLine & "Result: " & result, vbInformation, "Test Result"
End Sub

Function RemoveChars(ByVal inputString As String, ByVal leftCharsToRemove As Long, ByVal rightCharsToRemove As Long) As String
    Dim totalCharsToRemove As Long
    totalCharsToRemove = leftCharsToRemove + rightCharsToRemove
    
    If Len(inputString) <= totalCharsToRemove Then
        ' If the total characters to remove are more than the length of the string, return an empty string
        RemoveChars = ""
    Else
        ' Remove the left characters
        inputString = Mid(inputString, leftCharsToRemove + 1)
        
        ' Remove the right characters
        inputString = left(inputString, Len(inputString) - rightCharsToRemove)
        
        RemoveChars = inputString
    End If
End Function

' Helper function to check if a form exists
Public Function FormExists(FormName) As Boolean
    On Error Resume Next
    Dim frm As Form
    DoCmd.OpenForm FormName, , , , , acHidden
    Set frm = Forms(FormName)
    DoCmd.Close acForm, FormName, acSaveNo
    FormExists = (Err.number = 0)
    On Error GoTo 0
End Function

Public Function ReportExists(reportName) As Boolean
    On Error Resume Next
    Dim rpt As Report
    DoCmd.OpenReport reportName, acViewDesign, , , , acHidden
    Set rpt = Reports(reportName)
    DoCmd.Close acReport, reportName, acSaveNo
    ReportExists = (Err.number = 0)
    On Error GoTo 0
End Function

Public Function GetEndOfMonth(monthID, yearValue)

    If isFalse(monthID) Or isFalse(yearValue) Then
        GetEndOfMonth = Null
        Exit Function
    End If
    ' Check if the monthID is within a valid range (1 to 12)
    If monthID < 1 Or monthID > 12 Then
        GetEndOfMonth = Null
        Exit Function
    End If

    ' Construct the first day of the next month
    Dim firstDayNextMonth As Date
    firstDayNextMonth = DateSerial(yearValue, monthID + 1, 1)

    ' Subtract one day to get the last day of the specified month
    GetEndOfMonth = firstDayNextMonth - 1
End Function


Public Function ReturnResultArray(ResultMessage As String, Optional result As String = "Error") As Variant

    Dim Results(1) As String
    Results(0) = result
    Results(1) = ResultMessage
    
    ReturnResultArray = Results
    
End Function

Public Function IfError(Value) As Double
On Error GoTo ErrorHandler:
    IfError = IIf(IsError(Value), 0, Value)
    Exit Function
ErrorHandler:
    IfError = 0
End Function

Public Function EnumarateFields(tblName As String)

    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset(tblName)
        
    Dim fld As Field
    
    DoCmd.SetWarnings False
    For Each fld In rs.fields
        
        CurrentDb.Execute "INSERT INTO tblTableFields (TableName,FieldName) VALUES ('" & tblName & "','" & fld.Name & "')"
    
    Next fld
    DoCmd.SetWarnings True
    
End Function

Public Function Esc(str) As String
    Esc = EscapeString(str)
End Function

Public Sub p(str As String, Optional contains = "")

    PrintFields str, contains
    
End Sub

Sub CopyToClipboard(str)
    'Create a new DataObject to store the string
    Dim clipboard As Object
    Set clipboard = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")

    'Set the string to the DataObject's text property
    clipboard.SetText str
    On Error Resume Next
    'Copy the DataObject to the clipboard
    clipboard.PutInClipboard

    ''MsgBox "Copied to clipboard."
End Sub

Public Function PrintFields(tblName As String, Optional contains = "")
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset(tblName)
        
    Dim fld As Field
    Dim fields As New clsArray
    
    For Each fld In rs.fields
        If isFalse(contains) Then
            fields.Add fld.Name
        Else
            If fld.Name Like "*" & contains & "*" Then
                fields.Add fld.Name
            End If
        End If
    Next fld
    
    Dim result As String
    result = "TABLE: " & tblName & " Fields: " & fields.JoinArr(" | ")
    
    If fields.Count = 1 Then
        D fields.arr(0)
    End If
    
    Dim splitResult As New clsArray
    Dim currentIndex As Integer
    currentIndex = 1

    While currentIndex < Len(result)
        If Len(result) - currentIndex > 100 Then
            Dim nextSpace As Integer
            nextSpace = InStr(currentIndex + 100, result, "|")
            If nextSpace > 0 Then
                splitResult.Add Mid(result, currentIndex, nextSpace - currentIndex)
                currentIndex = nextSpace + 1
            Else
                splitResult.Add Mid(result, currentIndex, 100)
                currentIndex = currentIndex + 100
            End If
        Else
            splitResult.Add Mid(result, currentIndex, Len(result) - currentIndex + 1)
            currentIndex = Len(result) + 1
        End If
    Wend

    Dim finalResult As String, item
    finalResult = ""
    For Each item In splitResult.arr
        finalResult = finalResult & "''" & item & vbCrLf
    Next item
    
    Debug.Print finalResult
    
    ''CopyToClipboard finalResult
End Function

Public Function PrintRecordsetFields(rs As Recordset)

        
    Dim fld As Field
    Dim fields() As String
    Dim i As Integer
    
    For Each fld In rs.fields
        ReDim Preserve fields(i)
        fields(i) = fld.Name
        i = i + 1
    Next fld
    
    Debug.Print Join(fields, "|")
    
End Function

Public Function Divide(Numerator, Denominator) As Double
    
    If IsNull(Numerator) Or IsNull(Denominator) Then
        Divide = 0
        Exit Function
    End If
    
    If Numerator = 0 Or Denominator = 0 Then
        Divide = 0
        Exit Function
    End If
    
    Divide = Numerator / Denominator
    
End Function


Public Function ArrayLength(arr As Variant) As Integer
    'DEVELOPER: Ryan Wells (wellsr.com)
    'DESCRIPTION: Function to check if a value is in an array of values
    'INPUT: Pass the function a value to search for and an array of values of any data type.
    'OUTPUT: True if is in array, false otherwise
On Error GoTo ArrayLengthError: 'array is empty
        ArrayLength = UBound(arr) + 1
        Exit Function
ArrayLengthError:
    On Error GoTo 0
    ArrayLength = 0
End Function

Public Function LogIn(Optional ctl As IRibbonControl)
    g_UserID = 1
End Function

Public Function ShowError(ErrorStr As String)
    MsgBox ErrorStr, vbCritical + vbOKOnly
End Function

Public Function isFalse(Value) As Boolean
'On Error GoTo Err_isFalse
    
    If IsMissing(Value) Then
        isFalse = True
    Else
        isFalse = Value = "" Or IsNull(Value) Or IsEmpty(Value)
    End If

'Exit_isFalse:
'    Exit Function
'Err_isFalse:
'    LogError Err.number, Err.Description, "isFalse"
'    Resume Exit_isFalse
End Function

Public Function CdblNZ(val As Variant) As Double
    CdblNZ = CDbl(Nz(val, 0))
End Function

Public Function SetQueryDef()
    'Set recordsource -> tblAccOutInteractions
    ''qryLogs
    Dim logSQL As String
    logSQL = "SELECT * FROM qryLogs WHERE TableName = ""tblAccOutInteractions"" And EventName = ""ADD"""
    
    Dim sqlStr As String
    sqlStr = "SELECT tblAccOutInteractions.*,UserName,DateTime FROM tblAccOutInteractions LEFT JOIN (" & logSQL & ") As qryLogs ON tblAccOutInteractions.AccOutInteractionID = qryLogs.RecordID"
    
    Dim qDef As QueryDef
    Set qDef = CurrentDb.QueryDefs("qryAccOutInteractions")
    qDef.sql = sqlStr
End Function

Public Function PromptLogin()
    
    If isFalse(g_UserID) Then
        If MsgBox("Login using developer user?", vbYesNo) = vbYes Then
            LogIn
        End If
    End If
    
End Function


Public Function ReturnYesNo(TF As Boolean) As String
    If TF Then
        ReturnYesNo = "YES"
    Else
        ReturnYesNo = "NO"
    End If
End Function


Public Function GenerateUPCID() As String

    Dim randNumber As Long, upperbound, lowerbound, UPCStr As String
    upperbound = 9999999: lowerbound = 1
    
    Do Until UPCStr <> "" Or isPresent("tblOrders", "PackCycleID = '" & UPCStr & "'")
        randNumber = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
        UPCStr = Format$(randNumber, "0000000")
    Loop
    
    GenerateUPCID = UPCStr
    
End Function

Public Function ExitIfTrue(Condition As Boolean, Msg As String) As Boolean

    ExitIfTrue = False
    If Condition Then
        ShowError Msg
        ExitIfTrue = True
    End If
    
End Function
Public Function GenerateFieldNamesString(ByVal tblName As String, Optional ByVal IgnoreFields = "", Optional prefix As Variant = "") As String

    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset(tblName)
        
    Dim fld As Field
    Dim fields() As String, IgnoreArray() As String
    Dim i As Integer
    
    IgnoreArray = Split(IgnoreFields, ",")
    
    For Each fld In rs.fields
        If Not IsInArray(fld.Name, IgnoreArray) Then
            ReDim Preserve fields(i)
            If prefix <> "" Then
                fields(i) = prefix & "." & fld.Name
            Else
                fields(i) = fld.Name
            End If
            i = i + 1
        End If
    Next fld
    
    GenerateFieldNamesString = Join(fields, ", ")
    
End Function

Public Function GetTotalRecordCount(ByVal rs As Recordset) As Long

If rs.EOF Then
    GetTotalRecordCount = 0
    Exit Function
Else
    rs.MoveLast
    GetTotalRecordCount = rs.recordCount
    rs.MoveFirst
End If

End Function

Public Function InchToTwip(inch) As Double
    InchToTwip = inch * 1440
End Function

Public Function deleteTableIfExists(ByVal tableName As String) As Boolean

On Error GoTo Err_Handler:
    'Runs through all table names in CurrentDB and deletes table if name matches
    Dim db As DAO.Database
    Dim td As TableDef
    
    Set db = CurrentDb
    
    For Each td In db.TableDefs
        If td.Name = tableName Then
            db.TableDefs.Delete tableName
            db.TableDefs.refresh
            Set td = Nothing
            Set db = Nothing
            Exit For
        End If
    Next td
        
    Set td = Nothing
    Set db = Nothing
    
    deleteTableIfExists = True
    Exit Function
    
Err_Handler:
    
    If Err.number = 3211 Then
        ShowError """" & tableName & """ is open. Please close it first.."
        Exit Function
    End If

End Function

Public Sub deleteQueryIfExists(ByVal QueryName As String)
'Runs through all table names in CurrentDB and deletes table if name matches
    Dim db As DAO.Database
    Dim qdf As QueryDef

    Set db = CurrentDb

    For Each qdf In db.QueryDefs
        If qdf.Name = QueryName Then
            db.QueryDefs.Delete QueryName
            db.QueryDefs.refresh
            Set qdf = Nothing
            Set db = Nothing
            Exit For
        End If
    Next qdf
    
    Set qdf = Nothing
    Set db = Nothing
    
End Sub

Public Function DoesObjectExists(obj As Object)

    ''Checks wether a property exists within the parent obj
    ''obj => Parent object | propertyName => Name of the property
    Dim tempObj
On Error Resume Next
    Set tempObj = obj
    DoesObjectExists = (Err = 0)
On Error GoTo 0

End Function


Public Function DoesPropertyExists(obj As Object, PropertyName)
    ''Checks wether a property exists within the parent obj
    ''obj => Parent object | propertyName => Name of the property
    Dim tempObj
On Error Resume Next
    Set tempObj = obj(PropertyName)
    DoesPropertyExists = (Err = 0)
On Error GoTo 0

End Function

Function AddSpaces(pValue) As String
    'Update 20140723
    Dim xOut As String
    Dim i, xAsc
    xOut = VBA.left(pValue, 1)
    For i = 2 To VBA.Len(pValue)
       xAsc = VBA.Asc(VBA.Mid(pValue, i, 1))
       If xAsc >= 65 And xAsc <= 90 Then
          xOut = xOut & " " & VBA.Mid(pValue, i, 1)
       Else
          xOut = xOut & VBA.Mid(pValue, i, 1)
       End If
    Next
    AddSpaces = xOut
End Function

Function RemoveSpaces(pValue) As String
    'Update 20140723
    RemoveSpaces = replace(pValue, " ", "")
End Function

Public Function ReturnStringBasedOnType(fieldVal As Variant, ControlType As Integer) As String

    If IsNull(fieldVal) Then
        ReturnStringBasedOnType = "Null"
        Exit Function
    End If
    
    Select Case ControlType
        Case 10, 12:
            ReturnStringBasedOnType = """" & fieldVal & """"
        Case 8:
            ReturnStringBasedOnType = "#" & SQLDate(fieldVal) & "#"
        Case Else:
            ReturnStringBasedOnType = fieldVal
    End Select
    
End Function

Public Function myRandVal(nm As String) As String
' Randomizes Text Strings and/or numbers
' Usage: In a query - NameOfFieldRnd: myRandVal([NameOfField])
'        In a Form Control's "ControlSource" property - =myRandVal([NameOfField])
' Test in debug window - ?myRandVal("12345 UPPER & lower Case")
' Modified by Bob Raskew [URL="tel:11/1/2003"]11/1/2003[/URL] - Added number scrambling
Dim myChr As String
Dim myAsc As Integer
Dim i As Integer
    myRandVal = ""
    If Len(nm) = 0 Then
        myRandVal = vbNullString
        Exit Function
    Else
        For i = 1 To Len(nm)
            myChr = Mid(nm, i, 1)
            If Asc(myChr) >= 65 And Asc(myChr) <= 90 Then
                myAsc = Int((90 - 65 + 1) * Rnd + 65)
              ElseIf Asc(myChr) >= 97 And Asc(myChr) <= 122 Then
                myAsc = Int((122 - 97 + 1) * Rnd + 97)
              ElseIf Asc(myChr) >= 48 And Asc(myChr) <= 57 Then
                myAsc = Int((57 - 48 + 1) * Rnd + 48)
              Else
                myAsc = Asc(myChr)
            End If
            myChr = Chr(myAsc)
            myRandVal = myRandVal & myChr
        Next i
    End If
End Function

Public Function myRandEmail(nm As String) As String
' Randomizes Text Strings and/or numbers
' Usage: In a query - NameOfFieldRnd: myRandVal([NameOfField])
'        In a Form Control's "ControlSource" property - =myRandVal([NameOfField])
' Test in debug window - ?myRandVal("12345 UPPER & lower Case")
' Modified by Bob Raskew [URL="tel:11/1/2003"]11/1/2003[/URL] - Added number scrambling
    Dim splStr() As String

    myRandEmail = ""
    If Len(nm) = 0 Then
        myRandEmail = vbNullString
        Exit Function
    ElseIf InStr(1, nm, "@") Then
        splStr = Split(nm, "@")
        splStr(0) = left(CStr(10000000000# * Rnd), 10)
        splStr(1) = "example.com"
        myRandEmail = Join(splStr, "@")
    End If
End Function

Public Function makeQuery(sqlStr)

    Dim db As DAO.Database
    Dim qDef As DAO.QueryDef
    
    Set db = CurrentDb
    
    If DoesPropertyExists(db.QueryDefs, "qryTestQuery") Then
        Set qDef = db.QueryDefs("qryTestQuery")
    Else
        Set qDef = db.CreateQueryDef("qryTestQuery")
    End If

    qDef.sql = sqlStr
    qDef.Close
    db.Close
    
    DoCmd.OpenQuery "qryTestQuery", acViewDesign
    
End Function

Public Function GenerateUpdateStatements(targetFieldNames, targetTableName, updateFrom) As Variant

    Dim updatearr As New clsArray, filterArr As New clsArray, targetFieldArr As New clsArray, returnArr As New clsArray
    Dim targetField As Variant, trimmedtargetField As String, origField As String, tempField As String
    
    targetFieldArr.arr = targetFieldNames
    
    For Each targetField In targetFieldArr.arr
        trimmedtargetField = Trim(targetField)
        origField = targetTableName & "." & trimmedtargetField
        tempField = "[" & updateFrom & "]![" & trimmedtargetField & "]"
        
        updatearr.Add origField & " = " & tempField
        filterArr.Add origField & " <> " & tempField
        
    Next targetField
    
    Dim filterStatement As String, updateStatement As String
    updateStatement = updatearr.JoinArr
    filterStatement = filterArr.JoinArr(" OR ")
    
    returnArr.Add updateStatement
    returnArr.Add filterStatement
    
    GenerateUpdateStatements = returnArr.arr
    
End Function

Public Function HasProperty(obj As Object, strPropName) As Boolean
    'Purpose:   Return true if the object has the property.
    Dim varDummy As Variant
    
    On Error Resume Next
    varDummy = obj.Properties(strPropName)
    HasProperty = (Err.number = 0)
End Function


Public Function GenerateJoinObj(Source, LeftFields, Optional Alias = "", Optional RightFields = "", Optional JoinType = "") As clsJoin

    Dim joinObj As clsJoin
    Set joinObj = New clsJoin
    With joinObj
      .Source = Source
      .LeftFields = LeftFields
    End With
    
    If Alias <> "" Then joinObj.Alias = Alias
    If RightFields <> "" Then joinObj.RightFields = RightFields
    If JoinType <> "" Then joinObj.JoinType = JoinType
    
    Set GenerateJoinObj = joinObj
  
End Function

Public Function concat(ParamArray var() As Variant) As String
    Dim i As Integer
    Dim tmp As String
    For i = LBound(var) To UBound(var)
        tmp = tmp & var(i)
    Next
    concat = tmp
End Function

Public Function CenterString(xInput As String, xLength As Long)
    Dim xM As Variant
    xM = Space(((xLength / 2) - (Len(xInput) / 2) + 1)) + xInput
    CenterString = xM + Space(xLength - Len(xM))
End Function
