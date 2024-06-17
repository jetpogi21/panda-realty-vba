Attribute VB_Name = "Row Functions"
Option Compare Database
Option Explicit

Public Function GetTotalRetailValue(PurchaseOrderID As Variant) As Double
    If isFalse(PurchaseOrderID) Then
        GetTotalRetailValue = 0
        Exit Function
    End If
    
    Dim rs As Recordset
    
End Function

Public Function GetTotalPOSupplierCost(PurchaseOrderID As Variant) As Double
    If isFalse(PurchaseOrderID) Then
        GetTotalPOSupplierCost = 0
        Exit Function
    End If
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT SUM(TotalCost) As SumOfTotalCost from qryPurchaseOrderProducts where PurchaseOrderID = " & PurchaseOrderID & " And Active = -1")
    
    If rs.EOF Then
        GetTotalPOSupplierCost = 0
        Exit Function
    End If
    
    If isFalse(rs.fields("SumOfTotalCost")) Then
        GetTotalPOSupplierCost = 0
        Exit Function
    End If
    
    GetTotalPOSupplierCost = rs.fields("SumOfTotalCost")
    
End Function


Public Function GetTotalLinkedBackorderQTY(PurchaseOrderProductID As Variant) As Double
    If isFalse(PurchaseOrderProductID) Then
        GetTotalLinkedBackorderQTY = 0
        Exit Function
    End If
    
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT SUM(LinkedBackorderQTY) As SumOfLinkedBackorderQTY from tblProductBackorderLinks where PurchaseOrderProductID = " & _
        PurchaseOrderProductID & " And Active = -1")
    
    If rs.EOF Then
        GetTotalLinkedBackorderQTY = 0
        Exit Function
    End If
    
    If isFalse(rs.fields("SumOfLinkedBackorderQTY")) Then
        GetTotalLinkedBackorderQTY = 0
        Exit Function
    End If
    
    GetTotalLinkedBackorderQTY = rs.fields("SumOfLinkedBackorderQTY")
End Function


Public Function GetBackorderLinkCode(PurchaseOrderID, EstimatedDeliveryDate, LinkedBackorderQTY, Status) As String
    If isFalse(PurchaseOrderID) Then
        GetBackorderLinkCode = ""
        Exit Function
    End If
    
    If isFalse(EstimatedDeliveryDate) Then
        GetBackorderLinkCode = left(Status, 4) & "- PO:" & PurchaseOrderID & " -NODATE- " & " (" & LinkedBackorderQTY & ")"
    Else
        GetBackorderLinkCode = left(Status, 4) & "- PO:" & PurchaseOrderID & " ON " & Format(EstimatedDeliveryDate, "DD/MM") & " (" & LinkedBackorderQTY & ")"
    End If
    
End Function

Public Function ListOrderNumbers(ProjectionRecordID As Variant) As String
    If isFalse(ProjectionRecordID) Then
        ListOrderNumbers = ""
        Exit Function
    End If
    
    Dim OrderNumbers() As String
    Dim i As Integer
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT ProjectionRecordID,OrderID,OrderNumber FROM qryProjectionRecordBOs WHERE ProjectionRecordID = " & ProjectionRecordID & _
        " GROUP BY ProjectionRecordID,OrderID,OrderNumber")
    
    If rs.EOF Then
        ListOrderNumbers = ""
    Else
        Do Until rs.EOF
            ReDim Preserve OrderNumbers(i)
            OrderNumbers(i) = rs.fields("OrderNumber")
            i = i + 1
            rs.MoveNext
        Loop
        
        ListOrderNumbers = Join(OrderNumbers, ",")
    End If
    
End Function


Public Function ListBackordersLinked(PurchaseOrderProductID As Variant) As String
    ''Backorders - If a backorder link exists on this Purchase Order, check the order is Processing, display Order here in the following format:
    ''OrderNumber -OrderLocation(LinkedBackorderQTY)
    ''That way the warehouse staff can find the order and knows how many that specific order requires.
    
    ''qryProductBackorderLinks
    ''PurchaseOrderProductID | OrderNumber, Location, LinkedBackorderQTY, Status = 'Processing'
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM qryProductBackorderLinks WHERE PurchaseOrderProductID = " & PurchaseOrderProductID & _
        " And OrderStatus = ""Processing""")
        
    Dim BOLinked() As String
    Dim i As Integer
    Do Until rs.EOF
        ReDim Preserve BOLinked(i)
        BOLinked(i) = rs.fields("OrderNumber") & "-" & rs.fields("Location") & "(" & rs.fields("LinkedBackorderQTY") & ")"
        i = i + 1
        rs.MoveNext
    Loop
    
    ListBackordersLinked = Join(BOLinked, vbCrLf)
    
End Function

Public Function LatestStdComment(OrderID As Variant) As String

    If isFalse(OrderID) Then
        LatestStdComment = ""
        Exit Function
    End If
    
    ''Latest standard comment "[Date] [Show order comment where exception is false]
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM qryOrderComments WHERE OrderID = " & OrderID & " And Exception = 0 ORDER BY DateTime DESC")
    
    If rs.EOF Then
        LatestStdComment = vbNullString
    Else
        LatestStdComment = "Latest standard comment """ & Format$(rs.fields("DateTime"), "DD/MM/YYYY") & " " & rs.fields("Comment") & """"
    End If
    
End Function

Public Function LatestExceptionComment(OrderID As Variant) As String

    If isFalse(OrderID) Then
        LatestExceptionComment = ""
        Exit Function
    End If
    
    ''Latest standard comment "[Date] [Show order comment where exception is true]
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM qryOrderComments WHERE OrderID = " & OrderID & " And Exception = -1 ORDER BY DateTime DESC")
    
    If rs.EOF Then
        LatestExceptionComment = vbNullString
    Else
        
        LatestExceptionComment = "Latest exception comment """ & Format$(rs.fields("DateTime"), "DD/MM/YYYY") & " " & rs.fields("Comment") & """"
        
            If rs.fields("ExceptionCleared") = -1 Then
                LatestExceptionComment = LatestExceptionComment & vbCrLf & "Resolution: " & rs.fields("ExceptionResolution") & " at " & Format$(rs.fields("ExceptionClearedDateTime"), "hh:mm DD/MM/YYYY")
            End If
    End If
    
End Function

Public Function LatestOutageComment(OrderID As Variant) As String

    If isFalse(OrderID) Then
        LatestOutageComment = ""
        Exit Function
    End If
    
    ''Latest standard comment "[Date] [Show order comment where exception is false]
    Dim rs As Recordset
    Set rs = ReturnRecordset("SELECT * FROM qryAccOuts WHERE OrderID = " & OrderID & " ORDER BY UpdatedTimestamp DESC")
    
    If rs.EOF Then
        LatestOutageComment = vbNullString
    Else
        LatestOutageComment = "Latest Accidental Outage comment " & vbCrLf & " """ & Format$(rs.fields("UpdatedTimestamp"), "hh:ss DD/MM/YYYY") & " " & rs.fields("LastComment") & """"
    End If
    
End Function

Public Function ScheduledRunTime(PresetInterval As Variant, ScheduledTime As Variant, MinuteInterval As Variant) As String

    If isFalse(PresetInterval) Then
        ''If PresetInterval is Null then this is a custom duration interval
        ScheduledRunTime = "Every [" & MinuteInterval & "] Minutes"
        Exit Function
    Else
        ScheduledRunTime = PresetInterval & " @ " & ScheduledTime
        Exit Function
    End If
    
End Function

Public Function NextRunTime(LastRunAt As Variant, PresetInterval As Variant, ScheduledTime As Variant, IsSuspended As Boolean, MinuteInterval As Variant) As Variant
    
    If IsSuspended Then
        NextRunTime = Null
        Exit Function
    End If
    
    If isFalse(LastRunAt) Then
        ''The task is not yet run ever so run on specified scheduled date or now
        If isFalse(PresetInterval) Then
            ''If PresetInterval is Null then this is a custom duration interval
            NextRunTime = Now()
            Exit Function
        Else
            If time() > ScheduledTime Then
                ''The time is greather than the scheduled time then this is runable
                NextRunTime = Now()
                Exit Function
            Else
                NextRunTime = Null
                Exit Function
            End If
        End If
        
    End If
    
    If isFalse(PresetInterval) Then
        ''If PresetInterval is Null then this is a custom duration interval
        ''Compute using the Last Runat
        NextRunTime = DateAdd("n", MinuteInterval, LastRunAt)
        Exit Function
    Else
        ''Check if Task has been run based on the PresetInterval
        Dim lastRunYear As Long, currentYear As Long
        Dim lastRunMonth, currentMonth
        Dim lastRunWeek As Integer, currentWeek As Integer
        Dim lastRunFort As Integer, currentFort As Integer
        lastRunYear = DatePart("yyyy", SQLDate(LastRunAt)): currentYear = DatePart("yyyy", SQLDate(Now()))
        lastRunMonth = DatePart("m", SQLDate(LastRunAt)): currentMonth = DatePart("m", SQLDate(Now()))
        lastRunWeek = DatePart("ww", SQLDate(LastRunAt)): currentWeek = DatePart("ww", SQLDate(Now()))
        lastRunFort = GetCurrentFortnight(LastRunAt): currentFort = GetCurrentFortnight(Now())
        Select Case PresetInterval
            Case "Daily":
                If DateValue((LastRunAt)) <= Date Then
                    ''This means the task was already run today so proceed to next day with scheduled time
                    NextRunTime = DateAdd("d", 1, DateValue((LastRunAt))) + TimeValue(ScheduledTime)
                    Exit Function
                End If
            Case "Weekly":
                If lastRunWeek <= currentWeek And lastRunYear <= currentYear Then
                    ''This means the task was already run this week so proceed to the first day of next week scheduled time
                    NextRunTime = GetFirstDayOfWeek(lastRunYear, lastRunWeek + 1) + TimeValue(ScheduledTime)
                    Exit Function
                End If
            Case "Fortnightly":
                If lastRunFort <= currentFort And lastRunYear <= currentYear Then
                    ''This means the task was already run this forntight so proceed to the first day of next next week scheduled time
                    NextRunTime = GetFirstDayOfWeek(lastRunYear, lastRunWeek + 2) + TimeValue(ScheduledTime)
                    Exit Function
                End If
            Case "Monthly":
                 If lastRunMonth <= currentMonth And lastRunYear <= currentYear Then
                    ''This means the task was already run this week so proceed to the first day of next week scheduled time
                    NextRunTime = GetFirstDayOfNextMonth(LastRunAt) + TimeValue(ScheduledTime)
                    Exit Function
                End If
        End Select
        
        NextRunTime = Null
        Exit Function
        
    End If
    
End Function

Public Function RoundUp(vValue As Double) As Long
    RoundUp = -Int(-vValue)
End Function

Public Function GetCurrentFortnight(vDate As Variant) As Integer
    Dim currentWeek As Integer
    currentWeek = DatePart("ww", vDate)
    GetCurrentFortnight = RoundUp(currentWeek / 2)
End Function
