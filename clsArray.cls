VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private vArray() As String

Property Get arr() As Variant

    arr = vArray
    
End Property

Property Let arr(arrValue As Variant)
    
    If IsArray(arrValue) Then
        vArray = arrValue
    Else
        vArray = Split(arrValue, ",")
    End If
    
End Property

Public Function Add(arrValue As Variant, Optional Unique As Boolean = False)
    
    
    If Unique Then
        If IsInArray(arrValue, vArray) Then Exit Function
    End If
    
    Dim CurrentCount As Long
    CurrentCount = Count
    
    ReDim Preserve vArray(CurrentCount)
    vArray(CurrentCount) = arrValue

End Function

Public Function Count() As Long

    On Error GoTo ErrHandler:
    Count = UBound(vArray) + 1
    Exit Function
ErrHandler:
    Select Case Err.number
    Case 9:
        Count = 0
    End Select
    
End Function

Public Function JoinArr(Optional Delimiter As String = ",") As String
    JoinArr = Join(vArray, Delimiter)
End Function

Public Function items(index As Integer) As String
    items = vArray(index)
End Function

Public Function InArray(Value) As Boolean
    InArray = IsInArray(Value, vArray)
End Function

Public Function clearArr()

    Erase vArray
    
End Function

