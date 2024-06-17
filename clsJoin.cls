VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private vSource As String
Private vAlias As String
Private vJoinType As String
Private vRightFields As New clsArray
Private vLeftFields As New clsArray

Private Sub Class_Initialize()
    vJoinType = "INNER"
End Sub

Property Let Source(Value As String)
    vSource = Value
End Property

Property Let Alias(Value As String)
    vAlias = Value
End Property

Property Let JoinType(Value As String)
    vJoinType = Value
End Property

Property Let RightFields(Value As Variant)
    vRightFields.arr = Value
End Property

Property Let LeftFields(Value As Variant)
    vRightFields.arr = Value
    vLeftFields.arr = Value
End Property

Property Get Source() As String
    Source = vSource
End Property

Property Get Alias() As String
    Alias = vAlias
End Property

Property Get JoinType() As String
    JoinType = vJoinType
End Property

Property Get RightFields() As clsArray
    Set RightFields = vRightFields
End Property

Property Get LeftFields() As clsArray
    Set LeftFields = vLeftFields
End Property

