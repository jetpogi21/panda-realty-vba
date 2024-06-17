Attribute VB_Name = "ImportDataField Mod"
Option Compare Database
Option Explicit

Public Function ImportDataFieldModelFieldIDAfterUpdate(frm As Form)
    
    Dim ModelFieldID
    ModelFieldID = frm("ModelFieldID")
    
    If IsNull(ModelFieldID) Then Exit Function
    
    Dim PrimaryKey
    PrimaryKey = frm("ModelFieldID").Column(1)
    If Not PrimaryKey Like "*ID" Then Exit Function
    
    ''LookupTable, LookupField, ReturnField
    Dim rs As Recordset
    Set rs = GetModelByPrimaryKey(PrimaryKey)
    
    frm("ReturnField") = PrimaryKey
    frm("LookupTable") = GetTableNameByPrimaryKey(PrimaryKey)
    frm("LookupField") = rs.fields("MainField")
    
    
End Function
