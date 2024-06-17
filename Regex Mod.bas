Attribute VB_Name = "Regex Mod"
Option Compare Database
Option Explicit

Public Function RegexCreate(frm As Form, FormTypeID)

    Select Case FormTypeID
        Case 4: ''Data Entry Form
        Case 5: ''Datasheet Form
        Case 6: ''Main Form
        Case 7: ''Tabular Report
    End Select

End Function

'Public Function GetPropertyID(toMatch) As String
'
'    ''https://rpp.rpdata.com/rpp/property/detail.html?propertyId=16118042
'    Dim pattern: pattern = ".+propertyID=(\d+)"
'    ''get the ID=
'
'    Dim regex As Object
'    Dim theMatches As Object
'    Dim Match As Object
'    Set regex = New RegExp
'
'    regex.pattern = pattern
'    regex.Global = False
'    regex.IgnoreCase = True
'
'    If regex.Test(toMatch) Then
'        TestRegex = regex.Replace(toMatch, "$1")
'    End If
'
'End Function

''GetSanitizedName("Sai - Enquired 50 Pechey Street Chermside")
Public Function GetSanitizedName(toMatch) As String
    
    GetSanitizedName = toMatch
    Dim pattern: pattern = "^([a-zA-Z0-9 ]+)[ -_]+Enq.*"
    
    Dim regex As Object
    Dim theMatches As Object
    Dim match As Object
    Set regex = New RegExp
    
    regex.pattern = pattern
    regex.Global = False
    regex.IgnoreCase = True

    If regex.Test(toMatch) Then
        GetSanitizedName = regex.replace(toMatch, "$1")
    End If
    

End Function
