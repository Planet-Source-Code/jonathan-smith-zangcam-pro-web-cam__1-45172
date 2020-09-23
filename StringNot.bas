Attribute VB_Name = "modStringNot"
Option Explicit

Public Function StringNot(ByVal szData As String) As String

    '// Very simple, not too secure data protection scheme
    Dim i As Long
    Dim szEnc As String
    
    For i = 1 To Len(szData)
        szEnc = szEnc & Chr$(Abs(255 - Asc(Mid$(szData, i, 1))))
    Next i
    
    StringNot = szEnc
    
End Function
