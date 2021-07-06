Attribute VB_Name = "Module1"
Function getnum(text As String) As String
With CreateObject("VBScript.Regexp")
.Global = True
.Pattern = "[A-Z a-z]"
getnum = .Replace(text, "")
End With
End Function
