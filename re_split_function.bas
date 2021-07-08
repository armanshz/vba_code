Attribute VB_Name = "Module1"
Function re_split(str_to_split As String, pattern As String) As String
'Custom function to split a text string based on regex pattern
'Declaring regex object.
'Requires Tools > References > Microsoft VBScript Regular Expressions 5.5 to be checked
Dim re As Object
Set re = New RegExp
re.pattern = pattern
re.Global = True
' regex matches are stored in "matches" variable
Set matches = re.Execute(str_to_split)
'Initialise result variable to empty string
result = ""
'Concatenate result string with each Match in matches
For Each Match In matches
result = result + Match + ","
Next
'Remove trailing "," and set result = re_split which is displayed
re_split = Left(result, Len(result) - 1)
End Function

