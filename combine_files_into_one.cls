VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub combine()
Application.ScreenUpdating = False
Dim folderPath, filePath As String, i As Integer
folderPath = Left(ActiveWorkbook.FullName, Len(ActiveWorkbook.FullName) - Len(ActiveWorkbook.Name))
fileName = Dir(folderPath & "*.xlsx")


Do While fileName <> ""
filePath = folderPath & fileName
ActiveSheet.Range("A10000").End(xlUp).Offset(1, 0).Value = filePath
Workbooks.Open (filePath)
ActiveSheet.Range("A1").CurrentRegion.Copy Workbooks("main").ActiveSheet.Range("A10000").End(xlUp).Offset(1, 0)
Application.CutCopyMode = False
Workbooks(fileName).Close SaveChanges:=False
fileName = Dir()
Loop
Application.ScreenUpdating = True
End Sub
