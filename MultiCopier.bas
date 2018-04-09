Attribute VB_Name = "Module1"
Sub MultiCopy()
Dim CopyToNum As Integer
Dim I As Integer
CopyToNum = InputBox("How many copies would you like", "Excel Multi Copy", 1)

Dim SheetToCopy As Worksheet

Set SheetToCopy = ActiveSheet

For I = 1 To CopyToNum
     SheetToCopy.Copy After:=Sheets(Sheets.Count)
Next I
End Sub
