Attribute VB_Name = "Module1"
Sub Copier()
ThisWorkbook.ActiveSheet.Copy After:=Sheets(Sheets.Count)
End Sub
