Attribute VB_Name = "Module1"
Sub QuickPDFExport()
Dim File As String
File = ActiveWorkbook.FullName
ActiveWorkbook.Sheets.Select
ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=File, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
End Sub
