'Author: Angelo Sanches
'Origin date: 05/22/2018
'Reason for Code: is to quickly tally the totals for results in check workbook quickly
'Change Log
'{
'Editor: Angelo Sanches
'Change date: 05/22/2018
'Reason for Code Change: Pulled the sub(Method or function) out into its own space and to let it be called be the required counter as per page
' then I added a sub calling it for each page type. It also has been edited to allow for ignores or false positive handling.
',
'Editor: Angelo Sanches
'Change date: 05/22/2018
'Reason for Code Change: PDF export emoves file ending now
',
'Editor: Angelo Sanches
'Change date: 05/30/2018
'Reason for Code Change: Reordered the print out of counts to more closely match the share point
',
'Editor: Angelo Sanches
'Change date: 05/30/2018
'Reason for Code Change: Added AIRB-Waive to waved options to case cover
',
'Editor: Angelo Sanches
'Change date: 06/22/2018
'Reason for Code Change: Added Error handeling
',
'Editor: Angelo Sanches
'Change date: 06/22/2018
'Reason for Code Change: To fix the error related to hidden pages for the PDF print
',
'Editor: Angelo Sanches
'Change date: 06/27/2018
'Reason for Code Change: To adhere to a rescope on adoption. I have: Depreacated Copier, QuickCount, QuickCountINF,
'QuickCountIPRE; Imple. Quick AutoStats (this implents with new CellVector, PageResults objects and new AutoQuickCount, FindValueInRange.)
',
'Editor: Angelo Sanches
'Change date: 06/27/2018
'Reason for Code Change: Added a are you sure ok box to file overwrites on PDF.
',
'Editor: Angelo Sanches
'Change date: 06/29/2018
'Reason for Code Change: Changed the read out to match suggested readout with added fields.
',
'Editor: Angelo Sanches
'Change date: 06/29/2018
'Reason for Code Change: changed the search method to allow for multiple
'values to be looked for and added a config file that can be auto
'generated from user input. Also Removed and achived old depr. code.
'}
Option Explicit

Enum InputReturnAnswer
 None = 0
 Error = -1
 Ok = 1
 Cancel = 2
End Enum

Public Type ConfigReturn
    CustomerNameSearchHeaders As String
    RequestedBySearchHeaders As String
    FRNumberSearchHeaders As String
    DateSearchHeaders As String
    BranchNumberSearchHeaders As String
    DCASearchHeaders As String
    CommentSearchHeaders As String
    ResultsSearchHeaders As String
End Type

Type EasyConfigReturn
    Config_ As ConfigReturn
    SaveLocation As String
    Answer As InputReturnAnswer
End Type

Public Type CellVector
    X As Integer
    Y As Integer
End Type

Public Type PageResults
    SheetWasCounted As Boolean
    PassCount As Integer
    FailCount As Integer
    WAVECount As Integer
    NACount As Integer
    UnselectedCount As Integer
    TotalCount As Integer
    Unfilled As Integer
    Error As String
End Type

Public Type CountResults
    PassCount As Integer
    FailCount As Integer
    WAVECount As Integer
    NACount As Integer
    UnselectedCount As Integer
    TotalCount As Integer
    Unfilled As Integer
End Type

Public Type FindValuesResult
    Location As CellVector
    Success As Boolean
    ValueLocated As String
    FoundWith As String
End Type

Public Type config
    CustomerNameSearchHeaders() As String
    RequestedBySearchHeaders() As String
    FRNumberSearchHeaders() As String
    DateSearchHeaders() As String
    BranchNumberSearchHeaders() As String
    DCASearchHeaders() As String
    CommentSearchHeaders() As String
    ResultsSearchHeaders() As String
End Type

Public Type configAsString
    CustomerNameSearchHeaders As String
    RequestedBySearchHeaders As String
    FRNumberSearchHeaders As String
    DateSearchHeaders As String
    BranchNumberSearchHeaders As String
    DCASearchHeaders As String
    CommentSearchHeaders As String
    ResultsSearchHeaders As String
End Type


Public EasyConfigReturn As EasyConfigReturn
Public StrConfigReturn As configAsString


'Makes a export of the pdf with the same name in the same location
Sub QuickPDFExport()
    Dim CheckSheet As Integer
    Dim File As String
    Dim Selecter() As Integer
    Dim j As Integer
    j = 0
    
    
    'Grab the location & drop the file ending
    On Error GoTo ErrorHandler

    For CheckSheet = 1 To Sheets.Count
        'select a sheet.
        If Not (Sheets(CheckSheet).Visible = xlSheetHidden) Then
            ReDim Preserve Selecter(j)
            Selecter(j) = CheckSheet
            j = j + 1
        End If
    Next
    
    'select all the sheets
    Sheets(Selecter).Select
    File = ActiveWorkbook.FullName
    File = Mid(File, 1, Len(File) - 5)
    
    If Not (Len(Dir(File & ".pdf")) = 0) Then
        If MsgBox("This PDF already exists in the same place. Would you like to overwrite it?", vbYesNo) = vbNo Then Exit Sub
    End If
    
Retry:
    ' Export as a pdf
    ActiveSheet.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=File, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=True
    Exit Sub

ErrorHandler:

        If MsgBox("Would you like to try again from a specified file location?", vbYesNo) = vbYes Then
            Dim varResult As Variant
            varResult = Application.GetSaveAsFilename(FileFilter:="PDF(*.pdf), *.pdf", Title:="PDF Save", InitialFileName:=File)
            
            If varResult <> False Then
                File = varResult
                GoTo Retry
            Else
                Exit Sub
            End If
        Else
            Exit Sub
        End If

End Sub

' Makes multiple copies
Sub MultiCopier()
Dim CopyToNum As Double
Dim i As Integer

On Error GoTo InputErrorHandler
Retry:
CopyToNum = InputBox("How many copies would you like", "Excel Multi Copy", 1)
If (CopyToNum < 1) Then
    If MsgBox("The number entered is ethier a negative number or zero. It must be a positive whole number. Would you like to try again?", vbYesNo) = vbYes Then
        GoTo Retry
    Else
        Exit Sub
    End If
End If

If Not (Int(CopyToNum) = CopyToNum) Then
    If MsgBox("The number entered is decimal. It must be a positive whole number. Would you like to try again?", vbYesNo) = vbYes Then
        GoTo Retry
    Else
        Exit Sub
    End If
End If

On Error GoTo WorkErrorHandler
Dim SheetToCopy As Worksheet
Dim CountMade As Integer

Set SheetToCopy = ActiveSheet

For i = 1 To CopyToNum
     SheetToCopy.Copy After:=Sheets(Sheets.Count)
     CountMade = i
Next i
    Exit Sub
    
InputErrorHandler:
    If MsgBox("The Value entered is not a number. It must be a positive whole number. Would you like to try again?", vbYesNo) = vbYes Then
        GoTo Retry
    Else
        Exit Sub
    End If
WorkErrorHandler:
    'roll back
    MsgBox "Failed Copy all the Sheets. The name of the sheet was too long. Please shorten the name and try again. A roll back will be provided apon ok."
    Application.DisplayAlerts = False
    For i = 1 To CountMade
     Sheets(Sheets.Count).Delete
    Next i
    Application.DisplayAlerts = True
End Sub



Function FindValueInRange(LookFor As String, MyIn As Range, SNum As Integer) As CellVector
    Dim Brake As Boolean
    Dim Row, Col, i, j As Integer
    Dim ReturnVal As CellVector
    Row = -1
    Col = -1
    Brake = False
    LookFor = Trim(LookFor)
    For i = MyIn.Row To MyIn.Row + MyIn.Rows.Count - 1
        For j = MyIn.Column To MyIn.Column + MyIn.Columns.Count - 1
            If (Sheets(SNum).Cells(i, j).Value Like LookFor) Then
                Col = j
                Row = i
                Brake = True
                Exit For
            End If
        Next j
        If (Brake) Then
            Exit For
        End If
    Next i
    ReturnVal.X = Col
    ReturnVal.Y = Row
    FindValueInRange = ReturnVal
End Function

Function AutoQuickCount(SheetNum As Integer, StartRowResult As Integer, StartColResult As Integer, StartRowComment As Integer, StartColComment As Integer, PrintArea As Range) As PageResults
'*********************************** comparable string consts ***********************************************
'The comparable string for pass
Const PassResultStr = "Pass"
'The comparable string for Fail
Const FailResultStr = "Fail"
'The comparable string for Waved
Const WAVEResultStr = "Waived"
Const WAVEResultStr2 = "AIRB-Waive"
'The comparable string for N/A
Const NAResultStr = "N/A"
'The Note there is no comparable for an empty box and it is considered to be a N/A at he end

'*********************************** Other Consts not to ever edit ***********************************************
'This is the Var for seting the length. Note it is set above with the use of other consts DONT F WITH
'Const Length = CheckNumbers + StartRow - 1

'*********************************** The Vars ***********************************************
' The var for the drop down selection on a check for the resulting score
Dim Result As String
Dim Result2 As String

Dim CheckLine As Integer
'---------------the Result Counts-------------
Dim CountResult As PageResults
'set defaults
CountResult.FailCount = 0
CountResult.NACount = 0
CountResult.PassCount = 0
CountResult.TotalCount = 0
CountResult.Unfilled = 0
CountResult.UnselectedCount = 0
CountResult.WAVECount = 0
CountResult.SheetWasCounted = False
CountResult.Error = ""

'the var for the total number of sheets.
Dim SheetLength As Integer
Dim TotalLinesToIgnore As Integer

'the var sheet object.
Dim Sheet As Worksheet

Dim ChecklistEnd As Integer

ChecklistEnd = PrintArea.Row + PrintArea.Rows.Count - 1

'Instance a count and set all the vars to their starts.
SheetLength = Sheets.Count

'begin counting with iterating through the sheets.
Set Sheet = Sheets(SheetNum)
'only count it if it is visible.

If ((Not (Sheet.Visible = xlSheetHidden)) And StartRowResult = StartRowComment) Then
    CountResult.SheetWasCounted = True
    'for each check from the start point down them. Note start point is down.
    For CheckLine = StartRowResult + 1 To ChecklistEnd
        'add to the total count and grab a cell
        If Not (Sheet.Cells(CheckLine, StartColResult).EntireRow.Hidden) Then
            Result = Sheet.Cells(CheckLine, StartColResult)
            Result2 = Sheet.Cells(CheckLine, StartColComment)
            ' Cass out the checks into their counts
            If (Result = PassResultStr) Then
                CountResult.PassCount = CountResult.PassCount + 1
                CountResult.TotalCount = CountResult.TotalCount + 1
            ElseIf (Result = FailResultStr) Then
                CountResult.FailCount = CountResult.FailCount + 1
                CountResult.TotalCount = CountResult.TotalCount + 1
            ElseIf (Result = WAVEResultStr Or Result = WAVEResultStr2) Then
                CountResult.WAVECount = CountResult.WAVECount + 1
                CountResult.TotalCount = CountResult.TotalCount + 1
            ElseIf (Result = NAResultStr) Then
                CountResult.NACount = CountResult.NACount + 1
                CountResult.TotalCount = CountResult.TotalCount + 1
            ElseIf (Len(Result2) > 0) Then
                CountResult.UnselectedCount = CountResult.UnselectedCount + 1
                CountResult.TotalCount = CountResult.TotalCount + 1
            Else
                CountResult.Unfilled = CountResult.Unfilled + 1
            End If
            If (Sheet.Cells(CheckLine, StartColResult).MergeCells) Then
                CheckLine = CheckLine + ActiveCell.MergeArea.Columns.Count
            End If
        End If
    Next CheckLine
ElseIf Not (StartRowResult = StartRowComment) Then
    CountResult.Error = "The Rows do not match. It was disincuded from the count."
ElseIf Not (StartRowResult = -1 Or StartRowComment = -1 Or StartRowComment = -1 Or StartColComment = -1) Then
    CountResult.Error = "Failed to locate the start and end for this page. It was disincuded from the count."
Else
    CountResult.Error = "This page was hidden & assumed to not be need. It was disincuded from the count."
End If
    AutoQuickCount = CountResult
End Function


'ParamArray Arr() As Variant

Function FindAValueFromValues(MyIn As Range, SNum As Integer, LookFor() As String) As FindValuesResult
    Dim Result As FindValuesResult
    Dim i, j, n As Integer
    Dim CurrentString As String
    Result.Location.X = -1
    Result.Location.Y = -1
    Result.Success = False
    Result.ValueLocated = ""
    For n = LBound(LookFor) To UBound(LookFor)
        CurrentString = LookFor(n)
        CurrentString = Trim(CurrentString)
        For i = MyIn.Row To MyIn.Row + MyIn.Rows.Count - 1
            For j = MyIn.Column To MyIn.Column + MyIn.Columns.Count - 1
                If (Sheets(SNum).Cells(i, j).Value Like CurrentString) Then
                    Result.Location.X = j
                    Result.Location.Y = i
                    Result.Success = True
                    Result.ValueLocated = Sheets(SNum).Cells(i, j).Value
                    Result.FoundWith = CurrentString
                    FindAValueFromValues = Result
                    Exit Function
                End If
            Next j
        Next i
    Next n
    FindAValueFromValues = Result
End Function



Sub AutoStats()

Dim Config_ As config
Config_ = GetConfig

Dim Sheet As Worksheet
Dim CheckSheet As Integer

Dim CardArea As Range
Dim CommentsHeader As CellVector
Dim ResultsHeader As CellVector
Dim APageResult As PageResults
Dim CountResult As CountResults

Dim FRNumLoc As FindValuesResult
Dim BranchNumLoc As FindValuesResult
Dim DateLoc As FindValuesResult
Dim DCALoc As FindValuesResult
Dim CustomerNameLoc As FindValuesResult
Dim ReqByLoc As FindValuesResult

Dim FRNum As String
Dim BranchNum As String
Dim Date_ As String
Dim DCA_ As String
Dim CustomerName_ As String
Dim ReqBy_ As String

Dim Status As String

FRNum = "WAVE FR No: Not Found"
BranchNum = "Branch Number: Not Found"
Date_ = "DATE: Not Found"

DCA_ = "Data Compliance Analyst Who Completed Audit Verification: Not Found"
CustomerName_ = "Customer Name: Not Found"
ReqBy_ = "Requested By: Not Found"

CountResult.FailCount = 0
CountResult.NACount = 0
CountResult.PassCount = 0
CountResult.TotalCount = 0
CountResult.Unfilled = 0
CountResult.UnselectedCount = 0
CountResult.WAVECount = 0

For CheckSheet = 1 To Sheets.Count
    Set Sheet = Sheets(CheckSheet)
    If Not (Sheet.Visible = xlSheetHidden) Then
        'select a sheet.
        Set CardArea = Range(Worksheets(CheckSheet).PageSetup.PrintArea)
        
        CommentsHeader = FindAValueFromValues(CardArea, CheckSheet, Config_.CommentSearchHeaders).Location
        ResultsHeader = FindAValueFromValues(CardArea, CheckSheet, Config_.ResultsSearchHeaders).Location
        
        If Not (DateLoc.Success) Then
            DateLoc = FindAValueFromValues(CardArea, CheckSheet, Config_.DateSearchHeaders)
            If (DateLoc.Success) Then
                Date_ = DateLoc.ValueLocated
            End If
        End If
        
        If Not (FRNumLoc.Success) Then
            FRNumLoc = FindAValueFromValues(CardArea, CheckSheet, Config_.FRNumberSearchHeaders)
            If (FRNumLoc.Success) Then
                FRNum = FRNumLoc.ValueLocated
            End If
        End If
        
        If Not (BranchNumLoc.Success) Then
            BranchNumLoc = FindAValueFromValues(CardArea, CheckSheet, Config_.BranchNumberSearchHeaders)
            If (BranchNumLoc.Success) Then
                BranchNum = BranchNumLoc.ValueLocated
            End If
        End If
        
        If Not (DCALoc.Success) Then
            DCALoc = FindAValueFromValues(CardArea, CheckSheet, Config_.DCASearchHeaders)
            If (DCALoc.Success) Then
                DCA_ = DCALoc.ValueLocated
            End If
        End If
        
        If Not (CustomerNameLoc.Success) Then
            CustomerNameLoc = FindAValueFromValues(CardArea, CheckSheet, Config_.CustomerNameSearchHeaders)
            If (CustomerNameLoc.Success) Then
                CustomerName_ = CustomerNameLoc.ValueLocated
            End If
        End If
        
        If Not (ReqByLoc.Success) Then
            ReqByLoc = FindAValueFromValues(CardArea, CheckSheet, Config_.RequestedBySearchHeaders)
            If (ReqByLoc.Success) Then
                ReqBy_ = ReqByLoc.ValueLocated
            End If
        End If
        
        APageResult = AutoQuickCount(CheckSheet, ResultsHeader.Y, ResultsHeader.X, CommentsHeader.Y, CommentsHeader.X, CardArea)
        
        CountResult.FailCount = CountResult.FailCount + APageResult.FailCount
        CountResult.NACount = CountResult.NACount + APageResult.NACount
        CountResult.PassCount = CountResult.PassCount + APageResult.PassCount
        CountResult.TotalCount = CountResult.TotalCount + APageResult.TotalCount
        CountResult.Unfilled = CountResult.Unfilled + APageResult.Unfilled
        CountResult.UnselectedCount = CountResult.UnselectedCount + APageResult.UnselectedCount
        CountResult.WAVECount = CountResult.WAVECount + APageResult.WAVECount
    End If
Next CheckSheet

If (CountResult.FailCount = 0) Then
    Status = "Status: Pass"
Else
    Status = "Status: Fail"
End If

MsgBox _
CustomerName_ _
& vbNewLine & ReqBy_ _
& vbNewLine & "----------------------------------------------" _
& vbNewLine & FRNum _
& vbNewLine & Date_ _
& vbNewLine & BranchNum _
& vbNewLine & Status _
& vbNewLine & "----------------------------------------------" _
& vbNewLine & "Total Count: " & CountResult.TotalCount _
& vbNewLine & "Total N/A Count: " & (CountResult.NACount + CountResult.UnselectedCount) _
& vbNewLine & "Fail Count: " & CountResult.FailCount _
& vbNewLine & "Wavered Count: " & CountResult.WAVECount _
& vbNewLine & "Pass Count: " & CountResult.PassCount _
& vbNewLine & "----------------------------------------------" _
& vbNewLine & Replace(DCA_, Mid(DCALoc.FoundWith, 1, Len(DCALoc.FoundWith) - 1), "DCA:") _
, vbOKOnly, "Count Results"

End Sub



Sub AutoStatsSelected()

Dim Config_ As config
Config_ = GetConfig

Dim Sheet As Worksheet
Dim CheckSheet As Integer

Dim CardArea As Range
Dim CommentsHeader As CellVector
Dim ResultsHeader As CellVector
Dim APageResult As PageResults
Dim CountResult As CountResults

Dim FRNumLoc As FindValuesResult
Dim BranchNumLoc As FindValuesResult
Dim DateLoc As FindValuesResult
Dim DCALoc As FindValuesResult
Dim CustomerNameLoc As FindValuesResult
Dim ReqByLoc As FindValuesResult

Dim FRNum As String
Dim BranchNum As String
Dim Date_ As String
Dim DCA_ As String
Dim CustomerName_ As String
Dim ReqBy_ As String

Dim Status As String

FRNum = "WAVE FR No: Not Found"
BranchNum = "Branch Number: Not Found"
Date_ = "DATE: Not Found"

DCA_ = "Data Compliance Analyst Who Completed Audit Verification: Not Found"
CustomerName_ = "Customer Name: Not Found"
ReqBy_ = "Requested By: Not Found"

CountResult.FailCount = 0
CountResult.NACount = 0
CountResult.PassCount = 0
CountResult.TotalCount = 0
CountResult.Unfilled = 0
CountResult.UnselectedCount = 0
CountResult.WAVECount = 0

Dim sh As Object


 

For Each sh In ActiveWindow.SelectedSheets

    Set Sheet = Sheets(sh.Index)
    If Not (Sheet.Visible = xlSheetHidden) Then
        'select a sheet.
        Set CardArea = Range(Worksheets(sh.Index).PageSetup.PrintArea)
        
        CommentsHeader = FindAValueFromValues(CardArea, sh.Index, Config_.CommentSearchHeaders).Location
        ResultsHeader = FindAValueFromValues(CardArea, sh.Index, Config_.ResultsSearchHeaders).Location
        
        If Not (DateLoc.Success) Then
            DateLoc = FindAValueFromValues(CardArea, sh.Index, Config_.DateSearchHeaders)
            If (DateLoc.Success) Then
                Date_ = DateLoc.ValueLocated
            End If
        End If
        
        If Not (FRNumLoc.Success) Then
            FRNumLoc = FindAValueFromValues(CardArea, sh.Index, Config_.FRNumberSearchHeaders)
            If (FRNumLoc.Success) Then
                FRNum = FRNumLoc.ValueLocated
            End If
        End If
        
        If Not (BranchNumLoc.Success) Then
            BranchNumLoc = FindAValueFromValues(CardArea, sh.Index, Config_.BranchNumberSearchHeaders)
            If (BranchNumLoc.Success) Then
                BranchNum = BranchNumLoc.ValueLocated
            End If
        End If
        
        If Not (DCALoc.Success) Then
            DCALoc = FindAValueFromValues(CardArea, sh.Index, Config_.DCASearchHeaders)
            If (DCALoc.Success) Then
                DCA_ = DCALoc.ValueLocated
            End If
        End If
        
        If Not (CustomerNameLoc.Success) Then
            CustomerNameLoc = FindAValueFromValues(CardArea, sh.Index, Config_.CustomerNameSearchHeaders)
            If (CustomerNameLoc.Success) Then
                CustomerName_ = CustomerNameLoc.ValueLocated
            End If
        End If
        
        If Not (ReqByLoc.Success) Then
            ReqByLoc = FindAValueFromValues(CardArea, sh.Index, Config_.RequestedBySearchHeaders)
            If (ReqByLoc.Success) Then
                ReqBy_ = ReqByLoc.ValueLocated
            End If
        End If
        
        APageResult = AutoQuickCount(sh.Index, ResultsHeader.Y, ResultsHeader.X, CommentsHeader.Y, CommentsHeader.X, CardArea)
        
        CountResult.FailCount = CountResult.FailCount + APageResult.FailCount
        CountResult.NACount = CountResult.NACount + APageResult.NACount
        CountResult.PassCount = CountResult.PassCount + APageResult.PassCount
        CountResult.TotalCount = CountResult.TotalCount + APageResult.TotalCount
        CountResult.Unfilled = CountResult.Unfilled + APageResult.Unfilled
        CountResult.UnselectedCount = CountResult.UnselectedCount + APageResult.UnselectedCount
        CountResult.WAVECount = CountResult.WAVECount + APageResult.WAVECount
    End If
Next sh

If (CountResult.FailCount = 0) Then
    Status = "Status: Pass"
Else
    Status = "Status: Fail"
End If

MsgBox _
CustomerName_ _
& vbNewLine & ReqBy_ _
& vbNewLine & "----------------------------------------------" _
& vbNewLine & FRNum _
& vbNewLine & Date_ _
& vbNewLine & BranchNum _
& vbNewLine & Status _
& vbNewLine & "----------------------------------------------" _
& vbNewLine & "Total Count: " & CountResult.TotalCount _
& vbNewLine & "Total N/A Count: " & (CountResult.NACount + CountResult.UnselectedCount) _
& vbNewLine & "Fail Count: " & CountResult.FailCount _
& vbNewLine & "Wavered Count: " & CountResult.WAVECount _
& vbNewLine & "Pass Count: " & CountResult.PassCount _
& vbNewLine & "----------------------------------------------" _
& vbNewLine & Replace(DCA_, Mid(DCALoc.FoundWith, 1, Len(DCALoc.FoundWith) - 1), "DCA:") _
, vbOKOnly, "Count Results"

End Sub



Sub ReGenerateConfigFile()
    If MsgBox("This will auto generate a VBA config file. Please read the guide before using this method. Would you still like to proceed?", vbYesNo) = vbNo Then Exit Sub
    Dim FileOutput As String
    StrConfigReturn = GetStringRepOf()
    EasyConfig.Show
    If EasyConfigReturn.Answer = Ok Then
    FileOutput = BuildFileOutput()
    
    'Dim FilePath As String
    'Dim OldFilePath As String
    'FilePath = Environ("temp") & "\MacroConfig.bas"
    'OldFilePath = Environ("temp") & "\RecallMacroConfig.bas"

    Dim Bool As Boolean
    Bool = GenerateFile(FileOutput, EasyConfigReturn.SaveLocation)
    
    
    'Run time is not supported. ditched
    'ActiveWorkbook.VBProject.VBComponents("MacroConfig").Export ("OldFilePath")
    'ActiveWorkbook.VBProject.VBComponents.Remove (ActiveWorkbook.VBProject.VBComponents("MacroConfig"))
    'ActiveWorkbook.VBProject.VBComponents.Import FilePath
    
    End If
End Sub


Function GenerateFile(COut As String, FilePath As String) As Boolean
    Dim FileOutput As String
    

    ' The advantage of correctly typing fso as FileSystemObject is to make autocompletion
    ' (Intellisense) work, which helps you avoid typos and lets you discover other useful
    ' methods of the FileSystemObject
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim fileStream As TextStream

    ' Here the actual file is created and opened for write access
    Set fileStream = fso.CreateTextFile(FilePath)

    ' Write something to the file
    fileStream.WriteLine COut

    ' Close it, so it is not locked anymore
    fileStream.Close

    ' Here is another great method of the FileSystemObject that checks if a file exists
    ' Explicitly setting objects to Nothing should not be necessary in most cases, but if
    ' you're writing macros for Microsoft Access, you may want to uncomment the following
    ' two lines (see https://stackoverflow.com/a/517202/2822719 for details):
    Set fileStream = Nothing
    Set fso = Nothing
End Function

Public Function BuildFileOutput() As String
    Dim Output As String
    Dim Stringes() As String
        Output = "Attribute VB_Name = ""MacroConfig""" & vbNewLine & _
        "Option Explicit" & vbNewLine & _
        "Public Function GetConfig() As config" & vbNewLine & _
        "Dim ConfigReturn As config" & vbNewLine
        
        Output = Output & BuildOutputFor("BranchNumberSearchHeaders", EasyConfigReturn.Config_.BranchNumberSearchHeaders)
        
        Output = Output & BuildOutputFor("CommentSearchHeaders", EasyConfigReturn.Config_.CommentSearchHeaders)

        Output = Output & BuildOutputFor("CustomerNameSearchHeaders", EasyConfigReturn.Config_.CustomerNameSearchHeaders)
        
        Output = Output & BuildOutputFor("DateSearchHeaders", EasyConfigReturn.Config_.DateSearchHeaders)
        
        Output = Output & BuildOutputFor("DCASearchHeaders", EasyConfigReturn.Config_.DCASearchHeaders)
        
        Output = Output & BuildOutputFor("FRNumberSearchHeaders", EasyConfigReturn.Config_.FRNumberSearchHeaders)

        Output = Output & BuildOutputFor("RequestedBySearchHeaders", EasyConfigReturn.Config_.RequestedBySearchHeaders)
        
        Output = Output & BuildOutputFor("ResultsSearchHeaders", EasyConfigReturn.Config_.ResultsSearchHeaders)

        Output = Output & "GetConfig = ConfigReturn" & vbNewLine & _
        "End Function"
    BuildFileOutput = Output
End Function

Public Function BuildOutputFor(NameTo As String, Seperate As String) As String
    Dim Stringes() As String
    Dim Output As String
    Dim n As Integer
    BuildOutputFor = ""
    Stringes = Split(Seperate, "+")
    n = UBound(Stringes) - LBound(Stringes)
    If n < 0 Then n = 0
    Output = Output & "ReDim ConfigReturn." & NameTo & "(" & (n) & ")" & vbNewLine
    For n = LBound(Stringes) To UBound(Stringes)
        Output = Output & "ConfigReturn." & NameTo & "(" & n & ") = """ & Stringes(n) & """" & vbNewLine
    Next n
    BuildOutputFor = Output
End Function

Public Function GetStringRepOf() As configAsString
    Dim n As Integer
    Dim Config_ As config
    Dim StrConfig As configAsString
    Config_ = GetConfig()
    StrConfig.BranchNumberSearchHeaders = Config_.BranchNumberSearchHeaders(0)
    For n = LBound(Config_.BranchNumberSearchHeaders) + 1 To UBound(Config_.BranchNumberSearchHeaders)
        StrConfig.BranchNumberSearchHeaders = StrConfig.BranchNumberSearchHeaders & "+" & Config_.BranchNumberSearchHeaders(n)
    Next n
    StrConfig.CommentSearchHeaders = Config_.CommentSearchHeaders(0)
    For n = LBound(Config_.CommentSearchHeaders) + 1 To UBound(Config_.CommentSearchHeaders)
        StrConfig.CommentSearchHeaders = StrConfig.CommentSearchHeaders & "+" & Config_.CommentSearchHeaders(n)
    Next n
    StrConfig.CustomerNameSearchHeaders = Config_.CustomerNameSearchHeaders(0)
    For n = LBound(Config_.CustomerNameSearchHeaders) + 1 To UBound(Config_.CustomerNameSearchHeaders)
        StrConfig.CustomerNameSearchHeaders = StrConfig.CustomerNameSearchHeaders & "+" & Config_.CustomerNameSearchHeaders(n)
    Next n
    StrConfig.DateSearchHeaders = Config_.DateSearchHeaders(0)
    For n = LBound(Config_.DateSearchHeaders) + 1 To UBound(Config_.DateSearchHeaders)
        StrConfig.DateSearchHeaders = StrConfig.DateSearchHeaders & "+" & Config_.DateSearchHeaders(n)
    Next n
    StrConfig.DCASearchHeaders = Config_.DCASearchHeaders(0)
    For n = LBound(Config_.DCASearchHeaders) + 1 To UBound(Config_.DCASearchHeaders)
        StrConfig.DCASearchHeaders = StrConfig.DCASearchHeaders & "+" & Config_.DCASearchHeaders(n)
    Next n
    StrConfig.FRNumberSearchHeaders = Config_.FRNumberSearchHeaders(0)
    For n = LBound(Config_.FRNumberSearchHeaders) + 1 To UBound(Config_.FRNumberSearchHeaders)
        StrConfig.FRNumberSearchHeaders = StrConfig.FRNumberSearchHeaders & "+" & Config_.FRNumberSearchHeaders(n)
    Next n
    StrConfig.RequestedBySearchHeaders = Config_.RequestedBySearchHeaders(0)
    For n = LBound(Config_.RequestedBySearchHeaders) + 1 To UBound(Config_.RequestedBySearchHeaders)
        StrConfig.RequestedBySearchHeaders = StrConfig.RequestedBySearchHeaders & "+" & Config_.RequestedBySearchHeaders(n)
    Next n
    StrConfig.ResultsSearchHeaders = Config_.ResultsSearchHeaders(0)
    For n = LBound(Config_.ResultsSearchHeaders) + 1 To UBound(Config_.ResultsSearchHeaders)
        StrConfig.ResultsSearchHeaders = StrConfig.ResultsSearchHeaders & "+" & Config_.ResultsSearchHeaders(n)
    Next n
    GetStringRepOf = StrConfig
End Function
