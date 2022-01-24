Attribute VB_Name = "hiddenSubs"
Option Private Module
Sub Delete_Rows_Based_On_Value()
'PURPOSE: Apply a filter to a Range and delete rows based on value in column 1
'SOURCE: https://stackoverflow.com/a/23525044

Dim RowToTest As Long

'Iterates through rows that contain any content in column 2, runs from bottomon and ends by running on row 3, iterates upward 1 step at a time
For RowToTest = Cells(Rows.Count, 2).End(xlUp).Row To 3 Step -1

'values to check in column 1--if value is in column 1, it deletes the row; otherwise, it is ignored
With Cells(RowToTest, 1)
    If .Value = "n/a" _
    Or .Value = "Note:" _
    Or .Value = "" _
    Then _
    Rows(RowToTest).EntireRow.Delete
End With

'continues to iterate through each row as set up above with var RowToTest
Next RowToTest

End Sub
Sub FindReplaceEP()
'PURPOSE: Find & Replace text/values throughout entire workbook
'SOURCE: www.TheSpreadsheetGuru.com

Dim sht As Worksheet
Dim fnd As Variant
Dim rplc As Variant

fnd = "ESTABLISHED PATIENT"
rplc = "EP"

For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht

End Sub
Sub FindReplaceNP()
'PURPOSE: Find & Replace text/values throughout entire workbook
'SOURCE: www.TheSpreadsheetGuru.com

Dim sht As Worksheet
Dim fnd As Variant
Dim rplc As Variant

fnd = "NEW PATIENT"
rplc = "NP"

For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht

End Sub
Sub FindReplaceXC()
'PURPOSE: Find & Replace text/values throughout entire workbook
'SOURCE: www.TheSpreadsheetGuru.com

Dim sht As Worksheet
Dim fnd As Variant
Dim rplc As Variant

fnd = "EXCISION/PROCEDURE"
rplc = "XC"

For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht

End Sub
Sub FindReplaceQQ()
'PURPOSE: Find & Replace text/values throughout entire workbook
'SOURCE: www.TheSpreadsheetGuru.com

Dim sht As Worksheet
Dim fnd As Variant
Dim rplc As Variant

fnd = "NP FULL SKIN EXAM"
rplc = "QQ"

For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht

End Sub
Sub FindReplaceZZ()
'PURPOSE: Find & Replace text/values throughout entire workbook
'SOURCE: www.TheSpreadsheetGuru.com

Dim sht As Worksheet
Dim fnd As Variant
Dim rplc As Variant

fnd = "FULL SKIN EXAM - EP"
rplc = "ZZ"

For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht

End Sub
Sub FindReplaceLS()
'PURPOSE: Find & Replace text/values throughout entire workbook
'SOURCE: www.TheSpreadsheetGuru.com

Dim sht As Worksheet
Dim fnd As Variant
Dim rplc As Variant

fnd = "LASER"
rplc = "LS"

For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht

End Sub
Sub FindReplaceCOS()
'PURPOSE: Find & Replace text/values throughout entire workbook
'SOURCE: www.TheSpreadsheetGuru.com

Dim sht As Worksheet
Dim fnd As Variant
Dim rplc As Variant

fnd = "COSMETIC"
rplc = "COS"

For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht

End Sub
Sub FindReplaceBO()
'PURPOSE: Find & Replace text/values throughout entire workbook
'SOURCE: www.TheSpreadsheetGuru.com

Dim sht As Worksheet
Dim fnd As Variant
Dim rplc As Variant

fnd = "BOTOX"
rplc = "BO"

For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht

End Sub
Sub FindReplaceBX()
'PURPOSE: Find & Replace text/values throughout entire workbook
'SOURCE: www.TheSpreadsheetGuru.com

Dim sht As Worksheet
Dim fnd As Variant
Dim rplc As Variant

fnd = "BIOPSY"
rplc = "BX"

For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht

End Sub
Sub FindReplaceTELE()
'PURPOSE: Find & Replace text/values throughout entire workbook
'SOURCE: www.TheSpreadsheetGuru.com

Dim sht As Worksheet
Dim fnd As Variant
Dim rplc As Variant

fnd = "TELEMEDICINE"
rplc = "TELE"

For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next sht

End Sub
Sub DeleteEThruI()
'PURPOSE: Delete all content in Columns E through I
'SOURCE: https://github.com/hamartithia
    Columns("E:I").ClearContents
End Sub
