Attribute VB_Name = "cutPasteRows"
Sub CutPaste31Down()
Attribute CutPaste31Down.VB_ProcData.VB_Invoke_Func = " \n14"
'PURPOSE: Cut and paste rows 31 and below over to F3
'SOURCE: https://github.com/hamartithia
    Range("A31:D1000").Select
    Selection.Cut
    Range("F3").Select
    ActiveSheet.Paste
End Sub
