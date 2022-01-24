Attribute VB_Name = "Module3"
Sub RemoveNotesAndFormat()
'PURPOSE: Calls all the other Subs from Module2 to clean up the data and format as desired
'SOURCE: https://github.com/hamartithia
    Call Delete_Rows_Based_On_Value
    Call FindReplaceBO
    Call FindReplaceBX
    Call FindReplaceCOS
    Call FindReplaceEP
    Call FindReplaceLS
    Call FindReplaceNP
    Call FindReplaceQQ
    Call FindReplaceTELE
    Call FindReplaceXC
    Call FindReplaceZZ
    Call DeleteEThruI
End Sub
