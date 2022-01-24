Attribute VB_Name = "callSubs"
Sub RemoveNotesAndFormat()
'PURPOSE: Calls all the other Subs from hiddenSubs to clean up the data and format as desired
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
