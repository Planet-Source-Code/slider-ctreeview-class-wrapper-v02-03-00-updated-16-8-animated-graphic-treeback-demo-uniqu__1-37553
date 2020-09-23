Attribute VB_Name = "mLikeCompareText"
Option Explicit
Option Compare Text

Public Function LikeCompText(Str As String, Match As String) As Boolean
    LikeCompText = CBool(Str Like Match)
End Function


