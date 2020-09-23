Attribute VB_Name = "mLikeCompareBinary"
Option Explicit
Option Compare Binary

Public Function LikeCompBinary(Str As String, Match As String) As Boolean
    LikeCompBinary = CBool(Str Like Match)
End Function

