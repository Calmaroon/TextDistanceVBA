Attribute VB_Name = "modHamming"
'===============================================================================
' Hamming distance - port of textdistance.Hamming (Python textdistance 4.6)
'
' The number of positions at which two strings differ. Strings of different
' lengths are allowed: every extra character in the longer string counts as a
' mismatch, e.g. Hamming_Distance("test", "testing") = 3.
'
' Pure VBA with no host references, so it works in Excel, Access, Word, etc.
'
' Common arguments
'   s1, s2         Strings to compare.
'   CaseSensitive  True (default) compares exactly; False ignores case.
'                  (VBA addition - Python's textdistance has no such flag.)
'   Truncate       False (default) compares over the longer string's length;
'                  True compares only the first Min(Len(s1), Len(s2))
'                  characters, matching textdistance's truncate=True.
'===============================================================================
Option Explicit
Option Compare Binary

' Number of differing positions.
Public Function Hamming_Distance(ByVal s1 As String, ByVal s2 As String, _
                                 Optional ByVal CaseSensitive As Boolean = True, _
                                 Optional ByVal Truncate As Boolean = False) As Long
    Dim len1 As Long, len2 As Long, minLen As Long
    Dim i As Long, mismatches As Long

    If Not CaseSensitive Then
        s1 = LCase$(s1)
        s2 = LCase$(s2)
    End If

    If s1 = s2 Then Exit Function

    len1 = Len(s1)
    len2 = Len(s2)

    ' textdistance returns the maximum length when either string is empty,
    ' even with truncate=True
    If len1 = 0 Or len2 = 0 Then
        Hamming_Distance = IIf(len1 > len2, len1, len2)
        Exit Function
    End If

    minLen = IIf(len1 < len2, len1, len2)
    For i = 1 To minLen
        If Mid$(s1, i, 1) <> Mid$(s2, i, 1) Then mismatches = mismatches + 1
    Next i

    If Not Truncate Then mismatches = mismatches + Abs(len1 - len2)

    Hamming_Distance = mismatches
End Function

' Largest possible distance: the length of the longer string.
Public Function Hamming_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    Hamming_Maximum = IIf(Len(s1) > Len(s2), Len(s1), Len(s2))
End Function

' Number of matching positions: Maximum - Distance.
Public Function Hamming_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                   Optional ByVal CaseSensitive As Boolean = True, _
                                   Optional ByVal Truncate As Boolean = False) As Long
    Hamming_Similarity = Hamming_Maximum(s1, s2) - Hamming_Distance(s1, s2, CaseSensitive, Truncate)
End Function

' Distance scaled to 0 (identical) .. 1 (completely different).
' Two empty strings give 0.
Public Function Hamming_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                           Optional ByVal CaseSensitive As Boolean = True, _
                                           Optional ByVal Truncate As Boolean = False) As Double
    Dim maximum As Long
    maximum = Hamming_Maximum(s1, s2)
    If maximum = 0 Then Exit Function
    Hamming_NormalizedDistance = Hamming_Distance(s1, s2, CaseSensitive, Truncate) / maximum
End Function

' Similarity scaled to 0 (completely different) .. 1 (identical).
' Two empty strings give 1.
Public Function Hamming_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                             Optional ByVal CaseSensitive As Boolean = True, _
                                             Optional ByVal Truncate As Boolean = False) As Double
    Hamming_NormalizedSimilarity = 1 - Hamming_NormalizedDistance(s1, s2, CaseSensitive, Truncate)
End Function
