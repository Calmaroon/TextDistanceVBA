Attribute VB_Name = "modLevenshtein"
'===============================================================================
' Levenshtein distance - port of textdistance.Levenshtein (Python textdistance 4.6)
'
' The minimum number of single-character insertions, deletions and
' substitutions needed to turn one string into the other,
' e.g. Levenshtein_Distance("kitten", "sitting") = 3.
'
' Pure VBA with no host references, so it works in Excel, Access, Word, etc.
'
' Common arguments
'   s1, s2         Strings to compare.
'   CaseSensitive  True (default) compares exactly; False ignores case.
'                  (VBA addition - Python's textdistance has no such flag.)
'===============================================================================
Option Explicit
Option Compare Binary

' Minimum number of edits.
Public Function Levenshtein_Distance(ByVal s1 As String, ByVal s2 As String, _
                                     Optional ByVal CaseSensitive As Boolean = True) As Long
    Dim len1 As Long, len2 As Long
    Dim c1() As Long, c2() As Long
    Dim prev() As Long, cur() As Long, tmp() As Long
    Dim r As Long, c As Long, best As Long, cost As Long

    If Not CaseSensitive Then
        s1 = LCase$(s1)
        s2 = LCase$(s2)
    End If

    If s1 = s2 Then Exit Function

    len1 = Len(s1)
    len2 = Len(s2)
    If len1 = 0 Or len2 = 0 Then
        Levenshtein_Distance = len1 + len2
        Exit Function
    End If

    CharCodes s1, c1
    CharCodes s2, c2

    ' Two-row dynamic programming table
    ReDim prev(0 To len2)
    ReDim cur(0 To len2)
    For c = 0 To len2
        cur(c) = c
    Next c

    For r = 1 To len1
        tmp = prev: prev = cur: cur = tmp
        cur(0) = r
        For c = 1 To len2
            If c1(r) = c2(c) Then cost = 0 Else cost = 1
            best = prev(c - 1) + cost                               ' substitution
            If prev(c) + 1 < best Then best = prev(c) + 1           ' deletion
            If cur(c - 1) + 1 < best Then best = cur(c - 1) + 1     ' insertion
            cur(c) = best
        Next c
    Next r

    Levenshtein_Distance = cur(len2)
End Function

' Largest possible distance: the length of the longer string.
Public Function Levenshtein_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    Levenshtein_Maximum = IIf(Len(s1) > Len(s2), Len(s1), Len(s2))
End Function

' Maximum - Distance.
Public Function Levenshtein_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                       Optional ByVal CaseSensitive As Boolean = True) As Long
    Levenshtein_Similarity = Levenshtein_Maximum(s1, s2) - Levenshtein_Distance(s1, s2, CaseSensitive)
End Function

' Distance scaled to 0 (identical) .. 1 (completely different).
' Two empty strings give 0.
Public Function Levenshtein_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                               Optional ByVal CaseSensitive As Boolean = True) As Double
    Dim maximum As Long
    maximum = Levenshtein_Maximum(s1, s2)
    If maximum = 0 Then Exit Function
    Levenshtein_NormalizedDistance = Levenshtein_Distance(s1, s2, CaseSensitive) / maximum
End Function

' Similarity scaled to 0 (completely different) .. 1 (identical).
' Two empty strings give 1.
Public Function Levenshtein_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                                 Optional ByVal CaseSensitive As Boolean = True) As Double
    Levenshtein_NormalizedSimilarity = 1 - Levenshtein_NormalizedDistance(s1, s2, CaseSensitive)
End Function

' Character codes of s, 1-based, so comparisons avoid repeated Mid$ calls.
Private Sub CharCodes(ByVal s As String, codes() As Long)
    Dim i As Long
    ReDim codes(1 To Len(s))
    For i = 1 To Len(s)
        codes(i) = AscW(Mid$(s, i, 1))
    Next i
End Sub
