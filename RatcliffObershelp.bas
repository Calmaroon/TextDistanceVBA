Attribute VB_Name = "modRatcliffObershelp"
'===============================================================================
' Ratcliff-Obershelp similarity (Gestalt pattern matching) - port of
' textdistance.RatcliffObershelp (Python textdistance 4.6)
'
' Finds the longest common substring, then recursively the longest common
' substrings to its left and right. The similarity is twice the number of
' matched characters divided by the total length of both strings - the same
' idea as Python's difflib.SequenceMatcher.ratio():
'   RatcliffObershelp_Similarity("WIKIMEDIA", "WIKIMANIA") = 0.778
'
' Similarity is the primary value; the maximum is always 1, so
' Distance = 1 - Similarity.
'
' Pure VBA with no host references, so it works in Excel, Access, Word, etc.
' Self-contained: does not need LCSStr.bas.
'
' Common arguments
'   s1, s2         Strings to compare.
'   CaseSensitive  True (default) compares exactly; False ignores case.
'                  (VBA addition - Python's textdistance has no such flag.)
'===============================================================================
Option Explicit
Option Compare Binary

Public Function RatcliffObershelp_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                             Optional ByVal CaseSensitive As Boolean = True) As Double
    If Not CaseSensitive Then
        s1 = LCase$(s1)
        s2 = LCase$(s2)
    End If

    If s1 = s2 Then
        RatcliffObershelp_Similarity = 1
        Exit Function
    End If
    If Len(s1) = 0 Or Len(s2) = 0 Then Exit Function

    RatcliffObershelp_Similarity = 2 * MatchingCharacters(s1, s2) / (Len(s1) + Len(s2))
End Function

Public Function RatcliffObershelp_Distance(ByVal s1 As String, ByVal s2 As String, _
                                           Optional ByVal CaseSensitive As Boolean = True) As Double
    RatcliffObershelp_Distance = 1 - RatcliffObershelp_Similarity(s1, s2, CaseSensitive)
End Function

Public Function RatcliffObershelp_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                                       Optional ByVal CaseSensitive As Boolean = True) As Double
    RatcliffObershelp_NormalizedSimilarity = 1 - RatcliffObershelp_NormalizedDistance(s1, s2, CaseSensitive)
End Function

Public Function RatcliffObershelp_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                                     Optional ByVal CaseSensitive As Boolean = True) As Double
    RatcliffObershelp_NormalizedDistance = RatcliffObershelp_Distance(s1, s2, CaseSensitive)
End Function

' Always 1.
Public Function RatcliffObershelp_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    RatcliffObershelp_Maximum = 1
End Function

' Total length of the longest common substring plus the matches found
' recursively to its left and right.
Private Function MatchingCharacters(ByVal s1 As String, ByVal s2 As String) As Long
    Dim common As String, n As Long, p1 As Long, p2 As Long

    common = LongestCommonSubstring(s1, s2)
    n = Len(common)
    If n = 0 Then Exit Function

    p1 = InStr(1, s1, common, vbBinaryCompare)
    p2 = InStr(1, s2, common, vbBinaryCompare)
    MatchingCharacters = MatchingCharacters(Left$(s1, p1 - 1), Left$(s2, p2 - 1)) _
                       + n _
                       + MatchingCharacters(Mid$(s1, p1 + n), Mid$(s2, p2 + n))
End Function

' Longest common substring, choosing between equally long ones like Python's
' textdistance.LCSStr: the earliest in s1, except when either string has 200+
' characters, where it takes the earliest in the shorter string.
Private Function LongestCommonSubstring(ByVal s1 As String, ByVal s2 As String) As String
    Dim a As String, b As String
    Dim lenA As Long, lenB As Long, i As Long, j As Long
    Dim ca() As Long, cb() As Long, prev() As Long, cur() As Long, tmp() As Long
    Dim bestLen As Long, bestEnd As Long

    If (Len(s1) >= 200 Or Len(s2) >= 200) And Len(s2) < Len(s1) Then
        a = s2: b = s1
    Else
        a = s1: b = s2
    End If

    lenA = Len(a)
    lenB = Len(b)
    If lenA = 0 Or lenB = 0 Then Exit Function

    CharCodes a, ca
    CharCodes b, cb

    ' cur(j) = length of the common run ending at a(i) and b(j)
    ReDim prev(0 To lenB)
    ReDim cur(0 To lenB)
    For i = 1 To lenA
        For j = 1 To lenB
            If ca(i) = cb(j) Then
                cur(j) = prev(j - 1) + 1
                If cur(j) > bestLen Then
                    bestLen = cur(j)
                    bestEnd = i
                End If
            Else
                cur(j) = 0
            End If
        Next j
        tmp = prev: prev = cur: cur = tmp
    Next i

    If bestLen > 0 Then LongestCommonSubstring = Mid$(a, bestEnd - bestLen + 1, bestLen)
End Function

' Character codes of s, 1-based, so comparisons avoid repeated Mid$ calls.
Private Sub CharCodes(ByVal s As String, codes() As Long)
    Dim i As Long
    ReDim codes(1 To Len(s))
    For i = 1 To Len(s)
        codes(i) = AscW(Mid$(s, i, 1))
    Next i
End Sub
