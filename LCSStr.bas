Attribute VB_Name = "modLCSStr"
'===============================================================================
' Longest common substring - port of textdistance.LCSStr
' (Python textdistance 4.6)
'
' The longest run of consecutive characters appearing in both strings:
'   LCSStr_Substring("test", "text")   = "te"
'   LCSStr_Similarity("test", "text") = 2   (its length)
'
' Pure VBA with no host references, so it works in Excel, Access, Word, etc.
'
' Common arguments
'   s1, s2         Strings to compare.
'   CaseSensitive  True (default) compares exactly; False ignores case (the
'                  returned substring is then lower case).
'                  (VBA addition - Python's textdistance has no such flag.)
'===============================================================================
Option Explicit
Option Compare Binary

' The longest common substring itself (Python: textdistance.lcsstr(s1, s2)).
' When several exist, returns the same one as Python.
Public Function LCSStr_Substring(ByVal s1 As String, ByVal s2 As String, _
                                 Optional ByVal CaseSensitive As Boolean = True) As String
    If Not CaseSensitive Then
        s1 = LCase$(s1)
        s2 = LCase$(s2)
    End If
    LCSStr_Substring = LongestCommonSubstring(s1, s2)
End Function

' Length of the longest common substring.
Public Function LCSStr_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                  Optional ByVal CaseSensitive As Boolean = True) As Long
    LCSStr_Similarity = Len(LCSStr_Substring(s1, s2, CaseSensitive))
End Function

' Largest possible similarity: the length of the longer string.
Public Function LCSStr_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    LCSStr_Maximum = IIf(Len(s1) > Len(s2), Len(s1), Len(s2))
End Function

' Maximum - Similarity.
Public Function LCSStr_Distance(ByVal s1 As String, ByVal s2 As String, _
                                Optional ByVal CaseSensitive As Boolean = True) As Long
    LCSStr_Distance = LCSStr_Maximum(s1, s2) - LCSStr_Similarity(s1, s2, CaseSensitive)
End Function

' Distance scaled to 0 (identical) .. 1 (nothing in common).
' Two empty strings give 0.
Public Function LCSStr_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                          Optional ByVal CaseSensitive As Boolean = True) As Double
    Dim maximum As Long
    maximum = LCSStr_Maximum(s1, s2)
    If maximum = 0 Then Exit Function
    LCSStr_NormalizedDistance = LCSStr_Distance(s1, s2, CaseSensitive) / maximum
End Function

' Similarity scaled to 0 (nothing in common) .. 1 (identical).
' Two empty strings give 1.
Public Function LCSStr_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                            Optional ByVal CaseSensitive As Boolean = True) As Double
    LCSStr_NormalizedSimilarity = 1 - LCSStr_NormalizedDistance(s1, s2, CaseSensitive)
End Function

' Longest common substring, choosing between equally long ones like Python:
' the earliest in s1, except when either string has 200+ characters, where
' Python switches algorithm and takes the earliest in the shorter string.
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
