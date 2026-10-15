Attribute VB_Name = "modDamerauLevenshtein"
'===============================================================================
' Damerau-Levenshtein distance - port of textdistance.DamerauLevenshtein
' (Python textdistance 4.6)
'
' Like Levenshtein, but swapping two adjacent characters also counts as one
' edit, e.g. DamerauLevenshtein_Distance("abcd", "acbd") = 1 (Levenshtein: 2).
'
' Pure VBA with no host references, so it works in Excel, Access, Word, etc.
'
' Common arguments
'   s1, s2         Strings to compare.
'   CaseSensitive  True (default) compares exactly; False ignores case.
'                  (VBA addition - Python's textdistance has no such flag.)
'   Restricted     True (default) is the "optimal string alignment" variant:
'                  no substring is edited more than once.
'                  False allows it, so "ca" -> "abc" is 2 (ca -> ac -> abc)
'                  instead of 3. Matches textdistance's restricted argument.
'===============================================================================
Option Explicit
Option Compare Binary

' Minimum number of edits, counting adjacent transpositions as one edit.
Public Function DamerauLevenshtein_Distance(ByVal s1 As String, ByVal s2 As String, _
                                            Optional ByVal CaseSensitive As Boolean = True, _
                                            Optional ByVal Restricted As Boolean = True) As Long
    Dim len1 As Long, len2 As Long
    Dim c1() As Long, c2() As Long

    If Not CaseSensitive Then
        s1 = LCase$(s1)
        s2 = LCase$(s2)
    End If

    If s1 = s2 Then Exit Function

    len1 = Len(s1)
    len2 = Len(s2)
    If len1 = 0 Or len2 = 0 Then
        DamerauLevenshtein_Distance = IIf(len1 > len2, len1, len2)
        Exit Function
    End If

    CharCodes s1, c1
    CharCodes s2, c2
    If Restricted Then
        DamerauLevenshtein_Distance = RestrictedDistance(c1, c2, len1, len2)
    Else
        DamerauLevenshtein_Distance = UnrestrictedDistance(c1, c2, len1, len2)
    End If
End Function

' Largest possible distance: the length of the longer string.
Public Function DamerauLevenshtein_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    DamerauLevenshtein_Maximum = IIf(Len(s1) > Len(s2), Len(s1), Len(s2))
End Function

' Maximum - Distance.
Public Function DamerauLevenshtein_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                              Optional ByVal CaseSensitive As Boolean = True, _
                                              Optional ByVal Restricted As Boolean = True) As Long
    DamerauLevenshtein_Similarity = DamerauLevenshtein_Maximum(s1, s2) _
                                  - DamerauLevenshtein_Distance(s1, s2, CaseSensitive, Restricted)
End Function

' Distance scaled to 0 (identical) .. 1 (completely different).
' Two empty strings give 0.
Public Function DamerauLevenshtein_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                                      Optional ByVal CaseSensitive As Boolean = True, _
                                                      Optional ByVal Restricted As Boolean = True) As Double
    Dim maximum As Long
    maximum = DamerauLevenshtein_Maximum(s1, s2)
    If maximum = 0 Then Exit Function
    DamerauLevenshtein_NormalizedDistance = DamerauLevenshtein_Distance(s1, s2, CaseSensitive, Restricted) / maximum
End Function

' Similarity scaled to 0 (completely different) .. 1 (identical).
' Two empty strings give 1.
Public Function DamerauLevenshtein_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                                        Optional ByVal CaseSensitive As Boolean = True, _
                                                        Optional ByVal Restricted As Boolean = True) As Double
    DamerauLevenshtein_NormalizedSimilarity = 1 - DamerauLevenshtein_NormalizedDistance(s1, s2, CaseSensitive, Restricted)
End Function

' Optimal string alignment. c1/c2 are 0-based; d(i, j) is the distance between
' the first i+1 characters of s1 and the first j+1 characters of s2.
Private Function RestrictedDistance(c1() As Long, c2() As Long, ByVal len1 As Long, ByVal len2 As Long) As Long
    Dim d() As Long
    Dim i As Long, j As Long, cost As Long, best As Long

    ReDim d(-1 To len1 - 1, -1 To len2 - 1)
    For i = -1 To len1 - 1
        d(i, -1) = i + 1
    Next i
    For j = -1 To len2 - 1
        d(-1, j) = j + 1
    Next j

    For i = 0 To len1 - 1
        For j = 0 To len2 - 1
            If c1(i) = c2(j) Then cost = 0 Else cost = 1

            best = d(i - 1, j) + 1                                      ' deletion
            If d(i, j - 1) + 1 < best Then best = d(i, j - 1) + 1       ' insertion
            If d(i - 1, j - 1) + cost < best Then best = d(i - 1, j - 1) + cost ' substitution

            ' transposition
            If i > 0 And j > 0 Then
                If c1(i) = c2(j - 1) And c1(i - 1) = c2(j) Then
                    If d(i - 2, j - 2) + cost < best Then best = d(i - 2, j - 2) + cost
                End If
            End If

            d(i, j) = best
        Next j
    Next i

    RestrictedDistance = d(len1 - 1, len2 - 1)
End Function

' True Damerau-Levenshtein (Lowrance-Wagner). c1/c2 are 0-based.
Private Function UnrestrictedDistance(c1() As Long, c2() As Long, ByVal len1 As Long, ByVal len2 As Long) As Long
    Dim d() As Long
    Dim lastRow() As Long      ' last row (1-based) where each character code was seen in s1
    Dim i As Long, j As Long, i1 As Long, j1 As Long, db As Long
    Dim cost As Long, best As Long, maxDist As Long, v As Long

    maxDist = len1 + len2
    ReDim d(-1 To len1, -1 To len2)
    ReDim lastRow(0 To 65535)

    d(-1, -1) = maxDist
    For i = 0 To len1
        d(i, -1) = maxDist
        d(i, 0) = i
    Next i
    For j = 0 To len2
        d(-1, j) = maxDist
        d(0, j) = j
    Next j

    For i = 1 To len1
        db = 0
        For j = 1 To len2
            i1 = lastRow(c2(j - 1))
            j1 = db
            If c1(i - 1) = c2(j - 1) Then
                cost = 0
                db = j
            Else
                cost = 1
            End If

            best = d(i - 1, j - 1) + cost                               ' substitution
            If d(i, j - 1) + 1 < best Then best = d(i, j - 1) + 1       ' insertion
            If d(i - 1, j) + 1 < best Then best = d(i - 1, j) + 1       ' deletion
            v = d(i1 - 1, j1 - 1) + (i - i1) - 1 + (j - j1)             ' transposition
            If v < best Then best = v

            d(i, j) = best
        Next j
        lastRow(c1(i - 1)) = i
    Next i

    UnrestrictedDistance = d(len1, len2)
End Function

' Character codes of s, 0-based, as unsigned values 0..65535.
Private Sub CharCodes(ByVal s As String, codes() As Long)
    Dim i As Long
    ReDim codes(0 To Len(s) - 1)
    For i = 1 To Len(s)
        codes(i - 1) = AscW(Mid$(s, i, 1)) And &HFFFF&
    Next i
End Sub
