Attribute VB_Name = "modJaroWinkler"
'===============================================================================
' Jaro and Jaro-Winkler similarity - port of textdistance.Jaro and
' textdistance.JaroWinkler (Python textdistance 4.6)
'
' Scores from 0 (nothing in common) to 1 (identical) based on matching
' characters that are near each other and how many of them are out of order.
' Jaro-Winkler boosts the score of strings sharing a prefix (up to 4 chars),
' which makes it popular for matching names:
'   Jaro_Similarity("martha", "marhta")        = 0.944
'   JaroWinkler_Similarity("martha", "marhta") = 0.961
'
' Unlike edit distances, similarity is the primary value here; the maximum is
' always 1, so Distance = 1 - Similarity.
'
' Pure VBA with no host references, so it works in Excel, Access, Word, etc.
'
' Common arguments
'   s1, s2         Strings to compare.
'   CaseSensitive  True (default) compares exactly; False ignores case.
'                  (VBA addition - Python's textdistance has no such flag.)
'   LongTolerance  False (default). True applies the extra adjustment for long
'                  strings (textdistance's long_tolerance argument).
'   PrefixWeight   Jaro-Winkler only: boost per matching prefix character,
'                  default 0.1. Should not exceed 0.25.
'===============================================================================
Option Explicit
Option Compare Binary

'-------------------------------------------------------------------------------
' Jaro
'-------------------------------------------------------------------------------
Public Function Jaro_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                Optional ByVal CaseSensitive As Boolean = True, _
                                Optional ByVal LongTolerance As Boolean = False) As Double
    Jaro_Similarity = JaroCore(s1, s2, CaseSensitive, LongTolerance, False, 0.1)
End Function

Public Function Jaro_Distance(ByVal s1 As String, ByVal s2 As String, _
                              Optional ByVal CaseSensitive As Boolean = True, _
                              Optional ByVal LongTolerance As Boolean = False) As Double
    Jaro_Distance = 1 - Jaro_Similarity(s1, s2, CaseSensitive, LongTolerance)
End Function

Public Function Jaro_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                          Optional ByVal CaseSensitive As Boolean = True, _
                                          Optional ByVal LongTolerance As Boolean = False) As Double
    Jaro_NormalizedSimilarity = 1 - Jaro_NormalizedDistance(s1, s2, CaseSensitive, LongTolerance)
End Function

Public Function Jaro_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                        Optional ByVal CaseSensitive As Boolean = True, _
                                        Optional ByVal LongTolerance As Boolean = False) As Double
    Jaro_NormalizedDistance = Jaro_Distance(s1, s2, CaseSensitive, LongTolerance)
End Function

' Always 1.
Public Function Jaro_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    Jaro_Maximum = 1
End Function

'-------------------------------------------------------------------------------
' Jaro-Winkler
'-------------------------------------------------------------------------------
Public Function JaroWinkler_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                       Optional ByVal CaseSensitive As Boolean = True, _
                                       Optional ByVal LongTolerance As Boolean = False, _
                                       Optional ByVal PrefixWeight As Double = 0.1) As Double
    JaroWinkler_Similarity = JaroCore(s1, s2, CaseSensitive, LongTolerance, True, PrefixWeight)
End Function

Public Function JaroWinkler_Distance(ByVal s1 As String, ByVal s2 As String, _
                                     Optional ByVal CaseSensitive As Boolean = True, _
                                     Optional ByVal LongTolerance As Boolean = False, _
                                     Optional ByVal PrefixWeight As Double = 0.1) As Double
    JaroWinkler_Distance = 1 - JaroWinkler_Similarity(s1, s2, CaseSensitive, LongTolerance, PrefixWeight)
End Function

Public Function JaroWinkler_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                                 Optional ByVal CaseSensitive As Boolean = True, _
                                                 Optional ByVal LongTolerance As Boolean = False, _
                                                 Optional ByVal PrefixWeight As Double = 0.1) As Double
    JaroWinkler_NormalizedSimilarity = 1 - JaroWinkler_NormalizedDistance(s1, s2, CaseSensitive, LongTolerance, PrefixWeight)
End Function

Public Function JaroWinkler_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                               Optional ByVal CaseSensitive As Boolean = True, _
                                               Optional ByVal LongTolerance As Boolean = False, _
                                               Optional ByVal PrefixWeight As Double = 0.1) As Double
    JaroWinkler_NormalizedDistance = JaroWinkler_Distance(s1, s2, CaseSensitive, LongTolerance, PrefixWeight)
End Function

' Always 1.
Public Function JaroWinkler_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    JaroWinkler_Maximum = 1
End Function

'-------------------------------------------------------------------------------
' Shared implementation (textdistance.JaroWinkler.__call__)
'-------------------------------------------------------------------------------
Private Function JaroCore(ByVal s1 As String, ByVal s2 As String, ByVal CaseSensitive As Boolean, _
                          ByVal LongTolerance As Boolean, ByVal Winklerize As Boolean, _
                          ByVal PrefixWeight As Double) As Double
    Dim len1 As Long, len2 As Long, minLen As Long, searchRange As Long
    Dim c1() As Long, c2() As Long
    Dim flags1() As Boolean, flags2() As Boolean
    Dim i As Long, j As Long, k As Long, lo As Long, hi As Long
    Dim common As Long, trans As Long, prefix As Long
    Dim weight As Double

    If Not CaseSensitive Then
        s1 = LCase$(s1)
        s2 = LCase$(s2)
    End If

    If s1 = s2 Then
        JaroCore = 1
        Exit Function
    End If

    len1 = Len(s1)
    len2 = Len(s2)
    If len1 = 0 Or len2 = 0 Then Exit Function

    CharCodes s1, c1
    CharCodes s2, c2
    minLen = IIf(len1 < len2, len1, len2)
    searchRange = IIf(len1 > len2, len1, len2) \ 2 - 1
    If searchRange < 0 Then searchRange = 0

    ReDim flags1(0 To len1 - 1)
    ReDim flags2(0 To len2 - 1)

    ' Count and flag matching characters within the search range
    For i = 0 To len1 - 1
        lo = IIf(i - searchRange > 0, i - searchRange, 0)
        hi = IIf(i + searchRange < len2 - 1, i + searchRange, len2 - 1)
        For j = lo To hi
            If Not flags2(j) And c2(j) = c1(i) Then
                flags1(i) = True
                flags2(j) = True
                common = common + 1
                Exit For
            End If
        Next j
    Next i

    If common = 0 Then Exit Function

    ' Count transpositions: matched characters that appear in a different order
    k = 0
    For i = 0 To len1 - 1
        If flags1(i) Then
            For j = k To len2 - 1
                If flags2(j) Then
                    k = j + 1
                    Exit For
                End If
            Next j
            If c1(i) <> c2(j) Then trans = trans + 1
        End If
    Next i
    trans = trans \ 2

    weight = (common / len1 + common / len2 + (common - trans) / common) / 3

    If Not Winklerize Then GoTo Done
    If weight <= 0.7 Then GoTo Done

    ' Winkler boost for a common prefix of up to 4 characters
    Do While prefix < IIf(minLen < 4, minLen, 4)
        If c1(prefix) <> c2(prefix) Then Exit Do
        prefix = prefix + 1
    Loop
    If prefix > 0 Then weight = weight + prefix * PrefixWeight * (1 - weight)

    ' Optional adjustment for long strings: after the prefix, at least two more
    ' characters must agree and they must be over half of the rest
    If Not LongTolerance Or minLen <= 4 Then GoTo Done
    If common <= prefix + 1 Or 2 * common < minLen + prefix Then GoTo Done
    weight = weight + (1 - weight) * (common - prefix - 1) / (len1 + len2 - prefix * 2 + 2)

Done:
    JaroCore = weight
End Function

' Character codes of s, 0-based.
Private Sub CharCodes(ByVal s As String, codes() As Long)
    Dim i As Long
    ReDim codes(0 To Len(s) - 1)
    For i = 1 To Len(s)
        codes(i - 1) = AscW(Mid$(s, i, 1))
    Next i
End Sub
