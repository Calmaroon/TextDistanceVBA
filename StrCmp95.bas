Attribute VB_Name = "modStrCmp95"
'===============================================================================
' StrCmp95 similarity - port of textdistance.StrCmp95 (Python textdistance 4.6)
'
' The U.S. Census Bureau's string comparator: Jaro-Winkler with extra partial
' credit for characters often confused by people or OCR (O/0, I/1, M/N, ...),
' from 0 (different) to 1 (identical):
'   StrCmp95_Similarity("MARTHA", "MARHTA") = 0.961
'
' Leading/trailing whitespace is ignored and the comparison is always
' case-insensitive (there is no CaseSensitive argument).
' Note: VBA's UCase$ leaves a few characters alone that Python changes (e.g.
' German sharp s), and only 0-9 count as digits.
'
'   LongStrings    False (default). True applies the extra adjustment for long
'                  strings (textdistance's long_strings argument).
'
' Pure VBA with no host references, so it works in Excel, Access, Word, etc.
'===============================================================================
Option Explicit
Option Compare Binary

' Character pairs that earn partial credit (in both directions)
Private Const SIMILAR_PAIRS As String = _
    "|AE|AI|AO|AU|BV|EI|EO|EU|IO|IU|OU|IY|EY|CG|EF|WU|WV|XK|SZ|XS|QC|UV|MN|LI|" & _
    "QO|PR|IJ|2Z|5S|8B|1I|1L|0O|0Q|CK|GJ|"

Public Function StrCmp95_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                    Optional ByVal LongStrings As Boolean = False) As Double
    Dim len1 As Long, len2 As Long, minv As Long, searchRange As Long
    Dim flags1() As Long, flags2() As Long
    Dim i As Long, j As Long, k As Long, lo As Long, hi As Long
    Dim common As Long, trans As Long, similar As Long, prefix As Long
    Dim numSim As Double, weight As Double, c1 As String

    s1 = UCase$(StripSpace(s1))
    s2 = UCase$(StripSpace(s2))

    If s1 = s2 Then
        StrCmp95_Similarity = 1
        Exit Function
    End If
    len1 = Len(s1)
    len2 = Len(s2)
    If len1 = 0 Or len2 = 0 Then Exit Function

    If len1 > len2 Then
        searchRange = len1
        minv = len2
    Else
        searchRange = len2
        minv = len1
    End If
    ReDim flags1(0 To searchRange - 1)
    ReDim flags2(0 To searchRange - 1)
    searchRange = searchRange \ 2 - 1
    If searchRange < 0 Then searchRange = 0

    ' Count and flag matching characters within the search range
    For i = 0 To len1 - 1
        lo = IIf(i - searchRange > 0, i - searchRange, 0)
        hi = IIf(i + searchRange < len2 - 1, i + searchRange, len2 - 1)
        For j = lo To hi
            If flags2(j) = 0 And Mid$(s2, j + 1, 1) = Mid$(s1, i + 1, 1) Then
                flags2(j) = 1
                flags1(i) = 1
                common = common + 1
                Exit For
            End If
        Next j
    Next i
    If common = 0 Then Exit Function

    ' Count transpositions
    k = 0
    For i = 0 To len1 - 1
        If flags1(i) <> 0 Then
            For j = k To len2 - 1
                If flags2(j) <> 0 Then
                    k = j + 1
                    Exit For
                End If
            Next j
            If Mid$(s1, i + 1, 1) <> Mid$(s2, j + 1, 1) Then trans = trans + 1
        End If
    Next i
    trans = trans \ 2

    ' Partial credit for unmatched characters that are commonly confused
    If minv > common Then
        For i = 0 To len1 - 1
            c1 = Mid$(s1, i + 1, 1)
            If flags1(i) = 0 And InRange(c1) Then
                For j = 0 To len2 - 1
                    If flags2(j) = 0 Then
                        If InRange(Mid$(s2, j + 1, 1)) Then
                            If IsSimilarPair(c1, Mid$(s2, j + 1, 1)) Then
                                similar = similar + 3
                                flags2(j) = 2
                                Exit For
                            End If
                        End If
                    End If
                Next j
            End If
        Next i
    End If
    numSim = similar / 10# + common

    weight = (numSim / len1 + numSim / len2 + (common - trans) / common) / 3#
    If weight <= 0.7 Then GoTo Done

    ' Boost for up to 4 matching (non-digit) leading characters
    Do While prefix < IIf(minv < 4, minv, 4)
        c1 = Mid$(s1, prefix + 1, 1)
        If c1 <> Mid$(s2, prefix + 1, 1) Then Exit Do
        If IsDigitChar(c1) Then Exit Do
        prefix = prefix + 1
    Loop
    If prefix > 0 Then weight = weight + prefix * 0.1 * (1# - weight)

    ' Optional adjustment for long strings
    If Not LongStrings Or minv <= 4 Then GoTo Done
    If common <= prefix + 1 Or 2 * common < minv + prefix Then GoTo Done
    If IsDigitChar(Left$(s1, 1)) Then GoTo Done
    weight = weight + (1# - weight) * (common - prefix - 1) / (len1 + len2 - prefix * 2 + 2)

Done:
    StrCmp95_Similarity = weight
End Function

Public Function StrCmp95_Distance(ByVal s1 As String, ByVal s2 As String, _
                                  Optional ByVal LongStrings As Boolean = False) As Double
    StrCmp95_Distance = 1 - StrCmp95_Similarity(s1, s2, LongStrings)
End Function

Public Function StrCmp95_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                            Optional ByVal LongStrings As Boolean = False) As Double
    StrCmp95_NormalizedDistance = StrCmp95_Distance(s1, s2, LongStrings)
End Function

Public Function StrCmp95_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                              Optional ByVal LongStrings As Boolean = False) As Double
    StrCmp95_NormalizedSimilarity = 1 - StrCmp95_NormalizedDistance(s1, s2, LongStrings)
End Function

' Always 1.
Public Function StrCmp95_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    StrCmp95_Maximum = 1
End Function

' Characters with codes 1..90 (textdistance's _in_range).
Private Function InRange(ByVal ch As String) As Boolean
    Dim code As Long
    code = AscW(ch) And &HFFFF&
    InRange = code > 0 And code < 91
End Function

Private Function IsSimilarPair(ByVal a As String, ByVal b As String) As Boolean
    IsSimilarPair = InStr(1, SIMILAR_PAIRS, "|" & a & b & "|", vbBinaryCompare) > 0 _
                 Or InStr(1, SIMILAR_PAIRS, "|" & b & a & "|", vbBinaryCompare) > 0
End Function

Private Function IsDigitChar(ByVal ch As String) As Boolean
    IsDigitChar = Len(ch) = 1 And ch >= "0" And ch <= "9"
End Function

' Removes leading and trailing whitespace, like Python's str.strip().
Private Function StripSpace(ByVal s As String) As String
    Dim first As Long, last As Long
    first = 1
    last = Len(s)
    Do While first <= last
        If Not IsSpace(Mid$(s, first, 1)) Then Exit Do
        first = first + 1
    Loop
    Do While last >= first
        If Not IsSpace(Mid$(s, last, 1)) Then Exit Do
        last = last - 1
    Loop
    StripSpace = Mid$(s, first, last - first + 1)
End Function

' Characters Python's str.strip() treats as whitespace.
Private Function IsSpace(ByVal ch As String) As Boolean
    Select Case AscW(ch) And &HFFFF&
        Case 9 To 13, 28 To 32, &H85, &HA0, &H1680, &H2000 To &H200A, &H2028, &H2029, &H202F, &H205F, &H3000
            IsSpace = True
    End Select
End Function
