Attribute VB_Name = "modTokenBased"
'===============================================================================
' Token-based measures - port of textdistance's Jaccard, Sorensen, Tversky,
' Overlap, Tanimoto, Cosine, MongeElkan and Bag (Python textdistance 4.6)
'
' Each string is split into tokens (characters, character n-grams or words)
' and the measures compare the two bags of tokens:
'
'   Jaccard     shared / all tokens                      0 .. 1
'   Sorensen    2 * shared / (tokens1 + tokens2)         0 .. 1  (Sorensen-Dice)
'   Tversky     shared / (shared + Alpha*only1 + Beta*only2)     0 .. 1
'   Overlap     shared / tokens in the smaller bag       0 .. 1
'   Tanimoto    Log2(Jaccard)                            -inf .. 0
'   Cosine      shared / Sqr(tokens1 * tokens2)          0 .. 1  (Ochiai)
'   MongeElkan  average best Damerau-Levenshtein similarity of each token
'   Bag         distance: tokens left over in the bigger bag
'
' For all but Bag, similarity is the primary value. "Shared" counts a token as
' often as it appears in both strings (or once each with AsSet).
'
' Pure VBA with no host references, so it works in Excel, Access, Word, etc.
'
' Common arguments
'   s1, s2         Strings to compare.
'   CaseSensitive  True (default) compares exactly; False ignores case.
'                  (VBA addition - Python's textdistance has no such flag.)
'   Qval           What a token is: 1 (default) = single characters,
'                  2 or more = overlapping groups of that many characters,
'                  0 = words separated by whitespace (like Python's split()).
'   AsSet          False (default) counts repeated tokens; True counts each
'                  distinct token once.
'
' Differences from Python
'   - Where Python raises ZeroDivisionError (a non-empty string with no tokens,
'     e.g. "a" with Qval 2) the similarity is 0.
'   - VBA has no infinity: where Python returns -inf / inf (Tanimoto with
'     nothing shared, MongeElkan with an empty token list) these functions
'     return NEG_INF / -NEG_INF, the most negative / positive Double.
'===============================================================================
Option Explicit
Option Compare Binary

Public Const NEG_INF As Double = -1.79769313486231E+308

'===============================================================================
' Jaccard: shared / union
'===============================================================================
Public Function Jaccard_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                   Optional ByVal CaseSensitive As Boolean = True, _
                                   Optional ByVal Qval As Long = 1, _
                                   Optional ByVal AsSet As Boolean = False) As Double
    Dim count1 As Long, count2 As Long, common As Long
    Select Case QuickAnswer(s1, s2, CaseSensitive)
        Case 1: Jaccard_Similarity = 1
        Case -1
            CountTokens s1, s2, Qval, AsSet, count1, count2, common
            If count1 + count2 - common > 0 Then
                Jaccard_Similarity = common / (count1 + count2 - common)
            End If
    End Select
End Function

Public Function Jaccard_Distance(ByVal s1 As String, ByVal s2 As String, _
                                 Optional ByVal CaseSensitive As Boolean = True, _
                                 Optional ByVal Qval As Long = 1, _
                                 Optional ByVal AsSet As Boolean = False) As Double
    Jaccard_Distance = 1 - Jaccard_Similarity(s1, s2, CaseSensitive, Qval, AsSet)
End Function

Public Function Jaccard_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                           Optional ByVal CaseSensitive As Boolean = True, _
                                           Optional ByVal Qval As Long = 1, _
                                           Optional ByVal AsSet As Boolean = False) As Double
    Jaccard_NormalizedDistance = Jaccard_Distance(s1, s2, CaseSensitive, Qval, AsSet)
End Function

Public Function Jaccard_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                             Optional ByVal CaseSensitive As Boolean = True, _
                                             Optional ByVal Qval As Long = 1, _
                                             Optional ByVal AsSet As Boolean = False) As Double
    Jaccard_NormalizedSimilarity = 1 - Jaccard_NormalizedDistance(s1, s2, CaseSensitive, Qval, AsSet)
End Function

' Always 1.
Public Function Jaccard_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    Jaccard_Maximum = 1
End Function

'===============================================================================
' Sorensen (Sorensen-Dice): 2 * shared / (tokens1 + tokens2)
'===============================================================================
Public Function Sorensen_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                    Optional ByVal CaseSensitive As Boolean = True, _
                                    Optional ByVal Qval As Long = 1, _
                                    Optional ByVal AsSet As Boolean = False) As Double
    Dim count1 As Long, count2 As Long, common As Long
    Select Case QuickAnswer(s1, s2, CaseSensitive)
        Case 1: Sorensen_Similarity = 1
        Case -1
            CountTokens s1, s2, Qval, AsSet, count1, count2, common
            If count1 + count2 > 0 Then Sorensen_Similarity = 2 * common / (count1 + count2)
    End Select
End Function

Public Function Sorensen_Distance(ByVal s1 As String, ByVal s2 As String, _
                                  Optional ByVal CaseSensitive As Boolean = True, _
                                  Optional ByVal Qval As Long = 1, _
                                  Optional ByVal AsSet As Boolean = False) As Double
    Sorensen_Distance = 1 - Sorensen_Similarity(s1, s2, CaseSensitive, Qval, AsSet)
End Function

Public Function Sorensen_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                            Optional ByVal CaseSensitive As Boolean = True, _
                                            Optional ByVal Qval As Long = 1, _
                                            Optional ByVal AsSet As Boolean = False) As Double
    Sorensen_NormalizedDistance = Sorensen_Distance(s1, s2, CaseSensitive, Qval, AsSet)
End Function

Public Function Sorensen_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                              Optional ByVal CaseSensitive As Boolean = True, _
                                              Optional ByVal Qval As Long = 1, _
                                              Optional ByVal AsSet As Boolean = False) As Double
    Sorensen_NormalizedSimilarity = 1 - Sorensen_NormalizedDistance(s1, s2, CaseSensitive, Qval, AsSet)
End Function

' Always 1.
Public Function Sorensen_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    Sorensen_Maximum = 1
End Function

'===============================================================================
' Tversky: shared / (shared + Alpha * only-in-s1 + Beta * only-in-s2)
'
'   Alpha, Beta    Weights (default 1, 1 = Jaccard; 0.5, 0.5 = Sorensen).
'   Bias           Optional. When given, uses the symmetric variant:
'                  (shared + Bias) / (Alpha*Beta*(small - big) + Beta*big + shared + Bias)
'                  where small / big are the smaller / bigger token counts.
'===============================================================================
Public Function Tversky_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                   Optional ByVal CaseSensitive As Boolean = True, _
                                   Optional ByVal Qval As Long = 1, _
                                   Optional ByVal AsSet As Boolean = False, _
                                   Optional ByVal Alpha As Double = 1, _
                                   Optional ByVal Beta As Double = 1, _
                                   Optional ByVal Bias As Variant) As Double
    Dim count1 As Long, count2 As Long, common As Long
    Dim small As Double, big As Double, numerator As Double, denominator As Double

    Select Case QuickAnswer(s1, s2, CaseSensitive)
        Case 1: Tversky_Similarity = 1
        Case -1
            CountTokens s1, s2, Qval, AsSet, count1, count2, common
            If IsMissing(Bias) Then
                numerator = common
                denominator = common + Alpha * (count1 - common) + Beta * (count2 - common)
            Else
                small = IIf(count1 < count2, count1, count2)
                big = IIf(count1 > count2, count1, count2)
                numerator = common + CDbl(Bias)
                denominator = Alpha * Beta * (small - big) + big * Beta + numerator
            End If
            If denominator <> 0 Then Tversky_Similarity = numerator / denominator
    End Select
End Function

Public Function Tversky_Distance(ByVal s1 As String, ByVal s2 As String, _
                                 Optional ByVal CaseSensitive As Boolean = True, _
                                 Optional ByVal Qval As Long = 1, _
                                 Optional ByVal AsSet As Boolean = False, _
                                 Optional ByVal Alpha As Double = 1, _
                                 Optional ByVal Beta As Double = 1, _
                                 Optional ByVal Bias As Variant) As Double
    Tversky_Distance = 1 - Tversky_Similarity(s1, s2, CaseSensitive, Qval, AsSet, Alpha, Beta, Bias)
End Function

Public Function Tversky_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                           Optional ByVal CaseSensitive As Boolean = True, _
                                           Optional ByVal Qval As Long = 1, _
                                           Optional ByVal AsSet As Boolean = False, _
                                           Optional ByVal Alpha As Double = 1, _
                                           Optional ByVal Beta As Double = 1, _
                                           Optional ByVal Bias As Variant) As Double
    Tversky_NormalizedDistance = Tversky_Distance(s1, s2, CaseSensitive, Qval, AsSet, Alpha, Beta, Bias)
End Function

Public Function Tversky_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                             Optional ByVal CaseSensitive As Boolean = True, _
                                             Optional ByVal Qval As Long = 1, _
                                             Optional ByVal AsSet As Boolean = False, _
                                             Optional ByVal Alpha As Double = 1, _
                                             Optional ByVal Beta As Double = 1, _
                                             Optional ByVal Bias As Variant) As Double
    Tversky_NormalizedSimilarity = 1 - Tversky_NormalizedDistance(s1, s2, CaseSensitive, Qval, AsSet, Alpha, Beta, Bias)
End Function

' Always 1.
Public Function Tversky_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    Tversky_Maximum = 1
End Function

'===============================================================================
' Overlap: shared / tokens in the smaller bag
'===============================================================================
Public Function Overlap_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                   Optional ByVal CaseSensitive As Boolean = True, _
                                   Optional ByVal Qval As Long = 1, _
                                   Optional ByVal AsSet As Boolean = False) As Double
    Dim count1 As Long, count2 As Long, common As Long, smaller As Long
    Select Case QuickAnswer(s1, s2, CaseSensitive)
        Case 1: Overlap_Similarity = 1
        Case -1
            CountTokens s1, s2, Qval, AsSet, count1, count2, common
            smaller = IIf(count1 < count2, count1, count2)
            If smaller > 0 Then Overlap_Similarity = common / smaller
    End Select
End Function

Public Function Overlap_Distance(ByVal s1 As String, ByVal s2 As String, _
                                 Optional ByVal CaseSensitive As Boolean = True, _
                                 Optional ByVal Qval As Long = 1, _
                                 Optional ByVal AsSet As Boolean = False) As Double
    Overlap_Distance = 1 - Overlap_Similarity(s1, s2, CaseSensitive, Qval, AsSet)
End Function

Public Function Overlap_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                           Optional ByVal CaseSensitive As Boolean = True, _
                                           Optional ByVal Qval As Long = 1, _
                                           Optional ByVal AsSet As Boolean = False) As Double
    Overlap_NormalizedDistance = Overlap_Distance(s1, s2, CaseSensitive, Qval, AsSet)
End Function

Public Function Overlap_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                             Optional ByVal CaseSensitive As Boolean = True, _
                                             Optional ByVal Qval As Long = 1, _
                                             Optional ByVal AsSet As Boolean = False) As Double
    Overlap_NormalizedSimilarity = 1 - Overlap_NormalizedDistance(s1, s2, CaseSensitive, Qval, AsSet)
End Function

' Always 1.
Public Function Overlap_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    Overlap_Maximum = 1
End Function

'===============================================================================
' Tanimoto: Log2(Jaccard), from NEG_INF (nothing shared) to 0 (identical).
' Maximum is 1 as in textdistance, so Distance = 1 - Similarity >= 1.
'===============================================================================
Public Function Tanimoto_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                    Optional ByVal CaseSensitive As Boolean = True, _
                                    Optional ByVal Qval As Long = 1, _
                                    Optional ByVal AsSet As Boolean = False) As Double
    Dim j As Double
    j = Jaccard_Similarity(s1, s2, CaseSensitive, Qval, AsSet)
    If j = 0 Then
        Tanimoto_Similarity = NEG_INF
    Else
        Tanimoto_Similarity = Log(j) / Log(2#)
    End If
End Function

Public Function Tanimoto_Distance(ByVal s1 As String, ByVal s2 As String, _
                                  Optional ByVal CaseSensitive As Boolean = True, _
                                  Optional ByVal Qval As Long = 1, _
                                  Optional ByVal AsSet As Boolean = False) As Double
    Tanimoto_Distance = OneMinus(Tanimoto_Similarity(s1, s2, CaseSensitive, Qval, AsSet))
End Function

Public Function Tanimoto_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                            Optional ByVal CaseSensitive As Boolean = True, _
                                            Optional ByVal Qval As Long = 1, _
                                            Optional ByVal AsSet As Boolean = False) As Double
    Tanimoto_NormalizedDistance = Tanimoto_Distance(s1, s2, CaseSensitive, Qval, AsSet)
End Function

Public Function Tanimoto_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                              Optional ByVal CaseSensitive As Boolean = True, _
                                              Optional ByVal Qval As Long = 1, _
                                              Optional ByVal AsSet As Boolean = False) As Double
    Tanimoto_NormalizedSimilarity = OneMinus(Tanimoto_NormalizedDistance(s1, s2, CaseSensitive, Qval, AsSet))
End Function

' Always 1.
Public Function Tanimoto_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    Tanimoto_Maximum = 1
End Function

'===============================================================================
' Cosine (Ochiai): shared / Sqr(tokens1 * tokens2)
'===============================================================================
Public Function Cosine_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                  Optional ByVal CaseSensitive As Boolean = True, _
                                  Optional ByVal Qval As Long = 1, _
                                  Optional ByVal AsSet As Boolean = False) As Double
    Dim count1 As Long, count2 As Long, common As Long
    Select Case QuickAnswer(s1, s2, CaseSensitive)
        Case 1: Cosine_Similarity = 1
        Case -1
            CountTokens s1, s2, Qval, AsSet, count1, count2, common
            If count1 > 0 And count2 > 0 Then Cosine_Similarity = common / Sqr(CDbl(count1) * count2)
    End Select
End Function

Public Function Cosine_Distance(ByVal s1 As String, ByVal s2 As String, _
                                Optional ByVal CaseSensitive As Boolean = True, _
                                Optional ByVal Qval As Long = 1, _
                                Optional ByVal AsSet As Boolean = False) As Double
    Cosine_Distance = 1 - Cosine_Similarity(s1, s2, CaseSensitive, Qval, AsSet)
End Function

Public Function Cosine_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                          Optional ByVal CaseSensitive As Boolean = True, _
                                          Optional ByVal Qval As Long = 1, _
                                          Optional ByVal AsSet As Boolean = False) As Double
    Cosine_NormalizedDistance = Cosine_Distance(s1, s2, CaseSensitive, Qval, AsSet)
End Function

Public Function Cosine_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                            Optional ByVal CaseSensitive As Boolean = True, _
                                            Optional ByVal Qval As Long = 1, _
                                            Optional ByVal AsSet As Boolean = False) As Double
    Cosine_NormalizedSimilarity = 1 - Cosine_NormalizedDistance(s1, s2, CaseSensitive, Qval, AsSet)
End Function

' Always 1.
Public Function Cosine_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    Cosine_Maximum = 1
End Function

'===============================================================================
' MongeElkan: for each token of s1, the best restricted Damerau-Levenshtein
' similarity (longer length - distance) to any token of s2, summed and divided
' by the number of s1 tokens squared - exactly as textdistance computes it.
' The result is not limited to 1 when comparing words, and the maximum is
' always 2 (textdistance's MongeElkan.maximum).
'
'   Symmetric      False (default). True averages the s1->s2 and s2->s1 scores.
'===============================================================================
Public Function MongeElkan_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                      Optional ByVal CaseSensitive As Boolean = True, _
                                      Optional ByVal Qval As Long = 1, _
                                      Optional ByVal Symmetric As Boolean = False) As Double
    Dim t1() As String, t2() As String, n1 As Long, n2 As Long
    Dim forward As Double, backward As Double

    Select Case QuickAnswer(s1, s2, CaseSensitive)
        Case 1: MongeElkan_Similarity = 2     ' textdistance returns the maximum
        Case -1
            n1 = Tokenize(s1, Qval, t1)
            n2 = Tokenize(s2, Qval, t2)
            forward = MongeElkanCalc(t1, n1, t2, n2)
            If Not Symmetric Then
                MongeElkan_Similarity = forward
            Else
                backward = MongeElkanCalc(t2, n2, t1, n1)
                If forward = NEG_INF Or backward = NEG_INF Then
                    MongeElkan_Similarity = NEG_INF
                Else
                    MongeElkan_Similarity = (forward + backward) / 2
                End If
            End If
    End Select
End Function

' Always 2.
Public Function MongeElkan_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    MongeElkan_Maximum = 2
End Function

' Maximum - Similarity.
Public Function MongeElkan_Distance(ByVal s1 As String, ByVal s2 As String, _
                                    Optional ByVal CaseSensitive As Boolean = True, _
                                    Optional ByVal Qval As Long = 1, _
                                    Optional ByVal Symmetric As Boolean = False) As Double
    Dim sim As Double
    sim = MongeElkan_Similarity(s1, s2, CaseSensitive, Qval, Symmetric)
    If sim = NEG_INF Then
        MongeElkan_Distance = -NEG_INF
    Else
        MongeElkan_Distance = 2 - sim
    End If
End Function

' Distance / 2.
Public Function MongeElkan_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                              Optional ByVal CaseSensitive As Boolean = True, _
                                              Optional ByVal Qval As Long = 1, _
                                              Optional ByVal Symmetric As Boolean = False) As Double
    Dim d As Double
    d = MongeElkan_Distance(s1, s2, CaseSensitive, Qval, Symmetric)
    If d = -NEG_INF Then
        MongeElkan_NormalizedDistance = d
    Else
        MongeElkan_NormalizedDistance = d / 2
    End If
End Function

Public Function MongeElkan_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                                Optional ByVal CaseSensitive As Boolean = True, _
                                                Optional ByVal Qval As Long = 1, _
                                                Optional ByVal Symmetric As Boolean = False) As Double
    MongeElkan_NormalizedSimilarity = OneMinus(MongeElkan_NormalizedDistance(s1, s2, CaseSensitive, Qval, Symmetric))
End Function

'===============================================================================
' Bag distance: the number of tokens left over in the bigger bag after
' removing the shared ones. The maximum is the length of the longer string in
' characters (as in textdistance, even when Qval is not 1).
'===============================================================================
Public Function Bag_Distance(ByVal s1 As String, ByVal s2 As String, _
                             Optional ByVal CaseSensitive As Boolean = True, _
                             Optional ByVal Qval As Long = 1) As Long
    Dim count1 As Long, count2 As Long, common As Long
    If Not CaseSensitive Then
        s1 = LCase$(s1)
        s2 = LCase$(s2)
    End If
    CountTokens s1, s2, Qval, False, count1, count2, common
    Bag_Distance = IIf(count1 > count2, count1, count2) - common
End Function

' Length of the longer string.
Public Function Bag_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    Bag_Maximum = IIf(Len(s1) > Len(s2), Len(s1), Len(s2))
End Function

' Maximum - Distance.
Public Function Bag_Similarity(ByVal s1 As String, ByVal s2 As String, _
                               Optional ByVal CaseSensitive As Boolean = True, _
                               Optional ByVal Qval As Long = 1) As Long
    Bag_Similarity = Bag_Maximum(s1, s2) - Bag_Distance(s1, s2, CaseSensitive, Qval)
End Function

' Distance / Maximum; 0 when both strings are empty.
Public Function Bag_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                       Optional ByVal CaseSensitive As Boolean = True, _
                                       Optional ByVal Qval As Long = 1) As Double
    If Bag_Maximum(s1, s2) = 0 Then Exit Function
    Bag_NormalizedDistance = Bag_Distance(s1, s2, CaseSensitive, Qval) / Bag_Maximum(s1, s2)
End Function

Public Function Bag_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                         Optional ByVal CaseSensitive As Boolean = True, _
                                         Optional ByVal Qval As Long = 1) As Double
    Bag_NormalizedSimilarity = 1 - Bag_NormalizedDistance(s1, s2, CaseSensitive, Qval)
End Function

'===============================================================================
' Shared helpers
'===============================================================================

' textdistance's quick_answer for similarities, after applying CaseSensitive
' (s1 and s2 are changed in place): 1 = identical, 0 = one is empty,
' -1 = needs the full calculation.
Private Function QuickAnswer(s1 As String, s2 As String, ByVal CaseSensitive As Boolean) As Long
    If Not CaseSensitive Then
        s1 = LCase$(s1)
        s2 = LCase$(s2)
    End If
    If s1 = s2 Then
        QuickAnswer = 1
    ElseIf Len(s1) = 0 Or Len(s2) = 0 Then
        QuickAnswer = 0
    Else
        QuickAnswer = -1
    End If
End Function

' 1 - x, keeping the NEG_INF / -NEG_INF stand-ins for infinity.
Private Function OneMinus(ByVal x As Double) As Double
    If x = NEG_INF Then
        OneMinus = -NEG_INF
    ElseIf x = -NEG_INF Then
        OneMinus = NEG_INF
    Else
        OneMinus = 1 - x
    End If
End Function

' Token counts of s1 and s2 and how many they share (as multisets, or as sets
' of distinct tokens with AsSet).
Private Sub CountTokens(ByVal s1 As String, ByVal s2 As String, ByVal Qval As Long, ByVal AsSet As Boolean, _
                        ByRef count1 As Long, ByRef count2 As Long, ByRef common As Long)
    Dim t1() As String, t2() As String, n1 As Long, n2 As Long
    Dim i As Long, j As Long, c As Long

    n1 = Tokenize(s1, Qval, t1)
    n2 = Tokenize(s2, Qval, t2)
    SortStrings t1, 1, n1
    SortStrings t2, 1, n2
    count1 = CountSorted(t1, n1, AsSet)
    count2 = CountSorted(t2, n2, AsSet)

    ' Walk both sorted lists: each shared token counts min(count in s1, count in s2)
    ' times, or once with AsSet
    common = 0
    i = 1
    j = 1
    Do While i <= n1 And j <= n2
        c = StrComp(t1(i), t2(j), vbBinaryCompare)
        If c = 0 Then
            If Not AsSet Then
                common = common + 1
            ElseIf i = 1 Then
                common = common + 1
            ElseIf t1(i) <> t1(i - 1) Then
                common = common + 1
            End If
            i = i + 1
            j = j + 1
        ElseIf c < 0 Then
            i = i + 1
        Else
            j = j + 1
        End If
    Loop
End Sub

' Splits s into tokens (1-based in tokens), returning how many there are.
Private Function Tokenize(ByVal s As String, ByVal Qval As Long, tokens() As String) As Long
    Dim i As Long, n As Long, start As Long, atSpace As Boolean

    ReDim tokens(1 To IIf(Len(s) > 0, Len(s), 1))
    If Qval <= 0 Then
        ' Words: runs of non-whitespace, like Python's str.split()
        ' (i = Len(s) + 1 acts as a final whitespace to close the last word)
        For i = 1 To Len(s) + 1
            If i > Len(s) Then
                atSpace = True
            Else
                atSpace = IsSpace(Mid$(s, i, 1))
            End If
            If atSpace Then
                If start > 0 Then
                    n = n + 1
                    tokens(n) = Mid$(s, start, i - start)
                End If
                start = 0
            ElseIf start = 0 Then
                start = i
            End If
        Next i
    Else
        ' Overlapping groups of Qval characters
        For i = 1 To Len(s) - Qval + 1
            n = n + 1
            tokens(n) = Mid$(s, i, Qval)
        Next i
    End If
    Tokenize = n
End Function

' Characters Python's str.split() treats as whitespace.
Private Function IsSpace(ByVal ch As String) As Boolean
    Dim code As Long
    code = AscW(ch) And &HFFFF&
    Select Case code
        Case 9 To 13, 28 To 32, &H85, &HA0, &H1680, &H2000 To &H200A, &H2028, &H2029, &H202F, &H205F, &H3000
            IsSpace = True
    End Select
End Function

' Number of tokens, or of distinct tokens with AsSet (tokens must be sorted).
Private Function CountSorted(tokens() As String, ByVal n As Long, ByVal AsSet As Boolean) As Long
    Dim i As Long
    If Not AsSet Then
        CountSorted = n
        Exit Function
    End If
    For i = 1 To n
        If i = 1 Then
            CountSorted = 1
        ElseIf tokens(i) <> tokens(i - 1) Then
            CountSorted = CountSorted + 1
        End If
    Next i
End Function

' In-place quicksort of a(lo..hi) by binary string comparison.
Private Sub SortStrings(a() As String, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, pivot As String, tmp As String
    If lo >= hi Then Exit Sub
    pivot = a((lo + hi) \ 2)
    i = lo
    j = hi
    Do While i <= j
        Do While StrComp(a(i), pivot, vbBinaryCompare) < 0
            i = i + 1
        Loop
        Do While StrComp(a(j), pivot, vbBinaryCompare) > 0
            j = j - 1
        Loop
        If i <= j Then
            tmp = a(i): a(i) = a(j): a(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop
    If lo < j Then SortStrings a, lo, j
    If i < hi Then SortStrings a, i, hi
End Sub

' textdistance's MongeElkan._calc for two token lists.
Private Function MongeElkanCalc(a() As String, ByVal na As Long, b() As String, ByVal nb As Long) As Double
    Dim i As Long, j As Long, best As Long, sim As Long, total As Double

    If na = 0 Then Exit Function
    If nb = 0 Then
        ' No tokens to compare with: Python's max over nothing stays -inf
        MongeElkanCalc = NEG_INF
        Exit Function
    End If
    For i = 1 To na
        best = DLSimilarity(a(i), b(1))
        For j = 2 To nb
            sim = DLSimilarity(a(i), b(j))
            If sim > best Then best = sim
        Next j
        total = total + best
    Next i
    MongeElkanCalc = total / na / na
End Function

' Restricted Damerau-Levenshtein similarity: longer length - distance.
Private Function DLSimilarity(ByVal s1 As String, ByVal s2 As String) As Long
    Dim len1 As Long, len2 As Long, i As Long, j As Long, cost As Long, best As Long
    Dim d() As Long

    len1 = Len(s1)
    len2 = Len(s2)
    DLSimilarity = IIf(len1 > len2, len1, len2)
    If s1 = s2 Then Exit Function

    ' d(i, j) = distance between the first i+1 chars of s1 and first j+1 of s2
    ReDim d(-1 To len1 - 1, -1 To len2 - 1)
    For i = -1 To len1 - 1
        d(i, -1) = i + 1
    Next i
    For j = -1 To len2 - 1
        d(-1, j) = j + 1
    Next j
    For i = 0 To len1 - 1
        For j = 0 To len2 - 1
            If Mid$(s1, i + 1, 1) = Mid$(s2, j + 1, 1) Then cost = 0 Else cost = 1
            best = d(i - 1, j) + 1
            If d(i, j - 1) + 1 < best Then best = d(i, j - 1) + 1
            If d(i - 1, j - 1) + cost < best Then best = d(i - 1, j - 1) + cost
            If i > 0 And j > 0 Then
                If Mid$(s1, i + 1, 1) = Mid$(s2, j, 1) And Mid$(s1, i, 1) = Mid$(s2, j + 1, 1) Then
                    If d(i - 2, j - 2) + cost < best Then best = d(i - 2, j - 2) + cost
                End If
            End If
            d(i, j) = best
        Next j
    Next i

    DLSimilarity = DLSimilarity - d(len1 - 1, len2 - 1)
End Function
