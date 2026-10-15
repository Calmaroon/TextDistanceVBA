Attribute VB_Name = "modAlignment"
'===============================================================================
' Alignment scores - port of textdistance.NeedlemanWunsch, Gotoh and
' SmithWaterman (Python textdistance 4.6)
'
' These score the best alignment of the two strings: +1 for each matching
' pair of tokens, 0 for a mismatch, minus a penalty for each gap.
'
'   NeedlemanWunsch  global alignment, gaps cost GapCost
'   Gotoh            global alignment with affine gaps: opening a gap costs
'                    GapOpen, each further gap position costs GapExt
'   SmithWaterman    local alignment (scores never drop below 0)
'
' Similarity is the primary value. As in textdistance:
'   NeedlemanWunsch / Gotoh:  Distance = -Similarity, and the normalized values
'       use Minimum / Maximum (see each function).
'   SmithWaterman:  Distance = Maximum - Similarity.
'   Maximum / Minimum use the strings' lengths in characters, even when Qval
'   is not 1.
'
' Pure VBA with no host references, so it works in Excel, Access, Word, etc.
'
' Common arguments
'   s1, s2         Strings to compare.
'   CaseSensitive  True (default) compares exactly; False ignores case.
'                  (VBA addition - Python's textdistance has no such flag.)
'   Qval           What a token is: 1 (default) = single characters,
'                  2 or more = overlapping groups of that many characters,
'                  0 = words separated by whitespace.
'
' textdistance's custom sim_func is not supported (matches score 1, others 0).
'===============================================================================
Option Explicit
Option Compare Binary

' Stand-in for minus infinity inside the Gotoh tables
Private Const MINUS_INF As Double = -1E+300

'===============================================================================
' Needleman-Wunsch
'===============================================================================
Public Function NeedlemanWunsch_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                           Optional ByVal CaseSensitive As Boolean = True, _
                                           Optional ByVal Qval As Long = 1, _
                                           Optional ByVal GapCost As Double = 1) As Double
    Dim t1() As String, t2() As String, n1 As Long, n2 As Long
    Dim d() As Double, i As Long, j As Long, best As Double, v As Double

    If Not CaseSensitive Then
        s1 = LCase$(s1)
        s2 = LCase$(s2)
    End If
    n1 = Tokenize(s1, Qval, t1)
    n2 = Tokenize(s2, Qval, t2)

    ReDim d(0 To n1, 0 To n2)
    For i = 0 To n1
        d(i, 0) = -(i * GapCost)
    Next i
    For j = 0 To n2
        d(0, j) = -(j * GapCost)
    Next j
    For i = 1 To n1
        For j = 1 To n2
            best = d(i - 1, j - 1) + IIf(t1(i) = t2(j), 1, 0)     ' match / mismatch
            v = d(i - 1, j) - GapCost                             ' delete
            If v > best Then best = v
            v = d(i, j - 1) - GapCost                             ' insert
            If v > best Then best = v
            d(i, j) = best
        Next j
    Next i
    NeedlemanWunsch_Similarity = d(n1, n2)
End Function

' -Similarity.
Public Function NeedlemanWunsch_Distance(ByVal s1 As String, ByVal s2 As String, _
                                         Optional ByVal CaseSensitive As Boolean = True, _
                                         Optional ByVal Qval As Long = 1, _
                                         Optional ByVal GapCost As Double = 1) As Double
    NeedlemanWunsch_Distance = -NeedlemanWunsch_Similarity(s1, s2, CaseSensitive, Qval, GapCost)
End Function

' Length of the longer string.
Public Function NeedlemanWunsch_Maximum(ByVal s1 As String, ByVal s2 As String) As Double
    NeedlemanWunsch_Maximum = MaxLen(s1, s2)
End Function

' -Maximum * GapCost.
Public Function NeedlemanWunsch_Minimum(ByVal s1 As String, ByVal s2 As String, _
                                        Optional ByVal GapCost As Double = 1) As Double
    NeedlemanWunsch_Minimum = -MaxLen(s1, s2) * GapCost
End Function

' (Distance - Minimum) / (Maximum - Minimum); 0 when both strings are empty.
Public Function NeedlemanWunsch_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                                   Optional ByVal CaseSensitive As Boolean = True, _
                                                   Optional ByVal Qval As Long = 1, _
                                                   Optional ByVal GapCost As Double = 1) As Double
    Dim minimum As Double, maximum As Double
    maximum = NeedlemanWunsch_Maximum(s1, s2)
    minimum = NeedlemanWunsch_Minimum(s1, s2, GapCost)
    If maximum = 0 Then Exit Function
    NeedlemanWunsch_NormalizedDistance = (NeedlemanWunsch_Distance(s1, s2, CaseSensitive, Qval, GapCost) - minimum) _
                                         / (maximum - minimum)
End Function

' (Similarity - Minimum) / (Maximum * 2); 1 when both strings are empty.
Public Function NeedlemanWunsch_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                                     Optional ByVal CaseSensitive As Boolean = True, _
                                                     Optional ByVal Qval As Long = 1, _
                                                     Optional ByVal GapCost As Double = 1) As Double
    Dim minimum As Double, maximum As Double
    maximum = NeedlemanWunsch_Maximum(s1, s2)
    minimum = NeedlemanWunsch_Minimum(s1, s2, GapCost)
    If maximum = 0 Then
        NeedlemanWunsch_NormalizedSimilarity = 1
        Exit Function
    End If
    NeedlemanWunsch_NormalizedSimilarity = (NeedlemanWunsch_Similarity(s1, s2, CaseSensitive, Qval, GapCost) - minimum) _
                                           / (maximum * 2)
End Function

'===============================================================================
' Gotoh (Needleman-Wunsch with affine gap penalties)
'
'   GapOpen        Cost of the first position of a gap (default 1).
'   GapExt         Cost of each further position of the same gap (default 0.4).
'
' Python raises an error when exactly one string is empty; this returns the
' cost of a single gap of the other string's length instead.
'===============================================================================
Public Function Gotoh_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                 Optional ByVal CaseSensitive As Boolean = True, _
                                 Optional ByVal Qval As Long = 1, _
                                 Optional ByVal GapOpen As Double = 1, _
                                 Optional ByVal GapExt As Double = 0.4) As Double
    Dim t1() As String, t2() As String, n1 As Long, n2 As Long
    Dim d() As Double, p() As Double, q() As Double
    Dim i As Long, j As Long, sim As Double, best As Double

    If Not CaseSensitive Then
        s1 = LCase$(s1)
        s2 = LCase$(s2)
    End If
    n1 = Tokenize(s1, Qval, t1)
    n2 = Tokenize(s2, Qval, t2)

    ' d: best ending in a match/mismatch; p: ending in a gap in s2; q: gap in s1
    ReDim d(0 To n1, 0 To n2)
    ReDim p(0 To n1, 0 To n2)
    ReDim q(0 To n1, 0 To n2)
    p(0, 0) = MINUS_INF
    q(0, 0) = MINUS_INF
    For i = 1 To n1
        d(i, 0) = MINUS_INF
        p(i, 0) = -GapOpen - GapExt * (i - 1)
        q(i, 0) = MINUS_INF
    Next i
    For j = 1 To n2
        d(0, j) = MINUS_INF
        p(0, j) = MINUS_INF
        q(0, j) = -GapOpen - GapExt * (j - 1)
    Next j

    For i = 1 To n1
        For j = 1 To n2
            sim = IIf(t1(i) = t2(j), 1, 0)
            best = d(i - 1, j - 1)
            If p(i - 1, j - 1) > best Then best = p(i - 1, j - 1)
            If q(i - 1, j - 1) > best Then best = q(i - 1, j - 1)
            d(i, j) = best + sim
            p(i, j) = Larger(d(i - 1, j) - GapOpen, p(i - 1, j) - GapExt)
            q(i, j) = Larger(d(i, j - 1) - GapOpen, q(i, j - 1) - GapExt)
        Next j
    Next i

    Gotoh_Similarity = Larger(Larger(d(n1, n2), p(n1, n2)), q(n1, n2))
End Function

' -Similarity.
Public Function Gotoh_Distance(ByVal s1 As String, ByVal s2 As String, _
                               Optional ByVal CaseSensitive As Boolean = True, _
                               Optional ByVal Qval As Long = 1, _
                               Optional ByVal GapOpen As Double = 1, _
                               Optional ByVal GapExt As Double = 0.4) As Double
    Gotoh_Distance = -Gotoh_Similarity(s1, s2, CaseSensitive, Qval, GapOpen, GapExt)
End Function

' Length of the shorter string.
Public Function Gotoh_Maximum(ByVal s1 As String, ByVal s2 As String) As Double
    Gotoh_Maximum = MinLen(s1, s2)
End Function

' -Length of the shorter string.
Public Function Gotoh_Minimum(ByVal s1 As String, ByVal s2 As String) As Double
    Gotoh_Minimum = -MinLen(s1, s2)
End Function

' (Distance - Minimum) / (Maximum - Minimum); 0 when either string is empty.
Public Function Gotoh_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                         Optional ByVal CaseSensitive As Boolean = True, _
                                         Optional ByVal Qval As Long = 1, _
                                         Optional ByVal GapOpen As Double = 1, _
                                         Optional ByVal GapExt As Double = 0.4) As Double
    Dim maximum As Double
    maximum = Gotoh_Maximum(s1, s2)
    If maximum = 0 Then Exit Function
    Gotoh_NormalizedDistance = (Gotoh_Distance(s1, s2, CaseSensitive, Qval, GapOpen, GapExt) + maximum) _
                               / (2 * maximum)
End Function

' (Similarity - Minimum) / (Maximum * 2); 1 when either string is empty.
Public Function Gotoh_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                           Optional ByVal CaseSensitive As Boolean = True, _
                                           Optional ByVal Qval As Long = 1, _
                                           Optional ByVal GapOpen As Double = 1, _
                                           Optional ByVal GapExt As Double = 0.4) As Double
    Dim maximum As Double
    maximum = Gotoh_Maximum(s1, s2)
    If maximum = 0 Then
        Gotoh_NormalizedSimilarity = 1
        Exit Function
    End If
    Gotoh_NormalizedSimilarity = (Gotoh_Similarity(s1, s2, CaseSensitive, Qval, GapOpen, GapExt) + maximum) _
                                 / (maximum * 2)
End Function

'===============================================================================
' Smith-Waterman (local alignment). As in textdistance, the score is the value
' in the table's last cell, not the best cell anywhere.
'
'   GapCost        Cost of each gap position (default 1).
'===============================================================================
Public Function SmithWaterman_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                         Optional ByVal CaseSensitive As Boolean = True, _
                                         Optional ByVal Qval As Long = 1, _
                                         Optional ByVal GapCost As Double = 1) As Double
    Dim t1() As String, t2() As String, n1 As Long, n2 As Long
    Dim d() As Double, i As Long, j As Long, best As Double, v As Double

    If Not CaseSensitive Then
        s1 = LCase$(s1)
        s2 = LCase$(s2)
    End If
    n1 = Tokenize(s1, Qval, t1)
    n2 = Tokenize(s2, Qval, t2)

    ' textdistance's quick answer on the token lists
    If SameTokens(t1, n1, t2, n2) Then
        SmithWaterman_Similarity = IIf(n1 < n2, n1, n2)
        Exit Function
    End If
    If n1 = 0 Or n2 = 0 Then Exit Function

    ReDim d(0 To n1, 0 To n2)
    For i = 1 To n1
        For j = 1 To n2
            best = d(i - 1, j - 1) + IIf(t1(i) = t2(j), 1, 0)
            v = d(i - 1, j) - GapCost
            If v > best Then best = v
            v = d(i, j - 1) - GapCost
            If v > best Then best = v
            If best < 0 Then best = 0
            d(i, j) = best
        Next j
    Next i
    SmithWaterman_Similarity = d(n1, n2)
End Function

' Length of the shorter string.
Public Function SmithWaterman_Maximum(ByVal s1 As String, ByVal s2 As String) As Double
    SmithWaterman_Maximum = MinLen(s1, s2)
End Function

' Maximum - Similarity.
Public Function SmithWaterman_Distance(ByVal s1 As String, ByVal s2 As String, _
                                       Optional ByVal CaseSensitive As Boolean = True, _
                                       Optional ByVal Qval As Long = 1, _
                                       Optional ByVal GapCost As Double = 1) As Double
    SmithWaterman_Distance = MinLen(s1, s2) - SmithWaterman_Similarity(s1, s2, CaseSensitive, Qval, GapCost)
End Function

' Distance / Maximum; 0 when either string is empty.
Public Function SmithWaterman_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                                 Optional ByVal CaseSensitive As Boolean = True, _
                                                 Optional ByVal Qval As Long = 1, _
                                                 Optional ByVal GapCost As Double = 1) As Double
    If MinLen(s1, s2) = 0 Then Exit Function
    SmithWaterman_NormalizedDistance = SmithWaterman_Distance(s1, s2, CaseSensitive, Qval, GapCost) / MinLen(s1, s2)
End Function

Public Function SmithWaterman_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                                   Optional ByVal CaseSensitive As Boolean = True, _
                                                   Optional ByVal Qval As Long = 1, _
                                                   Optional ByVal GapCost As Double = 1) As Double
    SmithWaterman_NormalizedSimilarity = 1 - SmithWaterman_NormalizedDistance(s1, s2, CaseSensitive, Qval, GapCost)
End Function

'===============================================================================
' Helpers
'===============================================================================

Private Function MinLen(ByVal s1 As String, ByVal s2 As String) As Long
    MinLen = IIf(Len(s1) < Len(s2), Len(s1), Len(s2))
End Function

Private Function MaxLen(ByVal s1 As String, ByVal s2 As String) As Long
    MaxLen = IIf(Len(s1) > Len(s2), Len(s1), Len(s2))
End Function

Private Function Larger(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then Larger = a Else Larger = b
End Function

Private Function SameTokens(t1() As String, ByVal n1 As Long, t2() As String, ByVal n2 As Long) As Boolean
    Dim i As Long
    If n1 <> n2 Then Exit Function
    For i = 1 To n1
        If t1(i) <> t2(i) Then Exit Function
    Next i
    SameTokens = True
End Function

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
    Select Case AscW(ch) And &HFFFF&
        Case 9 To 13, 28 To 32, &H85, &HA0, &H1680, &H2000 To &H200A, &H2028, &H2029, &H202F, &H205F, &H3000
            IsSpace = True
    End Select
End Function
