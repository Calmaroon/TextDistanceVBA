Attribute VB_Name = "modLCSSeq"
'===============================================================================
' Longest common subsequence - port of textdistance.LCSSeq
' (Python textdistance 4.6)
'
' The longest sequence of characters appearing in both strings in the same
' order, not necessarily next to each other:
'   LCSSeq_Sequence("test", "text")   = "tet"
'   LCSSeq_Similarity("test", "text") = 3   (its length)
'
' Pure VBA with no host references, so it works in Excel, Access, Word, etc.
'
' Common arguments
'   s1, s2         Strings to compare.
'   CaseSensitive  True (default) compares exactly; False ignores case (the
'                  returned sequence is then lower case).
'                  (VBA addition - Python's textdistance has no such flag.)
'===============================================================================
Option Explicit
Option Compare Binary

' The longest common subsequence itself (Python: textdistance.lcsseq(s1, s2)).
' When several exist, returns the same one as Python.
Public Function LCSSeq_Sequence(ByVal s1 As String, ByVal s2 As String, _
                                Optional ByVal CaseSensitive As Boolean = True) As String
    Dim len1 As Long, len2 As Long, i As Long, j As Long, k As Long
    Dim c1() As Long, c2() As Long, lengths() As Long
    Dim result As String

    If Not CaseSensitive Then
        s1 = LCase$(s1)
        s2 = LCase$(s2)
    End If

    len1 = Len(s1)
    len2 = Len(s2)
    If len1 = 0 Or len2 = 0 Then Exit Function

    CharCodes s1, c1
    CharCodes s2, c2

    ' lengths(i, j) = LCS length of the first i chars of s1 and first j of s2
    ReDim lengths(0 To len1, 0 To len2)
    For i = 1 To len1
        For j = 1 To len2
            If c1(i) = c2(j) Then
                lengths(i, j) = lengths(i - 1, j - 1) + 1
            ElseIf lengths(i, j - 1) > lengths(i - 1, j) Then
                lengths(i, j) = lengths(i, j - 1)
            Else
                lengths(i, j) = lengths(i - 1, j)
            End If
        Next j
    Next i

    ' Walk back through the table, preferring the same moves as Python
    k = lengths(len1, len2)
    If k = 0 Then Exit Function
    result = String$(k, " ")
    i = len1
    j = len2
    Do While i <> 0 And j <> 0
        If lengths(i, j) = lengths(i - 1, j) Then
            i = i - 1
        ElseIf lengths(i, j) = lengths(i, j - 1) Then
            j = j - 1
        Else
            Mid$(result, k, 1) = Mid$(s1, i, 1)
            k = k - 1
            i = i - 1
            j = j - 1
        End If
    Loop

    LCSSeq_Sequence = result
End Function

' Length of the longest common subsequence.
Public Function LCSSeq_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                  Optional ByVal CaseSensitive As Boolean = True) As Long
    LCSSeq_Similarity = Len(LCSSeq_Sequence(s1, s2, CaseSensitive))
End Function

' Largest possible similarity: the length of the longer string.
Public Function LCSSeq_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    LCSSeq_Maximum = IIf(Len(s1) > Len(s2), Len(s1), Len(s2))
End Function

' Maximum - Similarity.
Public Function LCSSeq_Distance(ByVal s1 As String, ByVal s2 As String, _
                                Optional ByVal CaseSensitive As Boolean = True) As Long
    LCSSeq_Distance = LCSSeq_Maximum(s1, s2) - LCSSeq_Similarity(s1, s2, CaseSensitive)
End Function

' Distance scaled to 0 (identical) .. 1 (nothing in common).
' Two empty strings give 0.
Public Function LCSSeq_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                          Optional ByVal CaseSensitive As Boolean = True) As Double
    Dim maximum As Long
    maximum = LCSSeq_Maximum(s1, s2)
    If maximum = 0 Then Exit Function
    LCSSeq_NormalizedDistance = LCSSeq_Distance(s1, s2, CaseSensitive) / maximum
End Function

' Similarity scaled to 0 (nothing in common) .. 1 (identical).
' Two empty strings give 1.
Public Function LCSSeq_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                            Optional ByVal CaseSensitive As Boolean = True) As Double
    LCSSeq_NormalizedSimilarity = 1 - LCSSeq_NormalizedDistance(s1, s2, CaseSensitive)
End Function

' Character codes of s, 1-based, so comparisons avoid repeated Mid$ calls.
Private Sub CharCodes(ByVal s As String, codes() As Long)
    Dim i As Long
    ReDim codes(1 To Len(s))
    For i = 1 To Len(s)
        codes(i) = AscW(Mid$(s, i, 1))
    Next i
End Sub
