Attribute VB_Name = "modCosine"
'===============================================================================
' Cosine similarity (Ochiai coefficient) - port of textdistance.Cosine
' (Python textdistance 4.6)
'
' Treats each string as a bag of tokens (letters, letter n-grams or words) and
' returns  shared tokens / Sqr(tokens in s1 * tokens in s2),  from 0 to 1:
'   Cosine_Similarity("nelson", "neilsen")                = 0.772
'   Cosine_Similarity("Smith John", "John Smith", , 0)    = 1   (by words)
'
' Similarity is the primary value; the maximum is always 1, so
' Distance = 1 - Similarity.
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
'   AsSet          False (default) counts repeated tokens; True counts each
'                  distinct token once.
'
' When a non-empty string yields no tokens (e.g. "a" with Qval 2) Python
' raises ZeroDivisionError; this port returns 0 instead.
'===============================================================================
Option Explicit
Option Compare Binary

Public Function Cosine_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                  Optional ByVal CaseSensitive As Boolean = True, _
                                  Optional ByVal Qval As Long = 1, _
                                  Optional ByVal AsSet As Boolean = False) As Double
    Dim t1() As String, t2() As String
    Dim n1 As Long, n2 As Long, i As Long, j As Long
    Dim count1 As Long, count2 As Long, common As Long, c As Long

    If Not CaseSensitive Then
        s1 = LCase$(s1)
        s2 = LCase$(s2)
    End If

    If s1 = s2 Then
        Cosine_Similarity = 1
        Exit Function
    End If
    If Len(s1) = 0 Or Len(s2) = 0 Then Exit Function

    n1 = Tokenize(s1, Qval, t1)
    n2 = Tokenize(s2, Qval, t2)
    If n1 = 0 Or n2 = 0 Then Exit Function

    SortStrings t1, 1, n1
    SortStrings t2, 1, n2
    count1 = CountTokens(t1, n1, AsSet)
    count2 = CountTokens(t2, n2, AsSet)

    ' Walk both sorted lists: each shared token counts min(count in s1, count in s2)
    ' times, or once with AsSet
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

    Cosine_Similarity = common / Sqr(CDbl(count1) * count2)
End Function

Public Function Cosine_Distance(ByVal s1 As String, ByVal s2 As String, _
                                Optional ByVal CaseSensitive As Boolean = True, _
                                Optional ByVal Qval As Long = 1, _
                                Optional ByVal AsSet As Boolean = False) As Double
    Cosine_Distance = 1 - Cosine_Similarity(s1, s2, CaseSensitive, Qval, AsSet)
End Function

Public Function Cosine_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                            Optional ByVal CaseSensitive As Boolean = True, _
                                            Optional ByVal Qval As Long = 1, _
                                            Optional ByVal AsSet As Boolean = False) As Double
    Cosine_NormalizedSimilarity = 1 - Cosine_NormalizedDistance(s1, s2, CaseSensitive, Qval, AsSet)
End Function

Public Function Cosine_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                          Optional ByVal CaseSensitive As Boolean = True, _
                                          Optional ByVal Qval As Long = 1, _
                                          Optional ByVal AsSet As Boolean = False) As Double
    Cosine_NormalizedDistance = Cosine_Distance(s1, s2, CaseSensitive, Qval, AsSet)
End Function

' Always 1.
Public Function Cosine_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    Cosine_Maximum = 1
End Function

'-------------------------------------------------------------------------------
' Token helpers (textdistance's _get_sequences / Counter handling)
'-------------------------------------------------------------------------------

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
Private Function CountTokens(tokens() As String, ByVal n As Long, ByVal AsSet As Boolean) As Long
    Dim i As Long
    If Not AsSet Then
        CountTokens = n
        Exit Function
    End If
    For i = 1 To n
        If i = 1 Then
            CountTokens = 1
        ElseIf tokens(i) <> tokens(i - 1) Then
            CountTokens = CountTokens + 1
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
