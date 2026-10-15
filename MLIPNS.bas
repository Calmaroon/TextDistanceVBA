Attribute VB_Name = "modMLIPNS"
'===============================================================================
' MLIPNS similarity - port of textdistance.MLIPNS (Python textdistance 4.6)
'
' Modified Language-Independent Product Name Search: a yes/no match based on
' the Hamming distance. Returns 1 if the strings match, else 0:
'   MLIPNS_Similarity("Tomato", "Tamato") = 1
'   MLIPNS_Similarity("Tomato", "Potato") = 1   (2 mismatches are forgiven)
'   MLIPNS_Similarity("Tomato", "Banana") = 0
'
' Starting from the Hamming distance, up to MaxMismatches mismatches are
' forgiven one at a time; the strings match as soon as
'   1 - (length - mismatches) / length <= Threshold
' where length is the longer length (both reduced by each forgiven mismatch).
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
'   Threshold      Default 0.25.
'   MaxMismatches  Default 2.
'===============================================================================
Option Explicit
Option Compare Binary

Public Function MLIPNS_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                  Optional ByVal CaseSensitive As Boolean = True, _
                                  Optional ByVal Qval As Long = 1, _
                                  Optional ByVal Threshold As Double = 0.25, _
                                  Optional ByVal MaxMismatches As Long = 2) As Long
    Dim t1() As String, t2() As String, n1 As Long, n2 As Long
    Dim ham As Long, maxLen As Long, mismatches As Long, i As Long

    If Not CaseSensitive Then
        s1 = LCase$(s1)
        s2 = LCase$(s2)
    End If
    n1 = Tokenize(s1, Qval, t1)
    n2 = Tokenize(s2, Qval, t2)

    ' textdistance's quick answer on the token lists
    If SameTokens(t1, n1, t2, n2) Then
        MLIPNS_Similarity = 1
        Exit Function
    End If
    If n1 = 0 Or n2 = 0 Then Exit Function

    ' Hamming distance between the token lists (extra tokens count as mismatches)
    For i = 1 To IIf(n1 < n2, n1, n2)
        If t1(i) <> t2(i) Then ham = ham + 1
    Next i
    ham = ham + Abs(n1 - n2)
    maxLen = IIf(n1 > n2, n1, n2)

    Do While mismatches <= MaxMismatches
        If maxLen = 0 Then
            MLIPNS_Similarity = 1
            Exit Function
        End If
        If 1 - (maxLen - ham) / maxLen <= Threshold Then
            MLIPNS_Similarity = 1
            Exit Function
        End If
        mismatches = mismatches + 1
        ham = ham - 1
        maxLen = maxLen - 1
    Loop

    If maxLen = 0 Then MLIPNS_Similarity = 1
End Function

Public Function MLIPNS_Distance(ByVal s1 As String, ByVal s2 As String, _
                                Optional ByVal CaseSensitive As Boolean = True, _
                                Optional ByVal Qval As Long = 1, _
                                Optional ByVal Threshold As Double = 0.25, _
                                Optional ByVal MaxMismatches As Long = 2) As Long
    MLIPNS_Distance = 1 - MLIPNS_Similarity(s1, s2, CaseSensitive, Qval, Threshold, MaxMismatches)
End Function

Public Function MLIPNS_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                          Optional ByVal CaseSensitive As Boolean = True, _
                                          Optional ByVal Qval As Long = 1, _
                                          Optional ByVal Threshold As Double = 0.25, _
                                          Optional ByVal MaxMismatches As Long = 2) As Double
    MLIPNS_NormalizedDistance = MLIPNS_Distance(s1, s2, CaseSensitive, Qval, Threshold, MaxMismatches)
End Function

Public Function MLIPNS_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                            Optional ByVal CaseSensitive As Boolean = True, _
                                            Optional ByVal Qval As Long = 1, _
                                            Optional ByVal Threshold As Double = 0.25, _
                                            Optional ByVal MaxMismatches As Long = 2) As Double
    MLIPNS_NormalizedSimilarity = 1 - MLIPNS_NormalizedDistance(s1, s2, CaseSensitive, Qval, Threshold, MaxMismatches)
End Function

' Always 1.
Public Function MLIPNS_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    MLIPNS_Maximum = 1
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
