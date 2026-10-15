Attribute VB_Name = "modPhonetic"
'===============================================================================
' Phonetic measures - port of textdistance.MRA and textdistance.Editex
' (Python textdistance 4.6)
'
'   MRA     Match Rating Approach (Western Airlines surname matching). Each name
'           is reduced to a short codex, e.g. MRA_Codex("Catherine") = "CTHRN",
'           and the codexes are compared. Similarity runs 0 .. Maximum.
'   Editex  Edit distance where substituting letters that sound alike (e.g.
'           D/T, M/N, vowels) costs less than unrelated letters.
'
' Both compare upper-cased text, so they are always case-insensitive (there is
' no CaseSensitive argument). Note: VBA's UCase$ leaves a few characters alone
' that Python changes (e.g. German sharp s), so results can differ for those.
'
' Pure VBA with no host references, so it works in Excel, Access, Word, etc.
'===============================================================================
Option Explicit
Option Compare Binary

' Editex letter groups: letters in the same group sound alike
Private Const EDITEX_GROUPED As String = "AEIOUYBPCKQDTLRMNGJFVSXZ"
Private Const EDITEX_UNGROUPED As String = "HW"

'-------------------------------------------------------------------------------
' MRA
'-------------------------------------------------------------------------------

' The MRA codex of a name: upper case, vowels removed after the first letter,
' repeated letters collapsed, and at most 6 characters (first 3 + last 3).
Public Function MRA_Codex(ByVal s As String) As String
    Dim word As String, stripped As String, result As String, ch As String
    Dim i As Long

    If Len(s) = 0 Then Exit Function
    word = UCase$(s)

    stripped = Left$(word, 1)
    For i = 2 To Len(word)
        ch = Mid$(word, i, 1)
        If InStr(1, "AEIOU", ch, vbBinaryCompare) = 0 Then stripped = stripped & ch
    Next i

    result = Left$(stripped, 1)
    For i = 2 To Len(stripped)
        ch = Mid$(stripped, i, 1)
        If ch <> Right$(result, 1) Then result = result & ch
    Next i

    If Len(result) > 6 Then result = Left$(result, 3) & Right$(result, 3)
    MRA_Codex = result
End Function

' Similarity rating: matching codex characters (0 .. Maximum). Returns 0 if
' either string is empty or the codex lengths differ by more than 2.
Public Function MRA_Similarity(ByVal s1 As String, ByVal s2 As String) As Long
    Dim c1 As String, c2 As String, n1 As String, n2 As String
    Dim maxLen As Long, minLen As Long, pass As Long, i As Long

    If Len(s1) = 0 Or Len(s2) = 0 Then Exit Function
    c1 = MRA_Codex(s1)
    c2 = MRA_Codex(s2)
    maxLen = IIf(Len(c1) > Len(c2), Len(c1), Len(c2))
    If Abs(Len(c1) - Len(c2)) > 2 Then Exit Function

    ' Drop matching characters position by position (textdistance does this
    ' left to right, twice)
    For pass = 1 To 2
        minLen = IIf(Len(c1) < Len(c2), Len(c1), Len(c2))
        n1 = vbNullString
        n2 = vbNullString
        For i = 1 To minLen
            If Mid$(c1, i, 1) <> Mid$(c2, i, 1) Then
                n1 = n1 & Mid$(c1, i, 1)
                n2 = n2 & Mid$(c2, i, 1)
            End If
        Next i
        c1 = n1 & Mid$(c1, minLen + 1)
        c2 = n2 & Mid$(c2, minLen + 1)
    Next pass

    MRA_Similarity = maxLen - IIf(Len(c1) > Len(c2), Len(c1), Len(c2))
End Function

' Length of the longer codex.
Public Function MRA_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    Dim len1 As Long, len2 As Long
    len1 = Len(MRA_Codex(s1))
    len2 = Len(MRA_Codex(s2))
    MRA_Maximum = IIf(len1 > len2, len1, len2)
End Function

' Maximum - Similarity.
Public Function MRA_Distance(ByVal s1 As String, ByVal s2 As String) As Long
    MRA_Distance = MRA_Maximum(s1, s2) - MRA_Similarity(s1, s2)
End Function

' Distance / Maximum; 0 when both strings are empty.
Public Function MRA_NormalizedDistance(ByVal s1 As String, ByVal s2 As String) As Double
    Dim maximum As Long
    maximum = MRA_Maximum(s1, s2)
    If maximum = 0 Then Exit Function
    MRA_NormalizedDistance = MRA_Distance(s1, s2) / maximum
End Function

Public Function MRA_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String) As Double
    MRA_NormalizedSimilarity = 1 - MRA_NormalizedDistance(s1, s2)
End Function

'-------------------------------------------------------------------------------
' Editex
'
'   LocalMode      False (default). True makes deleting characters at the start
'                  of s1 free (textdistance's local=True).
'   MatchCost      Cost of identical letters (0).
'   GroupCost      Cost of letters in the same sound group (1); raised to
'                  MatchCost if lower.
'   MismatchCost   Cost of other letters (2); raised to GroupCost if lower.
'-------------------------------------------------------------------------------

Public Function Editex_Distance(ByVal s1 As String, ByVal s2 As String, _
                                Optional ByVal LocalMode As Boolean = False, _
                                Optional ByVal MatchCost As Long = 0, _
                                Optional ByVal GroupCost As Long = 1, _
                                Optional ByVal MismatchCost As Long = 2) As Long
    Dim len1 As Long, len2 As Long, i As Long, j As Long
    Dim maximum As Long, best As Long, v As Long
    Dim d() As Long

    If GroupCost < MatchCost Then GroupCost = MatchCost
    If MismatchCost < GroupCost Then MismatchCost = GroupCost

    If s1 = s2 Then Exit Function
    maximum = Editex_Maximum(s1, s2, LocalMode, MatchCost, GroupCost, MismatchCost)
    If Len(s1) = 0 Or Len(s2) = 0 Then
        Editex_Distance = maximum
        Exit Function
    End If

    ' Position 1 holds a leading space, as in textdistance
    s1 = " " & UCase$(s1)
    s2 = " " & UCase$(s2)
    len1 = Len(s1) - 1
    len2 = Len(s2) - 1

    ReDim d(0 To len1, 0 To len2)
    If Not LocalMode Then
        For i = 1 To len1
            d(i, 0) = d(i - 1, 0) + DeleteCost(Mid$(s1, i, 1), Mid$(s1, i + 1, 1), MatchCost, GroupCost, MismatchCost)
        Next i
    End If
    For j = 1 To len2
        d(0, j) = d(0, j - 1) + DeleteCost(Mid$(s2, j, 1), Mid$(s2, j + 1, 1), MatchCost, GroupCost, MismatchCost)
    Next j

    For i = 1 To len1
        For j = 1 To len2
            best = d(i - 1, j) + DeleteCost(Mid$(s1, i, 1), Mid$(s1, i + 1, 1), MatchCost, GroupCost, MismatchCost)
            v = d(i, j - 1) + DeleteCost(Mid$(s2, j, 1), Mid$(s2, j + 1, 1), MatchCost, GroupCost, MismatchCost)
            If v < best Then best = v
            v = d(i - 1, j - 1) + ReplaceCost(Mid$(s1, i + 1, 1), Mid$(s2, j + 1, 1), MatchCost, GroupCost, MismatchCost)
            If v < best Then best = v
            d(i, j) = best
        Next j
    Next i

    Editex_Distance = IIf(d(len1, len2) < maximum, d(len1, len2), maximum)
End Function

' Length of the longer string * MismatchCost.
Public Function Editex_Maximum(ByVal s1 As String, ByVal s2 As String, _
                               Optional ByVal LocalMode As Boolean = False, _
                               Optional ByVal MatchCost As Long = 0, _
                               Optional ByVal GroupCost As Long = 1, _
                               Optional ByVal MismatchCost As Long = 2) As Long
    If GroupCost < MatchCost Then GroupCost = MatchCost
    If MismatchCost < GroupCost Then MismatchCost = GroupCost
    Editex_Maximum = IIf(Len(s1) > Len(s2), Len(s1), Len(s2)) * MismatchCost
End Function

' Maximum - Distance.
Public Function Editex_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                  Optional ByVal LocalMode As Boolean = False, _
                                  Optional ByVal MatchCost As Long = 0, _
                                  Optional ByVal GroupCost As Long = 1, _
                                  Optional ByVal MismatchCost As Long = 2) As Long
    Editex_Similarity = Editex_Maximum(s1, s2, LocalMode, MatchCost, GroupCost, MismatchCost) _
                      - Editex_Distance(s1, s2, LocalMode, MatchCost, GroupCost, MismatchCost)
End Function

' Distance / Maximum; 0 when both strings are empty.
Public Function Editex_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                          Optional ByVal LocalMode As Boolean = False, _
                                          Optional ByVal MatchCost As Long = 0, _
                                          Optional ByVal GroupCost As Long = 1, _
                                          Optional ByVal MismatchCost As Long = 2) As Double
    Dim maximum As Long
    maximum = Editex_Maximum(s1, s2, LocalMode, MatchCost, GroupCost, MismatchCost)
    If maximum = 0 Then Exit Function
    Editex_NormalizedDistance = Editex_Distance(s1, s2, LocalMode, MatchCost, GroupCost, MismatchCost) / maximum
End Function

Public Function Editex_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                            Optional ByVal LocalMode As Boolean = False, _
                                            Optional ByVal MatchCost As Long = 0, _
                                            Optional ByVal GroupCost As Long = 1, _
                                            Optional ByVal MismatchCost As Long = 2) As Double
    Editex_NormalizedSimilarity = 1 - Editex_NormalizedDistance(s1, s2, LocalMode, MatchCost, GroupCost, MismatchCost)
End Function

' Cost of replacing a with b (textdistance's r_cost).
Private Function ReplaceCost(ByVal a As String, ByVal b As String, ByVal MatchCost As Long, _
                             ByVal GroupCost As Long, ByVal MismatchCost As Long) As Long
    If a = b Then
        ReplaceCost = MatchCost
    ElseIf InStr(1, EDITEX_GROUPED, a, vbBinaryCompare) = 0 Or InStr(1, EDITEX_GROUPED, b, vbBinaryCompare) = 0 Then
        ReplaceCost = MismatchCost
    ElseIf SameGroup(a, b) Then
        ReplaceCost = GroupCost
    Else
        ReplaceCost = MismatchCost
    End If
End Function

' Cost of deleting cur after prev (textdistance's d_cost): cheaper after H or W.
Private Function DeleteCost(ByVal prev As String, ByVal cur As String, ByVal MatchCost As Long, _
                            ByVal GroupCost As Long, ByVal MismatchCost As Long) As Long
    If prev <> cur And InStr(1, EDITEX_UNGROUPED, prev, vbBinaryCompare) > 0 Then
        DeleteCost = GroupCost
    Else
        DeleteCost = ReplaceCost(prev, cur, MatchCost, GroupCost, MismatchCost)
    End If
End Function

' True if a and b are in one of the Editex sound groups.
Private Function SameGroup(ByVal a As String, ByVal b As String) As Boolean
    Static groups As Variant
    Dim g As Variant
    If IsEmpty(groups) Then groups = Array("AEIOUY", "BP", "CKQ", "DT", "LR", "MN", "GJ", "FPV", "SXZ", "CSZ")
    For Each g In groups
        If InStr(1, g, a, vbBinaryCompare) > 0 And InStr(1, g, b, vbBinaryCompare) > 0 Then
            SameGroup = True
            Exit Function
        End If
    Next g
End Function
