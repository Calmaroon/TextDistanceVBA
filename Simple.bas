Attribute VB_Name = "modSimple"
'===============================================================================
' Simple measures - port of textdistance.Prefix, Postfix, Length, Identity and
' Matrix (Python textdistance 4.6)
'
'   Prefix    common start of both strings    Prefix_CommonPrefix("test", "text") = "te"
'   Postfix   common end of both strings      Postfix_CommonPostfix("ring", "sing") = "ing"
'   Length    difference in length            Length_Distance("test", "testing") = 3
'   Identity  1 if identical, else 0          Identity_Similarity("a", "a") = 1
'   Matrix    score looked up in a table      see Matrix_Similarity
'
' Pure VBA with no host references, so it works in Excel, Access, Word, etc.
'
' Common arguments
'   s1, s2         Strings to compare.
'   CaseSensitive  True (default) compares exactly; False ignores case (returned
'                  text is then lower case).
'                  (VBA addition - Python's textdistance has no such flag.)
'===============================================================================
Option Explicit
Option Compare Binary

'-------------------------------------------------------------------------------
' Prefix: similarity = length of the common prefix
'-------------------------------------------------------------------------------

' The common prefix itself (Python: textdistance.prefix(s1, s2)).
Public Function Prefix_CommonPrefix(ByVal s1 As String, ByVal s2 As String, _
                                    Optional ByVal CaseSensitive As Boolean = True) As String
    Dim n As Long, limit As Long
    If Not CaseSensitive Then
        s1 = LCase$(s1)
        s2 = LCase$(s2)
    End If
    limit = MinLen(s1, s2)
    Do While n < limit
        If Mid$(s1, n + 1, 1) <> Mid$(s2, n + 1, 1) Then Exit Do
        n = n + 1
    Loop
    Prefix_CommonPrefix = Left$(s1, n)
End Function

Public Function Prefix_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                  Optional ByVal CaseSensitive As Boolean = True) As Long
    Prefix_Similarity = Len(Prefix_CommonPrefix(s1, s2, CaseSensitive))
End Function

' Length of the longer string.
Public Function Prefix_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    Prefix_Maximum = MaxLen(s1, s2)
End Function

' Maximum - Similarity.
Public Function Prefix_Distance(ByVal s1 As String, ByVal s2 As String, _
                                Optional ByVal CaseSensitive As Boolean = True) As Long
    Prefix_Distance = MaxLen(s1, s2) - Prefix_Similarity(s1, s2, CaseSensitive)
End Function

' Distance / Maximum; two empty strings give 0.
Public Function Prefix_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                          Optional ByVal CaseSensitive As Boolean = True) As Double
    If MaxLen(s1, s2) = 0 Then Exit Function
    Prefix_NormalizedDistance = Prefix_Distance(s1, s2, CaseSensitive) / MaxLen(s1, s2)
End Function

Public Function Prefix_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                            Optional ByVal CaseSensitive As Boolean = True) As Double
    Prefix_NormalizedSimilarity = 1 - Prefix_NormalizedDistance(s1, s2, CaseSensitive)
End Function

'-------------------------------------------------------------------------------
' Postfix: similarity = length of the common suffix
'-------------------------------------------------------------------------------

' The common suffix itself (Python: textdistance.postfix(s1, s2)).
Public Function Postfix_CommonPostfix(ByVal s1 As String, ByVal s2 As String, _
                                      Optional ByVal CaseSensitive As Boolean = True) As String
    Dim n As Long, limit As Long, len1 As Long, len2 As Long
    If Not CaseSensitive Then
        s1 = LCase$(s1)
        s2 = LCase$(s2)
    End If
    len1 = Len(s1)
    len2 = Len(s2)
    limit = MinLen(s1, s2)
    Do While n < limit
        If Mid$(s1, len1 - n, 1) <> Mid$(s2, len2 - n, 1) Then Exit Do
        n = n + 1
    Loop
    Postfix_CommonPostfix = Right$(s1, n)
End Function

Public Function Postfix_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                   Optional ByVal CaseSensitive As Boolean = True) As Long
    Postfix_Similarity = Len(Postfix_CommonPostfix(s1, s2, CaseSensitive))
End Function

' Length of the longer string.
Public Function Postfix_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    Postfix_Maximum = MaxLen(s1, s2)
End Function

' Maximum - Similarity.
Public Function Postfix_Distance(ByVal s1 As String, ByVal s2 As String, _
                                 Optional ByVal CaseSensitive As Boolean = True) As Long
    Postfix_Distance = MaxLen(s1, s2) - Postfix_Similarity(s1, s2, CaseSensitive)
End Function

' Distance / Maximum; two empty strings give 0.
Public Function Postfix_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                           Optional ByVal CaseSensitive As Boolean = True) As Double
    If MaxLen(s1, s2) = 0 Then Exit Function
    Postfix_NormalizedDistance = Postfix_Distance(s1, s2, CaseSensitive) / MaxLen(s1, s2)
End Function

Public Function Postfix_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                             Optional ByVal CaseSensitive As Boolean = True) As Double
    Postfix_NormalizedSimilarity = 1 - Postfix_NormalizedDistance(s1, s2, CaseSensitive)
End Function

'-------------------------------------------------------------------------------
' Length: distance = difference in length (CaseSensitive has no effect, but is
' accepted for consistency)
'-------------------------------------------------------------------------------

Public Function Length_Distance(ByVal s1 As String, ByVal s2 As String, _
                                Optional ByVal CaseSensitive As Boolean = True) As Long
    Length_Distance = Abs(Len(s1) - Len(s2))
End Function

' Length of the longer string.
Public Function Length_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    Length_Maximum = MaxLen(s1, s2)
End Function

' Maximum - Distance, i.e. the length of the shorter string.
Public Function Length_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                  Optional ByVal CaseSensitive As Boolean = True) As Long
    Length_Similarity = MaxLen(s1, s2) - Length_Distance(s1, s2)
End Function

' Distance / Maximum; two empty strings give 0.
Public Function Length_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                          Optional ByVal CaseSensitive As Boolean = True) As Double
    If MaxLen(s1, s2) = 0 Then Exit Function
    Length_NormalizedDistance = Length_Distance(s1, s2) / MaxLen(s1, s2)
End Function

Public Function Length_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                            Optional ByVal CaseSensitive As Boolean = True) As Double
    Length_NormalizedSimilarity = 1 - Length_NormalizedDistance(s1, s2)
End Function

'-------------------------------------------------------------------------------
' Identity: similarity = 1 if the strings are identical, else 0
'-------------------------------------------------------------------------------

Public Function Identity_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                    Optional ByVal CaseSensitive As Boolean = True) As Long
    If Not CaseSensitive Then
        s1 = LCase$(s1)
        s2 = LCase$(s2)
    End If
    If s1 = s2 Then Identity_Similarity = 1
End Function

' Always 1.
Public Function Identity_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    Identity_Maximum = 1
End Function

Public Function Identity_Distance(ByVal s1 As String, ByVal s2 As String, _
                                  Optional ByVal CaseSensitive As Boolean = True) As Long
    Identity_Distance = 1 - Identity_Similarity(s1, s2, CaseSensitive)
End Function

Public Function Identity_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                            Optional ByVal CaseSensitive As Boolean = True) As Double
    Identity_NormalizedDistance = Identity_Distance(s1, s2, CaseSensitive)
End Function

Public Function Identity_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                              Optional ByVal CaseSensitive As Boolean = True) As Double
    Identity_NormalizedSimilarity = Identity_Similarity(s1, s2, CaseSensitive)
End Function

'-------------------------------------------------------------------------------
' Matrix: similarity looked up in a table of scores
'
'   Table          Optional 2-D array (or Excel range) with 3 columns:
'                  string 1, string 2, score. The pair (s1, s2) is looked up;
'                  if a pair appears more than once the last row wins.
'   MismatchCost   Score when the pair is not in the table and s1 <> s2 (0).
'   MatchCost      Score when the pair is not in the table and s1 = s2 (1).
'                  Also the maximum.
'   Symmetric      True (default) also looks up (s2, s1).
'
' Without a table this is MatchCost if s1 = s2, else MismatchCost.
' Example table (DNA scoring):  A A 10 / G G 7 / A G -1 ...
'   Matrix_Similarity("A", "G", Table) = -1
'-------------------------------------------------------------------------------

Public Function Matrix_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                  Optional ByVal Table As Variant, _
                                  Optional ByVal MismatchCost As Double = 0, _
                                  Optional ByVal MatchCost As Double = 1, _
                                  Optional ByVal Symmetric As Boolean = True, _
                                  Optional ByVal CaseSensitive As Boolean = True) As Double
    Dim values As Variant, row As Long, found As Boolean

    If Not CaseSensitive Then
        s1 = LCase$(s1)
        s2 = LCase$(s2)
    End If

    If Not IsMissing(Table) Then
        values = Table          ' a Range gives its .Value array here
        If Not IsEmpty(values) Then
            If Not IsTable(values) Then
                Err.Raise 5, "Matrix_Similarity", "Table must be a 2-D array or range with 3 columns"
            End If
            found = FindPair(values, s1, s2, CaseSensitive, row)
            If Not found And Symmetric Then found = FindPair(values, s2, s1, CaseSensitive, row)
            If found Then
                Matrix_Similarity = CDbl(values(row, LBound(values, 2) + 2))
                Exit Function
            End If
        End If
    End If

    If s1 = s2 Then
        Matrix_Similarity = MatchCost
    Else
        Matrix_Similarity = MismatchCost
    End If
End Function

' MatchCost.
Public Function Matrix_Maximum(ByVal s1 As String, ByVal s2 As String, _
                               Optional ByVal Table As Variant, _
                               Optional ByVal MismatchCost As Double = 0, _
                               Optional ByVal MatchCost As Double = 1, _
                               Optional ByVal Symmetric As Boolean = True, _
                               Optional ByVal CaseSensitive As Boolean = True) As Double
    Matrix_Maximum = MatchCost
End Function

' MatchCost - Similarity.
Public Function Matrix_Distance(ByVal s1 As String, ByVal s2 As String, _
                                Optional ByVal Table As Variant, _
                                Optional ByVal MismatchCost As Double = 0, _
                                Optional ByVal MatchCost As Double = 1, _
                                Optional ByVal Symmetric As Boolean = True, _
                                Optional ByVal CaseSensitive As Boolean = True) As Double
    Matrix_Distance = MatchCost - Matrix_Similarity(s1, s2, Table, MismatchCost, MatchCost, Symmetric, CaseSensitive)
End Function

' Distance / MatchCost; 0 when MatchCost is 0.
Public Function Matrix_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                          Optional ByVal Table As Variant, _
                                          Optional ByVal MismatchCost As Double = 0, _
                                          Optional ByVal MatchCost As Double = 1, _
                                          Optional ByVal Symmetric As Boolean = True, _
                                          Optional ByVal CaseSensitive As Boolean = True) As Double
    If MatchCost = 0 Then Exit Function
    Matrix_NormalizedDistance = Matrix_Distance(s1, s2, Table, MismatchCost, MatchCost, Symmetric, CaseSensitive) / MatchCost
End Function

Public Function Matrix_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                            Optional ByVal Table As Variant, _
                                            Optional ByVal MismatchCost As Double = 0, _
                                            Optional ByVal MatchCost As Double = 1, _
                                            Optional ByVal Symmetric As Boolean = True, _
                                            Optional ByVal CaseSensitive As Boolean = True) As Double
    Matrix_NormalizedSimilarity = 1 - Matrix_NormalizedDistance(s1, s2, Table, MismatchCost, MatchCost, Symmetric, CaseSensitive)
End Function

'-------------------------------------------------------------------------------
' Helpers
'-------------------------------------------------------------------------------

Private Function MinLen(ByVal s1 As String, ByVal s2 As String) As Long
    MinLen = IIf(Len(s1) < Len(s2), Len(s1), Len(s2))
End Function

Private Function MaxLen(ByVal s1 As String, ByVal s2 As String) As Long
    MaxLen = IIf(Len(s1) > Len(s2), Len(s1), Len(s2))
End Function

' True if values is a 2-D array with at least 3 columns.
Private Function IsTable(values As Variant) As Boolean
    Dim columns As Long
    If Not IsArray(values) Then Exit Function
    On Error GoTo NotTable
    columns = UBound(values, 2) - LBound(values, 2) + 1
    IsTable = columns >= 3
NotTable:
End Function

' Finds the last row of the table whose first two columns are (k1, k2).
Private Function FindPair(values As Variant, ByVal k1 As String, ByVal k2 As String, _
                          ByVal CaseSensitive As Boolean, ByRef FoundRow As Long) As Boolean
    Dim row As Long, c As Long, a As String, b As String
    c = LBound(values, 2)
    For row = UBound(values, 1) To LBound(values, 1) Step -1
        a = CStr(values(row, c))
        b = CStr(values(row, c + 1))
        If Not CaseSensitive Then
            a = LCase$(a)
            b = LCase$(b)
        End If
        If a = k1 And b = k2 Then
            FoundRow = row
            FindPair = True
            Exit Function
        End If
    Next row
End Function
