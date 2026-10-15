Attribute VB_Name = "modExcelUDF"
'===============================================================================
' Excel worksheet functions (UDFs) wrapping the TextDistanceVBA modules.
'
' Import this module into Excel only; the algorithm modules work without it.
' Arguments are declared As String, so Excel converts single cells, numbers
' and blanks to text itself, and returns #VALUE! for multi-cell ranges and
' error values without calling the function.
'
' Each bare name (=HAMMING, =JARO, ...) returns the same value as Python's
' textdistance.<name>(a, b): a distance for edit distances, a similarity for
' Jaro / Jaro-Winkler / Ratcliff-Obershelp / Identity / Matrix, and the
' matching text itself for LCSSEQ / LCSSTR / PREFIX / POSTFIX.
'
' Example:  =LEVENSHTEIN(A2, B2)            -> number of edits
'           =LEVENSHTEIN_NSIM(A2, B2, FALSE) -> case-insensitive similarity 0..1
'           =JARO_WINKLER(A2, B2)           -> similarity 0..1
'
' Requires all algorithm modules (Hamming.bas, Levenshtein.bas, ...) to be
' imported too.
'
' UDF names deliberately differ from the library names (Hamming_Distance etc.)
' because VBA names are case-insensitive and duplicates would be ambiguous.
'===============================================================================
Option Explicit
Option Compare Binary

Public Function HAMMING(ByVal Text1 As String, ByVal Text2 As String, _
                        Optional ByVal CaseSensitive As Boolean = True, _
                        Optional ByVal Truncate As Boolean = False) As Long
    HAMMING = Hamming_Distance(Text1, Text2, CaseSensitive, Truncate)
End Function

Public Function HAMMING_SIM(ByVal Text1 As String, ByVal Text2 As String, _
                            Optional ByVal CaseSensitive As Boolean = True, _
                            Optional ByVal Truncate As Boolean = False) As Long
    HAMMING_SIM = Hamming_Similarity(Text1, Text2, CaseSensitive, Truncate)
End Function

Public Function HAMMING_NDIST(ByVal Text1 As String, ByVal Text2 As String, _
                              Optional ByVal CaseSensitive As Boolean = True, _
                              Optional ByVal Truncate As Boolean = False) As Double
    HAMMING_NDIST = Hamming_NormalizedDistance(Text1, Text2, CaseSensitive, Truncate)
End Function

Public Function HAMMING_NSIM(ByVal Text1 As String, ByVal Text2 As String, _
                             Optional ByVal CaseSensitive As Boolean = True, _
                             Optional ByVal Truncate As Boolean = False) As Double
    HAMMING_NSIM = Hamming_NormalizedSimilarity(Text1, Text2, CaseSensitive, Truncate)
End Function

Public Function LEVENSHTEIN(ByVal Text1 As String, ByVal Text2 As String, _
                            Optional ByVal CaseSensitive As Boolean = True) As Long
    LEVENSHTEIN = Levenshtein_Distance(Text1, Text2, CaseSensitive)
End Function

Public Function LEVENSHTEIN_SIM(ByVal Text1 As String, ByVal Text2 As String, _
                                Optional ByVal CaseSensitive As Boolean = True) As Long
    LEVENSHTEIN_SIM = Levenshtein_Similarity(Text1, Text2, CaseSensitive)
End Function

Public Function LEVENSHTEIN_NDIST(ByVal Text1 As String, ByVal Text2 As String, _
                                  Optional ByVal CaseSensitive As Boolean = True) As Double
    LEVENSHTEIN_NDIST = Levenshtein_NormalizedDistance(Text1, Text2, CaseSensitive)
End Function

Public Function LEVENSHTEIN_NSIM(ByVal Text1 As String, ByVal Text2 As String, _
                                 Optional ByVal CaseSensitive As Boolean = True) As Double
    LEVENSHTEIN_NSIM = Levenshtein_NormalizedSimilarity(Text1, Text2, CaseSensitive)
End Function

Public Function DAMERAU_LEVENSHTEIN(ByVal Text1 As String, ByVal Text2 As String, _
                                    Optional ByVal CaseSensitive As Boolean = True, _
                                    Optional ByVal Restricted As Boolean = True) As Long
    DAMERAU_LEVENSHTEIN = DamerauLevenshtein_Distance(Text1, Text2, CaseSensitive, Restricted)
End Function

Public Function DAMERAU_LEVENSHTEIN_SIM(ByVal Text1 As String, ByVal Text2 As String, _
                                        Optional ByVal CaseSensitive As Boolean = True, _
                                        Optional ByVal Restricted As Boolean = True) As Long
    DAMERAU_LEVENSHTEIN_SIM = DamerauLevenshtein_Similarity(Text1, Text2, CaseSensitive, Restricted)
End Function

Public Function DAMERAU_LEVENSHTEIN_NDIST(ByVal Text1 As String, ByVal Text2 As String, _
                                          Optional ByVal CaseSensitive As Boolean = True, _
                                          Optional ByVal Restricted As Boolean = True) As Double
    DAMERAU_LEVENSHTEIN_NDIST = DamerauLevenshtein_NormalizedDistance(Text1, Text2, CaseSensitive, Restricted)
End Function

Public Function DAMERAU_LEVENSHTEIN_NSIM(ByVal Text1 As String, ByVal Text2 As String, _
                                         Optional ByVal CaseSensitive As Boolean = True, _
                                         Optional ByVal Restricted As Boolean = True) As Double
    DAMERAU_LEVENSHTEIN_NSIM = DamerauLevenshtein_NormalizedSimilarity(Text1, Text2, CaseSensitive, Restricted)
End Function

Public Function JARO(ByVal Text1 As String, ByVal Text2 As String, _
                     Optional ByVal CaseSensitive As Boolean = True, _
                     Optional ByVal LongTolerance As Boolean = False) As Double
    JARO = Jaro_Similarity(Text1, Text2, CaseSensitive, LongTolerance)
End Function

Public Function JARO_DIST(ByVal Text1 As String, ByVal Text2 As String, _
                          Optional ByVal CaseSensitive As Boolean = True, _
                          Optional ByVal LongTolerance As Boolean = False) As Double
    JARO_DIST = Jaro_Distance(Text1, Text2, CaseSensitive, LongTolerance)
End Function

Public Function JARO_WINKLER(ByVal Text1 As String, ByVal Text2 As String, _
                             Optional ByVal CaseSensitive As Boolean = True, _
                             Optional ByVal LongTolerance As Boolean = False, _
                             Optional ByVal PrefixWeight As Double = 0.1) As Double
    JARO_WINKLER = JaroWinkler_Similarity(Text1, Text2, CaseSensitive, LongTolerance, PrefixWeight)
End Function

Public Function JARO_WINKLER_DIST(ByVal Text1 As String, ByVal Text2 As String, _
                                  Optional ByVal CaseSensitive As Boolean = True, _
                                  Optional ByVal LongTolerance As Boolean = False, _
                                  Optional ByVal PrefixWeight As Double = 0.1) As Double
    JARO_WINKLER_DIST = JaroWinkler_Distance(Text1, Text2, CaseSensitive, LongTolerance, PrefixWeight)
End Function

Public Function LCSSEQ(ByVal Text1 As String, ByVal Text2 As String, _
                       Optional ByVal CaseSensitive As Boolean = True) As String
    LCSSEQ = LCSSeq_Sequence(Text1, Text2, CaseSensitive)
End Function

Public Function LCSSEQ_SIM(ByVal Text1 As String, ByVal Text2 As String, _
                           Optional ByVal CaseSensitive As Boolean = True) As Long
    LCSSEQ_SIM = LCSSeq_Similarity(Text1, Text2, CaseSensitive)
End Function

Public Function LCSSEQ_DIST(ByVal Text1 As String, ByVal Text2 As String, _
                            Optional ByVal CaseSensitive As Boolean = True) As Long
    LCSSEQ_DIST = LCSSeq_Distance(Text1, Text2, CaseSensitive)
End Function

Public Function LCSSEQ_NDIST(ByVal Text1 As String, ByVal Text2 As String, _
                             Optional ByVal CaseSensitive As Boolean = True) As Double
    LCSSEQ_NDIST = LCSSeq_NormalizedDistance(Text1, Text2, CaseSensitive)
End Function

Public Function LCSSEQ_NSIM(ByVal Text1 As String, ByVal Text2 As String, _
                            Optional ByVal CaseSensitive As Boolean = True) As Double
    LCSSEQ_NSIM = LCSSeq_NormalizedSimilarity(Text1, Text2, CaseSensitive)
End Function

Public Function LCSSTR(ByVal Text1 As String, ByVal Text2 As String, _
                       Optional ByVal CaseSensitive As Boolean = True) As String
    LCSSTR = LCSStr_Substring(Text1, Text2, CaseSensitive)
End Function

Public Function LCSSTR_SIM(ByVal Text1 As String, ByVal Text2 As String, _
                           Optional ByVal CaseSensitive As Boolean = True) As Long
    LCSSTR_SIM = LCSStr_Similarity(Text1, Text2, CaseSensitive)
End Function

Public Function LCSSTR_DIST(ByVal Text1 As String, ByVal Text2 As String, _
                            Optional ByVal CaseSensitive As Boolean = True) As Long
    LCSSTR_DIST = LCSStr_Distance(Text1, Text2, CaseSensitive)
End Function

Public Function LCSSTR_NDIST(ByVal Text1 As String, ByVal Text2 As String, _
                             Optional ByVal CaseSensitive As Boolean = True) As Double
    LCSSTR_NDIST = LCSStr_NormalizedDistance(Text1, Text2, CaseSensitive)
End Function

Public Function LCSSTR_NSIM(ByVal Text1 As String, ByVal Text2 As String, _
                            Optional ByVal CaseSensitive As Boolean = True) As Double
    LCSSTR_NSIM = LCSStr_NormalizedSimilarity(Text1, Text2, CaseSensitive)
End Function

Public Function RATCLIFF_OBERSHELP(ByVal Text1 As String, ByVal Text2 As String, _
                                   Optional ByVal CaseSensitive As Boolean = True) As Double
    RATCLIFF_OBERSHELP = RatcliffObershelp_Similarity(Text1, Text2, CaseSensitive)
End Function

Public Function RATCLIFF_OBERSHELP_DIST(ByVal Text1 As String, ByVal Text2 As String, _
                                        Optional ByVal CaseSensitive As Boolean = True) As Double
    RATCLIFF_OBERSHELP_DIST = RatcliffObershelp_Distance(Text1, Text2, CaseSensitive)
End Function

Public Function PREFIX(ByVal Text1 As String, ByVal Text2 As String, _
                       Optional ByVal CaseSensitive As Boolean = True) As String
    PREFIX = Prefix_CommonPrefix(Text1, Text2, CaseSensitive)
End Function

Public Function PREFIX_SIM(ByVal Text1 As String, ByVal Text2 As String, _
                           Optional ByVal CaseSensitive As Boolean = True) As Long
    PREFIX_SIM = Prefix_Similarity(Text1, Text2, CaseSensitive)
End Function

Public Function PREFIX_DIST(ByVal Text1 As String, ByVal Text2 As String, _
                            Optional ByVal CaseSensitive As Boolean = True) As Long
    PREFIX_DIST = Prefix_Distance(Text1, Text2, CaseSensitive)
End Function

Public Function PREFIX_NDIST(ByVal Text1 As String, ByVal Text2 As String, _
                             Optional ByVal CaseSensitive As Boolean = True) As Double
    PREFIX_NDIST = Prefix_NormalizedDistance(Text1, Text2, CaseSensitive)
End Function

Public Function PREFIX_NSIM(ByVal Text1 As String, ByVal Text2 As String, _
                            Optional ByVal CaseSensitive As Boolean = True) As Double
    PREFIX_NSIM = Prefix_NormalizedSimilarity(Text1, Text2, CaseSensitive)
End Function

Public Function POSTFIX(ByVal Text1 As String, ByVal Text2 As String, _
                        Optional ByVal CaseSensitive As Boolean = True) As String
    POSTFIX = Postfix_CommonPostfix(Text1, Text2, CaseSensitive)
End Function

Public Function POSTFIX_SIM(ByVal Text1 As String, ByVal Text2 As String, _
                            Optional ByVal CaseSensitive As Boolean = True) As Long
    POSTFIX_SIM = Postfix_Similarity(Text1, Text2, CaseSensitive)
End Function

Public Function POSTFIX_DIST(ByVal Text1 As String, ByVal Text2 As String, _
                             Optional ByVal CaseSensitive As Boolean = True) As Long
    POSTFIX_DIST = Postfix_Distance(Text1, Text2, CaseSensitive)
End Function

Public Function POSTFIX_NDIST(ByVal Text1 As String, ByVal Text2 As String, _
                              Optional ByVal CaseSensitive As Boolean = True) As Double
    POSTFIX_NDIST = Postfix_NormalizedDistance(Text1, Text2, CaseSensitive)
End Function

Public Function POSTFIX_NSIM(ByVal Text1 As String, ByVal Text2 As String, _
                             Optional ByVal CaseSensitive As Boolean = True) As Double
    POSTFIX_NSIM = Postfix_NormalizedSimilarity(Text1, Text2, CaseSensitive)
End Function

Public Function LENGTH(ByVal Text1 As String, ByVal Text2 As String, _
                       Optional ByVal CaseSensitive As Boolean = True) As Long
    LENGTH = Length_Distance(Text1, Text2, CaseSensitive)
End Function

Public Function LENGTH_SIM(ByVal Text1 As String, ByVal Text2 As String, _
                           Optional ByVal CaseSensitive As Boolean = True) As Long
    LENGTH_SIM = Length_Similarity(Text1, Text2, CaseSensitive)
End Function

Public Function LENGTH_NDIST(ByVal Text1 As String, ByVal Text2 As String, _
                             Optional ByVal CaseSensitive As Boolean = True) As Double
    LENGTH_NDIST = Length_NormalizedDistance(Text1, Text2, CaseSensitive)
End Function

Public Function LENGTH_NSIM(ByVal Text1 As String, ByVal Text2 As String, _
                            Optional ByVal CaseSensitive As Boolean = True) As Double
    LENGTH_NSIM = Length_NormalizedSimilarity(Text1, Text2, CaseSensitive)
End Function

Public Function IDENTITY(ByVal Text1 As String, ByVal Text2 As String, _
                         Optional ByVal CaseSensitive As Boolean = True) As Long
    IDENTITY = Identity_Similarity(Text1, Text2, CaseSensitive)
End Function

Public Function IDENTITY_DIST(ByVal Text1 As String, ByVal Text2 As String, _
                              Optional ByVal CaseSensitive As Boolean = True) As Long
    IDENTITY_DIST = Identity_Distance(Text1, Text2, CaseSensitive)
End Function

' Table: optional 3-column range (string 1, string 2, score).
Public Function MATRIX(ByVal Text1 As String, ByVal Text2 As String, _
                       Optional ByVal Table As Variant, _
                       Optional ByVal MismatchCost As Double = 0, _
                       Optional ByVal MatchCost As Double = 1, _
                       Optional ByVal Symmetric As Boolean = True, _
                       Optional ByVal CaseSensitive As Boolean = True) As Double
    MATRIX = Matrix_Similarity(Text1, Text2, Table, MismatchCost, MatchCost, Symmetric, CaseSensitive)
End Function

' Table: optional 3-column range (string 1, string 2, score).
Public Function MATRIX_DIST(ByVal Text1 As String, ByVal Text2 As String, _
                            Optional ByVal Table As Variant, _
                            Optional ByVal MismatchCost As Double = 0, _
                            Optional ByVal MatchCost As Double = 1, _
                            Optional ByVal Symmetric As Boolean = True, _
                            Optional ByVal CaseSensitive As Boolean = True) As Double
    MATRIX_DIST = Matrix_Distance(Text1, Text2, Table, MismatchCost, MatchCost, Symmetric, CaseSensitive)
End Function
