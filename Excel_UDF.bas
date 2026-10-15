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
' Jaro / Jaro-Winkler.
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
