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
' textdistance.<name>(a, b): a distance for edit distances, Bag and the
' compression measures (..._NCD), a similarity for Jaro / Jaro-Winkler /
' StrCmp95 / MLIPNS / the alignment scores / Ratcliff-Obershelp / Identity /
' Matrix / MRA and the token measures (Cosine, Jaccard, ...), and the
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

' MRA codex of a name (named MRACODEX because VBA names ignore case and
' MRA_Codex is the library function).
Public Function MRACODEX(ByVal Text As String) As String
    MRACODEX = MRA_Codex(Text)
End Function

Public Function MRA(ByVal Text1 As String, ByVal Text2 As String) As Long
    MRA = MRA_Similarity(Text1, Text2)
End Function

Public Function MRA_DIST(ByVal Text1 As String, ByVal Text2 As String) As Long
    MRA_DIST = MRA_Distance(Text1, Text2)
End Function

Public Function MRA_NDIST(ByVal Text1 As String, ByVal Text2 As String) As Double
    MRA_NDIST = MRA_NormalizedDistance(Text1, Text2)
End Function

Public Function MRA_NSIM(ByVal Text1 As String, ByVal Text2 As String) As Double
    MRA_NSIM = MRA_NormalizedSimilarity(Text1, Text2)
End Function

Public Function EDITEX(ByVal Text1 As String, ByVal Text2 As String, _
                       Optional ByVal LocalMode As Boolean = False, _
                       Optional ByVal MatchCost As Long = 0, _
                       Optional ByVal GroupCost As Long = 1, _
                       Optional ByVal MismatchCost As Long = 2) As Long
    EDITEX = Editex_Distance(Text1, Text2, LocalMode, MatchCost, GroupCost, MismatchCost)
End Function

Public Function EDITEX_SIM(ByVal Text1 As String, ByVal Text2 As String, _
                           Optional ByVal LocalMode As Boolean = False, _
                           Optional ByVal MatchCost As Long = 0, _
                           Optional ByVal GroupCost As Long = 1, _
                           Optional ByVal MismatchCost As Long = 2) As Long
    EDITEX_SIM = Editex_Similarity(Text1, Text2, LocalMode, MatchCost, GroupCost, MismatchCost)
End Function

Public Function EDITEX_NDIST(ByVal Text1 As String, ByVal Text2 As String, _
                             Optional ByVal LocalMode As Boolean = False, _
                             Optional ByVal MatchCost As Long = 0, _
                             Optional ByVal GroupCost As Long = 1, _
                             Optional ByVal MismatchCost As Long = 2) As Double
    EDITEX_NDIST = Editex_NormalizedDistance(Text1, Text2, LocalMode, MatchCost, GroupCost, MismatchCost)
End Function

Public Function EDITEX_NSIM(ByVal Text1 As String, ByVal Text2 As String, _
                            Optional ByVal LocalMode As Boolean = False, _
                            Optional ByVal MatchCost As Long = 0, _
                            Optional ByVal GroupCost As Long = 1, _
                            Optional ByVal MismatchCost As Long = 2) As Double
    EDITEX_NSIM = Editex_NormalizedSimilarity(Text1, Text2, LocalMode, MatchCost, GroupCost, MismatchCost)
End Function

Public Function COSINE(ByVal Text1 As String, ByVal Text2 As String, _
                       Optional ByVal CaseSensitive As Boolean = True, _
                       Optional ByVal Qval As Long = 1, _
                       Optional ByVal AsSet As Boolean = False) As Double
    COSINE = Cosine_Similarity(Text1, Text2, CaseSensitive, Qval, AsSet)
End Function

Public Function COSINE_DIST(ByVal Text1 As String, ByVal Text2 As String, _
                            Optional ByVal CaseSensitive As Boolean = True, _
                            Optional ByVal Qval As Long = 1, _
                            Optional ByVal AsSet As Boolean = False) As Double
    COSINE_DIST = Cosine_Distance(Text1, Text2, CaseSensitive, Qval, AsSet)
End Function

Public Function JACCARD(ByVal Text1 As String, ByVal Text2 As String, _
                        Optional ByVal CaseSensitive As Boolean = True, _
                        Optional ByVal Qval As Long = 1, _
                        Optional ByVal AsSet As Boolean = False) As Double
    JACCARD = Jaccard_Similarity(Text1, Text2, CaseSensitive, Qval, AsSet)
End Function

Public Function JACCARD_DIST(ByVal Text1 As String, ByVal Text2 As String, _
                             Optional ByVal CaseSensitive As Boolean = True, _
                             Optional ByVal Qval As Long = 1, _
                             Optional ByVal AsSet As Boolean = False) As Double
    JACCARD_DIST = Jaccard_Distance(Text1, Text2, CaseSensitive, Qval, AsSet)
End Function

Public Function SORENSEN(ByVal Text1 As String, ByVal Text2 As String, _
                         Optional ByVal CaseSensitive As Boolean = True, _
                         Optional ByVal Qval As Long = 1, _
                         Optional ByVal AsSet As Boolean = False) As Double
    SORENSEN = Sorensen_Similarity(Text1, Text2, CaseSensitive, Qval, AsSet)
End Function

Public Function SORENSEN_DIST(ByVal Text1 As String, ByVal Text2 As String, _
                              Optional ByVal CaseSensitive As Boolean = True, _
                              Optional ByVal Qval As Long = 1, _
                              Optional ByVal AsSet As Boolean = False) As Double
    SORENSEN_DIST = Sorensen_Distance(Text1, Text2, CaseSensitive, Qval, AsSet)
End Function

Public Function OVERLAP(ByVal Text1 As String, ByVal Text2 As String, _
                        Optional ByVal CaseSensitive As Boolean = True, _
                        Optional ByVal Qval As Long = 1, _
                        Optional ByVal AsSet As Boolean = False) As Double
    OVERLAP = Overlap_Similarity(Text1, Text2, CaseSensitive, Qval, AsSet)
End Function

Public Function OVERLAP_DIST(ByVal Text1 As String, ByVal Text2 As String, _
                             Optional ByVal CaseSensitive As Boolean = True, _
                             Optional ByVal Qval As Long = 1, _
                             Optional ByVal AsSet As Boolean = False) As Double
    OVERLAP_DIST = Overlap_Distance(Text1, Text2, CaseSensitive, Qval, AsSet)
End Function

Public Function TANIMOTO(ByVal Text1 As String, ByVal Text2 As String, _
                         Optional ByVal CaseSensitive As Boolean = True, _
                         Optional ByVal Qval As Long = 1, _
                         Optional ByVal AsSet As Boolean = False) As Double
    TANIMOTO = Tanimoto_Similarity(Text1, Text2, CaseSensitive, Qval, AsSet)
End Function

Public Function TANIMOTO_DIST(ByVal Text1 As String, ByVal Text2 As String, _
                              Optional ByVal CaseSensitive As Boolean = True, _
                              Optional ByVal Qval As Long = 1, _
                              Optional ByVal AsSet As Boolean = False) As Double
    TANIMOTO_DIST = Tanimoto_Distance(Text1, Text2, CaseSensitive, Qval, AsSet)
End Function

Public Function TVERSKY(ByVal Text1 As String, ByVal Text2 As String, _
                        Optional ByVal CaseSensitive As Boolean = True, _
                        Optional ByVal Qval As Long = 1, _
                        Optional ByVal AsSet As Boolean = False, _
                        Optional ByVal Alpha As Double = 1, _
                        Optional ByVal Beta As Double = 1, _
                        Optional ByVal Bias As Variant) As Double
    TVERSKY = Tversky_Similarity(Text1, Text2, CaseSensitive, Qval, AsSet, Alpha, Beta, Bias)
End Function

Public Function TVERSKY_DIST(ByVal Text1 As String, ByVal Text2 As String, _
                             Optional ByVal CaseSensitive As Boolean = True, _
                             Optional ByVal Qval As Long = 1, _
                             Optional ByVal AsSet As Boolean = False, _
                             Optional ByVal Alpha As Double = 1, _
                             Optional ByVal Beta As Double = 1, _
                             Optional ByVal Bias As Variant) As Double
    TVERSKY_DIST = Tversky_Distance(Text1, Text2, CaseSensitive, Qval, AsSet, Alpha, Beta, Bias)
End Function

Public Function MONGE_ELKAN(ByVal Text1 As String, ByVal Text2 As String, _
                            Optional ByVal CaseSensitive As Boolean = True, _
                            Optional ByVal Qval As Long = 1, _
                            Optional ByVal Symmetric As Boolean = False) As Double
    MONGE_ELKAN = MongeElkan_Similarity(Text1, Text2, CaseSensitive, Qval, Symmetric)
End Function

Public Function MONGE_ELKAN_DIST(ByVal Text1 As String, ByVal Text2 As String, _
                                 Optional ByVal CaseSensitive As Boolean = True, _
                                 Optional ByVal Qval As Long = 1, _
                                 Optional ByVal Symmetric As Boolean = False) As Double
    MONGE_ELKAN_DIST = MongeElkan_Distance(Text1, Text2, CaseSensitive, Qval, Symmetric)
End Function

Public Function BAG(ByVal Text1 As String, ByVal Text2 As String, _
                    Optional ByVal CaseSensitive As Boolean = True, _
                    Optional ByVal Qval As Long = 1) As Long
    BAG = Bag_Distance(Text1, Text2, CaseSensitive, Qval)
End Function

Public Function BAG_SIM(ByVal Text1 As String, ByVal Text2 As String, _
                        Optional ByVal CaseSensitive As Boolean = True, _
                        Optional ByVal Qval As Long = 1) As Long
    BAG_SIM = Bag_Similarity(Text1, Text2, CaseSensitive, Qval)
End Function

Public Function BAG_NDIST(ByVal Text1 As String, ByVal Text2 As String, _
                          Optional ByVal CaseSensitive As Boolean = True, _
                          Optional ByVal Qval As Long = 1) As Double
    BAG_NDIST = Bag_NormalizedDistance(Text1, Text2, CaseSensitive, Qval)
End Function

Public Function BAG_NSIM(ByVal Text1 As String, ByVal Text2 As String, _
                         Optional ByVal CaseSensitive As Boolean = True, _
                         Optional ByVal Qval As Long = 1) As Double
    BAG_NSIM = Bag_NormalizedSimilarity(Text1, Text2, CaseSensitive, Qval)
End Function

Public Function STRCMP95(ByVal Text1 As String, ByVal Text2 As String, _
                         Optional ByVal LongStrings As Boolean = False) As Double
    STRCMP95 = StrCmp95_Similarity(Text1, Text2, LongStrings)
End Function

Public Function STRCMP95_DIST(ByVal Text1 As String, ByVal Text2 As String, _
                              Optional ByVal LongStrings As Boolean = False) As Double
    STRCMP95_DIST = StrCmp95_Distance(Text1, Text2, LongStrings)
End Function

Public Function NEEDLEMAN_WUNSCH(ByVal Text1 As String, ByVal Text2 As String, _
                                 Optional ByVal CaseSensitive As Boolean = True, _
                                 Optional ByVal Qval As Long = 1, _
                                 Optional ByVal GapCost As Double = 1) As Double
    NEEDLEMAN_WUNSCH = NeedlemanWunsch_Similarity(Text1, Text2, CaseSensitive, Qval, GapCost)
End Function

Public Function NEEDLEMAN_WUNSCH_DIST(ByVal Text1 As String, ByVal Text2 As String, _
                                      Optional ByVal CaseSensitive As Boolean = True, _
                                      Optional ByVal Qval As Long = 1, _
                                      Optional ByVal GapCost As Double = 1) As Double
    NEEDLEMAN_WUNSCH_DIST = NeedlemanWunsch_Distance(Text1, Text2, CaseSensitive, Qval, GapCost)
End Function

Public Function NEEDLEMAN_WUNSCH_NSIM(ByVal Text1 As String, ByVal Text2 As String, _
                                      Optional ByVal CaseSensitive As Boolean = True, _
                                      Optional ByVal Qval As Long = 1, _
                                      Optional ByVal GapCost As Double = 1) As Double
    NEEDLEMAN_WUNSCH_NSIM = NeedlemanWunsch_NormalizedSimilarity(Text1, Text2, CaseSensitive, Qval, GapCost)
End Function

Public Function NEEDLEMAN_WUNSCH_NDIST(ByVal Text1 As String, ByVal Text2 As String, _
                                       Optional ByVal CaseSensitive As Boolean = True, _
                                       Optional ByVal Qval As Long = 1, _
                                       Optional ByVal GapCost As Double = 1) As Double
    NEEDLEMAN_WUNSCH_NDIST = NeedlemanWunsch_NormalizedDistance(Text1, Text2, CaseSensitive, Qval, GapCost)
End Function

Public Function GOTOH(ByVal Text1 As String, ByVal Text2 As String, _
                      Optional ByVal CaseSensitive As Boolean = True, _
                      Optional ByVal Qval As Long = 1, _
                      Optional ByVal GapOpen As Double = 1, _
                      Optional ByVal GapExt As Double = 0.4) As Double
    GOTOH = Gotoh_Similarity(Text1, Text2, CaseSensitive, Qval, GapOpen, GapExt)
End Function

Public Function GOTOH_DIST(ByVal Text1 As String, ByVal Text2 As String, _
                           Optional ByVal CaseSensitive As Boolean = True, _
                           Optional ByVal Qval As Long = 1, _
                           Optional ByVal GapOpen As Double = 1, _
                           Optional ByVal GapExt As Double = 0.4) As Double
    GOTOH_DIST = Gotoh_Distance(Text1, Text2, CaseSensitive, Qval, GapOpen, GapExt)
End Function

Public Function GOTOH_NSIM(ByVal Text1 As String, ByVal Text2 As String, _
                           Optional ByVal CaseSensitive As Boolean = True, _
                           Optional ByVal Qval As Long = 1, _
                           Optional ByVal GapOpen As Double = 1, _
                           Optional ByVal GapExt As Double = 0.4) As Double
    GOTOH_NSIM = Gotoh_NormalizedSimilarity(Text1, Text2, CaseSensitive, Qval, GapOpen, GapExt)
End Function

Public Function GOTOH_NDIST(ByVal Text1 As String, ByVal Text2 As String, _
                            Optional ByVal CaseSensitive As Boolean = True, _
                            Optional ByVal Qval As Long = 1, _
                            Optional ByVal GapOpen As Double = 1, _
                            Optional ByVal GapExt As Double = 0.4) As Double
    GOTOH_NDIST = Gotoh_NormalizedDistance(Text1, Text2, CaseSensitive, Qval, GapOpen, GapExt)
End Function

Public Function SMITH_WATERMAN(ByVal Text1 As String, ByVal Text2 As String, _
                               Optional ByVal CaseSensitive As Boolean = True, _
                               Optional ByVal Qval As Long = 1, _
                               Optional ByVal GapCost As Double = 1) As Double
    SMITH_WATERMAN = SmithWaterman_Similarity(Text1, Text2, CaseSensitive, Qval, GapCost)
End Function

Public Function SMITH_WATERMAN_DIST(ByVal Text1 As String, ByVal Text2 As String, _
                                    Optional ByVal CaseSensitive As Boolean = True, _
                                    Optional ByVal Qval As Long = 1, _
                                    Optional ByVal GapCost As Double = 1) As Double
    SMITH_WATERMAN_DIST = SmithWaterman_Distance(Text1, Text2, CaseSensitive, Qval, GapCost)
End Function

Public Function SMITH_WATERMAN_NSIM(ByVal Text1 As String, ByVal Text2 As String, _
                                    Optional ByVal CaseSensitive As Boolean = True, _
                                    Optional ByVal Qval As Long = 1, _
                                    Optional ByVal GapCost As Double = 1) As Double
    SMITH_WATERMAN_NSIM = SmithWaterman_NormalizedSimilarity(Text1, Text2, CaseSensitive, Qval, GapCost)
End Function

Public Function SMITH_WATERMAN_NDIST(ByVal Text1 As String, ByVal Text2 As String, _
                                     Optional ByVal CaseSensitive As Boolean = True, _
                                     Optional ByVal Qval As Long = 1, _
                                     Optional ByVal GapCost As Double = 1) As Double
    SMITH_WATERMAN_NDIST = SmithWaterman_NormalizedDistance(Text1, Text2, CaseSensitive, Qval, GapCost)
End Function

Public Function MLIPNS(ByVal Text1 As String, ByVal Text2 As String, _
                       Optional ByVal CaseSensitive As Boolean = True, _
                       Optional ByVal Qval As Long = 1, _
                       Optional ByVal Threshold As Double = 0.25, _
                       Optional ByVal MaxMismatches As Long = 2) As Long
    MLIPNS = MLIPNS_Similarity(Text1, Text2, CaseSensitive, Qval, Threshold, MaxMismatches)
End Function

Public Function ARITH_NCD(ByVal Text1 As String, ByVal Text2 As String, _
                          Optional ByVal CaseSensitive As Boolean = True, _
                          Optional ByVal LogBase As Double = 2) As Double
    ARITH_NCD = ArithNCD_Distance(Text1, Text2, CaseSensitive, LogBase)
End Function

Public Function ARITH_NCD_SIM(ByVal Text1 As String, ByVal Text2 As String, _
                              Optional ByVal CaseSensitive As Boolean = True, _
                              Optional ByVal LogBase As Double = 2) As Double
    ARITH_NCD_SIM = ArithNCD_Similarity(Text1, Text2, CaseSensitive, LogBase)
End Function

Public Function RLE_NCD(ByVal Text1 As String, ByVal Text2 As String, _
                        Optional ByVal CaseSensitive As Boolean = True) As Double
    RLE_NCD = RLENCD_Distance(Text1, Text2, CaseSensitive)
End Function

Public Function RLE_NCD_SIM(ByVal Text1 As String, ByVal Text2 As String, _
                            Optional ByVal CaseSensitive As Boolean = True) As Double
    RLE_NCD_SIM = RLENCD_Similarity(Text1, Text2, CaseSensitive)
End Function

Public Function BWTRLE_NCD(ByVal Text1 As String, ByVal Text2 As String, _
                           Optional ByVal CaseSensitive As Boolean = True) As Double
    BWTRLE_NCD = BWTRLENCD_Distance(Text1, Text2, CaseSensitive)
End Function

Public Function BWTRLE_NCD_SIM(ByVal Text1 As String, ByVal Text2 As String, _
                               Optional ByVal CaseSensitive As Boolean = True) As Double
    BWTRLE_NCD_SIM = BWTRLENCD_Similarity(Text1, Text2, CaseSensitive)
End Function

Public Function SQRT_NCD(ByVal Text1 As String, ByVal Text2 As String, _
                         Optional ByVal CaseSensitive As Boolean = True, _
                         Optional ByVal Qval As Long = 1) As Double
    SQRT_NCD = SqrtNCD_Distance(Text1, Text2, CaseSensitive, Qval)
End Function

Public Function SQRT_NCD_SIM(ByVal Text1 As String, ByVal Text2 As String, _
                             Optional ByVal CaseSensitive As Boolean = True, _
                             Optional ByVal Qval As Long = 1) As Double
    SQRT_NCD_SIM = SqrtNCD_Similarity(Text1, Text2, CaseSensitive, Qval)
End Function

Public Function ENTROPY_NCD(ByVal Text1 As String, ByVal Text2 As String, _
                            Optional ByVal CaseSensitive As Boolean = True, _
                            Optional ByVal Qval As Long = 1, _
                            Optional ByVal Coef As Double = 1, _
                            Optional ByVal LogBase As Double = 2) As Double
    ENTROPY_NCD = EntropyNCD_Distance(Text1, Text2, CaseSensitive, Qval, Coef, LogBase)
End Function

Public Function ENTROPY_NCD_SIM(ByVal Text1 As String, ByVal Text2 As String, _
                                Optional ByVal CaseSensitive As Boolean = True, _
                                Optional ByVal Qval As Long = 1, _
                                Optional ByVal Coef As Double = 1, _
                                Optional ByVal LogBase As Double = 2) As Double
    ENTROPY_NCD_SIM = EntropyNCD_Similarity(Text1, Text2, CaseSensitive, Qval, Coef, LogBase)
End Function
