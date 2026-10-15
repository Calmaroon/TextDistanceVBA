Attribute VB_Name = "modExcelUDF"
'===============================================================================
' Excel worksheet functions (UDFs) wrapping the TextDistanceVBA modules.
'
' Import this module into Excel only; the algorithm modules work without it.
' These wrappers accept cell references, numbers and blanks, and return
' #VALUE! instead of raising a VBA error on bad input (e.g. a multi-cell range).
'
' Example:  =HAMMING(A2, B2)          -> distance
'           =HAMMING_NSIM(A2, B2, FALSE) -> case-insensitive similarity 0..1
'
' UDF names deliberately differ from the library names (Hamming_Distance etc.)
' because VBA names are case-insensitive and duplicates would be ambiguous.
'===============================================================================
Option Explicit
Option Compare Binary

Public Function HAMMING(ByVal Text1 As Variant, ByVal Text2 As Variant, _
                        Optional ByVal CaseSensitive As Boolean = True, _
                        Optional ByVal Truncate As Boolean = False) As Variant
    Dim s1 As String, s2 As String
    If Not ToText(Text1, s1) Or Not ToText(Text2, s2) Then HAMMING = CVErr(xlErrValue): Exit Function
    HAMMING = Hamming_Distance(s1, s2, CaseSensitive, Truncate)
End Function

Public Function HAMMING_SIM(ByVal Text1 As Variant, ByVal Text2 As Variant, _
                            Optional ByVal CaseSensitive As Boolean = True, _
                            Optional ByVal Truncate As Boolean = False) As Variant
    Dim s1 As String, s2 As String
    If Not ToText(Text1, s1) Or Not ToText(Text2, s2) Then HAMMING_SIM = CVErr(xlErrValue): Exit Function
    HAMMING_SIM = Hamming_Similarity(s1, s2, CaseSensitive, Truncate)
End Function

Public Function HAMMING_NDIST(ByVal Text1 As Variant, ByVal Text2 As Variant, _
                              Optional ByVal CaseSensitive As Boolean = True, _
                              Optional ByVal Truncate As Boolean = False) As Variant
    Dim s1 As String, s2 As String
    If Not ToText(Text1, s1) Or Not ToText(Text2, s2) Then HAMMING_NDIST = CVErr(xlErrValue): Exit Function
    HAMMING_NDIST = Hamming_NormalizedDistance(s1, s2, CaseSensitive, Truncate)
End Function

Public Function HAMMING_NSIM(ByVal Text1 As Variant, ByVal Text2 As Variant, _
                             Optional ByVal CaseSensitive As Boolean = True, _
                             Optional ByVal Truncate As Boolean = False) As Variant
    Dim s1 As String, s2 As String
    If Not ToText(Text1, s1) Or Not ToText(Text2, s2) Then HAMMING_NSIM = CVErr(xlErrValue): Exit Function
    HAMMING_NSIM = Hamming_NormalizedSimilarity(s1, s2, CaseSensitive, Truncate)
End Function

' Converts a worksheet argument to text. Returns False for errors, arrays and
' multi-cell ranges.
Private Function ToText(ByVal Value As Variant, ByRef Result As String) As Boolean
    If IsObject(Value) Then
        If TypeName(Value) <> "Range" Then Exit Function
        If Value.Cells.CountLarge <> 1 Then Exit Function
        Value = Value.Value
    End If
    If IsError(Value) Or IsArray(Value) Then Exit Function
    If IsEmpty(Value) Or IsNull(Value) Then
        Result = vbNullString
    Else
        Result = CStr(Value)
    End If
    ToText = True
End Function
