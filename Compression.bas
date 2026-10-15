Attribute VB_Name = "modCompression"
'===============================================================================
' Normalized compression distances - port of textdistance.ArithNCD, RLENCD,
' BWTRLENCD, SqrtNCD and EntropyNCD (Python textdistance 4.6)
'
' NCD compares how well the two strings compress together versus apart:
'   distance = (C(s1 + s2) - Min(C(s1), C(s2))) / Max(C(s1), C(s2))
' where C is a "compressed size" and C(s1 + s2) is the smaller of the two
' concatenation orders. Distance is the primary value; the maximum is 1, so
' Similarity = 1 - Distance.
'
'   ArithNCD    C = bits needed by arithmetic coding (in digits of LogBase)
'   RLENCD      C = length after run-length encoding ("aaab" -> "3ab")
'   BWTRLENCD   C = length after a Burrows-Wheeler transform, then RLE
'   SqrtNCD     C = sum of the square roots of each token's count
'   EntropyNCD  C = Coef + Shannon entropy (in LogBase) of the tokens
'
' textdistance's zlib, bz2 and LZMA versions need compression libraries VBA
' does not have, so they are not included.
'
' Pure VBA with no host references, so it works in Excel, Access, Word, etc.
'
' Common arguments
'   s1, s2         Strings to compare.
'   CaseSensitive  True (default) compares exactly; False ignores case.
'                  (VBA addition - Python's textdistance has no such flag.)
'   Qval           SqrtNCD and EntropyNCD only: 1 (default) = characters,
'                  2 or more = overlapping character groups, 0 = words.
'===============================================================================
Option Explicit
Option Compare Binary

Private Const LIMB As Double = 10000000#    ' big integers: base-10^7 digits

' Which compressor Ncd uses
Private Const ncdArith As Long = 1
Private Const ncdRle As Long = 2
Private Const ncdBwtRle As Long = 3
Private Const ncdSqrt As Long = 4
Private Const ncdEntropy As Long = 5

'===============================================================================
' ArithNCD
'   LogBase        Digits the size is counted in (default 2 = bits).
'===============================================================================
Public Function ArithNCD_Distance(ByVal s1 As String, ByVal s2 As String, _
                                  Optional ByVal CaseSensitive As Boolean = True, _
                                  Optional ByVal LogBase As Double = 2) As Double
    ArithNCD_Distance = Ncd(ncdArith, s1, s2, CaseSensitive, 1, LogBase, 0, "")
End Function

Public Function ArithNCD_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                    Optional ByVal CaseSensitive As Boolean = True, _
                                    Optional ByVal LogBase As Double = 2) As Double
    ArithNCD_Similarity = 1 - ArithNCD_Distance(s1, s2, CaseSensitive, LogBase)
End Function

Public Function ArithNCD_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                            Optional ByVal CaseSensitive As Boolean = True, _
                                            Optional ByVal LogBase As Double = 2) As Double
    ArithNCD_NormalizedDistance = ArithNCD_Distance(s1, s2, CaseSensitive, LogBase)
End Function

Public Function ArithNCD_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                              Optional ByVal CaseSensitive As Boolean = True, _
                                              Optional ByVal LogBase As Double = 2) As Double
    ArithNCD_NormalizedSimilarity = 1 - ArithNCD_Distance(s1, s2, CaseSensitive, LogBase)
End Function

' Always 1.
Public Function ArithNCD_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    ArithNCD_Maximum = 1
End Function

'===============================================================================
' RLENCD
'===============================================================================
Public Function RLENCD_Distance(ByVal s1 As String, ByVal s2 As String, _
                                Optional ByVal CaseSensitive As Boolean = True) As Double
    RLENCD_Distance = Ncd(ncdRle, s1, s2, CaseSensitive, 1, 0, 0, "")
End Function

Public Function RLENCD_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                  Optional ByVal CaseSensitive As Boolean = True) As Double
    RLENCD_Similarity = 1 - RLENCD_Distance(s1, s2, CaseSensitive)
End Function

Public Function RLENCD_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                          Optional ByVal CaseSensitive As Boolean = True) As Double
    RLENCD_NormalizedDistance = RLENCD_Distance(s1, s2, CaseSensitive)
End Function

Public Function RLENCD_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                            Optional ByVal CaseSensitive As Boolean = True) As Double
    RLENCD_NormalizedSimilarity = 1 - RLENCD_Distance(s1, s2, CaseSensitive)
End Function

' Always 1.
Public Function RLENCD_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    RLENCD_Maximum = 1
End Function

'===============================================================================
' BWTRLENCD
'   Terminator     End-of-text marker for the transform (default ChrW(0)).
'                  Text that already contains it is not transformed.
'===============================================================================
Public Function BWTRLENCD_Distance(ByVal s1 As String, ByVal s2 As String, _
                                   Optional ByVal CaseSensitive As Boolean = True, _
                                   Optional ByVal Terminator As Variant) As Double
    Dim term As String
    If IsMissing(Terminator) Then
        term = ChrW(0)
    Else
        term = CStr(Terminator)
    End If
    BWTRLENCD_Distance = Ncd(ncdBwtRle, s1, s2, CaseSensitive, 1, 0, 0, term)
End Function

Public Function BWTRLENCD_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                     Optional ByVal CaseSensitive As Boolean = True, _
                                     Optional ByVal Terminator As Variant) As Double
    BWTRLENCD_Similarity = 1 - BWTRLENCD_Distance(s1, s2, CaseSensitive, Terminator)
End Function

Public Function BWTRLENCD_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                             Optional ByVal CaseSensitive As Boolean = True, _
                                             Optional ByVal Terminator As Variant) As Double
    BWTRLENCD_NormalizedDistance = BWTRLENCD_Distance(s1, s2, CaseSensitive, Terminator)
End Function

Public Function BWTRLENCD_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                               Optional ByVal CaseSensitive As Boolean = True, _
                                               Optional ByVal Terminator As Variant) As Double
    BWTRLENCD_NormalizedSimilarity = 1 - BWTRLENCD_Distance(s1, s2, CaseSensitive, Terminator)
End Function

' Always 1.
Public Function BWTRLENCD_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    BWTRLENCD_Maximum = 1
End Function

'===============================================================================
' SqrtNCD
'===============================================================================
Public Function SqrtNCD_Distance(ByVal s1 As String, ByVal s2 As String, _
                                 Optional ByVal CaseSensitive As Boolean = True, _
                                 Optional ByVal Qval As Long = 1) As Double
    SqrtNCD_Distance = Ncd(ncdSqrt, s1, s2, CaseSensitive, Qval, 0, 0, "")
End Function

Public Function SqrtNCD_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                   Optional ByVal CaseSensitive As Boolean = True, _
                                   Optional ByVal Qval As Long = 1) As Double
    SqrtNCD_Similarity = 1 - SqrtNCD_Distance(s1, s2, CaseSensitive, Qval)
End Function

Public Function SqrtNCD_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                           Optional ByVal CaseSensitive As Boolean = True, _
                                           Optional ByVal Qval As Long = 1) As Double
    SqrtNCD_NormalizedDistance = SqrtNCD_Distance(s1, s2, CaseSensitive, Qval)
End Function

Public Function SqrtNCD_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                             Optional ByVal CaseSensitive As Boolean = True, _
                                             Optional ByVal Qval As Long = 1) As Double
    SqrtNCD_NormalizedSimilarity = 1 - SqrtNCD_Distance(s1, s2, CaseSensitive, Qval)
End Function

' Always 1.
Public Function SqrtNCD_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    SqrtNCD_Maximum = 1
End Function

'===============================================================================
' EntropyNCD
'   Coef           Added to every entropy (default 1).
'   LogBase        Logarithm base (default 2).
'===============================================================================
Public Function EntropyNCD_Distance(ByVal s1 As String, ByVal s2 As String, _
                                    Optional ByVal CaseSensitive As Boolean = True, _
                                    Optional ByVal Qval As Long = 1, _
                                    Optional ByVal Coef As Double = 1, _
                                    Optional ByVal LogBase As Double = 2) As Double
    EntropyNCD_Distance = Ncd(ncdEntropy, s1, s2, CaseSensitive, Qval, LogBase, Coef, "")
End Function

Public Function EntropyNCD_Similarity(ByVal s1 As String, ByVal s2 As String, _
                                      Optional ByVal CaseSensitive As Boolean = True, _
                                      Optional ByVal Qval As Long = 1, _
                                      Optional ByVal Coef As Double = 1, _
                                      Optional ByVal LogBase As Double = 2) As Double
    EntropyNCD_Similarity = 1 - EntropyNCD_Distance(s1, s2, CaseSensitive, Qval, Coef, LogBase)
End Function

Public Function EntropyNCD_NormalizedDistance(ByVal s1 As String, ByVal s2 As String, _
                                              Optional ByVal CaseSensitive As Boolean = True, _
                                              Optional ByVal Qval As Long = 1, _
                                              Optional ByVal Coef As Double = 1, _
                                              Optional ByVal LogBase As Double = 2) As Double
    EntropyNCD_NormalizedDistance = EntropyNCD_Distance(s1, s2, CaseSensitive, Qval, Coef, LogBase)
End Function

Public Function EntropyNCD_NormalizedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                                Optional ByVal CaseSensitive As Boolean = True, _
                                                Optional ByVal Qval As Long = 1, _
                                                Optional ByVal Coef As Double = 1, _
                                                Optional ByVal LogBase As Double = 2) As Double
    EntropyNCD_NormalizedSimilarity = 1 - EntropyNCD_Distance(s1, s2, CaseSensitive, Qval, Coef, LogBase)
End Function

' Always 1.
Public Function EntropyNCD_Maximum(ByVal s1 As String, ByVal s2 As String) As Long
    EntropyNCD_Maximum = 1
End Function

'===============================================================================
' NCD (textdistance's _NCDBase.__call__)
'===============================================================================
Private Function Ncd(ByVal Kind As Long, ByVal s1 As String, ByVal s2 As String, _
                     ByVal CaseSensitive As Boolean, ByVal Qval As Long, ByVal LogBase As Double, _
                     ByVal Coef As Double, ByVal Terminator As String) As Double
    Dim size1 As Double, size2 As Double, concat As Double, other As Double, maxSize As Double
    Dim t1() As String, t2() As String, both() As String, n1 As Long, n2 As Long, i As Long

    If Not CaseSensitive Then
        s1 = LCase$(s1)
        s2 = LCase$(s2)
    End If

    Select Case Kind
        Case ncdSqrt, ncdEntropy
            ' Token counts: the order of concatenation does not matter
            n1 = Tokenize(s1, Qval, t1)
            n2 = Tokenize(s2, Qval, t2)
            ReDim both(1 To IIf(n1 + n2 > 0, n1 + n2, 1))
            For i = 1 To n1
                both(i) = t1(i)
            Next i
            For i = 1 To n2
                both(n1 + i) = t2(i)
            Next i
            size1 = CountSize(Kind, t1, n1, LogBase, Coef)
            size2 = CountSize(Kind, t2, n2, LogBase, Coef)
            concat = CountSize(Kind, both, n1 + n2, LogBase, Coef)
        Case Else
            size1 = TextSize(Kind, s1, LogBase, Terminator)
            size2 = TextSize(Kind, s2, LogBase, Terminator)
            concat = TextSize(Kind, s1 & s2, LogBase, Terminator)
            other = TextSize(Kind, s2 & s1, LogBase, Terminator)
            If other < concat Then concat = other
    End Select

    maxSize = IIf(size1 > size2, size1, size2)
    If maxSize = 0 Then Exit Function
    Ncd = (concat - IIf(size1 < size2, size1, size2)) / maxSize
End Function

' Compressed size of a string for the character-based compressors.
Private Function TextSize(ByVal Kind As Long, ByVal data As String, ByVal LogBase As Double, _
                          ByVal Terminator As String) As Double
    Select Case Kind
        Case ncdArith: TextSize = ArithSize(data, LogBase)
        Case ncdRle: TextSize = RleSize(data)
        Case ncdBwtRle: TextSize = RleSize(Bwt(data, Terminator))
    End Select
End Function

' Compressed size from token counts (tokens need not be sorted on entry).
Private Function CountSize(ByVal Kind As Long, tokens() As String, ByVal n As Long, _
                           ByVal LogBase As Double, ByVal Coef As Double) As Double
    Dim i As Long, runStart As Long, count As Long, p As Double, total As Double

    If n > 0 Then SortStrings tokens, 1, n
    runStart = 1
    For i = 1 To n + 1
        If i > n Then
            count = i - runStart
        ElseIf tokens(i) <> tokens(runStart) Then
            count = i - runStart
        Else
            count = 0
        End If
        If count > 0 Then
            If Kind = ncdSqrt Then
                total = total + Sqr(count)
            Else
                p = count / n
                total = total - p * (Log(p) / Log(LogBase))
            End If
            runStart = i
        End If
    Next i
    If Kind = ncdEntropy Then total = total + Coef
    CountSize = total
End Function

'===============================================================================
' Run-length encoding and Burrows-Wheeler transform
'===============================================================================

' Length of the RLE of data: runs of 3+ become count & char, runs of 1-2 stay.
Private Function RleSize(ByVal data As String) As Long
    Dim i As Long, runStart As Long, n As Long, total As Long
    runStart = 1
    For i = 2 To Len(data) + 1
        If i > Len(data) Then
            n = i - runStart
        ElseIf Mid$(data, i, 1) <> Mid$(data, runStart, 1) Then
            n = i - runStart
        Else
            n = 0
        End If
        If n > 0 Then
            If n > 2 Then
                total = total + Len(CStr(n)) + 1
            Else
                total = total + n
            End If
            runStart = i
        End If
    Next i
    RleSize = total
End Function

' textdistance's BWTRLENCD._compress before the RLE step.
Private Function Bwt(ByVal data As String, ByVal Terminator As String) As String
    Dim rotations() As String, i As Long, n As Long, result As String

    If Len(data) = 0 Then
        Bwt = Terminator
        Exit Function
    End If
    If InStr(1, data, Terminator, vbBinaryCompare) > 0 Then
        Bwt = data
        Exit Function
    End If

    data = data & Terminator
    n = Len(data)
    ReDim rotations(1 To n)
    For i = 1 To n
        rotations(i) = Mid$(data, i) & Left$(data, i - 1)
    Next i
    SortStrings rotations, 1, n
    result = String$(n, " ")
    For i = 1 To n
        Mid$(result, i, 1) = Right$(rotations(i), 1)
    Next i
    Bwt = result
End Function

'===============================================================================
' Arithmetic coding size (textdistance's ArithNCD._get_size), computed with
' integers: after encoding, start = S / D and width = W / D. The output is the
' first k where the smallest multiple of 1/2^k above start is below start +
' width; with r = S * 2^k mod D that is when D - r < W * 2^k.
'===============================================================================
Private Function ArithSize(ByVal data As String, ByVal LogBase As Double) As Long
    Dim n As Long, i As Long, j As Long, m As Long
    Dim chars() As String, counts() As Long, cums() As Long, idx() As Long, order() As Long
    Dim S() As Double, W() As Double, D() As Double, r() As Double, Wk() As Double, tmp() As Double
    Dim bits() As Byte, k As Long, ch As String, swap As Long, acc As Long

    n = Len(data)
    If n = 0 Then Exit Function

    ' Distinct characters in order of first appearance, with counts
    ReDim chars(1 To n)
    ReDim counts(1 To n)
    ReDim idx(1 To n)
    For i = 1 To n
        ch = Mid$(data, i, 1)
        For j = 1 To m
            If chars(j) = ch Then Exit For
        Next j
        If j > m Then
            m = m + 1
            chars(m) = ch
        End If
        counts(j) = counts(j) + 1
        idx(i) = j
    Next i

    ' Cumulative counts in Counter.most_common() order (count desc, stable)
    ReDim order(1 To m)
    For i = 1 To m
        order(i) = i
        For j = i To 2 Step -1
            If counts(order(j)) > counts(order(j - 1)) Then
                swap = order(j): order(j) = order(j - 1): order(j - 1) = swap
            Else
                Exit For
            End If
        Next j
    Next i
    ReDim cums(1 To m)
    For i = 1 To m
        cums(order(i)) = acc
        acc = acc + counts(order(i))
    Next i

    ' Encode: S = S*n + cum*W, W = W*count, D = D*n
    BigSet S, 0
    BigSet W, 1
    BigSet D, 1
    For i = 1 To n
        BigCopy W, tmp
        BigMulSmall tmp, cums(idx(i))
        BigMulSmall S, n
        BigAdd S, tmp
        BigMulSmall W, counts(idx(i))
        BigMulSmall D, n
    Next i
    If BigIsZero(S) Then Exit Function      ' output fraction 0 -> size 0

    ' Find k, collecting the bits of floor(S * 2^k / D)
    BigCopy S, r
    BigCopy W, Wk
    ReDim bits(0 To 0)
    k = 0
    Do
        BigCopy r, tmp
        BigAdd tmp, Wk
        If BigCmp(D, tmp) < 0 Then Exit Do
        BigMulSmall r, 2
        BigMulSmall Wk, 2
        k = k + 1
        ReDim Preserve bits(0 To k)
        If BigCmp(r, D) >= 0 Then
            BigSub r, D
            bits(k) = 1
        End If
    Loop

    ArithSize = CeilLog(OddPartLog(bits, k), LogBase)
End Function

' bits(1..k) are floor(S*2^k/D), most significant first. Returns the natural
' log of the odd part of that number + 1 (textdistance's reduced numerator),
' computed as Python's math.log does for an int: via the correctly rounded
' Double, or as log(mantissa) + exponent * log(2) when it exceeds a Double.
Private Function OddPartLog(bits() As Byte, ByVal k As Long) As Double
    Dim num() As Byte, i As Long, first As Long, last As Long, L As Long
    Dim mantissa As Double, roundUp As Boolean, sticky As Boolean

    ' num = bits + 1, with one extra leading position for a carry
    ReDim num(0 To k)
    For i = 1 To k
        num(i) = bits(i)
    Next i
    i = k
    Do While i >= 1
        If num(i) = 1 Then
            num(i) = 0
            i = i - 1
        Else
            Exit Do
        End If
    Loop
    num(i) = 1                          ' i = 0 when every bit carried

    first = 0
    Do While num(first) = 0
        first = first + 1
    Loop
    last = k
    Do While num(last) = 0              ' drop trailing zeros: the odd part
        last = last - 1
    Loop
    L = last - first + 1                ' bit length

    If L <= 53 Then
        For i = first To last
            mantissa = mantissa * 2 + num(i)
        Next i
        OddPartLog = Log(mantissa)
        Exit Function
    End If

    ' Round to 53 significant bits, ties to even
    For i = first To first + 52
        mantissa = mantissa * 2 + num(i)
    Next i
    For i = first + 54 To last
        If num(i) = 1 Then sticky = True
    Next i
    If num(first + 53) = 1 Then
        If sticky Then
            roundUp = True
        Else
            roundUp = (num(first + 52) = 1)
        End If
    End If
    If roundUp Then mantissa = mantissa + 1

    If mantissa = 2# ^ 53 Then          ' rounding carried into a new bit
        mantissa = 2# ^ 52
        L = L + 1
    End If
    If L <= 1024 Then
        OddPartLog = Log(mantissa * 2# ^ (L - 53))
    Else
        ' Beyond a Double: frexp-style mantissa in [0.5, 1) and exponent L
        OddPartLog = Log(mantissa / 2# ^ 53) + Log(2#) * L
    End If
End Function

' Ceiling(lnX / ln(LogBase)) as Python's math.ceil(math.log(x, base)).
Private Function CeilLog(ByVal lnX As Double, ByVal LogBase As Double) As Long
    CeilLog = -Int(-(lnX / Log(LogBase)))
End Function

'===============================================================================
' Minimal non-negative big integers: arrays of base-10^7 digits, least
' significant first (Double digits keep products exact).
'===============================================================================
Private Sub BigSet(a() As Double, ByVal v As Double)
    ReDim a(0 To 0)
    a(0) = v
End Sub

Private Sub BigCopy(src() As Double, dst() As Double)
    Dim i As Long
    ReDim dst(0 To UBound(src))
    For i = 0 To UBound(src)
        dst(i) = src(i)
    Next i
End Sub

Private Function BigIsZero(a() As Double) As Boolean
    Dim i As Long
    For i = 0 To UBound(a)
        If a(i) <> 0 Then Exit Function
    Next i
    BigIsZero = True
End Function

' a = a * m  (0 <= m < 10^8)
Private Sub BigMulSmall(a() As Double, ByVal m As Double)
    Dim i As Long, carry As Double, x As Double
    For i = 0 To UBound(a)
        x = a(i) * m + carry
        carry = Int(x / LIMB)
        a(i) = x - carry * LIMB
        If a(i) < 0 Then
            a(i) = a(i) + LIMB
            carry = carry - 1
        ElseIf a(i) >= LIMB Then
            a(i) = a(i) - LIMB
            carry = carry + 1
        End If
    Next i
    Do While carry > 0
        ReDim Preserve a(0 To UBound(a) + 1)
        a(UBound(a)) = carry - Int(carry / LIMB) * LIMB
        carry = Int(carry / LIMB)
    Loop
End Sub

' a = a + b
Private Sub BigAdd(a() As Double, b() As Double)
    Dim i As Long, carry As Double, x As Double
    If UBound(b) > UBound(a) Then
        ReDim Preserve a(0 To UBound(b))
    End If
    For i = 0 To UBound(a)
        x = a(i) + carry
        If i <= UBound(b) Then x = x + b(i)
        If x >= LIMB Then
            a(i) = x - LIMB
            carry = 1
        Else
            a(i) = x
            carry = 0
        End If
    Next i
    If carry > 0 Then
        ReDim Preserve a(0 To UBound(a) + 1)
        a(UBound(a)) = carry
    End If
End Sub

' a = a - b  (requires a >= b)
Private Sub BigSub(a() As Double, b() As Double)
    Dim i As Long, borrow As Double, x As Double
    For i = 0 To UBound(a)
        x = a(i) - borrow
        If i <= UBound(b) Then x = x - b(i)
        If x < 0 Then
            a(i) = x + LIMB
            borrow = 1
        Else
            a(i) = x
            borrow = 0
        End If
    Next i
End Sub

' -1, 0 or 1 as a <, = or > b.
Private Function BigCmp(a() As Double, b() As Double) As Long
    Dim i As Long, x As Double, y As Double, top As Long
    top = IIf(UBound(a) > UBound(b), UBound(a), UBound(b))
    For i = top To 0 Step -1
        x = 0
        y = 0
        If i <= UBound(a) Then x = a(i)
        If i <= UBound(b) Then y = b(i)
        If x <> y Then
            BigCmp = IIf(x < y, -1, 1)
            Exit Function
        End If
    Next i
End Function

'===============================================================================
' Tokens and sorting
'===============================================================================

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
