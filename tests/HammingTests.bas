Attribute VB_Name = "modHammingTests"
' Expected values generated from Python textdistance 4.6.2.
' Run Test_Hamming and check the Immediate window (Ctrl+G).
Option Explicit
Option Compare Binary

Private mPass As Long, mFail As Long

Public Sub Test_Hamming()
    mPass = 0: mFail = 0
    CheckL "Distance(""test"", ""text"", True, False)", Hamming_Distance("test", "text", True, False), 1
    CheckL "Similarity(""test"", ""text"", True, False)", Hamming_Similarity("test", "text", True, False), 3
    CheckD "NormDist(""test"", ""text"", True, False)", Hamming_NormalizedDistance("test", "text", True, False), 0.25
    CheckD "NormSim(""test"", ""text"", True, False)", Hamming_NormalizedSimilarity("test", "text", True, False), 0.75
    CheckL "Distance(""test"", ""testing"", True, False)", Hamming_Distance("test", "testing", True, False), 3
    CheckL "Similarity(""test"", ""testing"", True, False)", Hamming_Similarity("test", "testing", True, False), 4
    CheckD "NormDist(""test"", ""testing"", True, False)", Hamming_NormalizedDistance("test", "testing", True, False), 0.42857142857142855
    CheckD "NormSim(""test"", ""testing"", True, False)", Hamming_NormalizedSimilarity("test", "testing", True, False), 0.5714285714285714
    CheckL "Distance(""testing"", ""test"", True, False)", Hamming_Distance("testing", "test", True, False), 3
    CheckL "Similarity(""testing"", ""test"", True, False)", Hamming_Similarity("testing", "test", True, False), 4
    CheckD "NormDist(""testing"", ""test"", True, False)", Hamming_NormalizedDistance("testing", "test", True, False), 0.42857142857142855
    CheckD "NormSim(""testing"", ""test"", True, False)", Hamming_NormalizedSimilarity("testing", "test", True, False), 0.5714285714285714
    CheckL "Distance("""", """", True, False)", Hamming_Distance("", "", True, False), 0
    CheckL "Similarity("""", """", True, False)", Hamming_Similarity("", "", True, False), 0
    CheckD "NormDist("""", """", True, False)", Hamming_NormalizedDistance("", "", True, False), 0
    CheckD "NormSim("""", """", True, False)", Hamming_NormalizedSimilarity("", "", True, False), 1
    CheckL "Distance("""", ""abc"", True, False)", Hamming_Distance("", "abc", True, False), 3
    CheckL "Similarity("""", ""abc"", True, False)", Hamming_Similarity("", "abc", True, False), 0
    CheckD "NormDist("""", ""abc"", True, False)", Hamming_NormalizedDistance("", "abc", True, False), 1.0
    CheckD "NormSim("""", ""abc"", True, False)", Hamming_NormalizedSimilarity("", "abc", True, False), 0.0
    CheckL "Distance(""abc"", """", True, False)", Hamming_Distance("abc", "", True, False), 3
    CheckL "Similarity(""abc"", """", True, False)", Hamming_Similarity("abc", "", True, False), 0
    CheckD "NormDist(""abc"", """", True, False)", Hamming_NormalizedDistance("abc", "", True, False), 1.0
    CheckD "NormSim(""abc"", """", True, False)", Hamming_NormalizedSimilarity("abc", "", True, False), 0.0
    CheckL "Distance(""karolin"", ""kathrin"", True, False)", Hamming_Distance("karolin", "kathrin", True, False), 3
    CheckL "Similarity(""karolin"", ""kathrin"", True, False)", Hamming_Similarity("karolin", "kathrin", True, False), 4
    CheckD "NormDist(""karolin"", ""kathrin"", True, False)", Hamming_NormalizedDistance("karolin", "kathrin", True, False), 0.42857142857142855
    CheckD "NormSim(""karolin"", ""kathrin"", True, False)", Hamming_NormalizedSimilarity("karolin", "kathrin", True, False), 0.5714285714285714
    CheckL "Distance(""1011101"", ""1001001"", True, False)", Hamming_Distance("1011101", "1001001", True, False), 2
    CheckL "Similarity(""1011101"", ""1001001"", True, False)", Hamming_Similarity("1011101", "1001001", True, False), 5
    CheckD "NormDist(""1011101"", ""1001001"", True, False)", Hamming_NormalizedDistance("1011101", "1001001", True, False), 0.2857142857142857
    CheckD "NormSim(""1011101"", ""1001001"", True, False)", Hamming_NormalizedSimilarity("1011101", "1001001", True, False), 0.7142857142857143
    CheckL "Distance(""abcd"", ""abxyz"", True, False)", Hamming_Distance("abcd", "abxyz", True, False), 3
    CheckL "Similarity(""abcd"", ""abxyz"", True, False)", Hamming_Similarity("abcd", "abxyz", True, False), 2
    CheckD "NormDist(""abcd"", ""abxyz"", True, False)", Hamming_NormalizedDistance("abcd", "abxyz", True, False), 0.6
    CheckD "NormSim(""abcd"", ""abxyz"", True, False)", Hamming_NormalizedSimilarity("abcd", "abxyz", True, False), 0.4
    CheckL "Distance(""same"", ""same"", True, False)", Hamming_Distance("same", "same", True, False), 0
    CheckL "Similarity(""same"", ""same"", True, False)", Hamming_Similarity("same", "same", True, False), 4
    CheckD "NormDist(""same"", ""same"", True, False)", Hamming_NormalizedDistance("same", "same", True, False), 0.0
    CheckD "NormSim(""same"", ""same"", True, False)", Hamming_NormalizedSimilarity("same", "same", True, False), 1.0
    CheckL "Distance(""Hello"", ""hello"", True, False)", Hamming_Distance("Hello", "hello", True, False), 1
    CheckL "Similarity(""Hello"", ""hello"", True, False)", Hamming_Similarity("Hello", "hello", True, False), 4
    CheckD "NormDist(""Hello"", ""hello"", True, False)", Hamming_NormalizedDistance("Hello", "hello", True, False), 0.2
    CheckD "NormSim(""Hello"", ""hello"", True, False)", Hamming_NormalizedSimilarity("Hello", "hello", True, False), 0.8
    CheckL "Distance(""ABC"", ""xyz"", True, False)", Hamming_Distance("ABC", "xyz", True, False), 3
    CheckL "Similarity(""ABC"", ""xyz"", True, False)", Hamming_Similarity("ABC", "xyz", True, False), 0
    CheckD "NormDist(""ABC"", ""xyz"", True, False)", Hamming_NormalizedDistance("ABC", "xyz", True, False), 1.0
    CheckD "NormSim(""ABC"", ""xyz"", True, False)", Hamming_NormalizedSimilarity("ABC", "xyz", True, False), 0.0
    CheckL "Distance(""a"", ""b"", True, False)", Hamming_Distance("a", "b", True, False), 1
    CheckL "Similarity(""a"", ""b"", True, False)", Hamming_Similarity("a", "b", True, False), 0
    CheckD "NormDist(""a"", ""b"", True, False)", Hamming_NormalizedDistance("a", "b", True, False), 1.0
    CheckD "NormSim(""a"", ""b"", True, False)", Hamming_NormalizedSimilarity("a", "b", True, False), 0.0
    CheckL "Distance(""test"", ""text"", True, True)", Hamming_Distance("test", "text", True, True), 1
    CheckL "Similarity(""test"", ""text"", True, True)", Hamming_Similarity("test", "text", True, True), 3
    CheckD "NormDist(""test"", ""text"", True, True)", Hamming_NormalizedDistance("test", "text", True, True), 0.25
    CheckD "NormSim(""test"", ""text"", True, True)", Hamming_NormalizedSimilarity("test", "text", True, True), 0.75
    CheckL "Distance(""test"", ""testing"", True, True)", Hamming_Distance("test", "testing", True, True), 0
    CheckL "Similarity(""test"", ""testing"", True, True)", Hamming_Similarity("test", "testing", True, True), 7
    CheckD "NormDist(""test"", ""testing"", True, True)", Hamming_NormalizedDistance("test", "testing", True, True), 0.0
    CheckD "NormSim(""test"", ""testing"", True, True)", Hamming_NormalizedSimilarity("test", "testing", True, True), 1.0
    CheckL "Distance(""testing"", ""test"", True, True)", Hamming_Distance("testing", "test", True, True), 0
    CheckL "Similarity(""testing"", ""test"", True, True)", Hamming_Similarity("testing", "test", True, True), 7
    CheckD "NormDist(""testing"", ""test"", True, True)", Hamming_NormalizedDistance("testing", "test", True, True), 0.0
    CheckD "NormSim(""testing"", ""test"", True, True)", Hamming_NormalizedSimilarity("testing", "test", True, True), 1.0
    CheckL "Distance("""", """", True, True)", Hamming_Distance("", "", True, True), 0
    CheckL "Similarity("""", """", True, True)", Hamming_Similarity("", "", True, True), 0
    CheckD "NormDist("""", """", True, True)", Hamming_NormalizedDistance("", "", True, True), 0
    CheckD "NormSim("""", """", True, True)", Hamming_NormalizedSimilarity("", "", True, True), 1
    CheckL "Distance("""", ""abc"", True, True)", Hamming_Distance("", "abc", True, True), 3
    CheckL "Similarity("""", ""abc"", True, True)", Hamming_Similarity("", "abc", True, True), 0
    CheckD "NormDist("""", ""abc"", True, True)", Hamming_NormalizedDistance("", "abc", True, True), 1.0
    CheckD "NormSim("""", ""abc"", True, True)", Hamming_NormalizedSimilarity("", "abc", True, True), 0.0
    CheckL "Distance(""abc"", """", True, True)", Hamming_Distance("abc", "", True, True), 3
    CheckL "Similarity(""abc"", """", True, True)", Hamming_Similarity("abc", "", True, True), 0
    CheckD "NormDist(""abc"", """", True, True)", Hamming_NormalizedDistance("abc", "", True, True), 1.0
    CheckD "NormSim(""abc"", """", True, True)", Hamming_NormalizedSimilarity("abc", "", True, True), 0.0
    CheckL "Distance(""karolin"", ""kathrin"", True, True)", Hamming_Distance("karolin", "kathrin", True, True), 3
    CheckL "Similarity(""karolin"", ""kathrin"", True, True)", Hamming_Similarity("karolin", "kathrin", True, True), 4
    CheckD "NormDist(""karolin"", ""kathrin"", True, True)", Hamming_NormalizedDistance("karolin", "kathrin", True, True), 0.42857142857142855
    CheckD "NormSim(""karolin"", ""kathrin"", True, True)", Hamming_NormalizedSimilarity("karolin", "kathrin", True, True), 0.5714285714285714
    CheckL "Distance(""1011101"", ""1001001"", True, True)", Hamming_Distance("1011101", "1001001", True, True), 2
    CheckL "Similarity(""1011101"", ""1001001"", True, True)", Hamming_Similarity("1011101", "1001001", True, True), 5
    CheckD "NormDist(""1011101"", ""1001001"", True, True)", Hamming_NormalizedDistance("1011101", "1001001", True, True), 0.2857142857142857
    CheckD "NormSim(""1011101"", ""1001001"", True, True)", Hamming_NormalizedSimilarity("1011101", "1001001", True, True), 0.7142857142857143
    CheckL "Distance(""abcd"", ""abxyz"", True, True)", Hamming_Distance("abcd", "abxyz", True, True), 2
    CheckL "Similarity(""abcd"", ""abxyz"", True, True)", Hamming_Similarity("abcd", "abxyz", True, True), 3
    CheckD "NormDist(""abcd"", ""abxyz"", True, True)", Hamming_NormalizedDistance("abcd", "abxyz", True, True), 0.4
    CheckD "NormSim(""abcd"", ""abxyz"", True, True)", Hamming_NormalizedSimilarity("abcd", "abxyz", True, True), 0.6
    CheckL "Distance(""same"", ""same"", True, True)", Hamming_Distance("same", "same", True, True), 0
    CheckL "Similarity(""same"", ""same"", True, True)", Hamming_Similarity("same", "same", True, True), 4
    CheckD "NormDist(""same"", ""same"", True, True)", Hamming_NormalizedDistance("same", "same", True, True), 0.0
    CheckD "NormSim(""same"", ""same"", True, True)", Hamming_NormalizedSimilarity("same", "same", True, True), 1.0
    CheckL "Distance(""Hello"", ""hello"", True, True)", Hamming_Distance("Hello", "hello", True, True), 1
    CheckL "Similarity(""Hello"", ""hello"", True, True)", Hamming_Similarity("Hello", "hello", True, True), 4
    CheckD "NormDist(""Hello"", ""hello"", True, True)", Hamming_NormalizedDistance("Hello", "hello", True, True), 0.2
    CheckD "NormSim(""Hello"", ""hello"", True, True)", Hamming_NormalizedSimilarity("Hello", "hello", True, True), 0.8
    CheckL "Distance(""ABC"", ""xyz"", True, True)", Hamming_Distance("ABC", "xyz", True, True), 3
    CheckL "Similarity(""ABC"", ""xyz"", True, True)", Hamming_Similarity("ABC", "xyz", True, True), 0
    CheckD "NormDist(""ABC"", ""xyz"", True, True)", Hamming_NormalizedDistance("ABC", "xyz", True, True), 1.0
    CheckD "NormSim(""ABC"", ""xyz"", True, True)", Hamming_NormalizedSimilarity("ABC", "xyz", True, True), 0.0
    CheckL "Distance(""a"", ""b"", True, True)", Hamming_Distance("a", "b", True, True), 1
    CheckL "Similarity(""a"", ""b"", True, True)", Hamming_Similarity("a", "b", True, True), 0
    CheckD "NormDist(""a"", ""b"", True, True)", Hamming_NormalizedDistance("a", "b", True, True), 1.0
    CheckD "NormSim(""a"", ""b"", True, True)", Hamming_NormalizedSimilarity("a", "b", True, True), 0.0
    CheckL "Distance(""Hello"", ""hello"", False)", Hamming_Distance("Hello", "hello", False), 0
    CheckL "Distance(""TeSt"", ""tEXT"", False)", Hamming_Distance("TeSt", "tEXT", False), 1
    CheckL "Distance(""ABC"", ""abcd"", False)", Hamming_Distance("ABC", "abcd", False), 1
    CheckL "Maximum(""test"", ""testing"")", Hamming_Maximum("test", "testing"), 7
    Debug.Print "Hamming: " & mPass & " passed, " & mFail & " failed"
End Sub

Private Sub CheckL(ByVal Name As String, ByVal Actual As Long, ByVal Expected As Long)
    If Actual = Expected Then
        mPass = mPass + 1
    Else
        mFail = mFail + 1
        Debug.Print "FAIL " & Name & ": got " & Actual & ", expected " & Expected
    End If
End Sub

Private Sub CheckD(ByVal Name As String, ByVal Actual As Double, ByVal Expected As Double)
    If Abs(Actual - Expected) < 0.000000001 Then
        mPass = mPass + 1
    Else
        mFail = mFail + 1
        Debug.Print "FAIL " & Name & ": got " & Actual & ", expected " & Expected
    End If
End Sub
