Attribute VB_Name = "Excel_UDF"
Option Explicit
Function HAMMING(strInput1 As String, strInput2 As String, Optional boolCaseSensitive = True)
    HAMMING = HAMMING.distance(strInput1, strInput2, boolCaseSensitive)
End Function
Function HAMMING_DISTANCE(strInput1 As String, strInput2 As String, Optional boolCaseSensitive = True) As Long
    HAMMING_DISTANCE = HAMMING.distance(strInput1, strInput2, boolCaseSensitive)
End Function
Function HAMMING_SIMILARITY(strInput1 As String, strInput2 As String, Optional boolCaseSensitive = True) As Long
    HAMMING_DISTANCE = HAMMING.similarity(strInput1, strInput2, boolCaseSensitive)
End Function
