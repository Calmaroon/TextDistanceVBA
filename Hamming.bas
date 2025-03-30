Attribute VB_Name = "Hamming"
Option Explicit
Function Hamming(strInput1 As String, strInput2 As String, Optional boolCaseSensitive = True) As Long
    Hamming = distance(strInput1, strInput2, boolCaseSensitive)
End Function
Function distance(strInput1 As String, strInput2 As String, Optional boolCaseSensitive = True) As Long
    Dim i As Long
    Dim lngDistance As Long
    
    If Not boolCaseSensitive Then
        strInput1 = UCase$(strInput1)
        strInput2 = UCase$(strInput2)
    End If
    
    If Len(strInput1) <> Len(strInput2) Then
        distance = -1
        Exit Function
    End If
    
    For i = 1 To Len(strInput1)
        If Mid$(strInput1, i, 1) <> Mid$(strInput2, i, 1) Then
            lngDistance = lngDistance + 1
        End If
    Next i
    
    distance = lngDistance
End Function
Function similarity(strInput1 As String, strInput2 As String, Optional boolCaseSensitive = True) As Long
    Dim i As Long
    Dim lngSimiliarity As Long
    
    If Not boolCaseSensitive Then
        strInput1 = UCase$(strInput1)
        strInput2 = UCase$(strInput2)
    End If
    
    If Len(strInput1) <> Len(strInput2) Then
        similarity = -1
        Exit Function
    End If
    
    For i = 1 To Len(strInput1)
        If Mid$(strInput1, i, 1) = Mid$(strInput2, i, 1) Then
            lngSimiliarity = lngSimiliarity + 1
        End If
    Next i
    
    similarity = lngSimiliarity
End Function
Function normalized_distance(strInput1 As String, strInput2 As String, Optional boolCaseSensitive = True) As Double
    If Len(strInput1) <> Len(strInput2) Then
        normalized_distance = -1
        Exit Function
    End If
    
    normalized_distance = distance(strInput1, strInput2, boolCaseSensitive) / Len(strInput1)
End Function
Function normalized_similarity(strInput1 As String, strInput2 As String, Optional boolCaseSensitive = True) As Double
    If Len(strInput1) <> Len(strInput2) Then
        normalized_similarity = -1
        Exit Function
    End If
    
    normalized_similarity = similarity(strInput1, strInput2, boolCaseSensitive) / Len(strInput1)
End Function

