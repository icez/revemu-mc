Attribute VB_Name = "md_Convert"
Option Explicit
Type typeConv_Walk
    From As Coord
    To As Coord
End Type

'text
Function Conv2Arr(inString As String) As Byte()
    Dim strs() As Byte, i&
    ReDim strs(Len(inString) - 1)
    For i = 1 To Len(inString)
        strs(i - 1) = Asc(Mid(inString, i, 1))
    Next
    Conv2Arr = strs
End Function
Function Conv2Str(inArr() As Byte) As String
    Dim i&, res$
    For i = LBound(inArr) To UBound(inArr)
        res = res & Chr(inArr(i))
    Next
    Conv2Str = res
End Function
Function LngToChr(inLong As Long) As String
    Dim b(3) As Byte
    CopyMemory b(0), inLong, 4
    LngToChr = Chr(b(0)) & Chr(b(1)) & Chr(b(2)) & Chr(b(3))
End Function
Function IntToChr(inLong As Long) As String
    Dim c1 As Long, c2 As Long
    c1 = inLong Mod 256
    c2 = Int(inLong / 256) Mod 256
    IntToChr = Chr(c1) & Chr(c2)
End Function
Function ReverseHex(inHex As String) As String
    If (Len(inHex) Mod 2) <> 0 Then Exit Function
    Dim i As Long, tChr As String
    For i = 1 To Len(inHex) Step 2
        tChr = tChr & Mid(inHex, Len(inHex) - i, 2)
    Next
    ReverseHex = tChr
End Function
Function ChrtoHex(inString As String) As String ' "AB" > 4142
    Dim tstr As String, i As Long
    For i = 1 To Len(inString)
        If Len(Hex(Asc(Mid(inString, i, 1)))) = 1 Then tstr = tstr & "0" & Hex(Asc(Mid(inString, i, 1))) Else tstr = tstr & Hex(Asc(Mid(inString, i, 1)))
    Next
    ChrtoHex = tstr
End Function

'coord
Function Str3XY(inString As String) As Coord
    Str3XY.X = (Asc(Mid(inString, 1, 1)) * 4) + Int((Asc(Mid(inString, 2, 1)) And &HC0) / 64)
    Str3XY.Y = ((Asc(Mid(inString, 2, 1)) And &H3F) * 16) + Int((Asc(Mid(inString, 3, 1)) And &HF0) / 16)
End Function
Function Str5XY(inString As String) As typeConv_Walk
    Str5XY.From = Str3XY(Mid(inString, 1, 2) + Chr(Asc(Mid(inString, 3, 1)) And &HF0))
    Str5XY.To = Str3XY(Chr(((Asc(Mid(inString, 3, 1)) And &HF) * 16) + Int(Asc(Mid(inString, 4, 1)) / 16)) & Chr(((Asc(Mid(inString, 4, 1)) And &HF) * 16) + Int(Asc(Mid(inString, 5, 1)) / 16)) & Chr((Asc(Mid(inString, 5, 1)) And &HF) * 16))
End Function
