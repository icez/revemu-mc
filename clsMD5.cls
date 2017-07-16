VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MD5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=
'= Class Constants
'=
Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647

Private Const S11 = 7
Private Const S12 = 12
Private Const S13 = 17
Private Const S14 = 22
Private Const S21 = 5
Private Const S22 = 9
Private Const S23 = 14
Private Const S24 = 20
Private Const S31 = 4
Private Const S32 = 11
Private Const S33 = 16
Private Const S34 = 23
Private Const S41 = 6
Private Const S42 = 10
Private Const S43 = 15
Private Const S44 = 21


'=
'= Class Variables
'=
Private State(4) As Long
Private ByteCounter As Long
Private ByteBuffer(63) As Byte


'=
'= Class Properties
'=
Property Get RegisterA() As String
    RegisterA = State(1)
End Property

Property Get RegisterB() As String
    RegisterB = State(2)
End Property

Property Get RegisterC() As String
    RegisterC = State(3)
End Property

Property Get RegisterD() As String
    RegisterD = State(4)
End Property


'=
'= Class Functions
'=

'
' Function to quickly digest a file into a hex string
'
Public Function DigestFileToHexStr(FileName As String) As String
    Open FileName For Binary Access Read As #1
    MD5Init
    Do While Not EOF(1)
        Get #1, , ByteBuffer
        If Loc(1) < LOF(1) Then
            ByteCounter = ByteCounter + 64
            MD5Transform ByteBuffer
        End If
    Loop
    ByteCounter = ByteCounter + (LOF(1) Mod 64)
    Close #1
    MD5Final
    DigestFileToHexStr = GetValues
End Function

'
' Function to digest a text string and output the result as a string
' of hexadecimal characters.
'
Public Function DigestStrToHexStr(SourceString As String) As String
    MD5Init
    MD5Update Len(SourceString), StringToArray(SourceString)
    MD5Final
    DigestStrToHexStr = GetValues
End Function

Public Function DigestStrToChar(SourceString As String) As String
    MD5Init
    MD5Update Len(SourceString), StringToArray(SourceString)
    MD5Final
    DigestStrToChar = GetRawValues
End Function
'
' A utility function which converts a string into an array of
' bytes.
'
Private Function StringToArray(inString As String) As Byte()
    Dim i As Integer
    Dim bytBuffer() As Byte
    ReDim bytBuffer(Len(inString))
    For i = 0 To Len(inString) - 1
        bytBuffer(i) = Asc(Mid(inString, i + 1, 1))
    Next i
    StringToArray = bytBuffer
End Function


Public Function GetRawValues() As String
    GetRawValues = LongToChar(State(1)) & LongToChar(State(2)) & LongToChar(State(3)) & LongToChar(State(4))
End Function

Private Function LongToChar(num As Long) As String
        Dim a As Byte
        Dim b As Byte
        Dim c As Byte
        Dim d As Byte
        
        a = num And &HFF&
        b = (num And &HFF00&) \ 256
        c = (num And &HFF0000) \ 65536
        If num < 0 Then
            d = ((num And &H7F000000) \ 16777216) Or &H80&
        Else
            d = (num And &HFF000000) \ 16777216
        End If
        LongToChar = Chr(a) & Chr(b) & Chr(c) & Chr(d)
End Function


'
' Concatenate the four state vaules into one string
'
Public Function GetValues() As String
    GetValues = LongToString(State(1)) & LongToString(State(2)) & LongToString(State(3)) & LongToString(State(4))
End Function

'
' Convert a Long to a Hex string
'
Private Function LongToString(num As Long) As String
        Dim a As Byte
        Dim b As Byte
        Dim c As Byte
        Dim d As Byte
        
        a = num And &HFF&
        If a < 16 Then
            LongToString = "0" & Hex(a)
        Else
            LongToString = Hex(a)
        End If
               
        b = (num And &HFF00&) \ 256
        If b < 16 Then
            LongToString = LongToString & "0" & Hex(b)
        Else
            LongToString = LongToString & Hex(b)
        End If
        
        c = (num And &HFF0000) \ 65536
        If c < 16 Then
            LongToString = LongToString & "0" & Hex(c)
        Else
            LongToString = LongToString & Hex(c)
        End If
       
        If num < 0 Then
            d = ((num And &H7F000000) \ 16777216) Or &H80&
        Else
            d = (num And &HFF000000) \ 16777216
        End If
        
        If d < 16 Then
            LongToString = LongToString & "0" & Hex(d)
        Else
            LongToString = LongToString & Hex(d)
        End If
    
End Function

'
' Initialize the class
'   This must be called before a digest calculation is started
'
Public Sub MD5Init()
    ByteCounter = 0
    State(1) = UnsignedToLong(1732584193#)
    State(2) = UnsignedToLong(4023233417#)
    State(3) = UnsignedToLong(2562383102#)
    State(4) = UnsignedToLong(271733878#)
End Sub

'
' MD5 Final
'
Public Sub MD5Final()
    Dim dblBits As Double
    
    Dim padding(72) As Byte
    Dim lngBytesBuffered As Long
    
    padding(0) = &H80
    
    dblBits = ByteCounter * 8
    
    ' Pad out
    lngBytesBuffered = ByteCounter Mod 64
    If lngBytesBuffered <= 56 Then
        MD5Update 56 - lngBytesBuffered, padding
    Else
        MD5Update 120 - ByteCounter, padding
    End If
    
    
    padding(0) = UnsignedToLong(dblBits) And &HFF&
    padding(1) = UnsignedToLong(dblBits) \ 256 And &HFF&
    'padding(2) = 0
    'padding(3) = 0
    padding(2) = UnsignedToLong(dblBits) \ 65536 And &HFF&
    padding(3) = UnsignedToLong(dblBits) \ 16777216 And &HFF&
    padding(4) = 0
    padding(5) = 0
    padding(6) = 0
    padding(7) = 0
    
    MD5Update 8, padding
End Sub

'
' Break up input stream into 64 byte chunks
'
Public Sub MD5Update(InputLen As Long, InputBuffer() As Byte)
    Dim II As Integer
    Dim i As Integer
    Dim j As Integer
    Dim K As Integer
    Dim lngBufferedBytes As Long
    Dim lngBufferRemaining As Long
    Dim lngRem As Long
    
    lngBufferedBytes = ByteCounter Mod 64
    lngBufferRemaining = 64 - lngBufferedBytes
    ByteCounter = ByteCounter + InputLen
    ' Use up old buffer results first
    If InputLen >= lngBufferRemaining Then
        For II = 0 To lngBufferRemaining - 1
            ByteBuffer(lngBufferedBytes + II) = InputBuffer(II)
        Next II
        MD5Transform ByteBuffer
        
        lngRem = (InputLen) Mod 64
        ' The transfer is a multiple of 64 lets do some transformations
        For i = lngBufferRemaining To InputLen - II - lngRem Step 64
            For j = 0 To 63
                ByteBuffer(j) = InputBuffer(i + j)
            Next j
            MD5Transform ByteBuffer
        Next i
        lngBufferedBytes = 0
    Else
      i = 0
    End If
    
    ' Buffer any remaining input
    For K = 0 To InputLen - i - 1
        ByteBuffer(lngBufferedBytes + K) = InputBuffer(i + K)
    Next K
    
End Sub

'
' MD5 Transform
'
Private Sub MD5Transform(Buffer() As Byte)
    Dim X(16) As Long
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    
    a = State(1)
    b = State(2)
    c = State(3)
    d = State(4)
    
    Decode 64, X, Buffer

    ' Round 1
    FF a, b, c, d, X(0), S11, -680876936
    FF d, a, b, c, X(1), S12, -389564586
    FF c, d, a, b, X(2), S13, 606105819
    FF b, c, d, a, X(3), S14, -1044525330
    FF a, b, c, d, X(4), S11, -176418897
    FF d, a, b, c, X(5), S12, 1200080426
    FF c, d, a, b, X(6), S13, -1473231341
    FF b, c, d, a, X(7), S14, -45705983
    FF a, b, c, d, X(8), S11, 1770035416
    FF d, a, b, c, X(9), S12, -1958414417
    FF c, d, a, b, X(10), S13, -42063
    FF b, c, d, a, X(11), S14, -1990404162
    FF a, b, c, d, X(12), S11, 1804603682
    FF d, a, b, c, X(13), S12, -40341101
    FF c, d, a, b, X(14), S13, -1502002290
    FF b, c, d, a, X(15), S14, 1236535329
    
    ' Round 2
    GG a, b, c, d, X(1), S21, -165796510
    GG d, a, b, c, X(6), S22, -1069501632
    GG c, d, a, b, X(11), S23, 643717713
    GG b, c, d, a, X(0), S24, -373897302
    GG a, b, c, d, X(5), S21, -701558691
    GG d, a, b, c, X(10), S22, 38016083
    GG c, d, a, b, X(15), S23, -660478335
    GG b, c, d, a, X(4), S24, -405537848
    GG a, b, c, d, X(9), S21, 568446438
    GG d, a, b, c, X(14), S22, -1019803690
    GG c, d, a, b, X(3), S23, -187363961
    GG b, c, d, a, X(8), S24, 1163531501
    GG a, b, c, d, X(13), S21, -1444681467
    GG d, a, b, c, X(2), S22, -51403784
    GG c, d, a, b, X(7), S23, 1735328473
    GG b, c, d, a, X(12), S24, -1926607734
    
    ' Round 3
    HH a, b, c, d, X(5), S31, -378558
    HH d, a, b, c, X(8), S32, -2022574463
    HH c, d, a, b, X(11), S33, 1839030562
    HH b, c, d, a, X(14), S34, -35309556
    HH a, b, c, d, X(1), S31, -1530992060
    HH d, a, b, c, X(4), S32, 1272893353
    HH c, d, a, b, X(7), S33, -155497632
    HH b, c, d, a, X(10), S34, -1094730640
    HH a, b, c, d, X(13), S31, 681279174
    HH d, a, b, c, X(0), S32, -358537222
    HH c, d, a, b, X(3), S33, -722521979
    HH b, c, d, a, X(6), S34, 76029189
    HH a, b, c, d, X(9), S31, -640364487
    HH d, a, b, c, X(12), S32, -421815835
    HH c, d, a, b, X(15), S33, 530742520
    HH b, c, d, a, X(2), S34, -995338651
    
    ' Round 4
    II a, b, c, d, X(0), S41, -198630844
    II d, a, b, c, X(7), S42, 1126891415
    II c, d, a, b, X(14), S43, -1416354905
    II b, c, d, a, X(5), S44, -57434055
    II a, b, c, d, X(12), S41, 1700485571
    II d, a, b, c, X(3), S42, -1894986606
    II c, d, a, b, X(10), S43, -1051523
    II b, c, d, a, X(1), S44, -2054922799
    II a, b, c, d, X(8), S41, 1873313359
    II d, a, b, c, X(15), S42, -30611744
    II c, d, a, b, X(6), S43, -1560198380
    II b, c, d, a, X(13), S44, 1309151649
    II a, b, c, d, X(4), S41, -145523070
    II d, a, b, c, X(11), S42, -1120210379
    II c, d, a, b, X(2), S43, 718787259
    II b, c, d, a, X(9), S44, -343485551
    
    
    State(1) = LongOverflowAdd(State(1), a)
    State(2) = LongOverflowAdd(State(2), b)
    State(3) = LongOverflowAdd(State(3), c)
    State(4) = LongOverflowAdd(State(4), d)

'  /* Zeroize sensitive information.
'*/
'  MD5_memset ((POINTER)x, 0, sizeof (x));
    
End Sub

Private Sub Decode(Length As Integer, OutputBuffer() As Long, InputBuffer() As Byte)
    Dim intDblIndex As Integer
    Dim intByteIndex As Integer
    Dim dblSum As Double
    
    intDblIndex = 0
    For intByteIndex = 0 To Length - 1 Step 4
        dblSum = InputBuffer(intByteIndex) + _
                                    InputBuffer(intByteIndex + 1) * 256# + _
                                    InputBuffer(intByteIndex + 2) * 65536# + _
                                    InputBuffer(intByteIndex + 3) * 16777216#
        OutputBuffer(intDblIndex) = UnsignedToLong(dblSum)
        intDblIndex = intDblIndex + 1
    Next intByteIndex
End Sub

'
' FF, GG, HH, and II transformations for rounds 1, 2, 3, and 4.
' Rotation is separate from addition to prevent recomputation.
'
Private Function FF(a As Long, _
                    b As Long, _
                    c As Long, _
                    d As Long, _
                    X As Long, _
                    s As Long, _
                    ac As Long) As Long
    a = LongOverflowAdd4(a, (b And c) Or (Not (b) And d), X, ac)
    a = LongLeftRotate(a, s)
    a = LongOverflowAdd(a, b)
End Function

Private Function GG(a As Long, _
                    b As Long, _
                    c As Long, _
                    d As Long, _
                    X As Long, _
                    s As Long, _
                    ac As Long) As Long
    a = LongOverflowAdd4(a, (b And d) Or (c And Not (d)), X, ac)
    a = LongLeftRotate(a, s)
    a = LongOverflowAdd(a, b)
End Function

Private Function HH(a As Long, _
                    b As Long, _
                    c As Long, _
                    d As Long, _
                    X As Long, _
                    s As Long, _
                    ac As Long) As Long
    a = LongOverflowAdd4(a, b Xor c Xor d, X, ac)
    a = LongLeftRotate(a, s)
    a = LongOverflowAdd(a, b)
End Function

Private Function II(a As Long, _
                    b As Long, _
                    c As Long, _
                    d As Long, _
                    X As Long, _
                    s As Long, _
                    ac As Long) As Long
    a = LongOverflowAdd4(a, c Xor (b Or Not (d)), X, ac)
    a = LongLeftRotate(a, s)
    a = LongOverflowAdd(a, b)
End Function

'
' Rotate a long to the right
'
Function LongLeftRotate(value As Long, bits As Long) As Long
    Dim lngSign As Long
    Dim lngI As Long
    bits = bits Mod 32
    If bits = 0 Then LongLeftRotate = value: Exit Function
    For lngI = 1 To bits
        lngSign = value And &HC0000000
        value = (value And &H3FFFFFFF) * 2
        value = value Or ((lngSign < 0) And 1) Or (CBool(lngSign And _
                &H40000000) And &H80000000)
    Next
    LongLeftRotate = value
End Function

'
' Function to add two unsigned numbers together as in C.
' Overflows are ignored!
'
Private Function LongOverflowAdd(Val1 As Long, Val2 As Long) As Long
    Dim lngHighWord As Long
    Dim lngLowWord As Long
    Dim lngOverflow As Long

    lngLowWord = (Val1 And &HFFFF&) + (Val2 And &HFFFF&)
    lngOverflow = lngLowWord \ 65536
    lngHighWord = (((Val1 And &HFFFF0000) \ 65536) + ((Val2 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&
    LongOverflowAdd = UnsignedToLong((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))
End Function

'
' Function to add two unsigned numbers together as in C.
' Overflows are ignored!
'
Private Function LongOverflowAdd4(Val1 As Long, Val2 As Long, val3 As Long, val4 As Long) As Long
    Dim lngHighWord As Long
    Dim lngLowWord As Long
    Dim lngOverflow As Long

    lngLowWord = (Val1 And &HFFFF&) + (Val2 And &HFFFF&) + (val3 And &HFFFF&) + (val4 And &HFFFF&)
    lngOverflow = lngLowWord \ 65536
    lngHighWord = (((Val1 And &HFFFF0000) \ 65536) + _
                   ((Val2 And &HFFFF0000) \ 65536) + _
                   ((val3 And &HFFFF0000) \ 65536) + _
                   ((val4 And &HFFFF0000) \ 65536) + _
                   lngOverflow) And &HFFFF&
    LongOverflowAdd4 = UnsignedToLong((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))
End Function

'
' Convert an unsigned double into a long
'
Private Function UnsignedToLong(value As Double) As Long
        If value < 0 Or value >= OFFSET_4 Then Error 6 ' Overflow
        If value <= MAXINT_4 Then
          UnsignedToLong = value
        Else
          UnsignedToLong = value - OFFSET_4
        End If
      End Function

'
' Convert a long to an unsigned Double
'
Private Function LongToUnsigned(value As Long) As Double
        If value < 0 Then
          LongToUnsigned = value + OFFSET_4
        Else
          LongToUnsigned = value
        End If
End Function

