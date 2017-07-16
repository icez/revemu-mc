Attribute VB_Name = "Mod_Crypt"
Public EnMode As Integer
Public PatternCode As Long
Public IsEncryption As Boolean

Public CryptFrame As String
Public Crypt1stNum As Long
Public Crypt2ndNum As Long
Public CryptModNum As Long

Declare Sub encode Lib "encode.dll" (ByVal Version As Long, ByVal servetype As Long, ByVal servicetype As Long, ByVal username As String, ByVal password As String, ByVal Key As String, ByRef crypted As String)
Declare Sub encrypt Lib "revemuext.dll" (ByVal username As String, ByVal password As String, _
ByVal Key As String, ByRef crypt As Byte)
Declare Function getkey Lib "revemuext.dll" (ByVal crypt1 As String, ByVal crypt2 As String) As Long
Declare Function isjunk Lib "revemuext.dll" (ByVal pattern As Long, ByVal code As Long) As Integer
Declare Function junkpos Lib "revemuext.dll" (ByVal pattern As Long, ByVal code As Long, ByVal modcode As Long) As Integer

Public Sub set_CryptMode(ennum As Integer)
    If ennum > 0 Then
        IsEncryption = True
    Else
        IsEncryption = False
    End If
    Select Case ennum
        Case 1
            Crypt1stNum = 1391
            Crypt2ndNum = 1397
            CryptModNum = 13
        Case 2
            Crypt1stNum = 2483
            Crypt2ndNum = 12435
            CryptModNum = 17
    End Select
End Sub

Public Function MakePort(ByVal rawPort As String) As Long
On Error GoTo buggy
MakePort = CLng(Asc(Left(rawPort, 1))) + (CLng(Asc(Mid(rawPort, 2, 1))) * 256)
Exit Function
buggy:
MakePort = 0
End Function

Public Function GenJunk(Number) As String
    Dim tstr As String
    Dim i As Integer
    tstr = ""
    For i = 0 To Number - 1
        tstr = tstr & Chr(RandomNumber(255, 1))
    Next
    GenJunk = tstr
End Function

Public Function Encode_Crypt(packet As String, pattern As Long) As String
    Dim junkmask() As Byte
    Dim tstr As String
    Dim endframe As Long
    Dim loopin, loopout As Long
    loopout = 0
    ReDim junkmask(CryptModNum)
    If isjunk(pattern, Crypt1stNum) = 1 Then
        junkmask(junkpos(pattern, Crypt1stNum, CryptModNum)) = 1
    End If
    junkmask((pattern * Crypt2ndNum) Mod CryptModNum) = 1
    For loopin = 0 To Len(packet) - 1
        If junkmask(loopout Mod CryptModNum) = 1 Then
            tstr = tstr & Chr(RandomNumber(255, 1))
            loopout = loopout + 1
        End If
            tstr = tstr & Mid(packet, loopin + 1, 1)
            loopout = loopout + 1
    Next
    loopout = loopout + 4
    tstr = Make2Byte(pattern) & tstr
    tstr = Make2Byte(loopout) & tstr
    endframe = EndPOS(loopout)
    tstr = tstr & GenJunk(endframe - loopout)
    Encode_Crypt = tstr
End Function

Public Function EndPOS(lenght As Long) As Long
    Dim Number As Integer
    Number = 4
    Do While (lenght > Number)
        Number = Number + 8
    Loop
    EndPOS = Number
End Function

Public Function Decode_Crypt() As String
On Error GoTo errie
    Dim junkmask() As Byte
    Dim tstr As String
    Dim tstr2 As String
    Dim pattern As Long
    Dim lenght As Integer
    Dim packet As String
    Dim i As Integer
    packet = ""
restart:
    ReDim junkmask(CryptModNum)
    lenght = MakePort(Left(CryptFrame, 2))
    If Len(CryptFrame) < lenght Then
        GoTo endcode
    End If
    pattern = MakePort(Mid(CryptFrame, 3, 2))
    tstr = Mid(CryptFrame, 5, lenght - 4)
    CryptFrame = Right(CryptFrame, Len(CryptFrame) - EndPOS(MakePort(Left(CryptFrame, 2))))
    If isjunk(pattern, Crypt1stNum) = 1 Then junkmask(junkpos(pattern, Crypt1stNum, CryptModNum)) = 1
    If (pattern * Crypt2ndNum) >= 0 Then junkmask((pattern * Crypt2ndNum) Mod CryptModNum) = 1
    For i = 0 To lenght - 1
        If (junkmask(i Mod CryptModNum) = 0) Then packet = packet & Mid(tstr, i + 1, 1)
    Next
    If Len(CryptFrame) > 4 Then GoTo restart
endcode:
    Decode_Crypt = packet
    Exit Function
errie:
    Decode_Crypt = ""
End Function
