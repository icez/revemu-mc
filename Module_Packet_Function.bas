Attribute VB_Name = "Mod_Packet_Function"
Option Explicit

Public Sub Chat(Stri As String, Optional Color As Double = 0)
'print_errror "sub Chat"
Dim sel_start As Long
sel_start = Len(frmChat.txtChat.text)
frmChat.txtChat.SelStart = sel_start
frmChat.txtChat.SelLength = 0
frmChat.txtChat.SelText = Stri
frmChat.txtChat.SelText = vbCrLf

frmChat.txtChat.SelStart = sel_start
frmChat.txtChat.SelLength = Len(Stri)
frmChat.txtChat.SelColor = Color

If Len(frmChat.txtChat.text) > 5120 Then
    frmChat.txtChat.SelStart = 0
    frmChat.txtChat.SelLength = 2560
    frmChat.txtChat.SelText = ""
End If
frmChat.txtChat.SelStart = Len(frmChat.txtChat.text)
frmChat.txtChat.SelLength = 0
If (MDIfrmMain.mnuChatLog.CheckED) Then
Open ChatlogPath For Append As #1
    Print #1, "[" & Time & "] " & Left(Stri, Len(Stri))
Close #1
End If
End Sub

Public Sub Stat(Stri As String, Optional Color As Double = 0, Optional IsUnderline As Boolean = False, Optional IsBold As Boolean = False, Optional IsItalic As Boolean = False)
Dim stas As Long
stas = Len(frmMain.txtStatus.text)
frmMain.txtStatus.SelStart = stas
frmMain.txtStatus.SelText = Stri

frmMain.txtStatus.SelStart = Len(frmMain.txtStatus.text) - Len(Stri)
frmMain.txtStatus.SelLength = Len(Stri)

frmMain.txtStatus.SelColor = Color
If IsUnderline Then frmMain.txtStatus.SelUnderline = 1 Else frmMain.txtStatus.SelUnderline = 0
If IsBold Then frmMain.txtStatus.SelBold = 1 Else frmMain.txtStatus.SelBold = 0
If IsItalic Then frmMain.txtStatus.SelItalic = 1 Else frmMain.txtStatus.SelItalic = 0

If stas > 10240 Then
    frmMain.txtStatus.SelStart = 0
    frmMain.txtStatus.SelLength = 2560
    frmMain.txtStatus.SelText = ""
End If
frmMain.txtStatus.SelStart = Len(frmMain.txtStatus.text)
'If UseStatLog Then
'    Put #100, LOF(100), CStr(Replace(Stri, vbCrLf, vbCrLf & vbLf))
'End If
End Sub

Public Function MakeText(tstr As String) As String
    'print_errror "sub MakeText"
Dim i As Integer
i = 1
While (Asc(Mid(tstr, i, 1)) <> 0)
i = i + 1
Wend
MakeText = Left(tstr, i - 1)
End Function

Public Function MakeItemPos(Coords As Coord) As String
Dim tstr As String
Dim Offset As Integer
Dim newcoords As Coord
Dim CurAngle As Integer
Dim GetAngle As Integer
Dim i As Long
Offset = 1
For i = 0 To UBound(AllowCoord)
    newcoords.X = MapHeight - AllowCoord(i).Y
    newcoords.Y = AllowCoord(i).X
    If EvalNorm(newcoords, Coords) = Offset Then
        If CanGO(curPos, newcoords) Then
            CurAngle = Arctan((Coords.X - curPos.X), (Coords.Y - curPos.Y))
            GetAngle = Arctan((Coords.X - newcoords.X), (Coords.Y - newcoords.Y))
            If Abs(CurAngle - GetAngle) < 90 Then Exit For
        End If
    End If
Next

tstr = tstr + Chr(Int((newcoords.Y) / 4))
tstr = tstr + Chr(((newcoords.Y) Mod 4) * 64 + Int((newcoords.X) / 16))
tstr = tstr + Chr(((newcoords.X) Mod 16) * 16)
MakeItemPos = tstr
End Function

'Public Function GetLong(rawPort As String) As Long
'On Error GoTo errie
'Dim tst As Long
'Dim i As Integer
'''print_errror "sub getlong"
'tst = CLng(Asc(Mid(rawPort, 1, 1))) + (CLng(Asc(Mid(rawPort, 2, 1))) * 256) + (CLng(Asc(Mid(rawPort, 3, 1))) * 65536)
'If Asc(Mid(rawPort, 4, 1)) < 8 Then
'    For i = 1 To Asc(Mid(rawPort, 4, 1))
'        tst = tst + 16777216
'    Next
'End If
'GetLong = tst
'Exit Function
'errie:
''MsgBox Err.Description, vbCritical
'GetLong = 0
'End Function

Public Function MakeIP(rawIP As String) As String
'print_errror "sub MakeIP"
Dim str1 As String
Dim X As Integer
For X = 1 To 4
    str1 = str1 + CStr(Asc(Mid(rawIP, X, 1))) + "."
Next
str1 = Left(str1, Len(str1) - 1)
MakeIP = str1
End Function

Public Function MakeCoordPos(Coords As Coord) As String
Dim tstr As String
Dim Offset As Integer
Dim newcoords As Coord
Dim sparecoods As Coord
Dim found As Boolean
Dim CurAngle As Integer
Dim GetAngle As Integer
Dim i As Long
Offset = 1
found = False
For i = 0 To UBound(AllowCoord)
    newcoords.X = MapHeight - AllowCoord(i).Y
    newcoords.Y = AllowCoord(i).X
    If EvalNorm(newcoords, Coords) = Offset Then
        If CanGO(curPos, newcoords) Then
            sparecoods = newcoords
            CurAngle = Arctan((Coords.X - curPos.X), (Coords.Y - curPos.Y))
            GetAngle = Arctan((Coords.X - newcoords.X), (Coords.Y - newcoords.Y))
            If Abs(CurAngle - GetAngle) < 20 Then
                found = True
                Exit For
            End If
        End If
    End If
Next
If Not found Then newcoords = sparecoods
tstr = tstr + Chr(Int((newcoords.Y) / 4))
tstr = tstr + Chr(((newcoords.Y) Mod 4) * 64 + Int((newcoords.X) / 16))
tstr = tstr + Chr(((newcoords.X) Mod 16) * 16)
MakeCoordPos = tstr
End Function

Public Function MakeHex(rawLong As String) As String
''print_errror "sub MakeHex"
On Error Resume Next
Dim str1 As String
Dim X As Integer
For X = 1 To 4
    If Asc(Mid(rawLong, X, 1)) < 16 Then str1 = str1 + "0"
    str1 = str1 + Hex(Asc(Mid(rawLong, X, 1)))
Next
Err.Clear
MakeHex = str1
End Function

Public Function MakeHexName(ByVal rawLong As String) As String
''print_errror "sub MakeHexName"
On Error Resume Next
Dim str1 As String
Dim X As Integer
For X = Len(rawLong) To 1 Step -1
    If Asc(Mid(rawLong, X, 1)) < 16 Then str1 = str1 + "0"
    str1 = str1 + Hex(Asc(Mid(rawLong, X, 1)))
Next
Err.Clear
MakeHexName = str1

End Function

Public Function MakeString(rawString As String) As String
If InStr(rawString, Chr(0)) > 0 Then MakeString = Left(rawString, InStr(rawString, Chr(0)) - 1) Else MakeString = rawString
End Function

Public Function MakeCoordsSec(rawCoords As String) As Coord
If Len(rawCoords) < 3 Then GoTo errie
On Error GoTo errie
Dim xint As Long
Dim yint As Long
xint = (Asc(Mid(rawCoords, 2, 1)) And &H3) * 256
xint = xint + Asc(Mid(rawCoords, 3, 1))
yint = (Asc(Mid(rawCoords, 2, 1)) And &HFC) / 4
yint = yint + (Asc(Mid(rawCoords, 1, 1)) And &HF) * 64
MakeCoordsSec.Y = yint
MakeCoordsSec.X = xint
Exit Function
errie:
MakeCoordsSec.Y = 0
MakeCoordsSec.X = 0
End Function


Public Function MakeTickString() As String
''print_errror "sub MakeTickString"
On Error GoTo Out
Dim lng1 As Long
Dim X As Integer
Dim byt1(3) As Byte
lng1 = GetTickCount
CopyMemory byt1(0), lng1, 4
Dim str1 As String
For X = 0 To 3
    str1 = str1 + Chr(byt1(X))
Next
If MasterSelect.pkserver = 1 Then MakeTickString = str1 & Chr(Sex) Else MakeTickString = str1
Out:
End Function
