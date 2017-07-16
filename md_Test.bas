Attribute VB_Name = "md_Test"
Option Explicit
'R 0071 <charactor ID>.l <map name>.16B <ip>.l <port>.w
Type DP0071
    CharID As Long
    MapName As String * 16
    IP(3) As Byte
    Port As Integer
End Type

Sub Test()
    
    'sts = LngToChr(51200)
    'stp = Conv2Arr(sts)
    'CopyMemory Tmps, stp(0), 8
    'MsgBox Tmps
    'msgBox ChrtoHex(LngToChr2(-1024000))
    
    'Dim i&, ir$, sttime&
    'MsgBox ChrtoHex(LngToChr2(1048576))
    '515 500 515 515
    'sttime = GetTickCount
    'For i = 1 To 100000
    '    ir = LngToChr2(i)
    'Next
    'MsgBox CStr(GetTickCount - sttime)
    End
End Sub

