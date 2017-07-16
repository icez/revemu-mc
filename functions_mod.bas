Attribute VB_Name = "Mod_Tools"
Option Explicit

'*******************************************************************************************
'Functions
'*******************************************************************************************
' Julien Lecomte
' webmaster@amanitamuscaria.org
' http://www.amanitamuscaria.org
' Feel free to use, abuse or distribute. (USUS, FRUCTUS, & ABUSUS)
' If you improve it, tell me !
' Don't take credit for what you didn't create. Thanks.
'*******************************************************************************************
'


Private Const OFFSET_4 = 4294967296#

Public Enum PEN_MODE
    PM_SETTILE = 0
    PM_SETSTART = 1
    PM_SETEND = 2
End Enum

Public Enum TILE_HARDNESS
    TH_EASY = 1
    TH_NORMAL = 3
    TH_HARD = 6
    TH_VERYHARD = 9
    TH_UNWALKABLE = 10
End Enum

Public Enum PATH_MAP
    PATH_IMPOSSIBLE = -2
    PATH_EMPTY = -1
    PATH_HUGE = 2147483647
End Enum
    
Public Const SLOW_DOWN_VALUE = 10 '//Milliseconds

Public NUMBER_OF_TILES&
Public TILE_SIDE&

Private Const RDW_INVALIDATE = &H1
Private Const RDW_ERASE = &H4
Private Const RDW_UPDATENOW = &H100
Public Version As String

Public Function FormatTime$(lTimeMilliseconds&)
    Dim lMilliseconds&
    Dim lSeconds&
    Dim lMinutes&
    
    lSeconds = lTimeMilliseconds \ 1000
    lMilliseconds = lTimeMilliseconds - lSeconds * 1000
    lMinutes = lSeconds \ 60
    lSeconds = lSeconds - lMinutes * 60
    
    FormatTime = Right$("00" & lMinutes, 2) & "m " & Right$("00" & lSeconds, 2) & "s " & Right$("000" & lMilliseconds, 3)
End Function

Public Function Find_HealItem(Name As String) As Long
On Error GoTo errie
Dim X&, Y&, strSPL() As String
strSPL = Split(LCase(Name), "&")

For Y = 0 To UBound(strSPL)
    For X = 0 To UBound(AllInv)
        If AllInv(X).Amount > 0 And LCase(AllInv(X).Name) = strSPL(Y) Then
            Find_HealItem = X
            Exit Function
        End If
    Next
Next

errie:
Find_HealItem = 0
Err.Clear
End Function

'Public Sub 'print_errror(tstr As String)
    'Label1.Caption = tstr
    'Open "errorlog.txt" For Output As #5
    '   Print #5, "Application ERROR at [" & tstr & "]"
    'Close #5
'End Sub


