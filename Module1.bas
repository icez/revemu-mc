Attribute VB_Name = "Module1"
Option Explicit

Declare Function DESEncrypt Lib "desdll.dll" (ByVal strIn As String, ByVal Key As String, ByVal strOut As String) As Integer
Declare Function DESDecrypt Lib "desdll.dll" (ByVal strIn As String, ByVal Key As String, ByVal strOut As String) As Integer
Declare Function DESEncryptFile Lib "desdll.dll" _
(ByVal strIn As String, ByVal Key As String, ByVal strOut As String) As Integer
Declare Function DESDecryptFile Lib "desdll.dll" _
(ByVal strIn As String, ByVal Key As String, ByVal strOut As String) As Integer


Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, _
ByVal lpCaption As String, ByVal wType As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Declare Function ntohl Lib "wsock32" (ByVal netlong As Long) As Long
Declare Function ntohs Lib "wsock32" (ByVal netshort As Integer) As Integer
Declare Function GetTickCount Lib "kernel32" () As Long
 
Public strUser As String, StrPass As String, strServer As String, _
revkey As String, revpass As String, MapName As String, _
hwinfo As String, MapPath As String, TmpAggroName As String


Public Type MonsList
    ID As String * 4
End Type

Public Type UnknowList
    ID As String * 4
    Name As String
    NameID As String
End Type

Public Type Coord
    X As Long
    Y As Long
End Type

Type MonsterPos
    ID As String * 4
    Pos As Coord
    Speed As Integer
    NextPos As Coord
    Time As Long
    Endtime As Long
    NameID As Integer
    StatusA As Integer
    StatusB As Integer
    NoAttack As Boolean
    IsPet As Boolean
    IsAttack As Boolean
    IsFollow As Boolean
    IsTrap As Boolean
    CantGo As Boolean
    TargetID As String * 4
    Name As String
End Type

Type itemName
    ID As String * 4
    Pos As Coord
    Name As String
End Type

Type ItemPos
    ID As String * 4
    Pos As Coord
    Name As String * 24
End Type

Type Item
    CantKeep As Boolean
    Amount As Long
    Name As String
End Type

Type StatusChar
    Aspd As Integer
    Class As String
    ClassID As Integer
    Name As String
    BaseLV As Byte
    JobLV As Byte
    StatPoint As Integer
    HP As Integer
    MaxHP As Integer
    SP As Integer
    maxsp As Integer
    Weight As Integer
    MaxWeight As Integer
    Zeny As Double
    BaseExp As Double
    NextBaseEXP As Double
    JobExp As Double
    MaxJobEXP As Double
    ATK As Integer
    ATKp As Integer
    MaxMatk As Integer
    MinMatk As Integer
    Hit As Integer
    Crit As Integer
    Def As Integer
    Defp As Integer
    mDef As Integer
    mDefp As Integer
    Flee As Integer
    Fleep As Integer
    STR As Integer
    Strp As Integer
    AGI As Integer
    Agip As Integer
    VIT As Integer
    Vitp As Integer
    Intl As Integer
    Intp As Integer
    DEX As Integer
    Dexp As Integer
    LUK As Integer
    Lukp As Integer
    Party As String
    Guild As String
End Type

Type Skill
    Name As String
    ID As Integer
    Lv As Byte
    SP As Integer
    MaxLV As Byte
    Target As Integer
End Type

Type ItemInv
    Index As String * 2
    Type As String * 2
    Identified As Boolean
    Category As Byte
    Pos As Long
    Price As Long
    Amount As Long
    Name As String
    NameID As String
End Type

Type MonsterSlot
    NameID As Byte
End Type

Type Server
    Name As String
    IP As String
    Port As Integer
    number As Long
End Type

Type Itemlist
    Name As String
    Amount As Integer
End Type

Type skillmobs
    Packet As String
    rawname As String
    Lv As Byte
    MonsName As String
    number As Byte
    SP As Byte
End Type

Type UseItem
    Use As Boolean
    percent As Double
    Name As String
End Type

Public SPItem As UseItem
Public AutoShare As Boolean
Public CharIdStart As Byte
Public IsNomonsSit As Boolean
Public Dead As Boolean
Public Unknow() As UnknowList
Public RareItem() As Itemlist
Public MobSkill As skillmobs
Public UseSkillMobs As Boolean
Public MobSkill2 As skillmobs
Public UseSkillMobs2 As Boolean
'Public AutoWing As Boolean
Public giveuptime As Integer
Public SWeight1 As Boolean
Public Weight1 As Double
Public SWeight2 As Boolean
Public Weight2 As Double
Public Itempick() As Itemlist
Public RandomMove As Boolean
Public Movetime As Integer
Public Tanker As MonsterPos
Public PartyMode As Boolean
Public TankerID As String * 4
Public PortalTime As Integer
Public NumServ As Byte
Public ServerList() As Server
Public IsUseRange As Boolean
Public RangeSet As Integer
Public IsDamageDC As Boolean
Public DamageSet As Integer
Public SkillList() As Skill
Public SkillChar() As Skill
Public Players() As StatusChar
Public number As Integer
Public SkillNumber As Integer
Public IsUseSkill As Boolean
Public skillpacket As String
Public MageMode As Boolean
Public Range As Integer
Public AttCounter As Integer
Public Store() As ItemInv
Public AllInv() As ItemInv
Public IsAutoKill As Boolean
Public IsAutoPick As Boolean
'Public IsAutoSell  As Boolean
'Public IsAutoSell2 As Boolean
Public IsAutoRedz  As Boolean
Public IsAutoOrange As Boolean
Public IsAutorest As Boolean
Public InitLoad As Boolean
Public CounterTime As Integer
Public IsConnected As Boolean
Public IsSPWait As Boolean
Public SPWait As Double
Public IsSkillUse As Boolean
Public IsSelectSkill As Boolean
Public IsAutoDC As Boolean
Public HPSit As Double
Public IsSPSit As Boolean
Public SPSit As Double
Public HPRed As Double
Public HPOrange As Double
Public HPDC As Double
Public SkillSelect As Integer
Public SendAction As Boolean
Public UseWeapon As Boolean
Public Autoheal As Boolean
Public Automove As Boolean
Public healitem1 As String
Public healitem2 As String
Public HPHeal As Double
Public HealLV As Integer
Public Reconcount As Integer
Public DelayTime As Integer
Public WarpDelay As Integer
Public ResponseTime As Integer
Public killsteal As Boolean
Public AvoidWarp As Boolean
Public NomonsTime As Integer
Public NomonsWarp As Boolean
Public AutoDCCase As Integer
Public AutoDC2Case As Integer
Public ServerID As Byte
Public LoginIP As String
Public HPWait As Double
Public IsHPWait As Boolean
Public AcoHealName As String
Public IsWantHeal As Boolean


'RegEdit to Save Windows Position
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const REG_SZ = 1 'Unicode nul terminated string
Public Const REG_BINARY = 3 'Free form binary
Public Const REG_DWORD = 4 '32-bit number
Public Const ERROR_SUCCESS = 0&

Public Declare Function RegCloseKey Lib "advapi32.dll" _
(ByVal hKey As Long) As Long

Public Declare Function RegCreateKey Lib "advapi32.dll" _
Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey _
As String, phkResult As Long) As Long

Public Declare Function RegDeleteKey Lib "advapi32.dll" _
Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey _
As String) As Long

Public Declare Function RegDeleteValue Lib "advapi32.dll" _
Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal _
lpValueName As String) As Long

Public Declare Function RegOpenKey Lib "advapi32.dll" _
Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey _
As String, phkResult As Long) As Long

Public Declare Function RegQueryValueEx Lib "advapi32.dll" _
Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName _
As String, ByVal lpReserved As Long, lpType As Long, lpData _
As Any, lpcbData As Long) As Long

Public Declare Function RegSetValueEx Lib "advapi32.dll" _
Alias "RegSetValueExA" (ByVal hKey As Long, ByVal _
lpValueName As String, ByVal Reserved As Long, ByVal _
dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Sub SaveFormPos(frmSave As Form)
Dim strRegPath As String
strRegPath = "Software\" & App.CompanyName _
& "\" & App.Title & "\" & frmSave.Name

If frmSave.WindowState = vbMaximized Then
SaveSettingLong HKEY_CURRENT_USER, strRegPath, "Maximised", 1
DeleteValue HKEY_CURRENT_USER, strRegPath, "Left"
DeleteValue HKEY_CURRENT_USER, strRegPath, "Top"
DeleteValue HKEY_CURRENT_USER, strRegPath, "Width"
DeleteValue HKEY_CURRENT_USER, strRegPath, "Height"
Else
With frmSave
SaveSettingLong HKEY_CURRENT_USER, strRegPath, "Maximised", 0
SaveSettingLong HKEY_CURRENT_USER, strRegPath, "Left", .Left
SaveSettingLong HKEY_CURRENT_USER, strRegPath, "Top", .Top
SaveSettingLong HKEY_CURRENT_USER, strRegPath, "Width", .width
SaveSettingLong HKEY_CURRENT_USER, strRegPath, "Height", .height
End With
End If

End Sub

Public Sub LoadFormPosOnly(frmLoad As Form)
Dim strRegPath As String
Dim IsMax As Long
strRegPath = "Software\" & App.CompanyName _
& "\" & App.Title & "\" & frmLoad.Name
IsMax = GetSettingLong(HKEY_CURRENT_USER, strRegPath, "Maximised", 2)

Select Case IsMax
Case 0
With frmLoad
.Left = GetSettingLong(HKEY_CURRENT_USER, strRegPath, "Left", .Left)
.Top = GetSettingLong(HKEY_CURRENT_USER, strRegPath, "Top", .Top)
'.Move GetSettingLong(HKEY_CURRENT_USER, strRegPath, "Left", .Left), _
'GetSettingLong(HKEY_CURRENT_USER, strRegPath, "Top", .Top), _
'frmLoad.width, frmLoad.height
End With

Case 1
'frmLoad.WindowState = vbMaximized

Case 2
'MsgBox "There is no form data saved for this form"

End Select

End Sub

Public Sub LoadFormPos(frmLoad As Form)
Dim strRegPath As String
Dim IsMax As Long
strRegPath = "Software\" & App.CompanyName _
& "\" & App.Title & "\" & frmLoad.Name
IsMax = GetSettingLong(HKEY_CURRENT_USER, strRegPath, "Maximised", 2)

Select Case IsMax
Case 0
With frmLoad
.Move GetSettingLong(HKEY_CURRENT_USER, strRegPath, "Left", .Left), _
GetSettingLong(HKEY_CURRENT_USER, strRegPath, "Top", .Top), _
GetSettingLong(HKEY_CURRENT_USER, strRegPath, "Width", .width), _
GetSettingLong(HKEY_CURRENT_USER, strRegPath, "Height", .height)
End With

Case 1
'frmLoad.WindowState = vbMaximized

Case 2
'MsgBox "There is no form data saved for this form"

End Select

End Sub

Public Function GetSettingString(hKey As Long, _
strPath As String, strValue As String, Optional _
Default As String) As String
Dim hCurKey As Long
Dim lResult As Long
Dim lValueType As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim intZeroPos As Integer
Dim lRegResult As Long

'Set up default value
If Not IsEmpty(Default) Then
GetSettingString = Default
Else
GetSettingString = ""
End If

lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, _
lValueType, ByVal 0&, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

If lValueType = REG_SZ Then

strBuffer = String(lDataBufferSize, " ")
lResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, _
ByVal strBuffer, lDataBufferSize)

intZeroPos = InStr(strBuffer, Chr$(0))
If intZeroPos > 0 Then
GetSettingString = Left$(strBuffer, intZeroPos - 1)
Else
GetSettingString = strBuffer
End If

End If

Else
'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Function

Public Sub SaveSettingString(hKey As Long, strPath _
As String, strValue As String, strData As String)
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegCreateKey(hKey, strPath, hCurKey)

lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, _
ByVal strData, Len(strData))

If lRegResult <> ERROR_SUCCESS Then
'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function GetSettingLong(ByVal hKey As Long, _
ByVal strPath As String, ByVal strValue As String, _
Optional Default As Long) As Long

Dim lRegResult As Long
Dim lValueType As Long
Dim lBuffer As Long
Dim lDataBufferSize As Long
Dim hCurKey As Long

'Set up default value
If Not IsEmpty(Default) Then
GetSettingLong = Default
Else
GetSettingLong = 0
End If

lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lDataBufferSize = 4 '4 bytes = 32 bits = long

lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, _
lValueType, lBuffer, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

If lValueType = REG_DWORD Then
GetSettingLong = lBuffer
End If

Else
'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Function

Public Sub SaveSettingLong(ByVal hKey As Long, ByVal _
strPath As String, ByVal strValue As String, ByVal _
lData As Long)
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegCreateKey(hKey, strPath, hCurKey)

lRegResult = RegSetValueEx(hCurKey, strValue, 0&, _
REG_DWORD, lData, 4)

If lRegResult <> ERROR_SUCCESS Then
'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Sub

Public Sub DeleteValue(ByVal hKey As Long, _
ByVal strPath As String, ByVal strValue As String)
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegOpenKey(hKey, strPath, hCurKey)

lRegResult = RegDeleteValue(hCurKey, strValue)

lRegResult = RegCloseKey(hCurKey)

End Sub

'End RegEdit to Save Windows Position
