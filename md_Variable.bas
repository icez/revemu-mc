Attribute VB_Name = "md_Variable"
Option Explicit

Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer

Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Public LockmapList() As mcScriptLockmap
Type mcScriptLockmap
    LvMin As Byte
    LvMax As Byte
    MapName As String
End Type


Public FollowMode As mcFollowMode
'Public TankerMode As mcFollowMode
Type mcFollowMode
    AID As String * 4
    Name As String
    curPos As Coord
    Active As Boolean
    AutoBuff As Boolean
    NoAttack As Boolean
End Type
Public UseStatLog As Boolean
Public StartPriority As String
Public WaitEquipTele As Boolean
Public WaitEquipBack As Boolean
Public isWaitWarpSave As Boolean
Public LastGetStorage As Long
Public tmpEQOldPos As Integer
Public tmpEQTelePos As Integer
Public tmpEQOldName As String
Public tmpEQTeleName As String
Public TeleNothing As Boolean
Public UseAutoSpell As Boolean
Public AutoSpell_Name As String
'Public nKillSteal As Boolean
Public curEnableKey As Boolean
Public HauntedStep As Byte

Public AtkMode As Boolean
Public ForceTeleport As Boolean
Public isBackStore As Boolean
Public UseNPCBiDirect As Boolean
Public MobTeleNum As Integer
Public MODDelay As zMODDelay
Public MODDC As zMODDC
Public MODHead() As zMODHead
Public ItemCtrl() As MODItemCtrl
Public MODMLogN() As String
Public MODMLogM() As MODMonsterLog2
Public DetectFail As Boolean
Public SkillOnly As Boolean
Public StartTime As Long
Public RestartTime As Long
Public UseRestart As Boolean
Public IsUseProxy As Boolean
Public ProxyIP As String
Public ProxyPort As Long
Public ProxyType As Byte
Public CurConnIP As String
Public CurConnPort As Long
Public ProxyConn As Boolean
Public allData As String
Public ProxyUser As String
Public ProxyPass As String
Public ProxyStep As Byte

Type MODMonsterLog2
    Names() As String
    Amount() As Long
End Type

Type MODItemCtrl
    Name As String
    Price As String
    Lock As Boolean
    Reject As Boolean
End Type

Type zMODDelay
    DualLogin As Double
End Type

Type zMODDC
    DualLogin As Boolean
    DualLoginTime As Double
    AvoidTime As Double
    TAccept As Long
    TCalc As Long
    TItem As Long
    TNItem As Long
End Type

Type zMODHead
    Class As Integer
    ItemID As Integer
    Name As String
End Type

Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public Party() As zMODParty
'R 00fb <len>.w <party name>.24B {<ID>.l <nick>.24B <map name>.16B <leader>.B <offline>.B}.46B*
Type zMODParty
    ID As String * 4
    Name As String
    pos As Coord
    NextPos As Coord
    HpMin As Integer
    HPmax As Integer
    Admin As Boolean
    Map As String
    Online As Boolean
End Type

Private Type Map
    Name As String
    pos As Coord
End Type

Private Type Portal
    Src As Map
    Des As Map
End Type

Public PortalsInfo() As Portal
