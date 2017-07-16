VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIfrmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H00808080&
   Caption         =   "Revemu-MC"
   ClientHeight    =   7605
   ClientLeft      =   4140
   ClientTop       =   3525
   ClientWidth     =   9780
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   9780
      TabIndex        =   1
      Top             =   0
      Width           =   9780
      Begin VB.Frame Frame1 
         Height          =   1050
         Left            =   0
         TabIndex        =   2
         Top             =   -75
         Width           =   9780
         Begin VB.Image imgStatus 
            Height          =   480
            Index           =   0
            Left            =   120
            Top             =   420
            Width           =   480
         End
         Begin VB.Image imgStatus 
            Height          =   480
            Index           =   1
            Left            =   720
            Top             =   420
            Width           =   480
         End
         Begin VB.Image imgStatus 
            Height          =   480
            Index           =   2
            Left            =   1320
            Top             =   420
            Width           =   480
         End
         Begin VB.Image imgStatus 
            Height          =   480
            Index           =   3
            Left            =   1920
            Top             =   420
            Width           =   480
         End
         Begin VB.Image imgStatus 
            Height          =   480
            Index           =   4
            Left            =   2520
            Top             =   420
            Width           =   480
         End
         Begin VB.Image imgStatus 
            Height          =   480
            Index           =   5
            Left            =   3120
            Top             =   420
            Width           =   480
         End
         Begin VB.Image imgStatus 
            Height          =   480
            Index           =   6
            Left            =   3720
            Top             =   420
            Width           =   480
         End
         Begin VB.Image imgStatus 
            Height          =   480
            Index           =   7
            Left            =   4320
            Top             =   420
            Width           =   480
         End
         Begin VB.Image imgStatus 
            Height          =   480
            Index           =   8
            Left            =   4920
            Top             =   420
            Width           =   480
         End
         Begin VB.Image imgStatus 
            Height          =   480
            Index           =   9
            Left            =   5520
            Top             =   420
            Width           =   480
         End
         Begin VB.Image imgStatus 
            Height          =   480
            Index           =   10
            Left            =   6120
            Top             =   420
            Width           =   480
         End
         Begin VB.Image imgStatus 
            Height          =   480
            Index           =   11
            Left            =   6720
            Top             =   420
            Width           =   480
         End
         Begin VB.Image imgStatus 
            Height          =   480
            Index           =   12
            Left            =   7320
            Top             =   420
            Width           =   480
         End
         Begin VB.Image imgStatus 
            Height          =   480
            Index           =   13
            Left            =   7920
            Top             =   420
            Width           =   480
         End
         Begin VB.Label bState 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   180
            Width           =   45
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7230
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuCon 
      Caption         =   "&Connect"
      Begin VB.Menu login 
         Caption         =   "Login"
      End
      Begin VB.Menu mnuRecon 
         Caption         =   "Reconnect"
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu mnuReset 
         Caption         =   "Reset"
         Shortcut        =   +{INSERT}
      End
   End
   Begin VB.Menu mnuServer 
      Caption         =   "&Server"
   End
   Begin VB.Menu mnuReload 
      Caption         =   "&Reload"
      Begin VB.Menu mnuAll 
         Caption         =   "All *.txt"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuProfile 
         Caption         =   "Profile"
         Begin VB.Menu mnuAllProfile 
            Caption         =   "All Profile"
         End
         Begin VB.Menu mnuEqmons 
            Caption         =   "equip_monster.txt"
         End
         Begin VB.Menu mnuSelfSkill 
            Caption         =   "selfskill.txt"
         End
         Begin VB.Menu mnuNPCReload 
            Caption         =   "npc.txt"
         End
         Begin VB.Menu mnuRChatResp 
            Caption         =   "events.txt"
         End
         Begin VB.Menu mnuRecovery 
            Caption         =   "recovery_profile.txt"
         End
      End
      Begin VB.Menu mnuAttack 
         Caption         =   "attack.txt"
      End
      Begin VB.Menu mnuRarelist 
         Caption         =   "rarelist.txt"
      End
      Begin VB.Menu mnuDrop 
         Caption         =   "droplist.txt"
      End
      Begin VB.Menu mnuAvoid 
         Caption         =   "avoidlist.txt"
      End
      Begin VB.Menu mnuautoNpc 
         Caption         =   "autonpc.txt"
      End
      Begin VB.Menu mnuWarpList 
         Caption         =   "warplist.txt"
      End
      Begin VB.Menu mnuSell 
         Caption         =   "sell.txt"
      End
      Begin VB.Menu mnuKeep 
         Caption         =   "keep.txt"
      End
      Begin VB.Menu mnuBuy 
         Caption         =   "buy.txt"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "options.ini"
      End
      Begin VB.Menu mnuStatustxt 
         Caption         =   "status.txt"
      End
      Begin VB.Menu mnuSpecialStatus 
         Caption         =   "specialstatus.txt"
      End
      Begin VB.Menu mnuRMod 
         Caption         =   "mods.ini"
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Actions"
      Begin VB.Menu mnuUseRed 
         Caption         =   "Use Redz Pot"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuUseOrange 
         Caption         =   "Use Orangez Pots"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuTeleport 
         Caption         =   "Random Teleportation"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSit 
         Caption         =   "Sit"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuStand 
         Caption         =   "Stand"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuShare 
         Caption         =   "Set Party Sharing Exp"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuUnshare 
         Caption         =   "Unset Party Sharing Exp"
         Checked         =   -1  'True
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnRecShop 
         Caption         =   "Re-Create Shop"
      End
      Begin VB.Menu mnuSendRaw 
         Caption         =   "Send Raw Packet"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Return to Save Point"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuPriori 
         Caption         =   "Set Priority"
         Begin VB.Menu mnuPri 
            Caption         =   "Realtime"
            Index           =   0
         End
         Begin VB.Menu mnuPri 
            Caption         =   "High"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuPri 
            Caption         =   "AboveNormal"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu mnuPri 
            Caption         =   "Normal"
            Index           =   3
         End
         Begin VB.Menu mnuPri 
            Caption         =   "BelowNormal"
            Enabled         =   0   'False
            Index           =   4
         End
         Begin VB.Menu mnuPri 
            Caption         =   "Low"
            Index           =   5
         End
      End
   End
   Begin VB.Menu mnuLog 
      Caption         =   "&Log"
      Begin VB.Menu mnuChatLog 
         Caption         =   "ChatLog"
      End
      Begin VB.Menu mnuPKTLOG 
         Caption         =   "PacketLog"
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "&Options"
      Begin VB.Menu mnuCheatAggro 
         Caption         =   "Cheat Aggro!!"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAI 
         Caption         =   "AI Options"
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuAutoPick 
         Caption         =   "Auto pick up item."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMove 
         Caption         =   "Auto move."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAutoSell1 
         Caption         =   "Auto sell when no monster."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAutoSell2 
         Caption         =   "Auto sell when got Item."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRandomMove 
         Caption         =   "Random move when no monster every # second(s)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBackTown 
         Caption         =   "Back to to town when weight reach xx%"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSAttack 
         Caption         =   "Stop attack when weight reach xx%"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSPick 
         Caption         =   "Stop pick up when weight reach xx%"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHPSP 
         Caption         =   "HP/SP Options"
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuAutosit 
         Caption         =   "Auto sit when HP below"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHPWait 
         Caption         =   "Sit until HP reach"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSPSit 
         Caption         =   "Auto sit when SP below"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSPWait 
         Caption         =   "Sit until SP reach"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRedz 
         Caption         =   "Auto drink Redz when HP below"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOrange 
         Caption         =   "Auto drink Orangez when HP below"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAtt 
         Caption         =   "Attack Options"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuAutoKill 
         Caption         =   "Auto attack."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAutoSkill 
         Caption         =   "Auto skill use."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSkillMobs 
         Caption         =   "Use Bowling Bash Lv.10 when mobs > 3"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSkillMobs2 
         Caption         =   "Use Bowling Bash Lv.10 when mobs > 3"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRange 
         Caption         =   "Attack if Range is locked as xx blocks."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuKillSteal 
         Caption         =   "Attemp to kill steal."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWeapon 
         Caption         =   "Use Weapon to Attack (Mage/Wizard)."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHeal 
         Caption         =   "Auto heal."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTeleportloot 
         Caption         =   "Teleport Option"
         Shortcut        =   +{F6}
      End
      Begin VB.Menu mnuDamage 
         Caption         =   "Auto teleport when Damage over"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAutoDC 
         Caption         =   "Auto teleport When HP below"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNomons 
         Caption         =   "Auto teleport when no monster for "
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWarp 
         Caption         =   "Auto teleport when found warp/exit portal."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCast 
         Caption         =   "Cast Anywhere."
         Visible         =   0   'False
      End
      Begin VB.Menu mnu 
         Caption         =   "Modificate Option"
         Shortcut        =   +{F7}
      End
      Begin VB.Menu mnuOptMonsAtk 
         Caption         =   "Monster Attack"
         Shortcut        =   +{F8}
      End
      Begin VB.Menu mnuParty 
         Caption         =   "Party"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuWin 
      Caption         =   "&Windows"
      Begin VB.Menu mnuCas 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuPeople 
         Caption         =   "People"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuPlayer 
         Caption         =   "Player"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "Player Status"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuMonster 
         Caption         =   "Monster"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuNPC 
         Caption         =   "NPC"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuChat 
         Caption         =   "Chat"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuArmor 
         Caption         =   "Weapon/Armor"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuInv 
         Caption         =   "Inventory"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuItemD 
         Caption         =   "Item Description"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuMain 
         Caption         =   "Main"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuSkill 
         Caption         =   "Skill"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuMap 
         Caption         =   "Map"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuStat 
         Caption         =   "Stats Info"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnWCart 
         Caption         =   "Cart"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuChatShop 
         Caption         =   "Chatroom/Shop"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuWinParty 
         Caption         =   "Party"
      End
      Begin VB.Menu mnuGuild 
         Caption         =   "Guild"
      End
   End
End
Attribute VB_Name = "MDIfrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Error As Boolean

Private Sub MDIForm_Load()
On Error GoTo errie
LoadFormPos Me

'Me.Show
'Load frmConfOptions
'frmConfOptions.Show
'WaitTime 600000
'End
'frmRegistration.Check_Key
Error = False
Load_Response
ReDim CurRoute(0)
PetWinClose = False
MDIfrmMain.Caption = Version
CreatIcon Version
ResetMod
ReDim SkillChar(0)
ReDim MapRoute(0)
StartTime = GetTickCount

'mod
ReadModOption
Load_IPList
CheckEvent "OnStart", "nothingtocheck=True"

'avoid folder
Load_AvoidID
Load_Monswarplist
Load_Warplist
Load_Avoidlist

'table folder
Load_Special_Status
Load_Char_Status
Load_Emotion
Load_Rarelist
Load_Droplist
Load_Item
Load_Sell
Load_Buy
Load_Monster
Load_Attack
Load_SkillName
Load_Kafra

'maproute folder
Load_NPCWARP
Load_NPC

'profile folder
Load_NPC_Profile
Load_Equip_Profile
Load_SelfSkill_Profile
Load_Recovery_Profile

isWarpAll = False
AlwaySit = False
Dead = False
SWeight1 = False
SWeight2 = False
IsAutoPick = False
IsAutoKill = False
IsAutorest = False
'IsAutoSell = False
IsAutoRedz = False
IsAutoOrange = False
'IsAutoSell2 = False
IsConnected = False
IsSPWait = False
IsSkillUse = False
IsAutoDC = False
UseWeapon = False
IsDamageDC = False
IsUseRange = False
IsWantHeal = False
Automove = False
Autoheal = False
AvoidWarp = False
NomonsWarp = False
AutoAI = False

Load_Option
Load_User
Load_Server

If Not FileExists(MapPath & "\prontera.gat") Then
    MsgBox "You need to properly config 'gatpath' in options.ini"
    End
End If
'test_maproute
'Exit Sub
'end test

If Not isUseHaunted Then
    If Not IsConnected Or MasterSelect.Name = "" Then
        frmMasterServer.Visible = True
    Else
        If MasterSelect.IP = "" Then
            MsgBox "Error!!! Need to config your master server in 'table/server.txt' ", vbCritical
            Unload MDIfrmMain
        End If
        frmMain.Visible = True
        frmMain.Main_Init
        MDIfrmMain.mnuReset.Visible = True
        Dead = False
    End If
Else
    frmMain.Visible = True
    frmMain.Main_Init
    MDIfrmMain.mnuReset.Visible = True
    Dead = False
End If
StartZeny = 0
MDIfrmMain.StatusBar1.SimpleText = "Session Time : 00:00:00, Session EXP/JXP : 0/0, Session ZENY : 0"
LoadFormPos MDIfrmMain
If Error Then Unload MDIfrmMain
Exit Sub
errie:
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Msg As Long
    If MDIfrmMain.Visible = False Then
        Msg = X
        
        If Int(Msg / 15) = WM_LBUTTONDBLCLK Then
            Me.WindowState = 2
            Call Shell_NotifyIcon(NIM_DELETE, IconData)
            Me.Show
        End If
        If Int(Msg / 15) = WM_RBUTTONUP Then
            Me.WindowState = 2
            Call Shell_NotifyIcon(NIM_DELETE, IconData)
            Me.Show
        End If
    End If
End Sub

Private Sub MDIForm_Resize()
    If Me.WindowState = 1 Then
        Call Shell_NotifyIcon(NIM_SETVERSION, IconData)
        Call Shell_NotifyIcon(NIM_ADD, IconData)
        Me.Hide
        Exit Sub
    End If
    If Me.width > 1200 Then Frame1.width = Me.width - 120
    SaveFormPos Me
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Call Shell_NotifyIcon(NIM_DELETE, IconData)
'Unload encryption
SaveFormPos Me
End
Unload frmItem
Unload MDIfrmMain
End Sub

Public Sub CreatIcon(tstr As String)
With IconData
.cbSize = Len(IconData)
.hIcon = MDIfrmMain.Icon
.hWnd = MDIfrmMain.hWnd
.szTip = tstr & Chr(0)
.uCallbackMessage = WM_MOUSEMOVE
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE Or NIF_INFO Or NIF_STATE
.uID = vbNull
.uTimeoutOrVersion = NOTIFYICON_VERSION
.szInfoTitle = tstr & Chr(0)
.dwInfoFlags = NIF_INFO
End With
End Sub

Private Function MakeCoords(rawCoords As String) As Coord
On Error GoTo Out
Dim xint As Long
Dim yint As Long
yint = Asc(Mid(rawCoords, 1, 1)) * 4
yint = yint + (Asc(Mid(rawCoords, 2, 1)) And &HC0) / 64
xint = (Asc(Mid(rawCoords, 2, 1)) And &H3F) * 16
xint = xint + (Asc(Mid(rawCoords, 3, 1)) And &HF0) / 16
MakeCoords.Y = yint
MakeCoords.X = xint
Exit Function
Out:
MakeCoords.Y = 0
MakeCoords.X = 0
End Function

Private Sub Load_User()
On Error GoTo errie
Open App.Path & "\table\user.txt" For Input As #1
Dim Index As Integer
Dim tstr As String
Dim Options As String
Dim serv  As Integer
Dim char As Integer
serv = 50
char = 10
Do While Not EOF(1)
    Line Input #1, tstr
    Index = InStr(tstr, "=")
    If Index = 0 Then GoTo endloop
    Options = LCase(Trim(Left(tstr, Index - 1)))
    If Options = "autoconnect" Then
        If Trim(Right(tstr, Len(tstr) - Index)) = "1" Then
            IsConnected = True
        Else
            IsConnected = False
        End If
    ElseIf Options = "master_server_name" Then
        MasterSelect.Name = Trim(Right(tstr, Len(tstr) - Index))
    ElseIf Options = "id" Then
        frmLogin.txtUser.text = DeUser(Trim(Right(tstr, Len(tstr) - Index)), 1)
        strUser = frmLogin.txtUser.text
    ElseIf Options = "pass" Then
        frmLogin.txtPass.text = DeUser(Trim(Right(tstr, Len(tstr) - Index)), 2)
        StrPass = frmLogin.txtPass.text
    ElseIf Options = "server" Then
        If Trim(Right(tstr, Len(tstr) - Index)) <> "" Then serv = Val(Right(tstr, Len(tstr) - Index))
    ElseIf Options = "character" Then
        If Trim(Right(tstr, Len(tstr) - Index)) <> "" Then char = Val(Right(tstr, Len(tstr) - Index))
    End If
    If serv <> 50 And char <> 10 Then
        CharIdStart = char
        NumServ = serv
    End If
endloop:
Loop
Close (1)
Exit Sub
errie:
Close (1)
MsgBox Err.Description, vbCritical
frmLogin.txtUser.text = "user"
frmLogin.txtPass.text = "pass"
End Sub

Public Sub Save_User()
Open App.Path & "\table\user.txt" For Output As #1
Print #1, "[user]"
Print #1, "autoconnect = " & btol(IsConnected)
Print #1, "id = " & EnUser(strUser, 1)
Print #1, "pass = " & EnUser(StrPass, 2)
Print #1, "master_server_name = " & MasterSelect.Name
Print #1, "server = " & CStr(NumServ)
Print #1, "character = " & CStr(CharIdStart)
Close 1
End Sub

Private Sub login_Click()
Load_User
frmLogin.Visible = True
End Sub

Public Sub Save_Option()
Open App.Path & "\control\options.ini" For Output As #1

Print #1, "[Map Control Options]"
Print #1, "gatpath = " & MapPath
Print #1, "usewaypoint = " + btol(RandomMove) + " #" + CStr(Movetime) + "#"
Print #1, "savemapname = " & SaveMapName
Print #1, "lockmapname = " & IIf(Len(LockMapName) < 1, "0", LockMapName)
Print #1, "lockmap_x = " & LockXY.X
Print #1, "lockmap_y = " & LockXY.Y
Print #1, "lockmap_randx = " & LockXYRand.X
Print #1, "lockmap_randy = " & LockXYRand.Y
Print #1, "bi-direction_routing = " & btol(BiDirection)
Print #1, "use_kafra_warp = " & btol(UseNPCWarp)
Print #1, "npc_bi-direction_routing = " & btol(UseNPCBiDirect)
Print #1, "map_routing_time = " & Val(map_time_limit)
Print #1, "forcebuy = " & btol(ForceBuy)
Print #1, ""

Print #1, "[Startup Control Options]"
Print #1, "alwaysit = " & btol(AlwaySit) & " #" & ChatRoomName & "#"
Print #1, "autoshare = " & btol(AutoShare)
Print #1, "exall = " + btol(ExAll)
Print #1, "autorestart = " & btol(UseRestart) & " #" & CStr(RestartTime) & "#"
Print #1, "priority = " & StartPriority
Print #1, "#* note: priority have 4 values possible (realtime, high, normal, low). default is 'normal'."
Print #1, ""

Print #1, "[AI Control Options]"
Print #1, "autoai = " & btol(AutoAI)
Print #1, "nomonssit = " & btol(IsNomonsSit)
Print #1, "autokill = " + btol(IsAutoKill)
Print #1, "autoskill = " + btol(IsSkillUse)
Print #1, "autopick = " + btol(IsAutoPick)
Print #1, "automove = " + btol(Automove)
'Print #1, "wantheal = " + btol(IsWantHeal) & " #" & AcoHealName & "#"
'Print #1, "wantagi = " + btol(IsWantAgi) & " #" + CStr(WantAgiTime) + "#"
'Print #1, "wantbles = " + btol(IsWantBles) & " #" + CStr(WantBlesTime) + "#"
Print #1, "deadrecon = " & btol(DeadRecon)
Print #1, ""


Print #1, "[Party Mode]"
'Print #1, "follow_enable = " & btol(FollowMode.Active)
Print #1, "follow_target = " & FollowMode.Name
Print #1, "follow_autobuff = " & btol(FollowMode.AutoBuff)
'Print #1, "follow_noattack = " & btol(FollowMode.NoAttack)
Print #1, ""

'Print #1, "[Tanker Mode]"
'Print #1, "tanker_enable = " & btol(TankerMode.Active)
'Print #1, "tanker_target = " & btol(TankerMode.Name)
'Print #1, "tanker_autokill = " & btol(TankerMode.AutoBuff)

Print #1, "[Skill Use Control]"
Print #1, "autoheal = " + btol(Autoheal) + " #" + CStr(HPHeal * 100) + "%#" + CStr(HealLV) + "#"
Print #1, "useskillmobs = " + btol(UseSkillMobs) + " #" + MobSkill.rawname + " - " + CStr(MobSkill.Lv) + "#" + _
MobSkill.MonsName; " - " + CStr(MobSkill.number) + "#"
Print #1, "useskillmobs2 = " + btol(UseSkillMobs2) + " #" + MobSkill2.rawname + " - " + CStr(MobSkill2.Lv) + "#" + _
MobSkill2.MonsName; " - " + CStr(MobSkill2.number) + "#"
Print #1, "warpall = " & btol(isWarpAll)
Print #1, ""

Print #1, "[Monk]"
Print #1, "autospirits = " + btol(IsAutoSpirits) + " #" + CStr(SpSpirits * 100) + "%#" + CStr(BallSpirits) + "#"
Print #1, "autochaincombo = " & btol(CCSkill.Use) & " #" & CStr(Val(CCSkill.SP) * 100) & "%#" & CCSkill.Monster & " - " & CStr(CCSkill.Lv) & "#"
Print #1, "autofinishcombo = " & btol(FCSkill.Use) & " #" & CStr(Val(FCSkill.SP) * 100) & "%#" & FCSkill.Monster & " - " & CStr(FCSkill.Lv) & "#"
Print #1, ""

Print #1, "[Sage]"
Print #1, "autospell = " & btol(UseAutoSpell) & " #" & AutoSpell_Name & "#"
Print #1, ""

Print #1, "[HP/SP Options]"
Print #1, "hpsit = " + btol(IsAutorest) + " #" + CStr(HPSit * 100) + "%#"
Print #1, "hpwait = " + btol(IsHPWait) + " #" + CStr(HPWait * 100) + "%#"
Print #1, "spsit = " + btol(IsSPSit) + " #" + CStr(SPSit * 100) + "%#"
Print #1, "spwait = " + btol(IsSPWait) + " #" + CStr(SPWait * 100) + "%#"
Print #1, ""

Print #1, "[Item Use Control]"
Print #1, "autoitem = " & btol(AutoItem.Auto) & " #" & AutoItem.Name & "#" & CStr(AutoItem.Time) & "#"
'Print #1, "autowing = " + btol(AutoWing)
Print #1, "autored = " + btol(IsAutoRedz) + " #" + CStr(HPRed * 100) + "%#" + healitem1 + "#"
Print #1, "autoorange = " + btol(IsAutoOrange) + " #" + CStr(HPOrange * 100) + "%#" + healitem2 + "#"
With SPItem
Print #1, "autosp = " & btol(.Use) & " #" + CStr(.percent * 100) & "%#" & .Name & "#"
End With
Print #1, ""

Print #1, "[Teleport/Disconnect Control Options]"
Print #1, "do-nothing_teleport = " & btol(TeleNothing)
Print #1, "forceteleport = " & btol(ForceTeleport)
Print #1, "nomonswarp = " + btol(NomonsWarp) + " #" + CStr(NomonsTime) + "#"
Print #1, "autodc = " + btol(IsAutoDC) + " #" + CStr(HPDC * 100) + "%#" + CStr(AutoDCCase) + "#"
Print #1, "autodc2 = " + btol(IsDamageDC) + " #" + CStr(DamageSet) + "#" + CStr(AutoDC2Case) + "#"
Print #1, "telemob = " & MobTeleNum
Print #1, ""

'Print #1, "[Pet Control Options]"
'Print #1, "autofeed = " & btol(MyPet.AutoFeed) & " #" & CStr(MyPet.FeedLimit) & "#" & CStr(MyPet.Delay) & "#"
'Print #1, ""

Print #1, "[Avoid Control]"
Print #1, "warpjob = " + btol(JTele) + " #" + JobTele + "#"
Print #1, "avoidwarp = " + btol(AvoidWarp)
Print #1, "avoidgroundskillonposition = " + btol(GSonyou)
Print #1, "avoidgroundskillnearposition = " + btol(GSnearyou)
Print #1, "avoidmonstergroundskillonposition = " + btol(MGSonyou)
Print #1, "avoidmonstergroundskillnearposition = " + btol(MGSnearyou)
Print #1, ""

Print #1, "[Attack Control]"
Print #1, "killmob = " + btol(isKillmob)
Print #1, "onlyinlock = " & btol(AtkMode)
Print #1, "detectskillfail = " & btol(DetectFail)
Print #1, "killsteal = " & btol(killsteal)
Print #1, ""

Print #1, "[Long Range Control Options]"
Print #1, "useweapon = " + btol(UseWeapon)
Print #1, "userange = " + btol(IsUseRange) + " #" + CStr(RangeSet) + "#"
Print #1, "usemindistance = " + btol(useMinDistance) + " #" + CStr(MinDistance) + "#"
Print #1, ""

Print #1, "[Weight Control Options]"
Print #1, "stopattack = " + btol(SWeight1) + " #" + CStr(Weight1 * 100) + "%#"
Print #1, "stoppick = " + btol(SWeight2) + " #" + CStr(Weight2 * 100) + "%#"
Print #1, "backtown = " + btol(IsBackTown) + " #" + CStr(WeightBackTown * 100) + "%#"
Print #1, "backbuy = " & btol(isBackBuy)
Print #1, "backstorage = " & btol(isBackStore)
Print #1, ""

Print #1, "[Log Control]"
Print #1, "chatlog = " & btol(Auto_Chatlog)
'Print #1, "statlog = " & btol(UseStatLog)
Print #1, ""

Print #1, "[Form Control]"
Print #1, "disable_updatepeople = " & btol(Disable_frmPeople)
'Print #1, "statlog = " & btol(UseStatLog)
Print #1, ""

Print #1, "[Timing Control]"
Print #1, "giveuptime = " + CStr(giveuptime)
Print #1, "delay = " + CStr(DelayTime)
Print #1, "warpdelay = " + CStr(WarpDelay)
Print #1, "responsetime = " + CStr(ResponseTime / 60)
Print #1, ""

Print #1, "[Connection Control]"
Print #1, "useproxy = " & btol(IsUseProxy)
Print #1, "proxy_ip = " & ProxyIP
Print #1, "proxy_port = " & CStr(ProxyPort)
Print #1, "proxy_type = " & CStr(ProxyType)
Print #1, "proxy_user = " & ProxyUser
Print #1, "proxy_pass = " & ProxyPass
Print #1, "' proxy_type : 0 - HTTPS / 1 - SOCKS 4 / 2 - SOCKS 5"
Print #1, ""

Close 1
End Sub

Private Sub mnRecShop_Click()
        CreateShop
End Sub

Private Sub mnu_Click()
    frmModConfig.Visible = True
End Sub

Private Sub mnuAll_Click()
Load_NPC_Profile
Load_NPCWARP
Load_SkillName
Load_Special_Status
Load_Char_Status
Load_Emotion
Load_Equip_Profile
Load_SelfSkill_Profile
Load_Recovery_Profile
Load_Monswarplist
Load_Warplist
Load_Avoidlist
Load_Rarelist
Load_Droplist
Load_Monster
Load_Attack
Load_Item
Load_Sell
Load_Buy
Load_Kafra
ReadModOption
MDIfrmMain.Load_Option
End Sub

Private Sub mnuAllProfile_Click()
Load_NPC_Profile
Load_Equip_Profile
Load_SelfSkill_Profile
ReadEventList
Load_Recovery_Profile
End Sub

Private Sub mnuArmor_Click()
frmArmor.Visible = Not frmArmor.Visible
End Sub

Private Sub mnuAttack_Click()
    Load_Attack
End Sub


Private Sub mnuAutoDC_Click()
IsAutoDC = Not IsAutoDC
mnuAutoDC.CheckED = IsAutoDC
Save_Option
End Sub

Private Sub mnuAvoid_Click()
    Load_Avoidlist
End Sub

Private Sub mnuBackTown_Click()
    IsBackTown = Not IsBackTown
    mnuBackTown.CheckED = IsBackTown
End Sub

Private Sub mnuBuy_Click()
Load_Buy
End Sub

Private Sub mnuCast_Click()
mnuCast.CheckED = Not mnuCast.CheckED
End Sub

Private Sub mnuChatLog_Click()
mnuChatLog.CheckED = Not mnuChatLog.CheckED
Auto_Chatlog = mnuChatLog.CheckED
Save_Option
End Sub

Private Sub mnuChatShop_Click()
    frmChatRoom.Visible = True
End Sub

Private Sub mnuCheatAggro_Click()
    mnuCheatAggro.CheckED = Not mnuCheatAggro.CheckED
End Sub

Private Sub mnuDamage_Click()
IsDamageDC = Not IsDamageDC
mnuDamage.CheckED = IsDamageDC
Save_Option
End Sub

Private Sub mnuDrop_Click()
Load_Droplist
End Sub

Private Sub mnuEqmons_Click()
Load_Equip_Profile
End Sub

Private Sub mnuGuild_Click()
    frmGuild.Visible = Not frmGuild.Visible
End Sub

Private Sub mnuHeal_Click()
Autoheal = Not Autoheal
mnuHeal.CheckED = Autoheal
Save_Option
End Sub

Private Sub mnuHPWait_Click()
IsHPWait = Not IsHPWait
mnuHPWait.CheckED = IsHPWait
Save_Option
End Sub

Private Sub mnuKeep_Click()
Load_Kafra
End Sub

Private Sub mnuKillSteal_Click()
killsteal = Not killsteal
mnuKillSteal.CheckED = killsteal
Save_Option
End Sub

Private Sub mnuMap_Click()
    FrmField.Visible = True
End Sub

Private Sub mnuMonster_Click()
frmMonster.Visible = Not frmMonster.Visible
End Sub

Private Sub mnuNPC_Click()
frmNPC.Visible = Not frmNPC.Visible
End Sub

Private Sub mnuNPCReload_Click()
    Load_NPC_Profile
End Sub

Private Sub mnuOptMonsAtk_Click()
    frmConfAttack.Visible = True
End Sub

Private Sub mnuPeople_Click()
frmPeople.Visible = Not frmPeople.Visible
End Sub

Private Sub mnuPKTLOG_Click()
    mnuPKTLOG.CheckED = Not mnuPKTLOG.CheckED
End Sub

'Public Const REALTIME_PRIORITY_CLASS = &H100
'Public Const HIGH_PRIORITY_CLASS = &H80
'Public Const ABOVE_PRIORITY_CLASS = &H60
'Public Const IDLE_PRIORITY_CLASS = &H40
'Public Const NORMAL_PRIORITY_CLASS = &H20
'Public Const LOW_PRIORITY_CLASS = 0
'    SetPriorityClass GetCurrentProcess, HIGH_PRIORITY_CLASS
Private Sub mnuPri_Click(Index As Integer)
    Dim i&
    For i = 0 To 5
        mnuPri(i).CheckED = False
    Next
    mnuPri(Index).CheckED = True
    Select Case Index
        Case 0: SetPriorityClass GetCurrentProcess, REALTIME_PRIORITY_CLASS
        Case 1: SetPriorityClass GetCurrentProcess, HIGH_PRIORITY_CLASS
        Case 2: SetPriorityClass GetCurrentProcess, ABOVE_PRIORITY_CLASS
        Case 3: SetPriorityClass GetCurrentProcess, NORMAL_PRIORITY_CLASS
        Case 4: SetPriorityClass GetCurrentProcess, LOW_PRIORITY_CLASS
        Case 5: SetPriorityClass GetCurrentProcess, IDLE_PRIORITY_CLASS
    End Select
End Sub

Private Sub mnuRandomMove_Click()
RandomMove = Not RandomMove
mnuRandomMove.CheckED = RandomMove
Save_Option
End Sub

Private Sub mnuRarelist_Click()
Load_Rarelist
End Sub

Private Sub mnuMove_Click()
Automove = Not Automove
mnuMove.CheckED = Automove
Save_Option
End Sub

Private Sub mnuNomons_Click()
NomonsWarp = Not NomonsWarp
mnuNomons.CheckED = NomonsWarp
Save_Option
End Sub

Private Sub mnuOptions_Click()
Load_Option
End Sub

Private Sub mnuOrange_Click()
IsAutoOrange = Not IsAutoOrange
mnuOrange.CheckED = IsAutoOrange
Save_Option
End Sub

Private Sub mnuRange_Click()
IsUseRange = Not IsUseRange
mnuRange.CheckED = IsUseRange
Save_Option
End Sub

Private Sub mnuRChatResp_Click()
    ReadEventList
End Sub

Private Sub mnuRecon_Click()
frmMain.ResettoReCon
End Sub

Private Sub mnuRecovery_Click()
Load_Recovery_Profile
End Sub

Private Sub mnuRedz_Click()
IsAutoRedz = Not IsAutoRedz
mnuRedz.CheckED = IsAutoRedz
Save_Option
End Sub

Private Sub mnuAutoKill_Click()
IsAutoKill = Not IsAutoKill
mnuAutoKill.CheckED = IsAutoKill
Save_Option
End Sub

Private Sub mnuAutoPick_Click()
IsAutoPick = Not IsAutoPick
mnuAutoPick.CheckED = IsAutoPick
Save_Option
End Sub

'Private Sub mnuAutoSell1_Click()
'IsAutoSell = Not IsAutoSell
'mnuAutoSell1.CheckED = IsAutoSell
'Save_Option
'End Sub

'Private Sub mnuAutoSell2_Click()
'IsAutoSell2 = Not IsAutoSell2
'mnuAutoSell2.CheckED = IsAutoSell2
'Save_Option
'End Sub

Private Sub mnuAutosit_Click()
IsAutorest = Not IsAutorest
mnuAutosit.CheckED = IsAutorest
Save_Option
End Sub

Private Sub mnuAutoSkill_Click()
IsSkillUse = Not IsSkillUse
mnuAutoSkill.CheckED = IsSkillUse
Save_Option
End Sub

Private Sub mnuCas_Click()
Me.Arrange vbCascade
End Sub

Private Sub mnuChat_Click()
frmChat.Visible = Not frmChat.Visible
End Sub

Private Sub mnuInv_Click()
frmItem.Visible = Not frmItem.Visible
End Sub

Private Sub mnuMain_Click()
frmMain.Visible = Not frmMain.Visible
End Sub

Private Sub mnuOption_Click()
'frmOption.Visible = Not frmOption.Visible
End Sub

Private Sub mnuPlayer_Click()
frmPlayer.Visible = Not frmPlayer.Visible
End Sub


Private Sub mnuReset_Click()
frmMain.tmrAnswer_Timer
End Sub



Private Sub mnuRMod_Click()
    ReadModOption
End Sub

Private Sub mnuSAttack_Click()
    SWeight1 = Not SWeight1
    mnuSAttack = SWeight1
    Save_Option
End Sub

Private Sub mnuSave_Click()
frmMain.Warp_Save "System: Manual Warp to Save Point..."
End Sub



Private Sub mnuSelfSkill_Click()
Load_SelfSkill_Profile
End Sub

Private Sub mnuSell_Click()
Load_Sell
End Sub

Private Sub mnuSendRaw_Click()
frmSendRaw.Visible = True
End Sub

Private Sub mnuServer_Click()
    frmMasterServer.Visible = True
End Sub

Private Sub mnuShare_Click()
    frmMain.Set_Share
End Sub

Private Sub mnuSit_Click()
    frmMain.Send_Sit
End Sub

Private Sub mnuSkill_Click()
frmSkill.Visible = Not frmSkill.Visible
End Sub

Private Sub mnuSkillMobs_Click()
    UseSkillMobs = Not UseSkillMobs
    mnuSkillMobs.CheckED = UseSkillMobs
    Save_Option
End Sub

Private Sub mnuSkillMobs2_Click()
    UseSkillMobs2 = Not UseSkillMobs2
    mnuSkillMobs2.CheckED = UseSkillMobs2
    Save_Option
End Sub

Private Sub mnuSpecialStatus_Click()
Load_Special_Status
End Sub

Private Sub mnuSPick_Click()
    SWeight2 = Not SWeight2
    mnuSPick = SWeight2
    Save_Option
End Sub

Private Sub mnuSPSit_Click()
IsSPSit = Not IsSPSit
mnuSPSit.CheckED = IsSPSit
Save_Option
End Sub

Private Sub mnuSPWait_Click()
IsSPWait = Not IsSPWait
mnuSPWait.CheckED = IsSPWait
Save_Option
End Sub

Private Sub mnuStand_Click()
    frmMain.Send_Stand
End Sub

Private Sub mnuStat_Click()
frmStat.Visible = Not frmStat.Visible
End Sub



Private Sub mnuStatus_Click()
    frmStatus.Visible = Not frmStatus.Visible
End Sub

Private Sub mnuStatustxt_Click()
Load_Char_Status
End Sub

Private Sub mnuTeleport_Click()
 Teleport
End Sub

Private Sub mnuUnshare_Click()
frmMain.unSet_Share
End Sub

Private Sub mnuUseOrange_Click()
frmMain.UseOrangez
End Sub

Private Sub mnuUseRed_Click()
frmMain.UseRedz
End Sub

Private Sub mnuWarp_Click()
AvoidWarp = Not AvoidWarp
mnuWarp.CheckED = AvoidWarp
Save_Option
End Sub

Private Sub mnuWeapon_Click()
UseWeapon = Not UseWeapon
mnuWeapon.CheckED = UseWeapon
Save_Option
End Sub

Private Sub Save_Click()
frmMain.Warp_Save "System: Manual Warp to Save Point..."
End Sub

Public Sub Load_Option()
On Error GoTo Out:
Open App.Path & "\control\options.ini" For Input As #1
Dim text As String
Dim test As String
Dim tstr As String
Dim Index As Integer
Dim index2 As Integer
Dim index3 As Integer
Dim TmpMons As String
'killsteal = False
map_time_limit = 10
text = "="

Do While Not EOF(1)
    Input #1, tstr
    text = "="
    Index = InStr(1, tstr, text, vbTextCompare) - 1
    If Index > 0 Then
        If LCase(Trim(Left(tstr, Index))) = "onlyinlock" Then
            AtkMode = CBool(Trim(Right(tstr, Len(tstr) - Index - 1)))

        ElseIf LCase(Trim(Left(tstr, Index))) = "deadrecon" Then
            DeadRecon = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
        
        ElseIf LCase(Trim(Left(tstr, Index))) = "disable_updatepeople" Then
            Disable_frmPeople = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
        
        ElseIf LCase(Trim(Left(tstr, Index))) = "follow_enable" Then
            FollowMode.Active = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
        
        ElseIf LCase(Trim(Left(tstr, Index))) = "follow_autobuff" Then
            FollowMode.AutoBuff = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
        
        ElseIf LCase(Trim(Left(tstr, Index))) = "follow_noattack" Then
            FollowMode.NoAttack = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
        
        ElseIf LCase(Trim(Left(tstr, Index))) = "follow_target" Then
            FollowMode.Name = Trim(Right(tstr, Len(tstr) - Index - 1))
        
        ElseIf LCase(Trim(Left(tstr, Index))) = "priority" Then
            StartPriority = LCase(Trim(Right(tstr, Len(tstr) - Index - 1)))
            Select Case StartPriority
                Case "high"
                    mnuPri_Click 1
                Case "low"
                    mnuPri_Click 5
                Case "realtime"
                    mnuPri_Click 0
                Case Else
                    mnuPri_Click 3
                    StartPriority = "normal"
            End Select

        ElseIf LCase(Trim(Left(tstr, Index))) = "do-nothing_teleport" Then
            TeleNothing = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
            
        ElseIf LCase(Trim(Left(tstr, Index))) = "useproxy" Then
            IsUseProxy = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
            
        ElseIf LCase(Trim(Left(tstr, Index))) = "proxy_ip" Then
            ProxyIP = Trim(Right(tstr, Len(tstr) - Index - 1))
            
        ElseIf LCase(Trim(Left(tstr, Index))) = "proxy_user" Then
            ProxyUser = Trim(Right(tstr, Len(tstr) - Index - 1))
            
        ElseIf LCase(Trim(Left(tstr, Index))) = "proxy_pass" Then
            ProxyPass = Trim(Right(tstr, Len(tstr) - Index - 1))
            
        ElseIf LCase(Trim(Left(tstr, Index))) = "proxy_port" Then
            ProxyPort = Val(Trim(Right(tstr, Len(tstr) - Index - 1)))
            
        ElseIf LCase(Trim(Left(tstr, Index))) = "proxy_type" Then
            ProxyType = Val(Trim(Right(tstr, Len(tstr) - Index - 1)))
            
        ElseIf LCase(Trim(Left(tstr, Index))) = "onlyskill" Then
            SkillOnly = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))

        ElseIf LCase(Trim(Left(tstr, Index))) = "forceteleport" Then
            ForceTeleport = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))

        ElseIf LCase(Trim(Left(tstr, Index))) = "autorestart" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(tstr, "#") - 1
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then UseRestart = True Else UseRestart = False
                If Right(tstr, 1) = "#" Then
                    RestartTime = Val(Mid(tstr, index2 + 2, Len(tstr) - index2 - 2))
                Else
                    RestartTime = Val(Right(tstr, Len(tstr) - index2 - 1))
                End If
            End If

        ElseIf LCase(Trim(Left(tstr, Index))) = "telemob" Then
            MobTeleNum = Val(Trim(Right(tstr, Len(tstr) - Index - 1)))
            
        ElseIf LCase(Trim(Left(tstr, Index))) = "autokill" Then
            IsAutoKill = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
            mnuAutoKill.CheckED = IsAutoKill
        
        ElseIf LCase(Trim(Left(tstr, Index))) = "warpall" Then
                isWarpAll = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))

        ElseIf LCase(Trim(Left(tstr, Index))) = "use_kafra_warp" Then
                UseNPCWarp = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
                
        ElseIf LCase(Trim(Left(tstr, Index))) = "bi-direction_routing" Then
                BiDirection = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
                
        ElseIf LCase(Trim(Left(tstr, Index))) = "npc_bi-direction_routing" Then
                UseNPCBiDirect = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
                
        ElseIf LCase(Trim(Left(tstr, Index))) = "map_routing_time" Then
            map_time_limit = Val(Trim(Right(tstr, Len(tstr) - Index - 1)))
            If map_time_limit < 10 Then map_time_limit = 10

        ElseIf LCase(Trim(Left(tstr, Index))) = "backbuy" Then
                isBackBuy = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
                
        ElseIf LCase(Trim(Left(tstr, Index))) = "backstorage" Then
                isBackStore = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
        
        ElseIf LCase(Trim(Left(tstr, Index))) = "killmob" Then
                isKillmob = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))

        ElseIf LCase(Trim(Left(tstr, Index))) = "detectskillfail" Then
                DetectFail = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))

        'ElseIf LCase(Trim(Left(tstr, index))) = "autowing" Then
        '        AutoWing = CBool(Val(Trim(Right(tstr, Len(tstr) - index - 1))))
        
        ElseIf LCase(Trim(Left(tstr, Index))) = "autoshare" Then
                AutoShare = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))

        ElseIf LCase(Trim(Left(tstr, Index))) = "chatlog" Then
                Auto_Chatlog = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
                mnuChatLog.CheckED = Auto_Chatlog
                If Auto_Chatlog Then
                Open ChatlogPath For Append As #5
                    Print #5, "======================================================="
                    Print #5, "========= Session started at " & Date & " : " & Time & "========="
                    Print #5, "======================================================="
                    Print #5, ""
                Close #5
                End If
        
        'ElseIf LCase(Trim(Left(tstr, Index))) = "statlog" Then
        '        If UseStatLog Then Close #100
        '        UseStatLog = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
        '        If UseStatLog Then
        '            On Error Resume Next
        '            Open StatlogPath For Binary Access Write As #100
        '            Put #100, LOF(100), CStr("================================================" & Chr(13) & Chr(10) & Chr(10))
        '            Put #100, LOF(100), CStr("========== Session started at " & Date & " : " & Time & " ==========" & Chr(13) & Chr(10) & Chr(10))
        '            Put #100, LOF(100), CStr("================================================" & Chr(13) & Chr(10) & Chr(10))
        '            Put #100, LOF(100), CStr(Chr(13) & Chr(10) & Chr(13) & Chr(10))
        '            Err.Clear
        '            On Error GoTo Out
        '        End If

        ElseIf LCase(Trim(Left(tstr, Index))) = "nomonssit" Then
                IsNomonsSit = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))

        ElseIf LCase(Trim(Left(tstr, Index))) = "alwaysit" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(tstr, "#") - 1
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then AlwaySit = True Else AlwaySit = False
                If Right(tstr, 1) = "#" Then
                    ChatRoomName = Mid(tstr, index2 + 2, Len(tstr) - index2 - 2)
                Else
                    ChatRoomName = Right(tstr, Len(tstr) - index2 - 1)
                End If
            End If
        ElseIf LCase(Trim(Left(tstr, Index))) = "autoai" Then
            AutoAI = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
            FrmField.update_ImgAI

'-------------------------------------------------------------------------
        ElseIf LCase(Trim(Left(tstr, Index))) = "autospell" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    index2 = index2 + 2
                    UseAutoSpell = True
                    AutoSpell_Name = Mid(tstr, index2, Len(tstr) - index2)
                End If
            End If
'--------------------------------------------------------------------------
        ElseIf LCase(Trim(Left(tstr, Index))) = "gatpath" Then
            MapPath = Trim(Right(tstr, Len(tstr) - Index - 1))
            If MapPath = "" Then
                MsgBox "You need to set gatpath = x:\...\gat\!   ", vbCritical
                Error = True
                Exit Sub
            End If
            If Right(MapPath, 1) = "\" Then MapPath = Left(MapPath, Len(MapPath) - 1)

        ElseIf LCase(Trim(Left(tstr, Index))) = "delay" Then
            DelayTime = Val(Trim(Right(tstr, Len(tstr) - Index - 1)))
            If (DelayTime < 5) Then
                DelayTime = 5
                MsgBox "To avoid DoS, Delay Time back to default (5s.)", vbOKOnly, "Delay Time < 5 seconds!"
            End If

        ElseIf LCase(Trim(Left(tstr, Index))) = "warpdelay" Then
             WarpDelay = Val(Trim(Right(tstr, Len(tstr) - Index - 1)))
             
        ElseIf LCase(Trim(Left(tstr, Index))) = "responsetime" Then
             ResponseTime = Val(Trim(Right(tstr, Len(tstr) - Index - 1))) * 60
             
        ElseIf LCase(Trim(Left(tstr, Index))) = "useweapon" Then
            UseWeapon = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
            
        ElseIf LCase(Trim(Left(tstr, Index))) = "autopick" Then
                IsAutoPick = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))

        ElseIf LCase(Trim(Left(tstr, Index))) = "autoskill" Then
                IsSkillUse = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))

        ElseIf LCase(Trim(Left(tstr, Index))) = "avoidwarp" Then
                AvoidWarp = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))

        'ElseIf LCase(Trim(Left(tstr, index))) = "autosell" Then
        '        IsAutoSell = CBool(Val(Trim(Right(tstr, Len(tstr) - index - 1))))
        
        ElseIf LCase(Trim(Left(tstr, Index))) = "killsteal" Then
                killsteal = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
                'if curEnableKey Then MDIfrmMain.mnuRegis.CheckED = True

        'ElseIf LCase(Trim(Left(tstr, index))) = "autosell2" Then
        '        IsAutoSell2 = CBool(Val(Trim(Right(tstr, Len(tstr) - index - 1))))

        ElseIf LCase(Trim(Left(tstr, Index))) = "automove" Then
                Automove = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
                
        ElseIf LCase(Trim(Left(tstr, Index))) = "hpsit" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    IsAutorest = True
                    'mnuAutosit.Checked = IsAutorest
                End If
                test = Mid(tstr, index2 + 2, Len(tstr) - index2 - 3)
                HPSit = Val(Mid(tstr, index2 + 2, Len(tstr) - index2 - 3)) / 100
                'mnuAutosit.Caption = "Auto sit when HP below " + Mid(tstr, index2 + 2, Len(tstr) - index2 - 3) + "%."
            End If
        ElseIf LCase(Trim(Left(tstr, Index))) = "autoitem" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    AutoItem.Auto = True
                Else
                    AutoItem.Auto = False
                End If
                tstr = Right(tstr, Len(tstr) - index2 - 1)
                Index = InStr(tstr, "#")
                If Index > 0 Then
                    AutoItem.Name = Left(tstr, Index - 1)
                    tstr = Right(tstr, Len(tstr) - Index)
                    AutoItem.Time = Val(CStr(Left(tstr, Len(tstr) - 1)))
                End If
                If AutoItem.Name = "" Or AutoItem.Time = 0 Then AutoItem.Auto = False
            End If
        'ElseIf LCase(Trim(Left(tstr, Index))) = "autofeed" Then
        '    text = "#"
        '    Index = Index + 2
        '    index2 = InStr(1, tstr, text, vbTextCompare) - 1
        '    'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
        '    If (index2 > 0) Then
        '        If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
        '            MyPet.AutoFeed = True
        '        Else
        '            MyPet.AutoFeed = False
        '        End If
        ''        tstr = Right(tstr, Len(tstr) - index2 - 1)
        '        Index = InStr(tstr, "#")
        '        If Index > 0 Then
        '            MyPet.FeedLimit = Val(CStr(Left(tstr, Index - 1)))
        '            tstr = Right(tstr, Len(tstr) - Index)
        '            MyPet.Delay = Val(CStr(Left(tstr, Len(tstr) - 1)))
        '        End If
        '        If MyPet.FeedLimit = 0 Or MyPet.FeedLimit > 100 Then MyPet.FeedLimit = 35
        '        If MyPet.DelayFeed < 9 Then MyPet.DelayFeed = 9
        '    End If
        ElseIf LCase(Trim(Left(tstr, Index))) = "hpwait" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    IsHPWait = True
                    mnuHPWait.CheckED = IsHPWait
                End If
                test = Mid(tstr, index2 + 2, Len(tstr) - index2 - 3)
                HPWait = Val(Mid(tstr, index2 + 2, Len(tstr) - index2 - 3)) / 100
                'mnuHPWait.Caption = "Sit until HP reach " + Mid(tstr, index2 + 2, Len(tstr) - index2 - 3) + "%."
            End If
         ElseIf LCase(Trim(Left(tstr, Index))) = "spsit" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    IsSPSit = True
                    mnuSPSit.CheckED = IsSPSit
                End If
                index2 = index2 + 2
                SPSit = Val(Mid(tstr, index2, Len(tstr) - index2 - 1)) / 100
                'mnuSPSit.Caption = "Auto sit when SP below " + Mid(tstr, index2, Len(tstr) - index2) + "."
            End If
        ElseIf LCase(Trim(Left(tstr, Index))) = "spwait" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    IsSPWait = True
                    mnuSPWait.CheckED = IsSPWait
                End If
                index2 = index2 + 2
                SPWait = Val(Mid(tstr, index2, Len(tstr) - index2 - 1)) / 100
                'mnuSPWait.Caption = "Sit until SP reach " + Mid(tstr, index2, Len(tstr) - index2) + "."
            End If
        ElseIf LCase(Trim(Left(tstr, Index))) = "autosp" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                With SPItem
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    .Use = True
                Else
                    .Use = False
                End If
                index3 = InStr(index2 + 2, tstr, text, vbTextCompare)
                .percent = Val(Mid(tstr, index2 + 2, index3 - index2 - 3)) / 100
                .Name = Mid(tstr, index3 + 1, Len(tstr) - index3 - 1)
                End With
            End If
        ElseIf LCase(Trim(Left(tstr, Index))) = "autored" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    IsAutoRedz = True
                    mnuRedz.CheckED = IsAutoRedz
                End If
                index3 = InStr(index2 + 2, tstr, text, vbTextCompare)
                HPRed = Val(Mid(tstr, index2 + 2, index3 - index2 - 3)) / 100
                healitem1 = Mid(tstr, index3 + 1, Len(tstr) - index3 - 1)
                
            End If
        ElseIf LCase(Trim(Left(tstr, Index))) = "autoheal" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    Autoheal = True
                    mnuHeal.CheckED = Autoheal
                End If
                index3 = InStr(index2 + 2, tstr, text, vbTextCompare)
                HPHeal = Val(Mid(tstr, index2 + 2, index3 - index2 - 3)) / 100
                HealLV = Val(Mid(tstr, index3 + 1, Len(tstr) - index3 - 1))
                mnuHeal.Caption = "Auto heal (LV." + CStr(HealLV) + ") when HP below " + Mid(tstr, index2 + 2, index3 - index2 - 3) + "%."
            End If
        ElseIf LCase(Trim(Left(tstr, Index))) = "autoorange" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    IsAutoOrange = True
                    mnuOrange.CheckED = IsAutoOrange
                End If
                index3 = InStr(index2 + 2, tstr, text, vbTextCompare)
                HPOrange = Val(Mid(tstr, index2 + 2, index3 - index2 - 3)) / 100
                healitem2 = Mid(tstr, index3 + 1, Len(tstr) - index3 - 1)
                'mnuOrange.Caption = "Auto use " + healitem2 + " when HP below " + Mid(tstr, index2 + 2, index3 - index2 - 3) + "%."
                'mnuUseOrange.Caption = "Use " & healitem2
            End If
        
        ElseIf LCase(Trim(Left(tstr, Index))) = "autodc" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then IsAutoDC = True Else IsAutoDC = False
                mnuAutoDC.CheckED = IsAutoDC
                index3 = InStr(index2 + 2, tstr, text, vbTextCompare)
                HPDC = Val(Mid(tstr, index2 + 2, index3 - index2 - 3)) / 100
                AutoDCCase = Val(Mid(tstr, index3 + 1, Len(tstr) - index3 - 1))
            End If
        ElseIf LCase(Trim(Left(tstr, Index))) = "autodc2" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    IsDamageDC = True
                    mnuDamage.CheckED = IsDamageDC
                End If
                index3 = InStr(index2 + 2, tstr, text, vbTextCompare)
                DamageSet = Val(Mid(tstr, index2 + 2, index3 - index2 - 2))
                AutoDC2Case = Val(Mid(tstr, index3 + 1, Len(tstr) - index3 - 1))
                'If (AutoDC2Case = 0) Then
                '    mnuDamage.Caption = "Auto teleport when damage over " + Mid(tstr, index2 + 2, index3 - index2 - 2) + "."
                'Else
                '    mnuDamage.Caption = "Auto DC when damage over " + Mid(tstr, index2 + 2, index3 - index2 - 3) + "."
                'End If
            End If
        ElseIf LCase(Trim(Left(tstr, Index))) = "usemindistance" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    useMinDistance = True
                End If
                index2 = index2 + 2
                MinDistance = Val(Mid(tstr, index2, Len(tstr) - index2))
            End If
        ElseIf LCase(Trim(Left(tstr, Index))) = "userange" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    IsUseRange = True
                    mnuRange.CheckED = IsUseRange
                End If
                index2 = index2 + 2
                RangeSet = Val(Mid(tstr, index2, Len(tstr) - index2))
                'mnuRange.Caption = "Attack if distance below " + Mid(tstr, index2, Len(tstr) - index2) + " blocks."
            End If
        ElseIf LCase(Trim(Left(tstr, Index))) = "nomonswarp" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    NomonsWarp = True
                    mnuNomons.CheckED = NomonsWarp
                End If
                index2 = index2 + 2
                NomonsTime = Val(Mid(tstr, index2, Len(tstr) - index2))
                'mnuNomons.Caption = "Auto teleport when no monster for " + Mid(tstr, index2, Len(tstr) - index2) + " second(s)."
            End If
        ElseIf LCase(Trim(Left(tstr, Index))) = "usewaypoint" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    RandomMove = True
                    mnuRandomMove.CheckED = RandomMove
                End If
                index2 = index2 + 2
                Movetime = Val(Mid(tstr, index2, Len(tstr) - index2))
                'mnuRandomMove.Caption = "Go along waypoint when no monster every " + Mid(tstr, index2, Len(tstr) - index2) + " second(s)."
                
            End If
        ElseIf LCase(Trim(Left(tstr, Index))) = "party" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    index2 = index2 + 2
                    PartyMode = True
                    mnuParty.CheckED = True
                    text = Mid(tstr, index2, Len(tstr) - index2)
                    Tanker.ID = Chr(Val("&H" & Left(text, 2))) & Chr(Val("&H" & Mid(text, 3, 2))) & Chr(Val("&H" & Mid(text, 5, 2))) & Chr(Val("&H" & Right(tstr, 2)))
                    mnuParty.Caption = "Party ('" & MakeHexName(Tanker.ID) & "' is a Tanker)"
                Else
                    PartyMode = False
                    mnuParty.CheckED = False
                End If
                'Debug.Print MakeHexName(TankerID)
            End If
        ElseIf LCase(Trim(Left(tstr, Index))) = "stopattack" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    SWeight1 = True
                    mnuSAttack.CheckED = SWeight1
                End If
                Weight1 = Val(Mid(tstr, index2 + 2, Len(tstr) - index2 - 3)) / 100
                'mnuSAttack.Caption = "Stop attack when weight reach " + Mid(tstr, index2 + 2, Len(tstr) - index2 - 3) + "%."
            End If
        ElseIf LCase(Trim(Left(tstr, Index))) = "stoppick" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    SWeight2 = True
                    mnuSPick.CheckED = SWeight2
                End If
                Weight2 = Val(Mid(tstr, index2 + 2, Len(tstr) - index2 - 3)) / 100
                'mnuSPick.Caption = "Stop pick up when weight reach " + Mid(tstr, index2 + 2, Len(tstr) - index2 - 3) + "%."
            End If
        ElseIf LCase(Trim(Left(tstr, Index))) = "backtown" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    IsBackTown = True
                    mnuBackTown.CheckED = IsBackTown
                Else
                    IsBackTown = False
                End If
                WeightBackTown = Val(Mid(tstr, index2 + 2, Len(tstr) - index2 - 3)) / 100
                'mnuBackTown.Caption = "Back to town for sell when weight reach " + Mid(tstr, index2 + 2, Len(tstr) - index2 - 3) + "%."
            End If
' monk combo skill start
        ElseIf LCase(Trim(Left(tstr, Index))) = "autospirits" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    IsAutoSpirits = True
                End If
                index3 = InStr(index2 + 2, tstr, text, vbTextCompare)
                SpSpirits = Val(Mid(tstr, index2 + 2, index3 - index2 - 3)) / 100
                BallSpirits = Val(Mid(tstr, index3 + 1, Len(tstr) - index3 - 1))
            End If
        ElseIf LCase(Trim(Left(tstr, Index))) = "autochaincombo" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    CCSkill.Use = True
                End If
                index3 = InStr(index2 + 2, tstr, text, vbTextCompare)
                CCSkill.SP = Val(Mid(tstr, index2 + 2, index3 - index2 - 3)) / 100
                TmpMons = Mid(tstr, index3 + 1, Len(tstr) - index3 - 1)
                index3 = InStr(TmpMons, "-")
                CCSkill.Monster = Trim(Left(TmpMons, index3 - 1))
                CCSkill.Lv = Val(Right(TmpMons, Len(TmpMons) - index3 - 1))
                If CCSkill.Lv > 5 Then CCSkill.Lv = 5
            End If
        ElseIf LCase(Trim(Left(tstr, Index))) = "autofinishcombo" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    FCSkill.Use = True
                End If
                index3 = InStr(index2 + 2, tstr, text, vbTextCompare)
                FCSkill.SP = Val(Mid(tstr, index2 + 2, index3 - index2 - 3)) / 100
                TmpMons = Mid(tstr, index3 + 1, Len(tstr) - index3 - 1)
                index3 = InStr(TmpMons, "-")
                FCSkill.Monster = Trim(Left(TmpMons, index3 - 1))
                FCSkill.Lv = Val(Right(TmpMons, Len(TmpMons) - index3 - 1))
                If FCSkill.Lv > 5 Then FCSkill.Lv = 5
            End If
' monk combo skill end
        ElseIf LCase(Trim(Left(tstr, Index))) = "giveuptime" Then
             giveuptime = Val(Trim(Right(tstr, Len(tstr) - Index - 1)))
             If giveuptime < 4 Then giveuptime = 4
        ElseIf LCase(Trim(Left(tstr, Index))) = "useskillmobs2" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    UseSkillMobs2 = True
                Else
                    UseSkillMobs2 = False
                End If
                mnuSkillMobs2.CheckED = UseSkillMobs2
                index3 = InStr(index2 + 2, tstr, text, vbTextCompare)
                MobSkill2.rawname = Mid(tstr, index2 + 2, index3 - index2 - 2)
                TmpMons = Mid(tstr, index3 + 1, Len(tstr) - index3 - 1)
                index3 = InStr(MobSkill2.rawname, "-")
                MobSkill2.Lv = Val(Right(MobSkill2.rawname, Len(MobSkill2.rawname) - index3))
                MobSkill2.rawname = Trim(Left(MobSkill2.rawname, index3 - 1))
                index3 = InStr(TmpMons, "-")
                MobSkill2.MonsName = Trim(Left(TmpMons, index3 - 1))
                MobSkill2.number = Val(Right(TmpMons, Len(TmpMons) - index3))
                index3 = InStr(MobSkill2.MonsName, "&")
            End If
        ElseIf LCase(Trim(Left(tstr, Index))) = "useskillmobs" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    UseSkillMobs = True
                Else
                    UseSkillMobs = False
                End If
                mnuSkillMobs.CheckED = UseSkillMobs
                index3 = InStr(index2 + 2, tstr, text, vbTextCompare)
                MobSkill.rawname = Mid(tstr, index2 + 2, index3 - index2 - 2)
                TmpMons = Mid(tstr, index3 + 1, Len(tstr) - index3 - 1)
                index3 = InStr(MobSkill.rawname, "-")
                MobSkill.Lv = Val(Right(MobSkill.rawname, Len(MobSkill.rawname) - index3))
                MobSkill.rawname = Trim(Left(MobSkill.rawname, index3 - 1))
                index3 = InStr(TmpMons, "-")
                MobSkill.MonsName = Trim(Left(TmpMons, index3 - 1))
                MobSkill.number = Val(Right(TmpMons, Len(TmpMons) - index3))
                index3 = InStr(MobSkill.MonsName, "&")
                'ReDim MobName(0)
                'TmpMons = MobSkill.MonsName
                'While index3 > 0
                '    MobName(UBound(MobName)) = Left(MobSkill.MonsName, index3 - 1)
                '    MobSkill.MonsName = Right(MobSkill.MonsName, Len(MobSkill.MonsName) - index3)
                '    ReDim Preserve MobName(UBound(MobName) + 1)
                '    index3 = InStr(MobSkill.MonsName, "&")
                'Wend
                '    MobName(UBound(MobName)) = MobSkill.MonsName
                '    MobSkill.MonsName = TmpMons
                'mnuSkillMobs.Caption = "Use skill [" & MobSkill.rawname & "] lv." & CStr(MobSkill.lv) & " when " & MobSkill.monsname & " attack you > " & CStr(MobSkill.number)
            End If
        ElseIf LCase(Trim(Left(tstr, Index))) = "savemapname" Then
                SaveMapName = Trim(Right(tstr, Len(tstr) - (InStr(tstr, "=") + 1)))
                'Debug.Print SaveMapName
        ElseIf LCase(Trim(Left(tstr, Index))) = "lockmapname" Then
                LockMapName = Trim(Right(tstr, Len(tstr) - (InStr(tstr, "=") + 1)))
                If LockMapName = "0" Then LockMapName = ""
                'Debug.Print SaveMapName
        ElseIf LCase(Trim(Left(tstr, Index))) = "lockmap_x" Then
                LockXY.X = Val(Trim(Right(tstr, Len(tstr) - (InStr(tstr, "=") + 1))))
                'Debug.Print SaveMapName
        ElseIf LCase(Trim(Left(tstr, Index))) = "lockmap_y" Then
                LockXY.Y = Val(Trim(Right(tstr, Len(tstr) - (InStr(tstr, "=") + 1))))
                'Debug.Print SaveMapName
        ElseIf LCase(Trim(Left(tstr, Index))) = "lockmap_randx" Then
                LockXYRand.X = Val(Trim(Right(tstr, Len(tstr) - (InStr(tstr, "=") + 1))))
                If LockXYRand.X > 0 And LockXYRand.X < 1 Then LockXYRand.X = 2
                'Debug.Print SaveMapName
        ElseIf LCase(Trim(Left(tstr, Index))) = "lockmap_randy" Then
                LockXYRand.Y = Val(Trim(Right(tstr, Len(tstr) - (InStr(tstr, "=") + 1))))
                If LockXYRand.Y > 0 And LockXYRand.Y < 1 Then LockXYRand.Y = 2
                'Debug.Print SaveMapName
        ElseIf LCase(Trim(Left(tstr, Index))) = "forcebuy" Then
                ForceBuy = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
'=====new FT /devil copy
        ElseIf LCase(Trim(Left(tstr, Index))) = "exall" Then
                ExAll = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
                
        ElseIf LCase(Trim(Left(tstr, Index))) = "warpjob" Then
            text = "#"
            Index = Index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            If index2 > 0 Then
                If Trim(Mid(tstr, Index, index2 - Index)) = "1" Then
                    JTele = True
                End If
                index2 = index2 + 2
                JobTele = Mid(tstr, index2, Len(tstr) - index2)
            End If
        ElseIf LCase(Trim(Left(tstr, Index))) = "avoidgroundskillonposition" Then
                GSonyou = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
                
        ElseIf LCase(Trim(Left(tstr, Index))) = "avoidgroundskillnearposition" Then
                GSnearyou = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
                
        ElseIf LCase(Trim(Left(tstr, Index))) = "avoidmonstergroundskillonposition" Then
                MGSonyou = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
                
        ElseIf LCase(Trim(Left(tstr, Index))) = "avoidmonstergroundskillnearposition" Then
                MGSnearyou = CBool(Val(Trim(Right(tstr, Len(tstr) - Index - 1))))
                
'=================caseend
        End If
    End If
Loop
Close 1
If HPRed < HPOrange Then
    Dim tmphealname As String
    Dim tmphealhp As Boolean
    Dim tmpphp As Double
    tmphealname = healitem2
    tmphealhp = IsAutoOrange
    tmpphp = HPOrange
    healitem2 = healitem1
    IsAutoOrange = IsAutoRedz
    HPOrange = HPRed
    healitem1 = tmphealname
    IsAutoRedz = tmphealhp
    HPRed = tmpphp
    'mnuRedz.Caption = "Auto use " + healitem1 + " when HP below " + Mid(tstr, index2 + 2, index3 - index2 - 3) + "%."
End If
mnuUseRed.Caption = "Use " & healitem2
mnuUseOrange.Caption = "Use " & healitem1
Save_Option
Update_Option_Form
Update_AttackOption_Form
Update_HPOption_Form
Update_TeleOption_Form
Exit Sub
Out:
Close 1
MsgBox "Error!!! in 'control\options.ini' on loading [" & LCase(Trim(Left(tstr, Index))) & "] option : " & Err.Description, vbCritical
'Unload MDIfrmMain
'error = True
End Sub
Private Sub mnuTeleportloot_Click()
    frmTeleportOption.Visible = True
    Update_TeleOption_Form
End Sub

Public Sub Update_TeleOption_Form()
    With frmTeleportOption
        .update_imgDamageTele
        .update_imgHpTele
        .update_imgNomons
        '.update_ImgUseWing
        .update_imgJTele
        .update_imgWarpAll
    End With
End Sub

Private Sub mnuAtt_Click()
    frmAttackOption.Visible = True
    Update_AttackOption_Form
End Sub

Public Sub Update_AttackOption_Form()
    With frmAttackOption
        .update_imgAutoMobSkill
        .update_imgAutoRangeAttack
        .update_imgAutoSkill
        '.update_imgKS
        .update_imgAutoAttack
        .update_imgUseWeapon
        .update_imgKillMob
        .update_imgMindistance
    End With
End Sub

Private Sub mnuHPSP_Click()
    FrmHPSPOption.Visible = True
    Update_HPOption_Form
End Sub

Public Sub Update_HPOption_Form()
    With FrmHPSPOption
        .update_ImgAutositHP
        .update_imgSitUntilHP
        .update_imgAutoSitSP
        .update_imgSitUntilSP
        .update_imgUseItem1
        .update_imgUseItem2
        .update_imgSitNomons
    End With
End Sub

Private Sub mnuAI_Click()
    frmAIOption.Visible = True
    Update_Option_Form
End Sub

Public Sub Update_Option_Form()
   With frmAIOption
        .update_imgPick
        .update_imgMove
        .update_imgWayPoint
        .update_imgBackTown
        .update_imgBackBuy
        .update_imgStopPick
        .update_imgStopAttack
        .update_imgAlwaySit
        .update_imgExAll
        .update_imgGSonyou
        .update_imgGSnearyou
        .update_imgMGSonyou
        .update_imgMGSnearyou
   End With
End Sub

Private Sub mnuWinParty_Click()
    frmParty.Visible = Not frmParty.Visible
End Sub

Private Sub mnWCart_Click()
    frmCart.Visible = Not frmCart.Visible
End Sub
