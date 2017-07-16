VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIfrmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H00808080&
   Caption         =   "Revemu 0.87.1"
   ClientHeight    =   7800
   ClientLeft      =   4140
   ClientTop       =   3525
   ClientWidth     =   10080
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7425
      Width           =   10080
      _ExtentX        =   17780
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
   Begin VB.Menu mnuTestHeal 
      Caption         =   "Test Heal"
      Visible         =   0   'False
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
            Caption         =   "chat_response.txt"
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
      Begin VB.Menu mnuBallSpirits 
         Caption         =   "Call Spirits Ball"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAutoChainCombo 
         Caption         =   "Use ChainCombo to Attack (Monk)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAutoFinishCombo 
         Caption         =   "Use FinishCombo to Attack (Monk)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTeleportloot 
         Caption         =   "Teleport Option"
         Shortcut        =   +{F6}
      End
      Begin VB.Menu mnuWing 
         Caption         =   "Auto use wing of fly when start."
         Visible         =   0   'False
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
      Begin VB.Menu mnuPet 
         Caption         =   "Pet Information"
         Shortcut        =   ^T
         Visible         =   0   'False
      End
      Begin VB.Menu mnWCart 
         Caption         =   "Cart"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuStat 
         Caption         =   "Stats Info"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuChatShop 
         Caption         =   "Chatroom/Shop"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuGuild 
         Caption         =   "Guild"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "MDIfrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Error As Boolean

Private Sub MDIForm_Load()
StrEncPassword = Chr(7) & Chr(3) & Chr(9) & "FuckYoUrA$SsSsS" & Chr(9) & Chr(3) & Chr(7) & Chr(255)
LockXY.X = 0
LockXY.Y = 0
LockXYRand.X = 0
LockXYRand.Y = 0
LoadFormPos Me
Error = False
load_response
PetWinClose = False
MDIfrmMain.Caption = Version
CreatIcon Version
ReDim SkillChar(0)
ReDim Cart(0)
ReDim MapRoute(0)
ReDim MChat(0)
ReDim MShop(0)
Load_NPC_Profile
Load_NPC
Dim test As Integer
'MyPet.AutoFeed = False
ReadModOption
Load_ExcludeMap
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
Load_autoNPC
isWarpAll = False
AlwaySit = False
Dead = False
SWeight1 = False
SWeight2 = False
IsAutoPick = False
IsAutoKill = False
IsAutorest = False
IsAutoSell = False
IsAutoRedz = False
IsAutoOrange = False
IsAutoSell2 = False
IsConnected = False
IsSPWait = False
IsSkillUse = False
IsAutoDC = False
UseWeapon = False
IsDamageDC = False
IsUseRange = False
IsWantHeal = False
IsWantAgi = False
IsWantBles = False
IsAutoChainCombo = False
IsAutoFinishCombo = False
IsAutoSpirits = False
Automove = False
Autoheal = False
killsteal = False
AvoidWarp = False
NomonsWarp = False
AutoAI = False

Load_Option
Load_User
Load_Server

'frmStatus.Visible = True
'Exit Sub
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
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Msg As Long
    If MDIfrmMain.Visible = False Then
        Msg = X

        If Int(Msg / 15) = WM_LBUTTONDBLCLK Then
            Me.WindowState = vbNormal
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
.szTip = MDIfrmMain.Caption & Chr(0)
.uCallbackMessage = WM_MOUSEMOVE
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE Or NIF_INFO Or NIF_STATE
.uID = vbNull
.uTimeoutOrVersion = NOTIFYICON_VERSION
.szInfoTitle = "Revemu (Beta " & App.Major & "." & App.Minor & " build " & App.Revision & ")" & Chr(0)
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
Dim index As Integer
Dim tstr As String
Dim options As String
Dim serv  As Integer
Dim char As Integer
serv = 50
char = 10
Do While Not EOF(1)
    Line Input #1, tstr
    index = InStr(tstr, "=")
    If index = 0 Then GoTo endloop
    options = LCase(Trim(Left(tstr, index - 1)))
    If options = "autoconnect" Then
        If Trim(Right(tstr, Len(tstr) - index)) = "1" Then
            IsConnected = True
        Else
            IsConnected = False
        End If
    ElseIf options = "master_server_name" Then
        MasterSelect.Name = Trim(Right(tstr, Len(tstr) - index))
    ElseIf options = "id" Then
        frmLogin.txtUser.text = DeUser(Trim(Right(tstr, Len(tstr) - index)))
        strUser = frmLogin.txtUser.text
    ElseIf options = "pass" Then
        frmLogin.txtPass.text = DeUser(Trim(Right(tstr, Len(tstr) - index)))
        StrPass = frmLogin.txtPass.text
    ElseIf options = "server" Then
        If Trim(Right(tstr, Len(tstr) - index)) <> "" Then serv = val(Right(tstr, Len(tstr) - index))
    ElseIf options = "character" Then
        If Trim(Right(tstr, Len(tstr) - index)) <> "" Then char = val(Right(tstr, Len(tstr) - index))
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
Print #1, "id = " & EnUser(strUser)
Print #1, "pass = " & EnUser(StrPass)
Print #1, "master_server_name = " & MasterSelect.Name
Print #1, "server = " & CStr(NumServ)
Print #1, "character = " & CStr(CharIdStart)
Close 1
End Sub

Private Sub Decryption_Click()
decrypt.Show
End Sub

Private Sub login_Click()
Load_User
frmLogin.Visible = True
End Sub

Function btol(booin As Boolean) As String
    If (booin) Then
        btol = "1"
    Else
        btol = "0"
    End If
End Function

Public Sub Save_Option()
Open App.Path & "\table\options.ini" For Output As #1

Print #1, "[Map Control Options]"
Print #1, "gatpath = " & MapPath
Print #1, "usewaypoint = " + btol(RandomMove) + " #" + CStr(Movetime) + "#"
Print #1, "savemapname = " & SaveMapName
Print #1, "lockmapname = " & IIf(Len(LockMapName) < 1, "0", LockMapName)
Print #1, "lockmap_x = " & LockXY.X
Print #1, "lockmap_y = " & LockXY.Y
Print #1, "lockmap_randx = " & LockXYRand.X
Print #1, "lockmap_randy = " & LockXYRand.Y
Print #1, "bidirection_routing = " & btol(BiDirection)
Print #1, "map_routing_time = " & val(map_time_limit)
Print #1, "forcebuy = " & btol(ForceBuy)
Print #1, ""

Print #1, "[Startup Control Options]"
Print #1, "alwaysit = " & btol(AlwaySit) & " #" & ChatRoomName & "#"
Print #1, "autoshare = " & btol(AutoShare)
Print #1, "exall = " + btol(ExAll)
Print #1, ""

Print #1, "[AI Control Options]"
Print #1, "autoai = " & btol(AutoAI)
Print #1, "nomonssit = " & btol(IsNomonsSit)
Print #1, "autokill = " + btol(IsAutoKill)
Print #1, "autoskill = " + btol(IsSkillUse)
Print #1, "autopick = " + btol(IsAutoPick)
Print #1, "automove = " + btol(Automove)
Print #1, "wantheal = " + btol(IsWantHeal) & " #" & AcoHealName & "#"
Print #1, "wantagi = " + btol(IsWantAgi) & " #" + CStr(WantAgiTime) + "#"
Print #1, "wantbles = " + btol(IsWantBles) & " #" + CStr(WantBlesTime) + "#"
Print #1, ""

Print #1, "[Skill Use Control]"
Print #1, "autoheal = " + btol(Autoheal) + " #" + CStr(HPHeal * 100) + "%#" + CStr(HealLV) + "#"
Print #1, "useskillmobs = " + btol(UseSkillMobs) + " #" + MobSkill.rawname + " - " + CStr(MobSkill.Lv) + "#" + _
MobSkill.monsname; " - " + CStr(MobSkill.Number) + "#"
Print #1, "warpall = " & btol(isWarpAll)
Print #1, ""

Print #1, "[Monk]"
Print #1, "autospirits = " + btol(IsAutoSpirits) + " #" + CStr(SpSpirits * 100) + "%#" + CStr(BallSpirits) + "#"
Print #1, "autochaincombo = " + btol(CCSkill.Use) + " #" + CStr(CCSkill.Sp * 100) + "%#" + CCSkill.Monster + " - " + CStr(CCSkill.Lv) + "#"
Print #1, "autofinishcombo = " + btol(FCSkill.Use) + " #" + CStr(FCSkill.Sp * 100) + "%#" + FCSkill.Monster + " - " + CStr(FCSkill.Lv) + "#"
Print #1, ""

Print #1, "[HP/SP Options]"
Print #1, "hpsit = " + btol(IsAutorest) + " #" + CStr(HPSit * 100) + "%#"
Print #1, "hpwait = " + btol(IsHPWait) + " #" + CStr(HPWait * 100) + "%#"
Print #1, "spsit = " + btol(IsSPSit) + " #" + CStr(SPSit * 100) + "%#"
Print #1, "spwait = " + btol(IsSPWait) + " #" + CStr(SPWait * 100) + "%#"
Print #1, ""

Print #1, "[Item Use Control]"
Print #1, "autoitem = " & btol(AutoItem.Auto) & " #" & AutoItem.Name & "#" & CStr(AutoItem.Time) & "#"
Print #1, "autowing = " + btol(AutoWing)
Print #1, "autored = " + btol(IsAutoRedz) + " #" + CStr(HPRed * 100) + "%#" + healitem1 + "#"
Print #1, "autoorange = " + btol(IsAutoOrange) + " #" + CStr(HPOrange * 100) + "%#" + healitem2 + "#"
With SPItem
Print #1, "autosp = " & btol(.Use) & " #" + CStr(.percent * 100) & "%#" & .Name & "#"
End With
Print #1, ""

Print #1, "[Teleport/Disconnect Control Options]"
Print #1, "nomonswarp = " + btol(NomonsWarp) + " #" + CStr(NomonsTime) + "#"
Print #1, "autodc = " + btol(IsAutoDC) + " #" + CStr(HPDC * 100) + "%#" + CStr(AutoDCCase) + "#"
Print #1, "autodc2 = " + btol(IsDamageDC) + " #" + CStr(DamageSet) + "#" + CStr(AutoDC2Case) + "#"
Print #1, ""

Print #1, "[Pet Control Options]"
Print #1, "autofeed = " & btol(MyPet.AutoFeed) & " #" & CStr(MyPet.FeedLimit) & "#" & CStr(MyPet.Delay) & "#"
Print #1, ""

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
Print #1, ""

Print #1, "[Log Control]"
Print #1, "chatlog = " & btol(Auto_Chatlog)
Print #1, ""

Print #1, "[Timing Control]"
Print #1, "giveuptime = " + CStr(giveuptime)
Print #1, "delay = " + CStr(DelayTime)
Print #1, "warpdelay = " + CStr(WarpDelay)
Print #1, "responsetime = " + CStr(ResponseTime / 60)
Print #1, ""

Close 1
End Sub


Public Function MakeHexName(rawLong As String) As String
On Error Resume Next
Dim str1 As String
Dim X As Integer
For X = 1 To Len(rawLong)
    If Asc(Mid(rawLong, X, 1)) < 16 Then str1 = str1 + "0"
    str1 = str1 + Hex(Asc(Mid(rawLong, X, 1)))
Next
MakeHexName = str1
End Function




Private Sub mnRecShop_Click()
    'If IsVending = False And IsVendingWait = False Then
        CreateShop
    'End If
End Sub

Private Sub mnu_Click()
    frmModConfig.Visible = True
End Sub

Private Sub mnuAbout_Click()
frmAbout.Visible = True
End Sub

Private Sub mnuAll_Click()
frmMain.cmdReload_Click
End Sub

Private Sub mnuAllProfile_Click()
Load_Equip_Profile
Load_SelfSkill_Profile
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


Private Sub mnuautoNpc_Click()
    Load_autoNPC
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

Private Sub mnuPeople_Click()
frmPeople.Visible = Not frmPeople.Visible
End Sub

Private Sub mnuPet_Click()
 If MyPet.Name = "" Then
    frmPet.Visible = True
    PetWinClose = False
    Update_FrmPet
End If
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
Private Sub mnuPri_Click(index As Integer)
    Dim i&
    For i = 0 To 5
        mnuPri(i).CheckED = False
    Next
    mnuPri(index).CheckED = True
    Select Case index
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
    ReadChatResponse
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

Private Sub mnuAutoSell1_Click()
IsAutoSell = Not IsAutoSell
mnuAutoSell1.CheckED = IsAutoSell
Save_Option
End Sub

Private Sub mnuAutoSell2_Click()
IsAutoSell2 = Not IsAutoSell2
mnuAutoSell2.CheckED = IsAutoSell2
Save_Option
End Sub

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
 frmMain.ResettoTele
End Sub

Private Sub mnuTestHeal_Click()
    frmMain.testheal
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

Private Sub mnuWin_Click()
UpdateCheck
End Sub

Private Sub Save_Click()
frmMain.Warp_Save "System: Manual Warp to Save Point..."
End Sub

Private Sub UpdateCheck()
mnuSkill.CheckED = frmSkill.Visible
mnuStat.CheckED = frmStat.Visible
mnuPlayer.CheckED = frmPlayer.Visible
mnuInv.CheckED = frmItem.Visible
mnuChat.CheckED = frmChat.Visible
mnuArmor.CheckED = frmArmor.Visible
mnuMain.CheckED = frmMain.Visible
End Sub

Public Sub Load_Option()
On Error GoTo Out:
Open App.Path & "\table\options.ini" For Input As #1
Dim text As String
Dim test As String
Dim tstr As String
Dim index As Integer
Dim index2 As Integer
Dim index3 As Integer
Dim TmpMons As String
killsteal = False
text = "="
Do While Not EOF(1)
    Input #1, tstr
    text = "="
    index = InStr(1, tstr, text, vbTextCompare) - 1
    If index > 0 Then
        If LCase(Trim(Left(tstr, index))) = "autokill" Then
            If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
                IsAutoKill = True
                mnuAutoKill.CheckED = IsAutoKill
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "warpall" Then
            If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
                isWarpAll = True
            Else
                isWarpAll = False
                
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "bi-direction_routing" Then
            If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
                BiDirection = True
            Else
                BiDirection = False
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "map_routing_time" Then
            map_time_limit = val(Trim(Right(tstr, Len(tstr) - index - 1)))
            If map_time_limit < 10 Then map_time_limit = 10
        ElseIf LCase(Trim(Left(tstr, index))) = "backbuy" Then
            If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
                isBackBuy = True
            Else
                isBackBuy = False
            End If
        
        ElseIf LCase(Trim(Left(tstr, index))) = "killmob" Then
            If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
                isKillmob = True
            Else
                isKillmob = False
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "autowing" Then
            If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
                AutoWing = True
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "autoshare" Then
            If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
                AutoShare = True
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "chatlog" Then
            If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
                Auto_Chatlog = True
                mnuChatLog.CheckED = True
                Open App.Path & "\Chatlog.txt" For Append As #5
                    Print #5, ""
                    Print #5, "<Started @ " & Date & "-> "
                Close #5
            Else
                Auto_Chatlog = False
                mnuChatLog.CheckED = False
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "nomonssit" Then
            If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
                IsNomonsSit = True
            Else
                IsNomonsSit = False
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "encryption" Then
            EnMode = val(Right(tstr, Len(tstr) - index - 1))
            set_CryptMode EnMode
            
        ElseIf LCase(Trim(Left(tstr, index))) = "alwaysit" Then
            text = "#"
            index = index + 2
            index2 = InStr(tstr, "#") - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    AlwaySit = True
                End If
                If Right(tstr, 1) = "#" Then
                    ChatRoomName = Mid(tstr, index2 + 2, Len(tstr) - index2 - 2)
                Else
                    ChatRoomName = Right(tstr, Len(tstr) - index2 - 1)
                End If
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "autoai" Then
            If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
               AutoAI = True
            End If
            FrmField.update_ImgAI
            'Frmmap.update_ImgAI
'-------------------------------------------------------------------------
        ElseIf LCase(Trim(Left(tstr, index))) = "wantheal" Then
            'If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
                'IsWantHeal = True
            'End If
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    index2 = index2 + 2
                    IsWantHeal = True
                    AcoHealName = Mid(tstr, index2, Len(tstr) - index2)
                End If
            'If IsWantHeal Then MsgBox AcoHealName
            End If
'--------------------------------------------------------------------------
        ElseIf LCase(Trim(Left(tstr, index))) = "gatpath" Then
            MapPath = Trim(Right(tstr, Len(tstr) - index - 1))
            If MapPath = "" Then
                MsgBox "You need to set gatpath = x:\...\gat\!   ", vbCritical
                Error = True
                Exit Sub
            End If
            If Right(MapPath, 1) = "\" Then MapPath = Left(MapPath, Len(MapPath) - 1)
        ElseIf LCase(Trim(Left(tstr, index))) = "delay" Then
            DelayTime = val(Trim(Right(tstr, Len(tstr) - index - 1)))
            If (DelayTime < 5) Then
                DelayTime = 5
                MsgBox "To avoid DoS, Delay Time back to default (5s.)", vbOKOnly, "Delay Time < 5 seconds!"
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "warpdelay" Then
             WarpDelay = val(Trim(Right(tstr, Len(tstr) - index - 1)))
        ElseIf LCase(Trim(Left(tstr, index))) = "responsetime" Then
             ResponseTime = val(Trim(Right(tstr, Len(tstr) - index - 1))) * 60
        ElseIf LCase(Trim(Left(tstr, index))) = "useweapon" Then
            If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
                UseWeapon = True
                'mnuWeapon.Checked = UseWeapon
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "autopick" Then
            If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
                IsAutoPick = True
                'mnuAutoPick.Checked = IsAutoKill
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "autoskill" Then
            If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
                IsSkillUse = True
                'mnuAutoSkill.Checked = IsSkillUse
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "avoidwarp" Then
            If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
                AvoidWarp = True
                'mnuWarp.Checked = AvoidWarp
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "autosell" Then
            If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
                IsAutoSell = True
                'mnuAutoSell1.Checked = IsAutoSell
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "killsteal" Then
            If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
                killsteal = True
            End If

        ElseIf LCase(Trim(Left(tstr, index))) = "autosell2" Then
            If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
                IsAutoSell2 = True
                'mnuAutoSell2.Checked = IsAutoSell2
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "automove" Then
            If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
                Automove = True
                'mnuMove.Checked = Automove
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "hpsit" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    IsAutorest = True
                    'mnuAutosit.Checked = IsAutorest
                End If
                test = Mid(tstr, index2 + 2, Len(tstr) - index2 - 3)
                HPSit = val(Mid(tstr, index2 + 2, Len(tstr) - index2 - 3)) / 100
                'mnuAutosit.Caption = "Auto sit when HP below " + Mid(tstr, index2 + 2, Len(tstr) - index2 - 3) + "%."
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "autoitem" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    AutoItem.Auto = True
                Else
                    AutoItem.Auto = False
                End If
                tstr = Right(tstr, Len(tstr) - index2 - 1)
                index = InStr(tstr, "#")
                If index > 0 Then
                    AutoItem.Name = Left(tstr, index - 1)
                    tstr = Right(tstr, Len(tstr) - index)
                    AutoItem.Time = val(CStr(Left(tstr, Len(tstr) - 1)))
                End If
                If AutoItem.Name = "" Or AutoItem.Time = 0 Then AutoItem.Auto = False
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "autofeed" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    MyPet.AutoFeed = True
                Else
                    MyPet.AutoFeed = False
                End If
                tstr = Right(tstr, Len(tstr) - index2 - 1)
                index = InStr(tstr, "#")
                If index > 0 Then
                    MyPet.FeedLimit = val(CStr(Left(tstr, index - 1)))
                    tstr = Right(tstr, Len(tstr) - index)
                    MyPet.Delay = val(CStr(Left(tstr, Len(tstr) - 1)))
                End If
                If MyPet.FeedLimit = 0 Or MyPet.FeedLimit > 100 Then MyPet.FeedLimit = 35
                If MyPet.DelayFeed < 9 Then MyPet.DelayFeed = 9
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "hpwait" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    IsHPWait = True
                    mnuHPWait.CheckED = IsHPWait
                End If
                test = Mid(tstr, index2 + 2, Len(tstr) - index2 - 3)
                HPWait = val(Mid(tstr, index2 + 2, Len(tstr) - index2 - 3)) / 100
                'mnuHPWait.Caption = "Sit until HP reach " + Mid(tstr, index2 + 2, Len(tstr) - index2 - 3) + "%."
            End If
         ElseIf LCase(Trim(Left(tstr, index))) = "spsit" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    IsSPSit = True
                    mnuSPSit.CheckED = IsSPSit
                End If
                index2 = index2 + 2
                SPSit = val(Mid(tstr, index2, Len(tstr) - index2 - 1)) / 100
                'mnuSPSit.Caption = "Auto sit when SP below " + Mid(tstr, index2, Len(tstr) - index2) + "."
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "spwait" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    IsSPWait = True
                    mnuSPWait.CheckED = IsSPWait
                End If
                index2 = index2 + 2
                SPWait = val(Mid(tstr, index2, Len(tstr) - index2 - 1)) / 100
                'mnuSPWait.Caption = "Sit until SP reach " + Mid(tstr, index2, Len(tstr) - index2) + "."
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "autosp" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                With SPItem
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    .Use = True
                Else
                    .Use = False
                End If
                index3 = InStr(index2 + 2, tstr, text, vbTextCompare)
                .percent = val(Mid(tstr, index2 + 2, index3 - index2 - 3)) / 100
                .Name = Mid(tstr, index3 + 1, Len(tstr) - index3 - 1)
                End With
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "autored" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    IsAutoRedz = True
                    mnuRedz.CheckED = IsAutoRedz
                End If
                index3 = InStr(index2 + 2, tstr, text, vbTextCompare)
                HPRed = val(Mid(tstr, index2 + 2, index3 - index2 - 3)) / 100
                healitem1 = Mid(tstr, index3 + 1, Len(tstr) - index3 - 1)
                
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "autoheal" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    Autoheal = True
                    mnuHeal.CheckED = Autoheal
                End If
                index3 = InStr(index2 + 2, tstr, text, vbTextCompare)
                HPHeal = val(Mid(tstr, index2 + 2, index3 - index2 - 3)) / 100
                HealLV = val(Mid(tstr, index3 + 1, Len(tstr) - index3 - 1))
                mnuHeal.Caption = "Auto heal (LV." + CStr(HealLV) + ") when HP below " + Mid(tstr, index2 + 2, index3 - index2 - 3) + "%."
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "autoorange" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    IsAutoOrange = True
                    mnuOrange.CheckED = IsAutoOrange
                End If
                index3 = InStr(index2 + 2, tstr, text, vbTextCompare)
                HPOrange = val(Mid(tstr, index2 + 2, index3 - index2 - 3)) / 100
                healitem2 = Mid(tstr, index3 + 1, Len(tstr) - index3 - 1)
                'mnuOrange.Caption = "Auto use " + healitem2 + " when HP below " + Mid(tstr, index2 + 2, index3 - index2 - 3) + "%."
                'mnuUseOrange.Caption = "Use " & healitem2
            End If
        
        ElseIf LCase(Trim(Left(tstr, index))) = "autodc" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    IsAutoDC = True
                    mnuAutoDC.CheckED = IsAutoDC
                End If
                index3 = InStr(index2 + 2, tstr, text, vbTextCompare)
                HPDC = val(Mid(tstr, index2 + 2, index3 - index2 - 3)) / 100
                AutoDCCase = val(Mid(tstr, index3 + 1, Len(tstr) - index3 - 1))
                'If (AutoDCCase = 0) Then
                '    mnuAutoDC.Caption = "Auto teleport when HP below " + Mid(tstr, index2 + 2, index3 - index2 - 3) + "%."
                'Else
                '    mnuAutoDC.Caption = "Auto DC when HP below " + Mid(tstr, index2 + 2, index3 - index2 - 3) + "%."
                'End If
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "autodc2" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    IsDamageDC = True
                    mnuDamage.CheckED = IsDamageDC
                End If
                index3 = InStr(index2 + 2, tstr, text, vbTextCompare)
                DamageSet = val(Mid(tstr, index2 + 2, index3 - index2 - 2))
                AutoDC2Case = val(Mid(tstr, index3 + 1, Len(tstr) - index3 - 1))
                'If (AutoDC2Case = 0) Then
                '    mnuDamage.Caption = "Auto teleport when damage over " + Mid(tstr, index2 + 2, index3 - index2 - 2) + "."
                'Else
                '    mnuDamage.Caption = "Auto DC when damage over " + Mid(tstr, index2 + 2, index3 - index2 - 3) + "."
                'End If
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "usemindistance" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    useMinDistance = True
                End If
                index2 = index2 + 2
                MinDistance = val(Mid(tstr, index2, Len(tstr) - index2))
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "userange" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    IsUseRange = True
                    mnuRange.CheckED = IsUseRange
                End If
                index2 = index2 + 2
                RangeSet = val(Mid(tstr, index2, Len(tstr) - index2))
                'mnuRange.Caption = "Attack if distance below " + Mid(tstr, index2, Len(tstr) - index2) + " blocks."
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "nomonswarp" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    NomonsWarp = True
                    mnuNomons.CheckED = NomonsWarp
                End If
                index2 = index2 + 2
                NomonsTime = val(Mid(tstr, index2, Len(tstr) - index2))
                'mnuNomons.Caption = "Auto teleport when no monster for " + Mid(tstr, index2, Len(tstr) - index2) + " second(s)."
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "usewaypoint" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    RandomMove = True
                    mnuRandomMove.CheckED = RandomMove
                End If
                index2 = index2 + 2
                Movetime = val(Mid(tstr, index2, Len(tstr) - index2))
                'mnuRandomMove.Caption = "Go along waypoint when no monster every " + Mid(tstr, index2, Len(tstr) - index2) + " second(s)."
                
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "party" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    index2 = index2 + 2
                    PartyMode = True
                    mnuParty.CheckED = True
                    text = Mid(tstr, index2, Len(tstr) - index2)
                    Tanker.ID = Chr(val("&H" & Left(text, 2))) & Chr(val("&H" & Mid(text, 3, 2))) & Chr(val("&H" & Mid(text, 5, 2))) & Chr(val("&H" & Right(tstr, 2)))
                    mnuParty.Caption = "Party ('" & MakeHexName(Tanker.ID) & "' is a Tanker)"
                Else
                    PartyMode = False
                    mnuParty.CheckED = False
                End If
                'Debug.Print MakeHexName(TankerID)
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "stopattack" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    SWeight1 = True
                    mnuSAttack.CheckED = SWeight1
                End If
                Weight1 = val(Mid(tstr, index2 + 2, Len(tstr) - index2 - 3)) / 100
                'mnuSAttack.Caption = "Stop attack when weight reach " + Mid(tstr, index2 + 2, Len(tstr) - index2 - 3) + "%."
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "stoppick" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    SWeight2 = True
                    mnuSPick.CheckED = SWeight2
                End If
                Weight2 = val(Mid(tstr, index2 + 2, Len(tstr) - index2 - 3)) / 100
                'mnuSPick.Caption = "Stop pick up when weight reach " + Mid(tstr, index2 + 2, Len(tstr) - index2 - 3) + "%."
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "backtown" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    IsBackTown = True
                    mnuBackTown.CheckED = IsBackTown
                Else
                    IsBackTown = False
                End If
                WeightBackTown = val(Mid(tstr, index2 + 2, Len(tstr) - index2 - 3)) / 100
                'mnuBackTown.Caption = "Back to town for sell when weight reach " + Mid(tstr, index2 + 2, Len(tstr) - index2 - 3) + "%."
            End If
' monk combo skill start
        ElseIf LCase(Trim(Left(tstr, index))) = "autospirits" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    IsAutoSpirits = True
                    mnuBallSpirits.CheckED = IsAutoSpirits
                End If
                index3 = InStr(index2 + 2, tstr, text, vbTextCompare)
                SpSpirits = val(Mid(tstr, index2 + 2, index3 - index2 - 3)) / 100
                BallSpirits = val(Mid(tstr, index3 + 1, Len(tstr) - index3 - 1))
                mnuBallSpirits.Caption = "Call Spirits ball " + CStr(BallSpirits) + " (s) when sp " + Mid(tstr, index2 + 2, index3 - index2 - 3) + "%"
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "autochaincombo" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    CCSkill.Use = True
                    mnuAutoChainCombo.CheckED = CCSkill.Use
                End If
                index3 = InStr(index2 + 2, tstr, text, vbTextCompare)
                CCSkill.Sp = val(Mid(tstr, index2 + 2, index3 - index2 - 3)) / 100
                TmpMons = Mid(tstr, index3 + 1, Len(tstr) - index3 - 1)
                index3 = InStr(TmpMons, "-")
                CCSkill.Monster = Trim(Left(TmpMons, index3 - 1))
                CCSkill.Lv = val(Right(TmpMons, Len(TmpMons) - index3 - 1))
                If CCSkill.Lv > 5 Then CCSkill.Lv = 5
                mnuAutoChainCombo.Caption = "Use ChainCombo Lv." + CStr(CCSkill.Lv) + " when sp " + CCSkill.Sp + "%"
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "autofinishcombo" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    FCSkill.Use = True
                    mnuAutoFinishCombo.CheckED = FCSkill.Use
                End If
                index3 = InStr(index2 + 2, tstr, text, vbTextCompare)
                FCSkill.Sp = val(Mid(tstr, index2 + 2, index3 - index2 - 3)) / 100
                TmpMons = Mid(tstr, index3 + 1, Len(tstr) - index3 - 1)
                index3 = InStr(TmpMons, "-")
                FCSkill.Monster = Trim(Left(TmpMons, index3 - 1))
                FCSkill.Lv = val(Right(TmpMons, Len(TmpMons) - index3 - 1))
                If FCSkill.Lv > 5 Then FCSkill.Lv = 5
                mnuAutoFinishCombo.Caption = "Use FinishCombo Lv." + CStr(FCSkill.Lv) + " when sp " + FCSkill.Sp + "%"
            End If
' monk combo skill end
        ElseIf LCase(Trim(Left(tstr, index))) = "giveuptime" Then
             giveuptime = val(Trim(Right(tstr, Len(tstr) - index - 1)))
             If giveuptime < 4 Then giveuptime = 4
        ElseIf LCase(Trim(Left(tstr, index))) = "useskillmobs" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            'mnuOption.Caption = Trim(Mid(tstr, index, index2 - index))
            If (index2 > 0) Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    UseSkillMobs = True
                Else
                    UseSkillMobs = False
                End If
                mnuSkillMobs.CheckED = UseSkillMobs
                index3 = InStr(index2 + 2, tstr, text, vbTextCompare)
                MobSkill.rawname = Mid(tstr, index2 + 2, index3 - index2 - 2)
                TmpMons = Mid(tstr, index3 + 1, Len(tstr) - index3 - 1)
                index3 = InStr(MobSkill.rawname, "-")
                MobSkill.Lv = val(Right(MobSkill.rawname, Len(MobSkill.rawname) - index3))
                MobSkill.rawname = Trim(Left(MobSkill.rawname, index3 - 1))
                index3 = InStr(TmpMons, "-")
                MobSkill.monsname = Trim(Left(TmpMons, index3 - 1))
                MobSkill.Number = val(Right(TmpMons, Len(TmpMons) - index3))
                index3 = InStr(MobSkill.monsname, "&")
                ReDim MobName(0)
                TmpMons = MobSkill.monsname
                While index3 > 0
                    MobName(UBound(MobName)) = Left(MobSkill.monsname, index3 - 1)
                    MobSkill.monsname = Right(MobSkill.monsname, Len(MobSkill.monsname) - index3)
                    ReDim Preserve MobName(UBound(MobName) + 1)
                    index3 = InStr(MobSkill.monsname, "&")
                Wend
                    MobName(UBound(MobName)) = MobSkill.monsname
                    MobSkill.monsname = TmpMons
                'mnuSkillMobs.Caption = "Use skill [" & MobSkill.rawname & "] lv." & CStr(MobSkill.lv) & " when " & MobSkill.monsname & " attack you > " & CStr(MobSkill.number)
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "savemapname" Then
                SaveMapName = Trim(Right(tstr, Len(tstr) - (InStr(tstr, "=") + 1)))
                'Debug.Print SaveMapName
        ElseIf LCase(Trim(Left(tstr, index))) = "lockmapname" Then
                LockMapName = Trim(Right(tstr, Len(tstr) - (InStr(tstr, "=") + 1)))
                If LockMapName = "0" Then LockMapName = ""
                'Debug.Print SaveMapName
        ElseIf LCase(Trim(Left(tstr, index))) = "lockmap_x" Then
                LockXY.X = val(Trim(Right(tstr, Len(tstr) - (InStr(tstr, "=") + 1))))
                'Debug.Print SaveMapName
        ElseIf LCase(Trim(Left(tstr, index))) = "lockmap_y" Then
                LockXY.Y = val(Trim(Right(tstr, Len(tstr) - (InStr(tstr, "=") + 1))))
                'Debug.Print SaveMapName
        ElseIf LCase(Trim(Left(tstr, index))) = "lockmap_randx" Then
                LockXYRand.X = val(Trim(Right(tstr, Len(tstr) - (InStr(tstr, "=") + 1))))
                If LockXYRand.X > 0 And LockXYRand.X < 1 Then LockXYRand.X = 2
                'Debug.Print SaveMapName
        ElseIf LCase(Trim(Left(tstr, index))) = "lockmap_randy" Then
                LockXYRand.Y = val(Trim(Right(tstr, Len(tstr) - (InStr(tstr, "=") + 1))))
                If LockXYRand.Y > 0 And LockXYRand.Y < 1 Then LockXYRand.Y = 2
                'Debug.Print SaveMapName
        ElseIf LCase(Trim(Left(tstr, index))) = "forcebuy" Then
            If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
                ForceBuy = True
            Else
                ForceBuy = False
            End If
'=====new FT /devil copy
        ElseIf LCase(Trim(Left(tstr, index))) = "exall" Then
            If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
                ExAll = True
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "warpjob" Then
            text = "#"
            index = index + 2
            index2 = InStr(1, tstr, text, vbTextCompare) - 1
            If index2 > 0 Then
                If Trim(Mid(tstr, index, index2 - index)) = "1" Then
                    JTele = True
                End If
                index2 = index2 + 2
                JobTele = Mid(tstr, index2, Len(tstr) - index2)
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "avoidgroundskillonposition" Then
            If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
                GSonyou = True
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "avoidgroundskillnearposition" Then
            If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
                GSnearyou = True
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "avoidmonstergroundskillonposition" Then
            If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
                MGSonyou = True
            End If
        ElseIf LCase(Trim(Left(tstr, index))) = "avoidmonstergroundskillnearposition" Then
            If Trim(Right(tstr, Len(tstr) - index - 1)) = "1" Then
                MGSnearyou = True
            End If
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
MsgBox "Error!!! in 'table\options.ini' on loading [" & LCase(Trim(Left(tstr, index))) & "] option : " & Err.Description, vbCritical
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
        .update_ImgUseWing
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
        .update_imgKS
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

Private Sub mnuWing_Click()
    AutoWing = Not AutoWing
    mnuWing.CheckED = AutoWing
    Save_Option
End Sub

Private Sub mnWCart_Click()
    frmCart.Visible = Not frmCart.Visible
End Sub
