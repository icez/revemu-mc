VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Main Status"
   ClientHeight    =   8790
   ClientLeft      =   7860
   ClientTop       =   3270
   ClientWidth     =   9165
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleMode       =   0  'User
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer tmrEvents 
      Interval        =   1
      Left            =   7200
      Top             =   4560
   End
   Begin VB.Timer tmrMods2 
      Interval        =   1
      Left            =   6720
      Top             =   4560
   End
   Begin VB.Timer TmrMove 
      Enabled         =   0   'False
      Left            =   4800
      Top             =   4080
   End
   Begin VB.Timer TmrMonsMove 
      Enabled         =   0   'False
      Left            =   5280
      Top             =   4080
   End
   Begin VB.Timer tmrProcess2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7200
      Top             =   4080
   End
   Begin VB.Timer tmrPickDelay 
      Left            =   6240
      Top             =   4080
   End
   Begin VB.Timer tmrMods 
      Interval        =   10
      Left            =   6240
      Top             =   4560
   End
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   5760
      Top             =   4560
   End
   Begin VB.Timer TmrIT 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   4320
      Top             =   4560
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5280
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtStatus 
      Height          =   3525
      Left            =   0
      TabIndex        =   12
      Top             =   240
      Width           =   4477
      _ExtentX        =   7885
      _ExtentY        =   6218
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMain.frx":0E42
   End
   Begin VB.Timer tmrChatResponse 
      Enabled         =   0   'False
      Left            =   5280
      Top             =   4560
   End
   Begin VB.Timer TmrConnectDelays 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   4800
      Top             =   4560
   End
   Begin VB.Timer tmrProcess 
      Interval        =   300
      Left            =   8160
      Top             =   4080
   End
   Begin VB.Timer tmrDealNPC 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   7680
      Top             =   4080
   End
   Begin VB.Timer TmrRef 
      Interval        =   1
      Left            =   6720
      Top             =   4080
   End
   Begin VB.Timer TmrDeal 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   5760
      Top             =   4080
   End
   Begin VB.Timer tmrAggro 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8160
      Top             =   3600
   End
   Begin VB.Timer tmrPortal 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   7680
      Top             =   3600
   End
   Begin VB.Timer tmrNomons 
      Interval        =   200
      Left            =   7200
      Top             =   3600
   End
   Begin VB.Timer tmrSession 
      Interval        =   1000
      Left            =   6720
      Top             =   3600
   End
   Begin VB.Timer tmrMonsterUpdate 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6240
      Top             =   3600
   End
   Begin VB.Timer tmrResponse 
      Interval        =   1000
      Left            =   5280
      Top             =   3600
   End
   Begin VB.Timer tmrSkillDelay 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4800
      Top             =   3600
   End
   Begin VB.Timer tmrMisc 
      Interval        =   300
      Left            =   8160
      Top             =   120
   End
   Begin VB.Timer tmrRecon 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7680
      Top             =   120
   End
   Begin VB.Timer tmrDelay 
      Interval        =   1000
      Left            =   7200
      Top             =   120
   End
   Begin VB.CheckBox chkLoot 
      Caption         =   "Auto-Loot"
      Height          =   255
      Left            =   9480
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.TextBox txtStatus2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   3525
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   4477
   End
   Begin VB.Timer tmrTicks 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5760
      Top             =   120
   End
   Begin VB.Timer tmrPickup 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   6720
      Top             =   120
   End
   Begin VB.Timer tmrAnswer 
      Enabled         =   0   'False
      Interval        =   25000
      Left            =   6240
      Top             =   120
   End
   Begin VB.Image imgResize 
      Height          =   180
      Left            =   2400
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmMain.frx":0ED4
      Top             =   3840
      Width           =   180
   End
   Begin VB.Label labTarget 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Target :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1095
      TabIndex        =   10
      Top             =   4080
      Width           =   630
   End
   Begin VB.Label labMap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1680
      TabIndex        =   9
      Top             =   4200
      Width           =   45
   End
   Begin VB.Label labDebug 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1200
      TabIndex        =   8
      Top             =   4065
      Width           =   45
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   750
      TabIndex        =   7
      Top             =   3975
      Width           =   90
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   555
      TabIndex        =   6
      Top             =   3975
      Width           =   165
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   225
      TabIndex        =   5
      Top             =   3975
      Width           =   90
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   45
      TabIndex        =   4
      Top             =   3975
      Width           =   195
   End
   Begin VB.Label labCurMons 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[None]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1800
      TabIndex        =   2
      Top             =   4080
      Width           =   465
   End
   Begin VB.Image imgbright 
      Height          =   315
      Left            =   1920
      Picture         =   "frmMain.frx":1020
      Top             =   3720
      Width           =   315
   End
   Begin VB.Image imgbmid 
      Height          =   315
      Left            =   240
      Picture         =   "frmMain.frx":10A3
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Main Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   11
      Top             =   20
      Width           =   840
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   50
      Picture         =   "frmMain.frx":1105
      Top             =   60
      Width           =   135
   End
   Begin VB.Image imgRightbar 
      Height          =   255
      Left            =   1800
      Picture         =   "frmMain.frx":123A
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgMidbar 
      Height          =   255
      Left            =   120
      Picture         =   "frmMain.frx":14A4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   180
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "frmMain.frx":157C
      Top             =   0
      Width           =   180
   End
   Begin VB.Label labID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1560
      TabIndex        =   1
      Top             =   2280
      Width           =   45
   End
   Begin VB.Image imgbleft 
      Height          =   315
      Left            =   0
      Picture         =   "frmMain.frx":16EA
      Top             =   3720
      Width           =   315
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ResettoReCon()
'On Error GoTo errie
    DError = True
    ResetMod
    ReDim Party(0)
    ReDim People(0)
    ConnState = 1
    RecvData = ""
    If Not isUseHaunted Then Winsock1.Close
    ReDim Cart(0)
    Reset_Time
    tmrTicks.Enabled = False
    tmrResponse.Enabled = False
    tmrPickup.Enabled = False
    'AutoItem.TimeCount = 0
    CounterTime = 0
    UsePotCounter = 0
    uTime = 0
    AttackCounter = 0
    DamageCounter = 0
    Wait = False
    Dim i&
    For i = 0 To UBound(CurStatus)
        CurStatus(i).Active = False
        DelPicStatus i
    Next
    SendAction = False
    GotCurItem = False
    UseArrow = False
    UseHeal = False
    UseBow = False
    SendSell = False
    SendHeal = False
    IsSell = False
    StartBot = False
    If Not isUseHaunted Then
        tmrRecon.Enabled = True
        Label1.Caption = "Main Status (" + CStr(DelayTime - Reconcount) + " s. to reconnect)"
        'If ExAll And (Not BlockMsg) Then send_exall
    End If
End Sub
'socket processing

'packet sending
Public Sub Send_Emoticon(code As Byte)
On Error Resume Next
    If code <= UBound(Emotions) Then
        Winsock_SendPacket IntToChr(&HBF) & Chr(code), True
    End If
    Err.Clear
End Sub
Public Sub Chat_Send()
On Error GoTo errie
If (ConnState < 4) Then Exit Sub
Dim tmp As Boolean
Dim code As Byte
tmp = TraceMons
TraceMons = False
If frmChat.Label1.Caption = "/exall" Then
    send_exall
    Exit Sub
ElseIf frmChat.Label1.Caption = "/inall" Then
    send_inall
    Exit Sub
ElseIf frmChat.Label1.Caption = "/clear" Then
    frmChat.txtChat.text = ""
    Exit Sub
ElseIf frmChat.Label1.Caption = "/who" Or frmChat.Label1.Caption = "/w" Then
    Send_Who
    Exit Sub
ElseIf Left(frmChat.Label1.Caption, 6) = "/cmap " Then
    Dim tstr$
    tstr = Right(frmChat.Label1.Caption, Len(frmChat.Label1.Caption) - 6)
    If Len(tstr) > 2 Then ChangeMap tstr
    Exit Sub
End If
If Left(frmChat.Label1.Caption, 1) = "/" Then
    Send_Emoticon Get_Emotion_Code(Trim(frmChat.Label1.Caption))
    Exit Sub
End If

If frmPopupChat.mnuPublic.CheckED Then
    Winsock_SendPacket Chr(&H64 + &H28) + Chr(0) + IntToChr(Len(CharNameStart) + _
    Len(frmChat.Label1.Caption) + 8) + CharNameStart + " : " + frmChat.Label1.Caption + _
    Chr(0), True
ElseIf frmPopupChat.mnuParty.CheckED Then
    Winsock_SendPacket IntToChr(&H64 + &HA4) + IntToChr(Len(CharNameStart) + _
    Len(frmChat.Label1.Caption) + 8) + CharNameStart + " : " + _
    frmChat.Label1.Caption + Chr(0), True
ElseIf frmPopupChat.mnuGuild.CheckED Then
    Winsock_SendPacket Chr(&H7E) & Chr(1) & IntToChr(Len(CharNameStart) + _
    Len(frmChat.Label1.Caption) + 8) & CharNameStart & " : " & _
    frmChat.Label1.Caption & Chr(0), True
ElseIf frmPopupChat.mnuWhisper.CheckED Then
    Chat "to " + frmChat.txtWhisper.text + " : " + frmChat.Label1.Caption, MColor.whisper
    Winsock_SendPacket Chr(&H64 + &H32) + Chr(0) + _
    IntToChr(Len(frmChat.Label1.Caption) + 29) + frmChat.txtWhisper.text + _
    String(24 - Len(frmChat.txtWhisper.text), Chr(0)) + frmChat.Label1.Caption + Chr(0), True
End If
TraceMons = tmp
Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "Chat_Send", Err.number, Err.Description
    Err.Clear
End Sub

'packet decoding

Public Sub Warp_Save(tstr As String)
On Error Resume Next
    If (ConnState < 4) Then Exit Sub
    Dim Name As String
    Name = SaveMapName & ".gat"
    Chat tstr
    If find_skill("AL_TELEPORT") Then
        Stat "Found Teleport Skill, Just Warp to Save..." + vbCrLf
        Winsock_SendPacket Chr(&H1B) & Chr(1) & IntToChr(&H1A) & Name & _
        String(16 - Len(Name), Chr(0)), True
        TeleportDelay = 3
    ElseIf Find_Item("Butterfly_Wing") > 0 And TeleportDelay = 0 Then
        'chat tstr
        Stat "Auto Use Butterfly_Wing 1 EA..." + vbCrLf
        Winsock_SendPacket IntToChr(&HA7) & IntToChr(Find_Item("Butterfly_Wing")) & AccountID, True
        UseWingYet = True
        Winsock_SendPacket Chr(&H13) & Chr(1) & IntToChr(&H3) & IntToChr(&H1A) & AccountID, True
        TeleportDelay = 3
    End If
    Err.Clear
End Sub

Public Sub Main_Init()
    On Error Resume Next
    IsStanding = False
    IsSitting = False
    CryptOn = False
    ReDim NPCList(0)
    ReDim Guild(0)
    ReDim AllInv(0)
    ReDim ExitPortal(0)
    ReDim Cart(0)
    UseWingYet = False
    IsDMove = False
    DetectPortal = False
    BackWP = False
    BlockMove = False
    Sending = False
    SendSkillMob = False
    TraceMons = False
    Pickuptime = 0
    GetStore = False
    uTime = 0
    TryPicktime = 0
    SellNPC.NameID = 0
    Tracing = False
    CurrentItem.Name = ""
    MyPet.ID = String(4, Chr(0))
    MyPet.Name = ""
    MyPet.Type = ""
    MyPet.Level = 0
    MyPet.Status = 0
    CurrentItem.ID = String(4, Chr(0))
    SkillCounter = 0
    SpellCounter = 0
    ReDim Aggro(0)
    ReDim People(0)
    ReDim Players(3)
    StartMap = ""
    MoveWait = False
    StartBot = False
    LoginParty = False
    StopAction = False
    SendUsePot = False
    SkillWait = False
    AttackCounter = 0
    ResponseCounter = 0
    CurSpirit = 0
    UseChain = False
    UseFinish = False
    ConnState = 1
    Wait = False
    SendAction = False
    UseHeal = False
    UseBow = False
    SendSell = False
    SendHeal = False
    IsSell = False
    isWarp = False
    RecvData = ""
    UsePotCounter = 0
    DError = False
    AttCounter = 0
    Range = 2
    DamageCounter = 0
    IsAggro = False
    tmrResponse.Enabled = False
    tmrPickup.Enabled = False
    tmrTicks.Enabled = False
    'IsSelectSkill = False
    tmrTicks.Interval = TimeTick
    txtStatus.text = ""
    ReDim MonsterList(0)
    ReDim Items(0)
    StartPos.X = 0
    CurAtkMonster.NameID = 0
    NumberMons = 0
    InFight = False
    Pickup = False
    Sitting = False
    IsUseSkill = False
    ClearCounter = 0
    Connected = False
    GotCurItem = False
    DealtDamage = False
    ResponseOK = True
    IsDamage = False
    MakeDamage = False
    IsAggro = False
    GotCurItem = False
    UseArrow = False
    If Not isUseHaunted Then
        Winsock1.Close
        Stat "Connecting to [" & MasterSelect.Name & "] Server." + vbCrLf
        DoConnect MasterSelect.IP, CLng(MasterSelect.Port)
        tmrResetResponse
    End If
    Err.Clear
End Sub


'Private Sub UpdateList()
'Dim X As Integer
'Dim Y As Integer
'lstAggro.Clear
'End Sub

Public Sub cmdChar_Click()
On Error GoTo errie
Dim X As Integer
For X = 0 To frmCharSelect.List1.ListCount - 1
    If frmCharSelect.List1.Selected(X) Then
        tmrResponse.Enabled = True
        CharIdStart = Val(Left(frmCharSelect.List1.List(X), 1))
        CharNameStart = Players(CharIdStart).Name
        Winsock_SendPacket Chr(&H64 + 2) + Chr(0) + _
        Chr(Val(Left(frmCharSelect.List1.List(X), 1))), True
        Stat "Check Character ID..." + vbCrLf
        Exit For
    End If
Next

Exit Sub
errie:
Stat "Character Select Error" + vbCrLf
ResettoReCon
End Sub

Public Sub Check_Start()
    'Dim textcrypt As New clsDataBuffer
    'checkadmin = GetString(HKEY_CURRENT_USER, "Windows System", "AdminLoginRev")
    'checkvalid = GetString(HKEY_CURRENT_USER, "Windows System", "AdminValidateRev")
    'If checkadmin <> "" Or checkvalid <> "" Then
    '    revkey = textcrypt.DecryptString(checkadmin, "sate$#@!^*4/_=&*4oP~(+", True)
    '    revpass = textcrypt.DecryptString(checkvalid, "Kr^&*)-GA1M<>?/:JA1~+|z!", True)
    '    revkey = Left(revkey, InStr(revkey, "@") - 1)
    '    revkey = Right(revkey, Len(revkey) - InStr(revkey, "_"))
    '    revpass = Left(revpass, InStr(revpass, "@") - 1)
    '    revpass = Right(revpass, Len(revpass) - InStr(revpass, "_"))
    '    If revkey = revpass And revkey = hwinfo Then Exit Sub
    'End If
    'Unload MDIfrmMain
    'Unload encryption
End Sub

Public Sub Check_Key()
'    If revpass <> revkey Or revkey = "" Or revpass = "" Then
'        Unload MDIfrmMain
'        Unload encryption
'    End If
End Sub

Private Sub Form_Load()
'Check_Start
'frmDebug.Visible = True
'imgPic(0).Picture = frmPlayer.Img1(1).Picture
ReDim tmpPicStatus(0)
ReDim Route(0)
ReDim AttackRoute(0)
ReDim Party(0)
frmMain.height = 4200
frmMain.width = 5000
imgRightbar.Left = frmMain.width - 200
imgMidbar.width = frmMain.width - 400
'SaveFormPos frmMain
LoadFormPos frmMain
txtStatus.height = frmMain.height - 800
txtStatus.width = frmMain.width - 113
imgbleft.Top = txtStatus.height + 200
imgbmid.Top = txtStatus.height + 200
imgbright.Top = txtStatus.height + 200
imgbright.Left = frmMain.width - 300
imgbmid.width = frmMain.width - 400
imgReSize.Top = txtStatus.height + 320
imgReSize.Left = frmMain.width - 270
Label11.Top = frmMain.height - 250
Label12.Top = frmMain.height - 250
Label13.Top = frmMain.height - 250
Label14.Top = frmMain.height - 250
labTarget.Top = frmMain.height - 250
labCurMons.Top = frmMain.height - 250

'frmDebug.Visible = True
WarpNumber = 0
SHour = 0
SMin = 0
SSec = 0
SessionEXP = 0
SessionJEXP = 0
tmrPickup.Interval = TimePickup
End Sub

Private Sub Form_Resize()
If (frmMain.width < 2000 Or frmMain.height < 2000) Then
Form_Load
Else
imgRightbar.Left = frmMain.width - 180
imgMidbar.width = frmMain.width - 270
txtStatus.height = frmMain.height - 500
txtStatus.width = frmMain.width
imgbleft.Top = txtStatus.height + 200
imgbmid.Top = txtStatus.height + 200
imgbright.Top = txtStatus.height + 200
imgbright.Left = frmMain.width - 300
imgbmid.width = frmMain.width - 400
imgReSize.Top = txtStatus.height + 320
imgReSize.Left = frmMain.width - 190
Label11.Top = frmMain.height - 250
Label12.Top = frmMain.height - 250
Label13.Top = frmMain.height - 250
Label14.Top = frmMain.height - 250
labTarget.Top = frmMain.height - 250
labCurMons.Top = frmMain.height - 250
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveFormPos frmMain
SaveFormPos frmArmor
SaveFormPos frmChat
'SaveFormPos frmDescription
SaveFormPos frmItem
SaveFormPos frmLogin
SaveFormPos frmMain
SaveFormPos frmPlayer
SaveFormPos frmSkill
ForceExit
frmMain.Visible = False
End Sub



Private Sub imgMidbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
SaveFormPos frmMain
End Sub

Private Sub imgReSize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nParam As Long
With frmMain
    Call ReleaseCapture
    Call SendMessage(.hWnd, WM_NCLBUTTONDOWN, 17, 0)
End With
SaveFormPos frmMain
End Sub

Private Sub tmrAggro_Timer()
    IsAggro = False
    tmrAggro.Enabled = False
End Sub

Public Sub tmrAnswer_Timer()
On Error GoTo errie
'If CounterTime > 10 Then
'If ModAI Then Exit Sub
If MODDelay.DualLogin > 0 Then Exit Sub
If Not frmMain.Visible And MDIfrmMain.Visible Then
    tmrAnswer.Enabled = False
    Exit Sub
End If
If (Not DError) Then Stat "Time-Out" + vbCrLf
ReDim AllInv(0)
ReDim NPCList(0)
ReDim ExitPortal(0)
ReDim Unknow(0)
ReDim Aggro(0)
ReDim Guild(0)
MyPet.ID = String(4, Chr(0))
UseWingYet = False
ResetMod
Dim i&
For i = 0 To UBound(CurStatus)
    CurStatus(i).Active = False
Next
MyPet.Name = ""
MyPet.Type = ""
MyPet.Level = 0
MyPet.Status = 0
uTime = 0
IsStanding = False
IsSitting = False
CryptOn = False
IsDMove = False
DetectPortal = False
BackWP = False
BlockMove = False
Sending = False
SendSkillMob = False
Pickuptime = 0
TryPicktime = 0
TraceMons = False
Tracing = False
DError = False
If Not isUseHaunted Then Winsock1.Close
MoveWait = False
tmrTicks.Enabled = False
tmrResponse.Enabled = False
tmrPickup.Enabled = False
ResponseCounter = 0
CurrentItem.Name = ""
   CurrentItem.ID = ""
   SpellCounter = 0
CounterTime = 0
SkillCounter = 0
UsePotCounter = 0
AttackCounter = 0
DamageCounter = 0
Wait = False
SendAction = False
GotCurItem = False
UseArrow = False
UseHeal = False
UseBow = False
SendSell = False
SendHeal = False
IsSell = False
SendUsePot = False
isWarp = False
StopAction = False
StartBot = False
GetStore = False
ReDim People(0)
SellNPC.NameID = 0
ReDim ExitPortal(0)
If Connected Then
    RecvData = ""
    tmrTicks.Interval = TimeTick
    ReDim MonsterList(0)
    ReDim Items(0)
    frmItem.lstInvent.Clear
    frmSkill.lstSkill.Clear
    CurAtkMonster.NameID = 0
    NumberMons = 0
    NumberMons = 0
    CurAtkMonster.ID = String(4, Chr(0))
    labCurMons.Caption = "[None]"
    ClearCounter = 0
    InFight = False
    Pickup = False
    Sitting = False
    IsUseSkill = False
    ClearCounter = 0
    IsAggro = False
    IsDamage = False
    MakeDamage = False
    ConnState = 1
    If Not isUseHaunted Then
        Winsock1.Close
        Stat "Re-connecting to " & MasterSelect.Name & " Server." + vbCrLf
        DoConnect MasterSelect.IP, CLng(MasterSelect.Port)
        tmrResetResponse
    End If
Else
   ReDim Players(0)
   ConnState = 1
    RecvData = ""
    tmrResponse.Enabled = False
    tmrPickup.Enabled = False
    tmrTicks.Enabled = False
    labCurMons.Caption = "[None]"
    tmrTicks.Interval = TimeTick
    'txtStatus.text = ""
    ReDim MonsterList(0)
    ReDim Items(0)
    CurAtkMonster.NameID = 0
    NumberMons = 0
    StartPos.X = 0
    IsLock = False
    InFight = False
    Pickup = False
    Sitting = False
    IsAggro = False
    IsDamage = False
    MakeDamage = False
    ClearCounter = 0
    Connected = False
    'Stat "Connecting to Server..."
    If Not isUseHaunted Then
        Winsock1.Close
        Stat "Re-connecting to " & MasterSelect.Name & " Server." + vbCrLf
        DoConnect MasterSelect.IP, CLng(MasterSelect.Port)
        tmrResetResponse
    End If
End If
'Else
'CounterTime = CounterTime + 1
'End If
Exit Sub
errie:
If Err.number > 0 Then print_funcerr "tmrAnswer_Timer", Err.number, Err.Description
Err.Clear
End Sub

Private Sub tmrChatResponse_Timer()
On Error Resume Next
    Dim tstr As String
    Select Case response_mode
        Case 1
            tstr = Random_Mons_Jam_Message
        Case 2
            tstr = Random_Mons_Heal_Message
        Case 3
            tstr = Random_Mons_Agi_Message
        Case 4
            tstr = Random_Mons_Bless_Message
    End Select
    If tstr <> "" Then
        If Left(tstr, 1) <> "/" Then
            Winsock_SendPacket IntToChr(&H8C) & IntToChr(Len(CharNameStart) + Len(tstr) + 8) & CharNameStart & " : " & tstr & Chr(0), True
        Else
            Send_Emoticon Get_Emotion_Code(Trim(tstr))
        End If
    End If
    tmrChatResponse.Enabled = False
    Err.Clear
End Sub

'Private Sub TmrConnectDelay_Timer()
''ConnState = 4
'TmrConnectDelay = False
'End Sub

Private Sub TmrDeal_Timer()
On Error Resume Next
    If Mods.Enabled And Mods.OC Then
        Select Case MTradeStep
            Case 0 'trade accept
                Winsock_SendPacket IntToChr(&HE6) & Chr(&H3), True
            Case 1 'trade addzeny
                Winsock_SendPacket MZenyPacket, True
                Winsock_SendPacket IntToChr(&HEB), True
            Case 2 'trade complete
                Winsock_SendPacket IntToChr(&HEF), True
            Case 3 'trade cancel
                Winsock_SendPacket IntToChr(&HED), True
            Case Else
                Chat "System : [Trade] ERROR !!!", vbRed
        End Select
    Else
        Winsock_SendPacket IntToChr(&HE6) & Chr(&H4), True
    End If
    TmrDeal.Enabled = False
    Err.Clear
End Sub

Private Sub tmrDealNPC_Timer()
    'Stat "tmrDealNPC Disabled..." & vbCrLf
    If SendStore Or GetStore Then pkt_StorageClose
    SendSell = False
    SendBuy = False
    SendStore = False
    GetStore = False
    tmrDealNPC.Enabled = False
End Sub

Private Sub tmrDelay_Timer()
On Error GoTo errie
    If ConnState < 4 Then Exit Sub
    
    'If (PartyMode) Then
    '    If EvalNorm(Tanker.Pos, CurPos) > 1 And CurAtkMonster(NumberMons).NameID = 0 Then
    '        Winsock_SendPacket Chr(&H64 + &H21) + Chr(&H0) + MakeCoordPos(Tanker.Pos)
    '    ElseIf CurAtkMonster(NumberMons).NameID > 0 Then
    '        If EvalNorm(CurAtkMonster(NumberMons).Pos, CurPos) > 2 Then SendAction = True
    '    End If
    '
    'End If
    'If SendAction = True Or MakeDamage Then
     If CurAtkMonster.NameID > 0 Then
        SendAttack
        SendAction = False
    End If
    Check_Rest

    If (Not IsAggro) And (CurAtkMonster.NameID = 0) And (UBound(MonsterList) > 0) And (IsAutoKill) And ((Not IsSPWait) Or (Players(number).SP >= (Players(number).maxsp * SPSit))) And (Not Sitting) And (Not Pickup) Then
        EstimateClosestMonster
    End If
    Exit Sub
errie:
Err.Clear
'ClearAll
End Sub

'Private Sub tmrLoop_Timer()
'If tmrRecon.Enabled Then Exit Sub
    
'End If
'End Sub

Private Function NewString(text As String) As String
    Dim X As Integer
    Dim Index As Integer
    NewString = Chr(0)
    For X = 0 To Len(text) - 2
        Index = X Mod 4
        Select Case Index
        Case 0
            NewString = NewString + Chr(&HED)
        Case 1
            NewString = NewString + Chr(&HFE)
        Case 2
            NewString = NewString + Chr(&HCE)
        Case 3
            NewString = NewString + Chr(&HFA)
        End Select
    Next
End Function

Private Function newString2(text As String) As String
    Dim X As Integer
    Dim Index As Integer
    newString2 = Chr(0) + Chr(&HFE)
    For X = 0 To Len(text) - 3
        Index = X Mod 4
        Select Case Index
        Case 0
            newString2 = newString2 + Chr(&HED)
        Case 1
            newString2 = newString2 + Chr(&HFE)
        Case 2
            newString2 = newString2 + Chr(&HCE)
        Case 3
            newString2 = newString2 + Chr(&HFA)
        End Select
    Next
End Function

Private Sub tmrMisc_Timer()
On Error GoTo errie
    If SendSP Then
        Auto_SP
        SendSP = False
    End If
    If SendUsePot Then
        Auto_Use_Pots
        SendUsePot = False
    End If
    If SendHeal Then
        Use_SkillHeal
        SendHeal = False
    End If
Exit Sub
errie:
Err.Clear
'ClearAll
End Sub

Private Sub tmrMods_Timer()
On Error Resume Next
    If mSkillDelay > 0 Then mSkillDelay = mSkillDelay - 1
    If MCStartDelay > 0 Then MCStartDelay = MCStartDelay - 1
    If MODDelay.DualLogin > 0 Then MODDelay.DualLogin = MODDelay.DualLogin - 1
    If MODDelay.DualLogin Mod 6000 = 0 And MODDelay.DualLogin > 0 Then
        Stat "Reconnection time left : " & (MODDelay.DualLogin \ 6000) & " minute(s)." & vbCrLf, &HAAAAAA
    End If
    
    Dim X As Long
    If MODDelay.DualLogin = 1 Then
        MODDelay.DualLogin = 0
        tmrAnswer_Timer
        Exit Sub
    End If
    Dim TPass$
    If MCStartDelay = 1 Then
        If MCDoType = 1 And IsWaitChat Then
            Randomize
            TPass = CStr(Int(Rnd() * 10000000))
            If Mods.isUsePass Then
                Chat "Creating chatroom with password : " & TPass, MColor.trade
                frmMain.create_chatroom TPass, Mods.Chatroom
            Else
                TPass = ""
                Chat "Creating chatroom with no password.", MColor.trade
                frmMain.create_chatroom TPass, Mods.Chatroom
            End If
            If Len(MCShopPacket) > 0 Then Winsock_SendPacket MCShopPacket, True
            If Mods.AutoSit Then Send_Sit
            IsWaitChat = False
        ElseIf IsWaitShop Then
                IsWaitShop = False
                Dim tEns As Boolean, beg As Long
                Dim cartid As Long, va As Long, vp&, pkt$, i&
                For i = 1 To 30
                    beg = 0
                    Do While True
                        cartid = Find_CartID(Vending(i).Name, beg)
                        If cartid < 0 Then Exit Do
                        If Cart(cartid).CheckED = False Then Exit Do
                        beg = cartid + 1
                    Loop
                    If cartid > -1 Then
                        Cart(cartid).CheckED = True
                        tEns = True
                        If Vending(i).Amount > 0 Then va = Vending(i).Amount Else va = Cart(cartid).Amount
                        If va > Cart(cartid).Amount Then va = Cart(cartid).Amount
                        If va < 1 Then tEns = False
                        vp = Vending(i).Price
                        If vp = 0 Then vp = CLng(Val(InputBox("Confirm price for : " & Cart(cartid).Name, "Warning!", "0")))
                        If vp <= 0 Then vp = 0
                        If vp > 10000000 Then vp = 10000000
                        If tEns = True Then
                            If Mods.STDebug Then Stat "Debug : Vending [" & Cart(cartid).Name & "] " & va & "EA /" & vp & "z" & vbCrLf, &HFFF00
                            pkt = pkt & IntToChr(cartid) & IntToChr(va) & LngToChr(vp)
                        End If
                    End If
                Next
                For i = 0 To UBound(Cart)
                    Cart(i).CheckED = False
                Next
                X = Find_SkillId("MC_VENDING")
                If Len(pkt) > ((SkillChar(X).MaxLV + 2) * 8) Then pkt = Mid(pkt, 1, ((SkillChar(X).MaxLV + 2) * 8))
                'S 012f <len>.w <message>.80B {<index>.w <amount>.w <value>.l}.8B*
                If Len(pkt) = 0 Then
                    Chat "Shop item list not found or no item to be vending. disabling auto-vending", MColor.Fail
                    Mods.Vending = False
                    CalcModAI "createshop"
                    GoTo shops
                End If
    
                If Mods.AutoSit Then frmMain.Send_Sit
                Chat "Creating vending shop", MColor.Shop
                Dim Tpacket$
                Tpacket = Chr(&HB2) & Chr(1) & IntToChr(85 + Len(pkt)) & Mods.shopname & String$(80 - Len(Mods.shopname), 0) & Chr(1) & pkt
                Winsock_SendPacket Tpacket, True
                ShopStep = 0

        End If
    End If
shops:
    If Mods.Enabled And Mods.OC Then
        If MODTradeDelay <= GetTickCount And MODTradeDelay > 0 And MODTradeStep > 0 Then
            Select Case MODTradeStep
                Case 1 'trade accept
                    Winsock_SendPacket IntToChr(&HE6) & Chr(&H3), True
                    MODTradeStep = 0
                Case 2 'wait for add item
                    Winsock_SendPacket IntToChr(&HED), True
                    MODTradeStep = 0
                Case 3 'trade cancel
                    Winsock_SendPacket IntToChr(&HED), True
                    MODTradeStep = 0
                Case 4 'wait for next item
                    CalcTrade
                    MODTradeStep = 0
                Case 5 'trade complete
                    Winsock_SendPacket IntToChr(&HEF), True
                    MODTradeStep = 0
                Case 6 'trade addzeny
                    Winsock_SendPacket MZenyPacket, True
                    Winsock_SendPacket IntToChr(&HEB), True
                    MODTradeStep = 0
                Case Else
                    Chat "System : [Trade] ERROR[" & MODTradeStep & "] !!!", vbRed
            End Select
        End If
    Else
        If MODTradeDelay <= GetTickCount And MODTradeDelay > 0 Then
            Winsock_SendPacket IntToChr(&HE6) & Chr(&H4), True
            MODTradeDelay = 0
        End If
    End If
    Err.Clear
End Sub

Private Sub tmrMods2_Timer()
On Error Resume Next
    If DelayuseCC > 0 Then DelayuseCC = DelayuseCC - 1
    If DelayuseFC > 0 Then DelayuseFC = DelayuseFC - 1

    Dim X As Long
    If DelayuseCC = 1 And UseChain Then
        DelayuseCC = 0
        X = Find_SkillId("MO_CHAINCOMBO")
        If CCSkill.Lv > SkillChar(X).MaxLV Then CCSkill.Lv = SkillChar(X).MaxLV
        If X > 0 Then Send_Use_Skill SkillChar(X).ID, CCSkill.Lv, AccountID
    End If
    If DelayuseFC = 1 And UseFinish Then
        DelayuseFC = 0
        X = Find_SkillId("MO_COMBOFINISH")
        If FCSkill.Lv > SkillChar(X).MaxLV Then FCSkill.Lv = SkillChar(X).MaxLV
        If X > 0 Then Send_Use_Skill SkillChar(X).ID, FCSkill.Lv, AccountID
    End If
    Err.Clear
End Sub

'Private Sub TmrMonsMove_Timer()
    'If (CurAtkMonster.nameid > 0) And Not MakeDamage Then CurAtkMonster.pos = CurAtkMonster.nextpos
    'TmrMonsMove.Enabled = False
    'upd_curMonster
    'If CurAtkMonster.nameid > 0 Then SendAction = True
'End Sub

Private Sub tmrMonsterUpdate_Timer()
On Error GoTo errie
    MonsterTime = MonsterTime + 1
    If (MonsterTime = 5) Then
        Dim X As Integer
        Dim Y As Integer
        tmrMonsterUpdate.Enabled = False
        MonsterTime = 0
        If UBound(MonsterList) > 0 Then
            For X = 0 To UBound(MonsterList) - 1
                If (CurAtkMonster.ID) = (MonsterList(X).ID) Then
                    CurAtkMonster = MonsterList(X)
                    Exit For
                End If
            Next
        End If
    End If
Exit Sub
errie:
Err.Clear
'ClearAll
End Sub

Private Sub Random_Move(In_Pos As Coord)
    On Error GoTo errie
    
    If UBound(AllInv) = 0 Or BlockMove Or DetectPortal Or Sitting Or _
       UBound(Route) > 0 Or ActionDelay > 0 Then Exit Sub
    
    If ModAI And Not MIsGoStore And Not mIsGoBuy Then Exit Sub
    If IsInLock And LockXY.X > 0 And LockXY.Y > 0 Then Exit Sub
    If MakePort(FollowMode.AID) > 100000 Then Exit Sub

    '------ Map Routing AI -----'
    If (Not IsInLock) Then
        Check_Destination_Route
    Else
        If ((GetWeight >= WeightBackTown _
        And IsBackTown And (HaveStoreItem Or HaveSellItem)) Or _
        (isBackBuy And NeedBuy) Or (NeedGetStorage And isBackStore)) Then
            Check_Destination_Route
        Else
            ReDim CurRoute(0)
        End If
    End If
    
    ptStart.X = curPos.Y
    ptStart.Y = (MapHeight - curPos.X)
    ReDim Route(0)
    Dim test As Boolean
    Dim Index As Integer
    'Do we find the solutions ?
    Index = get_solutions(MapName)
     'yes, just Routing to that Point
    If CurRoute(0).Pos.X <> 0 Then
        test = IsMovePoint(CurRoute(0).Pos)
        If Not test Then FindNearerPoint CurRoute(0).Pos
        ptEnd.X = CurRoute(0).Pos.X
        ptEnd.Y = MapHeight - CurRoute(0).Pos.Y
        DelaymoveCounter = 0
        If EvalNorm(ptStart, ptEnd) < 5 Then Exit Sub
    'Start Botting Random Routing
    ElseIf Index >= 0 Then
        test = IsMovePoint(MapRoute(Index).Src.Pos)
        If Not test Then FindNearerPoint MapRoute(Index).Src.Pos
        ptEnd.X = MapRoute(Index).Src.Pos.X
        ptEnd.Y = MapHeight - MapRoute(Index).Src.Pos.Y
        DelaymoveCounter = 0
        If EvalNorm(ptStart, ptEnd) < 5 Then Exit Sub
    'Just go to the NPC on the same map
    ElseIf (ptEnd.X = 0 Or ptEnd.Y = 0) Or EvalNorm(ptStart, ptEnd) < 10 Then
        Do
            RandomPoint = RandomNumber(UBound(AllowCoord), 0)
            ptEnd = AllowCoord(RandomPoint)
        Loop While (ptEnd.X = 0 Or ptEnd.Y = 0) Or (EvalNorm(ptStart, ptEnd) < 35)
        DelaymoveCounter = 0
    End If
    
    '---------- End ------------'
    
    If CanusePath And Not tmrProcess2.Enabled Then
        frmMain.tmrProcess2.Enabled = True
        FrmField.Run_Search
        frmMain.tmrProcess2.Enabled = False
        FrmField.LabFrm.Caption = "Map - Destination (" & _
        CStr(Route(UBound(Route)).Y) & ":" & CStr(Route(UBound(Route)).X) & ")"
        IsRandomRoute = True
        If Route(UBound(Route)).Y = 0 Then ptEnd.X = 0
    Else
        If Not IsDMove Then DirectionID = (DirectionID + 1) Mod 8
        move_to Move_By_Direction(curPos, DirectionID)
        IsDMove = False
        IsRandomRoute = True
    End If
    Exit Sub
errie:
Stat "Random:" & Err.Description & vbCrLf
IsRandommove = True
End Sub

Private Sub tmrNomons_Timer()
On Error GoTo errie
If ConnState >= 4 And (Not Pickup) Then
    'Exit Sub
    Check_Rest
    
    'If (Not CanFindMonster() Or Not IsAutoKill) And (NomonsWarp Or RandomMove) And _
    (Not Sitting) Then
    If (Not CanFindMonster) And NomonsWarp And Not Sitting Then
        NomonsTimeCount = NomonsTimeCount + 1
    ElseIf Not SellMode Then
        NomonsTimeCount = 0
        OnRoute = False
    End If
    
    If NomonsTimeCount > 100 Then
        BlockMove = False
        'PlayerMoveTime = 0
    End If
    
    If (AutoAI And (CheckNPC Or tmrDealNPC.Enabled)) Or ModAI Then Exit Sub
    If Not AutoAI Or AlwaySit Or MakeDamage Or Sitting Or IsSitting Then Exit Sub
    
    'If (Not MoveOnly) And (NomonsTimeCount > (NomonsTime * 5)) And NomonsWarp _
    'And (Not SellMode) And (Not DetectPortal) And (People(0).NameID = 0) And _
    'Not IsOnWayPoint(CurPos) And IsInLock And Not ModAI Then
    '    Stat "Can't find monster, Teleport..." & vbCrLf
    '    Teleport
    '    NomonsTimeCount = 0
    'End If
    If Not MoveOnly And IsInLock And (NomonsTimeCount > (NomonsTime * 5)) And NomonsWarp And _
    Not SellMode And Not DetectPortal And Not IsOnWayPoint(curPos) Then
        NomonsTimeCount = 0
        Stat "Can't find monster, Teleport..." & vbCrLf
        Teleport
    End If
    
    If UBound(Route) > 0 And Current < UBound(Route) Or Pickup Or ActionDelay > 0 Then Exit Sub
    
    
    If IsNearFightPortal(curPos) And FightMap And MoveOnly And FightMode And _
    StartPoint >= UBound(WayPoint) - 2 Then
        move_to FightPortal(indexFight)
        Exit Sub
    ElseIf IsNearSellPortal(curPos) And BackMap And SellMode And StartPoint = 0 Then
        move_to SellPortal(indexSell)
        Exit Sub
    ElseIf CheckNPC() And Not SellMode Then
        Exit Sub
    End If
    
    Check_BackTown
    
    If (NomonsTimeCount >= 0) And RandomMove And CurAtkMonster.NameID = 0 Then
        Dim test1 As Boolean
        Dim test2 As Boolean
        test1 = IsOnWayPoint(curPos)
        test2 = CanFindMonster And (Not MoveOnly)
        If UBound(AllowCoord) > 0 And AutoAI And (Not CanUseWP Or Not _
        test1) And UBound(Route) = 0 And CurAtkMonster.NameID = 0 And _
        (Not test2 Or Not IsAutoKill) Then
            Random_Move curPos
            BlockMove = False
            'NomonsTimeCount = 0
            Exit Sub
        End If
        If SellMode Then
            Direction = SellDirection
        End If
        If StartPoint = UBound(WayPoint) And Not MoveOnly And Not SellMode Then
            Direction = "BW"
        ElseIf StartPoint = 0 And Not MoveOnly And Not SellMode Then
            Direction = "FW"
        End If
        Dim dis As Integer
        dis = EvalNorm(curPos, WayPoint(StartPoint))
        If EvalNorm(curPos, WayPoint(StartPoint)) < 4 Then
            If StartPoint + 1 <= UBound(WayPoint) And Direction = "FW" Then
                    StartPoint = StartPoint + 1
                    move_to WayPoint(StartPoint)
            ElseIf StartPoint - 1 >= 0 And Direction = "BW" Then
                    StartPoint = StartPoint - 1
                    move_to WayPoint(StartPoint)
            End If
            OnRoute = True
            BackWpCounter = 0
        ElseIf EvalNorm(curPos, WayPoint(StartPoint)) >= 4 And BlockMove Then
            Exit Sub
        ElseIf EvalNorm(curPos, WayPoint(StartPoint)) < 20 And BackWpCounter < 200 Then
                'Stat "Back to Waypoint at (" & CStr(WayPoint(StartPoint).y) & ":" & CStr(WayPoint(StartPoint).x) & ")..." & vbCrLf
                'move_to WayPoint(StartPoint)
                If BackWpCounter < 70 Then
                    move_to WayPoint(StartPoint)
                Else
                    ptStart.X = curPos.Y
                    ptStart.Y = MapHeight - curPos.X
                    ptEnd.X = WayPoint(StartPoint).Y
                    ptEnd.Y = (MapHeight - WayPoint(StartPoint).X)
                    If CanusePath And Not tmrProcess2.Enabled Then FrmField.Run_Search
                End If
        '        BackWP = True
                BackWpCounter = BackWpCounter + 1
        ElseIf (BackWpCounter > 200) And (Not MoveOnly) Then
                If BackWpCounter = 30 Then Stat "Can't reach closest waypoint... " & vbCrLf
                'If BackWpCounter > 10 Then ResettoReCon
                If BackWpCounter > 200 Then
                    Teleport
                ElseIf BackWpCounter < 200 And Set_Start_Point(curPos) Then
                    move_to WayPoint(StartPoint)
                End If
                BackWpCounter = BackWpCounter + 1
                'Stat "Can't reach closest waypoint, random move..." & vbCrLf
        ElseIf (EvalNorm(curPos, WayPoint(StartPoint)) >= 18 And EvalNorm(curPos, WayPoint(0)) > 0 And (Not IsOnWayPoint(curPos))) Then
                If Find_Near_Point(curPos) Then
                    'Stat "Back to Waypoint at (" & CStr(WayPoint(StartPoint).y) & ":" & CStr(WayPoint(StartPoint).x) & ")..." & vbCrLf
                    move_to WayPoint(StartPoint)
                    BackWP = True
                ElseIf SellMode And IsBackTown And UBound(WayPoint) > 0 Then
                    Stat "Can't find waypoint, Teleport..." & vbCrLf
                    Teleport
                'ElseIf EvalNorm(CurPos, WayPoint(StartPoint)) < 22 And Not DetectPortal Then
                '    Stat "Can't find waypoint, random move..." & vbCrLf
                '    Random_Move WayPoint(StartPoint)
                ElseIf Not DetectPortal And Not IsOnWayPoint(curPos) Then
                    'Stat "Can't find closest waypoint, random move..." & vbCrLf
                    'Random_Move CurPos
                End If
        Else
            If Find_Near_Point(curPos) Then
                    'Stat "Back to Waypoint at (" & CStr(WayPoint(StartPoint).y) & ":" & CStr(WayPoint(StartPoint).x) & ")..." & vbCrLf
                    move_to WayPoint(StartPoint)
            Else
                move_to WayPoint(StartPoint)
            End If
        End If
        'NomonsTimeCount = 0
        'BlockMove = True
    End If
Else
NomonsTimeCount = 0
End If
Exit Sub
errie:
Err.Clear
'ClearAll
End Sub

Private Sub tmrPickDelay_Timer()
    tmrPickDelay.Enabled = False
End Sub

Private Sub tmrPickup_Timer()
On Error GoTo errie
Dim X As Integer

If (ConnState < 4) Then Exit Sub

If CurrentItem.Name = "" Then
    EstimateClosestItem
ElseIf CurrentItem.Name <> "" Then
    SendPickup
End If

Exit Sub
errie:
Err.Clear
'ClearAll
End Sub

Private Sub Check_Rest()
    On Error GoTo errie
    Dim tmpsp As Double
    Dim CheckWeight As Double
    If (ConnState < 4) Or ((AlwaySit Or Pickup Or CurAtkMonster.NameID > 0 Or UBound(Aggro) > 0) And Not Sitting) Then Exit Sub
    If Players(number).MaxWeight > 0 Then CheckWeight = Players(number).Weight / Players(number).MaxWeight
    If CheckWeight >= 0.5 Then Exit Sub
    If (SPSit < 0.25) Then
        tmpsp = SPSit
    Else
        tmpsp = 0.25
    End If
    If RestDelay > 0 Or StandDelay > 0 Or DelayCheckRest > 0 Then Exit Sub
    
    If NumberMons < 0 Then NumberMons = 0
    
    
    If (CurAtkMonster.NameID = 0) And (Not Sitting) And (Players(number).SP < (Players(number).maxsp * tmpsp)) And (IsSPSit) And (Not MakeDamage) And (Not Pickup) Then
        'Sitting = True
        If SkillChar(0).MaxLV >= 3 Then
            'Send_Sit
            RestDelay = RandomNumber(3, 2)
        Else
            Sitting = True
        End If
        If Not IsSitting Then Stat "SP Restoring, Sitting down(Delay)... " + vbCrLf
        IsSitting = True
        IsStanding = False
    ElseIf (CurAtkMonster.NameID = 0) And (Not Sitting) And (Players(number).HP < (Players(number).MaxHP * HPSit)) And (IsAutorest) And (Not MakeDamage) And (Not Pickup) Then
        'Sitting = True
        If SkillChar(0).MaxLV >= 3 Then
            'Send_Sit
            RestDelay = RandomNumber(5, 3)
        Else
            Sitting = True
        End If
        If Not IsSitting Then Stat "HP Restoring, Sitting down(Delay)..." + vbCrLf
        IsSitting = True
        IsStanding = False
    ElseIf (Sitting) And (Players(number).SP >= (Players(number).maxsp * SPWait)) And IsSPWait And _
    ((Not IsHPWait) Or (Players(number).HP >= (Players(number).MaxHP * HPWait))) Then
        'Sitting = False
        If SkillChar(0).MaxLV >= 3 Then
            'Send_Stand
            StandDelay = RandomNumber(5, 3)
        Else
            Sitting = False
        End If
        If Not ModAI Then
            If Not IsStanding Then Stat "HP/SP restored, Stand up(Delay)..." + vbCrLf
            IsSitting = False
            IsStanding = True
        End If
    ElseIf (Sitting) And (Players(number).HP >= (Players(number).MaxHP * HPWait)) And IsHPWait And _
    ((Not IsSPWait) Or (Players(number).SP >= (Players(number).maxsp * SPWait))) Then
        If SkillChar(0).MaxLV >= 3 Then
            'Send_Stand
            StandDelay = RandomNumber(5, 3)
        Else
            Sitting = False
        End If
        If Not ModAI Then
            If Not IsStanding Then Stat "HP/SP restored, Stand up(Delay)..." + vbCrLf
            IsSitting = False
            IsStanding = True
        End If
    ElseIf IsNomonsSit And (Not CanFindMonster) And (Not Sitting) And ((Players(number).HP < (Players(number).MaxHP * HPWait) And IsHPWait)) And (IsAutorest) And (Not MakeDamage) And (Not Pickup) Then
        If SkillChar(0).MaxLV >= 3 Then
            'Send_Sit
            RestDelay = RandomNumber(5, 3)
        Else
            Sitting = True
        End If
        If Not IsSitting Then Stat "No monster, restoring HP for a while(Delay)..." + vbCrLf
        IsSitting = True
        IsStanding = False
    ElseIf IsNomonsSit And (Not CanFindMonster) And (Not Sitting) And ((Players(number).SP < (Players(number).maxsp * tmpsp) And (IsSPWait))) And (Not MakeDamage) And (Not Pickup) Then
        If SkillChar(0).MaxLV >= 3 Then
            'Send_Sit
            RestDelay = RandomNumber(5, 3)
        Else
            Sitting = True
        End If
        If Not IsSitting Then Stat "No monster, restoring SP for a while(Delay)..." + vbCrLf
        IsSitting = True
        IsStanding = False
    ElseIf UBound(MonsterList) > 0 And (Sitting) And ((Not IsAutorest) Or (Players(number).HP >= (Players(number).MaxHP * HPWait))) And ((Not IsSPWait) Or (Players(number).SP >= (Players(number).maxsp * SPWait))) Then
        If SkillChar(0).MaxLV >= 3 Then
            'Send_Stand
            StandDelay = RandomNumber(5, 3)
        Else
            Sitting = False
        End If
        If Not ModAI Then
            If Not IsStanding Then Stat "Found Monster, Stand up..." + vbCrLf
            IsSitting = False
            IsStanding = True
        End If
    Else
        IsSitting = False
        IsStanding = False
    End If
    'If (Not Sitting) And Dead Then
    '    winsock1.close
    '    Recvdata = ""
    '    ClearAll
    '    ConnState = 0
    '    Stat "You're at save point, Disconnect..." & vbCrLf
    'End If
    Exit Sub
errie:
Err.Clear
'ClearAll
End Sub

Private Sub Use_SkillHeal()
    If (ConnState < 4) Then Exit Sub
    Winsock_SendPacket Chr(&H13) + Chr(1) + Chr(HealLV) + Chr(0) + Chr(&H1C) + _
    Chr(0) + AccountID, True
End Sub

Private Sub tmrPortal_Timer()
PortalTime = PortalTime + 1
If MoveOnly Then
    PortalTime = 0
    tmrPortal.Enabled = False
ElseIf (PortalTime > 8) Then
    Stat "To avoid map change, Teleport Away..." & vbCrLf
    Teleport
    PortalTime = 0
    tmrPortal.Enabled = False
End If
End Sub

Private Sub TmrProcess_Timer()
    If Not isUseHaunted Then
        tmrProcess.Enabled = False
        Exit Sub
    End If
    Dim X As Long
    Dim i As Integer
    Dim Thread As String
    Thread = LCase(ProcessName)
    X = GetProcessByName(Thread)
    txtStatus.text = "Waiting for " & Thread & " to run..." & vbCrLf
    If X <> ProcessID And X <> 0 Then
        Stat "Found " & Thread & "..." & vbCrLf
        Dim hProcess As Long
        hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, X)
        i = InjectLibrary(hProcess, App.Path & "\Inject.dll")
        If i <> 0 Then
            Stat "Grabbed " & Thread & "..." & vbCrLf
            ProcessID = X
            Winsock1.Close
            Winsock1.LocalPort = 2350
            Winsock1.Listen
            tmrProcess.Enabled = False
        End If
    End If
End Sub

Private Sub tmrRecon_Timer()
    If MODDelay.DualLogin > 0 Then Exit Sub
    Reconcount = Reconcount + 1
    tmrResponse.Enabled = False
    If (Not isWarp) Then
        Label1.Caption = "Main Status (" + CStr(DelayTime - Reconcount) + " s. to reconnect)"
    Else
        Label1.Caption = "Main Status (" + CStr(WarpDelay - Reconcount) + " s. to reconnect)"
    End If
    If ((Reconcount >= DelayTime) And (Not isWarp)) Or (isWarp And (Reconcount >= WarpDelay)) Then
        tmrRecon.Enabled = False
        Reconcount = 0
        Label1.Caption = "Main Status"
        txtStatus.text = ""
        tmrAnswer_Timer
    End If
End Sub

Private Sub TmrRef_Timer()
    If ConnState < 4 Then Exit Sub
    Dim TimeRef As Long
    Dim X As Integer
    If CurAtkMonster.NameID > 0 Then
        If EvalNorm(curPos, CurAtkMonster.Pos) > 20 Then
            Clear_This_Mons 0
            Stat "Can't reach to target!" + vbCrLf
        End If
        If Not isKillmob And EvalNorm(oldSelectPos, CurAtkMonster.Pos) > 13 Then
            If UBound(MonsterList) > 0 Then
            For X = 0 To UBound(MonsterList) - 1
                If (MonsterList(X).ID = CurAtkMonster.ID) Then
                    MonsterList(X).IsFollow = True
                    'Exit For
                End If
                If EvalNorm(MonsterList(X).Pos, CurAtkMonster.Pos) < 5 Then
                    MonsterList(X).IsFollow = True
                End If
            Next
            End If
            Clear_This_Mons 0
            oldSelectPos = curPos
            Stat "This monster follow another player!" + vbCrLf
        End If
    End If
    
    Update_MonsterPos
    Update_PeoplePos
    Update_CurPos
    Check_Rest
    Check_ActionTime
    Check_ActionSkill
    
    If AutoAI And CheckNPC Then Exit Sub
    If AutoAI And Not ChkAtk Then GoTo next_case
    If AutoAI And (CurAtkMonster.NameID > 0 Or (UBound(Aggro) > 0 And ChkAtk)) And _
      (IsAutoKill And CanFindMonster()) Then
        If UBound(Route) > 0 Then
            ReDim Route(0)
            PlayerMoveTime = 0
            PlayerMoveTime = 0
            BlockMove = False
        End If
        If (CurAtkMonster.NameID = 0 Or IsRandomRoute) And (Not MoveOnly Or IsRandomRoute) And IsAutoKill And ChkAtk Then
            EstimateClosestMonster
            Exit Sub
        End If
        GoTo next_case
    End If
    If Pickup Or IsSitting Then Exit Sub
next_case:
    If DetectPortal And AutoAI Then Exit Sub
    Check_Route
End Sub

Private Sub tmrResponse_Timer()
On Error GoTo errie
    If IsRouting Then Exit Sub
    If MODDelay.DualLogin > 0 Then Exit Sub
    ResponseCounter = ResponseCounter + 1
    If ((ResponseCounter >= ResponseTime) And (ConnState = 4)) Or ((ResponseCounter >= 13) And _
    (ConnState < 4) And ConnState > 0) And Not ModAI And _
    ((Not IsVending Or Not IsWaitShop) Or (Not IsChatOC Or Not IsWaitChat)) Then
        Stat "Time-Out" + vbCrLf
        ResettoReCon
    End If
    Exit Sub
errie:
    Err.Clear
End Sub

Public Sub tmrResetResponse()
    ResponseCounter = 0
    tmrResponse.Enabled = True
End Sub

Private Sub tmrSession_Timer()
On Error Resume Next
    If (ConnState > 3) And MODDelay.DualLogin = 0 Then
        SSec = SSec + 1
        If (SSec > 59) Then
            SMin = SMin + 1
            SSec = 0
            If (SMin > 59) Then
                SHour = SHour + 1
                SMin = 0
            End If
        End If
    End If
    If DelaySelfSkill > 0 Then DelaySelfSkill = DelaySelfSkill - 1
    'If MyPet.AutoFeed And MyPet.DelayFeed > 0 Then
    '    MyPet.DelayFeed = MyPet.DelayFeed - 1
    'End If
    
    If delay_count > 0 Then delay_count = delay_count - 1
    
    If WarpSaveCount > 0 Then
        WarpSaveCount = WarpSaveCount - 1
        If WarpSaveCount = 0 Then Winsock_SendPacket Chr(&H13) & Chr(1) & IntToChr(3) & IntToChr(&H1A) & AccountID, True
    End If
    
    If DelayCheckRest > 0 Then DelayCheckRest = DelayCheckRest - 1
    If TeleportDelay > 0 Then TeleportDelay = TeleportDelay - 1
    If ActionDelay > 0 Then ActionDelay = ActionDelay - 1
    If RestDelay > 0 Then
        RestDelay = RestDelay - 1
        If RestDelay = 0 Then Send_Sit
    End If
    If StandDelay > 0 Then
        StandDelay = StandDelay - 1
        If StandDelay = 0 Then Send_Stand
    End If
    'If SSec Mod 3 = 0 And MyPet.Name <> "" Then Winsock_SendPacket Chr(&HA1) & Chr(1) & Chr(0), True
    'If SSec Mod 10 = 0 And MyPet.Name <> "" And MyPet.Status > 40 Then PetPerformance
    SessionTime = MakeTime
    'On Error Resume Next
    Dim STimeCount As Double
    STimeCount = (SHour * 3600) + (SMin * 60) + SSec
    MDIfrmMain.StatusBar1.SimpleText = "Session 'Time: [" & SessionTime & "], EXP/JXP: [" & FormatNumber(SessionEXP, 0, vbTrue) & "/" & FormatNumber(SessionJEXP, 0, vbTrue) & "]" & ", ZENY: [" & FormatNumber(Players(number).Zeny - StartZeny, 0, vbTrue, vbUseDefault, vbTrue) & "]'"
    Dim SLvUp As Double, SJLvUp As Double
    If SessionEXP > 0 Then SLvUp = ((Players(number).NextBaseEXP - Players(number).BaseExp) * STimeCount) / SessionEXP
    If SessionJEXP > 0 Then SJLvUp = ((Players(number).MaxJobEXP - Players(number).JobExp) * STimeCount) / SessionJEXP
    Dim msgbar$
    msgbar = "Exp,JExp/Hr: [" & FormatNumber((SessionEXP / STimeCount * 3600), 0, vbTrue) & "/" & FormatNumber((SessionJEXP / STimeCount * 3600), 0, vbTrue) & "], Next Lv/JLv: [" & ETA(SLvUp) & "/" & ETA(SJLvUp) & "]"
    If DeadMonsName <> "" Then msgbar = msgbar & ", Last Monster : '" & DeadMonsName & "' [" & FormatNumber(CurEXPMons, 0, vbTrue) & "/" & FormatNumber(CurJXPMons, 0, vbTrue) & "]"
    MDIfrmMain.bState.Caption = msgbar
    Err.Clear
End Sub

Private Sub tmrSkillDelay_Timer()
On Error Resume Next
    SkillDelay = SkillDelay + 1
    If (SkillDelay >= 2) Then
        SkillWait = False
        SkillDelay = 0
        tmrSkillDelay.Enabled = False
    End If
    Err.Clear
End Sub

Private Sub tmrTicks_Timer()
On Error GoTo errie
If ConnState = 0 Then
    ConnState = 4
    Stat "Loading Map..." & vbCrLf
    If StartZeny < 1 Then StartZeny = Players(number).Zeny
    If Not isUseHaunted Then
        Winsock_SendPacket IntToChr(&H7D), True
        Winsock_SendPacket IntToChr(&H7E) & MakeTickString, True
        Winsock_SendPacket Chr(&H8A) & Chr(1), True
    End If
ElseIf ConnState = 3 Then
    If StartZeny < 1 Then StartZeny = Players(number).Zeny
    ConnState = 4
    Stat "Loading Map ..." & vbCrLf
    Connected = True
    frmPlayer.Visible = True
    MDIfrmMain.mnuWin.Visible = True
    MDIfrmMain.mnuAction.Visible = True
    MDIfrmMain.mnuRecon.Visible = True
    If Not isUseHaunted Then Winsock_SendPacket IntToChr(&H7D), True
    OldPos.X = 0
    OldPos.Y = 0
    AI_AVCheck
    ReSetCounter = 0
ElseIf ConnState = 4 Then

    If IsWantHeal And (Players(number).HP < (Players(number).MaxHP * HPSit)) And AutoAI Then
        Winsock_SendPacket IntToChr(&H96) & _
        IntToChr(Len("HEAL " & AccountID) + 29) & AcoHealName & _
        String(24 - Len(AcoHealName), Chr(0)) & "HEAL " & AccountID & Chr(0), True
    End If
    
    If AutoAI And Not Sitting And CurAtkMonster.NameID = 0 And (OldPos.X = curPos.X) And (OldPos.Y = curPos.Y) Then
        noMoveCounter = noMoveCounter + 1
    Else
        noMoveCounter = 0
    End If

    If noMoveCounter Mod 40 = 35 And Not ModAI And Not IsRouting And MODDelay.DualLogin = 0 Then
        If TeleNothing Then
            Stat "Bot do nothing, Teleport..." & vbCrLf
            Teleport
        Else
            Stat "Bot do nothing, Reconnect..." & vbCrLf
            ResettoReCon
        End If
    End If
    
    If (Autoheal) And (UseHeal) And (Players(number).HP < (Players(number).MaxHP * HPHeal)) And AutoAI Then
        SendHeal = True
        tmrMisc.Enabled = True
    ElseIf ((IsAutoRedz) Or (IsAutoOrange)) And UBound(AllInv) > 0 And AutoAI Then
        SendUsePot = True
        tmrMisc.Enabled = True
    End If
    
    If SPItem.Use And UBound(AllInv) > 0 And AutoAI Then
        SendSP = True
        tmrMisc.Enabled = True
    End If
    
    ClearCounter = ClearCounter + 1
    If ClearCounter Mod 24 = 0 And ClearCounter > 0 Then
        Winsock_SendPacket IntToChr(&H7E) + MakeTickString, True
    End If
    If ClearCounter >= 62 Then ClearCounter = 0
   ' If Not IsAggro Then Check_Rest
    
    If (CurAtkMonster.NameID > 0) And (OldPos.X = curPos.X) And (OldPos.Y = curPos.Y) Then
        ReSetCounter = ReSetCounter + 1
    Else
        ReSetCounter = 0
    End If
    
    If ReSetCounter Mod 7 = 3 And ((CurAtkMonster.NameID > 0)) Then
        'MakeDamage = False
        'TraceMons = False
        BlockMove = False
        SendAttack
        'Winsock_SendPacket IntToChr(&H89) & CurAtkMonster.ID + Chr(7), True
    End If
    If ReSetCounter Mod 7 = 6 And ((CurAtkMonster.NameID > 0)) Then
        MakeDamage = False
        TraceMons = False
        upd_curMonster
        BlockMove = False
        SendAttack
    End If
    
    If (ReSetCounter > 20) Or ((ReSetCounter > giveuptime) And Not MakeDamage And (CurAtkMonster.NameID > 0) And Not IsAggro) Then
            Stat "Give up on this monster..." + vbCrLf
            Dim X As Integer
            If UBound(MonsterList) > 0 Then
            For X = 0 To UBound(MonsterList) - 1
                If (MonsterList(X).ID = CurAtkMonster.ID) Then
                    MonsterList(X).IsAttack = True
                End If
            Next
            End If
            Clear_This_Mons 0
            ReSetCounter = 0
    End If
        
    
    'If InFight And (CurAtkMonster(NumberMons).NameId = 0) Then InFight = False
    OldPos.X = curPos.X
    OldPos.Y = curPos.Y
    'If CurAtkMonster.NameID <> 0 Then SendAttack Else EstimateClosestMonster
    'If (CurAtkMonster(NumberMons).NameId > 0) And (Not MakeDamage) Then SendAttack
End If
Exit Sub
errie:
'Stat "Time Tick Error, waiting..." + vbCrLf
'ResettoReCon
'ClearAll
Err.Clear
End Sub


Private Sub tmrTime_Timer()
On Error Resume Next
    uTime = uTime + 1
    CheckEvent "OnEverySecond", "nothingtocheck=False"
    If StartBot Then
        Dim i&, X&
        If ((uTime Mod 3) = 0) And IsAutoSpirits And CurSpirit < BallSpirits And AutoAI And IsInLock Then
            X = Find_SkillId("CH_SOULCOLLECT")
            If X > 0 Then
                Send_Use_Skill SkillChar(X).ID, SkillChar(X).Lv, AccountID
            Else
                X = Find_SkillId("MO_CALLSPIRITS")
                If X > 0 Then Send_Use_Skill SkillChar(X).ID, SkillChar(X).Lv, AccountID
            End If
        End If
    End If
    If (GetTickCount > ((RestartTime * 60000) + StartTime)) And UseRestart Then
        Chat "Restarting . . . ", vbRed
        Shell "revemu-mc.exe", vbNormalFocus
        ForceExit
    End If
    If HauntedStep > 0 Then
        Dim MsgS$
        Select Case HauntedStep
            Case 5
                MsgS = "blueHello! You're using " & Version & "."
                SendClient Chr(&H9A) & Chr(0) & IntToChr(Len(MsgS) + 4) & MsgS
            Case 3
                MsgS = "blueThis program was written by iCeZ @ http://www.icez.net/"
                SendClient Chr(&H9A) & Chr(0) & IntToChr(Len(MsgS) + 4) & MsgS
            Case 1
                MsgS = "bluePlease visit http://www.icez.net for update details."
                SendClient Chr(&H9A) & Chr(0) & IntToChr(Len(MsgS) + 4) & MsgS
        End Select
        HauntedStep = HauntedStep - 1
    End If
    Err.Clear
End Sub

Private Sub Winsock1_Close()
DError = True
Stat "Socket closed.." & vbCrLf
If isUseHaunted Then tmrProcess.Enabled = True Else ResettoReCon
End Sub

Private Sub Winsock1_Connect()
On Error Resume Next
Stat "Connected!" & vbCrLf
If IsUseProxy Then
    ProxyConn = False
    allData = ""
    Select Case ProxyType
        Case 0 'https
            Stat "Requesting HTTPS proxy connection." & vbCrLf
            Winsock1.SendData "CONNECT " & CurConnIP & ":" & CurConnPort & " HTTP/1.0" & vbCrLf & vbCrLf
        Case 1 'socks4
            ProxyStep = 0
            Stat "Request connection with SOCKS4 proxy." & vbCrLf
            Winsock1.SendData Chr(4) & Chr(1) & HextoChr(ReverseHex(ChrtoHex(Mid(LngToChr(CLng(CurConnPort)), 1, 2)))) & IPToStr(CurConnIP) & "REVEMUMC" & Chr(0)
        Case 2 'socks5
            ProxyStep = 0
            If ProxyUser <> "0" Then
                Stat "Request authenticating with SOCKS5 proxy." & vbCrLf
                Winsock1.SendData Chr(5) & Chr(2) & Chr(0) & Chr(2)
            Else
                Stat "Request connection with SOCKS5 proxy." & vbCrLf
                Winsock1.SendData Chr(5) & Chr(1) & Chr(0)
            End If
    End Select
Else
    CheckAuth
End If
If Err.number > 0 Then print_funcerr "Winsock1_Connect", Err.number, Err.Description
Err.Clear
End Sub

Sub CheckAuth()
    If (WaitCheat = 1 Or WaitCheat = 3) And MasterSelect.pkserver = 0 Then
        Winsock_SendPacket Chr(&H72) & Chr(0) & AccountID & CharID & SessionID & MakeTickString & Chr(Sex), True
        Stat "Fake authenticating with map server..." & vbCrLf
        If WaitCheat = 1 Then WaitCheat = 2 Else WaitCheat = 0
    Else
        CheckAuth2
    End If
End Sub
Sub CheckAuth2()
On Error Resume Next
    Select Case ConnState
    Case 1
        If MasterSelect.IsLoginCrypt Then
            If MasterSelect.Encrypt = 3 Then
                Winsock_SendPacket HextoChr(MasterSelect.encRequest), True
            Else
                Winsock_SendPacket Chr(&HDB) & Chr(1), True
            End If
            Stat "Send request session key..." + vbCrLf
        Else
            Winsock_SendPacket Chr(&H64) & Chr(0) & Chr(MasterSelect.code) & Chr(0) & Chr(0) & Chr(0) & strUser & _
            String(24 - Len(strUser), Chr(0)) & StrPass & String(24 - Len(StrPass), Chr(0)) & Chr(MasterSelect.Version), True
            Stat "Verify your ID and Password..." + vbCrLf
        End If
        tmrResetResponse
        uTime = 0
    Case 2
        Winsock_SendPacket Chr(&H65) + Chr(0) + AccountID + SessionID + Chr(0) + Chr(0) + _
        Chr(0) + Chr(0) + Chr(0) + Chr(0) + Chr(Sex), True
        Stat "Verify your account..." + vbCrLf
        tmrResetResponse
        uTime = 0
    Case 3
        CryptOn = False
        If MasterSelect.pkserver = 1 Then
            Winsock_SendPacket Chr(&H72) + Chr(0) & Chr(0) & AccountID & Chr(0) & Chr(&H2C) & Chr(&HFC) & CharID & Chr(&H60) & Chr(0) & String(4, 255) & SessionID + _
            MakeTickString, True
            Stat "Authenticating PK Server..." + vbCrLf
        Else
            If MasterSelect.pkserver = 2 Then
                Winsock_SendPacket Chr(&H72) & Chr(0) & Chr(0) & Chr(0) & Chr(0) + AccountID & Chr(&HFA) & Chr(&H12) & Chr(0) & Chr(&H50) & Chr(&H83) + CharID & Chr(255) & Chr(255) + SessionID & MakeTickString & Chr(Sex), True
                Stat "Authenticating with map server..." + vbCrLf
            Else
                Winsock_SendPacket Chr(&H72) & Chr(0) & AccountID & CharID & SessionID & MakeTickString & Chr(Sex), True
                Stat "Verify player info..." & vbCrLf
            End If
        End If
        tmrResetResponse
        uTime = 0
    Case 0
        Winsock_SendPacket IntToChr(&H94) & AccountID, True
        Winsock_SendPacket Chr(&H64 + &HE) + Chr(0) + AccountID + CharID + SessionID + _
        MakeTickString + Chr(Sex), True
        Stat "Request Map Change..." + vbCrLf
        tmrResetResponse
        uTime = 0
    End Select
    If Err.number > 0 Then print_funcerr "CheckAuth", Err.number, Err.Description
    Err.Clear
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
If Winsock1.State <> 0 Then Winsock1.Close
Winsock1.Accept requestID
Stat "Ready to manual login from client now..." & vbCrLf
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error GoTo errie
tmrResetResponse
Dim tData As String
Winsock1.GetData tData

If IsUseProxy And ProxyConn = False And Not isUseHaunted Then
    allData = allData & tData
    Select Case ProxyType
        Case 0
            Stat tData
            If InStr(allData, vbCrLf & vbCrLf) > 0 Then
                Dim sspl() As String
                sspl = Split(tData, " ")
                If UBound(sspl) > 0 Then
                    If Val(sspl(1)) = 200 Then
                        Stat "Proxy connection successful." & vbCrLf
                        ProxyConn = True
                        CheckAuth
                    Else
                        Stat "Proxy connection error." & vbCrLf
                        Winsock1.Close
                        ResettoReCon
                    End If
                End If
            End If
        Case 1 'socks4
            If Left(tData, 1) = Chr(4) Then
                Select Case Asc(Mid(tData, 2, 1))
                    Case 90
                        Stat "SOCKS4 connection successful." & vbCrLf
                        ProxyConn = True
                        CheckAuth
                    Case Else
                        Stat "SOCKS4 Connection denied." & vbCrLf
                        Winsock1.Close
                        ResettoReCon
                End Select
            Else
                Stat "Error while talking to SOCKS4 Proxy." & vbCrLf
                Winsock1.Close
                ResettoReCon
            End If
            
        Case 2 'socks5
            Select Case ProxyStep
                Case 0
                    If Left(tData, 1) = Chr(5) Then
                        Select Case Asc(Mid(tData, 2, 1))
                            Case 0 'no authentication
                                ProxyStep = 2
                                Stat "Requesting a connection to server." & vbCrLf
                                Winsock1.SendData Chr(5) & Chr(1) & Chr(0) & Chr(1) & IPToStr(CurConnIP) & HextoChr(ReverseHex(ChrtoHex(Mid(LngToChr(CLng(CurConnPort)), 1, 2))))
                            Case 2 'normal authentication
                                Stat "Authenticating . . . " & vbCrLf
                                ProxyStep = 1
                                Winsock1.SendData Chr(1) & Chr(Len(ProxyUser)) & ProxyUser & Chr(Len(ProxyPass)) & ProxyPass
                            Case Else
                                Stat "Error while talking to SOCKS5 Proxy." & vbCrLf
                                Winsock1.Close
                                ResettoReCon
                        End Select
                    End If
                Case 1
                    If Mid(tData, 1, 1) = Chr(1) Then
                        If Mid(tData, 2, 1) = Chr(0) Then
                            ProxyStep = 2
                            Stat "Requesting a connection to server." & vbCrLf
                            Winsock1.SendData Chr(5) & Chr(1) & Chr(0) & Chr(1) & IPToStr(CurConnIP) & HextoChr(ReverseHex(ChrtoHex(Mid(LngToChr(CLng(CurConnPort)), 1, 2))))
                        Else
                            Stat "Error while authenticating to SOCKS5 Proxy." & vbCrLf
                            Winsock1.Close
                            ResettoReCon
                        End If
                    Else
                        Stat "Error while talking to SOCKS5 Proxy." & vbCrLf
                        Winsock1.Close
                        ResettoReCon
                    End If
                Case 2
                    If Mid(tData, 1, 1) = Chr(5) Then
                        Select Case Asc(Mid(tData, 2, 1))
                            Case 0
                                Stat "SOCKS5 connection successful." & vbCrLf
                                ProxyConn = True
                                CheckAuth
                            Case 1
                                Stat "SOCKS5 error: General SOCKS failure." & vbCrLf
                                Winsock1.Close
                                ResettoReCon
                            Case 2
                                Stat "SOCKS5 error: Connection not allowed." & vbCrLf
                                Winsock1.Close
                                ResettoReCon
                            Case 3
                                Stat "SOCKS5 error: Network unreachable." & vbCrLf
                                Winsock1.Close
                                ResettoReCon
                            Case 4
                                Stat "SOCKS5 error: Host unreachable." & vbCrLf
                                Winsock1.Close
                                ResettoReCon
                            Case 5
                                Stat "SOCKS5 error: Connection refused." & vbCrLf
                                Winsock1.Close
                                ResettoReCon
                            Case 6
                                Stat "SOCKS5 error: TTL expired." & vbCrLf
                                Winsock1.Close
                                ResettoReCon
                            Case 7
                                Stat "SOCKS5 error: Command not supported." & vbCrLf
                                Winsock1.Close
                                ResettoReCon
                            Case 8
                                Stat "SOCKS5 error: Address type not supported." & vbCrLf
                                Winsock1.Close
                                ResettoReCon
                            Case Else
                                Stat "SOCKS5 error: Unknown error." & vbCrLf
                                Winsock1.Close
                                ResettoReCon
                        End Select
                    End If
            End Select
    End Select
    Exit Sub
End If

If isUseHaunted And Len(tData) > 0 Then
    tData = ReArrangeData(tData)
    If Len(tData) = 0 Then Exit Sub
End If
'print_packet Recvdata
RecvData = RecvData & tData
If Len(RecvData) >= 4 Then
    If (Left(RecvData, 4)) = AccountID Then
        RecvData = Right(RecvData, Len(RecvData) - 4)
    End If
End If

If Len(RecvData) > 0 Then
    While Asc(Left(RecvData, 1)) = 0
        If Asc(Left(RecvData, 1)) = 0 Then RecvData = Mid(RecvData, 2, Len(RecvData) - 1)
        If Len(RecvData) = 0 Then GoTo endloop
    Wend
endloop:
End If


If Len(RecvData) >= 2 Then ParseData

Exit Sub
errie:
RecvData = ""
'If FrmField.Visible Then Load_Field MapName
FrmField.PicMap.Refresh
DError = True
tmrAnswer_Timer
End Sub

Private Sub Winsock1_Error(ByVal number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
DError = True
'If ConnState < 4 Then
'        Stat "Connection is aborted from Server.." & vbCrLf
'Else
        Stat ":" & Description & ".." & vbCrLf, vbRed, True
'End If
If isUseHaunted Then tmrProcess.Enabled = True Else ResettoReCon
End Sub

Public Function ReArrangeData(ByVal tData As String) As String
On Error GoTo errie
    Dim exportdata As String
    Dim Length As Long
    exportdata = ""
restart:
        Length = MakePort(Mid(tData, 2, 2))
        If Left(tData, 1) = Chr(&H53) Or Left(tData, 1) = Chr(&H4B) Then
            Winsock1.SendData Left(tData, Length + 3)
            If MDIfrmMain.mnuPKTLOG.CheckED And Length > 3 Then
                Open App.Path & "\packet.txt" For Append As #76
                    Print #76, "Send : "
                    Print #76, ConvPacketData(Mid(tData, 4, Length))
                    Print #76, ""
                Close #76
            End If
            If ConnState < 4 Then
                If Mid(tData, 4, 1) = Chr(&H66) Then
                    ConnState = 2
                    CharID = Asc(Mid(tData, 6, 1))
                    Stat "Character selected : " & Players(CharID).Name & " [" & Asc(Mid(tData, 6, 1)) & "]" & vbCrLf, vbBlue
                    number = CharID
                    UpdatePlayer
                ElseIf Mid(tData, 4, 1) = Chr(&H72) Then
                    ConnState = 3
                ElseIf Mid(tData, 4, 1) = Chr(&H7D) Then
                    ConnState = 4
                    HauntedStep = 6
                    Stat "Loading Map..." + vbCrLf
                    Connected = True
                    frmPlayer.Visible = True
                    MDIfrmMain.mnuWin.Visible = True
                    MDIfrmMain.mnuAction.Visible = True
                    MDIfrmMain.mnuRecon.Visible = True
                    'Winsock_SendPacket IntToChr(&H7D), True
                    OldPos.X = 0
                    OldPos.Y = 0
                    ReSetCounter = 0
                End If
            End If
        Else
            exportdata = exportdata & Mid(tData, 4, Length)
        End If
        tData = Right(tData, Len(tData) - (Length + 3))
        If Len(tData) > 0 Then GoTo restart
    ReArrangeData = exportdata
Exit Function
errie:
If Err.number > 0 Then print_funcerr "ReArrangeData", Err.number, Err.Description, "Length: " & Length & vbCrLf & "Data:" & ChrtoHex(tData)
Err.Clear
End Function

Sub EstimateClosestMonster()
On Error GoTo errie
Check_Rest
If Sitting Or IsSitting Or ActionDelay > 0 Then Exit Sub

If CurAtkMonster.NameID > 0 Then Exit Sub
If UBound(MonsterList) = 0 Then Exit Sub
If FollowMode.NoAttack And FollowMode.Active Then Exit Sub
If Not IsInLock Then Exit Sub
If Pickup Or (Not Check_Attack) Or SellMode Then Exit Sub ' Stat "1" & CStr(Pickup) & CStr(Not Check_Attack) & CStr(SellMode) & vbCrLf: Exit Sub
If BackWP Or (DetectPortal) Then Exit Sub ' Stat "2" & CStr(BackWP) & CStr(MoveOnly) & CStr(DetectPortal) & vbCrLf: Exit Sub
If (Not AutoAI) Or AlwaySit Or TmrDeal.Enabled Then Exit Sub ' Stat "3" & CStr(Not AutoAI) & CStr(AlwaySit) & CStr(TmrDeal.Enabled) & vbCrLf: Exit Sub
If ModAI Or Not ChkAtk Then Exit Sub ' Stat "4" & CStr(ModAI) & CStr(Not ChkAtk) & vbCrLf: Exit Sub

Dim i&, j&
Dim BestMons&, CurDist&, TmpDist&

BestMons = -1
CurDist = 20
Pickuptime = 0
TryPicktime = 0
Tracing = False

For i = 0 To UBound(MonsterList) - 1
    If MonsterList(i).NoAttack Then GoTo endloop
    If MonsterList(i).IsPet Then GoTo endloop
    If MonsterList(i).Time > 1 Then GoTo endloop
    If CloseAnyPlayer(MonsterList(i).Pos, MonsterList(i).NextPos) Then
        MonsterList(i).CantGo = True
        GoTo endloop
    End If
    If (MonsterList(i).IsFollow) And Not isKillmob Then GoTo endloop
    If (MonsterList(i).IsAttack) And Not killsteal Then GoTo endloop
    If Not CanAttackRoute(curPos, MonsterList(i).Pos) Then
        MonsterList(i).CantGo = True
        GoTo endloop
    End If
    
    TmpDist = EvalNorm(MonsterList(i).Pos, curPos)
    If TmpDist < CurDist Then
        CurDist = TmpDist
        BestMons = i
    End If
endloop:
Next

If BestMons < 0 Then Exit Sub

If Sitting And ((Not IsSPWait) Or (GetSP >= SPWait)) And ((Not IsAutorest) Or (GetHP >= HPWait)) Then
    Stat "You stand up..." + vbCrLf
    Winsock_SendPacket Chr(&H64 + &H25) + Chr(0) + Chr(0) + Chr(0) + Chr(0) + Chr(0) + Chr(3), True
End If
If CurDist > 10 And Not IsUseRange And Not isClassArcher Then
    move_to NearPos(MonsterList(BestMons).Pos, curPos, CurDist \ 2)
End If
If Return_MonsterName(MonsterList(BestMons).NameID) <> "" Then
    Aggro(0).ID = MonsterList(BestMons).ID
    CurAtkMonster = MonsterList(BestMons)
    CurMonsterName = Return_MonsterName(MonsterList(BestMons).NameID)
    Check_Equip CurMonsterName
    Check_Accessory CurMonsterName
    Plot_Dot MonsterList(BestMons).Pos, CurAtkColor
    SkillCounter = 0
    SpellCounter = 0
    Stat "Select [" + CurMonsterName + "] as a Target, Locking..." + vbCrLf
    oldSelectPos = CurAtkMonster.Pos
    NumberMons = 0
    Tracing = False
    SendAction = True
    NomonsTimeCount = 0
    InFight = False
    MakeDamage = False
    IsAggro = False
    IsDamage = False
    DamageCounter = 0
    AttackCounter = 0
    upd_curMonster
    SendAttack
End If

Exit Sub
errie:
If Err.number > 0 Then print_funcerr "EstimateClosestMonster", Err.number, Err.Description
Err.Clear
'ClearAll
End Sub


Public Sub Update_frmArmor(Itemname As String, code As Long)
    If code Mod 2 = 1 Then frmArmor.labHead.Caption = Itemname
    If code Mod 4 > 1 Then frmArmor.labRH.Caption = Itemname
    If code Mod 8 > 3 Then frmArmor.labRobe.Caption = Itemname
    If code Mod 16 > 7 Then frmArmor.labAcc1.Caption = Itemname
    If code Mod 32 > 15 Then frmArmor.labArmor.Caption = Itemname
    If code Mod 64 > 31 Then frmArmor.labLH.Caption = Itemname
    If code Mod 128 > 63 Then frmArmor.labShoes.Caption = Itemname
    If code Mod 256 > 127 Then frmArmor.labAcc2.Caption = Itemname
    If code Mod 512 > 255 Then frmArmor.labHead0.Caption = Itemname
    If code Mod 1024 > 511 Then frmArmor.LabHead1.Caption = Itemname
    If code Mod 2048 > 1023 Then frmArmor.labHead.Caption = Itemname
End Sub

Public Sub Equip_Item()
    With frmItem.lstInvent
        If .List(.ListIndex) <> "" Then
             Winsock_SendPacket IntToChr(&HA9) & IntToChr(Val(.List(.ListIndex))) & AllInv(Val(.List(.ListIndex))).Type, True
        End If
    End With
End Sub

Public Sub unEquip_Item()
    With frmItem.lstInvent
        If .List(.ListIndex) <> "" Then
             Winsock_SendPacket IntToChr(&HAB) & IntToChr(Val(.List(.ListIndex))), True
        End If
    End With
End Sub

Sub Send_Equip(ItemID As Integer)
    Winsock_SendPacket IntToChr(&HA9) & IntToChr(CLng(ItemID)) & AllInv(ItemID).Type, True
End Sub
Sub Send_unEquip(ItemID As Integer)
    Winsock_SendPacket IntToChr(&HAB) & IntToChr(CLng(ItemID)), True
End Sub

Public Sub Use_Item()
    With frmItem.lstInvent
        If .List(.ListIndex) <> "" Then
            Winsock_SendPacket IntToChr(&HA7) & IntToChr(Val(.List(.ListIndex))) & AccountID, True
        End If
    End With
End Sub

Public Sub Drop_Item()
On Error GoTo errie
    With frmItem.lstInvent
        If .List(.ListIndex) <> "" Then
            Winsock_SendPacket IntToChr(&HA2) & IntToChr(Val(.List(.ListIndex))) & IntToChr(Val(InputBox("Quantity?", "Enter number to Drop"))), True
        End If
    End With
Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "Drop_Item", Err.number, Err.Description
    Err.Clear
End Sub

Public Sub UpdateStats()
On Error Resume Next
If Connected Then
frmStat.labStr.Caption = CStr(Players(number).STR) & " + " & CStr(Players(number).Strp)
frmStat.labAgi.Caption = CStr(Players(number).AGI) & " + " & CStr(Players(number).Agip)
frmStat.labVit.Caption = CStr(Players(number).VIT) & " + " & CStr(Players(number).Vitp)
frmStat.labInt.Caption = CStr(Players(number).Intl) & " + " & CStr(Players(number).Intp)
frmStat.labDex.Caption = CStr(Players(number).DEX) & " + " & CStr(Players(number).Dexp)
frmStat.labLuk.Caption = CStr(Players(number).LUK) & " + " & CStr(Players(number).Lukp)
frmStat.labAtk.Caption = CStr(Players(number).ATK) & " + " & CStr(Players(number).ATKp)
frmStat.labMatk.Caption = CStr(Players(number).MinMatk) & " ~ " & CStr(Players(number).MaxMatk)
frmStat.labHit.Caption = CStr(Players(number).Hit)
frmStat.labCri.Caption = CStr(Players(number).Crit)
frmStat.labDef.Caption = CStr(Players(number).Def) & " + " & CStr(Players(number).Defp)
frmStat.labMdef.Caption = CStr(Players(number).mDef) & " ~ " & CStr(Players(number).mDefp)
frmStat.labAspd.Caption = CStr(Players(number).Aspd)
frmStat.labFlee.Caption = CStr(Players(number).Flee) & " + " & CStr(Players(number).Fleep)
If Val(frmStat.labStatPt.Caption) >= Val(frmStat.LabAgip) Then
    frmStat.imgUpAgi.Visible = True
Else
    frmStat.imgUpAgi.Visible = False
End If
If Val(frmStat.labStatPt.Caption) >= Val(frmStat.LabDexp) Then
    frmStat.imgUpDex.Visible = True
Else
    frmStat.imgUpDex.Visible = False
End If
If Val(frmStat.labStatPt.Caption) >= Val(frmStat.LabIntp) Then
    frmStat.imgUpInt.Visible = True
Else
    frmStat.imgUpInt.Visible = False
End If
If Val(frmStat.labStatPt.Caption) >= Val(frmStat.LabLuckp) Then
    frmStat.imgUpLuk.Visible = True
Else
     frmStat.imgUpLuk.Visible = False
End If
If Val(frmStat.labStatPt.Caption) >= Val(frmStat.labStrp) Then
    frmStat.ImgUpStr.Visible = True
Else
    frmStat.ImgUpStr.Visible = False
End If
If Val(frmStat.labStatPt.Caption) >= Val(frmStat.LabVitp) Then
    frmStat.imgUpVit.Visible = True
Else
    frmStat.imgUpVit.Visible = False
End If
End If
End Sub

Public Sub UpdateSkills()
On Error GoTo errie
frmSkill.lstSkill.Clear
'If Not Connected Then Exit Sub
Dim X As Integer
For X = 0 To UBound(SkillChar) - 1
    frmSkill.lstSkill.AddItem MakeHexName(IntToChr(CLng(SkillChar(X).ID))) + " : " + Get_SkillName(SkillChar(X).Name) + " LV." + CStr(SkillChar(X).MaxLV) + " " & IIf(SkillChar(X).SP > 0, "(" + CStr(SkillChar(X).SP) + " SP)", "") & " [" & GetSkillTargetType(SkillChar(X).Target) & "]"
Next
If (UBound(SkillChar) > 0) Then frmSkill.lstSkill.Selected(SkillSelect) = True
Exit Sub
errie:
Err.Clear
'ClearAll
End Sub

Function GetSkillTargetType(InVal As Integer) As String
On Error Resume Next
    Select Case InVal
        Case 0: GetSkillTargetType = "Passive"
        Case 1: GetSkillTargetType = "Enemy"
        Case 2: GetSkillTargetType = "Land skill"
        Case 4: GetSkillTargetType = "Self skill"
        Case 8: GetSkillTargetType = "Weapon"
        Case 16: GetSkillTargetType = "People / Monster"
        Case Else: GetSkillTargetType = "U:" & InVal
    End Select
End Function

Public Sub UpdatePlayer()
On Error GoTo errie
Dim percent As Double
If Not Connected Then Exit Sub
If (Players(number).NextBaseEXP > 0) And (Players(number).maxsp > 0) And (Players(number).MaxHP > 0) And (Players(number).MaxJobEXP > 0) And (Players(number).HP >= 0) And (Players(number).SP >= 0) And (Players(number).BaseExp >= 0) And (Players(number).JobExp >= 0) Then
    GetScriptLockmap
    frmPlayer.labBaseLv.Caption = CStr(Players(number).BaseLV)
    frmPlayer.labJobLv.Caption = CStr(Players(number).JobLV)
    frmPlayer.labPlayerName.Caption = Players(number).Name
    percent = (Players(number).BaseExp / Players(number).NextBaseEXP)
    frmPlayer.tabBaseEXP.width = percent * (frmPlayer.labtabBaseEXPBg.width - 25)
    frmPlayer.labtabBaseEXPBg.ToolTipText = CStr(Players(number).BaseExp) + "/" + CStr(Players(number).NextBaseEXP) + " (" & FormatNumber(percent, 2, vbTrue) + "%) "
    percent = (Players(number).JobExp / Players(number).MaxJobEXP)
    frmPlayer.tabJobEXP.width = percent * (frmPlayer.labtabJobExpBg.width - 25)
    frmPlayer.labtabJobExpBg.ToolTipText = CStr(Players(number).JobExp) + "/" + CStr(Players(number).MaxJobEXP) + " (" & FormatNumber(percent, 2, vbTrue) + "%) "
    frmPlayer.labZeny.Caption = FormatNumber(Players(number).Zeny, 0, vbTrue, vbTrue, vbTrue)
    frmPlayer.labSP.Caption = CStr(Players(number).SP) + "  /  " + CStr(Players(number).maxsp)
    frmPlayer.labSP.Caption = CStr(Players(number).SP) + "  /  " + CStr(Players(number).maxsp)
    frmPlayer.tabSP.width = (Players(number).SP / Players(number).maxsp) * (frmPlayer.tabSPbg.width - 20)
    frmPlayer.LabHP.Caption = CStr(Players(number).HP) + "  /  " + CStr(Players(number).MaxHP)
    frmPlayer.LabHP.Caption = CStr(Players(number).HP) + "  /  " + CStr(Players(number).MaxHP)
    frmPlayer.tabHP.width = (Players(number).HP / Players(number).MaxHP) * (frmPlayer.tabHPBg.width - 20)
    frmPlayer.labWeight.Caption = CStr(Players(number).Weight) + " / " + CStr(Players(number).MaxWeight)
    frmPlayer.labWeight.ToolTipText = FormatNumber((CLng(Players(number).Weight) * 100) / Players(number).MaxWeight, 2, vbTrue)
End If
Exit Sub
errie:
Err.Clear
'ClearAll
End Sub

Private Sub Auto_SP()
On Error GoTo errie
If (ConnState < 4) Then Exit Sub
Dim percent As Double
Dim X As Long
With SPItem
percent = Players(number).SP / Players(number).maxsp
X = 0
If (percent < .percent) Then X = Find_HealItem(.Name)
If (X > 0) And .Use Then
    Stat "Use " + AllInv(X).Name + vbCrLf
    Winsock_SendPacket IntToChr(&HA7) + IntToChr(X) + AccountID, True
End If
End With
Exit Sub
errie:
Err.Clear
'ClearAll
End Sub

Private Sub Auto_Use_Pots()
On Error GoTo errie
If (ConnState < 4) Then Exit Sub
Dim percent As Double
Dim X As Long
Dim Y As Long
percent = Players(number).HP / Players(number).MaxHP
X = 0
Y = 0
If (percent < HPOrange) Then X = Find_HealItem(healitem2)
If (percent < HPRed) And X = 0 Then Y = Find_HealItem(healitem1)
If (X > 0) And IsAutoOrange Then
    Stat "Use " + AllInv(X).Name + vbCrLf
    If Sitting Then Send_Stand
    Winsock_SendPacket IntToChr(&HA7) & IntToChr(X) & AccountID, True
End If
If (Y > 0) And (IsAutoRedz) Then
    Stat "Use " + AllInv(Y).Name + vbCrLf
    If Sitting Then Send_Stand
    Winsock_SendPacket IntToChr(&HA7) & IntToChr(Y) & AccountID, True
End If
      
Exit Sub
errie:
Err.Clear
'ClearAll
End Sub

Public Sub UseRedz()
On Error GoTo errie
If (ConnState < 4) Then Exit Sub
    Dim X As Long
    X = Find_HealItem(healitem2)
    If (X > 0) Then
        Stat "Use " + AllInv(X).Name + vbCrLf
        Winsock_SendPacket IntToChr(&HA7) & IntToChr(X) & AccountID, True
    Else
        Stat "Can't find item in the list..." & vbCrLf
    End If
Exit Sub
errie:
Err.Clear
'ClearAll
End Sub

Public Sub PetFeed()
On Error GoTo errie
If (ConnState < 4) Then Exit Sub
Winsock_SendPacket Chr(&HA1) & Chr(1) & Chr(1), True
Exit Sub
errie:
Err.Clear
'ClearAll
End Sub

Public Sub BackEgg()
On Error GoTo errie
If (ConnState < 4) Then Exit Sub
'Stat "Pet's Performance!" & vbCrLf
Winsock_SendPacket Chr(&HA1) & Chr(1) & Chr(3), True
Exit Sub
errie:
Err.Clear
'ClearAll
End Sub

Public Sub PetPerformance()
On Error GoTo errie
If (ConnState < 4) Then Exit Sub
'Stat "Pet's Performance!" & vbCrLf
Winsock_SendPacket Chr(&HA1) & Chr(1) & Chr(2), True
Exit Sub
errie:
Err.Clear
'ClearAll
End Sub

Public Sub UseOrangez()
On Error GoTo errie
If (ConnState < 4) Then Exit Sub
    Dim X As Long
    X = Find_HealItem(healitem1)
    If (X > 0) Then
        Stat "Use " + AllInv(X).Name + vbCrLf
        Winsock_SendPacket IntToChr(&HA7) & IntToChr(X) & AccountID, True
    Else
        Stat "Can't find item in the list..." & vbCrLf
    End If
Exit Sub
errie:
Err.Clear
'ClearAll
End Sub

Private Sub AutoSellItem()
On Error GoTo errie
If (ConnState < 4) Then Exit Sub
Dim X As Integer
For X = 0 To UBound(SelItem)
   If Find_Item(SelItem(X).Name) > 0 Then
      'Stat "Automatic Sell [" + Return_ItemName(SelItem(x).Name) + "] " + CStr(Inventory(Find_Item(SelItem(x).Name)).Amount) + " EA" + vbCrLf
      Winsock_SendPacket IntToChr(&H90) & SellNPC.ID & Chr(1) & _
      IntToChr(&HC5) & SellNPC.ID & Chr(1) & IntToChr(&HC9) & IntToChr(8) _
      & IntToChr(CLng(Find_Item(SelItem(X).Name))) & _
      IntToChr(CLng(AllInv(Find_Item(SelItem(X).Name)).Amount)), True
    End If
    
Next
Exit Sub
errie:
Err.Clear
'ClearAll
End Sub


'Public Sub SendRaw()
'Dim tstr As String
'Dim tstr2 As String
'tstr = Replace(frmSendRaw.txtRaw.text, " ", "")
'Dim X As Integer
'For X = 1 To Len(tstr) Step 2
'    tstr2 = tstr2 + Chr(Val("&h" + Mid(tstr, X, 2)))
'Next
'Winsock_SendPacket tstr2, True
'End Sub

Public Sub Update_Stat(StatID As Integer)
    Winsock_SendPacket IntToChr(&HBB) + IntToChr(&HD + StatID) + Chr(1), True
End Sub

Public Sub Update_SkillLV(SkillID As Integer)
   Winsock_SendPacket Chr(&H12) + Chr(1) + IntToChr(CLng(SkillID)), True
End Sub

Private Sub Check_SkillAttack()
On Error GoTo errie
    Dim X As Integer
    Dim Y As Integer
    IsSelectSkill = False
    For X = 0 To UBound(Attack) - 1
        If CurAtkMonster.NameID = Attack(X).ID Then
            If (Attack(X).Skill1 <> "") Then
                If (SkillCounter < Attack(X).UTime1) Or (Attack(X).UTime1 = 0) Then
                    skillpacket = Attack(X).Skill1
                    SPBound = Attack(X).sp1
                    IsSelectSkill = True
                    Exit Sub
                 ElseIf ((SkillCounter <= (Attack(X).UTime1 + Attack(X).UTime2)) Or (Attack(X).UTime2 = 0)) And Len(Attack(X).Skill2) > 0 Then
                    skillpacket = Attack(X).Skill2
                    SPBound = Attack(X).sp2
                    IsSelectSkill = True
                    Exit Sub
                End If
            End If
        End If
    Next
Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "Check_SkillAttack", Err.number, Err.Description
Err.Clear
'ClearAll
End Sub

Function Check_SelfSkill() As Boolean
On Error GoTo errie
    Dim X As Integer
    For X = 0 To UBound(Attack) - 1
        If CurAtkMonster.ID = Attack(X).ID Then
            If (Attack(X).Skill1 <> "") Then
                If (SkillCounter < Attack(X).UTime1) Or (Attack(X).UTime1 = 0) Then
                    Check_SelfSkill = IsSelfSkill(Attack(X).Spell1)
                    Exit Function
                ElseIf ((SkillCounter <= (Attack(X).UTime1 + Attack(X).UTime2)) Or (Attack(X).UTime2 = 0)) And Len(Attack(X).Skill2) > 0 Then
                    Check_SelfSkill = IsSelfSkill(Attack(X).Spell1)
                    Exit Function
                End If
            End If
        End If
    Next
    Check_SelfSkill = False
errie:
    Err.Clear
    Exit Function
End Function

Private Function Check_LandSkill() As Boolean
On Error GoTo errie
    Dim X As Integer
    Dim Y As Integer, SkillPos As Coord
    For X = 0 To UBound(Attack) - 1
        If CurAtkMonster.NameID = Attack(X).ID Then
            If (Attack(X).Skill1 <> "") Then
                Randomize
                If (SkillCounter < Attack(X).UTime1) Or (Attack(X).UTime1 = 0) Then
                    Y = Find_SkillId(Attack(X).Spell1)
                    If SkillChar(Y).Target = 2 Then
                        SkillPos = NextPos(curPos, CurAtkMonster.Pos)
                        If Rnd() * 2 >= 1 Then SkillPos = NextPos(SkillPos, CurAtkMonster.Pos)
                        Randomize
                        If Rnd() * 2 >= 1 Then SkillPos = NextPos(SkillPos, CurAtkMonster.Pos)
                        If Attack(X).Spell1 = "MG_SAFETYWALL" Or Attack(X).Spell1 = "AL_PNEUMA" _
                        Then
                            SkillPos = curPos
                        End If
                        'Stat "Casting : " & Attack(x).Spell1 & " at : " & CStr(SkillPos.y) & "," & CStr(SkillPos.x) + vbCrLf
                        skillpacket = IntToChr(&H116) & IntToChr(CLng(Attack(X).lv1)) & IntToChr(find_skill_id(Attack(X).Spell1)) & IntToChr(SkillPos.Y) & IntToChr(SkillPos.X)
                        SPBound = Attack(X).sp1
                        Check_LandSkill = True
                        Exit Function
                    End If
                ElseIf ((SkillCounter <= (Attack(X).UTime1 + Attack(X).UTime2)) Or (Attack(X).UTime2 = 0)) And Len(Attack(X).Skill2) > 0 Then
                    Y = Find_SkillId(Attack(X).Spell2)
                    If SkillChar(Y).Target = 2 Then
                        SkillPos = NextPos(curPos, CurAtkMonster.Pos)
                        If Rnd() * 2 >= 1 Then SkillPos = NextPos(SkillPos, CurAtkMonster.Pos)
                        Randomize
                        If Rnd() * 2 >= 1 Then SkillPos = NextPos(SkillPos, CurAtkMonster.Pos)
                        If Attack(X).Spell2 = "MG_SAFETYWALL" Or Attack(X).Spell2 = "AL_PNEUMA" _
                        Then
                            SkillPos = curPos
                        End If
                        'Stat "Casting : " & Attack(x).Spell1 & " at : " & CStr(SkillPos.y) & "," & CStr(SkillPos.x) + vbCrLf
                        skillpacket = MakePort(&H116) & MakePort(Attack(X).lv2) & MakePort(find_skill_id(Attack(X).Spell1)) & MakePort(SkillPos.Y) & MakePort(SkillPos.X)
                        SPBound = Attack(X).sp2
                        Check_LandSkill = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    Check_LandSkill = False
    Exit Function
errie:
    If Err.number > 0 Then print_funcerr "Check_LandSkill", Err.number, Err.Description
    Err.Clear
Check_LandSkill = False
End Function
Private Function IsLandSkill(srawname As String, slevel As Byte) As Boolean
On Error GoTo errie
    Dim X As Integer
    'Dim Y As Integer, SkillPos As Coord
    X = Find_SkillId(srawname)
    If X > 0 Then
        If SkillChar(X).Target = 2 Then
            skillpacket = IntToChr(&H116) & Chr(slevel) & Chr(0) & IntToChr(find_skill_id(srawname)) & IntToChr(curPos.Y) & IntToChr(curPos.X)
            IsLandSkill = True
            Exit Function
        End If
    End If
    IsLandSkill = False
    Exit Function
errie:
    If Err.number > 0 Then print_funcerr "IsLandSkill", Err.number, Err.Description
    Err.Clear
IsLandSkill = False
End Function
Function IsSelfSkill(srawname As String) As Boolean
On Error GoTo errie
    Dim X As Integer
    X = Find_SkillId(srawname)
    If X > 0 Then
        If SkillChar(X).Target = 4 Then
            IsSelfSkill = True
            Exit Function
        End If
    End If
    IsSelfSkill = False
    Exit Function
errie:
    If Err.number > 0 Then print_funcerr "IsSelfSkill", Err.number, Err.Description
    Err.Clear
    IsSelfSkill = False
End Function
Sub Check_ResetCounter()
On Error GoTo errie
    Dim i&
    For i = 0 To UBound(Attack)
        If Attack(i).Name = CurAtkMonster.Name Then
            If SkillCounter > (Attack(i).UTime1 + Attack(i).UTime2) And _
            Attack(i).UTime1 > 0 And Attack(i).UTime2 > 0 And Not UseWeapon Then
                SkillCounter = 0
            End If
        End If
    Next
Exit Sub
errie:
If Err.number > 0 Then print_funcerr "Check_ResetCounter", Err.number, Err.Description
Err.Clear
End Sub

Sub SendAttack()
On Error GoTo errie
IsRandommove = False
If StopAction Or (Not ChkAtk) Then Exit Sub
If (ConnState < 4) Then Exit Sub
    If UBound(Aggro) > MobTeleNum And MobTeleNum > 0 Then
        Stat "Get mobs [total: " & UBound(Aggro) & "], Teleport..." & vbCrLf
        Teleport
        Exit Sub
    End If
   If (Use_Set_Skill2) And UseSkillMobs2 Then
        If LCase(MobSkill2.rawname) = "warp" Or LCase(MobSkill2.rawname) = "teleport" Or LCase(MobSkill2.rawname) = "teleport" Then
            Stat "Get mobs2, Teleport..." & vbCrLf
            Teleport
            Exit Sub
        End If
        If CanUseSkill(MobSkill2.rawname, MobSkill2.Lv, Players(number).SP) Or Players(number).SP > 100 Then
            Stat "Get mobs2, Use skill [" & Get_SkillName(MobSkill2.rawname) & "] lv." & CStr(MobSkill2.Lv) & "..." & vbCrLf
            TraceMons = False
            If IsLandSkill(MobSkill2.rawname, MobSkill2.Lv) Then
                Winsock_SendPacket skillpacket, True
            ElseIf IsSelfSkill(MobSkill2.rawname) Then
                Winsock_SendPacket MobSkill2.Packet + AccountID, True
            Else
                Flood_Packet MobSkill2.Packet + CurAtkMonster.ID, 10
            End If
            Exit Sub
        Else
            Stat "Get mobs2, Try to use skill but not SP" & IIf(isKillmob, ", do normal attack", ", Teleport") & "..." & vbCrLf
            If Not isKillmob Then Teleport: Exit Sub
        End If
    End If
   If (Use_Set_Skill) And UseSkillMobs Then
        If LCase(MobSkill.rawname) = "warp" Then
            Stat "Get mobs, Teleport..." & vbCrLf
            Teleport
            Exit Sub
        End If
        If CanUseSkill(MobSkill.rawname, MobSkill.Lv, Players(number).SP) Or Players(number).SP > 100 Then
            Stat "Get mobs, Use skill [" & Get_SkillName(MobSkill.rawname) & "] lv." & CStr(MobSkill.Lv) & "..." & vbCrLf
            TraceMons = False
            If IsLandSkill(MobSkill.rawname, MobSkill.Lv) Then
                Winsock_SendPacket skillpacket, True
            ElseIf IsSelfSkill(MobSkill.rawname) Then
                Winsock_SendPacket MobSkill.Packet & AccountID, True
            Else
                Winsock_SendPacket MobSkill.Packet & CurAtkMonster.ID, True
            End If
            Exit Sub
        Else
            Stat "Get mobs, Try to use skill but no SP" & IIf(isKillmob, ", do normal attack", ", Teleport") & " ..." & vbCrLf
            If Not isKillmob Then Teleport: Exit Sub
        End If
    End If
    If IsNearPortal(CurAtkMonster.Pos, 6) Then
        Stat "Current monster is close to portal. Cancel target & teleport." & vbCrLf
        Clear_This_Mons 0
        'Dim lix&, lif As Boolean
       ' For lix = 0 To UBound(MonsterList)
       '     If MonsterList(lix).ID = CurAtkMonster.ID Then
       '         With MonsterList(lix)
       '             .CantGo = True
       '             .IsAttack = True
       '             .IsFollow = True
       '             .NoAttack = True
       '         End With
       '     End If
       ' Next
       ' lif = False
       ' For lix = 0 To UBound(Aggro)
       '     If Aggro(lix).ID = CurAtkMonster.ID Then lif = True
       '     If lif And lix < UBound(Aggro) Then Aggro(lix) = Aggro(lix + 1)
       ' Next
       ' If lif Then ReDim Preserve Aggro(UBound(Aggro) - 1)
       ' CurAtkMonster.NameID = 0
       ' CurAtkMonster.Name = ""
       ' CurAtkMonster.CantGo = True
       ' CurAtkMonster.IsAttack = True
       ' CurAtkMonster.IsFollow = True
       ' CurAtkMonster.NoAttack = True
        Teleport
        Exit Sub
    End If
    Dim melee As Boolean
    With Players(number)
    If (.Class = "Archer" Or .Class = "Hunter") Then
        ArcherMode = True
        MageMode = False
        Range = ArcherRange
        melee = False
    ElseIf (.Class = "Mage" Or .Class = "Wizard" Or .Class = "Sage" Or _
        .Class = "Acolyte" Or .Class = "Priest" Or .Class = "Super Novice") And (Not UseWeapon) Then
        ArcherMode = False
        MageMode = True
        Range = MageRange
        melee = False
    ElseIf (.Class = "Thief" Or ((.Class = "Acolyte" Or .Class = "Priest") And UseWeapon)) Or (UseWeapon And (.Class = "Mage" Or .Class = "Wizard" Or .Class = "Sage" Or .Class = "Super Novice")) Then
        ArcherMode = False
        MageMode = False
        melee = True
        Range = PRange
    ElseIf .Class = "Swordman" Or .Class = "Knight" Or .Class = "Crusader" Then
        ArcherMode = False
        MageMode = False
        melee = True
        Range = PRange
    Else
        ArcherMode = False
        MageMode = False
        melee = True
        Range = PRange
    End If
   
    If (IsUseRange) Then
        Range = RangeSet
        If .Class = "Mage" Or .Class = "Wizard" Or .Class = "Sage" Or _
        .Class = "Acolyte" Or .Class = "Priest" Then
            ArcherMode = False
            MageMode = True
            melee = False
        Else
            ArcherMode = True
            MageMode = False
            melee = False
        End If
    End If
    End With
   If (IsSkillUse) Then Check_SkillAttack

    If (CurAtkMonster.NameID > 0) Then
         '------------------------- Melee Mode ---------------------------------
            If ((Not MageMode) Or UseWeapon) And (Not ArcherMode) Or melee Then 'Tanker!
                If (((Players(number).SP > SPBound) And ((EvalNorm(CurAtkMonster.Pos, curPos) < 3)) _
                And (IsSelectSkill) And (IsSkillUse))) Then
                    TraceMons = False
                    If Check_LandSkill Then
                        Winsock_SendPacket skillpacket, True
                    Else
                        If Check_SelfSkill Then
                            Winsock_SendPacket skillpacket + AccountID, True
                        Else
                            Flood_Packet skillpacket + CurAtkMonster.ID, 10
                        End If
                    End If
                    SkillDelay = 0
                    tmrSkillDelay.Enabled = True
                ElseIf ((Not Automove) Or (EvalNorm(CurAtkMonster.Pos, curPos) <= (Range + 1))) Or (MakeDamage) Then
                    Winsock_SendPacket IntToChr(&H89) & CurAtkMonster.ID + Chr(7), True
                ElseIf (Automove) And Not BlockMove And Not MakeDamage Then
                    move_to NearPos(CurAtkMonster.Pos, curPos, Range)
                End If
        '----------------------- Mage and Wizard Mode -------------------------
            ElseIf MageMode And (Not ArcherMode) Then 'Is Mage?
                If (Players(number).SP > 14) And _
                    (EvalNorm(CurAtkMonster.Pos, curPos) < Range) And _
                    (EvalNorm(CurAtkMonster.Pos, curPos) >= MinDistance Or _
                    Not useMinDistance) And (IsSelectSkill) And (IsSkillUse) Then
                    If Check_LandSkill Then
                        Winsock_SendPacket skillpacket, True
                    Else
                        If Check_SelfSkill Then
                            Winsock_SendPacket skillpacket + AccountID, True
                        Else
                            Flood_Packet skillpacket + CurAtkMonster.ID, 10
                        End If
                    End If
                    SkillDelay = 0
                    tmrSkillDelay.Enabled = True
                ElseIf (Automove) And Not BlockMove And Not Casting Then
                    Winsock_SendPacket IntToChr(&H85) & MakeMagePos(CurAtkMonster.Pos), True
                End If
        '----------------------- Archer and Hunter Mode -----------------------
            ElseIf ArcherMode Then 'Is Archer?
                If (Players(number).SP > 14) And _
                    (EvalNorm(CurAtkMonster.Pos, curPos) < Range) And _
                    (EvalNorm(CurAtkMonster.Pos, curPos) >= MinDistance Or Not useMinDistance) And _
                    (IsSelectSkill) And (IsSkillUse) Then
                    If Check_LandSkill Then
                        Winsock_SendPacket skillpacket, True
                    Else
                        If Check_SelfSkill Then
                            Winsock_SendPacket skillpacket + AccountID, True
                        Else
                            Winsock_SendPacket skillpacket + CurAtkMonster.ID, True
                        End If
                    End If
                    UseArrow = True
                    SkillDelay = 0
                    tmrSkillDelay.Enabled = True
                'ElseIf (MakeDamage And EvalNorm(CurAtkMonster.pos, curPos) > 5) Or (Not Automove) Or (IsDamage And EvalNorm(CurAtkMonster.pos, curPos) < Range) Or _
                ((EvalNorm(CurAtkMonster.pos, curPos) < Range) And _
                    (EvalNorm(CurAtkMonster.pos, curPos) >= MinDistance Or _
                    Not useMinDistance) And Not BlockMove) Then
                ElseIf (Not IsSelectSkill Or Not IsSkillUse) And (EvalNorm(CurAtkMonster.Pos, curPos) < Range) And _
                    (EvalNorm(CurAtkMonster.Pos, curPos) >= MinDistance Or Not useMinDistance) Then
                    Winsock_SendPacket IntToChr(&H89) & CurAtkMonster.ID + Chr(7), True
                    UseArrow = True
                    TraceMons = True
                ElseIf (Automove) And Not BlockMove Then
                    Winsock_SendPacket IntToChr(&H85) & MakeMagePos(CurAtkMonster.Pos), True
                End If
            End If
    
    End If
    Exit Sub
errie:
Err.Clear
'ClearAll
NumberMons = 0
End Sub

Private Function rndCoord(inPos As Coord) As Coord
'startagain:
Dim xoff As Integer
Dim yoff As Integer
If curPos.X < inPos.X Then xoff = -1
If curPos.X > inPos.X Then xoff = 1
If curPos.Y < inPos.Y Then yoff = -1
If curPos.Y > inPos.Y Then yoff = 1
Dim tpos As Coord
tpos.X = inPos.X + xoff
tpos.Y = inPos.Y + yoff
If Abs(curPos.X - tpos.X) > 4 Then
    If curPos.X > tpos.X Then
        rndCoord.X = curPos.X - 5
    Else
        rndCoord.X = curPos.X + 5
    End If
Else
    rndCoord.X = tpos.X
End If
If Abs(curPos.Y - tpos.Y) > 4 Then
    If curPos.Y > tpos.Y Then
        rndCoord.Y = curPos.Y - 5
    Else
        rndCoord.Y = curPos.Y + 5
    End If
Else
    rndCoord.Y = tpos.Y
End If
End Function

'Public Sub BotDebug(text As String)
    'frmDebug.txtDebug = frmDebug.txtDebug & text & vbCrLf
    'frmDebug.txtDebug.SelStart = Len(frmChat.txtChat.text)
    'frmDebug.txtDebug.SelLength = 0
'End Sub

Public Function find_skill_id(ByVal tstr As String) As Integer
On Error Resume Next
    Dim X As Integer
    For X = 0 To UBound(SkillChar)
        If InStr(tstr, SkillChar(X).Name) > 0 And SkillChar(X).MaxLV > 0 Then
                find_skill_id = SkillChar(X).ID
                Exit Function
        End If
    Next
    find_skill_id = 0
    Err.Clear
End Function

Public Sub ClearAll()
   ptEnd.X = 0
   ptEnd.Y = 0
   'CryptFrame = ""
   'AutoItem.TimeCount = 0
   BackWpCounter = 0
   DetectPortal = False
   ReDim NPCList(0)
   ReDim ExitPortal(0)
   ReDim Unknow(0)
   ReDim Aggro(0)
   ReDim People(0)
   ReDim AllInv(0)
   ReDim Route(0)
   TraceMons = False
   BlockMove = False
   frmPeople.lstPeople.Clear
   Pickuptime = 0
   TryPicktime = 0
   Tracing = False
   CurrentItem.Name = ""
   CurrentItem.ID = ""
   SpellCounter = 0
   SendUsePot = False
   SkillWait = False
   SkillCounter = 0
   AttackCounter = 0
   ResponseCounter = 0
   Wait = False
   SendAction = False
   'UseHeal = False
   UseBow = False
   SendSell = False
   SendHeal = False
   IsSell = False
   isWarp = False
    UsePotCounter = 0
    DError = False
    AttCounter = 0
    DamageCounter = 0
    IsAggro = False
    tmrPickup.Enabled = False
    'IsSelectSkill = False
    ReDim MonsterList(0)
    ReDim Items(0)
    'ReDim Inventory(0)
    'ReDim Equipment(0)
    CurAtkMonster.NameID = 0
    NumberMons = 0
    labCurMons.Caption = "[None]"
    'Winsock_SendPacket (IntToChr(&H7D))
    InFight = False
    Pickup = False
    Sitting = False
    IsUseSkill = False
    ClearCounter = 0
    GotCurItem = False
    DealtDamage = False
    ResponseOK = True
    IsDamage = False
    MakeDamage = False
    IsAggro = False
    GotCurItem = False
    UseArrow = False
    'Recvdata = ""
End Sub

Public Sub Set_Share()
    Winsock_SendPacket (Chr(2) & Chr(1) & Chr(1) & IntToChr(0) & Chr(0)), True
    Winsock_SendPacket (IntToChr(&H7E) + MakeTickString), True
    ClearCounter = 0
    MDIfrmMain.mnuShare.CheckED = True
    MDIfrmMain.mnuUnshare.CheckED = False
End Sub

Public Sub unSet_Share()
    Winsock_SendPacket (Chr(2) & Chr(1) & Chr(0) & IntToChr(0) & Chr(0)), True
    Winsock_SendPacket (IntToChr(&H7E) + MakeTickString), True
    ClearCounter = 0
    MDIfrmMain.mnuShare.CheckED = False
    MDIfrmMain.mnuUnshare.CheckED = True
End Sub

Public Sub send_exall()
    Winsock_SendPacket (IntToChr(&HD0) & Chr(0)), True
End Sub

Public Sub send_inall()
    Winsock_SendPacket (IntToChr(&HD0) & Chr(1)), True
End Sub

Sub Clear_Mon_List(ID As String)
On Error GoTo errie
    Dim X As Integer
    Dim Y As Integer
    If UBound(MonsterList) = 0 Then Exit Sub
    For X = 0 To UBound(MonsterList) - 1
        If ID = MonsterList(X).ID Then
                CheckEvent "OnMonsterDisappear", "name=" & MonsterList(X).Name & Chr(0) & "posX=" & MonsterList(X).Pos.Y & Chr(0) & "posY=" & MonsterList(X).Pos.X
                Clear_Dot MonsterList(X).Pos
                TmpAggroName = Return_MonsterName(MonsterList(X).NameID)
                DeadMonPos = MonsterList(X).Pos
                For Y = X To UBound(MonsterList) - 1
                    MonsterList(Y) = MonsterList(Y + 1)
                Next
                ReDim Preserve MonsterList(UBound(MonsterList) - 1)
                upd_frmMonster
                Exit For
        End If
    Next
Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "Clear_Mon_List", Err.number, Err.Description
    Err.Clear
End Sub


Private Sub Check_BackTown()
On Error GoTo errie
    If Players(number).Weight = 0 Or Players(number).MaxWeight = 0 And (Not SellMode) Then
        FightMode = True
        SellMode = False
        If FightMode And MoveOnly Then
            Direction = FightDirection
        End If
        Exit Sub
    End If
    
    If isBackBuy And NeedBuy And UBound(AllInv) > 0 Then
        If Not MoveOnly And (Not TmrBackTown_Enabled) Then
            Warp_Save "System : [Need Buy], You set to back to town..."
            TmrBackTown_Enabled = True
            If UBound(WayPoint) = 0 Then Check_Destination_Route
        End If
    End If
    
    If Players(number).Weight >= (WeightBackTown * Players(number).MaxWeight) _
    And IsBackTown And HaveSellItem Then
        If SaveMapName <> "" And SaveMapName <> MapName _
    And Not MoveOnly And (Not TmrBackTown_Enabled) Then
        Warp_Save "System : [Overweight], You set to back to town..."
        TmrBackTown_Enabled = True
        If UBound(WayPoint) = 0 Then Check_Destination_Route
        End If
    End If
    If HaveSellItem And Players(number).Weight >= (WeightBackTown * Players(number).MaxWeight) And IsBackTown And (Not SellMode) Then
        If Not SellMode Then Chat "System : [Overweight], You set to back to town...", MColor.Fail
        SellMode = True
        FightMode = False
    ElseIf ((SellNPC.NameID = 0) And Not SellMode) Or (Not HaveSellItem) Then
        FightMode = True
        SellMode = False
        If FightMode And MoveOnly Then
            Direction = FightDirection
        End If
    End If
    
Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "Check_BackTown", Err.number, Err.Description
    Err.Clear
End Sub

'Public Function HaveStorageList() As Boolean
'On Error GoTo EndFunc
'    Dim X, i As Integer
'    Dim index As Integer
'    If NoStoreItem Then GoTo EndFunc
'    If Storage(0).Name = "" Then GoTo EndFunc
'    For X = 0 To UBound(GetStorageItem)
'            index = Find_StorageID(GetStorageItem(X).Name)
'            If index = 0 Then
'                HaveStorageList = True
'                Exit Function
'            ElseIf Storage(index).Amount > 0 Then
'                HaveStorageList = True
'                Exit Function
'            End If
'    Next
'EndFunc:
'    HaveStorageList = False
'End Function

Private Sub Send_Who()
    Winsock_SendPacket IntToChr(&HC1), True
End Sub


Private Sub print_header(Packet As String)
    ''print_errror "sub print_packet"
    Dim tstr As String
    Dim X As Integer
    tstr = ""
    For X = 1 To Len(Packet)
       If Asc(Mid(Packet, X, 1)) < 16 Then tstr = tstr + "0"
       tstr = tstr + Hex(Asc(Mid(RecvData, X, 1))) + " "
       If X Mod 16 = 0 Then tstr = tstr & vbCrLf
    Next
       tstr = Left(tstr, Len(tstr) - 1)
       Open App.Path & "\log\packetlog.txt" For Append As #1
       Print #1, " == Header Packet == " & vbCrLf + tstr
       Close #1
End Sub

Sub print_packet(Packet As String, text As String)
       Open App.Path & "\log\packetlog.txt" For Append As #1
       Print #1, " == " & text & " == "
       Print #1, ConvPacketData(Packet)
       Close #1
End Sub

Public Function Use_Set_Skill() As Boolean
On Error GoTo Out
        Dim mcount As Integer, i&
        mcount = 0
        For i = 0 To UBound(MonsterList)
            If InStr(MobSkill.MonsName, MonsterList(i).Name) > 0 And isInAggro(MonsterList(i).ID) Then mcount = mcount + 1
        Next
        If mcount > MobSkill.number Then
            skillpacket = MobSkill.Packet
            Use_Set_Skill = True
        End If

Out:
Err.Clear
End Function
Public Function Use_Set_Skill2() As Boolean
On Error GoTo Out
        skillpacket = MobSkill2.Packet
        Dim mcount As Integer, i&
        mcount = 0
        For i = 0 To UBound(MonsterList)
            If InStr(MobSkill2.MonsName, MonsterList(i).Name) > 0 And isInAggro(MonsterList(i).ID) Then mcount = mcount + 1
        Next
        If mcount > MobSkill2.number Then
            skillpacket = MobSkill.Packet
            Use_Set_Skill2 = True
        End If
Out:
Err.Clear
End Function
Public Function isInAggro(ID As String) As Boolean
On Error GoTo Out
    Dim i&
    For i = 0 To UBound(Aggro)
        If Aggro(i).ID = ID Then
            isInAggro = True
            Exit Function
        End If
    Next
Out:
Err.Clear
End Function

Private Sub Winsock1_SendComplete()
    Sending = False
End Sub

Private Sub Winsock1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    Sending = True
End Sub

Private Sub Send_To_Culvert()
    Winsock_SendPacket IntToChr(&H90) & CulVertNPC & Chr(1) & _
    IntToChr(&HB9) & CulVertNPC & IntToChr(&HB8) & CulVertNPC & _
    Chr(1), True
End Sub

Public Sub create_chatroom(ByVal pass As String, ByVal topic As String)
    Winsock_SendPacket IntToChr(&HD5) & IntToChr(15 + Len(topic)) & _
    IntToChr(&H14) & Chr(0) & pass & String(8 - Len(pass), Chr(0)) & topic, True
    Winsock_SendPacket IntToChr(&H94) & AccountID, True
End Sub

Public Sub edit_chatroom(ByVal pass As String, ByVal topic As String)
    Winsock_SendPacket IntToChr(&HDE) & IntToChr(15 + Len(topic)) & _
    IntToChr(&H14) & Chr(0) & pass & String(8 - Len(pass), Chr(0)) & topic, True
    'Winsock_SendPacket IntToChr(&H94) & AccountID, True
End Sub

Public Sub destroy_chatroom()
    If Not Sending Then Winsock_SendPacket IntToChr(&HE3), True
End Sub

Public Sub Send_Sit()
    If Not IsDamage Then Winsock_SendPacket IntToChr(&H89) & String(4, Chr(0)) & Chr(2), True
    DelayCheckRest = RandomNumber(3, 2)
End Sub

Public Sub Send_Stand()
    Winsock_SendPacket IntToChr(&H89) & String(4, Chr(0)) & Chr(3), True
    DelayCheckRest = RandomNumber(3, 2)
End Sub

Public Sub send_chat(ByVal txtword As String)
    Winsock_SendPacket Chr(&H8C) + Chr(0) + _
                        IntToChr(Len(CharNameStart) + Len(txtword) + 8) + _
                        CharNameStart + " : " + txtword + Chr(0), True
End Sub

Private Function IsClosePortal() As Boolean
On Error GoTo errie
    Dim i As Integer
    'If UBound(ExitPortal) = 0 Then GoTo endsub
    If ExitPortal(0).Pos.X = 0 Or ExitPortal(0).Pos.Y = 0 Then GoTo endsub
        For i = 0 To UBound(ExitPortal) - 1
            If EvalNorm(curPos, ExitPortal(i).Pos) < 5 Then
                IsClosePortal = True
                Exit Function
            End If
        Next
endsub:
        IsClosePortal = False
Exit Function
errie:
    If Err.number > 0 Then print_funcerr "IsClosePortal", Err.number, Err.Description
    Err.Clear
End Function

Private Function IsNearPortal(myPos As Coord, Optional dist As Integer = 12) As Boolean
On Error GoTo errie
    Dim i As Integer
    If UBound(ExitPortal) > 0 Then
        For i = 0 To UBound(ExitPortal) - 1
            If EvalNorm(myPos, ExitPortal(i).Pos) < dist Then
                IsNearPortal = True
                Exit Function
            End If
        Next
    End If
    IsNearPortal = False
Exit Function
errie:
    If Err.number > 0 Then print_funcerr "IsNearPortal", Err.number, Err.Description
    Err.Clear
End Function

Private Sub Update_CurPos()
On Error GoTo errie
    Dim Speed As Integer
    Dim TimeRef As Long
    'If TmrMove.Enabled And PlayerMoveTime > 0 Then
    If PlayerMoveTime > 0 Then
        If GetTickCount - PlayerMoveTime >= MovementSpeed Then
            Clear_Dot curPos
            curPos = NextPos(curPos, tmpPos)
            'If Abs(CurPos.X - tmpPos.X) > 0 Then
            '    CurPos.X = CurPos.X + Sgn(tmpPos.X - CurPos.X)
            'End If
            'If Abs(CurPos.Y - tmpPos.Y) > 0 Then
            '    CurPos.Y = CurPos.Y + Sgn(tmpPos.Y - CurPos.Y)
            'End If
            If FrmField.Visible Then Plot_Dot curPos, vbBlue
            If CurAtkMonster.NameID > 0 Then SendAttack
            Label14.Caption = curPos.X
            Label12.Caption = curPos.Y
            PlayerMoveTime = GetTickCount
            upd_curMonster
            noMoveCounter = 0
            If IsClosePortal And Not MoveOnly And AutoAI And Not IsOnWayPoint(curPos) Then Teleport
        End If
        If curPos.X = tmpPos.X And curPos.Y = tmpPos.Y Then
            'If CurAtkMonster.Nameid > 0 Then SendAttack
            PlayerMoveTime = 0
            'TmrMove.Enabled = False
            BlockMove = False
            BackWP = False
        End If
    Else
        PlayerMoveTime = 0
    End If
    
Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "Update_CurPos", Err.number, Err.Description
    Err.Clear
End Sub


Private Sub Update_MonsterPos()
On Error GoTo errie
    If UBound(MonsterList) > 0 Then
        Dim TimeRef As Long, ListSelected As Boolean
        Dim X As Integer
        ListSelected = False
        For X = 0 To UBound(MonsterList) - 1
            If MonsterList(X).Time > 0 Then
                'TimeRef = GetTickCount
                    'MonsterList(x).Time = GetTickCount
                If GetTickCount - MonsterList(X).Time >= MonsterList(X).Speed Then
                    Clear_Dot MonsterList(X).Pos
                        
                    If Abs(MonsterList(X).Pos.X - MonsterList(X).NextPos.X) > 0 Then MonsterList(X).Pos.X = MonsterList(X).Pos.X + Sgn(MonsterList(X).NextPos.X - MonsterList(X).Pos.X)
                    If Abs(MonsterList(X).Pos.Y - MonsterList(X).NextPos.Y) > 0 Then MonsterList(X).Pos.Y = MonsterList(X).Pos.Y + Sgn(MonsterList(X).NextPos.Y - MonsterList(X).Pos.Y)
                        
                    Dim dist&
                    dist = EvalNorm(curPos, MonsterList(X).Pos)
                    If MyPet.ID = MonsterList(X).ID Then
                        Plot_Dot MonsterList(X).Pos, &HFF00FF
                    ElseIf CurAtkMonster.ID <> MonsterList(X).ID Then
                        Plot_Dot MonsterList(X).Pos, vbRed
                        If Not ListSelected And MonsterList(X).IsAttack And CurAtkMonster.NameID < 1 Then
                            If IsSMAgg(MonsterList(X).Name) Then
                                ListSelected = True
                                Send_Stand
                                IsSitting = False
                                IsStanding = True
                                CurMonsterName = Return_MonsterName(MonsterList(X).NameID)
                                Stat "Select [" + CurMonsterName + "] as a Target, Locking..." + vbCrLf
                                SpellCounter = 0
                                SkillCounter = 0
                                CurAtkMonster = MonsterList(X)
                                Check_Equip CurMonsterName
                                Plot_Dot MonsterList(X).Pos, CurAtkColor
                                Aggro(0).ID = MonsterList(X).ID
                                oldSelectPos = CurAtkMonster.Pos
                                NumberMons = 0
                                Tracing = False
                                SendAction = True
                                SendAttack
                                upd_curMonster
                                NomonsTimeCount = 0
                                InFight = False
                                MakeDamage = False
                                IsAggro = False
                                IsDamage = False
                                DamageCounter = 0
                                AttackCounter = 0
                            End If
                        End If
                    Else
                        Plot_Dot MonsterList(X).Pos, CurAtkColor
                    End If
                    MonsterList(X).Time = GetTickCount
                    If MonsterList(X).Pos.X = MonsterList(X).NextPos.X And _
                    MonsterList(X).Pos.Y = MonsterList(X).NextPos.Y Then
                        MonsterList(X).Time = 0
                        If CanGO(curPos, MonsterList(X).Pos) Then MonsterList(X).CantGo = False
                    End If
                    If CurAtkMonster.ID = MonsterList(X).ID Then
                        CurAtkMonster.Pos = MonsterList(X).Pos
                        SendAttack
                    End If
                End If
            End If
        Next
    End If
Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "Update_MonsterPos", Err.number, Err.Description
    Err.Clear
End Sub

Private Sub Update_PeoplePos()
On Error GoTo errie
    Dim TimeRef As Long
    Dim X As Integer
    If UBound(People) > 0 Then
        For X = 0 To UBound(People) - 1
            If People(X).Time > 0 Then
                TimeRef = GetTickCount
                If TimeRef - People(X).Time >= People(X).Speed Then
                    If Not Disable_frmPeople Then Clear_Dot People(X).Pos
                    If Abs(People(X).Pos.X - People(X).NextPos.X) > 0 Then People(X).Pos.X = People(X).Pos.X + Sgn(People(X).NextPos.X - People(X).Pos.X)
                    If Abs(People(X).Pos.Y - People(X).NextPos.Y) > 0 Then People(X).Pos.Y = People(X).Pos.Y + Sgn(People(X).NextPos.Y - People(X).Pos.Y)
                    People(X).Time = TimeRef
                    If Not Disable_frmPeople Then Plot_Dot People(X).Pos, PColor
                    If People(X).Pos.X = People(X).NextPos.X And _
                    People(X).Pos.Y = People(X).NextPos.Y Then People(X).Time = 0
                End If
            End If
        Next
    End If
Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "Update_PeoplePos", Err.number, Err.Description
    Err.Clear
End Sub

Public Sub Send_GetStoreList()
    Winsock_SendPacket IntToChr(&HC5) & CurNPC & Chr(0), True
End Sub

Public Sub Send_Sell()
    Winsock_SendPacket IntToChr(&HC5) & CurNPC & Chr(1), True
End Sub

Public Sub Send_Talk(ID As String)
    Winsock_SendPacket IntToChr(&H90) & ID & Chr(1), True
End Sub

Public Sub Send_TalkContinue()
    Winsock_SendPacket IntToChr(&HB9) & CurNPC, True
End Sub

Public Sub Send_TalkCancel()
    Winsock_SendPacket Chr(&H46) & Chr(1) & CurNPC, True
End Sub

Public Sub Send_TalkResponse(Choice As Integer)
    Winsock_SendPacket IntToChr(&HB8) & CurNPC & Chr(Choice), True
End Sub

Public Sub Send_BuyList(tstr As String)
    Winsock_SendPacket IntToChr(&HC8) & IntToChr(Len(tstr) + 4) & tstr, True
End Sub

Public Sub Send_SellList(tstr As String)
    Winsock_SendPacket IntToChr(&HC9) & IntToChr(Len(tstr) + 4) & tstr, True
End Sub

Private Function Decode_00AA(inData As String) As String
On Error GoTo errie
    If Asc(Mid(inData, 7, 1)) = 1 Then
        Stat "You equip [" & AllInv(MakePort(Mid(inData, 3, 2))).Name & "]..." & vbCrLf
        AllInv(MakePort(Mid(inData, 3, 2))).Pos = MakePort(Mid(inData, 5, 2))
        Update_frmArmor AllInv(MakePort(Mid(inData, 3, 2))).Name, MakePort(Mid(inData, 5, 2))
        UpdateInventory
        'UpdateEquipment
    Else
        If Mods.STDebug Then Stat "You can't equip [" & AllInv(MakePort(Mid(inData, 3, 2))).Name & "]..." & vbCrLf
    End If
    If MakePort(Mid(inData, 3, 2)) = tmpEQTelePos And WaitEquipTele = True Then
        WaitEquipTele = False
        WaitEquipBack = True
        Teleport
    End If
    CheckEvent "OnEquipItem", "itemname=" & AllInv(MakePort(Mid(inData, 3, 2))).Name & Chr(0) & "success=" & CStr(CBool(Asc(Mid(inData, 7, 1))))
Decode_00AA = ""
Exit Function
errie:
Decode_00AA = "ERROR!!! [Decode_00AA] " & Err.Description
Err.Clear
End Function

Private Function Decode_00AC(inData As String) As String
On Error GoTo errie
    Stat "You unequip [" & AllInv(MakePort(Mid(inData, 3, 2))).Name & "]..." & vbCrLf
    AllInv(MakePort(Mid(inData, 3, 2))).Pos = 0
    Update_frmArmor "-", MakePort(Mid(inData, 5, 2))
    CheckEvent "OnUnEquipItem", "itemname=" & AllInv(MakePort(Mid(inData, 3, 2))).Name
    UpdateInventory
    Decode_00AC = ""
Exit Function
errie:
Decode_00AC = "ERROR!!! [Decode_00AC] " & Err.Description
Err.Clear
End Function

Private Function Decode_01F1(inData As String) As String
On Error GoTo errie
    'ReDim Storage(0)
    Stat Mid(inData, 5, MakePort(Mid(inData, 3, 2)) - 4) & vbCrLf
    Decode_01F1 = ""
    Exit Function
errie:
    Decode_01F1 = "ERROR!!! [Decode_01F1] " & Err.Description
    Err.Clear
End Function

Private Function Decode_00C7(inData As String) As String
On Error GoTo errie
    ReDim Store(0)
    Dim Price As Long
    Dim i As Integer
    Dim Index As Integer
    Dim NameID As String
    Dim Itemname As String
    Dim ChopNumber As Long
    ChopNumber = MakePort(Mid(inData, 3, 2))
    If Not AutoAI Then
        frmStoreBuy.lstItem.Clear
        For i = 5 To ChopNumber Step 10
            Index = MakePort(Mid(inData, i, 2))
            Price = MakePort(Mid(inData, i + 2, 4))
            If Price > 0 Then
                Store(UBound(Store)) = AllInv(Index)
                Store(UBound(Store)).Index = Index
                Store(UBound(Store)).Price = Price
                frmStoreBuy.lstItem.AddItem CStr(UBound(Store)) & " : " & AllInv(Index).Name & _
                " " & CStr(AllInv(Index).Amount) & " EA - " & Format(Price, "##,##") & " z."
                ReDim Preserve Store(UBound(Store) + 1)
            End If
        Next
        If UBound(Store) > 0 Then ReDim Preserve Store(UBound(Store) - 1)
        frmStoreBuy.Visible = True
        frmStoreBuy.LabName.Caption = "Tool Dealer : Sell Item List"
        frmTMPLst.Visible = True
        frmTMPLst.imgSell.Visible = True
    Else
        Itemname = ""
        For i = 5 To ChopNumber Step 10
            Index = MakePort(Mid(inData, i, 2))
            Price = MakePort(Mid(inData, i + 2, 4))
            If Price > 0 Then
                If Is_Sell(AllInv(Index).Name) Then
                    Itemname = Itemname & IntToChr(CLng(Index)) & IntToChr(AllInv(Index).Amount)
                End If
            End If
        Next
        If Len(Itemname) > 0 Then
            Send_SellList Itemname
            SendSell = True
            tmrDealNPC.Enabled = False
            tmrDealNPC.Enabled = True
        End If
    End If
Decode_00C7 = ""
Exit Function
errie:
Decode_00C7 = "ERROR!!! [Decode_00C7] " & Err.Description
Err.Clear
End Function


Private Function Decode_00C6(inData As String) As String
On Error GoTo errie
    ReDim Store(0)
    Dim Price As Long
    Dim i As Integer
    Dim NameID As String
    Dim Itemname As String
    Dim ChopNumber As Long
    ChopNumber = MakePort(Mid(inData, 3, 2))
    If ChopNumber = 4 Then Exit Function
    If Not AutoAI Then
        frmStoreBuy.lstItem.Clear
        For i = 5 To ChopNumber Step 11
            Price = MakePort(Mid(inData, i, 4))
            NameID = MakeHexName(Mid(inData, i + 9, 2))
            Itemname = Return_ItemName(NameID)
            Store(UBound(Store)).Price = Price
            Store(UBound(Store)).NameID = NameID
            Store(UBound(Store)).Name = Itemname
            frmStoreBuy.lstItem.AddItem CStr(UBound(Store)) & " : " & Itemname & _
            " " & Format(Price, "##,##") & " z."
            ReDim Preserve Store(UBound(Store) + 1)
        Next
        ReDim Preserve Store(UBound(Store) - 1)
        frmStoreBuy.Visible = True
        frmStoreBuy.LabName.Caption = "Tool Dealer : Buy Item List"
        frmTMPLst.Visible = True
        frmTMPLst.imgBuy.Visible = True
    Else
        Dim tstr As String
        tstr = ""
        For i = 5 To ChopNumber Step 11
            Price = MakePort(Mid(inData, i, 4))
            NameID = MakeHexName(Mid(inData, i + 9, 2))
            Itemname = LCase(Return_ItemName(NameID))
            Dim Index As Integer
            Dim index2 As Integer
            Index = Is_Buy(Itemname)
            index2 = Get_AmountItem(Itemname)
            If Index > 0 And index2 < Index Then
                tstr = tstr & IntToChr(CLng(Index - index2)) & _
                IntToChr(CLng(MakePort(Mid(inData, i + 9, 2))))
                If Mods.STSystem Then Stat "Request to buy : [" & Itemname & "] x" & (Index - index2) & "EA" & vbCrLf, &HCCCCCC
            End If
        Next
        'Debug.Print MakeHexName(tstr)
        If tstr <> "" Then
            Send_BuyList tstr
            tmrDealNPC.Enabled = False
            tmrDealNPC.Enabled = True
            SendBuy = True
        End If
    End If
Decode_00C6 = ""
Exit Function
errie:
Decode_00C6 = "ERROR!!! [Decode_00C6] " & Err.Description
Err.Clear
End Function

Private Function Decode_01A4(inData As String) As String
On Error GoTo errie
    Dim i As Integer
    Dim found As Boolean
    found = False
    For i = 0 To UBound(MonsterList)
        If Mid(inData, 4, 4) = MonsterList(i).ID Then
            MonsterList(i).IsPet = True
            found = True
            Exit For
        End If
    Next
    'If Not found Then
    '    MonsterList(UBound(MonsterList)).ID = Mid(InData, 4, 4)
    '    MonsterList(UBound(MonsterList)).IsPet = True
    '    ReDim Preserve MonsterList(UBound(MonsterList) + 1)
    'End If
    If Mid(inData, 4, 4) = CurAtkMonster.ID Then
        Clear_This_Mons 0
        Stat "This's pet, abort target..." & vbCrLf
    End If
    If Mid(inData, 4, 4) = LastID Then MyPet.ID = Mid(inData, 4, 4)
    LastID = Mid(inData, 4, 4)
Decode_01A4 = ""
Exit Function
errie:
Decode_01A4 = "ERROR!!! [Decode_01A4] " & Err.Description
Err.Clear
End Function


Private Function Decode_01A3(inData As String) As String
On Error GoTo errie
    If Asc(Mid(inData, 3, 1)) = 1 Then
            Stat "Feeds [" & MyPet.Name & "] with [" & Return_ItemName(MakeHexName(Mid(inData, 4, 2))) & "]" & vbCrLf
    Else
        Stat "Can't find [" & Return_ItemName(MakeHexName(Mid(inData, 4, 2))) & "] to feeds, Set your pet back to egg!..." & vbCrLf
            BackEgg
       End If
Decode_01A3 = ""
Exit Function
errie:
Decode_01A3 = "ERROR!!! [Decode_01A3] " & Err.Description
Err.Clear
End Function


Public Function Decode_0092(inData As String) As String
On Error GoTo errie
        Dim mapgot As String, eCase As String
        eCase = "oldMap=" & MapName & Chr(0) & "oldposX=" & curPos.Y & Chr(0) & "oldposY=" & curPos.X & Chr(0)
        'Stat "Map Changed..." + vbCrLf
        PlayerMoveTime = 0
        'TmrMove.Enabled = False
        BlockMove = False
        Stat "Got Map Server IP." + vbCrLf
        CanusePath = True
        tmrTicks.Enabled = False
        CryptOn = False
        'AutoItem.TimeCount = 0
        Reset_Time
        ConnState = 3
        mapgot = MakeString(Mid(inData, 3, 16))
        CurrentMap = mapgot
        If Not tmrPortal.Enabled Then
            tmrPortal.Enabled = False
            tmrNomons.Enabled = True
            StopAction = False
        End If
        mapgot = Left(mapgot, Len(mapgot) - 4)
        CurMIP = MakeIP(Mid(inData, 23, 4))
        CurMPort = MakePort(Mid(inData, 27, 2))
        UpdateMIP mapgot, CurMIP, CurMPort
        eCase = eCase & "newMap=" & mapgot & Chr(0) & "newposX=" & curPos.Y & Chr(0) & "newposY=" & curPos.X
        MapName = mapgot
        Load_WayPoint mapgot
        Load_Field mapgot
        ResetMod
        If Not IsInLock Then
            MoveOnly = True
        Else
            MoveOnly = False
        End If
        UseWingYet = False
        OldDot.X = 0
        OldDot.Y = 0
        Stat "Map changed to [" & MapName & "]..." & vbCrLf
        'If MapName = "gl_prison" Then killsteal = True: Chat "System : [Options] - Kill-steal have been enabled in map 'gl_prison'", vbRed
        frmMain.Label1.Caption = "Main Status - " & GetMapname(MapName)
        If Not isUseHaunted Then
            Winsock1.Close
            Stat "Connecting to " + CurMIP & ":" & CurMPort & " ..."
            DoConnect CurMIP, CLng(CurMPort)
        End If
        ClearAll
        inData = ""
        Decode_0092 = ""
        CheckEvent "OnMapChange", eCase
        Exit Function
errie:
    Decode_0092 = "ERROR!!! [Decode_0092] " & Err.Description
    Err.Clear
    ResettoReCon
End Function

Private Function Decode_0097(inData As String) As String
On Error GoTo errie
    Dim tData As String
    tData = Mid(inData, 29, MakePort(Mid(inData, 3, 2)) - 28)
    'If Left(tData, 4) = "HEAL" Then
    '    Winsock_SendPacket Chr(&H13) + Chr(1) + Chr(HealLV) + Chr(0) + _
    '    Chr(&H1C) + Chr(0) + Mid(InData, 34, 4), True
    'ElseIf Left(tData, 4) = "AGII" Then
    '    Winsock_SendPacket Chr(&H13) + Chr(1) + Chr(10) + Chr(0) + _
    '        Chr(&H1D) + Chr(0) + Mid(InData, 33, 4), True
    'ElseIf Left(tData, 4) = "BLES" Then
    '    Winsock_SendPacket Chr(&H13) + Chr(1) + Chr(10) + Chr(0) + _
    '        Chr(&H22) + Chr(0) + Mid(InData, 34, 4), True
    'End If
    Chat "from " + MakeString(Mid(inData, 5, 24)) + " : " + tData, MColor.whisper
    ParseCommand MakeString(Mid(inData, 5, 24)), tData
    CheckEvent "OnPrivateMessage", "name=" & MakeString(Mid(inData, 5, 24)) & Chr(0) & "message=" & tData
    Decode_0097 = ""
Exit Function
errie:
Decode_0097 = "ERROR!!! [Decode_0097] " & Err.Description
Err.Clear
End Function

Private Function Decode_0098(inData As String) As String
On Error GoTo errie
        Select Case Asc(Mid(inData, 3, 1))
            Case 1
                Chat "System : [" & frmChat.txtWhisper & "] is not online...", MColor.whisper
            Case 2
                Chat "System : [" & frmChat.txtWhisper & "] block your message...", MColor.Fail
        End Select
    Decode_0098 = ""
Exit Function
errie:
Decode_0098 = "ERROR!!! [Decode_0098] " & Err.Description
Err.Clear
End Function

Private Function Decode_009A(inData As String) As String
On Error GoTo errie
    Chat Mid(inData, 5, MakePort(Mid(inData, 3, 2)) - 4) + vbCrLf, MColor.gmannounce
    CheckEvent "OnGMAnnounceMessage", "message=" & Mid(inData, 5, MakePort(Mid(inData, 3, 2)) - 4)
    Decode_009A = ""
Exit Function
errie:
Decode_009A = "ERROR!!! [Decode_009A] " & Err.Description
Err.Clear
End Function

Private Function Decode_009E(inData As String) As String
On Error GoTo errie
    Dim OKPos As Coord
            Dim isPick As Byte
            Dim Rare As Boolean
            OKPos.X = MakePort(Mid(inData, 12, 2))
            OKPos.Y = MakePort(Mid(inData, 10, 2))
            Dim Name As String
            Name = Return_ItemName(MakeHexName(Mid(inData, 7, 2)))
            isPick = Is_Pickup(Name, Get_AmountItem(Name))
            'Debug.Print "Item= [" & name & "] - Have " & Get_AmountItem(name) & " EA IsPick=" & CStr(isPick)
            Rare = isRare(Name)
            If ((EvalNorm(OKPos, curPos) < 3) Or (EvalNorm(OKPos, DeadMonPos) < 3)) _
                And (IsAutoPick) And (CurAtkMonster.NameID = 0 Or UBound(Aggro) > 0) And _
                (Pickup Or tmrPickDelay.Enabled) Then

                
                If isPick > 0 Then Exit Function
                Items(UBound(Items)).ID = Mid(inData, 3, 4)
                Items(UBound(Items)).Pos.X = MakePort(Mid(inData, 12, 2))
                Items(UBound(Items)).Pos.Y = MakePort(Mid(inData, 10, 2))
                Items(UBound(Items)).Name = MakeHexName(Mid(inData, 7, 2))
                ReDim Preserve Items(UBound(Items) + 1)
                If Not Rare Then
                    Stat "Found [" & Name & "] at " & CStr(OKPos.Y) & ":" & CStr(OKPos.X) & " (Yours...)" & vbCrLf
                Else
                    Stat "Found [" & Name & "] at " & CStr(OKPos.Y) & ":" & CStr(OKPos.X) & " (Rare...)" & vbCrLf
                    Stat "Quickly Pickup [" & Name & "]..." & vbCrLf
                    move_to Items(UBound(Items) - 1).Pos
                    Flood_Packet Chr(&H9F) + Chr(&H0) + Mid(inData, 3, 4), 100
                End If
end_bar:                 'If CurrentItem.Name = "" Then EstimateClosestItem
            End If
Decode_009E = ""
Exit Function
Decode_009E = "ERROR!!! [Decode_009E] " & Err.Description
errie:
Err.Clear
End Function

Private Function Decode_01C8(inData As String) As String
On Error GoTo errie
    If Mid(inData, 7, 4) = AccountID Then
        Dim Index&, amountleft&
        Index = MakePort(Mid(inData, 3, 2))
        amountleft = MakePort(Mid(inData, 11, 2))
        Stat "You're using [" & AllInv(Index).Name & "] " & CStr(AllInv(Index).Amount - amountleft) & "EA..." & vbCrLf
        AllInv(Index).Amount = amountleft
        UpdateInventory
        Dim i&
        For i = 0 To UBound(Cart)
            If Cart(i).Name = AllInv(Index).Name Then
                CheckCartInv i
                Exit For
            End If
        Next
    End If
end_if:
    Decode_01C8 = ""
Exit Function
errie:
Decode_01C8 = "ERROR!!! [Decode_01C8] " & Err.Description
Err.Clear
End Function

Private Function Decode_00A0(inData As String) As String
On Error GoTo errie
    'R 00a0 <index>.w <amount>.w <item ID>.w <identify flag>.B <attribute?>.B <refine>.B <card>.4w <equip type>.w <type>.B <fail>.B
    '               3                    5                      7                       9                           10                          11               12                20                       22                  23
    Dim Index&, gotAmount&
    Index = MakePort(Mid(inData, 3, 2))
    gotAmount = MakePort(Mid(inData, 5, 2))
    If MakePort(Mid(inData, 7, 2)) > 0 Then
            If Index > UBound(AllInv) Then ReDim Preserve AllInv(Index)
            AllInv(Index).NameID = MakePort(Mid(inData, 7, 2))
            AllInv(Index).Name = MakeItemName(Mid(inData, 7, 2), Mid(inData, 12, 8), Mid(inData, 11, 1))
            AllInv(Index).Amount = AllInv(Index).Amount + gotAmount
            AllInv(Index).Category = Asc(Mid(inData, 22, 1))
            AllInv(Index).Identified = CBool(Asc(Mid(inData, 23, 1)))
            If Not SendStore Or GetStore Then
                Stat "You got [" & AllInv(Index).Name & "], " & CStr(gotAmount) & " EA" & vbCrLf, vbBlue
                If Not CartBuy Then
                    mIsGoBuy = False
                Else
                    If ModAI Then Check_Destination_Route
                End If
                CheckCartAI MakePort(Mid(inData, 3, 2))
            Else
                Stat "Can't add [" & AllInv(Index).Name & "], " & CStr(gotAmount) + " EA to [Kafra]" + vbCrLf
                Dim i As Integer
                If UBound(Kafra) = 0 Then GoTo end_if
                For i = 0 To UBound(Kafra) - 1
                    If AllInv(MakePort(Mid(inData, 3, 2))).Name = Kafra(i).Name Then
                        Kafra(i).CantKeep = True
                        Exit For
                    End If
                Next
end_if:
            End If
            labCurMons.Caption = "[None]"
            UpdateInventory
            'CurrentItem.id = ""
            'CurrentItem.Name = ""
            GotItem = True
        End If
    Decode_00A0 = ""
Exit Function
errie:
Decode_00A0 = "ERROR!!! [Decode_00A0] " & Err.Description
Err.Clear
End Function

Private Function Decode_00A1(inData As String) As String
On Error GoTo errie
    Dim X, Y As Integer
    If Mid(inData, 3, 4) = CurItem.ID Then GotCurItem = False
        For X = 0 To UBound(Items) - 1
            If Mid(inData, 3, 4) = Items(X).ID Then
                Stat "[" & Return_ItemName(Items(X).Name) & "] disappeared" & vbCrLf
                For Y = X To UBound(Items) - 1
                    Items(Y) = Items(Y + 1)
                Next
                ReDim Preserve Items(UBound(Items) - 1)
                Exit For
            End If
            If Mid(inData, 3, 4) = CurrentItem.ID Then
                CurrentItem.ID = ""
                CurrentItem.Name = ""
                labCurMons.Caption = "[None]"
                If UBound(Items) > 0 Then EstimateClosestItem
            End If
        Next
    Decode_00A1 = ""
Exit Function
errie:
Decode_00A1 = "ERROR!!! [Decode_00A1] " & Err.Description
Err.Clear
End Function

Private Function Decode_00A8(inData As String) As String
On Error GoTo errie
        If AllInv(MakePort(Mid(inData, 3, 2))).Amount <> MakePort(Mid(inData, 5, 2)) Then
            AllInv(MakePort(Mid(inData, 3, 2))).Amount = MakePort(Mid(inData, 5, 2))
            CheckCartAI MakePort(Mid(inData, 3, 2))
            UpdateInventory
        End If
        Decode_00A8 = ""
Exit Function
errie:
Decode_00A8 = "ERROR!!! [Decode_00A8] " & Err.Description
Err.Clear
End Function

Private Function Decode_00AF(inData As String) As String
On Error GoTo errie
    If (UBound(AllInv) >= MakePort(Mid(inData, 3, 2))) Then
        If AllInv(MakePort(Mid(inData, 3, 2))).Name <> "" Then
            AllInv(MakePort(Mid(inData, 3, 2))).Amount = AllInv(MakePort(Mid(inData, 3, 2))).Amount - MakePort(Mid(inData, 5, 2))
            CheckCartAI MakePort(Mid(inData, 3, 2))
            If AllInv(MakePort(Mid(inData, 3, 2))).Amount < 0 Then AllInv(MakePort(Mid(inData, 3, 2))).Amount = 0
            If (SendSell) And Not SendStore Then
                Stat "Sold [" & AllInv(MakePort(Mid(inData, 3, 2))).Name & "] " + CStr(MakePort(Mid(inData, 5, 2))) + " EA" + vbCrLf
            End If
        End If
    End If
    UpdateInventory
Decode_00AF = ""
Exit Function
errie:
Decode_00AF = "ERROR!!! [Decode_00AF] " & Err.Description
Err.Clear
End Function

Private Function Decode_00B0(inData As String) As String
On Error GoTo errie
        Dim value As Long
        Dim ID As Integer
        ID = MakePort(Mid(inData, 3, 2))
        value = MakePort(Mid(inData, 5, 4))
        'If Mods.STDebug Then Chat "Debug : [00B0] - " & ID & ":" & CStr(value)z
        Select Case ID
            Case 0
                If value > 0 Then MovementSpeed = value
            Case 4
                If Abs(value) > 0 Then
                    Chat "Talk & Skill was disabled by GM for " & value & " minute(s). Auto-terminate program."
                    Open App.Path & "\log\warning.txt" For Append As #8
                        Print #8, Date & "@" & Time & ":  Talk & Skill was disabled by GM for " & Abs(value) & " minute(s). Auto-terminate program."
                    Close 8
                    ForceExit
                End If
            Case 5
                If value > Players(number).MaxHP Then
                    Players(number).HP = Players(number).MaxHP
                Else
                    Players(number).HP = value
                End If
                frmPlayer.LabHP.Caption = CStr(Players(number).HP) + "  /  " + CStr(Players(number).MaxHP)
                If (Players(number).MaxHP > 0) And (Players(number).HP >= 0) Then frmPlayer.tabHP.width = (Players(number).HP / Players(number).MaxHP) * (frmPlayer.tabHPBg.width - 20)
                If (Players(number).HP / Players(number).MaxHP > 0.25) Then
                   frmPlayer.tabHP.BackColor = &HC000&
                Else
                    frmPlayer.tabHP.BackColor = &HC0&
                End If
                CheckEvent "OnHPChange", "nothingtocheck=False"
            Case 6
                If value > Players(number).MaxHP Then
                    Players(number).MaxHP = MakePort(Mid(inData, 5, 2))
                    frmPlayer.LabHP.Caption = CStr(Players(number).HP) + "  /  " + CStr(Players(number).MaxHP)
                    If (Players(number).HP / Players(number).MaxHP > 0.25) Then
                       frmPlayer.tabHP.BackColor = &HC000&
                    Else
                        frmPlayer.tabHP.BackColor = &HC0&
                    End If
                    CheckEvent "OnHPChange", "nothingtocheck=False"
                End If
            Case 7
                Players(number).SP = value
                frmPlayer.labSP.Caption = CStr(Players(number).SP) + "  /  " + CStr(Players(number).maxsp)
                If (Players(number).maxsp > 0) And (Players(number).SP >= 0) Then frmPlayer.tabSP.width = (Players(number).SP / Players(number).maxsp) * (frmPlayer.tabSPbg.width - 20)
                CheckEvent "OnSPChange", "nothingtocheck=False"
            Case 8
                Players(number).maxsp = value
                frmPlayer.labSP.Caption = CStr(Players(number).SP) + "  /  " + CStr(Players(number).maxsp)
                CheckEvent "OnSPChange", "nothingtocheck=False"
            Case 9
                Players(number).StatPoint = value
                frmStat.labStatPt.Caption = CStr(Players(number).StatPoint)
            Case 11
                Players(number).BaseLV = value
                frmPlayer.labBaseLv.Caption = CStr(Players(number).BaseLV)
                CheckEvent "OnLevelUp", "baseLevel=" & CStr(value)
                GetScriptLockmap
            Case 12
                frmSkill.labPts = CStr(value)
            Case 24
                Players(number).Weight = CInt(value / 10)
                frmPlayer.labWeight.Caption = CStr(Players(number).Weight) + "  /  " + CStr(Players(number).MaxWeight)
                CheckEvent "OnWeightChange", "nothingtocheck=True"
            Case 25
                Players(number).MaxWeight = CInt(value / 10)
                frmPlayer.labWeight.Caption = CStr(Players(number).Weight) + "  /  " + CStr(Players(number).MaxWeight)
                CheckEvent "OnWeightChange", "nothingtocheck=True"
            Case 41
                Players(number).ATK = value
            Case 42
                Players(number).ATKp = value
            Case 43
                Players(number).MaxMatk = value
            Case 44
                Players(number).MinMatk = value
            Case 45
                Players(number).Def = value
            Case 46
                Players(number).Defp = value
            Case 47
                Players(number).mDef = value
            Case 48
                Players(number).mDefp = value
            Case 49
                Players(number).Hit = value
            Case 50
                Players(number).Flee = value
            Case 51
                Players(number).Fleep = value
            Case 52
                Players(number).Crit = value
            Case 53
                Players(number).Aspd = 200 - CInt(value / 10)
            Case 55
                Players(number).JobLV = value
                CheckEvent "OnLevelUp", "jobLevel=" & CStr(value)
                GetScriptLockmap
                Check_JobBar
                frmPlayer.labJobLv.Caption = CStr(Players(number).JobLV)
            End Select
        UpdateStats
        Decode_00B0 = ""
        Exit Function
errie:
Decode_00B0 = "ERROR!!! [Decode_00B0] " & Err.Description
Err.Clear
End Function

Private Function Decode_00B1(inData As String) As String
On Error GoTo errie
Dim percent As Double
        Dim currentEXP As Long
        currentEXP = 0
        'Stat "4D Number(" & CStr(Asc(Mid(InData, 3, 1))) & ") =" & CStr(MakePort(Mid(InData, 5, 4))) & vbCrLf
        If Asc(Mid(inData, 3, 1)) = 1 Then
            'currentEXP = getlong(Mid(InData, 5, 4))
            currentEXP = MakePort(Mid(inData, 5, 4))
            oldBaseEXP = Players(number).BaseExp
            SkillWait = False
            'If (CurAtkMonster.NameId > 0) Then Clear_This_Mons
            If (currentEXP < oldBaseEXP) Then
                SessionEXP = SessionEXP + currentEXP + (OldNextBEXP - oldBaseEXP)
                'SessionEXP = SessionEXP + currentEXP
                'Players(Number).BaseLV = Players(Number).BaseLV + 1
                'frmPlayer.labBaseLv = CStr(Players(Number).BaseLV)
            Else
                If (oldBaseEXP >= 0) Then
                    SessionEXP = SessionEXP + (currentEXP - oldBaseEXP)
                    CurEXPMons = currentEXP - oldBaseEXP
                End If
            End If
            Players(number).BaseExp = currentEXP
            If (Players(number).NextBaseEXP > 0) And (Players(number).BaseExp >= 0) Then
            percent = (Players(number).BaseExp / Players(number).NextBaseEXP)
            frmPlayer.tabBaseEXP.width = percent * (frmPlayer.labtabBaseEXPBg.width - 25)
            Else
                'Winsock_SendPacket IntToChr(&H7D)
            End If
            frmPlayer.labtabBaseEXPBg.ToolTipText = Format(Players(number).BaseExp, "##,##") & "/" & Format(Players(number).NextBaseEXP, "##,##") & " (" & FormatNumber(percent * 100, 2, vbTrue) + "%) "
        ElseIf Asc(Mid(inData, 3, 1)) = 9 Then
            Players(number).StatPoint = MakePort(Mid(inData, 5, 2))
            frmStat.labStatPt.Caption = CStr(Players(number).StatPoint)
        ElseIf Asc(Mid(inData, 3, 1)) = 20 Then
            Players(number).Zeny = MakePort(Mid(inData, 5, 4))
            frmPlayer.labZeny.Caption = Format(Players(number).Zeny, "##,##")
        ElseIf Asc(Mid(inData, 3, 1)) = 22 Then
            'print_packet Left(InData, 8)
            OldNextBEXP = Players(number).NextBaseEXP
            'Players(number).NextBaseEXP = getlong(Mid(InData, 5, 4))
            Players(number).NextBaseEXP = MakePort(Mid(inData, 5, 4))
            percent = (Players(number).BaseExp / Players(number).NextBaseEXP)
            frmPlayer.tabBaseEXP.width = percent * (frmPlayer.labtabBaseEXPBg.width - 25)
            frmPlayer.labtabBaseEXPBg.ToolTipText = Format(Players(number).BaseExp, "##,##") & "/" & Format(Players(number).NextBaseEXP, "##,##") & " (" & FormatNumber(percent * 100, 2, vbTrue) + "%) "
        ElseIf Asc(Mid(inData, 3, 1)) = 23 Then
            OldNextJXP = Players(number).MaxJobEXP
            'Players(number).MaxJobEXP = getlong(Mid(InData, 5, 4))
            Players(number).MaxJobEXP = MakePort(Mid(inData, 5, 4))
            percent = (Players(number).JobExp / Players(number).MaxJobEXP)
            frmPlayer.tabJobEXP.width = percent * (frmPlayer.labtabJobExpBg.width - 25)
            frmPlayer.labtabJobExpBg.ToolTipText = Format(Players(number).JobExp, "##,##") + "/" + Format(Players(number).MaxJobEXP, "##,##") + " (" & FormatNumber(percent * 100, 2, vbTrue) + "%) "
        ElseIf Asc(Mid(inData, 3, 1)) = 2 Then
            'currentEXP = getlong(Mid(InData, 5, 4))
            currentEXP = MakePort(Mid(inData, 5, 4))
            oldJobEXP = Players(number).JobExp
            If (currentEXP < oldJobEXP) Then
                SessionJEXP = SessionJEXP + currentEXP + (OldNextJXP - oldJobEXP)
                'SessionJEXP = SessionJEXP + currentEXP
                'Players(Number).JobLV = Players(Number).JobLV + 1
                'frmPlayer.labJobLv = CStr(Players(Number).JobLV)
            Else
                If (oldJobEXP >= 0) Then
                    SessionJEXP = SessionJEXP + (currentEXP - oldJobEXP)
                    CurJXPMons = currentEXP - oldJobEXP
                    'Chat "Killed [" & DeadMonsName & "] got [" & Format(CurEXPMons, "##,##") & "/" & Format(CurJXPMons, "##,##")
                    'Stat "You got [" & Format(currentEXP - oldJobEXP, "##,##") & "] JXP..." & vbCrLf
                End If
            End If
            Players(number).JobExp = currentEXP
            'MDIfrmMain.StatusBar1.SimpleText = "Session Time : [" + MakeTime + "], Session EXP/JXP : [" + Format(SessionEXP, "##,##") & "/" & _
            'Format(SessionJEXP, "##,##") & "],Session Zeny : " & Format(Players(number).Zeny - StartZeny, "##,##") & " ,Last Monster : '" & DeadMonsName & "' [" & Format(CurEXPMons, "##,##") & "/" & Format(CurJXPMons, "##,##") & "]"
            If Players(number).MaxJobEXP > 0 And (Players(number).JobExp >= 0) Then
            percent = (Players(number).JobExp / Players(number).MaxJobEXP)
            frmPlayer.tabJobEXP.width = percent * (frmPlayer.labtabJobExpBg.width - 25)
            frmPlayer.labtabJobExpBg.ToolTipText = Format(Players(number).JobExp, "##,##") + "/" + Format(Players(number).MaxJobEXP, "##,##") + " (" & FormatNumber(percent * 100, 2, vbTrue) + "%) "
            Else
                'Winsock_SendPacket IntToChr(&H7D)
            End If
        Else
            'Chat "Unknow stat(&HB1 - " & Asc(Mid(InData, 3, 1)) & "): " & GetLong(Mid(InData, 5, 4))
        End If
    Decode_00B1 = ""
Exit Function
errie:
Decode_00B1 = "ERROR!!! [Decode_00B1] " & Err.Description
Err.Clear
End Function

Private Function Decode_00B4(inData As String) As String
On Error GoTo errie
    Dim message As String
    'Stat "Decode 00B4" & vbCrLf
    CurNPC = Mid(inData, 5, 4)
    message = Trim(Mid(inData, 9, MakePort(Mid(inData, 3, 2)) - 8))
    If Not AutoAI Then
        frmNPCMessage.Visible = True
        If Left(message, 1) = "[" Then
            frmNPCMessage.LabName.Caption = message
            frmNPCMessage.txtNPC.text = ""
        Else
            frmNPCMessage.txtNPC.text = frmNPCMessage.txtNPC.text & message
        End If
    End If
 Decode_00B4 = ""
Exit Function
errie:
Decode_00B4 = "ERROR!!! [Decode_00B4] " & Err.Description
Err.Clear
End Function

Private Function Decode_00B5(inData As String) As String
On Error GoTo errie
    CurNPC = Mid(inData, 3, 4)
    Dim Index As Integer
    If Not AutoAI Then
        frmNPCMessage.imgNext.Visible = True
        frmNPCMessage.btClose.Visible = False
        'frmSelectBuySell.Visible = True
        GoTo end_sub
    End If
    
    'If HaveSellItem Then
    '    Stat "Send Select Sell..." & vbCrLf
    '    tmrDealNPC.Enabled = False
    '    tmrDealNPC.Enabled = True
    '    Send_Sell
    '    Send_TalkCancel
    If npc_step <> "" Then
        Stat "Send talk continue..." & vbCrLf
        tmrDealNPC.Enabled = False
        tmrDealNPC.Enabled = True
        Send_TalkContinue
        Index = InStr(npc_step, " ")
        If Index > 0 Then npc_step = Right(npc_step, Len(npc_step) - Index)
    ElseIf HaveStoreItem Or HaveGetStorageItem Then
        tmrDealNPC.Enabled = False
        tmrDealNPC.Enabled = True
        Send_TalkContinue
    End If
end_sub:
Decode_00B5 = ""
Exit Function
errie:
Decode_00B5 = "ERROR!!! [Decode_00B5] " & Err.Description
Err.Clear
End Function

Private Function Decode_00B6(inData As String) As String
On Error GoTo errie
    CurNPC = Mid(inData, 3, 4)
    If Not AutoAI Then
        frmNPCMessage.btClose.Visible = True
         frmNPCMessage.imgNext.Visible = False
         GoTo end_sub
        'frmSelectBuySell.Visible = True
    End If
    
    If npc_step <> "" Then
        Stat "Your [action_script] is not correct or need to complete some quest?" & vbCrLf
        tmrDealNPC.Enabled = False
        tmrDealNPC.Enabled = True
    ElseIf HaveStoreItem Or HaveGetStorageItem Then
        tmrDealNPC.Enabled = False
        tmrDealNPC.Enabled = True
    End If
    Send_TalkCancel
    
end_sub:
Decode_00B6 = ""
Exit Function
errie:
Decode_00B6 = "ERROR!!! [Decode_00B6] " & Err.Description
Err.Clear
End Function

Private Function Decode_00B7(inData As String) As String
On Error GoTo errie
    Dim Index As Integer
    If Not AutoAI Then
        Dim ListText As String
        ListText = Mid(inData, 9, MakePort(Mid(inData, 3, 2)) - 8)
        Index = InStr(ListText, ":")
        frmNPCSelect.Visible = True
        frmNPCSelect.lstEvent.Clear
        frmNPCSelect.LabName.Caption = frmNPCMessage.LabName.Caption & " -> Select"
        Do While Index > 0
            frmNPCSelect.lstEvent.AddItem Trim(Left(ListText, Index - 1))
            ListText = Right(ListText, Len(ListText) - Index)
            Index = InStr(ListText, ":")
        Loop
    ElseIf npc_step <> "" Then
        Index = InStr(npc_step, " ")
        Dim tstr2 As String
        
        If Index > 0 Then
            tstr2 = Left(npc_step, Index)
            npc_step = Right(npc_step, Len(npc_step) - Index)
        Else
            tstr2 = npc_step
            npc_step = ""
        End If
        Stat "Send Talk select choice[" & Val(Right(tstr2, Len(tstr2) - 1)) & "]..." & vbCrLf
        Send_TalkResponse Val(Right(tstr2, Len(tstr2) - 1))
    ElseIf HaveStoreItem Or HaveGetStorageItem Then
        tmrDealNPC.Enabled = False
        tmrDealNPC.Enabled = True
        Send_TalkResponse 2
    End If
Decode_00B7 = ""
Exit Function
errie:
Decode_00B7 = "ERROR!!! [Decode_00B7] " & Err.Description
Err.Clear
End Function

Private Function Decode_00BD(inData As String) As String
On Error GoTo errie
    Players(number).StatPoint = MakePort(Mid(inData, 3, 2))
    Players(number).STR = Asc(Mid(inData, 5, 1))
    frmStat.labStrp.Caption = Asc(Mid(inData, 6, 1))
    Players(number).AGI = Asc(Mid(inData, 7, 1))
    frmStat.LabAgip.Caption = Asc(Mid(inData, 8, 1))
    Players(number).VIT = Asc(Mid(inData, 9, 1))
    frmStat.LabVitp.Caption = Asc(Mid(inData, 10, 1))
    Players(number).Intl = Asc(Mid(inData, 11, 1))
    frmStat.LabIntp.Caption = Asc(Mid(inData, 12, 1))
    Players(number).DEX = Asc(Mid(inData, 13, 1))
    frmStat.LabDexp.Caption = Asc(Mid(inData, 14, 1))
    Players(number).LUK = Asc(Mid(inData, 15, 1))
    frmStat.LabLuckp.Caption = Asc(Mid(inData, 16, 1))
    Players(number).ATK = MakePort(Mid(inData, 17, 2))
    Players(number).ATKp = MakePort(Mid(inData, 19, 2))
    Players(number).MinMatk = MakePort(Mid(inData, 21, 2))
    Players(number).MaxMatk = MakePort(Mid(inData, 23, 2))
    Players(number).Def = MakePort(Mid(inData, 25, 2))
    Players(number).Defp = MakePort(Mid(inData, 27, 2))
    Players(number).mDef = MakePort(Mid(inData, 29, 2))
    Players(number).mDefp = MakePort(Mid(inData, 31, 2))
    Players(number).Hit = MakePort(Mid(inData, 33, 2))
    Players(number).Flee = MakePort(Mid(inData, 35, 2))
    Players(number).Fleep = MakePort(Mid(inData, 37, 2))
    Players(number).Crit = MakePort(Mid(inData, 39, 2))
    UpdateStats
Decode_00BD = ""
Exit Function
errie:
Decode_00BD = "ERROR!!! [Decode_00BD] " & Err.Description
Err.Clear
End Function

Private Function Decode_00BE(inData As String) As String
    On Error GoTo errie
    Dim ID As Integer
    Dim Point As Byte
    ID = MakePort(Mid(inData, 3, 2))
    Point = Asc(Mid(inData, 5, 1))
    Select Case ID
        Case 32
            frmStat.labStrp.Caption = Point
        Case 33
            frmStat.LabAgip.Caption = Point
        Case 34
            frmStat.LabVitp.Caption = Point
        Case 35
            frmStat.LabIntp.Caption = Point
        Case 36
            frmStat.LabDexp.Caption = Point
        Case 37
            frmStat.LabLuckp.Caption = Point
    End Select
    UpdateStats
Decode_00BE = ""
Exit Function
errie:
Decode_00BE = "ERROR!!! [Decode_00BE] " & Err.Description
Err.Clear
End Function

Private Function Decode_00C0(inData As String) As String
On Error GoTo errie
    If Asc(Mid(inData, 7, 1)) <= UBound(Emotions) Then
        If Mid(inData, 3, 4) = AccountID Then
            Chat Players(number).Name & " : Send Emoticon [*" & Emotions(Asc(Mid(inData, 7, 1))).detail & "*]", MColor.Emotion
        Else
            Dim Name As String, txtChats$, i& ', ChrID As Long
            'ChrID = Conv4B(Mid(InData, 3, 4))
            Name = ""
            Name = Get_PeopleName(Mid(inData, 3, 4))
            If Name = "Unknow" Then
                Name = Get_MonsName(Mid(inData, 3, 4))
                For i = LBound(MonsterList) To UBound(MonsterList)
                    If Mid(inData, 3, 4) = MonsterList(i).ID Then
                        txtChats = "[" & EvalNorm(curPos, MonsterList(i).Pos) & " blks] "
                        Exit For
                    End If
                Next
            Else
                For i = LBound(People) To UBound(People)
                    If Mid(inData, 3, 4) = People(i).ID Then
                        txtChats = "[" & EvalNorm(curPos, People(i).Pos) & " blks] "
                        Exit For
                    End If
                Next
            End If
            'Chat txtChats, vbBlue
            If Mods.EmotionText Then Chat txtChats & Name & " : Send Emoticon [*" & Emotions(Asc(Mid(inData, 7, 1))).detail & "*]", MColor.Emotion
        End If
    'Else
    '    Chat Players(number).name & " : Send unknown Emoticon [" & CStr(Asc(Mid(InData, 7, 1))) & "]"
    End If
Decode_00C0 = ""
Exit Function
errie:
Decode_00C0 = "ERROR!!! [Decode_00C0] " & Err.Description
Err.Clear
End Function

Private Function Decode_00C2(inData As String) As String
On Error GoTo errie
    Chat "System : [" & Format(MakePort(Mid(inData, 3, 2)), "##,##") & "] connected to this server...", &H4080FF
Decode_00C2 = ""
Exit Function
errie:
Decode_00C2 = "ERROR!!! [Decode_00C2] " & Err.Description
Err.Clear
End Function
Private Function Decode_00C4(inData As String) As String
On Error GoTo errie
    CurNPC = Mid(inData, 3, 4)
    'Stat "Found your Tool Dealer..." & vbCrLf
    If Not AutoAI Then
        frmSelectBuySell.Visible = True
    ElseIf HaveSellItem Then
        tmrDealNPC.Enabled = False
        tmrDealNPC.Enabled = True
        Stat "Send request to sell item" & vbCrLf, &HAAAAAA
        Send_Sell
    ElseIf HaveBuyItem Then
        tmrDealNPC.Enabled = False
        tmrDealNPC.Enabled = True
        Stat "Send request to buy item" & vbCrLf, &HAAAAAA
        Send_GetStoreList
    End If
Decode_00C4 = ""
Exit Function
errie:
Decode_00C4 = "ERROR!!! [Decode_00C4] " & Err.Description
Err.Clear
End Function

Private Function Decode_00D2(inData As String) As String
On Error GoTo errie
    If Asc(Mid(inData, 3, 1)) = 0 Then
        If Mods.STSystem Then Chat "System : [Whisper], Block all message...", MColor.Fail
    ElseIf Asc(Mid(inData, 3, 1)) = 1 Then
        If Mods.STSystem Then Chat "System : [Whisper], Welcome all message...", MColor.Fail
    End If
Decode_00D2 = ""
Exit Function
errie:
Decode_00D2 = "ERROR!!! [Decode_00D2] " & Err.Description
Err.Clear
End Function

Private Function Decode_00D6() As String
On Error GoTo errie
    Chat "System : [chat room] Created...", MColor.Shop
    IsChatOC = True
Decode_00D6 = ""
Exit Function
errie:
Decode_00D6 = "ERROR!!! [Decode_00D6] " & Err.Description
Err.Clear
End Function

Private Function Decode_00F2(inData As String) As String
On Error GoTo errie
    Dim i&
    frmStorage.labNumber.Caption = CStr(MakePort(Mid(inData, 3, 2))) & "/" & CStr(MakePort(Mid(inData, 5, 2)))
    NoStoreItem = IIf(MakePort(Mid(inData, 3, 2)) > 0, False, True)
    For i = 0 To UBound(Cart)
        If Cart(i).Amount > 0 Then CheckCartStore CLng(i)
    Next
    MIsGoStore = False
Exit Function
Decode_00F2 = ""
Exit Function
errie:
Decode_00F2 = "ERROR!!! [Decode_00F2] " & Err.Description
Err.Clear
End Function

Private Function Decode_00F8() As String
On Error GoTo errie
    Unload frmStorage
Decode_00F8 = ""
Exit Function
errie:
Decode_00F8 = "ERROR!!! [Decode_00F8] " & Err.Description
Err.Clear
End Function



Private Function Decode_0101(inData As String) As String
On Error GoTo errie
    If Mods.STParty Then
        Select Case Asc(Mid(inData, 3, 1))
            Case 0
                Chat "System : [Party EXP] set to individual share...", &H4080FF
            Case 1
                Chat "System : [Party EXP] set to even share...", &H4080FF
            Case 2
                Chat "System : Can't set [Party EXP]...", &H4080FF
        End Select
    End If
Decode_0101 = ""
Exit Function
errie:
Decode_0101 = "ERROR!!! [Decode_0101] " & Err.Description
Err.Clear
End Function



Private Function Decode_010E(inData As String) As String
On Error GoTo errie
    Dim X As Integer
    For X = 0 To UBound(SkillChar) - 1
        If MakePort(Mid(inData, 3, 2)) = SkillChar(X).ID Then
            SkillChar(X).MaxLV = Asc(Mid(inData, 5, 1))
            UpdateSkills
            Exit For
        End If
    Next
Decode_010E = ""
Exit Function
errie:
Decode_010E = "ERROR!!! [Decode_010E] " & Err.Description
Err.Clear
End Function

Function Decode_0117(inData As String) As String
On Error GoTo errie
'R 0117 <skill ID>.w <src ID>.l <val>.w <X>.w <Y>.w <server tick>.l
'               3                   5                   9               11      13      15
'isulandskill
Dim X&, Y&, Val&, Src$, SkillID&, srcname$, SkillName$
SkillID = MakePort(Mid(inData, 3, 2))
Src = Mid(inData, 5, 4)
If MakePort(Src) = MakePort(AccountID) Then srcname = "You" Else srcname = "[" & Get_PeopleName(Src) & "]"
If SkillID - 1 < UBound(SkillIDName) Then SkillName = "[" & SkillIDName(SkillID - 1).Name & "]" Else SkillName = "[Unknown]"
Val = MakePort(Mid(inData, 9, 2))
X = MakePort(Mid(inData, 11, 2))
Y = MakePort(Mid(inData, 13, 2))
If Src = AccountID Then
    If (Not MakeDamage) Then
        Stat "You locked, [" + CurAtkMonster.Name + "] as a Target..." + vbCrLf
        IsLock = True
    End If
    MakeDamage = True
    SkillWait = False
    tmrSkillDelay.Enabled = False
    SkillCounter = SkillCounter + 1
    Check_ResetCounter
    Casting = False
    DamageCounter = 0
    AttackCounter = 0
    AttCounter = AttCounter + 1
End If
Stat srcname & " use skill " & SkillName & " at " & CStr(X) & ":" & CStr(Y) + vbCrLf
Exit Function
errie:
Decode_0117 = "ERROR!!! [Decode_0117] " & Err.Description
Err.Clear
End Function

Private Function Decode_0115(inData As String) As String
On Error GoTo errie
    If Asc(Mid(inData, 3, 1)) = 26 Or Asc(Mid(inData, 3, 1)) = 27 Then
        Chat "Someone make Warp Portal at" + CStr(MakeCoords(Mid(inData, 11, 3)).X) + ":" + CStr(MakeCoords(Mid(inData, 11, 3)).Y) + ", Avoid it..." + vbCrLf
        Teleport
        Exit Function
    End If
    If (CurMonsterName) <> "" And Mid(inData, 5, 4) = AccountID Then
        If (Not MakeDamage) Then
            Stat "You locked, [" + CurMonsterName + "] as a Target..." + vbCrLf
            IsLock = True
        End If
        If CStr(MakePort(Mid(inData, 29, 2))) > 0 And CStr(MakePort(Mid(inData, 29, 2))) < 20000 Then
            Stat "[" + frmSkill.Return_SkillName(Asc(Mid(inData, 3, 1))) + "] Skill to [" + Return_MonsterName(CurAtkMonster.NameID) + "], " + CStr(MakePort(Mid(inData, 29, 2))) + " Damage" + vbCrLf
            MakeDamage = True
            SkillCounter = SkillCounter + 1
            Check_ResetCounter
        ElseIf CStr(MakePort(Mid(inData, 29, 2))) = 0 Then
            Stat "[" + frmSkill.Return_SkillName(Asc(Mid(inData, 3, 1))) + "] Skill to [" + Return_MonsterName(CurAtkMonster.NameID) + "], " + "Miss!" + vbCrLf
        End If
        DamageCounter = 0
        AttackCounter = 0
        AttCounter = AttCounter + 1
    End If
Decode_0115 = ""
Exit Function
errie:
Decode_0115 = "ERROR!!! [Decode_0115] " & Err.Description
Err.Clear
End Function

Public Sub Auto_Recovery(Name As String)
On Error GoTo errie
    Dim i As Integer
    Dim X As Long
    If delay_count > 0 Then Exit Sub
    For i = 0 To UBound(ai_recovery)
        If InStr(ai_recovery(i).Name, Name) > 0 Then
            X = Find_HealItem(ai_recovery(i).RecovItem)
            If X > 0 Then
                Stat "Found [" & AllInv(X).Name & "] to recovery [" & Name & "]" & vbCrLf
                Winsock_SendPacket IntToChr(&HA7) & IntToChr(X) & AccountID, True
                delay_count = delay_recovery
            Else
                X = find_skill_id(ai_recovery(i).RecovSkill)
                If X > 0 Then
                    Stat "Found [" & SkillIDName(X).Name & "] to recovery [" & Name & "]" & vbCrLf
                    Send_Use_Skill X, 1, AccountID
                    delay_count = delay_recovery
                End If
            End If
        End If
    Next
errie:
Err.Clear
End Sub

Private Function Decode_0119(inData As String) As String
On Error GoTo errie
    'AI_AvoidID Mid(InData, 3, 4), "0119"
    Dim EvStat$
    EvStat = "isAvoidID=" & CStr(IsAvoidID(Mid(inData, 3, 4))) & Chr(0)
    If Mid(inData, 3, 4) = AccountID Then EvStat = EvStat & "name=You" & Chr(0) Else EvStat = EvStat & "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0)
    EvStat = EvStat & "AID=" & MakePort(Mid(inData, 3, 4)) & Chr(0)
    EvStat = EvStat & "P1=" & MakePort(Mid(inData, 7, 2)) & Chr(0)
    EvStat = EvStat & "P2=" & MakePort(Mid(inData, 9, 2)) & Chr(0)
    EvStat = EvStat & "P3=" & MakePort(Mid(inData, 11, 2))
    If (Mid(inData, 3, 4) = AccountID) Then
        'R 0119 <ID>.l <param1>.w <param2>.w <param3>.w ?.B
        '               3           7                   9                            11                     13
        Dim Index As Integer
        Index = MakePort(Mid(inData, 9, 2))
        If Index > UBound(CurCharStatus) Then ReDim Preserve CurCharStatus(Index)
        If CurCharStatus(Index).Name <> "" Then
            frmPlayer.labStatus.Caption = CurCharStatus(Index).Name
            If use_recovery_profile And Index > 0 Then Auto_Recovery CurCharStatus(Index).Name
        Else
            frmPlayer.labStatus.Caption = "Unknow_" & CStr(Index)
        End If
        Index = MakePort(Mid(inData, 11, 2))
        CheckCart Index
        If Mods.STDebug Then Chat "Debug : [0119] - param1:" & MakePort(Mid(inData, 7, 2)) & ", param2:" & MakePort(Mid(inData, 9, 2)) & ", param3:" & Index & ", param4:" & Asc(Mid(inData, 9, 1))
    Else
        Dim i As Long
        For i = LBound(MonsterList) To UBound(MonsterList)
            If MonsterList(i).ID = Mid(inData, 3, 4) Then
                MonsterList(i).StatusA = MakePort(Mid(inData, 7, 2))
                If Mid(inData, 3, 4) = CurAtkMonster.ID Then
                    CurAtkMonster.StatusA = MakePort(Mid(inData, 7, 2))
                    Stat MonsterList(i).Name & " is (-" & MonsterList(i).StatusA & "-)" + vbCrLf, &HCCCCCC
                    Exit For
                End If
                Stat MonsterList(i).Name & " is (" & MonsterList(i).StatusA & ")" + vbCrLf, &HCCCCCC
                Exit For
            End If
        Next
    End If
    CheckEvent "OnEffectChange", EvStat
    Decode_0119 = ""
Exit Function
errie:
Decode_0119 = "ERROR!!! [Decode_0119] " & Err.Description
Err.Clear
End Function

Private Function Decode_011A(inData As String) As String
On Error GoTo errie
    Dim SkillName As String
    SkillName = "Unknown"
    Dim i&, ID  As Integer, value As Long
    value = MakePort(Mid(inData, 5, 2))
    ID = MakePort(Mid(inData, 3, 2))
    If MakePort(Mid(inData, 3, 2)) - 1 < UBound(SkillIDName) Then SkillName = SkillIDName(MakePort(Mid(inData, 3, 2)) - 1).Name
    
    If (Mid(inData, 11, 4) = AccountID) And ((Mid(inData, 7, 4) = AccountID) Or (Mid(inData, 7, 4) = String$(4, 0))) Then
        Stat "You're using skill [" & SkillName & "].." & vbCrLf
        For i = 0 To UBound(AutoSkill)
            If ID = AutoSkill(i).ID Then
                AutoSkill(i).TimeCount = Int(GetTickCount() / 1000)
                DelaySelfSkill = 0
                Exit Function
            End If
        Next
    'R 011a <skill ID>.w <val>.w <dst ID>.l <src ID>.l <fail>.B
    '               3                       5               7               11              15
    ElseIf Mid(inData, 7, 4) = CurAtkMonster.ID And Mid(inData, 11, 4) = AccountID And CurAtkMonster.NameID <> 0 Then
        SkillCounter = SkillCounter + 1
        Stat "You use skill [" & SkillName & ",Lv:" & value & "][" & SkillCounter & "] to [" & CurAtkMonster.Name & "]..." & vbCrLf
        MakeDamage = True
'    ElseIf Mid(InData, 7, 4) = CurAtkMonster.ID And Mid(InData, 11, 4) <> AccountID And CurAtkMonster.NameID <> 0 Then
'        If Value > 0 And ID = 28 Then
'            Dim res&
'            res = -1
'            For i = 0 To UBound(People)
'                If People(i).ID = Mid(InData, 11, 4) Then
'                    res = i
'                    Exit For
'                End If
'            Next
'            If res = -1 Then
'                For i = 0 To UBound(MonsterList)
'                    If MonsterList(i).ID = Mid(InData, 11, 4) Then
'                        res = i
'                        Exit For
'                    End If
'                Next
'                If res = -1 Then
'                    Chat "[Unknown] heal your monster - HP Gained [" & Value & "]"
'                    CheckEvent "OnPeopleHealYourMonster", "name=Unknown" & Chr(0) & "hpgain=" & Value
'                Else
'                    Chat "[" & MonsterList(i).Name & "] heal your monster - HP Gained [" & Value & "]"
'                    CheckEvent "OnMonsterHealYourMonster", "name=" & MonsterList(i).Name & Chr(0) & "hpgain=" & Value
'                End If
'            Else
'                Chat "[" & IIf(People(res).Name <> "", People(res).Name, "Unknown") & "] heal your monster - HP Gained [" & Value & "]"
'                CheckEvent "OnPeopleHealYourMonster", "name=" & IIf(People(res).Name <> "", People(res).Name, "Unknown") & Chr(0) & "hpgain=" & Value
'            End If
'        End If
    End If
    Decode_011A = ""
Exit Function
errie:
Decode_011A = "ERROR!!! [Decode_011A] " & Err.Description
Err.Clear
End Function

Private Function Decode_013A(inData As String) As String
On Error GoTo errie
    PRange = MakePort(Mid(inData, 3, 2))
    Exit Function
errie:
Decode_013A = "ERROR!!! [Decode_013A] " & Err.Description
Err.Clear
End Function

Function Decode_013C(inData As String) As String
On Error GoTo errie
    'R 013c <ID>.w
    ArrowNumber = MakePort(Mid(inData, 3, 2))
    If UBound(AllInv) >= ArrowNumber Then Stat "You equipped arrow :[" & AllInv(ArrowNumber).Name & "]" & vbCrLf
Exit Function
errie:
Decode_013C = "ERROR!!! [Decode_013C] " & Err.Description
Err.Clear
End Function

Private Function Decode_013D(inData As String) As String
On Error GoTo errie
    If Asc(Mid(inData, 3, 1)) = 5 Then
        If Players(number).HP + MakePort(Mid(inData, 5, 2)) <= Players(number).MaxHP Then
            Players(number).HP = Players(number).HP + MakePort(Mid(inData, 5, 2))
        Else
            Players(number).HP = Players(number).MaxHP
        End If
        frmPlayer.LabHP.Caption = CStr(Players(number).HP) + "  /  " + CStr(Players(number).MaxHP)
        If (Players(number).MaxHP > 0) And (Players(number).HP >= 0) Then frmPlayer.tabHP.width = (Players(number).HP / Players(number).MaxHP) * (frmPlayer.tabHPBg.width - 20)
    End If
Decode_013D = ""
Exit Function
errie:
Decode_013D = "ERROR!!! [Decode_013D] " & Err.Description
Err.Clear
End Function

Private Function Decode_01D0(inData As String) As String
'spirit sphere update
' R 01D0 <acc.id>.l <ball>.w
On Error GoTo errie
    If Mid(inData, 3, 4) = AccountID Then CurSpirit = Asc(Mid(inData, 7, 1))
Decode_01D0 = ""
Exit Function
errie:
Decode_01D0 = "ERROR!!! [Decode_01D0] " & Err.Description
Err.Clear
End Function

Private Function Decode_013E(inData As String) As String
On Error GoTo errie
    Dim X&, Y&, i As Integer
    Dim Src, Target, SkillName As String
    Dim srcType&, desType&, skilldist&, SpellPos As Coord, APos As Coord
    Y = MakePort(Mid(inData, 11, 2))
    X = MakePort(Mid(inData, 13, 2))
    
    If (Mid(inData, 3, 4) = AccountID) Then
        Src = "You"
        srcType = 0
    ElseIf Get_MonsName(Mid(inData, 3, 4)) <> "Unknow" Then
        Src = "[" & Get_MonsName(Mid(inData, 3, 4)) & "]"
        srcType = 2
        For i = 0 To UBound(MonsterList) - 1
            If MonsterList(i).ID = Mid(inData, 3, 4) Then
                MonsterList(i).TargetID = Mid(inData, 7, 4)
                If Mid(inData, 7, 4) <> AccountID Then MonsterList(i).IsAttack = True
                Exit For
            End If
        Next
    ElseIf Get_PeopleName(Mid(inData, 3, 4)) <> "Unknow" Then
        Src = "[" & Get_PeopleName(Mid(inData, 3, 4)) & "]"
        srcType = 1
    Else
        Src = "[Unknow]"
    End If

    'mc new ai
    'srctype :  0 - you
    '                   1 - people
    '                   2 - monster
    'destype:  0 - you
    '                   1 - people
    '                   2 - monster
    '                   3 - ground
    '                   4 - unknown
    
    If MakePort(Mid(inData, 15, 2)) - 1 < UBound(SkillIDName) Then SkillName = SkillIDName(MakePort(Mid(inData, 15, 2)) - 1).Name
    Target = ""
    If (Mid(inData, 7, 4) = AccountID) Then
        Target = "You"
        If Src = "You" Then Target = "Yourself"
        desType = 0
    ElseIf Get_MonsName(Mid(inData, 7, 4)) <> "Unknow" Then
        Target = Target & "[" & Get_MonsName(Mid(inData, 7, 4)) & "]"
        desType = 2
    ElseIf Left(Get_PeopleName(Mid(inData, 7, 4)), 2) <> "U:" Then
        Target = Target & "[" & Get_PeopleName(Mid(inData, 7, 4)) & "]"
        desType = 1
    ElseIf X <> 0 And Y <> 0 Then
        Target = Target & "(" & CStr(Y) & ":" & CStr(X) & ")"
        desType = 3
        SpellPos.X = X
        SpellPos.Y = Y
        skilldist = EvalNorm(SpellPos, curPos)
        If Mid(inData, 3, 4) = AccountID Then isulandskill = True
        GoTo skip:
    Else
        desType = 4
        Target = "[U:" & MakePort(Mid(inData, 7, 4)) & "]"
    End If
    If Mid(inData, 3, 4) = Mid(inData, 7, 4) Then GoTo skip
    checkKS Mid(inData, 3, 4), Mid(inData, 7, 4)
skip:
    Stat Src & " is casting skill [" & SkillName & "] on " & Target & vbCrLf
    
    If srcType = 1 And desType = 3 Then
        If SpellPos.Y - curPos.Y >= 0 And SpellPos.X - curPos.X > 0 Then
        'spell on top or topright
            APos.Y = curPos.Y + RandomNumber(-4, -8)
            APos.X = curPos.X + RandomNumber(-4, -8)
        ElseIf SpellPos.Y - curPos.Y > 0 And SpellPos.X - curPos.X <= 0 Then
        'spell on right or bottomright
            APos.Y = curPos.Y + RandomNumber(-4, -8)
            APos.X = curPos.X + RandomNumber(8, 4)
        ElseIf SpellPos.Y - curPos.Y < 0 And SpellPos.X - curPos.X >= 0 Then
        'spell on left or topleft
            APos.Y = curPos.Y + RandomNumber(8, 4)
            APos.X = curPos.X + RandomNumber(-4, -8)
        ElseIf SpellPos.Y - curPos.Y <= 0 And SpellPos.X - curPos.X < 0 Then
        'spell on bottom or bottomleft
            APos.X = curPos.X + RandomNumber(8, 4)
            APos.Y = curPos.Y + RandomNumber(8, 4)
        End If
        'pos checking
        If skilldist < 3 Then
            If GSonyou Then
                move_to APos
                Chat "Avoid : Ground skill on your current position [move to (" & APos.Y & ":" & APos.X & ")]", vbRed
            End If
        Else
            If GSnearyou Then
                move_to APos
                Chat "Avoid : Ground skill near your current position [move to (" & APos.Y & ":" & APos.X & ")]", vbRed
            End If
        End If
    End If
    If srcType = 2 And desType = 3 Then
        If SpellPos.Y - curPos.Y >= 0 And SpellPos.X - curPos.X > 0 Then
        'spell on top or topright
            APos.Y = curPos.Y + RandomNumber(-4, -8)
            APos.X = curPos.X + RandomNumber(-4, -8)
            move_to APos
        ElseIf SpellPos.Y - curPos.Y > 0 And SpellPos.X - curPos.X <= 0 Then
        'spell on right or bottomright
            APos.Y = curPos.Y + RandomNumber(-4, -8)
            APos.X = curPos.X + RandomNumber(8, 4)
            move_to APos
        ElseIf SpellPos.Y - curPos.Y < 0 And SpellPos.X - curPos.X >= 0 Then
        'spell on left or topleft
            APos.Y = curPos.Y + RandomNumber(8, 4)
            APos.X = curPos.X + RandomNumber(-4, -8)
            move_to APos
        ElseIf SpellPos.Y - curPos.Y <= 0 And SpellPos.X - curPos.X < 0 Then
        'spell on bottom or bottomleft
            APos.X = curPos.X + RandomNumber(8, 4)
            APos.Y = curPos.Y + RandomNumber(8, 4)
            move_to APos
        End If
        If skilldist < 3 Then
            If MGSonyou Then
                move_to APos
                Chat "Avoid : Monster ground skill on your current position [move to (" & APos.Y & ":" & APos.X & ")]", vbRed
            End If
        Else
            If MGSnearyou Then
                move_to APos
                Chat "Avoid : Monster ground skill near your current position [move to (" & APos.Y & ":" & APos.X & ")]", vbRed
            End If
        End If
    End If
    
    If Mid(inData, 3, 4) = AccountID And (CurAtkMonster.NameID > 0) Then
        Casting = True
        If ((SpellCounter > 1 And (Not MakeDamage)) Or SpellCounter > 3) And Not isulandskill Then
            Stat "Problem with this monster, switched..." & vbCrLf
            If UBound(MonsterList) > 0 Then
                For i = 0 To UBound(MonsterList) - 1
                    If (MonsterList(i).ID = CurAtkMonster.ID) Then
                        MonsterList(i).IsAttack = True
                        Exit For
                    End If
                Next
                Clear_This_Mons 0
            End If
            CurAtkMonster.NameID = 0
            NumberMons = 0
            labCurMons.Caption = "[None]"
            If UBound(MonsterList) > 0 Then EstimateClosestMonster
            SpellCounter = 0
        End If
        'If isulandskill Then Casting = False
        SpellCounter = SpellCounter + 1
    End If
   
    
    If (Asc(Mid(inData, 15, 1)) = 26 Or Asc(Mid(inData, 15, 1)) = 27) And AvoidWarp And (Not MoveOnly) Then
            WarpPos.Y = Y
            WarpPos.X = X
            'Chat "[" & Get_PeopleName(Mid(InData, 3, 4)) & "] casting [Warp_Portal] at " + CStr(WarpPos.Y) + ":" + CStr(WarpPos.X) + " , Run Away..."
            WarpNumber = WarpNumber + 1
            If EvalNorm(WarpPos, curPos) < 3 Then
                'teleport
                Winsock_SendPacket IntToChr(&H85) & MakeMagePos(WarpPos), True
            End If
            
    End If
    
Decode_013E = ""
Exit Function
errie:
Decode_013E = "ERROR!!! [Decode_013E] " & Err.Description
Err.Clear
End Function

Private Function Decode_0141(inData As String) As String
On Error GoTo errie
    Dim ID As Integer
    Dim Val As Long
    Dim Val2 As Long
    ID = MakePort(Mid(inData, 3, 2))
    Val = MakePort(Mid(inData, 7, 2))
    Val2 = MakePort(Mid(inData, 11, 2))
    Select Case ID
        Case 13
            Players(number).STR = Val
            Players(number).Strp = Val2
        Case 14
            Players(number).AGI = Val
            Players(number).Agip = Val2
        Case 15
            Players(number).VIT = Val
            Players(number).Vitp = Val2
        Case 16
            Players(number).Intl = Val
            Players(number).Intp = Val2
        Case 17
            Players(number).DEX = Val
            Players(number).Dexp = Val2
        Case 18
            Players(number).LUK = Val
            Players(number).Lukp = Val2
    End Select
    UpdateStats
Decode_0141 = ""
Exit Function
errie:
Decode_0141 = "ERROR!!! [Decode_0141] - id[" & CStr(ID) & "] - " & CStr(MakePort(Mid(inData, 7, 2))) & " " & CStr(MakePort(Mid(inData, 11, 2))) & " " & Err.Description
Err.Clear
End Function

Private Function Decode_016F(inData As String) As String
On Error GoTo errie
If Mods.GuildText = True Then
    If MakeString(Mid(inData, 3, 60)) = "" Then
        Chat "[-]", MColor.guildannounce
    Else
        Chat "[" & MakeString(Mid(inData, 3, 60)) & "]", MColor.guildannounce
    End If
    If MakeString(Mid(inData, 63, 120)) = "" Then
        Chat "[-]", MColor.guildannounce
    Else
        Chat "[" & MakeString(Mid(inData, 63, 120)) & "]", MColor.guildannounce
    End If
End If
Decode_016F = ""
Exit Function
errie:
Decode_016F = "ERROR!!! [Decode_016F] " & Err.Description
Err.Clear
End Function

Private Function Decode_017F(inData As String) As String
On Error GoTo errie
    Chat "[Guild] " & Mid(inData, 5, MakePort(Mid(inData, 3, 2)) - 4), MColor.guildchat
    CheckEvent "OnGuildMessage", "name=" & Mid(inData, 5, InStr(5, inData, " : ") - 5) & Chr(0) & "message=" & Mid(inData, InStr(5, inData, " : "), Len(inData) - InStr(5, inData, " : "))
    'CheckChatResponse Mid(InData, 5, MakePort(Mid(InData, 3, 2)) - 4), 3, 0
Decode_017F = ""
Exit Function
errie:
Decode_017F = "ERROR!!! [Decode_017F] " & Err.Description
Err.Clear
End Function

Private Function Decode_0187() As String
On Error GoTo errie
    If ConnState < 4 Then ResettoReCon
    Decode_0187 = ""
    Exit Function
errie:
Decode_0187 = "ERROR!!! [Decode_0187] " & Err.Description
Err.Clear
ResettoReCon
End Function

Private Sub update_status()
    If frmStatus.Visible Then
        frmStatus.lstStatus.Clear
        Dim i As Integer
        For i = 0 To UBound(CurStatus)
            If CurStatus(i).Active Then
                If CurStatus(i).Name <> "" Then
                    frmStatus.lstStatus.AddItem CStr(i) & " : " & CurStatus(i).Name
                Else
                    frmStatus.lstStatus.AddItem CStr(i) & " : Unknown"
                End If
            End If
        Next
    End If
End Sub

Private Function Decode_01A2(inData As String) As String
On Error GoTo errie
    MyPet.Name = MakeString(Mid(inData, 3, 24))
    MyPet.Level = Asc(Mid(inData, 28, 1))
    MyPet.Status = Asc(Mid(inData, 30, 1))
    MyPet.Relation = MakePort(Mid(inData, 32, 2))
    'MDIfrmMain.mnuPet.Visible = True
    If MakePort(Mid(inData, 34, 2)) > 0 Then
        MyPet.Equipment = Return_ItemName(MakeHexName(Mid(inData, 34, 2)))
    Else
        MyPet.Equipment = "None"
    End If
Decode_01A2 = ""
Exit Function
errie:
Decode_01A2 = "ERROR!!! [Decode_01A2] " & Err.Description
Err.Clear
End Function

Private Function Decode_01B5(inData As String) As String
On Error GoTo errie
    Dim tstr$, timeleft As Double ',Hr&, Mi&, Se&, Credit&
    tstr = ""
    timeleft = MakePort(Mid(inData, 3, 4))
    If timeleft Mod 60 > 0 Then tstr = CStr(timeleft Mod 60) & " min(s)."
    timeleft = (timeleft - (timeleft Mod 60)) / 60
    If (timeleft Mod 24 > 0) Then tstr = CStr(timeleft Mod 24) & " hr(s)." & tstr
    timeleft = (timeleft - (timeleft Mod 24)) / 24
    If timeleft > 0 Then tstr = CStr(timeleft) & " Day(s)." & tstr
    If Len(tstr) > 0 Then Stat "Remaining time to play (day card) :" & tstr & vbCrLf
    tstr = ""
    timeleft = MakePort(Mid(inData, 7, 4))
    If timeleft Mod 60 > 0 Then tstr = CStr(timeleft Mod 60) & " min(s)."
    timeleft = (timeleft - (timeleft Mod 60)) / 60
    If (timeleft Mod 24 > 0) Then tstr = CStr(timeleft Mod 24) & " hr(s)." & tstr
    timeleft = (timeleft - (timeleft Mod 24)) / 24
    If timeleft > 0 Then tstr = CStr(timeleft) & " Day(s)." & tstr
    If Len(tstr) > 0 Then Stat "Remaining time to play (hours card) :" & tstr & vbCrLf
Decode_01B5 = ""
Exit Function
errie:
Decode_01B5 = "ERROR!!! [Decode_01B5] " & Err.Description
Err.Clear
End Function

Private Function Decode_01DC(inData As String) As String
On Error GoTo errie
        Stat "Got Session key..." & vbCrLf
        Dim md5Test As MD5
        Set md5Test = New MD5
        Dim login As String
        If MasterSelect.Encrypt = 1 Or MasterSelect.Encrypt = 3 Then
            login = Chr(&HDD) & Chr(&H1) & IntToChr(CLng(MasterSelect.code)) & IntToChr(0) & strUser & String(24 - Len(strUser), Chr(0))
            login = login & md5Test.DigestStrToChar(Mid(inData, 5, 16) & StrPass) & Chr(MasterSelect.Version)
        ElseIf MasterSelect.Encrypt = 4 Then
            login = Chr(&HDD) & Chr(&H1) & IntToChr(CLng(MasterSelect.code)) & IntToChr(0) & strUser & String(24 - Len(strUser), Chr(0))
            login = login & md5Test.DigestStrToChar(StrPass & Mid(inData, 5, 16)) & Chr(MasterSelect.Version)
        Else
            login = Chr(&HFA) & Chr(&H1) & IntToChr(CLng(MasterSelect.code)) & IntToChr(0) & strUser & String(24 - Len(strUser), Chr(0))
            login = login & md5Test.DigestStrToChar(Mid(inData, 5, 16) & StrPass) & Chr(MasterSelect.Version) & Chr(MasterSelect.enctype)
        End If
        IsLogin = False
        Stat "Verify your ID with MD5 hashing..." + vbCrLf
        If Not isUseHaunted Then Winsock_SendPacket login, True
Decode_01DC = ""
Exit Function
errie:
Decode_01DC = "ERROR!!! [Decode_01DC] " & Err.number & ":" & Err.Description
Err.Clear
End Function

Private Sub Decode_Else(inData As String)
        Stat "UNKNOWN PACKET !!!" + vbCrLf
        print_packet inData, "Unknown Packet"
        inData = ""
End Sub

' mod mc 0.1 build 1
Private Function Decode_0123(inData As String, Decode As String)
On Error GoTo errie
    Dim i, Packet As Integer
    Dim Itemname, Index As String
    Dim ItemData As String
    Dim ChopNumber As Long
    ChopNumber = MakePort(Mid(inData, 3, 2))
    IsCartOn = True
    'IsCartRecv = True
    If Decode = "0123" Then Packet = 10 Else Packet = 18
    For i = 5 To ChopNumber Step Packet
        Index = MakePort(Mid(inData, i, 2))
        If Index > UBound(Cart) Then ReDim Preserve Cart(Index)
        Itemname = Return_ItemName(MakeHexName(Mid(inData, i + 2, 2)))
        If Itemname = "" Then Itemname = "Unknow " & CStr(MakeHexName(Mid(inData, i + 2, 2)))
        Cart(Index).ID = Trim(STR(MakePort(Mid(inData, i + 2, 2))))
        Cart(Index).Name = Itemname
        Cart(Index).Amount = MakePort(Mid(inData, i + 6, 2))
        CheckCartInv CLng(Index)
        'ReDim Preserve Cart(UBound(Cart) + 1)
    Next
    UpdateCart
    CalcModAI "0123"
Exit Function
errie:
Decode_0123 = "ERROR!!! [Decode_0123] " & Err.Description
print_packet inData, "0123"
Err.Clear
End Function

Function Decode_0121(inData As String)
On Error GoTo errie
    'R 0121 <num>.w <num limit>.w <weight>.l <weight limit>l
    'kind of cart, weight and max weight.
    'IsCartOn = True
    'IsCartRecv = True
    CartNum = MakePort(Mid(inData, 3, 2))
    CartNumM = MakePort(Mid(inData, 5, 2))
    CartWeight = MakePort(Mid(inData, 7, 4))
    CartWeightM = MakePort(Mid(inData, 11, 4))
    Dim pcs&
    pcs = (CLng(CartWeight) * 100) / CartWeightM
    frmCart.lblCart.Caption = "Num: " & CartNum & "/" & CartNumM & " Weight: " & (CartWeight / 10) & "/" & (CartWeightM / 10)
    frmCart.lblCart.ToolTipText = "Weight " & pcs & "%"
Exit Function
errie:
Decode_0121 = "ERROR!!! [Decode_0121] " & Err.Description
print_packet inData, "0121"
Err.Clear
End Function

Private Function Decode_0124(inData As String)
On Error GoTo errie
    'R 0124 <index>.w <amount>.l <item ID>.w <identify flag>.B <attribute?>.B <refine>.B <card>.4w
    '               3                   5                   9                       11                          12                      13                  14,16,18,20
    'add item to cart.
    Dim Index As Integer
    Index = MakePort(Mid(inData, 3, 2))
    If Index > UBound(Cart) Then ReDim Preserve Cart(Index)
    IsCartOn = True
    'IsCartRecv = True
    Cart(Index).Name = MakeItemName(Mid(inData, 9, 2), Mid(inData, 14, 8), Mid(inData, 13, 1))
    Cart(Index).Pos = 0
    Cart(Index).Amount = Cart(Index).Amount + MakePort(Mid(inData, 5, 4))
    Cart(Index).ID = MakePort(Mid(inData, 9, 2))
    Cart(Index).Identified = CBool(Asc(Mid(inData, 11, 1)))
    'Cart(Index).Index = Index
    Cart(Index).Type = 0
    Cart(Index).Category = 0
    CheckCartInv CLng(Index)
    UpdateCart
Exit Function
errie:
Decode_0124 = "ERROR!!! [Decode_0124] " & Err.Description
print_packet inData, "0124"
Err.Clear
End Function

Private Function Decode_01C4(inData As String) As String
On Error GoTo errie
'R 00f4 <index>.w <amount>.l <type ID>.w <identify flag>.B <attribute?>.B <refine>.B <card>.4w
'               3                   5                   9                       11                          12                      13                  14
    Dim StorageIndex As Integer
    Dim NameID As String
    Dim Itemname As String
    StorageIndex = MakePort(Mid(inData, 3, 2))
    NameID = MakeHexName(Mid(inData, 9, 2))
    Itemname = MakeItemName(Mid(inData, 9, 2), Mid(inData, 15, 8), Mid(inData, 14, 1))
    Dim i As Integer
    If StorageIndex > UBound(Storage) Then
        ReDim Preserve Storage(StorageIndex)
        Storage(StorageIndex).Index = StorageIndex
        Storage(StorageIndex).NameID = NameID
        Storage(StorageIndex).Name = Itemname
        Storage(StorageIndex).Amount = MakePort(Mid(inData, 5, 4))
        Storage(StorageIndex).Identified = CBool(Asc(Mid(inData, 12, 1)))
        CheckCartStorage CLng(StorageIndex)
    Else
        'Storage(index).index = StorageIndex
        'Storage(index).Nameid = Nameid
        'Storage(index).name = Itemname
        Storage(StorageIndex).Amount = Storage(StorageIndex).Amount + MakePort(Mid(inData, 5, 4))
    End If
    upd_frmStorage
    Stat "Add [" & Itemname & "] " & CStr(MakePort(Mid(inData, 5, 4))) & " EA to Storage..." & vbCrLf
Decode_01C4 = ""
Exit Function
errie:
Decode_01C4 = "ERROR!!! [Decode_01C4] " & Err.Description
Err.Clear
End Function
Private Function Decode_01C5(inData As String)
On Error GoTo errie
    'C5 01 06 00 01 00 00 00 66 04 05 01 00 03 FF 00 03 05 41 90 01 00
    '      3     5           9     11 12    14 15    17    19    21
    'add item to cart.
    Dim Index As Integer
    Index = MakePort(Mid(inData, 3, 2))
    If Index > UBound(Cart) Then ReDim Preserve Cart(Index)
    'Dim Card() As Card_Profile, Itemname As String, i As Integer, name As String, TmpCard As String, Number As Long
    'Itemname = Return_ItemName(MakeHexName(Mid(InData, 9, 2)))
    Cart(Index).Name = MakeItemName(Mid(inData, 9, 2), Mid(inData, 15, 8), Mid(inData, 14, 1))
    Cart(Index).Pos = 0
    Cart(Index).Amount = Cart(Index).Amount + MakePort(Mid(inData, 5, 4))
    Cart(Index).ID = MakePort(Mid(inData, 9, 2))
    Cart(Index).Identified = CBool(Asc(Mid(inData, 12, 1)))
    'Cart(Index).Index = Index
    Cart(Index).Type = 0
    Cart(Index).Category = Asc(Mid(inData, 11, 1))
    CheckCartInv CLng(Index)
    UpdateCart
Exit Function
errie:
Decode_01C5 = "ERROR!!! [Decode_01C5] " & Err.Description
print_packet inData, "01C5"
Err.Clear
End Function

Private Function Decode_0125(inData As String)
On Error GoTo errie
'R 0125 <index>.w <amount>.l
    Dim Index As Integer
    'IsCartRecv = True
    Index = MakePort(Mid(inData, 3, 2))
    Cart(Index).Amount = Cart(Index).Amount - MakePort(Mid(inData, 5, 4))
    Stat "Cart item removed : [" & Cart(Index).Name & "] " & MakePort(Mid(inData, 5, 4)) & "EA" + vbCrLf, MColor.Shop
    UpdateCart
Exit Function
errie:
Decode_0125 = "ERROR!!! [Decode_0125] " & Err.Description
print_packet inData, "0125"
Err.Clear
End Function

Function Decode_0136(inData As String)
On Error GoTo errie
    'R 0136 <len>.w <ID>.l
    '{<value>.l <index>.w <amount>.w <type>.B <item ID>.w <identify flag>.B <attribute?>.B <refine>.B <card>.4w}.22B*
    ' 0                  4                  6                       8               9                       11                          12                      13                  14,16,18,20
    Dim pktLen As Long, shopid As Long, Index As Long, i As Long
    shopid = MakePort(Mid(inData, 5, 4))
    MyShopID = Mid(inData, 5, 4)
    pktLen = MakePort(Mid(inData, 3, 2))
    IsVending = True
    IsShopCreated = True
    For i = 0 To 12
        Shop(i).Name = ""
        Shop(i).Amount = 0
        Shop(i).ID = 0
        Shop(i).Index = i
        Shop(i).Price = 0
    Next
    For i = 9 To pktLen Step 22
        Index = MakePort(Mid(inData, i + 4, 2))
        Shop(Index).Name = MakeItemName(Mid(inData, i + 9, 2), Mid(inData, i + 14, 8), Mid(inData, i + 13, 1))
        Shop(Index).Amount = MakePort(Mid(inData, i + 6, 2))
        Shop(Index).Index = Index
        Shop(Index).ID = MakePort(Mid(inData, i + 9, 2))
        Shop(Index).Price = MakePort(Mid(inData, i, 4))
        Chat "Shop item list : [" & Shop(Index).Name & "] " & Shop(Index).Amount & " EA " & FormatNumber(Shop(Index).Price, 0, vbTrue, vbTrue, vbTrue) & "z", MColor.Shop
    Next
    UpdateShop
    CheckEvent "OnShopCreated", "nothingtocheck=True"
    Chat "Shop created. (AccID : " & shopid & ")", MColor.Shop
    Exit Function
errie:
Decode_0136 = "ERROR!!! [Decode_0136] " & Err.Description
print_packet inData, "0136"
Err.Clear
End Function

Function Decode_0137(inData As String)
On Error GoTo errie
    'R 0137 <index>.w <amount>.w
    Dim Index As Integer, i As Integer, Amount As Long, ccount As Long
    Index = MakePort(Mid(inData, 3, 2))
    Amount = MakePort(Mid(inData, 5, 2))
    Shop(Index).Amount = Shop(Index).Amount - Amount
    Chat "Sold : [" & Shop(Index).Name & "] " & Amount & " EA. Get " & FormatNumber(Amount * Shop(Index).Price, 0, vbTrue, vbTrue, vbTrue) & "z" & IIf(Shop(Index).Amount <= 0, " sold out", ""), MColor.shopsellitem
    UpdateShop
    ccount = 0
    For i = 0 To 13
        If Shop(i).Amount = 0 Then ccount = ccount + 1
    Next
    If ccount = 14 Then
        CheckEvent "OnShopSellItem", "itemname=" & Shop(Index).Name & Chr(0) & "count=" & CStr(Amount) & Chr(0) & "getZeny=" & (Amount * Shop(Index).Price) & Chr(0) & "isShopEmpty=True"
        Send_ShopClose
        Stat "Shop is empty, closed shop." + vbCrLf
        frmShop.Visible = False
        If Mods.dcshop Then
            End
        End If
        CalcShopAI
    Else
        CheckEvent "OnShopSellItem", "itemname=" & Shop(Index).Name & Chr(0) & "count=" & CStr(Amount) & Chr(0) & "getZeny=" & (Amount * Shop(Index).Price) & Chr(0) & "isShopEmpty=False"
    End If
    UpdateCart
    Exit Function
errie:
Decode_0137 = "ERROR!!! [Decode_0137] " & Err.Description
print_packet inData, "0137"
Err.Clear
End Function
Function Decode_012C(inData As String)
    'If Asc(Mid(InData, 3, 1)) = 0 Then
        Chat "Can't add item to cart : overweight - auto-storage routing", vbRed
        MIsGoStore = True
    'End If
End Function
Function Decode_012D(inData As String)
    IsVendingWait = False
    MaxShopAmount = MakePort(Mid(inData, 3, 2))
End Function

'trade section
Private Function Decode_00E5(inData As String) As String
On Error GoTo errie
    NResetTrade
    MTPartner = MakeString(Mid(inData, 3, 24))
    Chat "System : [Trade] Request : '" & MTPartner & "' - " & IIf(Mods.Enabled = True And Mods.OC = True, "Accepting...", "Automatic Cancel..."), MColor.trade
    MTradeStep = 0
    MODTradeDelay = GetTickCount + MODDC.TAccept
    MODTradeStep = 1
    Decode_00E5 = ""
Exit Function
errie:
Decode_00E5 = "ERROR!!! [Decode_00E5] " & Err.Description
Err.Clear
End Function
Function Decode_01F4(inData As String) As String
On Error GoTo errie
    NResetTrade
    MTPartner = MakeString(Mid(inData, 3, 24))
    If Mods.Enabled = True And Mods.OC = True Then
        Chat "System : [Trade] Request : '" & MTPartner & " [Lv:" & CStr(MakePort(Mid(inData, 31, 2))) & _
        " /AID: " & CStr(MakePort(Mid(inData, 27, 4))) & "]' - Accepting...", MColor.trade
    Else
        Chat "System : [Trade] Request : '" & MTPartner & " [Lv:" & CStr(MakePort(Mid(inData, 31, 2))) & _
        " /AID: " & CStr(MakePort(Mid(inData, 27, 4))) & "]' - Automatic Cancel...", MColor.trade
    End If
    MTradeStep = 0
    MODTradeDelay = GetTickCount + MODDC.TAccept
    MODTradeStep = 1
    Decode_01F4 = ""
Exit Function
errie:
Decode_01F4 = "ERROR!!! [Decode_01F4] " & Err.Description
Err.Clear
End Function
Function Decode_01F5(inData As String) As String
On Error GoTo errie
    Select Case Asc(Mid(inData, 3, 1))
        Case 0
            Chat "System : [Trade] Too far", MColor.trade
            MODTradeStep = 0
        Case 3
            Chat "System : [Trade] Allowed for trading", MColor.trade
            MODTradeStep = 2
            MODTradeDelay = GetTickCount + MODDC.TItem
        Case 4
            Chat "System : [Trade] Trade canceled. ", MColor.trade
            MODTradeStep = 0
    End Select
    Exit Function
errie:
Decode_01F5 = "ERROR!!! [Decode_01F5] " & Err.Description
print_packet inData, "01F5"
Err.Clear
End Function
Function Decode_00E7(inData As String) As String
On Error GoTo errie
    Select Case Asc(Mid(inData, 3, 1))
        Case 0
            Chat "System : [Trade] Too far", MColor.trade
            MODTradeStep = 0
        Case 3
            Chat "System : [Trade] Allowed for trading", MColor.trade
            MODTradeStep = 2
            MODTradeDelay = GetTickCount + MODDC.TItem
        Case 4
            Chat "System : [Trade] Trade canceled. ", MColor.trade
            MODTradeStep = 0
    End Select
    Exit Function
errie:
Decode_00E7 = "ERROR!!! [Decode_00E7] " & Err.Description
print_packet inData, "00E7"
Err.Clear
End Function
Function Decode_00E9(inData As String) As String
On Error GoTo errie
'R 00e9 <amount>.l <type ID>.w <identify flag>.B <attribute?>.B <refine>.B <card>.4w
'       1         3                 7                       9                               10                      11              12
    If MakePort(Mid(inData, 7, 2)) > 0 Then
        MTrade(UBound(MTrade)).Amount = MakePort(Mid(inData, 3, 4))
        MTrade(UBound(MTrade)).Identified = CBool(Asc(Mid(inData, 9, 1)))
        MTrade(UBound(MTrade)).ItemID = MakePort(Mid(inData, 7, 2))
        MTrade(UBound(MTrade)).Itemname = MakeItemName(Mid(inData, 7, 2), Mid(inData, 12, 8), Mid(inData, 11, 1))
        If (Not IsTradeAccept(MTrade(UBound(MTrade)).Itemname) Or MTrade(UBound(MTrade)).Identified = False) And Not Mods.OCnocalcmoney Then
            MODTradeStep = 3
            MODTradeDelay = GetTickCount + RandomNumber(3000, 1500)
            Chat "System : [Trade] '" & MTPartner & "' add item : " & MTrade(UBound(MTrade)).Itemname & " " & MTrade(UBound(MTrade)).Amount & " EA [rejecting]", MColor.trade
        Else
            Chat "System : [Trade] '" & MTPartner & "' add item : " & MTrade(UBound(MTrade)).Itemname & " " & MTrade(UBound(MTrade)).Amount & " EA", MColor.trade
            MODTradeStep = 4
            MODTradeDelay = GetTickCount + MODDC.TNItem
        End If
        ReDim Preserve MTrade(UBound(MTrade) + 1)
    Else
        Chat "System : [Trade] Your partner added " & MakePort(Mid(inData, 3, 4)) & " zeny", MColor.trade
    End If
    Exit Function
errie:
Decode_00E9 = "ERROR!!! [Decode_00E9] " & Err.Description
print_packet inData, "00E9"
Err.Clear
End Function
'Private Sub TmrIT_Timer()
'    TmrIT.Enabled = False
'    Chat "System : [Trade] Time-out on your trade partner, calculating for overcharge.", MColor.trade
'    CalcTrade
'End Sub
Function Decode_00EA(inData As String) As String
'R 00ea <index>.w <fail>.B
    If Asc(Mid(inData, 3, 1)) <> 0 Then Chat "System : [Trade] Failed to add an item", MColor.trade
End Function
Function Decode_00EC(inData As String) As String
On Error GoTo errie
'R 00ec <final>.B
If Asc(Mid(inData, 3, 1)) = 0 Then
    Chat "System : [Trade] You've completed trade." & IIf(MTStatus.Partner = True, " delaying . . . ", ""), MColor.trade
    If MTStatus.Partner = True Then
        MODTradeStep = 5
        MODTradeDelay = GetTickCount + MODDC.TCalc + RandomNumber(1000, 0)
    End If
    MTStatus.You = True
Else
    Chat "System : [Trade] " & MTPartner & " completed trade." & IIf(MTStatus.You = True, " delaying . . . ", ""), MColor.trade
    If MTStatus.You = True Then
        MODTradeStep = 5
        MODTradeDelay = GetTickCount + MODDC.TCalc + RandomNumber(1000, 0)
    Else
        CalcTrade
    End If
    MTStatus.Partner = True
End If
Exit Function
errie:
Decode_00EC = "ERROR!!! [Decode_00EC] " & Err.Description
print_packet inData, "00EC"
Err.Clear
End Function
Private Function Decode_00EE() As String
On Error GoTo errie
    MODTradeStep = 0
    Chat "System : [Trade] Trade is Cancelled...", MColor.trade
Decode_00EE = ""
Exit Function
errie:
Decode_00EE = "ERROR!!! [Decode_00EE] " & Err.Description
Err.Clear
End Function
Private Function Decode_00F0() As String
On Error GoTo errie
    TmrIT.Enabled = False
    Chat "System : [Trade] Completed.", MColor.trade
    If Players(number).Zeny < Mods.minzeny Then
        Mods.Vending = True
        Stat "Your money's below " & Mods.minzeny & ".", MColor.Fail
        If Mods.OCcreateshop Then
            Stat " Creating shop" + vbCrLf, vbBlue
            frmMain.destroy_chatroom
            If Sitting Then frmMain.Send_Stand
            Mods.Vending = True
            CreateShop
            Exit Function
        End If
        If Mods.OCdisconnect Then
            Stat " Exitting program" + vbCrLf, vbBlue
            End
            Exit Function
        End If
        Stat " Closing chatroom" + vbCrLf, vbBlue
        frmMain.destroy_chatroom
    End If
Decode_00F0 = ""
Exit Function
errie:
Decode_00F0 = "ERROR!!! [Decode_00F0] " & Err.Description
Err.Clear
End Function

Function Decode_00D7(inData As String)
On Error GoTo errie
    'R 00d7 <len>.w <owner ID>.l <chat ID>.l <limit>.w <users>.w <pub>.B <title>.?B
    '                  3             5                      9                   13              15              17              18 to (len - 22)
    Dim i&, redimMC As Boolean, chatid&
    For i = 0 To UBound(MChat)
        If (MChat(i).Owner = Mid(inData, 5, 4)) Or MChat(i).Visible = False Then
            redimMC = True
            chatid = i
            Exit For
        End If
    Next
    If Not redimMC Then
        ReDim Preserve MChat(UBound(MChat) + 1)
        chatid = UBound(MChat)
    End If
    MChat(chatid).Visible = True
    MChat(chatid).Title = Mid(inData, 18, Len(inData) - 17)
    MChat(chatid).CLimit = MakePort(Mid(inData, 13, 2))
    MChat(chatid).CUsers = MakePort(Mid(inData, 15, 2))
    MChat(chatid).ID = Mid(inData, 9, 4)
    MChat(chatid).IsPub = Asc(Mid(inData, 17, 1))
    MChat(chatid).Owner = Mid(inData, 5, 4)
    UpdateChatShop
    Exit Function
errie:
Decode_00D7 = "ERROR!!! [Decode_00D7] " & Err.Description
Err.Clear
End Function
Function Decode_00D8(inData As String)
'R 00d8 <chat ID>.l
On Error GoTo errie
    Dim i&, IsBack As Boolean
    IsBack = False
    For i = 0 To UBound(MChat)
        If Mid(inData, 3, 4) = MChat(i).ID Then IsBack = True
        If IsBack Then
            MChat(i) = MChat(IIf(i + 1 > UBound(MChat), i, i + 1))
        End If
    Next
    If IsBack Then ReDim Preserve MChat(IIf(UBound(MChat) > 1, UBound(MChat) - 1, 0))
    UpdateChatShop
    Exit Function
errie:
Decode_00D8 = "ERROR!!! [Decode_00D8] " & Err.Description
Err.Clear
End Function
Function Decode_0131(inData As String)
'R 0131 <ID>.l <message>.80B
On Error GoTo errie
    Dim i&, redimMS As Boolean, shopid&
    For i = 0 To UBound(MShop)
        If Mid(inData, 3, 4) = MShop(i).ID Or MShop(i).Visible = False Then
            redimMS = True
            shopid = i
            Exit For
        End If
    Next
    If Not redimMS Then
        ReDim Preserve MShop(UBound(MShop) + 1)
        shopid = UBound(MShop)
    End If
    MShop(shopid).Visible = True
    MShop(shopid).ID = Mid(inData, 3, 4)
    MShop(shopid).Name = MakeString(Mid(inData, 7, 80))
    UpdateChatShop
    Exit Function
errie:
Decode_0131 = "ERROR!!! [Decode_0131] " & Err.Description
Err.Clear
End Function
Function Decode_0132(inData As String)
'R 0132 <ID>.l
On Error GoTo errie
    If Mid(inData, 3, 4) = MyShopID Then
        IsVending = False
        IsShopCreated = False
        If Mods.STChat Then Chat "System : [shop] Your shop is closed.", MColor.Shop
        Exit Function
    End If
    Dim i&, IsBack As Boolean
    IsBack = False
    For i = 0 To UBound(MShop)
        If Mid(inData, 3, 4) = MShop(i).ID Then IsBack = True
        If IsBack Then
            MShop(i) = MShop(IIf(i + 1 > UBound(MShop), i, i + 1))
        End If
    Next
    If IsBack Then ReDim Preserve MShop(IIf(UBound(MShop) > 1, UBound(MShop) - 1, 0))
    UpdateChatShop
    Exit Function
errie:
Decode_0132 = "ERROR!!! [Decode_0132] " & Err.Description
Err.Clear
End Function
Function Decode_00DD(inData As String)
    If MakeString(Mid(inData, 5, 24)) = Players(number).Name Then
        Chat "System : [chat room] You leave the chat room", MColor.Shop
        IsChatOC = False
    End If
    Exit Function
End Function
Function Decode_00DF(inData As String)
    'R 00df <len>.w <owner ID>.l <chat ID>.l <limit>.w <users>.w <pub>.B <title>.?B
    '               3              5                    9                       13          15                  17             18
    If Mid(inData, 5, 4) = AccountID Then
        Chat "System : [chat room] Room changed : '" & Mid(inData, 18, Len(inData) - 17) & "'"
    End If
End Function

'skill fail
Function Decode_0110(inData As String) As String
On Error GoTo errie
    'R 0110 <skill ID>.w <basic type>.w ?.w <fail>.B <type>.B
    '               3                       5                           7       9           10
'    fail to use skill when fail=00?
'    type 00:basic type 01:lack of SP, 02:lack of HP, 03:no memo, 04:in delay
'    05:lack of money, 06:weapon does not satisfy, 07:no red gem, 08:no blue gem, 09:unknown
'    basic type 00:trade 01:emotion 02:sit down, 03:chat, 04:party
'    05:shout? 06:PK, 07:manner point
    Dim DSkID&, DSkType As Byte, DSkBType&, DSkFail As Byte, MsgS$
    DSkID = MakePort(Mid(inData, 3, 2))
    DSkType = Asc(Mid(inData, 10, 1))
    DSkBType = MakePort(Mid(inData, 5, 2))
    DSkFail = Asc(Mid(inData, 9, 1))
    
    If DSkFail = 0 And Mods.STSKFail Then Chat "System : [Skill] - Skill fail detected [Skill ID:" & DSkID & " / Type:" & DSkType & " / BType:" & DSkBType & "]", vbRed
    If DSkFail = 0 And SkillCounter > 0 And DetectFail Then SkillCounter = SkillCounter - 1
    
    'If DSkFail = 0 Then
    '    If DSkType = 0 Then
    '        Select Case DSkBType
    '            Case 0: MsgS = " [fail to use]"
    '            Case 1: MsgS = " [require more SP]"
    '            Case 2: MsgS = " [require more HP]"
    '            Case 3: MsgS = " [no memo point]"
    '            Case 4: MsgS = " [in skill delay]"
    '            Case 5: MsgS = " [require more ZENY]"
    '            Case 6: MsgS = " [incorrect weapon type]"
    '            Case 7: MsgS = " [require Red Gemstone]"
    '            Case 8: MsgS = " [require Blue Gemstone]"
    '            Case Else: MsgS = " [unknown:" & DSkBType & "]"
    '        End Select
    '    Else
    '        Select Case DSkBType
    '            Case 0: MsgS = " [" & DSkBType & "][fail to use]"
    '            Case 1: MsgS = " [" & DSkBType & "][basic skill : emotion]"
    '            Case 2: MsgS = " [" & DSkBType & "][basic skill : sit down]"
    '            Case 3: MsgS = " [" & DSkBType & "][basic skill : chat room creation]"
    '            Case 4: MsgS = " [" & DSkBType & "][basic skill : party creation]"
    '            Case 5: MsgS = " [" & DSkBType & "][basic skill : storage opening]"
    '            Case 6: MsgS = " [" & DSkBType & "][basic skill : pk]"
    '            Case 7: MsgS = " [" & DSkBType & "][basic skill : manner point]"
    '            Case Else: MsgS = " [basic skill : unknown:" & DSkBType & "]"
    '        End Select
    '    End If
    '    Dim i&
    '    For i = 0 To UBound(SkillChar)
    '        If SkillChar(i).ID = DSkID Then
    '            'MakeDamage = True
    '            If Mods.STStatus Then Stat "Skill failed : " & SkillChar(i).Name & MsgS + vbCrLf
    '            Exit Function
    '        End If
    '    Next
    '    'MakeDamage = True
    '    If Mods.STStatus Then Stat "Skill failed : Unknown [" & DSkID & "]" & MsgS + vbCrLf
    'End If
    Exit Function
errie:
Decode_0110 = "ERROR!!! [Decode_0110] " & Err.Description
Err.Clear
End Function

Private Function Decode_01B6(inData As String)
On Error GoTo errie
    With frmGuild
        .labAverage = MakePort(Mid(inData, 19, 4))
        .labExp = MakePort(Mid(inData, 23, 4))
        .labNextLv = MakePort(Mid(inData, 27, 4))
        .labLV = MakePort(Mid(inData, 7, 4))
        .labMaster = MakeString(Mid(inData, 71, 24))
        .labGuildName = MakeString(Mid(inData, 47, 24))
        .labMember = MakePort(Mid(inData, 11, 4)) & "/" & MakePort(Mid(inData, 15, 4))
    End With
Exit Function
errie:
Decode_01B6 = "ERROR!!! [Decode_01B6] " & Err.Description
Err.Clear
End Function
Private Function Decode_011C(inData As String)
On Error GoTo errie
    'R 011c <skill ID>.w <map1>.16B <map2>.16B <map3>.16B <map4>.16B
    '                 3                     5
    If MakePort(Mid(inData, 3, 2)) = 26 Then
        'S 011b <skill ID>.w <map name>.16B
        Stat "Got teleport skill effect, Teleport." & vbCrLf
        Winsock_SendPacket IntToChr(&H11B) & Mid(inData, 3, 18), True
    End If
Exit Function
errie:
Decode_011C = "ERROR!!! [Decode_011C] " & Err.Description
Err.Clear
End Function
Private Function Decode_016C(inData As String)
On Error GoTo errie
        Players(number).Guild = Mid(inData, 20, 24)
        frmGuild.LabGuild.Caption = Players(number).Guild & " - Guild"
Exit Function
errie:
Decode_016C = "ERROR!!! [Decode_016C] " & Err.Description
Err.Clear
End Function

Function Decode_01D7(inData As String)
'On Error GoTo errie
    'If Mid(InData, 3, 4) = AccountID Then
        'Chat "01D7 update > " & ChrtoHex(Mid(InData, 7, 5))
        'If MakePort(Mid(InData, 7, 4)) = 2 Then IsCartOn = True
    'End If
    'Exit Function
'errie:
'Decode_01D7 = "ERROR!!! [Decode_01D7] " & Err.Description
'Err.Clear
End Function

Function Decode_0139(inData As String)
On Error GoTo errie
    Dim nPos As Coord
    nPos.X = MakePort(Mid(inData, 9, 2))
    nPos.Y = MakePort(Mid(inData, 7, 2))
    move_to NearPos(nPos, curPos, PRange)
    'SendAttack
    Exit Function
errie:
Decode_0139 = "ERROR!!! [Decode_0139] " & Err.Description
Err.Clear
End Function

Sub UpdateChatShop()
On Error GoTo errie
'print_errror "sub UpdateChatShop"
    Dim i&, j&, isCC&
    frmChatRoom.lstChatroom.Clear
    For i = 0 To UBound(MChat)
        If MChat(i).Visible Then
            isCC = -1
            For j = 0 To UBound(People)
                If People(j).ID = MChat(i).Owner Then isCC = j
            Next
            If isCC > -1 Then
                frmChatRoom.lstChatroom.AddItem "[chat] -'" & MChat(i).Title & "'" & IIf(MChat(i).IsPub = 0, "*", "") & "[" & MChat(i).CUsers & "/" & MChat(i).CLimit & "] [" & People(isCC).Name & " (" & EvalNorm(curPos, People(isCC).Pos) & " blks)]"
            Else
                frmChatRoom.lstChatroom.AddItem "[chat] -'" & MChat(i).Title & "'" & IIf(MChat(i).IsPub = 0, "*", "") & "[" & MChat(i).CUsers & "/" & MChat(i).CLimit & "] [Unknown]"
            End If
        End If
    Next
    For i = 0 To UBound(MShop)
        If MShop(i).Visible Then
            isCC = -1
            For j = 0 To UBound(People)
                If People(j).ID = MShop(i).ID Then isCC = j
            Next
            If isCC > -1 Then
                frmChatRoom.lstChatroom.AddItem "[shop] -'" & MShop(i).Name & "' [" & People(isCC).Name & " (" & EvalNorm(curPos, People(isCC).Pos) & " blks)] [" & MakeHex(MShop(i).ID) & "]"
            Else
                frmChatRoom.lstChatroom.AddItem "[shop] -'" & MShop(i).Name & "' [Unknown] [" & MakeHex(MShop(i).ID) & "]"
            End If
        End If
    Next
Exit Sub
errie:
Stat "Error in UpdateChatShop : " & Err.Description + vbCrLf
Err.Clear
End Sub

Public Sub Send_ShopClose()
        Dim i&, tcid&
    For i = 0 To UBound(Shop)
        If Shop(i).Amount > 0 Then
            tcid = Find_CartID(Shop(i).Name)
            If tcid < 0 Then
                ReDim Preserve Cart(UBound(Cart) + 1)
                tcid = UBound(Cart)
            End If
            Cart(tcid).Amount = Cart(tcid).Amount + Shop(i).Amount
        End If
    Next
    UpdateCart
    Winsock_SendPacket IntToChr(&H12E), True
    IsVending = False
    IsShopCreated = False
    frmShop.Visible = False
End Sub
Public Sub Send_Guildinfo(i As Integer)
        Winsock_SendPacket Chr(&H4F) + Chr(1) + Chr(i) + Chr(0) + Chr(0) + Chr(0), True
End Sub
Public Sub Send_GuildRequest()
    Winsock_SendPacket Chr(&H4D) + Chr(1), True
End Sub

Private Function Decode_0196(inData As String) As String
On Error GoTo errie
    Dim X As Long
    If (Mid(inData, 5, 4) = AccountID) Then
        Dim tmp$, stid&
        stid = MakePort(Mid(inData, 3, 2))
        If stid > UBound(CurStatus) Then ReDim Preserve CurStatus(stid)
        If Mods.STStatus Then Chat "Your status changed [" & IIf(Len(CurStatus(stid).Name) = 0, "Unknown:" & stid, CurStatus(stid).Name) & " (" & IIf(Asc(Mid(inData, 9, 1)) = 1, "on", "off") & ")]: "
        tmp = CurStatus(stid).Active
        CurStatus(stid).Active = CBool(Asc(Mid(inData, 9, 1)))
        CheckEvent "OnYourStatusChange", "StatusName=" & IIf(Len(CurStatus(stid).Name) = 0, "Unknown:" & stid, CurStatus(stid).Name) & Chr(0) & "StatusNum=" & stid & Chr(0) & "isActive=" & CStr(CurStatus(stid).Active)
        If tmp <> CurStatus(stid).Active Then update_status
        If Asc(Mid(inData, 9, 1)) = 1 Then AddPicStatus stid Else DelPicStatus stid
        If stid = 89 And Asc(Mid(inData, 9, 1)) = 1 Then
            If DelayuseCC = 0 Then DelayuseCC = 2
            If DelayuseFC = 0 Then DelayuseFC = 5
        End If
    End If
   Decode_0196 = ""
Exit Function
errie:
Decode_0196 = "ERROR!!! [Decode_0196] " & Err.Description
Err.Clear
End Function

Private Sub ParseData()
Dim intA As Integer
Dim tstr As String
Dim ChopNumber As Long
'Dim TmpData As String
Dim DebugTstr As String
Dim X As Integer
Dim Y As Integer
Dim addval As Integer
Dim found As Boolean
Dim found2 As Boolean
Dim NoJump As Boolean
Dim tcoords As Coord
Dim tlng As Long
Dim teststr As String
Dim distance As Integer
Dim tmpname As String
Dim errlen As Boolean
'If Not IsLogin Then
'    On Error GoTo runtime5
'Else
'    On Error Resume Next
'End If
restart:
IsLogin = False
If Len(RecvData) < 2 Then Exit Sub

ChopNumber = 0
DebugTstr = ""
Dim debugdata As Long, curLen As Long
ChopNumber = GetPacketLen(ChrtoHex(Mid(RecvData, 1, 2)))
If ChopNumber < 0 Then
    If Len(RecvData) > 3 Then
        ChopNumber = MakePort(Mid(RecvData, 3, 2))
    Else
        NoJump = True
        GoTo doCheckEnd
    End If
End If
If ChopNumber = 0 Then ChopNumber = Len(RecvData): errlen = True
If Len(RecvData) < ChopNumber Then
    NoJump = True
    GoTo doCheckEnd
End If
tmpdata = Left(RecvData, ChopNumber)
RecvData = IIf(Len(RecvData) = ChopNumber, "", Right(RecvData, Len(RecvData) - ChopNumber))

If MDIfrmMain.mnuPKTLOG.CheckED Then
        Open App.Path & "\packet.txt" For Append As #11
            Print #11, "Receive : "
            Print #11, ConvPacketData(tmpdata)
            Print #11, ""
        Close #11
End If

debugdata = MakePort(Mid(tmpdata, 1, 2))
Select Case debugdata
Case &H69 'response after initial connect which holds new servers
        DebugTstr = Decode_0069(tmpdata)

Case &H6A
        DebugTstr = Decode_006A(tmpdata)

Case &H6B   'char data
        DebugTstr = Decode_006B(tmpdata)

Case &H6C
        DebugTstr = Decode_006C()

Case &H71
        DebugTstr = Decode_0071(tmpdata)

Case &H73 'cur pos
        DebugTstr = Decode_0073(tmpdata)

Case &H77

Case &H78
        DebugTstr = Decode_0078(tmpdata)

Case &H79
        DebugTstr = Decode_0079(tmpdata)

Case &H7B
        DebugTstr = Decode_007B(tmpdata)

Case &H7C
        DebugTstr = Decode_007C(tmpdata)

Case &H7F

Case &H80
        DebugTstr = Decode_0080(tmpdata)

Case &H81 'Disconnect From Server
        DebugTstr = Decode_0081(tmpdata)

Case &H87 'Current Position
        DebugTstr = Decode_0087(tmpdata)

Case &H88 'Current Position
        DebugTstr = Decode_0088(tmpdata)

Case &H89

Case &H8A 'Action Information
        DebugTstr = Decode_008A(tmpdata)

Case &H8D 'Chat
        DebugTstr = Decode_008D(tmpdata)

Case &H8E 'Chat
        DebugTstr = Decode_008E(tmpdata)

Case &H8F

 Case &H91
        DebugTstr = Decode_0091(tmpdata)

Case &H92 'Map Server Information
        DebugTstr = Decode_0092(tmpdata)

Case &H93

Case &H95 'Name Received
        DebugTstr = Decode_0095(tmpdata)

Case &H96

Case &H97 'Whisper Message
        DebugTstr = Decode_0097(tmpdata)

Case &H98 'Chat Response Information
        DebugTstr = Decode_0098(tmpdata)

Case &H9A 'Chat
        DebugTstr = Decode_009A(tmpdata)

Case &H9C

Case &H9D

Case &H9E 'Item Dropped
        DebugTstr = Decode_009E(tmpdata)

Case &HA0 'Got Item
        DebugTstr = Decode_00A0(tmpdata)

Case &HA1 'Item Disappeared
        DebugTstr = Decode_00A1(tmpdata)

Case &HA3 'Your Inventory Information
        DebugTstr = Decode_00A3(tmpdata)

Case &HA4 'Equipment Inventory
        DebugTstr = Decode_00A4(tmpdata)

Case &HA5 'Storage Information (Kafra)
        DebugTstr = Decode_00A5(tmpdata)

Case &HA6 'Storage Information (Kafra)
        DebugTstr = Decode_00A6(tmpdata)

Case &HA8 'Item Amount Update
        DebugTstr = Decode_00A8(tmpdata)

Case &HAA 'Equip Action Information
        DebugTstr = Decode_00AA(tmpdata)

Case &HAC
        DebugTstr = Decode_00AC(tmpdata)

Case &HAF 'Item left
        DebugTstr = Decode_00AF(tmpdata)

Case &HB0 'Status Information (Str, AGi,. Etc)
        DebugTstr = Decode_00B0(tmpdata)

Case &HB1 'Status Information (EXP, Zeny)
        DebugTstr = Decode_00B1(tmpdata)

Case &HB3

Case &HB4 'NPC Message
        DebugTstr = Decode_00B4(tmpdata)

Case &HB5 'NPC Continue Talking '
        DebugTstr = Decode_00B5(tmpdata)

Case &HB6 'NPC Close Talking
        DebugTstr = Decode_00B6(tmpdata)

Case &HB7 ' NPC: Choice List '
        DebugTstr = Decode_00B7(tmpdata)

Case &HBC
        
Case &HBD
        DebugTstr = Decode_00BD(tmpdata)
        
Case &HBE
        DebugTstr = Decode_00BE(tmpdata)

Case &HC0
        DebugTstr = Decode_00C0(tmpdata)

Case &HC1

Case &HC2
        DebugTstr = Decode_00C2(tmpdata)

Case &HC3

Case &HC4 'Send Buy/Sell
        DebugTstr = Decode_00C4(tmpdata)

Case &HC6 'Store Open
        DebugTstr = Decode_00C6(tmpdata)

Case &HC7
        DebugTstr = Decode_00C7(tmpdata)

Case &HC8 ' Figure Out '

Case &HCA

Case &HCB

Case &HCD
        Open App.Path & "\warning.txt" For Append As #8
        Chat Date & "@" & Time & ": [GM] Kicked you from server, Closed program..."
        Close 8
        End
Case &HD1

Case &HD2 'Whisper Response Information
        DebugTstr = Decode_00D2(tmpdata)

Case &HD6
        Decode_00D6
Case &HD7
        DebugTstr = Decode_00D7(tmpdata)

Case &HD8
        DebugTstr = Decode_00D8(tmpdata)

Case &HDA

Case &HDD
        DebugTstr = Decode_00DD(tmpdata)

Case &HDF
        DebugTstr = Decode_00DF(tmpdata)

Case &HE4

'mod-mc trading info
Case &HE5
        DebugTstr = Decode_00E5(tmpdata)

Case &HE7 ' You Cancel Deal '
        DebugTstr = Decode_00E7(tmpdata)

Case &HE9 ' Player add Money '
        DebugTstr = Decode_00E9(tmpdata)

Case &HEA
        DebugTstr = Decode_00EA(tmpdata)

Case &HEC 'Finalize Deal
        DebugTstr = Decode_00EC(tmpdata)

Case &HEE
        DebugTstr = Decode_00EE()

Case &HF0
        DebugTstr = Decode_00F0()

Case &HF2
        DebugTstr = Decode_00F2(tmpdata)

Case &HF4 'Storage Amount Update
        DebugTstr = Decode_00F4(tmpdata)

Case &HF6 'Item in Storgae Left
        DebugTstr = Decode_00F6(tmpdata)

Case &HF8

Case &HFB
        DebugTstr = Decode_00FB(tmpdata)

Case &HFD
        DebugTstr = Decode_00FD(tmpdata)

Case &HFE
        DebugTstr = Decode_00FE(tmpdata)

Case &H100

Case &H101
        DebugTstr = Decode_0101(tmpdata)

Case &H102

Case &H104
        DebugTstr = Decode_0104(tmpdata)

Case &H105
        DebugTstr = Decode_0105(tmpdata)

Case &H106
        DebugTstr = Decode_0106(tmpdata)

Case &H107
        DebugTstr = Decode_0107(tmpdata)

Case &H109
        DebugTstr = Decode_0109(tmpdata)

Case &H10A
        DebugTstr = Decode_010A(tmpdata)

Case &H10B
        DebugTstr = Decode_010B(tmpdata)

Case &H10C

Case &H10E
        DebugTstr = Decode_010E(tmpdata)

Case &H10F
        DebugTstr = Decode_010F(tmpdata)

Case &H110
        DebugTstr = Decode_0110(tmpdata)

Case &H111

Case &H114
        DebugTstr = Decode_0114(tmpdata)

Case &H115 'Skill Use Information
        DebugTstr = Decode_0115(tmpdata)

Case &H117
        DebugTstr = Decode_0117(tmpdata)

Case &H119 'Status Change
        DebugTstr = Decode_0119(tmpdata)

Case &H11A
        DebugTstr = Decode_011A(tmpdata)

Case &H11C
        DebugTstr = Decode_011C(tmpdata)

Case &H11E

Case &H11F

Case &H120

Case &H121
        DebugTstr = Decode_0121(tmpdata)

Case &H122
        DebugTstr = Decode_0122(tmpdata)

Case &H123
        DebugTstr = Decode_0123(tmpdata, "0123")

Case &H124
        DebugTstr = Decode_0124(tmpdata)

Case &H125
        DebugTstr = Decode_0125(tmpdata)

Case &H12C
        DebugTstr = Decode_012C(tmpdata)

Case &H12D
        DebugTstr = Decode_012D(tmpdata)

Case &H131
        DebugTstr = Decode_0131(tmpdata)

Case &H132
        DebugTstr = Decode_0132(tmpdata)

Case &H133

Case &H136
        DebugTstr = Decode_0136(tmpdata)

Case &H137
        DebugTstr = Decode_0137(tmpdata)

Case &H139
        DebugTstr = Decode_0139(tmpdata)

Case &H13A
        DebugTstr = Decode_013A(tmpdata)

Case &H13B

Case &H13C
        DebugTstr = Decode_013C(tmpdata)

Case &H13D 'HP Update
        DebugTstr = Decode_013D(tmpdata)

Case &H13E 'Skill Use Detail
        DebugTstr = Decode_013E(tmpdata)

Case &H141
        DebugTstr = Decode_0141(tmpdata)

Case &H144

Case &H145

Case &H147

Case &H148

Case &H14B

Case &H14C
        DebugTstr = Decode_014C(tmpdata)

Case &H14E

Case &H150
        DebugTstr = Decode_0150(tmpdata)

Case &H152 'guild emblem

Case &H154 'Guild
        DebugTstr = Decode_0154(tmpdata)

Case &H156

Case &H15A

Case &H15C

Case &H162 'guild skill list

Case &H166 'Guild
        DebugTstr = Decode_0166(tmpdata)

Case &H16A

Case &H16C
        DebugTstr = Decode_016C(tmpdata)

Case &H16D

Case &H16F
        DebugTstr = Decode_016F(tmpdata)

Case &H174

Case &H179

Case &H17F 'Guild Message
        DebugTstr = Decode_017F(tmpdata)
Case &H180

Case &H182

Case &H183

Case &H187
        DebugTstr = Decode_0187()
Case &H18A

Case &H18B

Case &H18F

Case &H192

Case &H191

Case &H192

Case &H194

Case &H195 'Players Name (Party, Guild) Information
        DebugTstr = Decode_0195(tmpdata)
Case &H196 'Status Changed
        DebugTstr = Decode_0196(tmpdata)
Case &H198

Case &H199

Case &H19A

Case &H19B

Case &H19E

Case &H1A2
        DebugTstr = Decode_01A2(tmpdata)
Case &H1A3
       DebugTstr = Decode_01A3(tmpdata)
Case &H1A4
        DebugTstr = Decode_01A4(tmpdata)
Case &H1AA

Case &H1AB

Case &H1AC

Case &H1B0

Case &H1B3

Case &H1B5
       DebugTstr = Decode_01B5(tmpdata)

Case &H1B6
        DebugTstr = Decode_01B6(tmpdata)

Case &H1B9

Case &H1C4 '???
       DebugTstr = Decode_01C4(tmpdata)

Case &H1C5
       DebugTstr = Decode_01C5(tmpdata)

Case &H1C8 'Update Item
       DebugTstr = Decode_01C8(tmpdata)

Case &H1C9 'New

Case &H1CF 'New

Case &H1D0 'New
       DebugTstr = Decode_01D0(tmpdata)
        
Case &H1D2 'New

Case &H1D6 'New

Case &H1D7
       DebugTstr = Decode_01D7(tmpdata)
       
Case &H1D8
       DebugTstr = Decode_0078(tmpdata)
       
Case &H1D9
       DebugTstr = Decode_0079(tmpdata)
       
Case &H1DA
       DebugTstr = Decode_007B(tmpdata)
       
Case &H1DC
        DebugTstr = Decode_01DC(tmpdata)

Case &H1DE
       DebugTstr = Decode_01DE(tmpdata)
      
Case &H1E1 'New-5.0

Case &H1E6 'New-5.0

Case &H1EB 'New-5.0

Case &H1EE 'New-5.0
        DebugTstr = Decode_01EE(tmpdata)

Case &H1EF 'Cart Information New-5.0
        DebugTstr = Decode_0123(tmpdata, "01EF")

Case &H1F0 'Storage Information (Kafra) new 5.0 (replaced 00A5)
        DebugTstr = Decode_01F0(tmpdata)

Case &H1F1
        DebugTstr = Decode_01F1(tmpdata)

Case &H1F2

Case &H1F4 'trade information new kunroon patch (replaced 00E5)
        DebugTstr = Decode_01F4(tmpdata)

Case &H1F5 'trade information new kunroon patch (replaced 00E7)
        DebugTstr = Decode_01F5(tmpdata)

Case &H1FA

Case &H201 'friend list

Case &H206 '???

Case Else
    If errlen Then
        Chat "Packet decoding error: UNKNOWN LENGTH!!" & vbCrLf, vbRed
        print_packet lastPacket, "Packet Pre-Processing before unknown"
        Decode_Else tmpdata
    Else
        print_packet lastPacket, "Packet Pre-Processing before unknown"
        Decode_Else tmpdata
    End If
End Select
If DebugTstr <> "" Then
    Stat DebugTstr & vbCrLf
    Open App.Path & "\log\errorlog.txt" For Append As #1
    Print #1, " == " & Error & "(" & Date & ")@" & Time & " == "
    Print #1, DebugTstr
    Close #1
End If
ChopErrorCounter = 0
'If Not NoJump Then Stat "Debug " & DebugTstr & vbCrLf
doCheckEnd:
lastPacket = tmpdata
If (Len(RecvData) >= 2) And (Not NoJump) Then
    'DoEvents
    GoTo restart
End If
Exit Sub
runtime5:
    
    'If Len(RecvData) >= chopnumber Then
    '    RecvData = Right(RecvData, Len(RecvData) - chopnumber)
    'Else
    Dim tos As String
    tos = "(" & Err.number & ") " & Err.Description
    Err.Clear
    'If IsLogin Then GoTo rundd:
    Stat "[" & DebugTstr & "] parse data Error..." & tos & vbCrLf
    print_packet RecvData, "Error Packet - " & tos
    RecvData = ""
    'End If
    'If (Len(Recvdata) > 0) Then GoTo restart
    If FrmField.Visible Then Load_Field MapName
    FrmField.PicMap.Refresh
    ClearAll
    RecvData = ""
End Sub

Private Sub tmrEvents_Timer()
On Error GoTo errie
    ProcessAction
    Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "tmrEvents_Timer", Err.number, Err.Description
    Err.Clear
End Sub

Public Sub Send_AddParty()
    With frmPeople.lstPeople
        If .List(.ListIndex) <> "" Then
            Winsock_SendPacket IntToChr(&HFC) & People(.ListIndex).ID, True
        End If
    End With
End Sub

Public Sub Send_KickParty()
    With frmParty.lstParty
        If .List(.ListIndex) <> "" Then
            'Winsock_SendPacket IntToChr(&H103) & Party(.ListIndex).Id & Party(.ListIndex).Name, True
            'Chat Val(.ListIndex)
        End If
    End With
End Sub

Public Sub Send_LeaveParty()
    With frmParty.lstParty
        If .List(.ListIndex) <> "" Then
            frmParty.lstParty.Clear
            ReDim Preserve Party(0)
            Winsock_SendPacket IntToChr(&H1), True
        End If
    End With
End Sub

Public Sub Send_AddGuild()
    With frmPeople.lstPeople
        If .List(.ListIndex) <> "" Then
            Winsock_SendPacket IntToChr(&H16B) & People(.ListIndex).ID, True
        End If
    End With
End Sub

'Public Sub Send_KickGuild()
    'With frmParty.lstParty
    '    If .List(.ListIndex) <> "" Then
            'Winsock_SendPacket IntToChr(&H103) & Party(.ListIndex).Id & Party(.ListIndex).Name, True
            'Chat Val(.ListIndex)
    '    End If
    'End With
'End Sub

'Public Sub Send_LeaveGuild()
    'With frmParty.lstParty
    '    If .List(.ListIndex) <> "" Then
    '        Winsock_SendPacket IntToChr(&H1), True
    '    End If
    'End With
'End Sub

'R 00fb <len>.w <party name>.24B {<ID>.l <nick>.24B <map name>.16B <leader>.B <offline>.B}.46B*
Private Function Decode_00FB(inData As String)
On Error GoTo errie
    ReDim Party(0)
    Dim i As Integer
    Dim ChopNumber As Long
    Players(number).Party = MakeString(Mid(inData, 5, 24))
    ChopNumber = MakePort(Mid(inData, 3, 2))
    frmParty.LabParty.Caption = Players(number).Party & " - Party"
    For i = 29 To ChopNumber Step 46
        '{<ID>.l <nick>.24B <map name>.16B <leader>.B <offline>.B}.46B*
        '0            4                    28                                44                 45
        If Party(UBound(Party)).Name <> "" Then ReDim Preserve Party(UBound(Party) + 1)
        Party(UBound(Party)).ID = Mid(inData, i, 4)
        Party(UBound(Party)).Name = MakeString(Mid(inData, i + 4, 24))
        Party(UBound(Party)).Map = MakeString(Mid(inData, i + 28, 16))
        Party(UBound(Party)).Admin = Not CBool(Asc(Mid(inData, i + 44, 1)))
        Party(UBound(Party)).Online = CBool(Asc(Mid(inData, i + 45, 1)))
    Next
    UpdateParty
Exit Function
errie:
Decode_00FB = "ERROR!!! [Decode_00FB] " & Err.Description
Err.Clear
End Function
Private Function Decode_00FD(inData As String)
On Error GoTo errie
    Dim Name As String
    Name = MakeString(Mid(inData, 3, 24))
    Select Case Asc(Mid(inData, 27, 1))
        Case 0
            Chat "System : [party] - [" & Name & "] already have party"
        Case 1
            Chat "System : [party] - [" & Name & "] denied request"
        Case 2
            Chat "System : [party] - [" & Name & "] Accept"
    End Select
    Decode_00FD = ""
Exit Function
errie:
Decode_00FD = "ERROR!!! [Decode_00FD] " & Err.Description
Err.Clear
End Function
Private Function Decode_00FE(inData As String) As String
On Error GoTo errie
    Chat "System : [party] Incoming request to join party [" & MakeString(Mid(inData, 7, 24)) & "] from [" & Get_PeopleName(Mid(inData, 3, 4)) & "] , Automatic cancel...", MColor.trade
Decode_00FE = ""
Exit Function
errie:
Decode_00FE = "ERROR!!! [Decode_00FE] " & Err.Description
Err.Clear
End Function
Private Function Decode_0104(inData As String)
On Error GoTo errie
    Players(number).Party = MakeString(Mid(inData, 16, 24))
    Dim found As Boolean
    'frmParty.LabParty.Caption = Players(number).Party & " - Party"
    Dim i As Integer
    For i = 0 To UBound(Party)
        If Party(i).ID = Mid(inData, 3, 4) Then
            found = True
            Clear_Dot Party(i).Pos
            Party(i).Name = MakeString(Mid(inData, 40, 24))
            Party(i).Pos.Y = MakePort(Mid(inData, 11, 2))
            Party(i).Pos.X = MakePort(Mid(inData, 13, 2))
            Party(i).Online = Not CBool(Asc(Mid(inData, 15, 1)))
            Party(i).Map = MakeString(Mid(inData, 64, 16))
            'If FollowMode.Active Then FollowCheck 2
            If MapName = Party(i).Map Then Plot_Dot Party(i).Pos, &H404040
        End If
    Next
    If Not found Then
        If Party(UBound(Party)).Name <> "" Then ReDim Preserve Party(UBound(Party) + 1)
        With Party(UBound(Party))
            .Name = MakeString(Mid(inData, 40, 24))
            .Pos.Y = MakePort(Mid(inData, 11, 2))
            .Pos.X = MakePort(Mid(inData, 13, 2))
            .Online = Not CBool(Asc(Mid(inData, 15, 1)))
            .Map = MakeString(Mid(inData, 64, 16))
        End With
    End If
    UpdateParty
Exit Function
errie:
Decode_0104 = "ERROR!!! [Decode_0104] " & Err.Description
Err.Clear
End Function
Private Function Decode_0105(inData As String)
On Error GoTo errie
    Dim X As Integer, isRedim As Boolean
    If Mid(inData, 3, 4) = AccountID Then
        Chat "You left the party"
        Players(number).Party = ""
        ReDim Party(0)
    Else
        Chat "System : [Party] - " & MakeString(Mid(inData, 7, 24)) & " left the party"
        For X = 0 To UBound(Party)
            If Party(X).ID = Mid(inData, 3, 4) Then isRedim = True
            If isRedim Then Party(X) = Party(X + 1)
        Next
        ReDim Preserve Party(UBound(Party) - 1)
    End If
    UpdateParty
Exit Function
errie:
Decode_0105 = "ERROR!!! [Decode_0105] " & Err.Description
Err.Clear
End Function
Private Function Decode_0106(inData As String)
On Error GoTo errie
    Dim i As Integer
    Dim ParHP&, HealPos&, dist&
    For i = 0 To UBound(Party)
        If Party(i).ID = Mid(inData, 3, 4) Then
            Party(i).HpMin = MakePort(Mid(inData, 7, 2))
            Party(i).HPmax = MakePort(Mid(inData, 9, 2))
            If Party(i).HPmax > 0 Then ParHP = (CLng(Party(i).HpMin) * 100) \ CLng(Party(i).HPmax)
            'If FollowMode.Active Then FollowCheck 2
            If ParHP < 70 And FollowMode.AutoBuff Then
                HealPos = Find_SkillId("AL_HEAL")
                dist = EvalNorm(curPos, Party(i).Pos)
                If dist < 10 And HealPos > -1 Then
                    If Mods.STParty Then Chat "System : [Party] - Autobuff 'HEAL' on [" & Party(i).Name & "]"
                    Send_Use_Skill SkillChar(HealPos).ID, SkillChar(HealPos).MaxLV, Party(i).ID
                End If
            End If
            CheckEvent "OnPartyHPChange", "AID=" & MakePort(Mid(inData, 3, 4)) & Chr(0) & "name=" & Party(i).Name & Chr(0) & "posX=" & Party(i).Pos.Y & Chr(0) & "posY=" & Party(i).Pos.X & Chr(0) & "curHP=" & Party(i).HpMin & Chr(0) & "maxHP=" & Party(i).HPmax & Chr(0) & "percentHP=" & ParHP & Chr(0) & "distance=" & EvalNorm(curPos, Party(i).Pos)
        End If
    Next
    UpdateParty
Exit Function
errie:
Decode_0106 = "ERROR!!! [Decode_0106] " & Err.Description
Err.Clear
End Function
Private Function Decode_0107(inData As String)
On Error GoTo errie
    Dim i As Long
    For i = 0 To UBound(Party)
        If Party(i).ID = Mid(inData, 3, 4) Then
            Clear_Dot Party(i).Pos
            Party(i).Pos.Y = CInt(MakePort(Mid(inData, 7, 2)))
            Party(i).Pos.X = CInt(MakePort(Mid(inData, 9, 2)))
            Plot_Dot Party(i).Pos, &HFF00FF
            Dim partyHP&
            'If FollowMode.Active Then FollowCheck 2
            If Party(i).HPmax > 0 Then partyHP = (CLng(Party(i).HpMin) * 100) \ CLng(Party(i).HPmax)
            CheckEvent "OnPartyMove", "AID=" & MakePort(Mid(inData, 3, 4)) & Chr(0) & "name=" & Party(i).Name & Chr(0) & "posX=" & Party(i).Pos.Y & Chr(0) & "posY=" & Party(i).Pos.X & Chr(0) & "curHP=" & Party(i).HpMin & Chr(0) & "maxHP=" & Party(i).HPmax & Chr(0) & "percentHP=" & partyHP & Chr(0) & "distance=" & EvalNorm(curPos, Party(i).Pos)
        End If
    Next
    UpdateParty
Exit Function
errie:
Decode_0107 = "ERROR!!! [Decode_0107] " & Err.Description
Err.Clear
End Function
Private Function Decode_0109(inData As String) As String
On Error GoTo errie
    Chat "[Party] " & Mid(inData, 9, MakePort(Mid(inData, 3, 2)) - 8), MColor.Party
    'FollowMsgCheck Mid(inData, 5, 4), Mid(inData, 9, InStr(9, inData, " : ") - 9), Mid(inData, InStr(9, inData, " : ") + 3, Len(inData) - InStr(9, inData, " : ") - 3)
    CheckEvent "OnPartyMessage", "name=" & Mid(inData, 9, InStr(9, inData, " : ") - 9) & Chr(0) & "message=" & Mid(inData, InStr(9, inData, " : ") + 2, Len(inData) - InStr(9, inData, " : ") - 2) & Chr(0) & "AID=" & MakePort(Mid(inData, 5, 4))
    'CheckChatResponse Mid(InData, 9, MakePort(Mid(InData, 3, 2)) - 8), 2, 0
Decode_0109 = ""
Exit Function
errie:
Decode_0109 = "ERROR!!! [Decode_0109] " & Err.Description
Err.Clear
End Function

'mvp
Private Function Decode_010A(inData As String)
On Error GoTo errie
    Chat "You got MVP Item [" & Return_ItemName(MakeHexName(Mid(inData, 3, 2))) & "]"
    Decode_010A = ""
Exit Function
errie:
Decode_010A = "ERROR!!! [Decode_010A] " & Err.Description
Err.Clear
End Function
Private Function Decode_010B(inData As String)
On Error GoTo errie
    Chat "You're MVP!!! Special exp : " & CStr(CLng(Mid(inData, 3, 4)))
    Decode_010B = ""
Exit Function
errie:
Decode_010B = "ERROR!!! [Decode_010B] " & Err.Description
Err.Clear
End Function

Public Sub UpdateParty()
On Error GoTo errie
''print_errror "sub UpdateParty"
    frmParty.lstParty.Clear
    If Len(Players(number).Party) = 0 Then Exit Sub
    frmParty.LabParty.Caption = Players(number).Party & " - Party"
    Dim X As Integer, tstr$
    For X = 0 To UBound(Party)
        With Party(X)
            If Len(.Name) > 0 Then
                tstr = .Name & " [" & IIf(.Online, .Map & " (" & .Pos.Y & "/" & .Pos.X & ")", "Offline") & "], (" & Party(X).HpMin & "/" & Party(X).HPmax & ")"
                If .Admin Then tstr = tstr & ", [M]"
                frmParty.lstParty.AddItem tstr
            End If
        End With
    Next
Exit Sub
errie:
'If Err.number > 0 Then print_funcerr "UpdateParty", Err.number, Err.Description
Err.Clear
End Sub
