VERSION 5.00
Begin VB.Form frmAttackOption 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2460
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtMinDistance 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1054
         SubFormatType   =   1
      EndProperty
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
      Left            =   2175
      MaxLength       =   2
      TabIndex        =   21
      Text            =   "48"
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox imgMinDistance 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "frmAttackOption.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   20
      Top             =   1320
      Width           =   225
   End
   Begin VB.PictureBox imgKillMob 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "frmAttackOption.frx":011B
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   18
      Top             =   1560
      Width           =   225
   End
   Begin VB.TextBox txtLV 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1054
         SubFormatType   =   1
      EndProperty
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
      Left            =   3450
      MaxLength       =   2
      TabIndex        =   17
      Text            =   "10"
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox imgUseWeapon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "frmAttackOption.frx":0236
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   15
      Top             =   1800
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.TextBox txtSkill 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1054
         SubFormatType   =   0
      EndProperty
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
      Left            =   1490
      TabIndex        =   14
      Text            =   "SM_BASH"
      Top             =   600
      Width           =   1620
   End
   Begin VB.TextBox txtMonster 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1054
         SubFormatType   =   0
      EndProperty
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
      Left            =   480
      TabIndex        =   6
      Text            =   "Goblin"
      Top             =   840
      Width           =   1630
   End
   Begin VB.PictureBox imgAutoRangeAttack 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "frmAttackOption.frx":0351
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   5
      Top             =   1080
      Width           =   225
   End
   Begin VB.PictureBox imgAutoMobSkill 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "frmAttackOption.frx":046C
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   4
      Top             =   600
      Width           =   225
   End
   Begin VB.TextBox txtMonsNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1054
         SubFormatType   =   1
      EndProperty
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
      Left            =   3225
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "2"
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox imgAutoSkill 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1440
      Picture         =   "frmAttackOption.frx":0587
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   360
      Width           =   225
   End
   Begin VB.PictureBox imgAutoAttack 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "frmAttackOption.frx":06A2
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   360
      Width           =   225
   End
   Begin VB.TextBox txtBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1054
         SubFormatType   =   1
      EndProperty
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
      Left            =   2170
      MaxLength       =   2
      TabIndex        =   0
      Text            =   "48"
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Attack if distance above       block(s)"
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
      Left            =   405
      TabIndex        =   22
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Attemt to kill mob trained"
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
      Left            =   405
      TabIndex        =   19
      Top             =   1560
      Width           =   1740
   End
   Begin VB.Label LabUseWeapon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Use Weapon (Mage Cls/Acolyte Cls/S.Novice)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   210
      Left            =   405
      TabIndex        =   16
      Top             =   1800
      Width           =   3315
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "["
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
      Left            =   405
      TabIndex        =   13
      Top             =   840
      Width           =   45
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "] attack you > "
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
      Left            =   2175
      TabIndex        =   12
      Top             =   840
      Width           =   1020
   End
   Begin VB.Image Image6 
      Height          =   135
      Left            =   3960
      Picture         =   "frmAttackOption.frx":07BD
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image4 
      Height          =   135
      Left            =   70
      Picture         =   "frmAttackOption.frx":08F2
      Top             =   60
      Width           =   135
   End
   Begin VB.Label LabStopPick 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Attack if distance below       block(s)"
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
      Left            =   405
      TabIndex        =   11
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Attack Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   15
      Width           =   1215
   End
   Begin VB.Image Image5 
      Height          =   300
      Left            =   3480
      Picture         =   "frmAttackOption.frx":0A27
      Top             =   2100
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto use skill [                                      ] lv.      when "
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
      Left            =   405
      TabIndex        =   9
      Top             =   600
      Width           =   3750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Skill Use"
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
      Left            =   1725
      TabIndex        =   8
      Top             =   360
      Width           =   990
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Attack"
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
      Left            =   405
      TabIndex        =   7
      Top             =   360
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   0
      Picture         =   "frmAttackOption.frx":0CD9
      Top             =   2040
      Width           =   4200
   End
   Begin VB.Image Image2 
      Height          =   2160
      Left            =   0
      Picture         =   "frmAttackOption.frx":0E1D
      Stretch         =   -1  'True
      Top             =   240
      Width           =   4200
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "frmAttackOption.frx":11EE
      Top             =   0
      Width           =   4200
   End
End
Attribute VB_Name = "frmAttackOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub


Private Sub Image5_Click()
    MDIfrmMain.Save_Option
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image5.Picture = LoadPicture(App.Path & "\interface\bt_change_c.gif")
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image5.Picture = LoadPicture(App.Path & "\interface\bt_change.gif")
End Sub

Private Sub Image6_Click()
    Unload frmAttackOption
End Sub

Public Sub update_imgAutoAttack()
    If IsAutoKill Then
        imgAutoAttack.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgAutoAttack.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
End Sub

Private Sub imgAutoAttack_Click()
    IsAutoKill = Not IsAutoKill
    update_imgAutoAttack
    MDIfrmMain.Save_Option
End Sub

Public Sub update_imgAutoSkill()
    If IsSkillUse Then
        imgAutoSkill.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgAutoSkill.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
End Sub

Private Sub imgAutoSkill_Click()
    IsSkillUse = Not IsSkillUse
    update_imgAutoSkill
    MDIfrmMain.Save_Option
End Sub

Public Sub update_imgAutoMobSkill()
    If UseSkillMobs Then
        imgAutoMobSkill.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgAutoMobSkill.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
    txtSkill.text = MobSkill.rawname
    txtMonster.text = MobSkill.monsname
    txtMonsNumber.text = MobSkill.number
    txtLV = MobSkill.Lv
End Sub

Private Sub imgAutoMobSkill_Click()
    UseSkillMobs = Not UseSkillMobs
    update_imgAutoMobSkill
    MDIfrmMain.Save_Option
End Sub



Public Sub update_imgAutoRangeAttack()
    If IsUseRange Then
        imgAutoRangeAttack.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgAutoRangeAttack.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
    txtBlock.text = RangeSet
End Sub

Private Sub imgAutoRangeAttack_Click()
    IsUseRange = Not IsUseRange
    update_imgAutoRangeAttack
    MDIfrmMain.Save_Option
End Sub



Public Sub update_imgKS()
    If killsteal Then
        imgKS.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgKS.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
End Sub

Public Sub update_imgKillMob()
    If isKillmob Then
        imgKillMob.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgKillMob.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
End Sub

Private Sub imgKillMob_Click()
    isKillmob = Not isKillmob
    update_imgKillMob
    MDIfrmMain.Save_Option
End Sub

'Private Sub imgKS_Click()
'    killsteal = Not killsteal
'    update_imgKS
'    MDIfrmMain.Save_Option
'End Sub

Public Sub update_imgUseWeapon()
    If UseWeapon Then
        imgUseWeapon.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgUseWeapon.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
End Sub

Public Sub update_imgMindistance()
    If useMinDistance Then
        imgMinDistance.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgMinDistance.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
    TxtMinDistance.text = MinDistance
End Sub

Private Sub imgMinDistance_Click()
    useMinDistance = Not useMinDistance
    update_imgMindistance
    MDIfrmMain.Save_Option
End Sub

Private Sub imgUseWeapon_Click()
    UseWeapon = Not UseWeapon
    update_imgUseWeapon
    MDIfrmMain.Save_Option
End Sub

Private Sub txtBlock_Change()
    RangeSet = Val(txtBlock.text)
End Sub

Private Sub txtLV_Change()
    MobSkill.Lv = Val(txtLV.text)
End Sub

Private Sub TxtMinDistance_Change()
    MinDistance = Val(TxtMinDistance.text)
End Sub

Private Sub txtMonsNumber_Change()
    MobSkill.number = txtMonsNumber.text
End Sub

Private Sub txtSkill_Change()
    MobSkill.rawname = txtSkill.text
End Sub
