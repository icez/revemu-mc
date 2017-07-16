VERSION 5.00
Begin VB.Form FrmHPSPOption 
   BorderStyle     =   0  'None
   Caption         =   "HP/SP Option"
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2820
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox ImgSitNomons 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "FrmHPSPOption.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   25
      Top             =   1330
      Width           =   225
   End
   Begin VB.TextBox txtHealLv 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   255
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   24
      Text            =   "10"
      Top             =   2040
      Width           =   255
   End
   Begin VB.PictureBox imgHeal 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "FrmHPSPOption.frx":011B
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   22
      Top             =   2040
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.TextBox txtHealHP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   255
      Left            =   2800
      MaxLength       =   2
      TabIndex        =   21
      Text            =   "48"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox txtItem2HP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   255
      Left            =   3700
      MaxLength       =   2
      TabIndex        =   20
      Text            =   "48"
      Top             =   1830
      Width           =   255
   End
   Begin VB.PictureBox imgItem2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "FrmHPSPOption.frx":0236
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   18
      Top             =   1830
      Width           =   225
   End
   Begin VB.TextBox txtItem2Name 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   255
      Left            =   1200
      TabIndex        =   17
      Text            =   "Red_Potion"
      Top             =   1830
      Width           =   1215
   End
   Begin VB.TextBox txtItem1HP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   255
      Left            =   3700
      MaxLength       =   2
      TabIndex        =   16
      Text            =   "48"
      Top             =   1590
      Width           =   255
   End
   Begin VB.TextBox txtAutositHP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   255
      Left            =   2200
      MaxLength       =   3
      TabIndex        =   14
      Text            =   "48"
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox imgAutoSitHP 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "FrmHPSPOption.frx":0351
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   13
      Top             =   360
      Width           =   225
   End
   Begin VB.TextBox txtSitUntilHP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   255
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   11
      Text            =   "48"
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox imgSitUntilHP 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "FrmHPSPOption.frx":046C
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   10
      Top             =   600
      Width           =   225
   End
   Begin VB.TextBox txtAutoSitSP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   255
      Left            =   2200
      MaxLength       =   3
      TabIndex        =   8
      Text            =   "48"
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox imgAutoSitSP 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "FrmHPSPOption.frx":0587
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   7
      Top             =   840
      Width           =   225
   End
   Begin VB.PictureBox imgUseItem1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "FrmHPSPOption.frx":06A2
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   3
      Top             =   1590
      Width           =   225
   End
   Begin VB.TextBox txtItem1Name 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Text            =   "Red_Herb"
      Top             =   1590
      Width           =   1215
   End
   Begin VB.TextBox txtSitUntilSP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   255
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "48"
      Top             =   1100
      Width           =   255
   End
   Begin VB.PictureBox imgSitUntilSP 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "FrmHPSPOption.frx":07BD
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   1100
      Width           =   225
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Sit when no monster"
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
      TabIndex        =   26
      Top             =   1330
      Width           =   1890
   End
   Begin VB.Label LabHeal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto heal lv.       when HP below       % "
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
      TabIndex        =   23
      Top             =   2040
      Width           =   2880
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto use [                            ] when HP below       % "
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
      Top             =   1830
      Width           =   3765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Sit when HP below       %"
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
      TabIndex        =   15
      Top             =   360
      Width           =   2235
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sit until HP reach       %"
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
      TabIndex        =   12
      Top             =   600
      Width           =   1680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Sit when SP below       %"
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
      Top             =   840
      Width           =   2235
   End
   Begin VB.Image Image5 
      Height          =   300
      Left            =   3480
      Picture         =   "FrmHPSPOption.frx":08D8
      ToolTipText     =   "Change edited value."
      Top             =   2460
      Width           =   630
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   0
      Picture         =   "FrmHPSPOption.frx":0B8A
      Top             =   2400
      Width           =   4200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "HP/SP Options"
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
      TabIndex        =   6
      Top             =   15
      Width           =   1215
   End
   Begin VB.Label LabStopPick 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto use [                            ] when HP below       % "
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
      TabIndex        =   5
      Top             =   1590
      Width           =   3765
   End
   Begin VB.Label labBackTown 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sit until SP reach       %"
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
      TabIndex        =   4
      Top             =   1095
      Width           =   1680
   End
   Begin VB.Image Image4 
      Height          =   135
      Left            =   70
      Picture         =   "FrmHPSPOption.frx":0CCE
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image6 
      Height          =   135
      Left            =   3960
      Picture         =   "FrmHPSPOption.frx":0E03
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   2160
      Left            =   0
      Picture         =   "FrmHPSPOption.frx":0F38
      Stretch         =   -1  'True
      Top             =   240
      Width           =   4200
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "FrmHPSPOption.frx":1309
      Top             =   0
      Width           =   4200
   End
End
Attribute VB_Name = "FrmHPSPOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
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
    Unload FrmHPSPOption
End Sub

Public Sub update_ImgAutositHP()
    If IsAutorest Then
        imgAutoSitHP.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgAutoSitHP.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
    txtAutositHP.text = HPSit * 100
End Sub

Private Sub imgAutoSitHP_Click()
    IsAutorest = Not IsAutorest
    update_ImgAutositHP
    MDIfrmMain.Save_Option
End Sub

Public Sub update_imgSitUntilHP()
    If IsHPWait Then
        imgSitUntilHP.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgSitUntilHP.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
    txtSitUntilHP.text = HPWait * 100
End Sub

Public Sub update_imgHeal()
    If Autoheal Then
        imgHeal.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgHeal.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
    txtHealLv.text = HealLV
    txtHealHP.text = HPHeal * 100
End Sub

Private Sub imgHeal_Click()
    Autoheal = Not Autoheal
    update_imgHeal
    MDIfrmMain.Save_Option
End Sub

Public Sub update_imgSitNomons()
    If IsNomonsSit Then
        ImgSitNomons.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        ImgSitNomons.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
End Sub


Private Sub ImgSitNomons_Click()
    IsNomonsSit = Not IsNomonsSit
    update_imgSitNomons
    MDIfrmMain.Save_Option
End Sub

Private Sub imgSitUntilHP_Click()
    IsHPWait = Not IsHPWait
    update_imgSitUntilHP
    MDIfrmMain.Save_Option
End Sub

Public Sub update_imgAutoSitSP()
    If IsSPSit Then
        imgAutoSitSP.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgAutoSitSP.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
    txtAutoSitSP.text = SPSit * 100
End Sub


Private Sub imgAutoSitSP_Click()
    IsSPSit = Not IsSPSit
    update_imgAutoSitSP
    MDIfrmMain.Save_Option
End Sub

Public Sub update_imgSitUntilSP()
    If IsSPWait Then
        imgSitUntilSP.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgSitUntilSP.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
    txtSitUntilSP.text = SPWait * 100
End Sub

Private Sub imgSitUntilSP_Click()
    IsSPWait = Not IsSPWait
    update_imgSitUntilSP
    MDIfrmMain.Save_Option
End Sub

Public Sub update_imgUseItem1()
    If IsAutoRedz Then
        imgUseItem1.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgUseItem1.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
    txtItem1Name.text = healitem1
    txtItem1HP.text = HPRed * 100
End Sub

Private Sub imgUseItem1_Click()
    IsAutoRedz = Not IsAutoRedz
    update_imgUseItem1
    MDIfrmMain.Save_Option
End Sub

Public Sub update_imgUseItem2()
    If IsAutoOrange Then
        imgItem2.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgItem2.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
    txtItem2Name.text = healitem2
    txtItem2HP.text = HPOrange * 100
End Sub

Private Sub imgItem2_Click()
    IsAutoOrange = Not IsAutoOrange
    update_imgUseItem2
    MDIfrmMain.Save_Option
End Sub

Private Sub txtAutositHP_Change()
    HPSit = Val(txtAutositHP.text) / 100
End Sub

Private Sub txtAutoSitSP_Change()
    SPSit = Val(txtAutoSitSP.text) / 100
End Sub

Private Sub txtHealHP_Change()
    HPHeal = Val(txtHealHP.text) / 100
End Sub

Private Sub txtHealLv_Change()
    HealLV = Val(txtHealLv.text)
End Sub

Private Sub txtItem1HP_Change()
    HPRed = Val(txtItem1HP.text) / 100
End Sub

Private Sub txtItem1Name_Change()
    healitem1 = txtItem1Name.text
End Sub

Private Sub txtItem2HP_Change()
    HPOrange = Val(txtItem2HP.text) / 100
End Sub

Private Sub txtItem2Name_Change()
    healitem2 = txtItem2Name.text
End Sub

Private Sub txtSitUntilHP_Change()
    HPWait = txtSitUntilHP.text / 100
End Sub

Private Sub txtSitUntilSP_Change()
    SPWait = txtSitUntilSP.text / 100
End Sub
