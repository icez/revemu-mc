VERSION 5.00
Begin VB.Form frmGuild 
   BorderStyle     =   0  'None
   ClientHeight    =   4890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1230
      Left            =   0
      Picture         =   "frmGuild.frx":0000
      ScaleHeight     =   1230
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   255
      Width           =   300
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   300
      TabIndex        =   2
      Top             =   240
      Width           =   4215
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   4215
      End
      Begin VB.Label lab1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name  :"
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
         Left            =   45
         TabIndex        =   22
         Top             =   10
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lv  :"
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
         Left            =   45
         TabIndex        =   21
         Top             =   250
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Members  :"
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
         Left            =   45
         TabIndex        =   20
         Top             =   490
         Width           =   945
      End
      Begin VB.Label labGuildName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   720
         TabIndex        =   19
         Top             =   15
         Width           =   45
      End
      Begin VB.Label labLV 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   480
         TabIndex        =   18
         Top             =   255
         Width           =   45
      End
      Begin VB.Label labMember 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   1080
         TabIndex        =   17
         Top             =   495
         Width           =   45
      End
      Begin VB.Label labAverage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   1245
         TabIndex        =   16
         Top             =   735
         Width           =   45
      End
      Begin VB.Label labExp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   600
         TabIndex        =   15
         Top             =   975
         Width           =   45
      End
      Begin VB.Label labMaster 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   900
         TabIndex        =   14
         Top             =   1215
         Width           =   45
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Average LV  :"
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
         Left            =   45
         TabIndex        =   13
         Top             =   730
         Width           =   1095
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp  :"
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
         Left            =   45
         TabIndex        =   12
         Top             =   975
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Master  :"
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
         Left            =   45
         TabIndex        =   11
         Top             =   1215
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Castle  :"
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
         Left            =   45
         TabIndex        =   10
         Top             =   1455
         Width           =   660
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Next LV  :"
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
         Left            =   45
         TabIndex        =   9
         Top             =   1695
         Width           =   765
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Castle  :"
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
         Left            =   45
         TabIndex        =   8
         Top             =   2160
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Castle  :"
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
         Left            =   45
         TabIndex        =   7
         Top             =   1920
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label labCastle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   840
         TabIndex        =   6
         Top             =   1455
         Width           =   45
      End
      Begin VB.Label labNextLv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   960
         TabIndex        =   5
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "sss"
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
         Left            =   840
         TabIndex        =   4
         Top             =   1920
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "sss"
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
         Left            =   840
         TabIndex        =   3
         Top             =   2160
         Visible         =   0   'False
         Width           =   270
      End
   End
   Begin VB.ListBox lstGuild 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      IntegralHeight  =   0   'False
      ItemData        =   "frmGuild.frx":0185
      Left            =   300
      List            =   "frmGuild.frx":0187
      TabIndex        =   1
      Top             =   240
      Width           =   4200
   End
   Begin VB.Image imgReSize 
      Height          =   180
      Left            =   3240
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmGuild.frx":0189
      Top             =   3720
      Width           =   180
   End
   Begin VB.Image imgclose 
      Height          =   135
      Left            =   3120
      Picture         =   "frmGuild.frx":02D5
      Top             =   60
      Width           =   135
   End
   Begin VB.Label LabGuild 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Guild List"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   23
      Top             =   15
      Width           =   645
   End
   Begin VB.Image imgMidbar 
      Height          =   255
      Left            =   170
      Picture         =   "frmGuild.frx":040A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   180
   End
   Begin VB.Image Image5 
      Height          =   135
      Left            =   50
      Picture         =   "frmGuild.frx":04E2
      Top             =   60
      Width           =   135
   End
   Begin VB.Image imgbright 
      Height          =   420
      Left            =   2400
      Picture         =   "frmGuild.frx":0617
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgbmid 
      Height          =   420
      Left            =   120
      Picture         =   "frmGuild.frx":068C
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgbleft 
      Height          =   420
      Left            =   0
      Picture         =   "frmGuild.frx":06E4
      Top             =   3480
      Width           =   150
   End
   Begin VB.Image Image1 
      Height          =   120
      Left            =   0
      Picture         =   "frmGuild.frx":0763
      Top             =   2400
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image10 
      Height          =   255
      Left            =   0
      Picture         =   "frmGuild.frx":07A6
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgRightbar 
      Height          =   255
      Left            =   1560
      Picture         =   "frmGuild.frx":0914
      Top             =   0
      Width           =   180
   End
End
Attribute VB_Name = "frmGuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.height = 4500
    Me.width = 4300
    imgRightbar.Left = Me.width - 200
    imgMidbar.width = Me.width - 350
    LoadFormPos Me
    Frame1.height = Me.height - 880
    Frame1.width = Me.width
    Label3.height = Frame1.height
    Label3.width = Me.width
    lstGuild.height = Me.height - 880
    lstGuild.width = Me.width
    imgbleft.Top = lstGuild.height + 200
    imgbmid.Top = lstGuild.height + 200
    imgbright.Top = lstGuild.height + 200
    imgbright.Left = Me.width - 300
    imgbmid.width = Me.width - 400
    imgReSize.Top = lstGuild.height + 320
    imgReSize.Left = Me.width - 270
    Me.height = lstGuild.height + 880
    Dim tx, ty, tw, th As Integer
    Dim pw, ph As Long
    frmGuild.AutoRedraw = True
    tw = Int(frmGuild.width / Image1.width) + 1
    th = Int((frmGuild.height) / Image1.height) + 1
    pw = Image1.width
    ph = Image1.height
    For tx = 0 To tw
        For ty = 0 To th
             frmGuild.PaintPicture Image1.Picture, tx * pw, ty * ph
        Next ty
    Next tx
    frmGuild.AutoRedraw = False
    If StartBot Then
        frmMain.Send_GuildRequest
        frmMain.Send_Guildinfo 1
        frmMain.Send_Guildinfo 0
    End If
    UpdateGuild
End Sub

Private Sub Form_Resize()
    If (Me.width < 2000 Or Me.height < 2000) Then
        Form_Load
    Else
    imgRightbar.Left = Me.width - 180
    imgMidbar.width = Me.width - 320
    Frame1.height = Me.height - 650
    Frame1.width = Me.width
    lstGuild.height = Me.height - 650
    lstGuild.width = Me.width
    imgbleft.Top = lstGuild.height + 240
    imgbmid.Top = lstGuild.height + 240
    imgclose.Left = Me.width - 200
    imgbright.Top = lstGuild.height + 240
    imgbright.Left = Me.width - 100
    imgbmid.width = Me.width - 200
    imgReSize.Top = lstGuild.height + 480
    imgReSize.Left = Me.width - 182
    If (Me.height + 650) < MDIfrmMain.height Then Me.height = lstGuild.height + 650
    Dim tx, ty, tw, th As Integer
    Dim pw, ph As Long
    frmGuild.AutoRedraw = True
    tw = Int(frmGuild.width / Image1.width) + 1
    th = Int((frmGuild.height) / Image1.height) + 1
    pw = Image1.width
    ph = Image1.height
    For tx = 0 To tw
        For ty = 0 To th
             frmGuild.PaintPicture Image1.Picture, tx * pw, ty * ph
        Next ty
    Next tx
    frmGuild.AutoRedraw = False
End If
End Sub

Private Sub imgclose_Click()
    SaveFormPos Me
    Unload Me
End Sub

Private Sub imgMidbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ReleaseCapture
    Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    SaveFormPos Me
End Sub

Private Sub imgReSize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ReleaseCapture
    Call SendMessage(hWnd, WM_NCLBUTTONDOWN, 17, 0)
    SaveFormPos Me
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y > 420 Then
        ViewGuild = 0
    Else
        ViewGuild = 1
    End If
    Update_frmGuild
    UpdateGuild
End Sub

