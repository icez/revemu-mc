VERSION 5.00
Begin VB.Form frmMonster 
   BorderStyle     =   0  'None
   Caption         =   "Monster List"
   ClientHeight    =   4065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4065
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstMonster 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3150
      IntegralHeight  =   0   'False
      ItemData        =   "frmMonster.frx":0000
      Left            =   0
      List            =   "frmMonster.frx":0002
      TabIndex        =   0
      Top             =   240
      Width           =   4200
   End
   Begin VB.Image imgReSize 
      Height          =   180
      Left            =   2040
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmMonster.frx":0004
      Top             =   3480
      Width           =   180
   End
   Begin VB.Image imgclose 
      Height          =   135
      Left            =   3980
      Picture         =   "frmMonster.frx":0150
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Monster List"
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
      TabIndex        =   1
      Top             =   15
      Width           =   885
   End
   Begin VB.Image imgMidbar 
      Height          =   255
      Left            =   170
      Picture         =   "frmMonster.frx":0285
      Stretch         =   -1  'True
      Top             =   0
      Width           =   180
   End
   Begin VB.Image Image5 
      Height          =   135
      Left            =   50
      Picture         =   "frmMonster.frx":035D
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "frmMonster.frx":0492
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgRightbar 
      Height          =   255
      Left            =   1560
      Picture         =   "frmMonster.frx":0600
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgbmid 
      Height          =   420
      Left            =   120
      Picture         =   "frmMonster.frx":086A
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgbright 
      Height          =   420
      Left            =   1560
      Picture         =   "frmMonster.frx":08C2
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgbleft 
      Height          =   420
      Left            =   0
      Picture         =   "frmMonster.frx":0937
      Top             =   3480
      Width           =   150
   End
End
Attribute VB_Name = "frmMonster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Height = 4500
Me.Width = 4300
imgRightbar.Left = Me.Width - 200
imgMidbar.Width = Me.Width - 350
LoadFormPos Me
lstMonster.Height = Me.Height - 880
lstMonster.Width = Me.Width
imgbleft.Top = lstMonster.Height + 200
imgbmid.Top = lstMonster.Height + 200
imgbright.Top = lstMonster.Height + 200
imgbright.Left = Me.Width - 300
imgbmid.Width = Me.Width - 400
imgReSize.Top = lstMonster.Height + 320
imgReSize.Left = Me.Width - 270
Me.Height = lstMonster.Height + 880
End Sub

Private Sub Form_Resize()
If (Me.Width < 2000 Or Me.Height < 2000) Then
Form_Load
Else
imgRightbar.Left = Me.Width - 180
imgMidbar.Width = Me.Width - 320
lstMonster.Height = Me.Height - 650
lstMonster.Width = Me.Width
imgbleft.Top = lstMonster.Height + 240
imgbmid.Top = lstMonster.Height + 240
imgclose.Left = Me.Width - 200
imgbright.Top = lstMonster.Height + 240
imgbright.Left = Me.Width - 100
imgbmid.Width = Me.Width - 200
imgReSize.Top = lstMonster.Height + 480
imgReSize.Left = Me.Width - 182
If (Me.Height + 650) < MDIfrmMain.Height Then Me.Height = lstMonster.Height + 650
End If
End Sub

Private Sub imgclose_Click()
    SaveFormPos Me
    Unload Me
End Sub

Private Sub imgMidbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
SaveFormPos Me
End Sub

Private Sub imgReSize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ReleaseCapture
    Call SendMessage(hwnd, WM_NCLBUTTONDOWN, 17, 0)
    SaveFormPos Me
End Sub

