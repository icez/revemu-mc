VERSION 5.00
Begin VB.Form frmNPCMessage 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4950
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNPC 
      Appearance      =   0  'Flat
      Height          =   3135
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
   Begin VB.Image BtClose 
      Height          =   300
      Left            =   3480
      Picture         =   "frmResponse.frx":0000
      Top             =   3840
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image imgNext 
      Height          =   300
      Left            =   3480
      Picture         =   "frmResponse.frx":03C2
      Top             =   3480
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image imgReSize 
      Height          =   180
      Left            =   2040
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmResponse.frx":078F
      Top             =   3480
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label labName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NPC"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   15
      Width           =   330
   End
   Begin VB.Image imgclose 
      Height          =   135
      Left            =   3980
      Picture         =   "frmResponse.frx":08DB
      Top             =   60
      Width           =   135
   End
   Begin VB.Image imgMidbar 
      Height          =   255
      Left            =   170
      Picture         =   "frmResponse.frx":0A10
      Stretch         =   -1  'True
      Top             =   0
      Width           =   180
   End
   Begin VB.Image Image5 
      Height          =   135
      Left            =   50
      Picture         =   "frmResponse.frx":0AE8
      Top             =   60
      Width           =   135
   End
   Begin VB.Image imgRightbar 
      Height          =   255
      Left            =   1560
      Picture         =   "frmResponse.frx":0C1D
      Top             =   0
      Width           =   180
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "frmResponse.frx":0E87
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgbmid 
      Height          =   420
      Left            =   120
      Picture         =   "frmResponse.frx":0FF5
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgbright 
      Height          =   420
      Left            =   1560
      Picture         =   "frmResponse.frx":104D
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgbleft 
      Height          =   420
      Left            =   0
      Picture         =   "frmResponse.frx":10C2
      Top             =   3480
      Width           =   150
   End
End
Attribute VB_Name = "frmNPCMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btClose_Click()
    frmMain.Send_TalkCancel
    btClose.Visible = False
    Unload Me
End Sub

Private Sub Form_Load()
imgRightbar.Left = Me.Width - 200
imgMidbar.Width = Me.Width - 350
LoadFormPos Me
Me.Height = 3000
Me.Width = 4000
txtNPC.Height = Me.Height - 880
txtNPC.Width = Me.Width
imgbleft.Top = txtNPC.Height + 200
imgbmid.Top = txtNPC.Height + 200
imgbright.Top = txtNPC.Height + 200
imgbright.Left = Me.Width - 300
imgbmid.Width = Me.Width - 400
imgReSize.Top = txtNPC.Height + 320
imgReSize.Left = Me.Width - 270
Me.Height = txtNPC.Height + 880
imgNext.Top = txtNPC.Height + 50
imgNext.Left = Me.Width - 800
btClose.Top = txtNPC.Height + 50
btClose.Left = Me.Width - 800
End Sub

Private Sub Form_Resize()
If (Me.Width < 2000 Or Me.Height < 2000) Then
Form_Load
Else
imgRightbar.Left = Me.Width - 180
imgMidbar.Width = Me.Width - 320
txtNPC.Height = Me.Height - 650
txtNPC.Width = Me.Width
imgbleft.Top = txtNPC.Height + 240
imgbmid.Top = txtNPC.Height + 240
imgclose.Left = Me.Width - 200
imgbright.Top = txtNPC.Height + 240
imgbright.Left = Me.Width - 100
imgbmid.Width = Me.Width - 200
imgReSize.Top = txtNPC.Height + 480
imgReSize.Left = Me.Width - 182
imgNext.Top = txtNPC.Height + 280
imgNext.Left = Me.Width - 800
btClose.Top = txtNPC.Height + 280
btClose.Left = Me.Width - 800
If (Me.Height + 650) < MDIfrmMain.Height Then Me.Height = txtNPC.Height + 650
End If
End Sub

Private Sub imgclose_Click()
    frmMain.Send_TalkCancel
    btClose.Visible = False
    SaveFormPos Me
    Unload Me
End Sub

Private Sub imgMidbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
SaveFormPos Me
End Sub

Private Sub imgNext_Click()
    imgNext.Visible = False
    'btClose.Visible = True
    frmMain.Send_TalkContinue
End Sub

Private Sub imgReSize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ReleaseCapture
    Call SendMessage(hwnd, WM_NCLBUTTONDOWN, 17, 0)
    SaveFormPos Me
End Sub
