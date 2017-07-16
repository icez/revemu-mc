VERSION 5.00
Begin VB.Form frmParty 
   BorderStyle     =   0  'None
   Caption         =   "Party List"
   ClientHeight    =   4065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4065
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstParty 
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
      Height          =   3255
      IntegralHeight  =   0   'False
      ItemData        =   "frmParty.frx":0000
      Left            =   0
      List            =   "frmParty.frx":0002
      TabIndex        =   0
      Top             =   240
      Width           =   4200
   End
   Begin VB.Image imgReSize 
      Height          =   180
      Left            =   3240
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmParty.frx":0004
      Top             =   3720
      Width           =   180
   End
   Begin VB.Image Image5 
      Height          =   135
      Left            =   50
      Picture         =   "frmParty.frx":0150
      Top             =   60
      Width           =   135
   End
   Begin VB.Image imgclose 
      Height          =   135
      Left            =   3120
      Picture         =   "frmParty.frx":0285
      Top             =   60
      Width           =   135
   End
   Begin VB.Label LabParty 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Party List"
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
      TabIndex        =   1
      Top             =   15
      Width           =   645
   End
   Begin VB.Image imgRightbar 
      Height          =   255
      Left            =   1560
      Picture         =   "frmParty.frx":03BA
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgbright 
      Height          =   420
      Left            =   2400
      Picture         =   "frmParty.frx":0624
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgbmid 
      Height          =   420
      Left            =   120
      Picture         =   "frmParty.frx":0699
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgbleft 
      Height          =   420
      Left            =   0
      Picture         =   "frmParty.frx":06F1
      Top             =   3480
      Width           =   150
   End
   Begin VB.Image imgMidbar 
      Height          =   255
      Left            =   170
      Picture         =   "frmParty.frx":0770
      Stretch         =   -1  'True
      Top             =   0
      Width           =   180
   End
   Begin VB.Image Image10 
      Height          =   255
      Left            =   0
      Picture         =   "frmParty.frx":0848
      Top             =   0
      Width           =   180
   End
End
Attribute VB_Name = "frmParty"
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
    lstParty.height = Me.height - 880
    lstParty.width = Me.width
    imgbleft.Top = lstParty.height + 200
    imgbmid.Top = lstParty.height + 200
    imgbright.Top = lstParty.height + 200
    imgbright.Left = Me.width - 300
    imgbmid.width = Me.width - 400
    imgReSize.Top = lstParty.height + 320
    imgReSize.Left = Me.width - 270
    Me.height = lstParty.height + 880
    frmMain.UpdateParty
End Sub

Private Sub Form_Resize()
    If (Me.width < 2000 Or Me.height < 2000) Then
        Form_Load
    Else
    imgRightbar.Left = Me.width - 180
    imgMidbar.width = Me.width - 320
    lstParty.height = Me.height - 650
    lstParty.width = Me.width
    imgbleft.Top = lstParty.height + 240
    imgbmid.Top = lstParty.height + 240
    imgclose.Left = Me.width - 200
    imgbright.Top = lstParty.height + 240
    imgbright.Left = Me.width - 100
    imgbmid.width = Me.width - 200
    imgReSize.Top = lstParty.height + 480
    imgReSize.Left = Me.width - 182
    If (Me.height + 650) < MDIfrmMain.height Then _
        Me.height = lstParty.height + 650
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

Private Sub lstParty_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        If lstParty.List(lstParty.ListIndex) <> "" Then
            frmPopupChat.mnuKickParty.Visible = True
            frmPopupChat.mnuLeaveParty.Visible = True
            frmPopupChat.mnuGuilds.Visible = False
            Me.PopupMenu frmPopupChat.mnuPartylist
        End If
    End If
End Sub
