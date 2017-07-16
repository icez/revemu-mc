VERSION 5.00
Begin VB.Form frmPeople 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Other Players"
   ClientHeight    =   5850
   ClientLeft      =   7485
   ClientTop       =   5415
   ClientWidth     =   4200
   Icon            =   "frmPeople.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrPeople 
      Interval        =   2000
      Left            =   1440
      Top             =   4200
   End
   Begin VB.ListBox lstPeople 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   3150
      IntegralHeight  =   0   'False
      ItemData        =   "frmPeople.frx":0E42
      Left            =   0
      List            =   "frmPeople.frx":0E44
      TabIndex        =   0
      Top             =   240
      Width           =   4200
   End
   Begin VB.Image imgclose 
      Height          =   135
      Left            =   3980
      Picture         =   "frmPeople.frx":0E46
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player List"
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
      Width           =   750
   End
   Begin VB.Image Image5 
      Height          =   135
      Left            =   50
      Picture         =   "frmPeople.frx":0F7B
      Top             =   60
      Width           =   135
   End
   Begin VB.Image imgReSize 
      Height          =   180
      Left            =   2040
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmPeople.frx":10B0
      Top             =   3480
      Width           =   180
   End
   Begin VB.Image imgRightbar 
      Height          =   255
      Left            =   1560
      Picture         =   "frmPeople.frx":11FC
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgMidbar 
      Height          =   255
      Left            =   170
      Picture         =   "frmPeople.frx":1466
      Stretch         =   -1  'True
      Top             =   0
      Width           =   180
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "frmPeople.frx":153E
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgbmid 
      Height          =   420
      Left            =   120
      Picture         =   "frmPeople.frx":16AC
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgbright 
      Height          =   420
      Left            =   1560
      Picture         =   "frmPeople.frx":1704
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgbleft 
      Height          =   420
      Left            =   0
      Picture         =   "frmPeople.frx":1779
      Top             =   3480
      Width           =   150
   End
End
Attribute VB_Name = "frmPeople"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
upd_frmPeople
Me.height = 4500
Me.width = 4300
imgRightbar.Left = Me.width - 200
imgMidbar.width = Me.width - 350
LoadFormPos Me
lstPeople.height = Me.height - 880
lstPeople.width = Me.width
imgbleft.Top = lstPeople.height + 200
imgbmid.Top = lstPeople.height + 200
imgbright.Top = lstPeople.height + 200
imgbright.Left = Me.width - 300
imgbmid.width = Me.width - 400
imgReSize.Top = lstPeople.height + 320
imgReSize.Left = Me.width - 270
Me.height = lstPeople.height + 880
End Sub



Private Sub Form_Resize()
If (Me.width < 2000 Or Me.height < 2000) Then
Form_Load
Else
imgRightbar.Left = Me.width - 180
imgMidbar.width = Me.width - 320
lstPeople.height = Me.height - 650
lstPeople.width = Me.width
imgbleft.Top = lstPeople.height + 240
imgbmid.Top = lstPeople.height + 240
imgclose.Left = Me.width - 200
imgbright.Top = lstPeople.height + 240
imgbright.Left = Me.width - 100
imgbmid.width = Me.width - 200
imgReSize.Top = lstPeople.height + 480
imgReSize.Left = Me.width - 182
If (Me.height + 650) < MDIfrmMain.height Then Me.height = lstPeople.height + 650
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIfrmMain.mnuInv.CheckED = False
SaveFormPos Me
End Sub

Private Sub imgclose_Click()
Me.Visible = False
End Sub

Private Sub imgMidbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
SaveFormPos Me
End Sub

Private Sub imgReSize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, 17, 0)
SaveFormPos Me
End Sub

Private Sub tmrPeople_Timer()
    upd_frmPeople
End Sub
