VERSION 5.00
Begin VB.Form frmShop 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstShop 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3150
      IntegralHeight  =   0   'False
      ItemData        =   "frmShop.frx":0000
      Left            =   0
      List            =   "frmShop.frx":0002
      TabIndex        =   0
      Top             =   240
      Width           =   4200
   End
   Begin VB.Image Image5 
      Height          =   135
      Left            =   50
      Picture         =   "frmShop.frx":0004
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shop"
      BeginProperty Font 
         Name            =   "Tahoma"
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
      Width           =   360
   End
   Begin VB.Image imgclose 
      Height          =   135
      Left            =   3980
      Picture         =   "frmShop.frx":0139
      Top             =   60
      Width           =   135
   End
   Begin VB.Image imgReSize 
      Height          =   180
      Left            =   2040
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmShop.frx":026E
      Top             =   3480
      Width           =   180
   End
   Begin VB.Image imgRightbar 
      Height          =   255
      Left            =   1560
      Picture         =   "frmShop.frx":03BA
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgMidbar 
      Height          =   255
      Left            =   170
      Picture         =   "frmShop.frx":0624
      Stretch         =   -1  'True
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgbleft 
      Height          =   420
      Left            =   0
      Picture         =   "frmShop.frx":06FC
      Top             =   3480
      Width           =   150
   End
   Begin VB.Image imgbright 
      Height          =   420
      Left            =   1560
      Picture         =   "frmShop.frx":077B
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgbmid 
      Height          =   420
      Left            =   120
      Picture         =   "frmShop.frx":07F0
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "frmShop.frx":0848
      Top             =   0
      Width           =   180
   End
End
Attribute VB_Name = "frmShop"
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
lstShop.height = Me.height - 880
lstShop.width = Me.width
imgbleft.Top = lstShop.height + 200
imgbmid.Top = lstShop.height + 200
imgbright.Top = lstShop.height + 200
imgbright.Left = Me.width - 300
imgbmid.width = Me.width - 400
imgResize.Top = lstShop.height + 320
imgResize.Left = Me.width - 270
Me.height = lstShop.height + 880
'If UBound(NPCList) > 0 Then
'    frmMain.UpdateNPC
'End If
End Sub
Private Sub Form_Resize()
If (Me.width < 2000 Or Me.height < 2000) Then
Form_Load
Else
imgRightbar.Left = Me.width - 180
imgMidbar.width = Me.width - 320
lstShop.height = Me.height - 650
lstShop.width = Me.width
imgbleft.Top = lstShop.height + 240
imgbmid.Top = lstShop.height + 240
imgclose.Left = Me.width - 200
imgbright.Top = lstShop.height + 240
imgbright.Left = Me.width - 100
imgbmid.width = Me.width - 200
imgResize.Top = lstShop.height + 480
imgResize.Left = Me.width - 182
If (Me.height + 650) < MDIfrmMain.height Then Me.height = lstShop.height + 650
End If
If frmMain.Visible Then
    If IsVending Then UpdateShop Else Me.Hide
End If
End Sub

Private Sub imgclose_Click()
    frmMain.Send_ShopClose
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

