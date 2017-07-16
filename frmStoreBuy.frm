VERSION 5.00
Begin VB.Form frmStoreBuy 
   BorderStyle     =   0  'None
   Caption         =   "Tool Dealer : Buy Item"
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4395
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   3150
      IntegralHeight  =   0   'False
      ItemData        =   "frmStoreBuy.frx":0000
      Left            =   0
      List            =   "frmStoreBuy.frx":0007
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   240
      Width           =   4200
   End
   Begin VB.Image imgReSize 
      Height          =   180
      Left            =   2040
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmStoreBuy.frx":0014
      Top             =   3480
      Width           =   180
   End
   Begin VB.Image imgclose 
      Height          =   135
      Left            =   3980
      Picture         =   "frmStoreBuy.frx":0160
      Top             =   60
      Width           =   135
   End
   Begin VB.Label LabName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Dealer : Buy Item"
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
      TabIndex        =   2
      Top             =   15
      Width           =   1560
   End
   Begin VB.Image Image5 
      Height          =   135
      Left            =   50
      Picture         =   "frmStoreBuy.frx":0295
      Top             =   60
      Width           =   135
   End
   Begin VB.Image imgMidbar 
      Height          =   255
      Left            =   170
      Picture         =   "frmStoreBuy.frx":03CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   180
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "frmStoreBuy.frx":04A2
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgRightbar 
      Height          =   255
      Left            =   1560
      Picture         =   "frmStoreBuy.frx":0610
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgbleft 
      Height          =   420
      Left            =   0
      Picture         =   "frmStoreBuy.frx":087A
      Top             =   3480
      Width           =   150
   End
   Begin VB.Image imgbright 
      Height          =   420
      Left            =   1560
      Picture         =   "frmStoreBuy.frx":08F9
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgbmid 
      Height          =   420
      Left            =   120
      Picture         =   "frmStoreBuy.frx":096E
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   120
   End
   Begin VB.Label labNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   45
   End
End
Attribute VB_Name = "frmStoreBuy"
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
lstItem.Height = Me.Height - 880
lstItem.Width = Me.Width
imgbleft.Top = lstItem.Height + 200
imgbmid.Top = lstItem.Height + 200
imgbright.Top = lstItem.Height + 200
imgbright.Left = Me.Width - 300
imgbmid.Width = Me.Width - 400
imgReSize.Top = lstItem.Height + 320
imgReSize.Left = Me.Width - 270
Me.Height = lstItem.Height + 880
End Sub

Private Sub Form_Resize()
If (Me.Width < 2000 Or Me.Height < 2000) Then
Form_Load
Else
imgRightbar.Left = Me.Width - 180
imgMidbar.Width = Me.Width - 320
lstItem.Height = Me.Height - 650
lstItem.Width = Me.Width
imgbleft.Top = lstItem.Height + 240
imgbmid.Top = lstItem.Height + 240
imgclose.Left = Me.Width - 200
imgbright.Top = lstItem.Height + 240
imgbright.Left = Me.Width - 100
imgbmid.Width = Me.Width - 200
imgReSize.Top = lstItem.Height + 480
labNumber.Top = lstItem.Height + 400
imgReSize.Left = Me.Width - 182
If (Me.Height + 650) < MDIfrmMain.Height Then Me.Height = lstItem.Height + 650
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

Private Sub lstItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lstItem.OLEDrag
End Sub

Private Sub lstItem_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectMove
    Data.SetData lstItem
End Sub
