VERSION 5.00
Begin VB.Form frmStatus 
   BorderStyle     =   0  'None
   Caption         =   "frmStatus"
   ClientHeight    =   4065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4065
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   3150
      IntegralHeight  =   0   'False
      ItemData        =   "frmStatus.frx":0000
      Left            =   0
      List            =   "frmStatus.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   4200
   End
   Begin VB.Image imgclose 
      Height          =   135
      Left            =   3980
      Picture         =   "frmStatus.frx":0004
      Top             =   60
      Width           =   135
   End
   Begin VB.Image imgReSize 
      Height          =   180
      Left            =   2040
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmStatus.frx":0139
      Top             =   3480
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      Width           =   465
   End
   Begin VB.Image imgRightbar 
      Height          =   255
      Left            =   1560
      Picture         =   "frmStatus.frx":0285
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgbmid 
      Height          =   420
      Left            =   120
      Picture         =   "frmStatus.frx":04EF
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgbright 
      Height          =   420
      Left            =   1560
      Picture         =   "frmStatus.frx":0547
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgbleft 
      Height          =   420
      Left            =   0
      Picture         =   "frmStatus.frx":05BC
      Top             =   3480
      Width           =   150
   End
   Begin VB.Image imgMidbar 
      Height          =   255
      Left            =   170
      Picture         =   "frmStatus.frx":063B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   180
   End
   Begin VB.Image Image5 
      Height          =   135
      Left            =   50
      Picture         =   "frmStatus.frx":0713
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "frmStatus.frx":0848
      Top             =   0
      Width           =   180
   End
End
Attribute VB_Name = "frmStatus"
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
lstStatus.Height = Me.Height - 880
lstStatus.Width = Me.Width
imgbleft.Top = lstStatus.Height + 200
imgbmid.Top = lstStatus.Height + 200
imgbright.Top = lstStatus.Height + 200
imgbright.Left = Me.Width - 300
imgbmid.Width = Me.Width - 400
imgReSize.Top = lstStatus.Height + 320
imgReSize.Left = Me.Width - 270
Me.Height = lstStatus.Height + 880
'If UBound(NPCList) > 0 Then
'    frmMain.UpdateNPC
'End If
End Sub

Private Sub Form_Resize()
If (Me.Width < 2000 Or Me.Height < 2000) Then
Form_Load
Else
imgRightbar.Left = Me.Width - 180
imgMidbar.Width = Me.Width - 320
lstStatus.Height = Me.Height - 650
lstStatus.Width = Me.Width
imgbleft.Top = lstStatus.Height + 240
imgbmid.Top = lstStatus.Height + 240
imgclose.Left = Me.Width - 200
imgbright.Top = lstStatus.Height + 240
imgbright.Left = Me.Width - 100
imgbmid.Width = Me.Width - 200
imgReSize.Top = lstStatus.Height + 480
imgReSize.Left = Me.Width - 182
If (Me.Height + 650) < MDIfrmMain.Height Then Me.Height = lstStatus.Height + 650
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
