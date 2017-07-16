VERSION 5.00
Begin VB.Form frmNPCSelect 
   BorderStyle     =   0  'None
   Caption         =   "NPC"
   ClientHeight    =   4380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4380
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstEvent 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   3150
      IntegralHeight  =   0   'False
      ItemData        =   "frmNPCSelect.frx":0000
      Left            =   0
      List            =   "frmNPCSelect.frx":0002
      TabIndex        =   0
      Top             =   240
      Width           =   4200
   End
   Begin VB.Image ImgOk 
      Height          =   300
      Left            =   2640
      Picture         =   "frmNPCSelect.frx":0004
      Top             =   3480
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image ImgCancel 
      Height          =   300
      Left            =   3360
      Picture         =   "frmNPCSelect.frx":0291
      Top             =   3480
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image imgReSize 
      Height          =   180
      Left            =   2040
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmNPCSelect.frx":053B
      Top             =   3480
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgclose 
      Height          =   135
      Left            =   3980
      Picture         =   "frmNPCSelect.frx":0687
      Top             =   60
      Width           =   135
   End
   Begin VB.Label LabName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NPC: Select"
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
      Width           =   840
   End
   Begin VB.Image Image5 
      Height          =   135
      Left            =   50
      Picture         =   "frmNPCSelect.frx":07BC
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "frmNPCSelect.frx":08F1
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgMidbar 
      Height          =   255
      Left            =   170
      Picture         =   "frmNPCSelect.frx":0A5F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgRightbar 
      Height          =   255
      Left            =   1560
      Picture         =   "frmNPCSelect.frx":0B37
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgbmid 
      Height          =   420
      Left            =   120
      Picture         =   "frmNPCSelect.frx":0DA1
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgbright 
      Height          =   420
      Left            =   1560
      Picture         =   "frmNPCSelect.frx":0DF9
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgbleft 
      Height          =   420
      Left            =   0
      Picture         =   "frmNPCSelect.frx":0E6E
      Top             =   3480
      Width           =   150
   End
End
Attribute VB_Name = "frmNPCSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
imgRightbar.Left = Me.Width - 200
imgMidbar.Width = Me.Width - 350
LoadFormPos Me
Me.Height = 3000
Me.Width = 4000
lstEvent.Height = Me.Height - 880
lstEvent.Width = Me.Width
imgbleft.Top = lstEvent.Height + 200
imgbmid.Top = lstEvent.Height + 200
imgbright.Top = lstEvent.Height + 200
imgbright.Left = Me.Width - 300
imgbmid.Width = Me.Width - 400
imgReSize.Top = lstEvent.Height + 320
imgReSize.Left = Me.Width - 270
Me.Height = lstEvent.Height + 880
ImgCancel.Top = lstEvent.Height + 50
ImgOk.Top = lstEvent.Height + 50
ImgCancel.Left = Me.Width - 800
ImgOk.Left = Me.Width - 1200
End Sub

Private Sub Form_Resize()
If (Me.Width < 2000 Or Me.Height < 2000) Then
Form_Load
Else
imgRightbar.Left = Me.Width - 180
imgMidbar.Width = Me.Width - 320
lstEvent.Height = Me.Height - 650
lstEvent.Width = Me.Width
imgbleft.Top = lstEvent.Height + 240
imgbmid.Top = lstEvent.Height + 240
imgclose.Left = Me.Width - 200
imgbright.Top = lstEvent.Height + 240
imgbright.Left = Me.Width - 100
imgbmid.Width = Me.Width - 200
imgReSize.Top = lstEvent.Height + 480
imgReSize.Left = Me.Width - 182
ImgCancel.Top = lstEvent.Height + 280
ImgOk.Top = lstEvent.Height + 280
ImgCancel.Left = Me.Width - 800
ImgOk.Left = Me.Width - 1500
If (Me.Height + 650) < MDIfrmMain.Height Then Me.Height = lstEvent.Height + 650
End If
End Sub

Private Sub imgclose_Click()
    frmMain.Send_TalkCancel
    SaveFormPos Me
    Unload Me
End Sub

Private Sub imgMidbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
SaveFormPos Me
End Sub

Private Sub ImgOk_Click()
    Dim X As Integer
    For X = 0 To lstEvent.ListCount - 1
        If lstEvent.Selected(X) Then
            frmMain.Send_Talk NPCList(X).ID
            Exit For
        End If
    Next
End Sub

Private Sub imgReSize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ReleaseCapture
    Call SendMessage(hwnd, WM_NCLBUTTONDOWN, 17, 0)
    SaveFormPos Me
End Sub

Private Sub lstEvent_DblClick()
    frmMain.Send_TalkResponse lstEvent.ListIndex + 1
    Unload Me
End Sub
