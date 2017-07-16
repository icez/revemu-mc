VERSION 5.00
Begin VB.Form frmNPC 
   BorderStyle     =   0  'None
   Caption         =   "NPC List"
   ClientHeight    =   4065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4065
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstNPC 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   3150
      IntegralHeight  =   0   'False
      ItemData        =   "frmNPC.frx":0000
      Left            =   0
      List            =   "frmNPC.frx":0002
      TabIndex        =   0
      Top             =   240
      Width           =   4200
   End
   Begin VB.Image imgclose 
      Height          =   135
      Left            =   3980
      Picture         =   "frmNPC.frx":0004
      Top             =   60
      Width           =   135
   End
   Begin VB.Image imgReSize 
      Height          =   180
      Left            =   2040
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmNPC.frx":0139
      Top             =   3480
      Width           =   180
   End
   Begin VB.Image Image5 
      Height          =   135
      Left            =   50
      Picture         =   "frmNPC.frx":0285
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NPC List"
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
      Width           =   600
   End
   Begin VB.Image imgMidbar 
      Height          =   255
      Left            =   170
      Picture         =   "frmNPC.frx":03BA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   180
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "frmNPC.frx":0492
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgbleft 
      Height          =   420
      Left            =   0
      Picture         =   "frmNPC.frx":0600
      Top             =   3480
      Width           =   150
   End
   Begin VB.Image imgbright 
      Height          =   420
      Left            =   1560
      Picture         =   "frmNPC.frx":067F
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgbmid 
      Height          =   420
      Left            =   120
      Picture         =   "frmNPC.frx":06F4
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgRightbar 
      Height          =   255
      Left            =   1560
      Picture         =   "frmNPC.frx":074C
      Top             =   0
      Width           =   180
   End
End
Attribute VB_Name = "frmNPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.height = 4500
Me.width = 4300
imgRightbar.Left = Me.width - 200
imgMidbar.width = Me.width - 350
LoadFormPos Me
lstNPC.height = Me.height - 880
lstNPC.width = Me.width
imgbleft.Top = lstNPC.height + 200
imgbmid.Top = lstNPC.height + 200
imgbright.Top = lstNPC.height + 200
imgbright.Left = Me.width - 300
imgbmid.width = Me.width - 400
imgReSize.Top = lstNPC.height + 320
imgReSize.Left = Me.width - 270
Me.height = lstNPC.height + 880
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
lstNPC.height = Me.height - 650
lstNPC.width = Me.width
imgbleft.Top = lstNPC.height + 240
imgbmid.Top = lstNPC.height + 240
imgclose.Left = Me.width - 200
imgbright.Top = lstNPC.height + 240
imgbright.Left = Me.width - 100
imgbmid.width = Me.width - 200
imgReSize.Top = lstNPC.height + 480
imgReSize.Left = Me.width - 182
If (Me.height + 650) < MDIfrmMain.height Then Me.height = lstNPC.height + 650
End If
If frmMain.Visible Then
    If UBound(NPCList) > 0 Then UpdateNPC
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

Private Sub lstNPC_DblClick()
    frmMain.Send_Talk NPCList(lstNPC.ListIndex).ID
End Sub
