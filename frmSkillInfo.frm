VERSION 5.00
Begin VB.Form frmSkillInfo 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Skill Description"
   ClientHeight    =   1755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4155
   Icon            =   "frmSkillInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   3960
      Picture         =   "frmSkillInfo.frx":0E42
      Top             =   120
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   1395
      Left            =   0
      Picture         =   "frmSkillInfo.frx":0F77
      Top             =   360
      Width           =   4170
   End
   Begin VB.Label labSkill 
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
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   45
   End
   Begin VB.Image imgBar 
      Height          =   375
      Left            =   0
      Picture         =   "frmSkillInfo.frx":1303
      Top             =   0
      Width           =   4170
   End
End
Attribute VB_Name = "frmSkillInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
LoadFormPos frmSkillInfo
End Sub

Private Sub Image2_Click()
frmSkillInfo.Visible = False
End Sub

Private Sub imgBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
SaveFormPos frmSkillInfo
End Sub
