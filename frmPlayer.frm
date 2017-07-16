VERSION 5.00
Begin VB.Form frmPlayer 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Basic Info"
   ClientHeight    =   1785
   ClientLeft      =   6510
   ClientTop       =   6180
   ClientWidth     =   4200
   Icon            =   "frmPlayer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleMode       =   0  'User
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Begin VB.Label labStatus 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   720
      Width           =   1215
   End
   Begin VB.Image Image4 
      Height          =   135
      Left            =   3980
      Picture         =   "frmPlayer.frx":0E42
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Basic Info"
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
      TabIndex        =   16
      Top             =   15
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   50
      Picture         =   "frmPlayer.frx":0F77
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   0
      Picture         =   "frmPlayer.frx":10AC
      Top             =   0
      Width           =   4200
   End
   Begin VB.Label labZeny 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   2205
      TabIndex        =   0
      Top             =   1530
      Width           =   90
   End
   Begin VB.Label labtabJobExpBg 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   120
      Left            =   1200
      TabIndex        =   15
      Top             =   1305
      Width           =   1725
   End
   Begin VB.Label labtabBaseEXPBg 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   120
      Left            =   1200
      TabIndex        =   14
      Top             =   1140
      Width           =   1725
   End
   Begin VB.Shape tabHP 
      BackColor       =   &H00F2D7C9&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080FF80&
      Height          =   90
      Left            =   1680
      Top             =   360
      Width           =   1005
   End
   Begin VB.Shape tabSP 
      BackColor       =   &H00F2D7C9&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080FF80&
      Height          =   90
      Left            =   1680
      Top             =   690
      Width           =   795
   End
   Begin VB.Label labPlayerName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   285
      Width           =   45
   End
   Begin VB.Label labClass 
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
      Height          =   210
      Left            =   120
      TabIndex        =   12
      Top             =   495
      Width           =   45
   End
   Begin VB.Label LabHPText 
      BackStyle       =   0  'Transparent
      Caption         =   "HP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1455
      TabIndex        =   11
      Top             =   465
      Width           =   255
   End
   Begin VB.Shape tabHPBg 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H0080FF80&
      Height          =   105
      Left            =   1665
      Top             =   345
      Width           =   1215
   End
   Begin VB.Label LabHP 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0  /  0"
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
      Left            =   1680
      TabIndex        =   10
      Top             =   465
      Width           =   1125
   End
   Begin VB.Shape tabSPbg 
      BackStyle       =   1  'Opaque
      Height          =   105
      Left            =   1665
      Top             =   675
      Width           =   1215
   End
   Begin VB.Label labSPtext 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SP"
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
      Left            =   1455
      TabIndex        =   9
      Top             =   795
      Width           =   195
   End
   Begin VB.Label labSP 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0  /  0"
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
      Left            =   1680
      TabIndex        =   8
      Top             =   795
      Width           =   1125
   End
   Begin VB.Label labBLVtext 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Base Lv."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   645
   End
   Begin VB.Label labJobLvtext 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Job Lv."
      Height          =   195
      Left            =   345
      TabIndex        =   6
      Top             =   1245
      Width           =   525
   End
   Begin VB.Label labBaseLv 
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
      Height          =   210
      Left            =   930
      TabIndex        =   5
      Top             =   1080
      Width           =   45
   End
   Begin VB.Label labJobLv 
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
      Height          =   210
      Left            =   930
      TabIndex        =   4
      Top             =   1245
      Width           =   45
   End
   Begin VB.Label labWeighttext 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Weight :"
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
      Left            =   60
      TabIndex        =   3
      Top             =   1530
      Width           =   585
   End
   Begin VB.Label labWeight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
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
      Left            =   705
      TabIndex        =   2
      Top             =   1530
      Width           =   225
   End
   Begin VB.Label LabZenytext 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zeny :"
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
      Left            =   1695
      TabIndex        =   1
      Top             =   1530
      Width           =   465
   End
   Begin VB.Shape tabBaseEXP 
      BackColor       =   &H0000BAFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      Height          =   90
      Left            =   1215
      Top             =   1155
      Width           =   15
   End
   Begin VB.Shape tabJobEXP 
      BackColor       =   &H0000BAFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00004040&
      Height          =   90
      Left            =   1215
      Top             =   1320
      Width           =   15
   End
   Begin VB.Image Image1 
      Height          =   1545
      Left            =   0
      Picture         =   "frmPlayer.frx":1655
      Top             =   240
      Width           =   4200
   End
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function Return_Skill_Name(ID As Integer) As String
    If ID = &H5 Then
    Return_Skill_Name = "Bash"
    ElseIf ID = &HD Then
    Return_Skill_Name = "Soul Strike"
    End If
End Function


Private Sub Form_Load()
frmMain.UpdatePlayer
MDIfrmMain.mnuPlayer.CheckED = True
LoadFormPos frmPlayer
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIfrmMain.mnuPlayer.CheckED = False
SaveFormPos frmPlayer
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
SaveFormPos frmPlayer
End Sub

Private Sub Image4_Click()
frmPlayer.Visible = False
End Sub
