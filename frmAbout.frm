VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   ClientHeight    =   5340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   4695
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dsfsdfsdf"
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   720
   End
   Begin VB.Image Image4 
      Height          =   135
      Left            =   3960
      Picture         =   "frmAbout.frx":0E42
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   0
      Picture         =   "frmAbout.frx":0F77
      Top             =   4920
      Width           =   4200
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   50
      Picture         =   "frmAbout.frx":10BB
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   0
      Top             =   20
      Width           =   105
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "frmAbout.frx":11F0
      Top             =   0
      Width           =   4200
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmAbout.Caption = "About Rêvemu 0.87.1"
Label1.Caption = "About Rêvemu 0.87.1"
Label2.Caption = "Rêvemu Team: Dash, Lunar|Shadow and Kyne"
Text1.text = Text1.text & "All Rêvemu and Rêvemu forum users must agree to the following:" + vbCrLf & vbCrLf
Text1.text = Text1.text & "1) All matters regarding any version of ROE, ROP, K-Bot (including K-OC Bot), and Rêvemu are to be SOLELY discussed ONLY on these forums. If the forums are down, then you must wait. Posting such info on other sites could be considered a breach of this agreement." + vbCrLf & vbCrLf
Text1.text = Text1.text & "2) All Rêvemu users MUST agree to keep ALL information, materials, and data given to them private. This means you cannot discuss your unique 'key' for using Rêvemu, and you cannot send the Rêvemu, ROE, K-Bot, ROP, or K-OC Bot or ANY of their associated files to ANYONE else. ALL of these include not doing any of the previous with anyone, including other Rêvemu users. This also means you cannot host such files." + vbCrLf & vbCrLf
Text1.text = Text1.text & "3) All users must agree that they will report sites hosting any ROE, ROP, K-Bot, K-OC Bot, or Rêvemu related materials. They also must agree that they will report TRUTHFULLY any members breaching any of these agreements. Should we be provided with false info regarding a user, the info-provider will most likely be banned under the conditions of term 4." + vbCrLf & vbCrLf
Text1.text = Text1.text & "4) All users must agree that breaching the previous statements can mean a permanent ban from this forum, the upcoming chat room, and from the R?vemu. They also must agree they will abide by any future revisions to this agreement or face the previous penalty."
End Sub




Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub Image4_Click()
frmAbout.Visible = False
End Sub
