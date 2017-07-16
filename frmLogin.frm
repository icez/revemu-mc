VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   1800
   ClientLeft      =   4185
   ClientTop       =   2025
   ClientWidth     =   4200
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CheckBox chkSave 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "keep"
      DisabledPicture =   "frmLogin.frx":0E42
      DownPicture     =   "frmLogin.frx":0EBA
      ForeColor       =   &H00000000&
      Height          =   200
      Left            =   3360
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmLogin.frx":101D
      TabIndex        =   2
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.TextBox txtPass 
      Appearance      =   0  'Flat
      BackColor       =   &H00F3F7F7&
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
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "•"
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      BackColor       =   &H00F3F7F7&
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
      Height          =   195
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Image Image4 
      Height          =   135
      Left            =   3980
      Picture         =   "frmLogin.frx":1095
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Logon"
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
      TabIndex        =   3
      Top             =   20
      Width           =   450
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   50
      Picture         =   "frmLogin.frx":11CA
      Top             =   60
      Width           =   135
   End
   Begin VB.Image imgTitle 
      Height          =   255
      Left            =   0
      Picture         =   "frmLogin.frx":12FF
      Top             =   0
      Width           =   4200
   End
   Begin VB.Image imgLogin 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2760
      Picture         =   "frmLogin.frx":18A8
      Top             =   1440
      Width           =   630
   End
   Begin VB.Image imgExit 
      Height          =   300
      Left            =   3480
      Picture         =   "frmLogin.frx":1C7F
      Top             =   1440
      Width           =   630
   End
   Begin VB.Image Image2 
      Height          =   1155
      Left            =   0
      Picture         =   "frmLogin.frx":1F29
      Top             =   240
      Width           =   4200
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   0
      Picture         =   "frmLogin.frx":2368
      Top             =   1400
      Width           =   4200
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
'SaveFormPos frmLogin
LoadFormPos frmLogin
'txtPass.Text = MakeHex(frmMain.MakeTickString)
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveFormPos frmLogin
End Sub

Private Sub Image4_Click()
frmLogin.Visible = False
End Sub

Private Sub imgExit_Click()
frmLogin.Visible = False
End Sub

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgExit.Picture = LoadPicture(App.Path & "\interface\bt_cancel_c.gif")
End Sub

Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgExit.Picture = LoadPicture(App.Path & "\interface\bt_cancel.gif")
End Sub

Private Sub imgLogin_Click()
    strUser = txtUser.text
    strPass = txtPass.text
    If (chkSave.value = 1) Then MDIfrmMain.Save_User
    frmLogin.Visible = False
    frmMain.Visible = True
    frmMain.Main_Init
    MDIfrmMain.mnuReset.Visible = True
    Dead = False
End Sub

Private Sub imgLogin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgLogin.Picture = LoadPicture(App.Path & "\interface\bt_login_c.gif")
End Sub

Private Sub imgLogin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgLogin.Picture = LoadPicture(App.Path & "\interface\bt_login.gif")
End Sub

Private Sub imgTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
SaveFormPos frmLogin
End Sub
