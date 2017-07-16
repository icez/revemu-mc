VERSION 5.00
Begin VB.Form frmMainServer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Select cRO Main Server"
   ClientHeight    =   1695
   ClientLeft      =   6645
   ClientTop       =   7845
   ClientWidth     =   4200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMainServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstcROServer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1080
      ItemData        =   "frmMainServer.frx":0E42
      Left            =   0
      List            =   "frmMainServer.frx":0E55
      TabIndex        =   0
      Top             =   240
      Width           =   4200
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   50
      Picture         =   "frmMainServer.frx":0EEF
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Main cRO Server"
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
      Width           =   1725
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   3975
      Picture         =   "frmMainServer.frx":1024
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   0
      Picture         =   "frmMainServer.frx":1159
      Top             =   0
      Width           =   4200
   End
   Begin VB.Image Image5 
      Height          =   300
      Left            =   3480
      Picture         =   "frmMainServer.frx":1702
      Top             =   1350
      Width           =   630
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   2760
      Picture         =   "frmMainServer.frx":19AC
      Top             =   1350
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   0
      Picture         =   "frmMainServer.frx":1C39
      Top             =   1320
      Width           =   4200
   End
End
Attribute VB_Name = "frmMainServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
LoadFormPos frmMainServer
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveFormPos frmMainServer
End Sub

Private Sub Image2_Click()
frmMainServer.Visible = False
End Sub

Private Sub Image4_Click()
frmMainServer.Visible = False
frmLogin.Visible = True
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Picture = LoadPicture("interface\bt_ok_c.gif")
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Picture = LoadPicture("interface\bt_ok.gif")
End Sub

Private Sub Image5_Click()
    Unload Me
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Picture = LoadPicture("interface\bt_cancel_c.gif")
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Picture = LoadPicture("interface\bt_cancel.gif")
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
SaveFormPos frmMainServer
End Sub

Private Sub lstcROServer_Click()
Dim X As Integer
For X = 0 To lstcROServer.ListCount - 1
   If lstcROServer.Selected(X) Then
       Select Case X
            Case 0
                ServerID = 1
            Case 1
                ServerID = 6
                LoginIP = "61.220.56.147"
            Case 2
                ServerID = 7
                LoginIP = "61.220.62.30"
            Case 3
                ServerID = 7
                LoginIP = "61.220.62.28"
            Case 4
                ServerID = 7
                LoginIP = "61.220.62.26"
       End Select
       Exit For
   End If
Next
End Sub
