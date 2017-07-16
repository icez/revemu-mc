VERSION 5.00
Begin VB.Form frmSendRaw 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Send Raw Packet"
   ClientHeight    =   1155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   Icon            =   "frmSendRaw.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtRaw 
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
      Height          =   525
      Left            =   0
      MaxLength       =   512
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
   Begin VB.Image Image6 
      Height          =   135
      Left            =   3980
      Picture         =   "frmSendRaw.frx":0E42
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Send Raw Packet"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   20
      Width           =   1305
   End
   Begin VB.Image Image5 
      Height          =   135
      Left            =   50
      Picture         =   "frmSendRaw.frx":0F77
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   2760
      Picture         =   "frmSendRaw.frx":10AC
      Top             =   810
      Width           =   630
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   3480
      Picture         =   "frmSendRaw.frx":1339
      Top             =   810
      Width           =   630
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   0
      Picture         =   "frmSendRaw.frx":15E3
      Top             =   0
      Width           =   4200
   End
   Begin VB.Image Image4 
      Height          =   420
      Left            =   0
      Picture         =   "frmSendRaw.frx":1B8C
      Top             =   760
      Width           =   4200
   End
End
Attribute VB_Name = "frmSendRaw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
LoadFormPos frmSendRaw
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveFormPos frmSendRaw
End Sub

Private Sub Image1_Click()
frmMain.SendRaw
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\interface\bt_ok_c.gif")
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\interface\bt_ok.gif")
End Sub

Private Sub Image2_Click()
frmSendRaw.Visible = False
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\interface\bt_cancel_c.gif")
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\interface\bt_cancel.gif")
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
SaveFormPos frmSendRaw
End Sub

Private Sub Image6_Click()
frmSendRaw.Visible = False
End Sub
