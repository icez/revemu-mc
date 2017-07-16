VERSION 5.00
Begin VB.Form frmServer 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Select Servers"
   ClientHeight    =   1950
   ClientLeft      =   6255
   ClientTop       =   6495
   ClientWidth     =   4215
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox LstServer 
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
      Height          =   1290
      ItemData        =   "frmServer.frx":0E42
      Left            =   0
      List            =   "frmServer.frx":0E44
      TabIndex        =   0
      Top             =   240
      Width           =   4210
   End
   Begin VB.Image Image6 
      Height          =   135
      Left            =   3980
      Picture         =   "frmServer.frx":0E46
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Server"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   15
      Width           =   1215
   End
   Begin VB.Image Image4 
      Height          =   135
      Left            =   50
      Picture         =   "frmServer.frx":0F7B
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "frmServer.frx":10B0
      Top             =   0
      Width           =   4200
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   3480
      Picture         =   "frmServer.frx":1659
      Top             =   1580
      Width           =   630
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   2760
      Picture         =   "frmServer.frx":1903
      Top             =   1575
      Width           =   630
   End
   Begin VB.Image Image5 
      Height          =   420
      Left            =   0
      Picture         =   "frmServer.frx":1B90
      Top             =   1520
      Width           =   4200
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
LoadFormPos frmServer
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveFormPos frmServer
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
SaveFormPos frmServer
End Sub

Private Sub Image2_Click()
Dim X As Integer
For X = 0 To LstServer.ListCount - 1
    If LstServer.Selected(X) Then
        NumServ = X
        Exit For
    End If
Next
frmServer.Visible = False
Stat "CSrv:Connecting to " & ServerList(NumServ).IP & ":" & CStr(ServerList(NumServ).Port) & "...."
CurCIP = ServerList(NumServ).IP
DoConnect ServerList(NumServ).IP, CLng(ServerList(NumServ).Port)
frmMain.tmrResponse.Enabled = True
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\interface\bt_ok_c.gif")
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\interface\bt_ok.gif")
End Sub

Private Sub Image3_Click()
    Unload Me
    End
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = LoadPicture(App.Path & "\interface\bt_cancel_c.gif")
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = LoadPicture(App.Path & "\interface\bt_cancel.gif")
End Sub

Private Sub Image6_Click()
frmServer.Visible = False
End Sub

