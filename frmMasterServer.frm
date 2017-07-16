VERSION 5.00
Begin VB.Form frmMasterServer 
   BorderStyle     =   0  'None
   Caption         =   "Select Master Server"
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1740
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
      ItemData        =   "frmMasterServer.frx":0000
      Left            =   0
      List            =   "frmMasterServer.frx":0002
      TabIndex        =   0
      Top             =   240
      Width           =   4200
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   2760
      Picture         =   "frmMasterServer.frx":0004
      Top             =   1380
      Width           =   630
   End
   Begin VB.Image Image5 
      Height          =   300
      Left            =   3480
      Picture         =   "frmMasterServer.frx":0291
      Top             =   1380
      Width           =   630
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   3975
      Picture         =   "frmMasterServer.frx":053B
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Master Server"
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
      Width           =   1530
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   50
      Picture         =   "frmMasterServer.frx":0670
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   0
      Picture         =   "frmMasterServer.frx":07A5
      Top             =   1320
      Width           =   4200
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   0
      Picture         =   "frmMasterServer.frx":08E9
      Top             =   0
      Width           =   4200
   End
End
Attribute VB_Name = "frmMasterServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Dim i As Integer
    lstcROServer.Clear
    For i = 0 To UBound(ROServer)
        If ROServer(i).Name <> "" Then
            lstcROServer.AddItem ROServer(i).Name
            If ROServer(i).Name = MasterSelect.Name Then lstcROServer.ListIndex = i
        End If
    Next
    LoadFormPos Me
End Sub

Private Sub Image2_Click()
    Unload Me
End Sub

Private Sub Image4_Click()
    MasterSelect = ROServer(lstcROServer.ListIndex)
    Unload Me
    frmLogin.Visible = True
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Picture = LoadPicture(App.Path & "\interface\bt_ok_c.gif")
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Picture = LoadPicture(App.Path & "\interface\bt_ok.gif")
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
SaveFormPos frmMasterServer
End Sub

Private Sub Image5_Click()
Unload Me
frmLogin.Visible = True
End Sub



Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Picture = LoadPicture(App.Path & "\interface\bt_cancel_c.gif")
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Picture = LoadPicture(App.Path & "\interface\bt_cancel.gif")
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
SaveFormPos frmMasterServer
End Sub

Private Sub lstcROServer_DblClick()
    MasterSelect = ROServer(lstcROServer.ListIndex)
    Unload Me
    frmLogin.Visible = True
End Sub
