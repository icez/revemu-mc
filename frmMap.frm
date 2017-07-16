VERSION 5.00
Begin VB.Form frmMap 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "MAP"
   ClientHeight    =   3930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4245
   Icon            =   "frmMap.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "frmMap"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   262
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   283
   ShowInTaskbar   =   0   'False
   Begin VB.Label labMap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MAP"
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
      TabIndex        =   0
      Top             =   15
      Width           =   330
   End
   Begin VB.Image Image5 
      Height          =   135
      Left            =   3675
      Picture         =   "frmMap.frx":0E42
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   50
      Picture         =   "frmMap.frx":0F77
      Top             =   60
      Width           =   135
   End
   Begin VB.Image imgMidbar 
      Height          =   255
      Left            =   120
      Picture         =   "frmMap.frx":10AC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto AI"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   3585
      Width           =   525
   End
   Begin VB.Image imgAI 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   120
      Picture         =   "frmMap.frx":1184
      Top             =   3570
      Width           =   225
   End
   Begin VB.Image imgRightbar 
      Height          =   255
      Left            =   3675
      Picture         =   "frmMap.frx":1278
      Top             =   0
      Width           =   180
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   0
      Picture         =   "frmMap.frx":14E2
      Top             =   0
      Width           =   180
   End
   Begin VB.Shape Player 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   1875
      Shape           =   3  'Circle
      Top             =   1680
      Width           =   180
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   0
      Picture         =   "frmMap.frx":1650
      Stretch         =   -1  'True
      Top             =   3540
      Width           =   3840
   End
   Begin VB.Image Image1 
      Height          =   7680
      Left            =   0
      Top             =   0
      Width           =   7680
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    'Image1.Top = -(100 * 15)
    'Image1.Picture = LoadPicture("map\alb2trea.JPG")
    frmMap.Width = 256 * 15
    frmMap.Height = 256 * 15
    Player.Left = Int(frmMap.ScaleWidth / 2)
    Player.Top = Int(frmMap.ScaleHeight / 2)
    'MapName = "new_zone01"
    Refresh_MAP 236, 273
    'frmMap.Width = Image1.Width
    'frmMap.Height = Image1.Height + 500
    'frmMap.Caption = CStr(Int(Image1.Width / 15)) & ":" & CStr(Int(Image1.Height / 15))
    'Image1.Left = frmMap.Left - Player.Left
    'Image1.Top = frmMap.Top - Player.Top
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim pos As Coord
    Dim w As Double
    Dim h As Double
    w = Image1.Width / mapW
    h = Image1.Height / mapH
    pos.Y = Int(X / 15) / w
    pos.X = Int((Image1.Height - (Int(Y / 15)))) / h
    move_to pos
    WalkMap = True
End Sub

Public Sub Refresh_MAP(ByVal Y As Long, ByVal X As Long)
    On Error GoTo errie
    'set_MapScale "new_zone01"
    Dim w As Double
    Dim h As Double
    w = Image1.Width / mapW
    h = Image1.Height / mapH
    Image1.Top = Int(frmMap.ScaleHeight / 2) - (Image1.Height - (Y * h))
    Image1.Left = Int(frmMap.ScaleWidth / 2) - (X * w)
    'frmMap.Caption = "MAP (" & CStr(X) & ":" & CStr(Y) & ")"
errie:
    
End Sub


Public Sub update_ImgAI()
    If AutoAI Then
        imgAI.Picture = LoadPicture("interface\on.gif")
    Else
        imgAI.Picture = LoadPicture("interface\off.gif")
    End If
End Sub

Private Sub Image5_Click()
    Unload frmMap
End Sub

Private Sub imgAI_Click()
    AutoAI = Not AutoAI
    update_ImgAI
End Sub

Private Sub imgMidbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ReleaseCapture
    Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub labMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ReleaseCapture
    Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub
