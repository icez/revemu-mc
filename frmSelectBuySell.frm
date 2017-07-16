VERSION 5.00
Begin VB.Form frmSelectBuySell 
   BorderStyle     =   0  'None
   Caption         =   "Select Buy/Sell"
   ClientHeight    =   660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   660
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image4 
      Height          =   300
      Left            =   2760
      Picture         =   "frmSelectBuySell.frx":0000
      Top             =   300
      Width           =   630
   End
   Begin VB.Image Image5 
      Height          =   300
      Left            =   3480
      Picture         =   "frmSelectBuySell.frx":0289
      Top             =   300
      Width           =   630
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   3975
      Picture         =   "frmSelectBuySell.frx":0642
      Top             =   60
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Do you wan to Buy or Sell ?"
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
      Width           =   2025
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   50
      Picture         =   "frmSelectBuySell.frx":0777
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   0
      Picture         =   "frmSelectBuySell.frx":08AC
      Top             =   240
      Width           =   4200
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   0
      Picture         =   "frmSelectBuySell.frx":09F0
      Top             =   0
      Width           =   4200
   End
End
Attribute VB_Name = "frmSelectBuySell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Image4_Click()
    frmMain.Send_GetStoreList
    Unload Me
    SaveFormPos Me
End Sub

Private Sub Image5_Click()
    frmMain.Send_Sell
    Unload Me
    SaveFormPos Me
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ReleaseCapture
    Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    SaveFormPos Me
End Sub
