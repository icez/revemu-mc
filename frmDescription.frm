VERSION 5.00
Begin VB.Form frmDescription 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Item Description"
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4110
   Icon            =   "frmDescription.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
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
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   3940
      Picture         =   "frmDescription.frx":0E42
      Top             =   40
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   1530
      Left            =   120
      Picture         =   "frmDescription.frx":0F77
      Top             =   120
      Width           =   1155
   End
   Begin VB.Image Image2 
      Height          =   1365
      Left            =   0
      Picture         =   "frmDescription.frx":10D4
      Top             =   360
      Width           =   4110
   End
   Begin VB.Label labName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   45
   End
   Begin VB.Image imgBar 
      Height          =   345
      Left            =   0
      Picture         =   "frmDescription.frx":1289
      Top             =   0
      Width           =   4110
   End
End
Attribute VB_Name = "frmDescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
LoadFormPos frmDescription
End Sub

Private Sub Image1_Click()
frmDescription.Visible = False
End Sub

Private Sub imgBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
SaveFormPos frmDescription
End Sub
