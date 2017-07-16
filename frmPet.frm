VERSION 5.00
Begin VB.Form frmPet 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Pet"
   ClientHeight    =   1680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3300
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1680
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1700
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3320
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Equipment :"
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
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label LabEQ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Backpack"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1200
         TabIndex        =   10
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Relation :"
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
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   750
      End
      Begin VB.Label LabRelate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Embarrass"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   960
         TabIndex        =   8
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label LabStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Satiety"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   840
         TabIndex        =   7
         Top             =   840
         Width           =   480
      End
      Begin VB.Label Lab3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
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
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Lab2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level  :"
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
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   585
      End
      Begin VB.Label lab1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name  :"
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
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   600
      End
      Begin VB.Image Image5 
         Height          =   135
         Left            =   50
         Picture         =   "frmPet.frx":0000
         Top             =   60
         Width           =   135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pet Information"
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
         Top             =   15
         Width           =   1065
      End
      Begin VB.Image imgclose 
         Height          =   135
         Left            =   3120
         Picture         =   "frmPet.frx":0135
         Top             =   60
         Width           =   135
      End
      Begin VB.Label LabLevel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "44"
         Height          =   195
         Left            =   840
         TabIndex        =   2
         Top             =   600
         Width           =   180
      End
      Begin VB.Label LabName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Draco (Earth Pettie)"
         Height          =   195
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   1395
      End
      Begin VB.Image imgMidbar 
         Height          =   255
         Left            =   165
         Picture         =   "frmPet.frx":026A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3060
      End
      Begin VB.Image Image10 
         Height          =   255
         Left            =   0
         Picture         =   "frmPet.frx":0342
         Top             =   0
         Width           =   180
      End
      Begin VB.Image imgRightbar 
         Height          =   255
         Left            =   3120
         Picture         =   "frmPet.frx":04B0
         Top             =   0
         Width           =   180
      End
   End
End
Attribute VB_Name = "frmPet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
 LoadFormPos Me
 If MyPet.Name <> "" Then Update_FrmPet
End Sub

Private Sub Form_Unload(Cancel As Integer)
 SaveFormPos Me
End Sub



Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then
        Me.PopupMenu frmPopupChat.mnuPet
    End If
End Sub

Private Sub imgclose_Click()
PetWinClose = True
Unload Me
End Sub

Private Sub imgMidbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
SaveFormPos Me
End Sub


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
SaveFormPos frmItem
End Sub

