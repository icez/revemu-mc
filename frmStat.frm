VERSION 5.00
Begin VB.Form frmStat 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Status"
   ClientHeight    =   1785
   ClientLeft      =   5280
   ClientTop       =   3795
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
   Icon            =   "frmStat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Begin VB.Label labAspd 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   3650
      TabIndex        =   23
      Top             =   1020
      Width           =   500
   End
   Begin VB.Label labLuk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   825
      TabIndex        =   19
      Top             =   1515
      Width           =   45
   End
   Begin VB.Label labDex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   840
      TabIndex        =   4
      Top             =   1275
      Width           =   45
   End
   Begin VB.Label labInt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   825
      TabIndex        =   3
      Top             =   1035
      Width           =   45
   End
   Begin VB.Label labVit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   825
      TabIndex        =   2
      Top             =   795
      Width           =   45
   End
   Begin VB.Label labAgi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   825
      TabIndex        =   1
      Top             =   555
      Width           =   45
   End
   Begin VB.Label labStr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   825
      TabIndex        =   0
      Top             =   300
      Width           =   45
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   3980
      Picture         =   "frmStat.frx":0E42
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   50
      Picture         =   "frmStat.frx":0F77
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player Status"
      Height          =   210
      Left            =   240
      TabIndex        =   22
      Top             =   20
      Width           =   960
   End
   Begin VB.Image imgTitle 
      Height          =   255
      Left            =   0
      Picture         =   "frmStat.frx":10AC
      Top             =   0
      Width           =   4200
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   50
      Picture         =   "frmStat.frx":1655
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Logon"
      Height          =   210
      Left            =   240
      TabIndex        =   21
      Top             =   20
      Width           =   450
   End
   Begin VB.Image Image4 
      Height          =   135
      Left            =   3980
      Picture         =   "frmStat.frx":178A
      Top             =   60
      Width           =   135
   End
   Begin VB.Label LabLuckp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   210
      Left            =   1530
      TabIndex        =   20
      Top             =   1500
      Width           =   195
   End
   Begin VB.Label labStatPt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   210
      Left            =   2880
      TabIndex        =   18
      Top             =   1260
      Width           =   1275
   End
   Begin VB.Image imgUpLuk 
      Height          =   135
      Left            =   1375
      Picture         =   "frmStat.frx":18BF
      Top             =   1530
      Width           =   90
   End
   Begin VB.Image imgUpDex 
      Height          =   135
      Left            =   1375
      Picture         =   "frmStat.frx":1971
      Top             =   1305
      Width           =   90
   End
   Begin VB.Image imgUpInt 
      Height          =   135
      Left            =   1375
      Picture         =   "frmStat.frx":1A23
      Top             =   1050
      Width           =   90
   End
   Begin VB.Image imgUpVit 
      Height          =   135
      Left            =   1375
      Picture         =   "frmStat.frx":1AD5
      Top             =   825
      Width           =   90
   End
   Begin VB.Image imgUpAgi 
      Height          =   135
      Left            =   1375
      Picture         =   "frmStat.frx":1B87
      Top             =   570
      Width           =   90
   End
   Begin VB.Image ImgUpStr 
      Height          =   135
      Left            =   1375
      Picture         =   "frmStat.frx":1C39
      Top             =   345
      Width           =   90
   End
   Begin VB.Label LabDexp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   210
      Left            =   1530
      TabIndex        =   17
      Top             =   1260
      Width           =   195
   End
   Begin VB.Label LabIntp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   210
      Left            =   1530
      TabIndex        =   16
      Top             =   1020
      Width           =   195
   End
   Begin VB.Label LabVitp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   210
      Left            =   1530
      TabIndex        =   15
      Top             =   780
      Width           =   195
   End
   Begin VB.Label LabAgip 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   210
      Left            =   1530
      TabIndex        =   14
      Top             =   540
      Width           =   195
   End
   Begin VB.Label labStrp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   210
      Left            =   1530
      TabIndex        =   13
      Top             =   300
      Width           =   195
   End
   Begin VB.Label labPt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3120
      TabIndex        =   12
      Top             =   1260
      Width           =   300
   End
   Begin VB.Label labCri 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   2130
      TabIndex        =   11
      Top             =   1035
      Width           =   825
   End
   Begin VB.Label labFlee 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   3495
      TabIndex        =   10
      Top             =   795
      Width           =   645
   End
   Begin VB.Label labHit 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2130
      TabIndex        =   9
      Top             =   795
      Width           =   825
   End
   Begin VB.Label labMdef 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   3495
      TabIndex        =   8
      Top             =   555
      Width           =   645
   End
   Begin VB.Label labMatk 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2130
      TabIndex        =   7
      Top             =   555
      Width           =   825
   End
   Begin VB.Label labDef 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   3495
      TabIndex        =   6
      Top             =   315
      Width           =   645
   End
   Begin VB.Label labAtk 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   2130
      TabIndex        =   5
      Top             =   315
      Width           =   825
   End
   Begin VB.Image imgBG 
      Height          =   1545
      Left            =   0
      Picture         =   "frmStat.frx":1CEB
      Top             =   240
      Width           =   4200
   End
End
Attribute VB_Name = "frmStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
LoadFormPos frmStat
MDIfrmMain.mnuStat.Checked = frmStat.Visible
frmMain.UpdateStats
MDIfrmMain.mnuStat.Checked = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIfrmMain.mnuStat.Checked = False
SaveFormPos frmStat
End Sub

Private Sub Image3_Click()
frmStat.Visible = False
End Sub



Private Sub imgTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
SaveFormPos frmStat
End Sub

Private Sub imgUpAgi_Click()
If Val(labStatPt.Caption) >= Val(LabAgip) Then
    frmMain.Update_Stat (1)
Else
    Stat "You need some stat points" + vbCrLf
End If
End Sub

Private Sub imgUpDex_Click()
If Val(labStatPt.Caption) >= Val(LabDexp) Then
frmMain.Update_Stat (4)
Else
    Stat "You need some stat points" + vbCrLf
End If
End Sub

Private Sub imgUpInt_Click()
If Val(labStatPt.Caption) >= Val(LabIntp) Then
    frmMain.Update_Stat (3)
Else
    Stat "You need some stat points" + vbCrLf
End If
End Sub

Private Sub imgUpLuk_Click()
If Val(labStatPt.Caption) >= Val(LabLuckp) Then
frmMain.Update_Stat (5)
Else
    Stat "You need some stat points" + vbCrLf
End If
End Sub

Private Sub ImgUpStr_Click()
If Val(labStatPt.Caption) >= Val(labStrp) Then
    frmMain.Update_Stat (0)
Else
    Stat "You need some stat points" + vbCrLf
End If
End Sub

Private Sub imgUpVit_Click()
If Val(labStatPt.Caption) >= Val(LabVitp) Then
    frmMain.Update_Stat (2)
Else
    Stat "You need some stat points" + vbCrLf
End If
End Sub


Private Sub labStr_Click()
    frmMain.Update_Stat (3)
End Sub
