VERSION 5.00
Begin VB.Form frmCharSelect 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Char Selects"
   ClientHeight    =   1860
   ClientLeft      =   6645
   ClientTop       =   7845
   ClientWidth     =   4200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCharSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1200
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   4200
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   50
      Picture         =   "frmCharSelect.frx":0E42
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Character"
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
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   3975
      Picture         =   "frmCharSelect.frx":0F77
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   0
      Picture         =   "frmCharSelect.frx":10AC
      Top             =   0
      Width           =   4200
   End
   Begin VB.Image Image5 
      Height          =   300
      Left            =   3480
      Picture         =   "frmCharSelect.frx":1655
      Top             =   1480
      Width           =   630
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   2760
      Picture         =   "frmCharSelect.frx":18FF
      Top             =   1480
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   0
      Picture         =   "frmCharSelect.frx":1B8C
      Top             =   1440
      Width           =   4200
   End
End
Attribute VB_Name = "frmCharSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
LoadFormPosOnly frmCharSelect
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveFormPos frmCharSelect
End Sub

Private Sub Image2_Click()
frmCharSelect.Visible = False
End Sub

Private Sub Image4_Click()
frmCharSelect.Visible = False
frmMain.cmdChar_Click
IsConnected = True
MDIfrmMain.Caption = Players(number).Name & " - Powered by " & Version
'StartZeny = Players(number).Zeny
If frmLogin.chkSave.value Then MDIfrmMain.Save_User
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Picture = LoadPicture(App.Path & "\interface\bt_ok_c.gif")
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Picture = LoadPicture(App.Path & "\interface\bt_ok.gif")
End Sub

Private Sub Image5_Click()
frmCharSelect.Visible = False
End Sub



Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Picture = LoadPicture(App.Path & "\interface\bt_cancel_c.gif")
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Picture = LoadPicture(App.Path & "\interface\bt_cancel.gif")
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
SaveFormPos frmCharSelect
End Sub

Private Sub List1_Click()
Dim X As Integer
For X = 0 To List1.ListCount - 1
   If List1.Selected(X) Then
       number = Val(Left(List1.List(X), 1))
       frmPlayer.labBaseLv.Caption = CStr(Players(Val(Left(List1.List(X), 1))).BaseLV)
       frmPlayer.labJobLv.Caption = CStr(Players(Val(Left(List1.List(X), 1))).JobLV)
       frmPlayer.labPlayerName.Caption = Players(Val(Left(List1.List(X), 1))).Name
       frmPlayer.labSP.Caption = CStr(Players(number).Sp) + "  /  " + CStr(Players(number).maxsp)
       frmPlayer.labSP.Caption = CStr(Players(number).Sp) + "  /  " + CStr(Players(number).maxsp)
       frmPlayer.tabSP.width = (Players(number).Sp / IIf(Players(number).maxsp > 0, Players(number).maxsp, 1)) * (frmPlayer.tabSPbg.width - 20)
       frmPlayer.LabHP.Caption = CStr(Players(number).HP) + "  /  " + CStr(Players(number).MaxHP)
       frmPlayer.LabHP.Caption = CStr(Players(number).HP) + "  /  " + CStr(Players(number).MaxHP)
       frmPlayer.tabHP.width = (Players(number).HP / IIf(Players(number).MaxHP > 0, Players(number).MaxHP, 1)) * (frmPlayer.tabHPBg.width - 20)
       frmPlayer.labZeny.Caption = Format(Players(number).Zeny, "##,##")
       frmStat.labStatPt.Caption = CStr(Players(number).StatPoint)
       frmPlayer.labClass.Caption = Players(number).Class
       Exit For
   End If
Next
oldBaseEXP = Players(number).BaseExp
Check_JobBar
If Players(number).BaseLV = 99 Then
    frmPlayer.labtabBaseEXPBg.Visible = False
    frmPlayer.tabBaseEXP.Visible = False
Else
    frmPlayer.labtabBaseEXPBg.Visible = True
    frmPlayer.tabBaseEXP.Visible = True
End If
End Sub
