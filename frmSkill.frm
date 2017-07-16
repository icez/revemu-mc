VERSION 5.00
Begin VB.Form frmSkill 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Select Skill "
   ClientHeight    =   4890
   ClientLeft      =   5925
   ClientTop       =   6180
   ClientWidth     =   4245
   ControlBox      =   0   'False
   Icon            =   "frmSkill.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstSkill 
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
      Height          =   1080
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   4200
   End
   Begin VB.Label labPts 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
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
      Left            =   1440
      TabIndex        =   2
      Top             =   1680
      Width           =   105
   End
   Begin VB.Label labSPts 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Skill Pts : "
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
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   795
   End
   Begin VB.Image Image3 
      Height          =   360
      Left            =   80
      Picture         =   "frmSkill.frx":0E42
      Top             =   1680
      Width           =   360
   End
   Begin VB.Image imgReSize 
      Height          =   180
      Left            =   1920
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmSkill.frx":118B
      Top             =   1320
      Width           =   180
   End
   Begin VB.Image imgclose 
      Height          =   135
      Left            =   3980
      Picture         =   "frmSkill.frx":12D7
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image5 
      Height          =   135
      Left            =   50
      Picture         =   "frmSkill.frx":140C
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Skill List"
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
      Width           =   570
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   3480
      Picture         =   "frmSkill.frx":1541
      Top             =   1680
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   0
      Picture         =   "frmSkill.frx":17CE
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgMidbar 
      Height          =   255
      Left            =   170
      Picture         =   "frmSkill.frx":193C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgRightbar 
      Height          =   255
      Left            =   1560
      Picture         =   "frmSkill.frx":1A14
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgbleft 
      Height          =   420
      Left            =   0
      Picture         =   "frmSkill.frx":1C7E
      Top             =   1320
      Width           =   150
   End
   Begin VB.Image imgbright 
      Height          =   420
      Left            =   1440
      Picture         =   "frmSkill.frx":1CFD
      Top             =   1320
      Width           =   120
   End
   Begin VB.Image imgbmid 
      Height          =   420
      Left            =   120
      Picture         =   "frmSkill.frx":1D72
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   120
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   3480
      Picture         =   "frmSkill.frx":1DCA
      Top             =   1680
      Visible         =   0   'False
      Width           =   630
   End
End
Attribute VB_Name = "frmSkill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function Return_SkillName(ID As Byte) As String
Dim i As Integer
For i = 0 To UBound(SkillChar) - 1
            If SkillChar(i).ID = ID Then
                Return_SkillName = Get_SkillName(SkillChar(i).Name)
                Exit For
            End If
Next
End Function

Private Sub Form_Default()
frmSkill.height = 3000
frmSkill.width = 3000
imgRightbar.Left = frmSkill.width - 180
imgMidbar.width = frmSkill.width - 320
lstSkill.height = frmSkill.height - 650
lstSkill.width = frmSkill.width
imgclose.Left = frmSkill.width - 200
imgbleft.Top = lstSkill.height + 220
imgbmid.Top = lstSkill.height + 220
imgbright.Top = lstSkill.height + 220
imgbright.Left = frmSkill.width - 120
imgbmid.width = frmSkill.width - 100
imgReSize.Top = lstSkill.height + 480
imgReSize.Left = frmSkill.width - 180
Image1.Top = lstSkill.height + 300
Image2.Top = lstSkill.height + 300
Image3.Top = lstSkill.height + 250
Image1.Left = frmSkill.width - 1500
Image2.Left = frmSkill.width - 800
labSPts.Top = lstSkill.height + 300
labPts.Top = lstSkill.height + 300
MDIfrmMain.mnuSkill.CheckED = True
End Sub

Private Sub Form_Load()
frmMain.UpdateSkills
frmSkill.height = 3000
frmSkill.width = 3000
LoadFormPos frmSkill
imgRightbar.Left = frmSkill.width - 180
imgMidbar.width = frmSkill.width - 320
lstSkill.height = frmSkill.height - 650
lstSkill.width = frmSkill.width
imgclose.Left = frmSkill.width - 200
imgbleft.Top = lstSkill.height + 220
imgbmid.Top = lstSkill.height + 220
imgbright.Top = lstSkill.height + 220
imgbright.Left = frmSkill.width - 120
imgbmid.width = frmSkill.width - 100
imgReSize.Top = lstSkill.height + 480
imgReSize.Left = frmSkill.width - 180
Image1.Top = lstSkill.height + 300
Image2.Top = lstSkill.height + 300
Image3.Top = lstSkill.height + 250
Image1.Left = frmSkill.width - 1500
Image2.Left = frmSkill.width - 800
labSPts.Top = lstSkill.height + 300
labPts.Top = lstSkill.height + 300
MDIfrmMain.mnuSkill.CheckED = True
End Sub

Private Sub Form_Resize()
If (frmSkill.width < 2500 Or frmSkill.height < 2500) Then
Form_Default
Else
imgRightbar.Left = frmSkill.width - 180
imgMidbar.width = frmSkill.width - 320
lstSkill.height = frmSkill.height - 650
lstSkill.width = frmSkill.width
imgclose.Left = frmSkill.width - 200
imgbleft.Top = lstSkill.height + 220
imgbmid.Top = lstSkill.height + 220
imgbright.Top = lstSkill.height + 220
imgbright.Left = frmSkill.width - 120
imgbmid.width = frmSkill.width - 100
imgReSize.Top = lstSkill.height + 480
imgReSize.Left = frmSkill.width - 180
Image1.Top = lstSkill.height + 300
Image2.Top = lstSkill.height + 300
Image3.Top = lstSkill.height + 250
Image1.Left = frmSkill.width - 800
Image2.Left = frmSkill.width - 800
labSPts.Top = lstSkill.height + 300
labPts.Top = lstSkill.height + 300
If (frmSkill.height + 650) < MDIfrmMain.height Then frmSkill.height = lstSkill.height + 650
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIfrmMain.mnuSkill.CheckED = False
SaveFormPos frmSkill
End Sub




Private Sub Image1_Click()
Dim X As Integer
SkillNumber = 5000
skillpacket = ""
IsUseSkill = False
For X = 0 To lstSkill.ListCount - 1
   If lstSkill.Selected(X) Then
       If (SkillChar(X).SP > 0) Then
       SkillNumber = X
       frmSkillLV.Top = frmSkill.Top + 500
       frmSkillLV.Left = frmSkill.Left + 400
       frmSkillLV.Visible = True
       frmSkill.Enabled = False
       Else
       Stat "Can't use this Skill..." + vbCrLf
       End If
       Exit For
   End If
Next
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\interface\bt_ok_c.gif")
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\interface\bt_ok.gif")
End Sub

Private Sub Image2_Click()
frmSkill.Visible = False
End Sub

Private Sub Image3_Click()
Dim X As Integer
For X = 0 To lstSkill.ListCount - 1
   If lstSkill.Selected(X) Then
       If (SkillChar(X).ID > 0) And (SkillChar(X).MaxLV < 10) Then
        frmMain.Update_SkillLV SkillChar(X).ID
        'SkillChar(X).MaxLV = SkillChar(X).MaxLV + 1
        frmMain.UpdateSkills
       ElseIf (SkillChar(X).ID > 0) And (SkillChar(X).MaxLV < 9) Then
        frmMain.Update_SkillLV SkillChar(X).ID
        'SkillChar(X).MaxLV = SkillChar(X).MaxLV + 1
        frmMain.UpdateSkills
       Else
        Stat "This skill reach max LV!" + vbCrLf
       End If
       Exit For
   End If
Next
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = LoadPicture(App.Path & "\interface\bt_jobup_c.gif")
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = LoadPicture(App.Path & "\interface\bt_jobup.gif")
End Sub

Private Sub imgclose_Click()
frmSkill.Visible = False
End Sub

Private Sub imgMidbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
SaveFormPos frmSkill
End Sub

Private Sub imgReSize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(frmSkill.hWnd, WM_NCLBUTTONDOWN, 17, 0)
SaveFormPos frmSkill
End Sub

Private Sub lstSkill_Click()
Dim X As Integer
For X = 0 To lstSkill.ListCount - 1
   If lstSkill.Selected(X) Then
       If (SkillChar(X).ID > 0) And (SkillChar(X).MaxLV < 10) Then
        Image3.Visible = True
       ElseIf (SkillChar(X).ID > 0) And (SkillChar(X).MaxLV < 9) Then
        Image3.Visible = True
       Else
        Image3.Visible = False
       End If
       SkillSelect = X
       Exit For
   End If
Next
End Sub

Private Sub View_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
View.Picture = LoadPicture(App.Path & "\interface\bt_info_c.gif")
End Sub

Private Sub View_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
View.Picture = LoadPicture(App.Path & "\interface\bt_info.gif")
End Sub
