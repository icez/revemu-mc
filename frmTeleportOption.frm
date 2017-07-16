VERSION 5.00
Begin VB.Form frmTeleportOption 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2100
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox imgWarpAll 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "frmTeleportOption.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   12
      Top             =   1080
      Width           =   225
   End
   Begin VB.PictureBox imgJTele 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "frmTeleportOption.frx":011B
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   11
      Top             =   1320
      Width           =   225
   End
   Begin VB.TextBox txtJTele 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1054
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Text            =   "Acolyte&Priest&Novice"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.PictureBox imgNomons 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "frmTeleportOption.frx":0236
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   5
      Top             =   855
      Width           =   225
   End
   Begin VB.TextBox txtNoMons 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3180
      TabIndex        =   4
      Text            =   "15"
      Top             =   840
      Width           =   230
   End
   Begin VB.PictureBox imgDamageTele 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "frmTeleportOption.frx":0351
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   3
      Top             =   360
      Width           =   225
   End
   Begin VB.TextBox txtDamageTele 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1054
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2760
      MaxLength       =   4
      TabIndex        =   2
      Text            =   "200"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtHPtele 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1054
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2540
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "48"
      Top             =   615
      Width           =   255
   End
   Begin VB.PictureBox imgHpTele 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "frmTeleportOption.frx":046C
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   615
      Width           =   225
   End
   Begin VB.Label labWarpAll 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teleport away from all people."
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
      Left            =   405
      TabIndex        =   14
      Top             =   1080
      Width           =   2190
   End
   Begin VB.Label labPtele 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teleport away from Job"
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
      Left            =   405
      TabIndex        =   13
      Top             =   1320
      Width           =   1725
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Teleport Options"
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
      TabIndex        =   9
      Top             =   15
      Width           =   1215
   End
   Begin VB.Label labNomons 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto teleport when no monster every       sec(s)"
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
      Left            =   405
      TabIndex        =   8
      Top             =   855
      Width           =   3510
   End
   Begin VB.Label LabStopPick 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto teleport when damage over"
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
      Left            =   405
      TabIndex        =   7
      Top             =   360
      Width           =   2385
   End
   Begin VB.Label labBackTown 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto teleport when hp below       %"
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
      Left            =   405
      TabIndex        =   6
      Top             =   615
      Width           =   2580
   End
   Begin VB.Image Image5 
      Height          =   300
      Left            =   3480
      Picture         =   "frmTeleportOption.frx":0587
      Top             =   1740
      Width           =   630
   End
   Begin VB.Image Image4 
      Height          =   135
      Left            =   70
      Picture         =   "frmTeleportOption.frx":0839
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image6 
      Height          =   135
      Left            =   3960
      Picture         =   "frmTeleportOption.frx":096E
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "frmTeleportOption.frx":0AA3
      Top             =   0
      Width           =   4200
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   0
      Picture         =   "frmTeleportOption.frx":104C
      Top             =   1680
      Width           =   4200
   End
   Begin VB.Image Image2 
      Height          =   1800
      Left            =   0
      Picture         =   "frmTeleportOption.frx":1190
      Top             =   240
      Width           =   4200
   End
End
Attribute VB_Name = "frmTeleportOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Public Sub update_imgNomons()
    If NomonsWarp Then
        imgNomons.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgNomons.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
     txtNoMons.text = NomonsTime
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ReleaseCapture
    Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub Image5_Click()
    MDIfrmMain.Save_Option
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image5.Picture = LoadPicture(App.Path & "\interface\bt_change_c.gif")
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image5.Picture = LoadPicture(App.Path & "\interface\bt_change.gif")
End Sub

Private Sub Image6_Click()
    Unload frmTeleportOption
End Sub

Private Sub imgNomons_Click()
    NomonsWarp = Not NomonsWarp
    update_imgNomons
    MDIfrmMain.Save_Option
End Sub

'Public Sub update_ImgUseWing()
'    If AutoWing Then
'        ImgUseWing.Picture = LoadPicture(App.Path & "\interface\on.gif")
'    Else
'        ImgUseWing.Picture = LoadPicture(App.Path & "\interface\off.gif")
'    End If
'End Sub

'Private Sub ImgUseWing_Click()
'    AutoWing = Not AutoWing
'    update_ImgUseWing
'    MDIfrmMain.Save_Option
'End Sub

Public Sub update_imgDamageTele()
    If IsDamageDC Then
        imgDamageTele.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgDamageTele.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
    txtDamageTele.text = DamageSet
End Sub

Private Sub imgDamageTele_Click()
    IsDamageDC = Not IsDamageDC
    update_imgDamageTele
    MDIfrmMain.Save_Option
End Sub

Public Sub update_imgHpTele()
    If IsAutoDC Then
        imgHpTele.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgHpTele.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
    txtHPtele.text = HPDC * 100
End Sub

Private Sub imgHpTele_Click()
    IsAutoDC = Not IsAutoDC
    update_imgHpTele
    MDIfrmMain.Save_Option
End Sub

Private Sub imgWarpAll_Click()
    WarpAll = Not WarpAll
    update_imgWarpAll
    MDIfrmMain.Save_Option
End Sub

Public Sub update_imgWarpAll()
    If WarpAll Then
        imgWarpAll.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgWarpAll.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
End Sub

Private Sub imgJTele_Click()
    JTele = Not JTele
    update_imgJTele
    MDIfrmMain.Save_Option
End Sub
'SRJ - add 0.0.18
Public Sub update_imgJTele()
    If JTele Then
        imgJTele.Picture = LoadPicture(App.Path & "\interface\on.gif")
        JTele = True
    Else
        imgJTele.Picture = LoadPicture(App.Path & "\interface\off.gif")
        JTele = False
    End If
    txtJTele.text = JobTele
End Sub

Private Sub txtDamageTele_Change()
    DamageSet = Val(txtDamageTele.text)
End Sub

Private Sub txtHPtele_Change()
    HPDC = Val(txtHPtele.text) / 100
End Sub

Private Sub txtJTele_Change()
    JobTele = txtJTele
End Sub

Private Sub txtNoMons_Change()
    NomonsTime = Val(txtNoMons.text)
End Sub
