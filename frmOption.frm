VERSION 5.00
Begin VB.Form frmAIOption 
   BorderStyle     =   0  'None
   Caption         =   "Revemu Options"
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4020
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox imgExAll 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "frmOption.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   30
      Top             =   3330
      Width           =   225
   End
   Begin VB.PictureBox imgGSonyou 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "frmOption.frx":011B
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   29
      Top             =   2370
      Width           =   225
   End
   Begin VB.PictureBox imgGSnearyou 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "frmOption.frx":0236
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   28
      Top             =   2610
      Width           =   225
   End
   Begin VB.PictureBox imgMGSonyou 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "frmOption.frx":0351
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   27
      Top             =   2850
      Width           =   225
   End
   Begin VB.PictureBox imgMGSnearyou 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "frmOption.frx":046C
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   26
      Top             =   3090
      Width           =   225
   End
   Begin VB.PictureBox imgBackBuy 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "frmOption.frx":0587
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   20
      Top             =   2120
      Width           =   225
   End
   Begin VB.TextBox txtChatRoomName 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2040
      TabIndex        =   19
      Text            =   "<AFK>"
      Top             =   1850
      Width           =   1095
   End
   Begin VB.PictureBox imgAlwaySit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "frmOption.frx":06A2
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   17
      Top             =   1860
      Width           =   225
   End
   Begin VB.PictureBox imgBackTown 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "frmOption.frx":07BD
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   15
      Top             =   1100
      Width           =   225
   End
   Begin VB.TextBox txtBackTown 
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
      Left            =   2730
      MaxLength       =   2
      TabIndex        =   14
      Text            =   "48"
      Top             =   1100
      Width           =   255
   End
   Begin VB.TextBox txtStopAttack 
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
      Left            =   2730
      MaxLength       =   2
      TabIndex        =   13
      Text            =   "48"
      Top             =   1600
      Width           =   255
   End
   Begin VB.PictureBox imgStopAttack 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "frmOption.frx":08D8
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   11
      Top             =   1600
      Width           =   225
   End
   Begin VB.TextBox txtStopPick 
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
      Left            =   2730
      MaxLength       =   2
      TabIndex        =   10
      Text            =   "48"
      Top             =   1350
      Width           =   255
   End
   Begin VB.PictureBox imgStopPick 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "frmOption.frx":09F3
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   8
      Top             =   1350
      Width           =   225
   End
   Begin VB.TextBox txtWaypoint 
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
      Left            =   3240
      TabIndex        =   7
      Text            =   "15"
      Top             =   850
      Width           =   230
   End
   Begin VB.PictureBox imgWayPoint 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "frmOption.frx":0B0E
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   5
      Top             =   850
      Width           =   225
   End
   Begin VB.PictureBox imgMove 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "frmOption.frx":0C29
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   3
      Top             =   600
      Width           =   225
   End
   Begin VB.PictureBox ImgPick 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      Picture         =   "frmOption.frx":0D44
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   360
      Width           =   225
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto block whisper"
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
      TabIndex        =   31
      Top             =   3330
      Width           =   1410
   End
   Begin VB.Label labGSonyou 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Avoid player's skill at your position"
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
      TabIndex        =   25
      Top             =   2370
      Width           =   2490
   End
   Begin VB.Label labGSnearyou 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Avoid player's skill near your position"
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
      TabIndex        =   24
      Top             =   2610
      Width           =   2685
   End
   Begin VB.Label labMGSonyou 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Avoid monster's skill at your position"
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
      TabIndex        =   23
      Top             =   2850
      Width           =   2625
   End
   Begin VB.Label labMGSnearyou 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Avoid monster's skill near your position"
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
      TabIndex        =   22
      Top             =   3090
      Width           =   2820
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Back Town to Buy"
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
      TabIndex        =   21
      Top             =   2115
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alway Sit and make ""                           "" Chat Room"
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
      TabIndex        =   18
      Top             =   1860
      Width           =   3645
   End
   Begin VB.Image Image6 
      Height          =   135
      Left            =   3960
      Picture         =   "frmOption.frx":0E5F
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image4 
      Height          =   135
      Left            =   70
      Picture         =   "frmOption.frx":0F94
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image5 
      Height          =   300
      Left            =   3480
      Picture         =   "frmOption.frx":10C9
      ToolTipText     =   "Save all edited value."
      Top             =   3660
      Width           =   630
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   0
      Picture         =   "frmOption.frx":137B
      Top             =   3600
      Width           =   4200
   End
   Begin VB.Label labBackTown 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Back Town when weight reach       %"
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
      TabIndex        =   16
      Top             =   1095
      Width           =   2760
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stop attack when weight reach       %"
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
      TabIndex        =   12
      Top             =   1605
      Width           =   2745
   End
   Begin VB.Label LabStopPick 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stop pickup when weight reach       %"
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
      TabIndex        =   9
      Top             =   1350
      Width           =   2775
   End
   Begin VB.Label labWaypoint 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Use waypoint when no monster every       sec(s)"
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
      Top             =   855
      Width           =   3585
   End
   Begin VB.Label labMove 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Move"
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
      TabIndex        =   4
      Top             =   615
      Width           =   780
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "AI Options"
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
      TabIndex        =   2
      Top             =   15
      Width           =   1215
   End
   Begin VB.Label labPickup 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Pickup"
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
      TabIndex        =   1
      Top             =   375
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "frmOption.frx":14BF
      Top             =   0
      Width           =   4200
   End
   Begin VB.Image Image7 
      Height          =   3960
      Left            =   0
      Picture         =   "frmOption.frx":1A68
      Stretch         =   -1  'True
      Top             =   240
      Width           =   4200
   End
End
Attribute VB_Name = "frmAIOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Public Sub update_imgPick()
    If IsAutoPick Then
        ImgPick.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        ImgPick.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
End Sub

Public Sub update_imgMove()
    If Automove Then
        imgMove.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgMove.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
End Sub

Public Sub update_imgWayPoint()
    If RandomMove Then
        imgWayPoint.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgWayPoint.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
    txtWaypoint.text = Movetime
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
    Unload frmAIOption
End Sub

Public Sub update_imgBackTown()
    If IsBackTown Then
        imgBackTown.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgBackTown.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
    txtBackTown.text = WeightBackTown * 100
End Sub

Public Sub update_imgBackBuy()
    If isBackBuy Then
        imgBackBuy.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgBackBuy.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
End Sub

Public Sub update_imgAlwaySit()
    If AlwaySit Then
        imgAlwaySit.Picture = LoadPicture(App.Path & "\interface\on.gif")
        If ConnState > 3 Then
            frmMain.create_chatroom "test", ChatRoomName
            frmMain.Send_Sit
        End If
    Else
        imgAlwaySit.Picture = LoadPicture(App.Path & "\interface\off.gif")
        If ConnState > 3 Then frmMain.destroy_chatroom
    End If
        txtChatRoomName.text = ChatRoomName
End Sub

Private Sub imgAlwaySit_Click()
    AlwaySit = Not AlwaySit
    update_imgAlwaySit
End Sub

Private Sub imgBackBuy_Click()
    isBackBuy = Not isBackBuy
    update_imgBackBuy
End Sub

Private Sub imgBackTown_Click()
    IsBackTown = Not IsBackTown
    update_imgBackTown
End Sub

Private Sub imgExAll_Click()
    ExAll = Not ExAll
    update_imgExAll
    MDIfrmMain.Save_Option
End Sub

Private Sub imgGSonyou_Click()
    GSonyou = Not GSonyou
    update_imgGSonyou
    MDIfrmMain.Save_Option
End Sub
Public Sub update_imgGSonyou()
    If GSonyou Then
        imgGSonyou.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgGSonyou.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
End Sub
Private Sub imgMGSonyou_Click()
    MGSonyou = Not MGSonyou
    update_imgMGSonyou
    MDIfrmMain.Save_Option
End Sub
Private Sub imgGSnearyou_Click()
    GSnearyou = Not GSnearyou
    update_imgGSnearyou
    MDIfrmMain.Save_Option
End Sub
Public Sub update_imgGSnearyou()
    If GSnearyou Then
        imgGSnearyou.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgGSnearyou.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
End Sub
Public Sub update_imgMGSonyou()
    If MGSonyou Then
        imgMGSonyou.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgMGSonyou.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
End Sub
Private Sub imgMGSnearyou_Click()
    MGSnearyou = Not MGSnearyou
    update_imgMGSnearyou
    MDIfrmMain.Save_Option
End Sub
Public Sub update_imgMGSnearyou()
    If MGSnearyou Then
        imgMGSnearyou.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgMGSnearyou.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
End Sub

Public Sub update_imgExAll()
    If ExAll Then
        imgExAll.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgExAll.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
End Sub


Private Sub imgMove_Click()
    Automove = Not Automove
    update_imgMove
    MDIfrmMain.Save_Option
End Sub

Private Sub ImgPick_Click()
    IsAutoPick = Not IsAutoPick
    update_imgPick
    MDIfrmMain.Save_Option
End Sub

Public Sub update_imgStopPick()
    If SWeight2 Then
        imgStopPick.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgStopPick.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
    txtStopPick.text = Weight2 * 100
End Sub

Private Sub imgStopPick_Click()
    SWeight2 = Not SWeight2
    update_imgStopPick
    MDIfrmMain.Save_Option
End Sub

Public Sub update_imgStopAttack()
    If SWeight1 Then
        imgStopAttack.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgStopAttack.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
    txtStopAttack.text = Weight1 * 100
End Sub

Private Sub imgStopAttack_Click()
    SWeight1 = Not SWeight1
    update_imgStopAttack
    MDIfrmMain.Save_Option
End Sub

Private Sub imgWayPoint_Click()
    RandomMove = Not RandomMove
    update_imgWayPoint
    MDIfrmMain.Save_Option
End Sub

Private Sub txtBackTown_Change()
    WeightBackTown = val(txtBackTown.text) / 100
End Sub

Private Sub txtStopAttack_Change()
    Weight1 = val(txtStopAttack.text) / 100
End Sub

Private Sub txtStopPick_Change()
    Weight2 = val(txtStopPick.text) / 100
End Sub

Private Sub txtWaypoint_Change()
    Movetime = val(txtWaypoint.text)
End Sub
