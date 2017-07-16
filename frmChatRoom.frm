VERSION 5.00
Begin VB.Form frmChatRoom 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4455
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstChatroom 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3150
      IntegralHeight  =   0   'False
      ItemData        =   "frmChatRoom.frx":0000
      Left            =   0
      List            =   "frmChatRoom.frx":0002
      TabIndex        =   0
      Top             =   240
      Width           =   4200
   End
   Begin VB.Image imgReSize 
      Height          =   180
      Left            =   2040
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmChatRoom.frx":0004
      Top             =   3480
      Width           =   180
   End
   Begin VB.Image imgclose 
      Height          =   135
      Left            =   3980
      Picture         =   "frmChatRoom.frx":0150
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chat room/shop list"
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
      Width           =   1380
   End
   Begin VB.Image imgbmid 
      Height          =   420
      Left            =   120
      Picture         =   "frmChatRoom.frx":0285
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgbright 
      Height          =   420
      Left            =   1560
      Picture         =   "frmChatRoom.frx":02DD
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgbleft 
      Height          =   420
      Left            =   0
      Picture         =   "frmChatRoom.frx":0352
      Top             =   3480
      Width           =   150
   End
   Begin VB.Image imgMidbar 
      Height          =   255
      Left            =   170
      Picture         =   "frmChatRoom.frx":03D1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   180
   End
   Begin VB.Image Image5 
      Height          =   135
      Left            =   50
      Picture         =   "frmChatRoom.frx":04A9
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "frmChatRoom.frx":05DE
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgRightbar 
      Height          =   255
      Left            =   1560
      Picture         =   "frmChatRoom.frx":074C
      Top             =   0
      Width           =   180
   End
End
Attribute VB_Name = "frmChatRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.height = 4500
Me.width = 4300
imgRightbar.Left = Me.width - 200
imgMidbar.width = Me.width - 350
LoadFormPos Me
lstChatroom.height = Me.height - 880
lstChatroom.width = Me.width
imgbleft.Top = lstChatroom.height + 200
imgbmid.Top = lstChatroom.height + 200
imgbright.Top = lstChatroom.height + 200
imgbright.Left = Me.width - 300
imgbmid.width = Me.width - 400
imgReSize.Top = lstChatroom.height + 320
imgReSize.Left = Me.width - 270
Me.height = lstChatroom.height + 880
'If UBound(NPCList) > 0 Then
'    frmMain.UpdateNPC
'End If
End Sub
Private Sub Form_Resize()
If (Me.width < 2000 Or Me.height < 2000) Then
Form_Load
Else
imgRightbar.Left = Me.width - 180
imgMidbar.width = Me.width - 320
lstChatroom.height = Me.height - 650
lstChatroom.width = Me.width
imgbleft.Top = lstChatroom.height + 240
imgbmid.Top = lstChatroom.height + 240
imgclose.Left = Me.width - 200
imgbright.Top = lstChatroom.height + 240
imgbright.Left = Me.width - 100
imgbmid.width = Me.width - 200
imgReSize.Top = lstChatroom.height + 480
imgReSize.Left = Me.width - 182
If (Me.height + 650) < MDIfrmMain.height Then Me.height = lstChatroom.height + 650
End If
If frmMain.Visible Then
    If UBound(MChat) Or UBound(MShop) Then frmMain.UpdateChatShop
End If
End Sub

Private Sub imgclose_Click()
    SaveFormPos Me
    Unload Me
End Sub

Private Sub imgMidbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
SaveFormPos Me
End Sub

Private Sub imgReSize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ReleaseCapture
    Call SendMessage(hWnd, WM_NCLBUTTONDOWN, 17, 0)
    SaveFormPos Me
End Sub

Private Sub lstChatroom_DblClick()
'tmp
With lstChatroom
    If .ListIndex < 0 Then Exit Sub
    Dim CInfo$
    CInfo = .List(.ListIndex)
    If Left$(CInfo, 6) = "[shop]" Then
        CInfo = Left(Right(CInfo, 9), 8)
    End If
End With
End Sub
