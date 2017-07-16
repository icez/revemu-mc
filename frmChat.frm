VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmChat 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Chat"
   ClientHeight    =   5355
   ClientLeft      =   2415
   ClientTop       =   7860
   ClientWidth     =   4920
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   3375
      Left            =   0
      TabIndex        =   7
      Top             =   240
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   5953
      _Version        =   393217
      BackColor       =   16777215
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmChat.frx":0E42
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.OptionButton optParty 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Party"
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
      Left            =   3720
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton optWhisper 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Private"
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
      Left            =   2040
      TabIndex        =   4
      Top             =   4440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OptionButton optPublic 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Public"
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
      Left            =   2880
      TabIndex        =   3
      Top             =   4440
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtWhisper 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   120
      MaxLength       =   24
      TabIndex        =   0
      Top             =   3690
      Width           =   1185
   End
   Begin VB.TextBox txtSay 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3600
      Width           =   3075
   End
   Begin VB.Image ImgDialog 
      Height          =   270
      Left            =   4560
      Picture         =   "frmChat.frx":0EBE
      Top             =   4440
      Width           =   120
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chat"
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
      TabIndex        =   6
      Top             =   15
      Width           =   330
   End
   Begin VB.Image imgClose 
      Height          =   135
      Left            =   1920
      Picture         =   "frmChat.frx":10B0
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   50
      Picture         =   "frmChat.frx":11E5
      Top             =   60
      Width           =   135
   End
   Begin VB.Image imgResize 
      Height          =   180
      Left            =   4680
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmChat.frx":131A
      Top             =   3720
      Width           =   180
   End
   Begin VB.Image imgbright 
      Height          =   360
      Left            =   4440
      Picture         =   "frmChat.frx":1466
      Top             =   3600
      Width           =   510
   End
   Begin VB.Image imgbmid 
      Height          =   360
      Left            =   1680
      Picture         =   "frmChat.frx":1527
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   270
   End
   Begin VB.Image imgRightBar 
      Height          =   255
      Left            =   1200
      Picture         =   "frmChat.frx":1597
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgMidBar 
      Height          =   255
      Left            =   120
      Picture         =   "frmChat.frx":1801
      Stretch         =   -1  'True
      Top             =   0
      Width           =   180
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "frmChat.frx":18D9
      Top             =   0
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   3840
      TabIndex        =   2
      Top             =   4005
      Width           =   45
   End
   Begin VB.Image imgbLeft 
      Height          =   360
      Left            =   0
      Picture         =   "frmChat.frx":1A47
      Top             =   3600
      Width           =   1680
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
frmChat.height = 2500
frmChat.width = 2500
LoadFormPos frmChat
imgRightbar.Left = frmChat.width - 180
imgMidbar.width = frmChat.width - 300
txtChat.height = frmChat.height - 600
txtChat.width = frmChat.width
imgbleft.Top = txtChat.height + 240
imgbmid.Top = txtChat.height + 240
imgbmid.width = frmChat.width - 200
imgbright.Top = txtChat.height + 240
imgbright.Left = frmChat.width - 480
imgReSize.Top = txtChat.height + 400
imgReSize.Left = frmChat.width - 200
txtWhisper.Top = txtChat.height + 350
txtSay.Top = txtChat.height + 350
optWhisper.Top = txtChat.height - 100
optPublic.Top = txtChat.height - 100
optParty.Top = txtChat.height - 100
optWhisper.Left = txtChat.width - 2500
optPublic.Left = txtChat.width - 1600
optParty.Left = txtChat.width - 800
txtSay.width = txtChat.width - 2200
ImgDialog.Left = frmChat.width - 400
ImgDialog.Top = txtChat.height + 280
End Sub

Private Sub Form_Resize()
If (frmChat.width < 2200 Or frmChat.height < 2200) Then
Form_Load
Else
imgRightbar.Left = frmChat.width - 180
imgMidbar.width = frmChat.width - 300
txtChat.height = frmChat.height - 600
txtChat.width = frmChat.width
imgbleft.Top = txtChat.height + 240
imgbmid.Top = txtChat.height + 240
imgbmid.width = frmChat.width - 200
imgbright.Top = txtChat.height + 240
imgbright.Left = frmChat.width - 480
imgclose.Left = frmChat.width - 200
imgReSize.Top = txtChat.height + 400
imgReSize.Left = frmChat.width - 200
txtWhisper.Top = txtChat.height + 350
txtSay.Top = txtChat.height + 350
optWhisper.Top = txtChat.height - 100
optPublic.Top = txtChat.height - 100
optParty.Top = txtChat.height - 100
optWhisper.Left = txtChat.width - 2500
optPublic.Left = txtChat.width - 1600
optParty.Left = txtChat.width - 800
txtSay.width = txtChat.width - 2200
ImgDialog.Left = frmChat.width - 400
ImgDialog.Top = txtChat.height + 280
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIfrmMain.mnuChat.CheckED = False
SaveFormPos frmChat
End Sub

Private Sub imgclose_Click()
frmChat.Visible = False
End Sub

Private Sub ImgDialog_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.PopupMenu frmPopupChat.mnuChat
End Sub

Private Sub imgMidbar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
SaveFormPos frmChat
End Sub

Private Sub imgReSize_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    Call ReleaseCapture
    Call SendMessage(frmChat.hwnd, WM_NCLBUTTONDOWN, 17, 0)
    SaveFormPos frmChat
End If
End Sub

Private Sub txtSay_Change()
If txtSay.text <> "" Then
    If CStr(Asc(Right(txtSay.text, 1))) = "10" Then
        frmMain.Chat_Send
        txtSay.text = ""
    Else
        Label1.Caption = txtSay.text
    End If
End If
End Sub
