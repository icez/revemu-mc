VERSION 5.00
Begin VB.Form frmModConfig 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOptWin 
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   7
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdOptWin 
      Caption         =   "Buying"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdOptWin 
      Caption         =   "Vending"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdOptWin 
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.Frame famOption 
      BackColor       =   &H80000009&
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3975
      Begin VB.CheckBox chkAutoSit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Auto-Sit when create shop/chatroom"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox txtCDelay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H000000FF&
         Height          =   235
         Left            =   2880
         TabIndex        =   13
         Text            =   "5"
         Top             =   1080
         Width           =   390
      End
      Begin VB.TextBox txtVDelay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H000000FF&
         Height          =   235
         Left            =   2520
         TabIndex        =   12
         Text            =   "5"
         Top             =   600
         Width           =   390
      End
      Begin VB.CheckBox chkChatroom 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Auto-Chatroom create with delay            sec"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   3735
      End
      Begin VB.CheckBox chkVending 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Auto-Shop create with delay            sec"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   3735
      End
      Begin VB.CheckBox chkEnabled 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Enabled?"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   3735
      End
      Begin VB.CheckBox chkDCShop 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Disconnect when shop close?"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   3735
      End
      Begin VB.CheckBox chkNewMapType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Use new map type?"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   3735
      End
   End
   Begin VB.Frame famOption 
      BackColor       =   &H80000009&
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   3975
      Begin VB.CheckBox chkSTSys 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "System message?"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1560
         Width           =   3735
      End
      Begin VB.CheckBox chkSTWalk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Walking message?"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1320
         Width           =   3735
      End
      Begin VB.CheckBox chkSTStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Status change text?"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   3735
      End
      Begin VB.CheckBox chkSTChat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Chat room/shop  [appear/disappear] text?"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   840
         Width           =   3735
      End
      Begin VB.CheckBox chkEmotion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Emotion?"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   3735
      End
      Begin VB.CheckBox chkGuildAnn 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Guild Announce?"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame famOption 
      BackColor       =   &H80000009&
      Caption         =   "Buying"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   3975
      Begin VB.CheckBox chkNoCalc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "No calculate money (accept all item with 0z)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox txtCTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H000000FF&
         Height          =   235
         Left            =   1200
         TabIndex        =   33
         Text            =   "10000"
         Top             =   1080
         Width           =   2685
      End
      Begin VB.CheckBox chkDCChat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Auto-Disconnect when chatroom close?"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox txtBelow 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H000000FF&
         Height          =   235
         Left            =   2640
         TabIndex        =   20
         Text            =   "10000"
         Top             =   360
         Width           =   1005
      End
      Begin VB.CheckBox chkCreateShop 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Create shop when zeny below                          z"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Chatroom Title"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.Frame famOption 
      BackColor       =   &H80000009&
      Caption         =   "Vending"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   3975
      Begin VB.TextBox txtITNPC 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H000000FF&
         Height          =   235
         Index           =   0
         Left            =   1680
         TabIndex        =   31
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtITAmount 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H000000FF&
         Height          =   235
         Index           =   0
         Left            =   1680
         TabIndex        =   29
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtITPrice 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H000000FF&
         Height          =   235
         Index           =   0
         Left            =   1680
         TabIndex        =   28
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtITName 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H000000FF&
         Height          =   235
         Index           =   0
         Left            =   1680
         TabIndex        =   27
         Top             =   360
         Width           =   2175
      End
      Begin VB.ListBox lstVL 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NPC : "
         Height          =   255
         Left            =   720
         TabIndex        =   30
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "<< Click at this list to select an item to edit"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   720
         TabIndex        =   26
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Amount : "
         Height          =   255
         Left            =   720
         TabIndex        =   25
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Price : "
         Height          =   255
         Left            =   720
         TabIndex        =   24
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name : "
         Height          =   255
         Left            =   720
         TabIndex        =   23
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Image Image5 
      Height          =   300
      Left            =   3480
      Picture         =   "frmModConfig.frx":0000
      Top             =   2820
      Width           =   630
   End
   Begin VB.Image Image6 
      Height          =   135
      Left            =   3960
      Picture         =   "frmModConfig.frx":02B2
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image4 
      Height          =   135
      Left            =   70
      Picture         =   "frmModConfig.frx":03E7
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mods Options"
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
      TabIndex        =   0
      Top             =   15
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "frmModConfig.frx":051C
      Top             =   0
      Width           =   4200
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   0
      Picture         =   "frmModConfig.frx":0AC5
      Top             =   2760
      Width           =   4200
   End
   Begin VB.Image Image2 
      Height          =   3000
      Left            =   0
      Picture         =   "frmModConfig.frx":0C09
      Top             =   240
      Width           =   45
   End
   Begin VB.Image Image7 
      Height          =   3000
      Left            =   4080
      Picture         =   "frmModConfig.frx":15AB
      Top             =   240
      Width           =   120
   End
End
Attribute VB_Name = "frmModConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAutoSit_Click()
    Mods.AutoSit = CBool(chkAutoSit.value)
End Sub

Private Sub chkChatroom_Click()
    Mods.OC = CBool(chkChatroom.value)
End Sub

Private Sub chkCreateShop_Click()
    Mods.OCcreateshop = CBool(chkCreateShop.value)
End Sub

Private Sub chkDCChat_Click()
    Mods.OCdisconnect = CBool(chkDCChat.value)
End Sub

Private Sub chkDCShop_Click()
    Mods.dcshop = CBool(chkDCShop.value)
End Sub

Private Sub chkEmotion_Click()
    Mods.EmotionText = CBool(chkEmotion.value)
End Sub

Private Sub chkEnabled_Click()
    Mods.Enabled = CBool(chkEnabled.value)
End Sub

Private Sub chkGuildAnn_Click()
    Mods.GuildText = CBool(chkGuildAnn.value)
End Sub

Private Sub chkNoCalc_Click()
    Mods.OCnocalcmoney = CBool(chkNoCalc.value)
End Sub

Private Sub chkSTChat_Click()
    Mods.STChat = CBool(chkSTChat.value)
End Sub

Private Sub chkSTStatus_Click()
    Mods.STStatus = CBool(chkSTStatus.value)
End Sub

Private Sub chkSTSys_Click()
    Mods.STSystem = CBool(chkSTSys.value)
End Sub

Private Sub chkSTWalk_Click()
    Mods.STWalk = CBool(chkSTWalk.value)
End Sub

Private Sub chkVending_Click()
    Mods.Vending = CBool(chkVending.value)
End Sub

Private Sub cmdOptWin_Click(index As Integer)
    Dim i As Long
    For i = 0 To 3
        famOption(i).Visible = False
    Next
    famOption(index).Visible = True
End Sub

Private Sub Form_Load()
    'tmp
    cmdOptWin_Click 0
    Dim i As Long
    For i = 1 To 29
        Load txtITName(i)
        Load txtITPrice(i)
        Load txtITAmount(i)
        Load txtITNPC(i)
    Next
    lstVL.Clear
    For i = 1 To 30
        lstVL.AddItem CStr(i)
    Next
    lstVL.ListIndex = 0
    lstVL_Click
    RefreshMC
    LoadFormPos frmModConfig
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ReleaseCapture
    Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub Image5_Click()
    SaveModConfig
    If IsShopCreated Then frmMain.Send_ShopClose
    If IsChatOC Then frmMain.destroy_chatroom
    CalcModAI "modconfig"
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image5.Picture = LoadPicture(App.Path & "\interface\bt_change_c.gif")
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image5.Picture = LoadPicture(App.Path & "\interface\bt_change.gif")
End Sub

Private Sub Image6_Click()
    Unload frmModConfig
End Sub

Private Sub lstVL_Click()
    Dim i&
    For i = 0 To txtITAmount.UBound
        txtITAmount(i).Visible = False
        txtITName(i).Visible = False
        txtITNPC(i).Visible = False
        txtITPrice(i).Visible = False
    Next
    i = lstVL.ListIndex
    txtITAmount(i).Visible = True
    txtITName(i).Visible = True
    txtITNPC(i).Visible = True
    txtITPrice(i).Visible = True
End Sub

Private Sub txtBelow_Change()
If Not IsNumeric(txtBelow.text) Then txtBelow.text = "0"
Mods.minzeny = Val(txtBelow.text)
End Sub

Private Sub txtBelow_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtCDelay_Change()
If Not IsNumeric(txtCDelay.text) Then txtCDelay.text = "0"
Mods.OCdelay = Val(txtCDelay.text)
End Sub

Private Sub txtCDelay_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtCTitle_Change()
    Mods.Chatroom = txtCTitle.text
End Sub

Private Sub txtITAmount_Change(index As Integer)
If Not IsNumeric(txtITAmount(index).text) Then txtITAmount(index).text = "0"
Vending(index + 1).Amount = Val(txtITAmount(index).text)
End Sub

Private Sub txtITAmount_KeyPress(index As Integer, KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtITName_Change(index As Integer)
    Vending(index + 1).Name = txtITName(index).text
End Sub

Private Sub txtITNPC_Change(index As Integer)
    Vending(index + 1).NPC = txtITNPC(index).text
End Sub

Private Sub txtITPrice_Change(index As Integer)
If Not IsNumeric(txtITPrice(index).text) Then txtITPrice(index).text = "0"
Vending(index + 1).Price = Val(txtITPrice(index).text)
End Sub

Private Sub txtITPrice_KeyPress(index As Integer, KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtVDelay_Change()
If Not IsNumeric(txtVDelay.text) Then txtVDelay.text = "0"
Mods.Vendingdelay = Val(txtVDelay.text)
End Sub

Private Sub txtVDelay_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub
