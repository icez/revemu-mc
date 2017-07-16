VERSION 5.00
Begin VB.Form frmCart 
   BorderStyle     =   0  'None
   Caption         =   "Cart"
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstCart 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   0
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   240
      Width           =   4200
   End
   Begin VB.Image imgReSize 
      Height          =   180
      Left            =   2040
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmCart.frx":0000
      Top             =   3480
      Width           =   180
   End
   Begin VB.Image Image5 
      Height          =   135
      Left            =   50
      Picture         =   "frmCart.frx":014C
      Top             =   60
      Width           =   135
   End
   Begin VB.Image imgclose 
      Height          =   135
      Left            =   3980
      Picture         =   "frmCart.frx":0281
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cart"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   15
      Width           =   315
   End
   Begin VB.Label lblCart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   45
   End
   Begin VB.Image imgbmid 
      Height          =   420
      Left            =   120
      Picture         =   "frmCart.frx":03B6
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgbright 
      Height          =   420
      Left            =   1560
      Picture         =   "frmCart.frx":040E
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgbleft 
      Height          =   420
      Left            =   0
      Picture         =   "frmCart.frx":0483
      Top             =   3480
      Width           =   150
   End
   Begin VB.Image imgMidbar 
      Height          =   255
      Left            =   170
      Picture         =   "frmCart.frx":0502
      Stretch         =   -1  'True
      Top             =   0
      Width           =   180
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "frmCart.frx":05DA
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgRightbar 
      Height          =   255
      Left            =   1560
      Picture         =   "frmCart.frx":0748
      Top             =   0
      Width           =   180
   End
End
Attribute VB_Name = "frmCart"
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
lstCart.height = Me.height - 880
lstCart.width = Me.width
imgbleft.Top = lstCart.height + 200
imgbmid.Top = lstCart.height + 200
imgbright.Top = lstCart.height + 200
imgbright.Left = Me.width - 300
imgbmid.width = Me.width - 400
lblCart.Top = imgbmid.Top + 120
'120
imgReSize.Top = lstCart.height + 320
imgReSize.Left = Me.width - 270
Me.height = lstCart.height + 880
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
lstCart.height = Me.height - 650
lstCart.width = Me.width
imgbleft.Top = lstCart.height + 240
imgbmid.Top = lstCart.height + 240
imgclose.Left = Me.width - 200
imgbright.Top = lstCart.height + 240
lblCart.Top = imgbmid.Top + 120
imgbright.Left = Me.width - 100
imgbmid.width = Me.width - 200
imgReSize.Top = lstCart.height + 480
imgReSize.Left = Me.width - 182
If (Me.height + 650) < MDIfrmMain.height Then Me.height = lstCart.height + 650
End If
End Sub

Private Sub imgclose_Click()
'    frmMain.Send_ShopClose
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

Private Sub lstCart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lstCart.OLEDrag
End Sub

Private Sub lstCart_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strList As String
    Dim Index As Integer
    Dim Amount As Long
    Amount = 0
    ' Check the format of the DataObject
    If Not data.GetFormat(vbCFText) Then Exit Sub
    ' Retrieve the text from the DataObject
    strList = data.GetData(vbCFText)
    If strList = "" Then Exit Sub
    If strList = frmStorage.lstStorage.List(frmStorage.lstStorage.ListIndex) Then
        Index = Val(frmStorage.lstStorage.List(frmStorage.lstStorage.ListIndex))
        If Storage(Index).Amount > 1 Then
            Amount = Val(InputBox("Enter number to get", "Get from storage", CStr(Cart(Index).Amount)))
        Else
            Amount = 1
        End If
        If Amount > 0 Then pkt_CartFromKafra Index, Amount
    End If
    If strList = frmItem.lstInvent.List(frmItem.lstInvent.ListIndex) Then
        Index = Val(frmItem.lstInvent.List(frmItem.lstInvent.ListIndex))
        If AllInv(Index).Amount > 1 Then
            Amount = Val(InputBox("Enter number to add", "Get from inventory", CStr(AllInv(Index).Amount)))
        Else
            Amount = 1
        End If
        If Amount > 0 Then pkt_CartGet Index, Amount
    End If
    ' If the item was not dropped on itself
    strList = ""
End Sub

Private Sub lstCart_OLEStartDrag(data As DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectMove
    data.SetData lstCart
End Sub
