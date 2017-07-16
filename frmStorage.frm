VERSION 5.00
Begin VB.Form frmStorage 
   BorderStyle     =   0  'None
   Caption         =   "Storage List"
   ClientHeight    =   4980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstStorage 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   3150
      IntegralHeight  =   0   'False
      ItemData        =   "frmStorage.frx":0000
      Left            =   0
      List            =   "frmStorage.frx":0002
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   240
      Width           =   4200
   End
   Begin VB.Image imgclose 
      Height          =   135
      Left            =   3980
      Picture         =   "frmStorage.frx":0004
      Top             =   60
      Width           =   135
   End
   Begin VB.Image btClose 
      Height          =   300
      Left            =   3480
      Picture         =   "frmStorage.frx":0139
      Top             =   3480
      Width           =   630
   End
   Begin VB.Label labNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   45
   End
   Begin VB.Image imgReSize 
      Height          =   180
      Left            =   2040
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmStorage.frx":04FB
      Top             =   3480
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Storage List"
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
      Width           =   870
   End
   Begin VB.Image Image5 
      Height          =   135
      Left            =   50
      Picture         =   "frmStorage.frx":0647
      Top             =   60
      Width           =   135
   End
   Begin VB.Image imgMidbar 
      Height          =   255
      Left            =   170
      Picture         =   "frmStorage.frx":077C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   180
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "frmStorage.frx":0854
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgRightbar 
      Height          =   255
      Left            =   1560
      Picture         =   "frmStorage.frx":09C2
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgbmid 
      Height          =   420
      Left            =   120
      Picture         =   "frmStorage.frx":0C2C
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgbright 
      Height          =   420
      Left            =   1560
      Picture         =   "frmStorage.frx":0C84
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgbleft 
      Height          =   420
      Left            =   0
      Picture         =   "frmStorage.frx":0CF9
      Top             =   3480
      Width           =   150
   End
End
Attribute VB_Name = "frmStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btClose_Click()
    pkt_StorageClose
End Sub

Private Sub Form_Load()
Me.height = 4500
Me.width = 4300
imgRightbar.Left = Me.width - 200
imgMidbar.width = Me.width - 350
LoadFormPos Me
lstStorage.height = Me.height - 880
lstStorage.width = Me.width
imgbleft.Top = lstStorage.height + 200
imgbmid.Top = lstStorage.height + 200
imgbright.Top = lstStorage.height + 200
imgbright.Left = Me.width - 300
imgbmid.width = Me.width - 400
imgReSize.Top = lstStorage.height + 320
imgReSize.Left = Me.width - 270
Me.height = lstStorage.height + 880
End Sub

Private Sub Form_Resize()
If (Me.width < 2000 Or Me.height < 2000) Then
Form_Load
Else
imgRightbar.Left = Me.width - 180
imgMidbar.width = Me.width - 320
lstStorage.height = Me.height - 650
lstStorage.width = Me.width
imgbleft.Top = lstStorage.height + 240
imgbmid.Top = lstStorage.height + 240
imgclose.Left = Me.width - 200
imgbright.Top = lstStorage.height + 240
imgbright.Left = Me.width - 100
imgbmid.width = Me.width - 200
imgReSize.Top = lstStorage.height + 480
labNumber.Top = lstStorage.height + 400
imgReSize.Left = Me.width - 182
btClose.Top = lstStorage.height + 280
btClose.Left = Me.width - 800
If (Me.height + 650) < MDIfrmMain.height Then Me.height = lstStorage.height + 650
End If
End Sub

Private Sub imgclose_Click()
    pkt_StorageClose
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

Private Sub lstStorage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lstStorage.OLEDrag
End Sub

Private Sub lstStorage_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strList As String
    Dim Index As Integer
    Dim Amount As Long
    Amount = 0
    ' Check the format of the DataObject
    If Not data.GetFormat(vbCFText) Then Exit Sub
    ' Retrieve the text from the DataObject
    strList = data.GetData(vbCFText)
    If strList = "" Then Exit Sub
    If strList = frmItem.lstInvent.List(frmItem.lstInvent.ListIndex) Then
        Index = Val(frmItem.lstInvent.List(frmItem.lstInvent.ListIndex))
        If AllInv(Index).Amount > 1 Then
            Amount = Val(InputBox("Enter number to add", "Add to storage", CStr(AllInv(Index).Amount)))
        Else
            Amount = 1
        End If
        If Amount > 0 Then pkt_StorageAdd Index, Amount
    End If
    If strList = frmCart.lstCart.List(frmCart.lstCart.ListIndex) Then
        Index = Val(frmCart.lstCart.List(frmCart.lstCart.ListIndex))
        If Cart(Index).Amount > 1 Then
            Amount = Val(InputBox("Enter number to add", "Add from cart to storage", CStr(Cart(Index).Amount)))
        Else
            Amount = 1
        End If
        If Amount > 0 Then pkt_CartToKafra Index, Amount
    End If
    strList = ""
End Sub

Private Sub lstStorage_OLEStartDrag(data As DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectMove
    data.SetData lstStorage
End Sub
