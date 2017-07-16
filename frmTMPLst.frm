VERSION 5.00
Begin VB.Form frmTMPLst 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4200
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstStorage 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   3150
      IntegralHeight  =   0   'False
      ItemData        =   "frmTMPLst.frx":0000
      Left            =   0
      List            =   "frmTMPLst.frx":0002
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   240
      Width           =   4200
   End
   Begin VB.Image imgBuy 
      Height          =   300
      Left            =   3480
      Picture         =   "frmTMPLst.frx":0004
      Top             =   3480
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image imgSell 
      Height          =   300
      Left            =   2880
      Picture         =   "frmTMPLst.frx":028D
      Top             =   3480
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image imgReSize 
      Height          =   180
      Left            =   2040
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmTMPLst.frx":0646
      Top             =   3480
      Width           =   180
   End
   Begin VB.Image imgclose 
      Height          =   135
      Left            =   3980
      Picture         =   "frmTMPLst.frx":0792
      Top             =   60
      Width           =   135
   End
   Begin VB.Label labStore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Drag Item to buy/sell here"
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
      TabIndex        =   2
      Top             =   15
      Width           =   1830
   End
   Begin VB.Image Image5 
      Height          =   135
      Left            =   50
      Picture         =   "frmTMPLst.frx":08C7
      Top             =   60
      Width           =   135
   End
   Begin VB.Image imgMidbar 
      Height          =   255
      Left            =   170
      Picture         =   "frmTMPLst.frx":09FC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   180
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "frmTMPLst.frx":0AD4
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgRightbar 
      Height          =   255
      Left            =   1560
      Picture         =   "frmTMPLst.frx":0C42
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgbleft 
      Height          =   420
      Left            =   0
      Picture         =   "frmTMPLst.frx":0EAC
      Top             =   3480
      Width           =   150
   End
   Begin VB.Image imgbright 
      Height          =   420
      Left            =   1560
      Picture         =   "frmTMPLst.frx":0F2B
      Top             =   3480
      Width           =   120
   End
   Begin VB.Image imgbmid 
      Height          =   420
      Left            =   120
      Picture         =   "frmTMPLst.frx":0FA0
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   120
   End
   Begin VB.Label labNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   45
   End
End
Attribute VB_Name = "frmTMPLst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private tmpStore() As ItemInv

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
imgResize.Top = lstStorage.height + 320
imgResize.Left = Me.width - 270
Me.height = lstStorage.height + 880
ReDim tmpStore(0)
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
imgResize.Top = lstStorage.height + 480
labNumber.Top = lstStorage.height + 400
imgResize.Left = Me.width - 182
imgSell.Top = lstStorage.height + 280
imgSell.Left = Me.width - 800
imgBuy.Top = lstStorage.height + 280
imgBuy.Left = Me.width - 800
If (Me.height + 650) < MDIfrmMain.height Then Me.height = lstStorage.height + 650
End If
End Sub

Private Sub imgBuy_Click()
    Dim i As Integer
    Dim tstr As String
    For i = 0 To UBound(tmpStore) - 1
        tstr = tstr & IntToChr(tmpStore(i).Amount) _
        & IntToChr(CLng(Val("&H" & tmpStore(i).NameID)))
    Next
    frmMain.Send_BuyList tstr
    Unload Me
    Unload frmStoreBuy
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

Private Sub Update_TmpStore()
    Dim i As Integer
    lstStorage.Clear
    For i = 0 To UBound(tmpStore) - 1
        lstStorage.AddItem CStr(i) & " : " & tmpStore(i).Name & " " & _
        CStr(tmpStore(i).Amount) & " EA"
    Next
End Sub

Private Sub imgSell_Click()
    Dim i As Integer
    Dim tstr As String
    For i = 0 To UBound(tmpStore) - 1
        tstr = tstr & IntToChr(CLng(tmpStore(i).Index)) _
        & IntToChr(CLng(Val("&H" & tmpStore(i).Amount)))
    Next
    frmMain.Send_SellList tstr
    Unload Me
    Unload frmStoreBuy
End Sub

Private Sub lstStorage_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strList As String
    Dim Index As Integer
    Dim Amount As Long
    Dim i As Integer
    Dim found As Boolean
    Amount = 0
    ' Check the format of the DataObject
    If Not data.GetFormat(vbCFText) Then Exit Sub
    ' Retrieve the text from the DataObject
    strList = data.GetData(vbCFText)
    If strList = "" Then Exit Sub
    If strList = frmStoreBuy.lstItem.List(frmStoreBuy.lstItem.ListIndex) Then
        Index = Val(frmStoreBuy.lstItem.List(frmStoreBuy.lstItem.ListIndex))
        Amount = Val(InputBox("Enter number to buy", "Buy Item", CStr(1)))
        If UBound(tmpStore) > 0 Then
            found = False
            For i = 0 To UBound(tmpStore) - 1
                If (Store(Index).NameID = tmpStore(i).NameID) Then
                    tmpStore(i).Amount = tmpStore(i).Amount + Amount
                    found = True
                    Exit For
                End If
            Next
            If Not found Then
                tmpStore(UBound(tmpStore)) = Store(Index)
                tmpStore(UBound(tmpStore)).Amount = Amount
                ReDim Preserve tmpStore(UBound(tmpStore) + 1)
            End If
        Else
            tmpStore(UBound(tmpStore)) = Store(Index)
            tmpStore(UBound(tmpStore)).Amount = Amount
            ReDim Preserve tmpStore(UBound(tmpStore) + 1)
        End If
        'If amount > 0 Then frmMain.pkt_StorageGet Storage(index).index, amount
    End If
    ' If the item was not dropped on itself
    Update_TmpStore
    strList = ""
End Sub
