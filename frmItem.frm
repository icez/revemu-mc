VERSION 5.00
Begin VB.Form frmItem 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Use/Sell Inventory"
   ClientHeight    =   5850
   ClientLeft      =   7485
   ClientTop       =   5415
   ClientWidth     =   4200
   Icon            =   "frmItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Palette         =   "frmItem.frx":0E42
   Picture         =   "frmItem.frx":0E93
   ScaleHeight     =   5850
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1230
      Left            =   -20
      Picture         =   "frmItem.frx":0ED6
      ScaleHeight     =   1230
      ScaleWidth      =   300
      TabIndex        =   2
      Top             =   255
      Width           =   300
   End
   Begin VB.ListBox lstInvent 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DragIcon        =   "frmItem.frx":105F
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
      ItemData        =   "frmItem.frx":1EA1
      Left            =   300
      List            =   "frmItem.frx":1EA8
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   240
      Width           =   4320
   End
   Begin VB.Image Image1 
      Height          =   120
      Left            =   0
      Picture         =   "frmItem.frx":1EB2
      Top             =   1560
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgclose 
      Height          =   135
      Left            =   3980
      Picture         =   "frmItem.frx":1EF5
      Top             =   60
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory"
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
      Width           =   675
   End
   Begin VB.Image Image5 
      Height          =   135
      Left            =   50
      Picture         =   "frmItem.frx":202A
      Top             =   60
      Width           =   135
   End
   Begin VB.Image imgReSize 
      Height          =   180
      Left            =   2040
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmItem.frx":215F
      Top             =   3480
      Width           =   180
   End
   Begin VB.Image imgRightbar 
      Height          =   255
      Left            =   1560
      Picture         =   "frmItem.frx":22AB
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgMidbar 
      Height          =   255
      Left            =   170
      Picture         =   "frmItem.frx":2515
      Stretch         =   -1  'True
      Top             =   0
      Width           =   180
   End
   Begin VB.Image Image10 
      Height          =   255
      Left            =   0
      Picture         =   "frmItem.frx":25ED
      Top             =   0
      Width           =   180
   End
   Begin VB.Image imgbmid 
      Height          =   420
      Left            =   120
      Picture         =   "frmItem.frx":275B
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image imgbright 
      Height          =   420
      Left            =   1560
      Picture         =   "frmItem.frx":27B3
      Top             =   3480
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image imgbleft 
      Height          =   420
      Left            =   0
      Picture         =   "frmItem.frx":2828
      Top             =   3480
      Visible         =   0   'False
      Width           =   150
   End
End
Attribute VB_Name = "frmItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mhWndSubClassed As Long
Private mWndProcNext As Long

Private Sub Form_Load()
If frmMain.Visible Then UpdateInventory
frmItem.height = 4500
frmItem.width = 4300
imgRightbar.Left = frmItem.width - 200
imgMidbar.width = frmItem.width - 350
LoadFormPos frmItem
lstInvent.height = frmItem.height - 880
lstInvent.width = frmItem.width
imgbleft.Top = lstInvent.height + 200
imgbmid.Top = lstInvent.height + 200
imgbright.Top = lstInvent.height + 200
imgbright.Left = frmItem.width - 300
imgbmid.width = frmItem.width - 400
imgReSize.Top = lstInvent.height + 320
imgReSize.Left = frmItem.width - 270
'frmItem.Height = lstInvent.Height + 880
Dim tx As Integer
Dim ty As Integer
Dim tw As Integer
Dim th As Integer
Dim pw As Long
Dim ph As Long
frmItem.AutoRedraw = True
tw = Int(frmItem.width / Image1.width) + 1
th = Int((frmItem.height) / Image1.height) + 1
pw = Image1.width
ph = Image1.height
For tx = 0 To tw
    For ty = 0 To th
         frmItem.PaintPicture Image1.Picture, tx * pw, ty * ph
    Next ty
Next tx
frmItem.AutoRedraw = False
'SubClass Me.hwnd
'Inithwnd = Me.hwnd
End Sub


Private Sub Form_Resize()

If (frmItem.width < 2000 Or frmItem.height < 2000) Then
    'Form_Load
Else
imgRightbar.Left = frmItem.width - 180
imgMidbar.width = frmItem.width - 320
lstInvent.height = frmItem.height - 500
lstInvent.width = frmItem.width - 300
imgbleft.Top = lstInvent.height + 240
imgbmid.Top = lstInvent.height + 240
imgclose.Left = frmItem.width - 200
imgbright.Top = lstInvent.height + 240
imgbright.Left = frmItem.width - 100
imgbmid.width = frmItem.width - 200
imgReSize.Top = lstInvent.height + 320
imgReSize.Left = frmItem.width - 182
Dim tx As Integer
Dim ty As Integer
Dim tw As Integer
Dim th As Integer
Dim pw As Long
Dim ph As Long
frmItem.AutoRedraw = True
tw = Int(frmItem.width / Image1.width) + 1
th = Int((frmItem.height) / Image1.height) + 1
pw = Image1.width
ph = Image1.height
For tx = 0 To tw
    For ty = 0 To th
         frmItem.PaintPicture Image1.Picture, tx * pw, ty * ph
    Next ty
Next tx
frmItem.AutoRedraw = False
If (frmItem.height + 650) < MDIfrmMain.height Then frmItem.height = lstInvent.height + 500
End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
UnSubClass
MDIfrmMain.mnuInv.CheckED = False
SaveFormPos frmItem
End Sub

Private Sub SubClass(hWnd)
   Dim lResult As Long
   ' First we will make sure that the form is not allready subclassed.
   ' Else we will get an error and that's pretty bad when subclassing.
   UnSubClass
        
   ' Now we will redirect all messages regarding the form to our own
   ' WindowProc instead of VB's using AddressOf (Works with VB5 or higher)
   mWndProcNext = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf SubWndProc)
        
   ' If all went as planned mWndProc is now true and we can proceed.
   If mWndProcNext Then
      mhWndSubClassed = hWnd
      lResult = SetWindowLong(hWnd, GWL_USERDATA, ObjPtr(Me))
   End If
End Sub

Private Sub UnSubClass()
   If mWndProcNext Then
      SetWindowLong mhWndSubClassed, GWL_WNDPROC, mWndProcNext
      mWndProcNext = 0
   End If
End Sub

Friend Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, _
       ByVal wParam As Long, ByVal lParam As Long) As Long
        
   ' This is our WindowProc. All messages to the form will be
   ' sent here for us to use before VB get's them.
   Select Case uMsg
   Case WM_SIZE
      ' Form is being resized
      Debug.Print "Formens størrelse ændres"
   Case WM_GETMINMAXINFO
      ' Make sure the for is not resized smaller than
      ' 400x250 pixels.
      Debug.Print "MINMAXINFO Changed"
                
      Dim mmiT As MINMAXINFO
                
      ' Copy the parameter lParam to a local variable
      ' so that we can play around with it
      CopyMemory mmiT, ByVal lParam, Len(mmiT)
                
      ' Minimium width and height. Remember that API works with
      ' pixels instead of twips.
      mmiT.ptMinTrackSize.X = 300
      mmiT.ptMinTrackSize.Y = 300
                
      ' Copy modified results back to parameter
      CopyMemory ByVal lParam, mmiT, Len(mmiT)
                
      ' In this case we don't want VB to handle resizeing.
      ' We will exit the function without sending the message
      ' to VB.
                
      Exit Function
                
   End Select
        
   ' Redirect the messages back to VB. All unhandled messages should be
   ' redirected to VB so that we can still use VB's events.
        
   WindowProc = CallWindowProc(mWndProcNext, hWnd, uMsg, wParam, _
                ByVal lParam)
        
End Function
'-- End --'


Private Sub imgclose_Click()
frmItem.Visible = False
End Sub

Private Sub imgMidbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
SaveFormPos frmItem
End Sub

Private Sub imgReSize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(frmItem.hWnd, WM_NCLBUTTONDOWN, 17, 0)
SaveFormPos frmItem
End Sub

Private Sub lstInvent_DblClick()
    If lstInvent.List(lstInvent.ListIndex) <> "" And ViewState = 0 Then frmMain.Use_Item
End Sub

Private Sub lstInvent_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Button = vbLeftButton Then
        lstInvent.OLEDrag
    Else
        If lstInvent.List(lstInvent.ListIndex) <> "" Then
            Select Case ViewState
                Case 0
                    frmPopupChat.mnuUse.Visible = True
                    frmPopupChat.mnuDrop.Visible = True
                    frmPopupChat.mnuCart.Visible = IsCartOn
                Case 1
                    If AllInv(Val(lstInvent.List(lstInvent.ListIndex))).Pos = 0 Then
                        frmPopupChat.mnuEquip.Visible = True
                        frmPopupChat.mnuDrop.Visible = True
                        frmPopupChat.mnuCart.Visible = IsCartOn
                    Else
                        frmPopupChat.mnuUEquip.Visible = True
                        frmPopupChat.mnuCart.Visible = False
                    End If
                Case 2
                    frmPopupChat.mnuDrop.Visible = True
                    frmPopupChat.mnuCart.Visible = IsCartOn
            End Select
            Me.PopupMenu frmPopupChat.mnuItem
        End If
        frmPopupChat.mnuUse.Visible = False
        frmPopupChat.mnuDrop.Visible = False
        frmPopupChat.mnuEquip.Visible = False
        frmPopupChat.mnuCart.Visible = False
        frmPopupChat.mnuUEquip.Visible = False
    End If
    Err.Clear
End Sub

Private Sub lstInvent_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
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
            Amount = Val(InputBox("Enter number to get", "Get from storage", CStr(Storage(Index).Amount)))
        Else
            Amount = 1
        End If
        If Amount > 0 Then pkt_StorageGet Storage(Index).Index, Amount
    End If
    If strList = frmCart.lstCart.List(frmCart.lstCart.ListIndex) Then
        Index = Val(frmCart.lstCart.List(frmCart.lstCart.ListIndex))
        If Cart(Index).Amount > 1 Then
            Amount = Val(InputBox("Enter number to get", "Get from cart", CStr(Cart(Index).Amount)))
            If Amount > Cart(Index).Amount Then Amount = Cart(Index).Amount
        Else
            Amount = 1
        End If
        If Amount > 0 Then
            pkt_CartTake Index, Amount
            Stat "Cart item to inventory : "
            Stat "[" & Cart(Index).Name & "]", vbBlue
            Stat " " & Amount & "EA" + vbCrLf
        End If
    End If
    ' If the item was not dropped on itself
    strList = ""
End Sub

Private Sub lstInvent_OLEStartDrag(data As DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectMove
    data.SetData lstInvent
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y > 810 Then
        ViewState = 2
    ElseIf Y > 420 Then
        ViewState = 1
    Else
        ViewState = 0
    End If
    Update_FrmItem
    UpdateInventory
End Sub
