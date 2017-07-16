VERSION 5.00
Begin VB.Form FrmField 
   BorderStyle     =   0  'None
   Caption         =   "MAP"
   ClientHeight    =   4995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   333
   ScaleMode       =   0  'User
   ScaleWidth      =   333
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   0
      Top             =   255
      Width           =   1695
   End
   Begin VB.Label LabXY 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LabXY"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2760
      TabIndex        =   3
      Top             =   2760
      Width           =   1200
   End
   Begin VB.Image imgAI 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   60
      Picture         =   "FrmField.frx":0000
      Top             =   2760
      Width           =   225
   End
   Begin VB.Label LabAutoAI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto AI"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   2775
      Width           =   525
   End
   Begin VB.Image ImgEndBar 
      Height          =   300
      Left            =   0
      Picture         =   "FrmField.frx":00F4
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   4200
   End
   Begin VB.Image ImgClose 
      Height          =   135
      Left            =   3960
      Picture         =   "FrmField.frx":0238
      Top             =   60
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   50
      Picture         =   "FrmField.frx":036D
      Top             =   60
      Width           =   135
   End
   Begin VB.Label LabFrm 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Map"
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
      Width           =   300
   End
   Begin VB.Image ImgTopBar 
      Height          =   255
      Left            =   0
      Picture         =   "FrmField.frx":04A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4200
   End
End
Attribute VB_Name = "FrmField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    LoadFormPosOnly Me
    If ConnState > 3 Then
        Load_Field MapName
        Plot_Dot curPos, vbBlue
    End If
End Sub


Public Sub Run_Search() 'Start Routing
On Error GoTo errie
    Set objTree = New clsTree
    objTree.bAStar = False
    IsRouting = True
    If Path_Search Then
        Path_ChoosePath
    Else
        ptEnd.X = 0
        ptEnd.Y = 0
    End If
errie:
    Set objTree = Nothing
    IsRouting = False
End Sub

Public Function Path_Search() As Boolean
    Dim lNodeX As Long, lNodeY As Long
    
    Dim bShowPath As Boolean, bSlowDown As Boolean
    Dim bAchieved As Boolean
    
        'If objTree Is Nothing Then Exit Function
        
        objTree.StartSearch ptStart.X, ptStart.Y, ptEnd.X, ptEnd.Y, aiMap()
        objTree.Load_Portal MapName
        
        Do Until objTree.NextNode Or bAchieved
            lNodeLooped = lNodeLooped + 1            'Measures
            bAchieved = objTree.UpdateCurrentNode
            DoEvents
        Loop
        objTree.BackTracePath
        Path_Search = bAchieved
        
End Function

Public Sub Path_PaintNode(ByVal lx As Long, ByVal ly As Long, Optional bColorIt As Boolean = False)
    Dim Pt As Coord
    Pt.X = lx
    Pt.Y = ly
    Tile_Paint Pt, vbRed
End Sub

Public Sub Path_ChoosePath()                 'The path is ready to be traced by now
    Dim lx As Long, ly As Long, rc As Long
    Dim tmpx As Long, tmpy As Long
    Dim StartCod As Coord, EndCod As Coord
    Dim isportal As Integer
    lx = ptStart.X
    ly = ptStart.Y
    Dim i As Long
    Dim block As Byte
    Dim X, Y As Long
    Dim j As Integer
    Dim FoundPortal As Boolean
    FoundPortal = False
    Dim IsMapIN As Boolean
    IsMapIN = Check_MapInPortal(MapName)
    block = 1
    i = 0
    Dim tmpPt As Coord
    Dim tmpcurpt As Coord
    ReDim Route(0)
    Do
        rc = objTree.PathStepNext(lx, ly, isportal) 'now step forward one at a time
        
        If IsMapIN Then
            If j > 0 Then
                j = j - 1
                If j = 0 Then FoundPortal = False
            End If
            If FoundPortal Then GoTo endwhile
        End If
        
        If isportal > -1 Then
            tmpPt.X = lx
            tmpPt.Y = ly
        End If
        tmpcurpt.X = lx
        tmpcurpt.Y = ly
        If Not IsMapIN And tmpPt.X <> 0 Then
            If EvalNorm(tmpPt, tmpcurpt) < 3 Then GoTo endloopwhile
        End If
        If ((i Mod 3 = 0 Or (isportal > -1)) And Not FoundPortal) Then
            'With Route(UBound(Route))
            tmpx = lx
            tmpy = ly
           
            If isportal > -1 Then GoTo endloop
            StartCod.X = lx
            StartCod.Y = ly
            For X = lx - block To lx + block
                For Y = ly - block To ly + block
                    If X >= 0 And Y >= 0 And (X <> lx And Y <> ly) Then
                        If aiMap(X, Y) > 0 Then
                            If X < lx Then
                                If lx + block < xDim Then
                                    EndCod.X = lx + block
                                    EndCod.Y = ly
                                    If NewCanGO(StartCod, EndCod) Then tmpx = lx + block
                                End If
                            ElseIf X > lx Then
                                If lx - block > 0 Then
                                    EndCod.X = lx - block
                                    EndCod.Y = ly
                                    If NewCanGO(StartCod, EndCod) Then tmpx = lx - block
                                End If
                            End If
                            If Y > ly Then
                                If ly - block > 0 Then
                                    EndCod.X = lx
                                    EndCod.Y = ly - block
                                    If NewCanGO(StartCod, EndCod) Then tmpy = ly - block
                                End If
                            ElseIf Y < ly Then
                                If ly + block > 0 Then
                                    EndCod.X = lx
                                    EndCod.Y = ly + block
                                    If NewCanGO(StartCod, EndCod) Then tmpy = ly + block
                                End If
                            End If
                        End If
                    End If
                    If tmpx <> lx And tmpy <> ly Then GoTo endloop
                    Next
                Next
endloop:
                'lx = tmpx
                'ly = tmpy
            If i <> 0 And rc <> 0 Then
                Route(UBound(Route)).X = FrmField.PicMap.height - (tmpy)
                Route(UBound(Route)).Y = (tmpx)
                ReDim Preserve Route(UBound(Route) + 1)
            ElseIf rc = 0 Then
                Route(UBound(Route)).X = FrmField.PicMap.height - (tmpy)
                Route(UBound(Route)).Y = (tmpx)
                ReDim Preserve Route(UBound(Route) + 1)
            Else
                Route(UBound(Route)) = curPos
                ReDim Preserve Route(UBound(Route) + 1)
            End If
                    
            'End With
        End If
        
endwhile:
        If isportal > -1 And IsMapIN Then
            FoundPortal = Not FoundPortal
            'Debug.Print "Portals"
            'Debug.Print CStr(FrmField.PicMap.Height - ly) & ":" & CStr(lx)
            If Not FoundPortal Then
                FoundPortal = True
                j = 5
            End If
        End If
        'Path_PaintNode lX, lY, True
        i = i + 1
endloopwhile:
    Loop Until rc = 0                     'the last node has no link forward (like a string termination)
    ReDim Preserve Route(UBound(Route) - 1)
    'Debug.Print "Route"
    'For i = 0 To UBound(Route)
    '    Debug.Print Route(i).x & ":" & Route(i).y
    'Next
    Current = 0
End Sub

Public Sub Draw_Map()
    Dim i As Integer
    Dim j As Integer
    Dim PtPass As Coord
    'FrmField.Visible = False
    PicMap.width = MapWidth * TILE_SIDE
    PicMap.height = MapHeight * TILE_SIDE

    For i = 0 To UBound(aiMap, 1)
        For j = 0 To UBound(aiMap, 2)
            PtPass.X = i
            PtPass.Y = j
            Tile_Paint PtPass, Return_Map_Color(PtPass), False
        Next
    Next
    ImgTopBar.width = MapWidth
    imgclose.Left = MapWidth - 15
    ImgEndBar.width = MapWidth
    ImgEndBar.Top = MapHeight + 17
    imgAI.Top = MapHeight + 19
    LabXY.Top = MapHeight + 19
    LabXY.Left = MapWidth - LabXY.width - 10
    LabAutoAI.Top = MapHeight + 20
    FrmField.width = PicMap.width * 15
    FrmField.height = (PicMap.height + ImgEndBar.height + ImgTopBar.height) * 15
End Sub

Public Sub update_ImgAI()
    If AutoAI Then
        imgAI.Picture = LoadPicture(App.Path & "\interface\on.gif")
    Else
        imgAI.Picture = LoadPicture(App.Path & "\interface\off.gif")
    End If
End Sub

Private Sub Form_Resize()
SaveFormPos Me
End Sub

Private Sub Form_Terminate()
SaveFormPos Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveFormPos Me
End Sub

Private Sub imgAI_Click()
    AutoAI = Not AutoAI
    update_ImgAI
End Sub

Private Sub imgclose_Click()
Unload FrmField
End Sub

Private Sub ImgTopBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
SaveFormPos FrmField
End Sub

Private Sub PicMap_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "r" Then Draw_Map
End Sub

Private Sub PicMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Pts As Coord
    Pts.X = X
    Pts.Y = Y
    If AutoAI Then
        MsgBox "Please turn off autoAI first!!!  ", vbInformation
        Exit Sub
    End If
    If ConnState < 4 Then
    '    MsgBox "Client need to connected to the server first!!!  ", vbExclamation
    '    Exit Sub
    End If
    If Button = vbLeftButton Then
        'If Not Mode Then
        '    ptStart = Pts
        '    Tile_Paint Pts, vbGreen, True
        '    Mode = Not Mode
        'Else
            'Dim test  As String
            'test = MakeMagePos(Pts)
            'Exit Sub
            ptStart.X = curPos.Y
            ptStart.Y = (FrmField.PicMap.height - curPos.X)
            'ptStart.x = 210
            'ptStart.y = 183
            ptEnd = Pts
            If aiMap(ptEnd.X, ptEnd.Y) <> 0 Then
                MsgBox "Can't move to that coordinate!!!  ", vbExclamation
                Exit Sub
            End If
            If EvalNorm(ptStart, ptEnd) < 15 Then
                Pts.X = FrmField.PicMap.height - Y
                Pts.Y = X
                If CanGO(curPos, Pts) Then move_to Pts
                Exit Sub
            Else
                IsRandomRoute = False
            End If
            
        '    Tile_Paint ptEnd, vbRed, True
        '    Mode = Not Mode
            Run_Search
            'Routing
        'End If
    End If
End Sub

Private Sub PicMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim pos As Coord
    pos.X = X
    pos.Y = FrmField.PicMap.height - Y
    LabXY.Caption = CStr(pos.X) & ":" & CStr(pos.Y)
End Sub

