Attribute VB_Name = "Mod_Pathfinding"
Public Const HIGH_PRIORITY_CLASS = &H80
Public Const IDLE_PRIORITY_CLASS = &H40
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const REALTIME_PRIORITY_CLASS = &H100

Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long


Public objTree As clsTree
Public aiMap() As Byte
Public ptStart As Coord
Public ptEnd As Coord
Public Mode As Boolean
Public MapWidth As Integer
Public MapHeight As Integer
Public Current As Long
Public Route() As Coord
Public AllowCoord() As Coord
Public OldDot As Coord
Public GoOnRoute As Boolean
Public IsRandomRoute As Boolean
Public CanusePath As Boolean



Public Function Arctan(Y As Integer, X As Integer) As Integer
    Dim tmpAngle As Integer
    If X <> 0 Then
        tmpAngle = Atn(Y / X) * (180 / (3.1416))
        If Sgn(Y) = -1 And Sgn(X) = -1 Then
            tmpAngle = tmpAngle - 180
        ElseIf Sgn(Y) = 1 And Sgn(X) = -1 Then
            tmpAngle = 180 + tmpAngle
        ElseIf Sgn(Y) = 0 And Sgn(X) = 1 Then
            tmpAngle = 0
        ElseIf Sgn(Y) = 0 And Sgn(X) = -1 Then
            tmpAngle = 180
        End If
    Else
        If Sgn(Y) = 1 Then
            tmpAngle = 90
        ElseIf Sgn(Y) = -1 Then
            tmpAngle = -90
        End If
    End If
    Arctan = tmpAngle
End Function

Public Sub FindNearerPoint(ByRef Pt As Coord)
    Dim Des As Coord
    Dim i, j As Integer
    Dim xLD, xUD, xDim, yLD, yUD, yDim As Long
    xLD = LBound(aiMap, 1)
    xUD = UBound(aiMap, 1)
    xDim = xUD - xLD + 1
    
    yLD = LBound(aiMap, 2)
    yUD = UBound(aiMap, 2)
    yDim = yUD - yLD + 1
    Des.X = Pt.X
    Des.Y = MapHeight - Pt.Y
    If aiMap(Des.X, Des.Y) = 0 Then Exit Sub
    For i = Des.X - 2 To Des.X + 2
        For j = Des.Y - 2 To Des.Y + 2
            If i >= 0 And j >= 0 And i < xDim And j < yDim Then
                If aiMap(i, j) = 0 Then
                    Pt.X = i
                    Pt.Y = MapHeight - j
                End If
            End If
        Next
    Next
End Sub

Public Function IsMovePoint(Pt As Coord) As Boolean
    Dim Des As Coord
    Des.X = Pt.X
    Des.Y = MapHeight - Pt.Y
    If aiMap(Des.X, Des.Y) = 0 Then
        IsMovePoint = True
    Else
        IsMovePoint = False
    End If
End Function

Public Function CanGO(pt1 As Coord, Pt As Coord) As Boolean
    Dim Des As Coord
    Dim Src As Coord
    Dim count As Integer
    Src.X = pt1.Y
    Src.Y = MapHeight - pt1.X
    Des.X = Pt.Y
    Des.Y = MapHeight - Pt.X
    count = 0
    Do
        If aiMap(Src.X + Sgn(Des.X - Src.X), Src.Y) = 0 Then Src.X = Src.X + Sgn(Des.X - Src.X)
        If aiMap(Src.X, Src.Y + Sgn(Des.Y - Src.Y)) = 0 Then Src.Y = Src.Y + Sgn(Des.Y - Src.Y)
        count = count + 1
        If count > 500 Then
            CanGO = False
            Exit Function
        End If
        If Src.X = Des.X And Src.Y = Des.Y Then GoTo endloop
        If Src.X = Des.X Then
            If aiMap(Src.X, Src.Y + Sgn(Des.Y - Src.Y)) <> 0 Then
                CanGO = False
                Exit Function
            End If
        End If
        If Src.Y = Des.Y Then
           If aiMap(Src.X + Sgn(Des.X - Src.X), Src.Y) <> 0 Then
                CanGO = False
                Exit Function
            End If
        End If
endloop:
    Loop While (Src.X <> Des.X Or Src.Y <> Des.Y)
    CanGO = True
End Function

Public Function CanAttackRoute(pt1 As Coord, Pt As Coord) As Boolean
    Dim Des As Coord
    Dim Src As Coord
    Dim count As Integer
    Src.X = pt1.Y
    Src.Y = MapHeight - pt1.X
    Des.X = Pt.Y
    Des.Y = MapHeight - Pt.X
    count = 0
    Do
        If aiMap(Src.X + Sgn(Des.X - Src.X), Src.Y) = 0 Then Src.X = Src.X + Sgn(Des.X - Src.X)
        If aiMap(Src.X, Src.Y + Sgn(Des.Y - Src.Y)) = 0 Then Src.Y = Src.Y + Sgn(Des.Y - Src.Y)
        count = count + 1
        If count > 30 Then
            CanAttackRoute = False
            Exit Function
        End If
        If Src.X = Des.X And Src.Y = Des.Y Then GoTo endloop
        If Src.X = Des.X Then
            If aiMap(Src.X, Src.Y + Sgn(Des.Y - Src.Y)) <> 0 Then
                CanAttackRoute = False
                Exit Function
            End If
        End If
        If Src.Y = Des.Y Then
           If aiMap(Src.X + Sgn(Des.X - Src.X), Src.Y) <> 0 Then
                CanAttackRoute = False
                Exit Function
            End If
        End If
endloop:
    Loop While (Src.X <> Des.X Or Src.Y <> Des.Y)
    CanAttackRoute = True
End Function

Public Sub Tile_Paint(ByRef ptWhich As Coord, ByVal lColor As Long, Optional bRefresh As Boolean = True)
    SetPixel FrmField.PicMap.hDC, ptWhich.X, ptWhich.Y, lColor
End Sub

Public Function Return_Map_Color(Pt As Coord) As Long
        Select Case aiMap(Pt.X, Pt.Y)
            Case O
                Return_Map_Color = &H9BDFC4
            Case 1
                Return_Map_Color = 3551021
            Case 5
                Return_Map_Color = &HB5B6B5
        End Select
End Function

Public Sub Clear_Dot(Pt As Coord, Optional RFMap As Boolean = True)
On Error GoTo errie
    'If Not FrmField.Visible Then Exit Sub
    Dim i, j, size As Integer
    Dim Pts As Coord
    size = 1
    For i = -size To size
        For j = -size To size
            Pts.X = Pt.Y + i
            Pts.Y = (FrmField.PicMap.height - Pt.X) + j
            If Pts.X > 0 And Pts.Y > 0 Then Tile_Paint Pts, Return_Map_Color(Pts), True
        Next
    Next
    If RFMap Then FrmField.PicMap.Refresh
    Exit Sub
errie:
    Err.Clear
End Sub


Public Sub Plot_Dot(Pt As Coord, Color As Long)
On Error GoTo errie
    'If Not FrmField.Visible Then Exit Sub
    Dim i, j, size As Integer
    Dim Pts As Coord
    size = 1
    For i = -size To size
        For j = -size To size
            Pts.X = Pt.Y + i
            Pts.Y = (FrmField.PicMap.height - Pt.X) + j
            If Pts.X > 0 And Pts.Y > 0 Then Tile_Paint Pts, Color, True
        Next
    Next
    FrmField.PicMap.Refresh
    Exit Sub
errie:
    Err.Clear
End Sub

Public Sub Plot_Dot3(Pt As Coord, Color As Long)
On Error GoTo errie
    'If Not FrmField.Visible Then Exit Sub
    Dim i, j, size As Integer
    Dim Pts As Coord
    size = 1
    For i = -size To size
        For j = -size To size
            Pts.X = Pt.Y + i
            Pts.Y = (FrmField.PicMap.height - Pt.X) + j
            If Pts.X > 0 And Pts.Y > 0 Then Tile_Paint Pts, Color, True
        Next
    Next
    Exit Sub
errie:
    Err.Clear
End Sub


Public Sub Plot_Dot2(Pt As Coord, Color As Long)
On Error GoTo errie
    'If Not FrmField.Visible Then Exit Sub
    'Dim i, j, size As Integer
    Dim Pts As Coord
    'size = 0
    'For i = -size To size
    '    For j = -size To size
            Pts.X = Pt.Y ' + i
            Pts.Y = (FrmField.PicMap.height - Pt.X) ' + j
            If Pts.X > 0 And Pts.Y > 0 Then Tile_Paint Pts, Color, True
    '    Next
    'Next
    FrmField.PicMap.Refresh
    Exit Sub
errie:
    Err.Clear
End Sub

Public Sub Set_Nearer_Point(Pt As Coord)
    Dim i As Integer
    Dim distance As Integer
    Dim OldDis As Integer
    OldDis = 9999
    For i = 0 To UBound(Route)
        distance = EvalNorm(Pt, Route(i))
        If distance < OldDis And distance < 10 Then
            OldDis = distance
            Exit For
        End If
    Next
    If OldDis = 0 Then
        Current = i + 1
    Else
        Current = i
    End If
End Sub


Public Sub Load_Field(field As String)
On Error GoTo errie
    Dim lfile As Long
    Dim test() As Byte
    TILE_SIDE = 1
    lfile = FreeFile
    'MapPath = "C:\REVMEU\gat"
    Open MapPath & "\" & field & ".gat" For Binary Access Read As lfile
    Get lfile, , MapWidth
    Get lfile, , MapHeight
    ReDim aiMap(MapWidth - 1, MapHeight - 1)
    ReDim test(MapWidth - 2, MapHeight - 2) As Byte
    Stat ". ", vbRed, False, True
    ReDim AllowCoord(0)
    Get lfile, , aiMap()
    Close lfile
    Dim i As Integer
    Dim j As Integer
    For i = 0 To UBound(aiMap, 1) - 1
        For j = 0 To UBound(aiMap, 2) - 1
            If aiMap(i, j) = &HFF Then
                test(i, UBound(aiMap, 2) - j - 1) = 0
            ElseIf aiMap(i, j) = &H80 Then
                test(i, UBound(aiMap, 2) - j - 1) = 5
            ElseIf aiMap(i, j) = &H0 Then
                test(i, UBound(aiMap, 2) - j - 1) = 1
            End If
            If aiMap(i, j) = &HFF Then
                AllowCoord(UBound(AllowCoord)).X = i
                AllowCoord(UBound(AllowCoord)).Y = UBound(aiMap, 2) - j - 1
                ReDim Preserve AllowCoord(UBound(AllowCoord) + 1)
            End If
        Next
    Next
    Stat ". ", vbRed, False, True
    ReDim aiMap(MapWidth - 2, MapHeight - 2)
    ReDim Preserve AllowCoord(UBound(AllowCoord) - 1)
    MapWidth = MapWidth - 2
    MapHeight = MapHeight - 2
    aiMap() = test()
    FrmField.Visible = True
    Stat ". ", vbRed, False, True
    FrmField.Draw_Map
    Check_RouteMap
    Exit Sub
errie:
    Close lfile
    MsgBox "Error!!! on loading " & MapPath & "\" & field & ".gat", vbCritical
    Unload MDIfrmMain
End Sub

