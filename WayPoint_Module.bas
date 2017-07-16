Attribute VB_Name = "WayPoint_Module"
Public useMinDistance As Boolean
Public MinDistance As Integer

Public PlayerMoveTime As Long
Public MovementSpeed As Integer
Public ServerName As String
Public IsDMove As Boolean
Public DirectionID As Byte
Public LastJamNumber As Byte
Public ChatRoomName As String
Public AlwaySit As Boolean
Public PacketData As String
Public WayPoint() As Coord
Public FightPortal() As Coord
Public FightDirection As String
Public SellDirection As String
Public SellPortal() As Coord
Public StartPoint As Long
Public Direction As String
Public CanUseWP As Boolean
Public MoveOnly As Boolean
Public FightMap As Boolean
Public BackMap As Boolean
Public SellMode As Boolean
Public FightMode As Boolean
Public SelItem() As Item
Public BuyItem() As BackCode
Public Kafra() As Item
Public IsBackTown As Boolean
Public WeightBackTown As Double
Public IsCulvert As Boolean
Public CulVertNPC As String
Public indexFight As Integer
Public indexSell As Integer
Public SaveMapName As String
Public mapH As Double
Public mapW As Double
Public curPos As Coord
Public WalkMap As Boolean
Public AutoAI As Boolean

Public Function MakeMagePos(Coords As Coord) As String
Dim tstr As String
Dim Offset As Integer
Dim newcoords As Coord
Dim CurAngle As Integer
Dim GetAngle As Integer
Dim i As Long
If useMinDistance Then
    Offset = MinDistance
Else
    Offset = 5
End If

For i = 0 To UBound(AllowCoord)
    newcoords.X = MapHeight - AllowCoord(i).Y
    newcoords.Y = AllowCoord(i).X
    If EvalNorm(newcoords, Coords) = Offset Then
        If CanGO(curPos, newcoords) Then
            CurAngle = Arctan((Coords.X - curPos.X), (Coords.Y - curPos.Y))
            GetAngle = Arctan((Coords.X - newcoords.X), (Coords.Y - newcoords.Y))
            If Abs(CurAngle - GetAngle) < 45 Then Exit For
        End If
    End If
Next
tstr = tstr + Chr(Int((newcoords.Y) / 4))
tstr = tstr + Chr(((newcoords.Y) Mod 4) * 64 + Int((newcoords.X) / 16))
tstr = tstr + Chr(((newcoords.X) Mod 16) * 16)
MakeMagePos = tstr
End Function

Public Sub GetTeleportName(ByRef MapName() As String)
Dim i, count  As Integer
On Error GoTo errie
    count = UBound(MapName)
    Dim maplist, Name As String
    Dim Index As Integer
    Dim index2 As Integer
    Open App.Path & "\Table\MapName.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, maplist
        Index = InStr(maplist, "#")
        index2 = InStr(maplist, "^")
        If Index > 0 And index2 > 0 Then
            Name = Left(maplist, InStr(maplist, ".gat") - 1)
            maplist = Trim(Right(maplist, Len(maplist) - index2))
            For i = 0 To UBound(MapName)
                If MapName(i) = maplist Then
                MapName(i) = Name
                count = count - 1
                If count = 0 Then GoTo errie
                Exit For
                End If
            Next
        End If
    Loop
errie:
    Close 1
End Sub

Public Function GetMapname(ByVal Name As String) As String
    On Error GoTo errie
    Dim maplist As String
    Dim Index As Integer
    Dim index2 As Integer
    Open App.Path & "\table\MapName.txt" For Input As #1
    Do While Not EOF(1)
        Input #1, maplist
        Index = InStr(maplist, "#")
        index2 = InStr(maplist, "$")
        If LCase(Left(maplist, Len(Name))) = LCase(Name) Then
                If Index > 0 And index2 > 0 Then
                    Name = Mid(maplist, Index + 1, index2 - Index - 1)
                    maplist = Trim(Right(maplist, Len(maplist) - index2))
                    mapW = Val(Left(maplist, InStr(maplist, ":") - 1))
                    mapH = Val(Right(maplist, Len(maplist) - InStr(maplist, ":")))
                ElseIf Index > 0 And index2 = 0 Then
                    Name = Trim(Right(maplist, Len(maplist) - Index))
                End If
                GetMapname = Name
                Close 1
                Exit Function
            Exit Do
        End If
    Loop
errie:
    Close 1
    'Stat "Mapname.txt need to update..." & vbCrLf
    GetMapname = Name
End Function

Public Function LoadMAP(ByVal Name As String) As Boolean
    On Error GoTo errie
    frmMap.labMap.Caption = GetMapname(Name)
    Dim Result As Boolean
    Result = False
    frmMap.Image1.Picture = LoadPicture(MapPath & "\" & Name & ".jpg")
    Result = True
errie:
    'If Not Result Then Stat "Can't find " & MapPath & "\" & name & ".jpg" & vbCrLf
    LoadMAP = Result
End Function

Public Function MakeCoordString(Coords As Coord) As String
Dim tstr As String
tstr = tstr + Chr(Int(Coords.Y / 4))
tstr = tstr + Chr((Coords.Y Mod 4) * 64 + Int(Coords.X / 16))
tstr = tstr + Chr((Coords.X Mod 16) * 16)
MakeCoordString = tstr
End Function

Public Sub move_to(Point As Coord, Optional PosRange As Byte = 0)
    'OnRoute = True
    Dim mPos As Coord, i As Byte
    mPos = Point
    If PosRange > 0 Then
        For i = 1 To PosRange
            mPos = NextPos(mPos, curPos)
        Next
    End If
    IsRandommove = False
    Winsock_SendPacket Chr(&H85) & Chr(0) & MakeCoordString(mPos), True
    If PosRange > 0 Then ReDim Route(0)
End Sub

Public Sub Load_WayPoint(Name As String)
On Error GoTo Out
Dim tstr As String
Dim Index As Integer
Dim index2 As Integer
Dim text As String
Dim tmpstr As String
ReDim WayPoint(0)
'Close 1
IsCulvert = False
MoveOnly = False
FightMap = True
FightDirection = "FW"
ReDim FightPortal(0)
ReDim SellPortal(0)
Open App.Path & "\waypoint\" & Name & ".wap" For Input As #1
Do While Not EOF(1)
    Line Input #1, tstr
    If InStr(tstr, "=") Then
        Index = InStr(tstr, "=")
        If Index > 0 Then
            text = LCase(Trim(Left(tstr, Index - 1)))
            Select Case text
                Case "moveonly"
                    If Trim(Right(tstr, Len(tstr) - Index)) = "1" Then
                        MoveOnly = True
                    Else
                        MoveOnly = False
                    End If
                Case "fightmode"
                    If Trim(Right(tstr, Len(tstr) - Index)) = "FW" Then
                        FightDirection = "FW"
                    ElseIf Trim(Right(tstr, Len(tstr) - Index)) = "BW" Then
                        FightDirection = "BW"
                    End If
                    If Direction = "" Then Direction = FightDirection
                Case "sellmode"
                    If Trim(Right(tstr, Len(tstr) - Index)) = "FW" Then
                        SellDirection = "FW"
                    ElseIf Trim(Right(tstr, Len(tstr) - Index)) = "BW" Then
                        SellDirection = "BW"
                    End If
                Case "gomap"
                    index2 = InStr(tstr, "#")
                    
                    If index2 = 0 Then GoTo ok2
                    Dim test As String
                    test = Trim(Mid(tstr, Index + 1, index2 - Index - 1))
                    If Trim(Mid(tstr, Index + 1, index2 - Index - 1)) = "1" Then
                        FightMap = True
                        tmpstr = Trim(Right(tstr, Len(tstr) - index2))
                        While (index2 > 0)
                            index2 = InStr(tmpstr, "#")
                            If index2 > 0 Then
                                tstr = Left(tmpstr, index2 - 1)
                                tmpstr = Trim(Right(tmpstr, Len(tmpstr) - index2))
                            Else
                                tstr = tmpstr
                            End If
                            If InStr(tstr, ":") > 0 Then
                                FightPortal(UBound(FightPortal)).Y = Val(Left(tstr, InStr(tstr, ":") - 1))
                                FightPortal(UBound(FightPortal)).X = Val(Right(tstr, Len(tstr) - InStr(tstr, ":")))
                                ReDim Preserve FightPortal(UBound(FightPortal) + 1)
                            End If
                        Wend
                                ReDim Preserve FightPortal(UBound(FightPortal) - 1)
                    Else
ok2:
                        FightMap = False
                    End If
                Case "culvert"
                    If Trim(Right(tstr, Len(tstr) - Index)) = "1" Then
                        IsCulvert = True
                    End If
                Case "backmap"
                    index2 = InStr(tstr, "#")
                    If index2 = 0 Then GoTo ok
                    If Trim(Mid(tstr, Index + 1, index2 - Index - 1)) = "1" Then
                        BackMap = True
                        tmpstr = Trim(Right(tstr, Len(tstr) - index2))
                        While (index2 > 0)
                            index2 = InStr(tmpstr, "#")
                            If index2 > 0 Then
                                tstr = Left(tmpstr, index2 - 1)
                                tmpstr = Trim(Right(tmpstr, Len(tmpstr) - index2))
                            Else
                                tstr = tmpstr
                            End If
                            If InStr(tstr, ":") > 0 Then
                                SellPortal(UBound(SellPortal)).Y = Val(Left(tstr, InStr(tstr, ":") - 1))
                                SellPortal(UBound(SellPortal)).X = Val(Right(tstr, Len(tstr) - InStr(tstr, ":")))
                                ReDim Preserve SellPortal(UBound(SellPortal) + 1)
                            End If
                        Wend
                                ReDim Preserve SellPortal(UBound(SellPortal) - 1)
                    Else
ok:
                        BackMap = False
                    End If
            End Select
        End If
    ElseIf InStr(tstr, ":") > 0 Then
        WayPoint(UBound(WayPoint)).Y = Val(Left(tstr, InStr(tstr, ":") - 1))
        WayPoint(UBound(WayPoint)).X = Val(Right(tstr, Len(tstr) - InStr(tstr, ":")))
        ReDim Preserve WayPoint(UBound(WayPoint) + 1)
    End If
Loop
ReDim Preserve WayPoint(UBound(WayPoint) - 1)
Close 1
CanUseWP = True
Exit Sub
Out:
Close 1
CanUseWP = False
End Sub

Public Function EvalNorm(coord1 As Coord, coord2 As Coord) As Long
On Error GoTo Out
    Dim a As Long
    Dim X As Long
    Dim Y As Long
    
    X = Abs(coord1.X - coord2.X) * Abs(coord1.X - coord2.X)
    Y = Abs(coord1.Y - coord2.Y) * Abs(coord1.Y - coord2.Y)
    EvalNorm = Sqr(X + Y)
    Exit Function
Out:
    EvalNorm = 0
End Function

Public Function Make_Start_Point(P As Coord) As Boolean
On Error GoTo Out
Dim i As Long
Dim X As Integer
Dim Y As Integer
Dim closest As Integer
Y = 55
If UBound(WayPoint) = 0 Then GoTo Out
Dim found As Boolean
    For i = 0 To UBound(WayPoint)
        X = EvalNorm(P, WayPoint(i))
        If X < Y Then
            StartPoint = i
            Y = X
        End If
    Next
    If Y < 15 Then
        Make_Start_Point = True
        Exit Function
    End If
Out:
    Make_Start_Point = False
End Function

Public Function Find_Near_Point(P As Coord) As Boolean
Dim i As Long
Dim X As Integer
Dim Y As Integer
Y = 55
Dim found As Boolean
    For i = 0 To UBound(WayPoint)
        X = EvalNorm(P, WayPoint(i))
        If X < Y And X > 0 Then
            StartPoint = i
            Y = X
        End If
    Next
    If Y < 12 Then
        Find_Near_Point = True
        Exit Function
    End If
    Find_Near_Point = False
End Function

Public Function IsOnWayPoint(P As Coord) As Boolean
Dim found As Boolean
found = False
If UBound(WayPoint) = 0 Then
    IsOnWayPoint = False
    Exit Function
End If
    For i = 0 To UBound(WayPoint)
        If EvalNorm(P, WayPoint(i)) < 15 Then
            found = True
            Exit For
        End If
    Next
    IsOnWayPoint = found
End Function

Public Function IsWalkOnWayPoint(P As Coord) As Boolean
Dim found As Boolean
found = False
    For i = 0 To UBound(WayPoint)
        If EvalNorm(P, WayPoint(i)) = 0 Then
            found = True
            Exit For
        End If
    Next
    IsWalkOnWayPoint = found
End Function

Public Function IsNearWayPoint(P As Coord) As Boolean
Dim found As Boolean
found = False
    For i = 0 To UBound(WayPoint)
        If EvalNorm(P, WayPoint(i)) < 15 Then
            found = True
            Exit For
        End If
    Next
    IsNearWayPoint = found
End Function

Public Function IsNearFightPortal(P As Coord) As Boolean
Dim found As Boolean
found = False
If FightPortal(0).X = 0 And FightPortal(0).Y = 0 Then GoTo endfight
    For i = 0 To UBound(FightPortal)
        If EvalNorm(P, FightPortal(i)) < 8 Then
            indexFight = i
            found = True
            Exit For
        End If
    Next
endfight:
    IsNearFightPortal = found
End Function

Public Function IsNearSellPortal(P As Coord) As Boolean
Dim found As Boolean
found = False
If SellPortal(0).X = 0 And SellPortal(0).Y = 0 Then GoTo endsell
    For i = 0 To UBound(SellPortal)
        If EvalNorm(P, SellPortal(i)) < 8 Then
            indexSell = i
            found = True
            Exit For
        End If
    Next
endsell:
    IsNearSellPortal = found
End Function

Public Function Set_Start_Point(P As Coord) As Boolean
Dim i As Long
Dim X As Integer
Dim Y As Integer
    For i = 0 To UBound(WayPoint)
        X = EvalNorm(P, WayPoint(i))
        If X < 3 Then
            StartPoint = i
            Set_Start_Point = True
            Exit Function
        End If
    Next
    Set_Start_Point = False
End Function

Public Function Move_By_Direction(P As Coord, Did As Byte) As Coord
    Select Case Did
    Case 0
        P.X = P.X + RandomNumber(8, 3)
    Case 1
        P.X = P.X + RandomNumber(8, 3)
        P.Y = P.Y - RandomNumber(8, 3)
    Case 2
        P.Y = P.Y - RandomNumber(8, 3)
    Case 3
        P.X = P.X - RandomNumber(8, 3)
        P.Y = P.Y - RandomNumber(8, 3)
    Case 4
        P.X = P.X - RandomNumber(8, 3)
    Case 5
        P.X = P.X - RandomNumber(8, 3)
        P.Y = P.Y + RandomNumber(8, 3)
    Case 6
        P.Y = P.Y + RandomNumber(8, 3)
    Case 7
        P.X = P.X + RandomNumber(8, 3)
        P.Y = P.Y + RandomNumber(8, 3)
    End Select
    Move_By_Direction = P
End Function
