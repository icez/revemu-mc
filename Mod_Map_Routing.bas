Attribute VB_Name = "Mod_Map_Routing"
Type MyPortal
    Name As String
    Pos As Coord
End Type

Type MyinPortal
    Src As MyPortal
    Des As MyPortal
End Type

Public MapRoute() As MyinPortal
Public CurRoute() As MyPortal
Public IsRouting As Boolean
Public BiDirection As Boolean
Public ForceBuy As Boolean
Public map_time_limit As Integer

Public LockMapName As String

Function MakePort(ByVal rawPort As String) As Long
    Dim tmpMP As Long
    tmpMP = CLng(Format("&H" & ReverseHex(ChrtoHex(rawPort))))
    If Len(rawPort) = 2 And tmpMP > 32768 Then tmpMP = tmpMP - 65536
    MakePort = tmpMP
End Function

Public Sub test_maproute()

Dim j As Integer
    MapName = "geffen"
    LockMapName = "prontera"
        
    Do_Search_Map_Routing MapName, "geffen_in"
For j = 0 To UBound(ai_npc)
    If Check_MapInPortal(ai_npc(j).location) Then

        MapName = "amatsu"
        LockMapName = ai_npc(j).location
        
        Do_Search_Map_Routing MapName, LockMapName
        
        'Replace_Map_Routing 1, LockMapName, ai_npc(j).pos
       
    End If
Next

End Sub

Public Sub Go_NearerPoint(pt1 As Coord, Pt2 As Coord)
    Dim PtCur As Coord
    Dim ptEnd As Coord
    Dim diff As Integer
    PtCur.X = pt1.Y
    PtCur.Y = FrmField.PicMap.height - pt1.X
    ptEnd.X = Pt2.Y
    ptEnd.Y = FrmField.PicMap.height - Pt2.X
    If aiMap(ptEnd.X, ptEnd.Y) > 0 Then
       If aiMap(ptEnd.X, ptEnd.Y + Sgn(PtCur.Y - ptEnd.Y)) = 0 Then
            ptEnd.X = FrmField.PicMap.height - (ptEnd.Y + Sgn(PtCur.Y - ptEnd.Y))
            ptEnd.Y = Pt2.Y
       ElseIf aiMap(ptEnd.X + Sgn(PtCur.X - ptEnd.X), ptEnd.Y) = 0 Then
            ptEnd.X = FrmField.PicMap.height - (ptEnd.Y)
            ptEnd.Y = Pt2.Y + Sgn(PtCur.X - Pt2.Y)
       End If
    Else
            ptEnd.X = Pt2.X
            ptEnd.Y = Pt2.Y
    End If
    move_to ptEnd
End Sub

Public Function Find_MapPortal(Name As String) As Boolean
    On Error GoTo errie
    Dim tstr As String
    Open App.Path & "\maproute\portals.txt" For Input As #10
    Do While Not EOF(10)
        Line Input #10, tstr
        tstr = Trim(tstr)
        If InStr(tstr, Name & " ") = 1 Then
            Find_MapPortal = True
            Close 10
            Exit Function
        End If
    Loop
    Close 10
    Find_MapPortal = False
    Exit Function
errie:
    Close 10
    Find_MapPortal = False
    MsgBox "Error!!! on loading 'maproute\portals.txt' (Find_MapPortal)", vbCritical
End Function

Public Function get_errcode(code As Byte) As String
    Select Case code
        Case 0
            get_errcode = "Sucessful"
        Case 1
            get_errcode = "time-out"
        Case 2
            get_errcode = "error"
    End Select
End Function

Public Sub Do_Search_Map_Routing(start As String, Destination As String)
    
    Dim tmpMapRoute() As MyinPortal
    Dim errcode As Byte
    Dim select_path As Byte
    If Find_MapPortal(start) Then
        If Find_MapPortal(Destination) Then
            
            errcode = Map_Routing(start, Destination)
            Stat "1# Searching [" & start & "] to [" & Destination & "]... "
            If errcode = 0 Then
                Stat "got " & CStr(UBound(MapRoute) + 1) & " Map(s)"
                select_path = 1
            Else
                Stat get_errcode(errcode)
            End If
            Stat vbCrLf
            
            If BiDirection Or errcode <> 0 Then
                ReDim tmpMapRoute(UBound(MapRoute))
                tmpMapRoute() = MapRoute()
                ReDim MapRoute(0)
                Stat "2# Searching with bi-direction... "
                errcode = Map_Routing(Destination, start)
                If errcode = 0 Then
                    Stat "got " & CStr(UBound(MapRoute) + 1) & " Map(s)"
                Else
                    Stat get_errcode(errcode)
                End If
                Stat vbCrLf
                If (UBound(MapRoute) < UBound(tmpMapRoute) Or tmpMapRoute(0).Src.Name = "") And MapRoute(0).Src.Name <> "" Then
                    ReDim tmpMapRoute(UBound(MapRoute))
                    For i = 0 To UBound(MapRoute)
                        tmpMapRoute(UBound(MapRoute) - i).Des = MapRoute(i).Src
                        tmpMapRoute(UBound(MapRoute) - i).Src = MapRoute(i).Des
                    Next
                    MapRoute() = tmpMapRoute()
                    select_path = 2
                Else
                    ReDim MapRoute(UBound(tmpMapRoute))
                    MapRoute() = tmpMapRoute()
                End If
            End If
            
            If UseNPCWarp Then
                ReDim tmpMapRoute(UBound(MapRoute))
                tmpMapRoute() = MapRoute()
                Stat "3# Searching with npc warp... "
                errcode = Map_Routing(start, Destination, , , , , True)
                If errcode = 0 Then
                    Stat "got " & CStr(UBound(MapRoute) + 1) & " Map(s)"
                Else
                    Stat get_errcode(errcode)
                End If
                Stat vbCrLf
                If tmpcode = 0 Then
                    If Not ((UBound(MapRoute) < UBound(tmpMapRoute) Or tmpMapRoute(0).Src.Name = "") And MapRoute(0).Src.Name <> "") Then
                        ReDim MapRoute(UBound(tmpMapRoute))
                        MapRoute() = tmpMapRoute()
                    Else
                        select_path = 3
                    End If
                End If
                
                If UseNPCBiDirect Or errcode <> 0 Then
                    ReDim tmpMapRoute(UBound(MapRoute))
                    tmpMapRoute() = MapRoute()
                    ReDim MapRoute(0)
                    Stat "4# Searching with bi-direction npc warp ... "
                    errcode = Map_Routing(Destination, start, , , , , True)
                    If errcode = 0 Then
                        Stat "got " & CStr(UBound(MapRoute) + 1) & " Map(s)"
                    Else
                        Stat get_errcode(errcode)
                    End If
                    Stat vbCrLf
                    If (UBound(MapRoute) < UBound(tmpMapRoute) Or tmpMapRoute(0).Src.Name = "") And MapRoute(0).Src.Name <> "" Then
                        ReDim tmpMapRoute(UBound(MapRoute))
                        For i = 0 To UBound(MapRoute)
                            tmpMapRoute(UBound(MapRoute) - i).Des = MapRoute(i).Src
                            tmpMapRoute(UBound(MapRoute) - i).Src = MapRoute(i).Des
                        Next
                        MapRoute() = tmpMapRoute()
                        select_path = 4
                    Else
                        ReDim MapRoute(UBound(tmpMapRoute))
                        MapRoute() = tmpMapRoute()
                    End If
                End If
                'If UseNPCBiDirect Then
                '    ReDim tmpMapRoute(UBound(MapRoute))
                '    tmpMapRoute() = MapRoute()
                '    Stat "4# Searching with bi-direction npc warp... "
                '    errcode = Map_Routing(Destination, start, , , , , True)
                '    If errcode = 0 Then
                '        Stat "got " & CStr(UBound(MapRoute) + 1) & " Map(s)"
                '    Else
                '        Stat get_errcode(errcode)
                '    End If
                '    Stat vbCrLf
                '    If errcode = 0 Then
                '        If Not ((UBound(MapRoute) < UBound(tmpMapRoute) Or tmpMapRoute(0).Src.Name = "") And MapRoute(0).Src.Name <> "") Then
                '            ReDim tmpMapRoute(UBound(MapRoute))
                '            For i = 0 To UBound(MapRoute)
                '                tmpMapRoute(UBound(MapRoute) - i).Des = MapRoute(i).Src
                '                tmpMapRoute(UBound(MapRoute) - i).Src = MapRoute(i).Des
                '            Next
                '            MapRoute() = tmpMapRoute()
                '            select_path = 4
                '        End If
                '    End If
                'End If
            End If
            If select_path > 0 Then
                Stat "Successfull Map Route, Select searching number " & CStr(select_path) & "..." & vbCrLf
                For i = 0 To UBound(MapRoute)
                    Stat "From [" & MapRoute(i).Src.Name & " (" & MapRoute(i).Src.Pos.X & "," & MapRoute(i).Src.Pos.Y & ")] to [" & MapRoute(i).Des.Name & " (" & MapRoute(i).Des.Pos.X & "," & MapRoute(i).Des.Pos.Y & ")]" & vbCrLf, &H909090, True
                Next
            Else
                Stat "AI Map routing can't find any path..." & vbCrLf
            End If
        Else
            Stat "Can't find any portals for [" & Destination & "], Routing Failed..." & vbCrLf
        End If
    Else
        Stat "Can't find any portals for [" & start & "], Routing Failed..." & vbCrLf
    End If
End Sub

Public Sub Replace_Map_Routing(Mode As Byte, Name As String, Pts As Coord, Optional i As Integer = 0, Optional Name2 As String = "", Optional swap As Boolean = False)
    
    Dim tmpPortal As MyPortal

    'Are we on Discontious Map ?
    'If Check_MapInPortal(Name) Then 'Yes
        'Use extended routing to check where's the portals?
        'Mode 1 use for go to Discontinuous MAP
        'Mode 2 use for go out frim Discontinuous MAP
        Dim tmpPt As Coord
        If swap Then
            tmpPt.X = Pts.Y
            tmpPt.Y = Pts.X
        Else
            tmpPt.X = Pts.X
            tmpPt.Y = Pts.Y
        End If
        
        Extend_Map_Routing Mode, Name, tmpPt, tmpPortal, Name2
        
        If tmpPortal.Pos.X <> 0 And tmpPortal.Pos.Y <> 0 Then
            'We can find the Portals
            Dim Index As Integer
            'Mode 2 if we need to go out then set the portals to the 1st position
            'Mode 1 If we need to go to discontinuous map so set the last position
            Index = 0
            If Mode = 1 Then Index = UBound(MapRoute)
            If i > 0 Then Index = i
            MapRoute(Index).Src.Pos.X = tmpPortal.Pos.X
            MapRoute(Index).Src.Pos.Y = tmpPortal.Pos.Y
        End If
    'End If
    
End Sub


Public Function get_warpsolutions(MapName As String, Pos As Coord) As Boolean
    Dim i As Integer
    Dim myPos As Coord
    myPos.X = Pos.Y
    myPos.Y = Pos.X
    If MapRoute(0).Src.Name <> "" Then
        For i = 0 To UBound(MapRoute)
            If MapName = MapRoute(i).Src.Name Then
                If MapRoute(i).Src.Pos.X = myPos.X And MapRoute(i).Src.Pos.Y = myPos.Y Then
                    get_warpsolutions = True
                    Exit Function
                End If
            End If
        Next
    End If
    get_warpsolutions = False
End Function

Public Function get_buy_solutions(MapName As String) As Integer
    Dim i As Integer
    If MapRoute(0).Des.Name <> "" Then
        For i = 0 To UBound(MapRoute)
            If MapName = MapRoute(i).Des.Name Then
                get_buy_solutions = i
                Exit Function
            End If
        Next
    End If
    get_buy_solutions = -1
End Function

Public Function get_solutions(MapName As String) As Integer
On Error GoTo errie
    Dim i As Integer
    If MapRoute(0).Src.Name <> "" Then
        For i = 0 To UBound(MapRoute)
            If MapName = MapRoute(i).Src.Name Then
                get_solutions = i
                Exit Function
            End If
        Next
    End If
errie:
    Err.Clear
    get_solutions = -1
End Function


Public Function Map_Routing(start As String, Destination As String, Optional Sx As Long = 0, Optional Sy As Long = 0, Optional Dx As Long = 0, Optional Dy As Long = 0, Optional UseWarp As Boolean = False) As Byte
On Error GoTo errie
    Dim retcode As Byte
    Dim aiMapRoute As clsMapRoute
    Set aiMapRoute = New clsMapRoute

    IsRouting = True
    ReDim MapRoute(0)
    retcode = aiMapRoute.Map_Search(start, Destination, map_time_limit, Sx, Sy, Dx, Dy, UseWarp)
    
    Select Case retcode
        Case 0
            Dim i As Integer
            Dim continue As Boolean
            i = 0
            Do
                If MapRoute(0).Src.Name <> "" Then ReDim Preserve MapRoute(UBound(MapRoute) + 1)
                continue = aiMapRoute.MapStepNext(i, MapRoute(UBound(MapRoute)).Src.Name, MapRoute(UBound(MapRoute)).Src.Pos.X, MapRoute(UBound(MapRoute)).Src.Pos.Y, MapRoute(UBound(MapRoute)).Des.Name, MapRoute(UBound(MapRoute)).Des.Pos.X, MapRoute(UBound(MapRoute)).Des.Pos.Y)
                i = i + 1
            Loop Until Not continue
            If Sx <> 0 And Sy <> 0 Then
                MapRoute(0).Src.Pos.X = Sx
                MapRoute(0).Src.Pos.Y = Sy
            End If
            'Stat "Successfull Map Routing!," & CStr(UBound(MapRoute) + 1) & " Map(s) to [" & Destination & "]" & vbCrLf
        Case 1
            'Stat "Warning: MAP Routing time-out" & vbCrLf
        Case 2
            'Stat "Warning: MAP Routing problem" & vbCrLf
    End Select
    IsRouting = False
    Map_Routing = retcode
    Exit Function
errie:
    IsRouting = False
    Map_Routing = 2
End Function

Public Function Check_IsManyPortal(Name As String, Name2 As String) As Boolean
    On Error GoTo errie
    Dim tstr As String
    Dim Index As Integer
    Dim i As Integer
    ReDim myInPortals(0)
    Dim count As Byte
    Open App.Path & "\maproute\portals.txt" For Input As #10
    Do While Not EOF(10)
        Line Input #10, tstr
        tstr = Trim(tstr)
        Index = InStr(tstr, Name & " ")
        If Index = 0 Or Index > 1 Then GoTo end_loop
        Index = InStr(Index + 1, tstr, Name2 & " ")
        If Index = 0 Then GoTo end_loop
        count = count + 1
        If count > 1 Then
            Check_IsManyPortal = True
            Close 10
            Exit Function
        End If
end_loop:
    Loop
    Close 10
    Check_IsManyPortal = False
    Exit Function
errie:
    Close 10
    Check_IsManyPortal = False
    MsgBox "Error!!! on loading 'maproute\portals.txt' (Check_IsManyPortal) : " & Err.Description, vbCritical
End Function

Public Function Check_MapInPortal(Name As String) As Boolean
    On Error GoTo errie
    Dim tstr As String
    Dim Index As Integer
    Dim i As Integer
    ReDim myInPortals(0)
    Dim DesName As String
    Open App.Path & "\maproute\portals.txt" For Input As #10
    Do While Not EOF(10)
        Line Input #10, tstr
        tstr = Trim(tstr)
        Index = InStr(tstr, Name & " ")
        If Index = 0 Or Index > 1 Then GoTo end_loop
        Index = InStr(Index + 1, tstr, Name & " ")
        If Index = 0 Then GoTo end_loop
        Check_MapInPortal = True
        Close 10
        Exit Function
end_loop:
    Loop
    Close 10
    Check_MapInPortal = False
    Exit Function
errie:
    Close 10
    Check_MapInPortal = False
    MsgBox "Error!!! on loading 'maproute\portals.txt' (Check_MapInPortal) : " & Err.Description, vbCritical
End Function


Public Sub Load_MapInPortal(Mode As Byte, Name As String, ByRef myInPortals() As MyinPortal)
    On Error GoTo errie
    Dim tstr As String
    Dim Index As Integer, LCount As Long
    Dim i As Integer
    ReDim myInPortals(0)
    LCount = 0
    Open App.Path & "\maproute\portals.txt" For Input As #10
    Do While Not EOF(10)
        Line Input #10, tstr
        tstr = Trim(tstr)
        LCount = LCount + 1
        Select Case Mode
            Case 0
                Index = InStr(tstr, Name & " ")
                If (Index = 0) Then GoTo end_loop
                Index = InStr(Index + 1, tstr, Name)
                If (Index = 0) Then GoTo end_loop
            Case 1
                Index = InStr(tstr, Name & " ")
                If (Index <= 1) Then GoTo end_loop
            Case 2
                Index = InStr(tstr, Name)
                If (Index = 0 Or Index > 1) Then GoTo end_loop
                Index = InStr(Index + 1, tstr, Name & " ")
                If (Index > 0) Then GoTo end_loop
        End Select
start_portal:
        If myInPortals(0).Des.Name <> "" Then ReDim Preserve myInPortals(UBound(myInPortals) + 1)
        For i = 1 To 6
            Index = InStr(tstr, " ")
            Select Case i
                Case 1
                    myInPortals(UBound(myInPortals)).Src.Name = Left(tstr, Index - 1)
                Case 2
                    myInPortals(UBound(myInPortals)).Src.Pos.X = Val(Left(tstr, Index - 1))
                Case 3
                    myInPortals(UBound(myInPortals)).Src.Pos.Y = Val(Left(tstr, Index - 1))
                Case 4
                    myInPortals(UBound(myInPortals)).Des.Name = Left(tstr, Index - 1)
                Case 5
                    myInPortals(UBound(myInPortals)).Des.Pos.X = Val(Left(tstr, Index - 1))
                Case 6
                    myInPortals(UBound(myInPortals)).Des.Pos.Y = Val(tstr)
            End Select
            If i < 6 Then
                tstr = Trim(Right(tstr, Len(tstr) - Index))
            End If
        Next
end_loop:
    Loop
    Close 10
    Exit Sub
errie:
    Close 10
    MsgBox "Error!!! on loading 'maproute\portals.txt' (Load_MapInPortal) Line:" & LCount & vbCrLf & vbCrLf & Err.Description, vbCritical
End Sub

Public Sub Load_Field2(field As String, ByRef tmpMap() As Byte, width As Integer, myheight As Integer)
On Error GoTo errie
    Dim lfile As Long
    Dim test() As Byte
    TILE_SIDE = 1
    lfile = FreeFile
    'MapPath = "C:\REVMEU\gat"
    Open MapPath & "\" & field & ".gat" For Binary Access Read As lfile
    Get lfile, , width
    Get lfile, , myheight
    ReDim tmpMap(width - 1, myheight - 1)
    ReDim test(width - 2, myheight - 2) As Byte
    Get lfile, , tmpMap()
    Close lfile
    Dim i As Integer
    Dim j As Integer
    For i = 0 To UBound(tmpMap, 1) - 1
        For j = 0 To UBound(tmpMap, 2) - 1
            If tmpMap(i, j) = &HFF Then
                test(i, UBound(tmpMap, 2) - j - 1) = 0
            ElseIf tmpMap(i, j) = &H80 Then
                test(i, UBound(tmpMap, 2) - j - 1) = 5
            ElseIf tmpMap(i, j) = &H0 Then
                test(i, UBound(tmpMap, 2) - j - 1) = 1
            End If
        Next
    Next
    ReDim tmpMap(width - 2, myheight - 2)
    tmpMap() = test()
    Exit Sub
errie:
    Close lfile
    MsgBox "Error!!! on loading " & MapPath & "\" & field & ".gat", vbCritical
    Unload MDIfrmMain
End Sub

Public Sub Extend_Map_Routing(Mode As Byte, Name As String, inPt1 As Coord, ByRef myDestination As MyPortal, Optional Name2 As String = "", Optional i As Integer = -1)
    Dim pt1 As Coord
    Dim Pt2 As Coord
    Dim Pt3 As Coord
    Dim Pt4 As Coord
    Dim width As Integer
    Dim height As Integer
    Dim w As Integer
    Dim h As Integer
    Dim tmpindex As Integer
    Dim myAIMap() As Byte
    Dim tmpMap() As Byte
    Dim tmpPortal() As MyinPortal
    'Load_Field MapName
    tmpindex = 0
    Load_Field2 Name, myAIMap(), width, height
    Load_MapInPortal 1, Name, tmpPortal()
    If i > 0 Then
        tmpindex = i
        Load_Field2 Name2, tmpMap(), w, h
        Pt3.X = MapRoute(tmpindex - 1).Des.Pos.X
        Pt3.Y = h - MapRoute(tmpindex - 1).Des.Pos.Y
    End If
    pt1.X = inPt1.Y
    pt1.Y = height - inPt1.X
    Dim Index As Integer
    Dim found As Boolean
    Dim AllCanGo As Boolean
    found = False
    AllCanGo = True
    Dim dis As Integer
    dis = 500
    For i = 0 To UBound(tmpPortal)
        Pt2.X = tmpPortal(i).Des.Pos.X
        Pt2.Y = height - tmpPortal(i).Des.Pos.Y
        If (Name2 = "" Or Name2 = tmpPortal(i).Src.Name) Then
            'MsgBox tmpPortal(i).Des.pos.x & ":" & CStr(tmpPortal(i).Des.pos.y) & " Routing OK!", vbOKOnly
            If Pt3.X <> 0 And Pt3.Y <> 0 Then
                Pt4.X = tmpPortal(i).Src.Pos.X
                Pt4.Y = h - tmpPortal(i).Src.Pos.Y
                If Not Routing(Pt3, Pt4, Name2, tmpMap) Then
                    AllCanGo = False
                    GoTo end_loop
                Else
                    MapRoute(tmpindex).Src.Pos.X = tmpPortal(i).Src.Pos.X
                    MapRoute(tmpindex).Src.Pos.Y = tmpPortal(i).Src.Pos.Y
                End If
            End If
            If Routing(pt1, Pt2, Name, myAIMap) Then
                    If EvalNorm(pt1, Pt2) < dis Then
                        Index = i
                    End If
                    dis = EvalNorm(pt1, Pt2)
                    found = True
            Else
                AllCanGo = False
            End If
        End If
end_loop:
    Next
    If found And Not AllCanGo Then
        If Mode = 2 Then
            myDestination = tmpPortal(Index).Des
        ElseIf Mode = 1 Then
            myDestination = tmpPortal(Index).Src
        End If
    Else
        myDestination.Name = ""
        myDestination.Pos.X = 0
        myDestination.Pos.Y = 0
    End If
End Sub

Public Function Routing(pStart As Coord, pEnd As Coord, Name As String, Map() As Byte) As Boolean      'Start Routing
    On Error GoTo errie
    Dim newTree As clsTree
    Set newTree = New clsTree
    newTree.bAStar = False
    Dim lNodeX As Long, lNodeY As Long
    Dim bShowPath As Boolean, bSlowDown As Boolean
    Dim bAchieved As Boolean
    
    'If objTree Is Nothing Then Exit Function
    'pStart.y = 21
    newTree.bAStar = False
    bAchieved = newTree.RunSearch(pStart.X, pStart.Y, pEnd.X, pEnd.Y, Map(), Name)
    'newTree.Load_Portal Name
        
    'Do Until newTree.NextNode Or bAchieved
    '    bAchieved = newTree.UpdateCurrentNode
    '    DoEvents
    'Loop
    'newTree.BackTracePath
    
    If bAchieved Then
        Routing = True
    Else
        Routing = False
    End If
errie:
    Set newTree = Nothing
End Function


