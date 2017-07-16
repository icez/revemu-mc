Attribute VB_Name = "md_Decode"
Option Explicit


Function Decode_0069(inData As String) As String 'Serverlist
On Error GoTo errie
'Text1.Text = Mid(InData, 2, 10)
        Dim i As Integer
        Dim tstr As String
        Dim X As Integer
        Dim p0069 As px0069
        Dim pex As px0069ex
        Dim bArr() As Byte
        bArr = Conv2Arr(Mid(inData, 1, 47))
        CopyMemory p0069, bArr(0), 47
        
        SessionID = Mid(inData, 5, 4)
        AccountID = Mid(inData, 9, 4)
        Sex = Asc(Mid(inData, 47, 1))

        ConnState = 2
        If Not isUseHaunted Then frmMain.Winsock1.Close
        Stat "Usable Account!" + vbCrLf
        frmServer.LstServer.Clear
        
        'X = Int((p0069.Length - 47) \ chopserver) - 1
        'print_packet InData
        'ReDim ServerList(X)
        
        ReDim ServerList(0)
        For X = 48 To Len(inData) Step 32
            bArr = Conv2Arr(Mid(inData, X, 32))
            CopyMemory pex, bArr(0), 32
            With ServerList(UBound(ServerList))
                .IP = MakeIP(pex.IP)
                .Name = MakeString(pex.Name)
                .number = pex.Players
                .Port = pex.Port
            End With
            ReDim Preserve ServerList(UBound(ServerList) + 1)
        Next
        If UBound(ServerList) > 0 Then ReDim Preserve ServerList(UBound(ServerList) - 1)
        For X = 0 To UBound(ServerList)
            frmServer.LstServer.AddItem (ServerList(X).Name) & "- (" & CStr(ServerList(X).number) & " Users)"
        Next
        inData = ""
        If Not isUseHaunted Then
            If IsConnected And NumServ <= UBound(ServerList) Then
                    Stat "CSrv:Connecting to " & ServerList(NumServ).IP & ":" & CStr(ServerList(NumServ).Port) & "...."
                    DoConnect ServerList(NumServ).IP, CLng(ServerList(NumServ).Port)
                    CurCIP = ServerList(NumServ).IP
            Else
                frmServer.LstServer.Selected(0) = True
                frmMain.tmrResponse.Enabled = False
                frmServer.Visible = True
            End If
        End If
Decode_0069 = ""
Exit Function
errie:
Decode_0069 = "ERROR!!! [Decode_0069] " & Err.Description
Err.Clear
End Function

Function Decode_006A(inData As String) As String
On Error GoTo errie
    If (ConnState < 2) Then
        Dim ReConRes As Boolean
        ReConRes = True
            Select Case Asc(Mid(inData, 3, 1))
                Case 0
                    Stat "Sorry we can't find your ID.Please check again..."
                    ReConRes = False
                Case 1
                    Stat "Incorrect Password.Please check again..."
                    ReConRes = False
                Case 2
                    Stat "Sorry, this ID was expired..."
                Case 3
                    Stat "Server connection denied..."
                Case 4
                    Stat "No E-Mail Certification for this ID!..."
                    ReConRes = False
                Case 5
                    Stat "Your client is not the latest version. Checking [" & (MasterSelect.code + 1) & "]...", MColor.Fail
                    MasterSelect.code = MasterSelect.code + 1
                Case 6
                    Stat "You've been disconnect by GM!!" + vbCrLf, vbRed
                    Open App.Path & "\warning.txt" For Append As #8
                        Chat Date & "@" & Time & ": You've been disconnect by [GM], Closed program..."
                        Print #8, Date & "@" & Time & ": You've been disconnect by [GM], Closed program..."
                    Close 8
                    End
            End Select
            If ReConRes Then frmMain.ResettoReCon Else frmLogin.Show: frmMain.Winsock1.Close
            Exit Function
        End If
Decode_006A = ""
Exit Function
errie:
Decode_006A = "ERROR!!! [Decode_006A] " & Err.Description
Err.Clear
frmMain.ResettoReCon
End Function

Function Decode_006B(inData As String) As String  'Character Data
On Error GoTo errie
        Dim i As Integer
        Dim X As Integer
        Dim start As Integer
        Dim pex As px006Bex, bArr() As Byte

        frmCharSelect.List1.Clear
        ReDim Players(0)
        'print_packet InData
        i = 0
        start = 5
        If MasterSelect.start > 0 Then start = MasterSelect.start
        'If ServerID = 3 Or ServerID = 0 Then start = 25
        For X = start To Len(inData) - 105 Step 106
            bArr = Conv2Arr(Mid(inData, X, 106))
            CopyMemory pex, bArr(0), 106
            If pex.Index > UBound(Players) Then ReDim Preserve Players(pex.Index)
            Players(pex.Index).Name = MakeString(pex.Name)
            Players(pex.Index).Class = Return_Class(CLng(pex.jobID))
            Players(pex.Index).ClassID = pex.jobID
            Players(pex.Index).BaseLV = pex.levelBase
            Players(pex.Index).HP = pex.HP
            Players(pex.Index).MaxHP = pex.HPmax
            Players(pex.Index).SP = pex.SP
            Players(pex.Index).maxsp = pex.SPmax
            MovementSpeed = pex.walkSpeed
            Players(pex.Index).JobLV = pex.levelJOB
            Players(pex.Index).BaseExp = pex.expBASE
            Players(pex.Index).StatPoint = pex.StatusPoint
            Players(pex.Index).JobExp = pex.expJOB
            'oldJobEXP = Players(i).JobExp
            Players(pex.Index).Zeny = pex.Zeny
            'Players(i).StatPoint = MakePort(Mid(InData, X + 44, 2))
            frmCharSelect.List1.AddItem CStr(pex.Index) + " : " + MakeString(pex.Name) & " [" & Players(pex.Index).Class & "]"
        Next
        'ReDim Preserve Players(UBound(Players) - 1)

        Stat "Got Character ID... " + vbCrLf
        frmMain.tmrResponse.Enabled = False
        inData = ""
        'frmPlayer.Visible = True
        
        If isUseHaunted Then
            frmCharSelect.Visible = False
        Else
            If IsConnected And CharIdStart <= UBound(Players) Then
                If Players(CharIdStart).Name <> "" Then
                    frmMain.tmrResetResponse
                    number = CharIdStart
                    GetScriptLockmap
                    CharNameStart = Players(number).Name
                    MDIfrmMain.Caption = Players(number).Name & " - Powered by " & Version
                    MDIfrmMain.CreatIcon Players(number).Name & " - Powered by " & Version
                    frmPlayer.labBaseLv.Caption = CStr(Players(number).BaseLV)
                    frmPlayer.labJobLv.Caption = CStr(Players(number).JobLV)
                    frmPlayer.labPlayerName.Caption = Players(number).Name
                    frmPlayer.labSP.Caption = CStr(Players(number).SP) + "  /  " + CStr(Players(number).maxsp)
                    frmPlayer.labSP.Caption = CStr(Players(number).SP) + "  /  " + CStr(Players(number).maxsp)
                    frmPlayer.tabSP.width = (Players(number).SP / Players(number).maxsp) * (frmPlayer.tabSPbg.width - 20)
                    frmPlayer.LabHP.Caption = CStr(Players(number).HP) + "  /  " + CStr(Players(number).MaxHP)
                    frmPlayer.LabHP.Caption = CStr(Players(number).HP) + "  /  " + CStr(Players(number).MaxHP)
                    frmPlayer.tabHP.width = (Players(number).HP / Players(number).MaxHP) * (frmPlayer.tabHPBg.width - 20)
                    frmPlayer.labZeny.Caption = Format(Players(number).Zeny, "##,##")
                    frmStat.labStatPt.Caption = CStr(Players(number).StatPoint)
                    frmPlayer.labClass.Caption = Players(number).Class
                    MageMode = False
                    Range = 2
                    With frmAttackOption
                        .imgUseWeapon.Visible = True
                        .LabUseWeapon.ForeColor = 0
                    End With
                    If isClassMage Then
                        MageMode = True
                        Range = 8
                        MDIfrmMain.mnuWeapon.Visible = True
                    End If
                    If isClassAco Then
                        UseHeal = True
                        With FrmHPSPOption
                            .imgHeal.Visible = True
                            .LabHeal.ForeColor = 0
                        End With
                    End If
                    If isClassArcher Then
                            Range = 6
                    End If
                'If Not Connected Then
                    oldBaseEXP = Players(number).BaseExp
                    oldJobEXP = Players(number).JobExp
                'End If
                    Check_JobBar
                    If Players(number).BaseLV = 99 Then
                        frmPlayer.labtabBaseEXPBg.Visible = False
                        frmPlayer.tabBaseEXP.Visible = False
                    Else
                        frmPlayer.labtabBaseEXPBg.Visible = True
                        frmPlayer.tabBaseEXP.Visible = True
                    End If
                    Winsock_SendPacket IntToChr(&H66) & Chr(CharIdStart), True
                    Stat "Verify Character ID" + vbCrLf
                Else
                    frmCharSelect.Visible = True
                End If
            Else
                frmCharSelect.Visible = True
            End If
        End If
        Decode_006B = ""
        Exit Function
errie:
Decode_006B = "ERROR!!! [Decode_006B] " & Err.Description
Err.Clear
End Function

Function Decode_006C() As String
On Error GoTo errie
    Stat "Error Login to Game Server..." & vbCrLf
    frmMain.ResettoReCon
Decode_006C = ""
Exit Function
errie:
Decode_006C = "ERROR!!! [Decode_006C] " & Err.Description
frmMain.ResettoReCon
Err.Clear
End Function

Function Decode_0071(inData As String) As String
On Error GoTo errie
    Dim mapgot$
    CharID = Mid(inData, 3, 4)
    If Not isUseHaunted Then frmMain.Winsock1.Close
    Stat "Got map server IP " & IIf(WaitCheat = 2, "fake map!! ", "") & "[loading data "
    
    CanusePath = True
    mapgot = MakeString(Mid(inData, 7, 16))
    
    If (StartMap = "") Then StartMap = mapgot
    CurrentMap = mapgot
    OldMap = mapgot
    
    If WaitCheat = 2 And Not isUseHaunted Then mapgot = WaitCMap Else mapgot = Left(mapgot, Len(mapgot) - 4)
    MapName = mapgot
    Stat ". ", vbRed, False, True
    Load_WayPoint mapgot
    Stat "/ ", vbRed, False, True
    Load_Field mapgot
    Stat ". ", vbRed, False, True
    If Not IsInLock Then MoveOnly = True Else MoveOnly = False
    frmMain.Label1.Caption = "Main Status - " & GetMapname(MapName)
    ConnState = 3
    If WaitCheat = 2 And Not isUseHaunted Then
        Dim i&, j&
        For i = 0 To UBound(IPList)
            If IPList(i).CIP = CurCIP Then
                For j = 0 To UBound(IPList(i).MList)
                    If WaitCMap = IPList(i).MList(j).MName Then
                        CurMIP = IPList(i).MList(j).MIP
                        CurMPort = CLng(IPList(i).MList(j).MPort)
                        WaitCheat = 3
                        Exit For
                    End If
                Next
                Exit For
            End If
        Next
    Else
        CurMIP = MakeIP(Mid(inData, 23, 4))
        CurMPort = MakePort(Mid(inData, 27, 2))
    End If
    Stat "/ ", vbRed, False, True
    Stat "success!]" & vbCrLf & "Connecting to " & CurMIP & ":" & CurMPort & " ..." + vbCrLf
    UpdateMIP MapName, CurMIP, CurMPort
    If Not isUseHaunted Then DoConnect CurMIP, CLng(CurMPort)
    Decode_0071 = ""
    Exit Function
errie:
Decode_0071 = "ERROR!!! [Decode_0071] " & Err.Description
Err.Clear
frmMain.ResettoReCon
End Function

Function Decode_0073(inData As String) As String
On Error GoTo errie
    MoveWait = False
        Clear_Dot curPos
        PlayerMoveTime = 0
        BlockMove = False
        curPos = MakeCoords(Mid(inData, 7, 3))
        Clear_Dot DisPos
        DisPos = curPos
        OldPos = curPos
        CheckEvent "OnMapChange", "oldMap=" & Chr(0) & "oldposX=0" & Chr(0) & "oldposY=0" & Chr(0) & "newMap=" & MapName & Chr(0) & "newposX=" & curPos.Y & Chr(0) & "newposY=" & curPos.X & Chr(0) & "pktHeader=0073"
        If Make_Start_Point(curPos) Then
            If EvalNorm(curPos, WayPoint(StartPoint)) = 0 Then
                Stat "You're on waypoint at (" & CStr(WayPoint(StartPoint).Y) & ":" & CStr(WayPoint(StartPoint).X) & ")" & vbCrLf
            Else
                Stat CStr(EvalNorm(curPos, WayPoint(StartPoint))) & " Block(s) far from closest waypoint at (" & CStr(WayPoint(StartPoint).Y) & ":" & CStr(WayPoint(StartPoint).X) & ")" & vbCrLf
            End If
        End If
        frmMain.Label14.Caption = curPos.X
        frmMain.Label12.Caption = curPos.Y
        'If frmMap.Visible Then frmMap.Refresh_MAP CurPos.X, CurPos.Y
        If FrmField.Visible Then Plot_Dot curPos, vbBlue
        If StartPos.X = 0 Then StartPos = curPos
        upd_curMonster
        If CurrentItem.Name <> "" Then
            frmMain.labCurMons.Caption = "[" + Return_ItemName(CurrentItem.Name) + "], " _
            & CStr(EvalNorm(CurrentItem.Pos, curPos)) + " Blocks"
            SendPickup
        End If
        If (CurAtkMonster.NameID > 0) Then frmMain.SendAttack
    
        frmMain.tmrTicks.Enabled = True
        Decode_0073 = ""
        Exit Function
errie:
    Decode_0073 = "ERROR!!! [Decode_0073] " & Err.Description
    Err.Clear
End Function

Function Decode_007C(inData As String) As String
On Error GoTo errie
    Dim tmpcoord As Coord
    Dim Monstmppos As Coord
    Dim IsPet As Boolean
    Dim NameID As Integer
    Dim X As Integer
    IsPet = False
    NameID = MakePort(Mid(inData, 21, 2))
    Monstmppos = MakeCoordsSec(Mid(inData, 37, 3))
    tmpcoord = Monstmppos
        If MakePort(Mid(inData, 21, 2)) < &H389 Then
            AI_AvoidID Mid(inData, 3, 4)
            If UBound(People) > 0 Then
            For X = 0 To UBound(People) - 1
                If Mid(inData, 3, 4) = People(X).ID Then
                    If Not Disable_frmPeople Then Clear_Dot People(X).Pos
                    People(X).Pos = tmpcoord
                    People(X).NextPos = Monstmppos
                    People(X).Time = GetTickCount
                    People(X).Speed = MakePort(Mid(inData, 7, 2))
                    If Not Disable_frmPeople Then Plot_Dot People(X).Pos, PColor
                    'If frmPeople.Visible = True Then
                    UpdatePeople
                    'End If
                    If IsAvoidID(Mid(inData, 3, 4)) Then
                        CheckEvent "OnGMMove", "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0) & "job=" & People(X).Class & Chr(0) & "startX=" & People(X).Pos.Y & Chr(0) & "startY=" & People(X).Pos.X & Chr(0) & "endX=" & People(X).NextPos.Y & Chr(0) & "endY=" & People(X).NextPos.X & Chr(0) & "AID=" & MakePort(Mid(inData, 3, 4))
                    ElseIf IsAvoid(People(X).Name) And Len(People(X).Name) > 0 Then
                        CheckEvent "OnAvoidListMove", "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0) & "job=" & People(X).Class & Chr(0) & "startX=" & People(X).Pos.Y & Chr(0) & "startY=" & People(X).Pos.X & Chr(0) & "endX=" & People(X).NextPos.Y & Chr(0) & "endY=" & People(X).NextPos.X & Chr(0) & "AID=" & MakePort(Mid(inData, 3, 4))
                    ElseIf isWarpList(People(X).Name) And Len(People(X).Name) > 0 Then
                        CheckEvent "OnWarpListMove", "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0) & "job=" & People(X).Class & Chr(0) & "startX=" & People(X).Pos.Y & Chr(0) & "startY=" & People(X).Pos.X & Chr(0) & "endX=" & People(X).NextPos.Y & Chr(0) & "endY=" & People(X).NextPos.X & Chr(0) & "AID=" & MakePort(Mid(inData, 3, 4))
                    Else
                        CheckEvent "OnPlayerMove", "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0) & "job=" & People(X).Class & Chr(0) & "startX=" & People(X).Pos.Y & Chr(0) & "startY=" & People(X).Pos.X & Chr(0) & "endX=" & People(X).NextPos.Y & Chr(0) & "endY=" & People(X).NextPos.X & Chr(0) & "AID=" & MakePort(Mid(inData, 3, 4))
                    End If
                    GoTo skip4
                End If
            Next
            End If
                People(UBound(People)).ID = Mid(inData, 3, 4)
                People(UBound(People)).Pos = tmpcoord
                People(UBound(People)).NextPos = Monstmppos
                People(UBound(People)).Speed = MakePort(Mid(inData, 7, 2))
                People(UBound(People)).Time = GetTickCount
                If Not Disable_frmPeople Then Plot_Dot People(UBound(People)).Pos, PColor
                People(UBound(People)).NameID = MakePort(Mid(inData, 21, 2))
                'print_packet Left(InData, 41), People(UBound(People)).nameid
                If Asc(Mid(inData, 50, 1)) = 0 Then
                    People(UBound(People)).Sex = " <F>"
                Else
                    People(UBound(People)).Sex = " <M>"
                End If
                People(UBound(People)).Class = Return_Class(People(UBound(People)).NameID)
                If IsAvoidID(Mid(inData, 3, 4)) Then
                    CheckEvent "OnGMMove", "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0) & "job=" & People(UBound(People)).Class & Chr(0) & "startX=" & People(UBound(People)).Pos.Y & Chr(0) & "startY=" & People(UBound(People)).Pos.X & Chr(0) & "endX=" & People(UBound(People)).NextPos.Y & Chr(0) & "endY=" & People(UBound(People)).NextPos.X & Chr(0) & "AID=" & MakePort(Mid(inData, 3, 4))
                ElseIf IsAvoid(People(UBound(People)).Name) And Len(People(UBound(People)).Name) > 0 Then
                    CheckEvent "OnAvoidListMove", "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0) & "job=" & People(UBound(People)).Class & Chr(0) & "startX=" & People(UBound(People)).Pos.Y & Chr(0) & "startY=" & People(UBound(People)).Pos.X & Chr(0) & "endX=" & People(UBound(People)).NextPos.Y & Chr(0) & "endY=" & People(UBound(People)).NextPos.X & Chr(0) & "AID=" & MakePort(Mid(inData, 3, 4))
                ElseIf isWarpList(People(UBound(People)).Name) And Len(People(UBound(People)).Name) > 0 Then
                    CheckEvent "OnWarpListMove", "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0) & "job=" & People(UBound(People)).Class & Chr(0) & "startX=" & People(UBound(People)).Pos.Y & Chr(0) & "startY=" & People(UBound(People)).Pos.X & Chr(0) & "endX=" & People(UBound(People)).NextPos.Y & Chr(0) & "endY=" & People(UBound(People)).NextPos.X & Chr(0) & "AID=" & MakePort(Mid(inData, 3, 4))
                Else
                    CheckEvent "OnPlayerMove", "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0) & "job=" & People(UBound(People)).Class & Chr(0) & "startX=" & People(UBound(People)).Pos.Y & Chr(0) & "startY=" & People(UBound(People)).Pos.X & Chr(0) & "endX=" & People(UBound(People)).NextPos.Y & Chr(0) & "endY=" & People(UBound(People)).NextPos.X & Chr(0) & "AID=" & MakePort(Mid(inData, 3, 4))
                End If
                'CharHair
                Winsock_SendPacket IntToChr(&H64 + &H30) + Mid(inData, 3, 4), True
                ReDim Preserve People(UBound(People) + 1)
                UpdatePeople
        End If
        Dim found As Boolean
        Dim found2 As Boolean
        Dim islist As Boolean
        Dim tmpname As String
skip4:
        found = False
        islist = False
        IsPet = False
        'If Asc(Mid(InData, 23, 1)) > 0 Then IsPet = True
        For X = 0 To UBound(Attack)
            If MakePort(Mid(inData, 21, 2)) = Attack(X).ID Then
                found = True
                Exit For
            End If
        Next
        For X = 0 To UBound(Monsters)
            If MakePort(Mid(inData, 21, 2)) = Monsters(X).ID Then
                islist = True
                tmpname = Monsters(X).Name
                Exit For
            End If
        Next
        'If islist Then
        If Mid(inData, 3, 4) = MyPet.ID And MyPet.Type = "" Then MyPet.Type = tmpname
            found2 = False
            If UBound(MonsterList) > 0 Then
            For X = 0 To UBound(MonsterList) - 1
                If Mid(inData, 3, 4) = MonsterList(X).ID Then
                    found2 = True
                    If Not found Then MonsterList(X).NoAttack = True
                    'If IsPet Then MonsterList(X).IsPet = True
                    Clear_Dot MonsterList(X).Pos
                    MonsterList(X).Pos = MakeCoords(Mid(inData, 37, 3))
                    MonsterList(X).IsAttack = False
                    'MonsterList(X).IsPet = True
                    IsPet = MonsterList(X).IsPet
                    If MyPet.ID = MonsterList(X).ID Then
                        Plot_Dot MonsterList(X).Pos, 16711935
                    ElseIf CurAtkMonster.ID <> MonsterList(X).ID Then
                        Plot_Dot MonsterList(X).Pos, vbRed
                    Else
                        Plot_Dot MonsterList(X).Pos, CurAtkColor
                    End If
                    MonsterList(X).NextPos = Monstmppos
                    MonsterList(X).StatusA = MakePort(Mid(inData, 9, 2))
                    MonsterList(X).StatusB = MakePort(Mid(inData, 11, 2))
                    MonsterList(X).Speed = MakePort(Mid(inData, 7, 2))
                    If (MakePort(Mid(inData, 9, 2)) > 0 And MakePort(Mid(inData, 9, 2)) < 5) Or MakePort(Mid(inData, 11, 2)) > 0 Then MonsterList(X).IsTrap = True
                    MonsterList(X).Time = GetTickCount()
                    MonsterList(X).Endtime = MonsterList(X).Time + (EvalNorm(tmpcoord, Monstmppos) * MonsterList(X).Speed)
                    If CanGO(curPos, MonsterList(X).Pos) Then MonsterList(X).CantGo = True
                    CheckEvent "OnMonsterMove", "name=" & MonsterList(X).Name & Chr(0) & "startX=" & MonsterList(X).Pos.Y & Chr(0) & "startY=" & MonsterList(X).Pos.X & Chr(0) & "endX=" & MonsterList(X).NextPos.Y & Chr(0) & "endY=" & MonsterList(X).NextPos.X
                    'If (CurAtkMonster.ID = Mid(InData, 3, 4)) Then CurAtkMonster = MonsterList(UBound(MonsterList))
                    upd_frmMonster
                    Exit For
                End If
            Next
            End If
                If Mid(inData, 3, 4) = CurAtkMonster.ID Then
                    CurAtkMonster.Pos = MakeCoords(Mid(inData, 37, 3))
                    CurAtkMonster.NextPos = Monstmppos
                    CurAtkMonster.Speed = MakePort(Mid(inData, 7, 2))
                    CurAtkMonster.Time = GetTickCount()
                    CurAtkMonster.Endtime = CurAtkMonster.Time + (EvalNorm(tmpcoord, Monstmppos) * CurAtkMonster.Speed)
                    'TmrMonsMove.Interval = EvalNorm(tmpcoord, Monstmppos) * MakePort(Mid(InData, 7, 2))
                    'TmrMonsMove.Enabled = True
                    upd_curMonster
                End If
           
            If (Not Sitting) And (Not Pickup) And (Not InFight) And (CurAtkMonster.NameID > 0) Then SendAction = True
            If Not found2 Then
                'If (UBound(MonsterList) < 5) Then
                If Not found Then MonsterList(UBound(MonsterList)).NoAttack = True
                'If IsPet Then MonsterList(UBound(MonsterList)).IsPet = True
                MonsterList(UBound(MonsterList)).ID = Mid(inData, 3, 4)
                MonsterList(UBound(MonsterList)).Pos = MakeCoords(Mid(inData, 37, 3))
                MonsterList(UBound(MonsterList)).IsAttack = False
                'MonsterList(UBound(MonsterList)).IsPet = True
                If MyPet.ID = MonsterList(UBound(MonsterList)).ID Then
                    Plot_Dot MonsterList(X).Pos, 16711935
                Else
                    Plot_Dot MonsterList(UBound(MonsterList)).Pos, vbRed
                End If
                MonsterList(UBound(MonsterList)).NextPos = Monstmppos
                MonsterList(UBound(MonsterList)).Time = EvalNorm(tmpcoord, Monstmppos) * MakePort(Mid(inData, 7, 2))
                MonsterList(UBound(MonsterList)).NameID = MakePort(Mid(inData, 21, 2))
                MonsterList(UBound(MonsterList)).Speed = MakePort(Mid(inData, 7, 2))
                MonsterList(UBound(MonsterList)).StatusA = MakePort(Mid(inData, 9, 2))
                'MonsterList(UBound(MonsterList)).StatusB = MakePort(Mid(InData, 11, 2))
                MonsterList(UBound(MonsterList)).StatusB = 10
                If (MakePort(Mid(inData, 9, 2)) > 0 And MakePort(Mid(inData, 9, 2)) < 5) Or MakePort(Mid(inData, 11, 2)) > 0 Then MonsterList(UBound(MonsterList)).IsTrap = True
                MonsterList(UBound(MonsterList)).IsAttack = False
                If islist Then MonsterList(UBound(MonsterList)).Name = tmpname
                X = UBound(MonsterList)
                CheckEvent "OnMonsterMove", "name=" & MonsterList(X).Name & Chr(0) & "startX=" & MonsterList(X).Pos.Y & Chr(0) & "startY=" & MonsterList(X).Pos.X & Chr(0) & "endX=" & MonsterList(X).NextPos.Y & Chr(0) & "endY=" & MonsterList(X).NextPos.X
                ReDim Preserve MonsterList(UBound(MonsterList) + 1)
                upd_frmMonster
                'End If
            End If
            'Check_Rest
            If ((CurAtkMonster.NameID = 0) And (Not Sitting) And (Not Pickup)) Then
                If (Not IsAggro) And (UBound(MonsterList) > 0) And (IsAutoKill) And ((Not IsSPWait) Or (Not Sitting)) Then
                    frmMain.EstimateClosestMonster
                End If
            End If
         
        'End If

        If Not islist And NameID > 20 And Not IsPet Then
            Stat "Unknow Monster " & MakeHexName(Mid(inData, 21, 2)) & vbCrLf
            Winsock_SendPacket IntToChr(&H94) & Mid(inData, 3, 4), True
        End If
    Decode_007C = ""
Exit Function
errie:
Decode_007C = "ERROR!!! [Decode_007C] " & Err.Description
Err.Clear
End Function

Function Decode_0080(inData As String) As String
On Error GoTo errie
    Dim dis_code As Byte, s As Long
    Dim X, Y As Integer
        dis_code = Asc(Mid(inData, 7, 1))
        If UBound(NPCList) > 0 Then
            For X = 0 To UBound(NPCList) - 1
                If Mid(inData, 3, 4) = NPCList(X).ID Then
                    Clear_Dot NPCList(X).Pos
                    For Y = X To UBound(NPCList) - 1
                        NPCList(Y) = NPCList(Y + 1)
                    Next
                    ReDim Preserve NPCList(UBound(NPCList) - 1)
                    If frmNPC.Visible = True Then
                        UpdateNPC
                    End If
                    Exit For
                End If
            Next
        End If
        If MyPet.ID = Mid(inData, 3, 4) And MyPet.Name <> "" Then
            Stat "Your pet is gone!..." & vbCrLf
            MyPet.Name = ""
        End If
        If UBound(People) > 0 Then
            For X = 0 To UBound(People) - 1
                If Mid(inData, 3, 4) = People(X).ID Then
                    Clear_Dot People(X).Pos
                    For Y = X To UBound(People) - 1
                        People(Y) = People(Y + 1)
                    Next
                    If IsAvoidID(People(X).ID) Then
                        CheckEvent "OnGMDisappear", "name=" & People(X).Name & Chr(0) & "job=" & People(X).Class & Chr(0) & "posX=" & People(X).Pos.Y & Chr(0) & "posY=" & People(X).Pos.X
                    ElseIf IsAvoid(People(X).Name) And Len(People(X).Name) > 0 Then
                        CheckEvent "OnAvoidListDisappear", "name=" & People(X).Name & Chr(0) & "job=" & People(X).Class & Chr(0) & "posX=" & People(X).Pos.Y & Chr(0) & "posY=" & People(X).Pos.X
                    ElseIf isWarpList(People(X).Name) And Len(People(X).Name) > 0 Then
                        CheckEvent "OnWarpListDisappear", "name=" & People(X).Name & Chr(0) & "job=" & People(X).Class & Chr(0) & "posX=" & People(X).Pos.Y & Chr(0) & "posY=" & People(X).Pos.X
                    Else
                        CheckEvent "OnPeopleDisappear", "name=" & People(X).Name & Chr(0) & "job=" & People(X).Class & Chr(0) & "posX=" & People(X).Pos.Y & Chr(0) & "posY=" & People(X).Pos.X
                    End If
                    ReDim Preserve People(UBound(People) - 1)
                    'If frmPeople.Visible Then
                    UpdatePeople
                    Exit For
                End If
            Next
        End If
            Dim found As Boolean
            frmMain.Clear_Mon_List Mid(inData, 3, 4)
            found = False
        'end remove
            If Mid(inData, 3, 4) = CurAtkMonster.ID Then
                Clear_Dot CurAtkMonster.Pos
                DeadMonsName = CurMonsterName
                Stat "Your target, [" & CurMonsterName & "]"
                Select Case dis_code
                    Case 0
                        Stat " Disappeared..." & vbCrLf, vbRed
                        If Not MakeDamage Then
                            Clear_This_Mons 1
                        End If
                    Case 1
                        Stat " Dead..." + vbCrLf, vbRed
                        Clear_This_Mons 1
                    Case Else
                        Stat " Teleport... " & vbCrLf, vbRed
                        Clear_This_Mons 1
                End Select
                found = True
            End If
            If UBound(Aggro) > 0 Then
            For X = 0 To UBound(Aggro) - 1
                If Mid(inData, 3, 4) = Aggro(X).ID Then
                    frmMain.tmrPickDelay.Enabled = True
                    If (Not found) Then
                        Select Case dis_code
                            Case 0
                                Stat "Your target, [" + TmpAggroName + "] Disappeared..." & vbCrLf
                            Case 1
                                Stat "Your target, [" + TmpAggroName + "] Dead..." & vbCrLf
                            Case Else
                                Stat "Your target, [" + TmpAggroName + "] Teleport..." & vbCrLf
                        End Select
                    End If
                    For Y = X To UBound(Aggro) - 1
                        Aggro(Y) = Aggro(Y + 1)
                    Next
                    ReDim Preserve Aggro(UBound(Aggro) - 1)
                    Exit For
                End If
            Next
        End If
        If UBound(Aggro) = 0 Then IsAggro = False
        'remove portallist
        If UBound(ExitPortal) > 0 Then
            For X = 0 To UBound(ExitPortal) - 1
                If Mid(inData, 3, 4) = ExitPortal(X).ID Then
                    Stat "Exit Portal(" & MakeHexName(ExitPortal(X).ID) & "), Disappeared..." & vbCrLf
                    For Y = X To UBound(ExitPortal) - 1
                        ExitPortal(Y) = ExitPortal(Y + 1)
                    Next
                    ReDim Preserve ExitPortal(UBound(ExitPortal) - 1)
                    Exit For
                End If
            Next
        End If
        'end remove
        If UBound(ExitPortal) = 0 Then DetectPortal = False
        'frmDebug.txtDebug = "Aggro number = " + CStr(UBound(Aggro))
        
        If Mid(inData, 3, 4) = AccountID Then
            Stat "You're Dead!, waiting..." + vbCrLf
            If ModAI And Mods.ReconWhenDead Then
                frmMain.ResettoReCon
            ElseIf Not ModAI Then
                If DeadRecon Then frmMain.ResettoReCon
            End If
            'Exit Sub
        End If
Decode_0080 = ""
Exit Function
errie:
Decode_0080 = "ERROR!!! [Decode_0080] " & Err.Description
Err.Clear
End Function

Function Decode_0081(inData As String) As String
On Error GoTo errie
        If Asc(Mid(inData, 3, 1)) = 2 Then
                Stat "Someone login to your account!..." & vbCrLf, vbRed
                Open App.Path & "\log\dual-login.log" For Append As #8
                Print #8, ""
                Print #8, Date & "@" & Time & ": Someone login to your account!"
                If MODDC.DualLogin Then
                    If MODDC.DualLoginTime > 0 Then
                        Stat "Delaying to login for " & MODDC.DualLoginTime & " minutes" & vbCrLf
                        Print #8, "Delaying to login for " & MODDC.DualLoginTime & " minutes"
                        MODDelay.DualLogin = MODDC.DualLoginTime * 6000
                        ConnState = 1
                        frmMain.Winsock1.Close
                        Close 8
                        Exit Function
                    Else
                        Print #8, "Close program."
                        End
                    End If
                End If
                
                Close 8
        ElseIf Asc(Mid(inData, 3, 1)) = 6 Then
                Stat "You need to pay for this account!, Disconnect..."
                Open App.Path & "\log\dual-login.log" For Append As #8
                Print #8, ""
                Print #8, Date & "@" & Time & ": You need to pay for this account!, Disconnect..."
                Close 8
                End
        ElseIf Asc(Mid(inData, 3, 1)) = 8 Then
                Stat "Server still recognize your current session!, Disconnect..."
                Open App.Path & "\log\dual-login.log" For Append As #8
                Print #8, ""
                Print #8, Date & "@" & Time & ": Server still recognize your current session!"
                Close 8
        Else
                Stat "Disconnected from server for some reason...[" & CStr(Asc(Mid(inData, 3, 1))) & "]" & vbCrLf
        End If
        frmMain.ResettoReCon
        Decode_0081 = ""
Exit Function
errie:
Decode_0081 = "ERROR!!! [Decode_0081] " & Err.Description
Err.Clear
frmMain.ResettoReCon
End Function

Function Decode_0087(inData As String) As String
On Error GoTo errie
    Dim tmpdistance As Integer
        Dim tmpCurpos As Coord, WalkMsg$
        tmpCurpos = MakeCoords(Mid(inData, 7, 3))
        OldPos = tmpCurpos
        tmpPos = MakeCoordsSec(Mid(inData, 9, 3))
        tmpdistance = EvalNorm(tmpCurpos, tmpPos)
        Clear_Dot curPos
        curPos = tmpCurpos
        If CurAtkMonster.NameID = 0 And (IsRandommove Or OnRoute Or WalkMap Or GoOnRoute) And tmpdistance > 0 And CurAtkMonster.NameID = 0 Then
            If GoOnRoute And IsRandomRoute Then
                IsDMove = True
                If CanusePath Then
                    If (MapName <> LockMapName And LockMapName <> "") Or Not IsInLock Then
                        WalkMsg = "Walking to lockmap from " + CStr(MakeCoords(Mid(inData, 7, 3)).Y) + ":" + CStr(MakeCoords(Mid(inData, 7, 3)).X) + " to " + CStr(tmpPos.Y) + ":" + CStr(tmpPos.X) + vbCrLf
                    Else
                        WalkMsg = "Random routing from " + CStr(MakeCoords(Mid(inData, 7, 3)).Y) + ":" + CStr(MakeCoords(Mid(inData, 7, 3)).X) + " to " + CStr(tmpPos.Y) + ":" + CStr(tmpPos.X) + vbCrLf
                    End If
                Else
                    WalkMsg = "Random move from " + CStr(MakeCoords(Mid(inData, 7, 3)).Y) + ":" + CStr(MakeCoords(Mid(inData, 7, 3)).X) + " to " + CStr(tmpPos.Y) + ":" + CStr(tmpPos.X) + vbCrLf
                End If
            ElseIf GoOnRoute And (Not IsRandomRoute) Then
                WalkMsg = "Walk from " + CStr(MakeCoords(Mid(inData, 7, 3)).Y) + ":" + CStr(MakeCoords(Mid(inData, 7, 3)).X) + " to " + CStr(tmpPos.Y) + ":" + CStr(tmpPos.X) + vbCrLf
            Else
                WalkMsg = "Go along waypoint from " + CStr(tmpCurpos.Y) + ":" + CStr(tmpCurpos.X) + " to " + CStr(tmpPos.Y) + ":" + CStr(tmpPos.X) + vbCrLf
            End If
        End If
        If Mods.STWalk And Len(WalkMsg) > 0 Then Stat WalkMsg
        If tmpdistance > 0 Then
            PlayerMoveTime = GetTickCount()
            PlayerEndMoveTime = PlayerMoveTime + (MovementSpeed * tmpdistance)
            BlockMove = True
        Else
            BlockMove = False
        End If
        IsRandommove = False
        OnRoute = False
        WalkMap = False
        GoOnRoute = False
        BackWpCounter = 0
        If FrmField.Visible Then Plot_Dot curPos, vbBlue
        ReSetCounter = 0
        upd_curMonster
        Wait = False
        frmMain.Label14.Caption = curPos.X
        frmMain.Label12.Caption = curPos.Y
        If Tracing Then SendPickup
        If (CurAtkMonster.NameID = 0 And (IsAutoKill) And UBound(MonsterList) > 0) And (Not IsAggro) And ((Not IsSPWait) Or (Not Sitting)) And (Not Pickup) Then frmMain.EstimateClosestMonster
        If CurAtkMonster.NameID > 0 Then SendAction = True
        ReSetCounter = 0
    Decode_0087 = ""
    Exit Function
errie:
    Decode_0087 = "ERROR!!! [Decode_0087] " & Err.Description
    Err.Clear
End Function

Function Decode_0088(inData As String) As String
On Error GoTo errie
    Dim X As Integer
    If Mid(inData, 3, 4) = AccountID Then
        'Stat "Moving interrupted..." & vbCrLf
        PlayerMoveTime = 0
        frmMain.TmrMove.Enabled = False
        Clear_Dot curPos
        curPos.Y = MakePort(Mid(inData, 7, 2))
        curPos.X = MakePort(Mid(inData, 9, 2))
        frmMain.Label14.Caption = curPos.X
        frmMain.Label12.Caption = curPos.Y
        If FrmField.Visible Then Plot_Dot curPos, vbBlue
        If CurrentItem.Name <> "" Then
            frmMain.labCurMons.Caption = "[" + Return_ItemName(CurrentItem.Name) + "], " _
            & CStr(EvalNorm(CurrentItem.Pos, curPos)) + " Blocks"
            SendPickup
        End If
        If (CurAtkMonster.NameID > 0) Then frmMain.SendAttack
        frmMain.tmrTicks.Enabled = True
    ElseIf UBound(MonsterList) > 0 Then
        For X = 0 To UBound(MonsterList) - 1
            If (MonsterList(X).ID = Mid(inData, 3, 4)) Then
                Clear_Dot MonsterList(X).Pos
                MonsterList(X).Pos.Y = MakePort(Mid(inData, 7, 2))
                MonsterList(X).Pos.X = MakePort(Mid(inData, 9, 2))
                If CurAtkMonster.ID = MonsterList(X).ID Then
                    CurAtkMonster = MonsterList(X)
                    upd_curMonster
                    Plot_Dot MonsterList(X).Pos, CurAtkColor
                Else
                    Plot_Dot MonsterList(X).Pos, vbRed
                End If
                MonsterList(X).Time = 0
                Exit For
            End If
        Next
      End If
      Decode_0088 = ""
      Exit Function
errie:
      Decode_0088 = "ERROR!!! [Decode_0088] " & Err.Description
      Err.Clear
End Function

Function Decode_008A(inData As String) As String
On Error GoTo errie
    'R 008a <src ID>.l <dst ID>.l <server tick>.l <src speed>.l <dst speed>.l <param1>.w <param2>.w <type>.B <param3>.w
    '               3                   7               11                      15                          19                      23                      25                  27              28
'Type DP008A
'SrcID As Long
'DstID As Long
'End Type
    Dim P1&, P2&, P3&
    Dim aType As Integer
    Dim Src$, Des$, EvStat$, X&
    
    P1 = MakePort(Mid(inData, 23, 2))
    P2 = MakePort(Mid(inData, 25, 2))
    P3 = MakePort(Mid(inData, 28, 2))
    aType = Asc(Mid(inData, 27, 1))
    
    If Mid(inData, 7, 4) <> Chr(0) + Chr(0) + Chr(0) + Chr(0) And (aType <> 1 Or aType <> 2 Or aType <> 3) Then checkKS Mid(inData, 3, 4), Mid(inData, 7, 4)
    If Mid(inData, 3, 4) = AccountID Or Mid(inData, 7, 4) = AccountID Then
        InFight = True
        Select Case aType
            Case 1 'pickup item
                Dim ITName$
                ITName = IsItem(Mid(inData, 7, 4))
                If ITName <> "" Then
                    Stat "You pick up [" & ITName & "]..." & vbCrLf
                    GoTo end8A
                End If
                CheckEvent "OnPickUpItem", "item=" & ITName
            Case 2 'sit down
                If Not Sitting Then Stat "You're sitting..." & vbCrLf
                Sitting = True
                IsSitting = False
                IsStanding = False
                CheckEvent "OnSit", "nothingtocheck=False"
            Case 3 'stand up
                If Sitting Then Stat "You're standing..." & vbCrLf
                Sitting = False
                IsSitting = False
                IsStanding = False
                CheckEvent "OnStand", "nothingtocheck=False"
            Case Else
                If Mid(inData, 3, 4) = AccountID Then
                    Src = "You"
                ElseIf Get_MonsName(Mid(inData, 3, 4)) <> "Unknow" Then
                    Src = "[" & Get_MonsName(Mid(inData, 3, 4)) & "]"
                ElseIf Left(Get_PeopleName(Mid(inData, 3, 4)), 2) <> "U:" Then
                    Src = "[" & Get_PeopleName(Mid(inData, 3, 4)) & "]"
                Else
                    Src = "[*Unknown/" & MakePort(Mid(inData, 3, 4)) & "*]"
                End If
                EvStat = "isSourceAvoidID=" & IsAvoidID(Mid(inData, 3, 4))
                EvStat = EvStat & Chr(0) & "SourceID=" & MakePort(Mid(inData, 3, 4))
                EvStat = EvStat & Chr(0) & "Source=" & Src
                
                If Mid(inData, 7, 4) = AccountID Then
                    Des = "You"
                ElseIf Get_MonsName(Mid(inData, 7, 4)) <> "Unknown" Then
                    Des = "[" & Get_MonsName(Mid(inData, 7, 4)) & "]"
                ElseIf Left(Get_PeopleName(Mid(inData, 7, 4)), 2) <> "U:" Then
                    Des = "[" & Get_PeopleName(Mid(inData, 7, 4)) & "]"
                Else
                    Des = "[*Unknown/" & MakePort(Mid(inData, 7, 4)) & "*]"
                End If
                EvStat = EvStat & Chr(0) & "isDestinationAvoidID=" & IsAvoidID(Mid(inData, 7, 4))
                EvStat = EvStat & Chr(0) & "DestinationID=" & MakePort(Mid(inData, 7, 4))
                EvStat = EvStat & Chr(0) & "Destination=" & Des

                If CurStatus(34).Active And Mid(inData, 7, 4) = AccountID Then
                    P1 = P1 / 7
                    P3 = P3 / 7
                End If

                EvStat = EvStat & Chr(0) & "Damage=" & P1
                EvStat = EvStat & Chr(0) & "LeftDamage=" & P3
                EvStat = EvStat & Chr(0) & "Hit=" & P2
                EvStat = EvStat & Chr(0) & "AttackType=" & aType

                CheckEvent "OnAttack", EvStat

                If Not MakeDamage And Mid(inData, 3, 4) = AccountID Then
                    Stat Src & " locked, " + Des + " as a Target..." + vbCrLf
                    MakeDamage = True
                    IsLock = True
                End If
                Stat Src & " attack to " & Des
                If aType = 11 Then
                    Stat ", Lucky!", vbBlue
                Else
                    If P1 = 0 Then Stat ", Miss!", vbBlue Else Stat ", " & CStr(P1) & IIf(P3 > 0, "+" & CStr(P3), "") & " Damage", IIf(Mid(inData, 3, 4) = AccountID, vbBlue, vbRed)
                    If aType = 10 Then Stat " Critical!", vbBlue
                End If
                Stat vbCrLf
                ReSetCounter = 0
                AttackCounter = 0
                PlayerMoveTime = 0
                BlockMove = False
                
                If Mid(inData, 3, 4) = AccountID And CurAtkMonster.ID = Mid(inData, 7, 4) And P1 > 0 Then DamageCounter = 0
                If (Not isPlayer(Mid(inData, 3, 4))) And Mid(inData, 7, 4) = AccountID Then
                    If Pickup And Not HaveRare Then
                        frmMain.tmrPickup.Enabled = False
                        Pickup = False
                        Pickuptime = 0
                        TryPicktime = 0
                    End If
                    If P1 > 0 Then
                        PlayerMoveTime = 0
                        BlockMove = False
                    End If
                    If Sitting Then
                        Stat "Monster Attacks, You stand up..." + vbCrLf
                        Winsock_SendPacket Chr(&H89) + Chr(0) + Chr(0) + Chr(0) + Chr(0) + _
                        Chr(0) + Chr(3), True
                        IsStanding = True
                        IsSitting = False
                    End If
                    DamageCounter = DamageCounter + P1
                    If Mid(inData, 3, 4) = CurAtkMonster.ID And Not MakeDamage Then
                        TraceMons = False
                        upd_curMonster
                    Else
                        For X = 0 To UBound(MonsterList) - 1
                            If Mid(inData, 3, 4) = MonsterList(X).ID And Not HaveRare Then
                                If Not MonsterList(X).IsAttack Then
                                    'current monster didn't attack to another people
                                    If (CurAtkMonster.NameID = 0) Or (Not MakeDamage And EvalNorm(curPos, CurAtkMonster.Pos) > EvalNorm(curPos, MonsterList(X).Pos)) Then
                                        CurAtkMonster = MonsterList(X)
                                        oldSelectPos = CurAtkMonster.Pos
                                        CurMonsterName = MonsterList(X).Name
                                        Check_Equip CurMonsterName
                                        Check_Accessory CurMonsterName
                                        upd_curMonster
                                    End If
                                    frmMain.tmrAggro.Enabled = False
                                    IsAggro = True
                                ElseIf MonsterList(X).IsAttack Then
                                    'current monster attacked to another people but it attack me too. fight back (no killsteal detection)
                                    If (CurAtkMonster.NameID = 0) Or (Not MakeDamage And (EvalNorm(curPos, CurAtkMonster.Pos) > EvalNorm(curPos, MonsterList(X).Pos))) Then
                                        CurAtkMonster = MonsterList(X)
                                        oldSelectPos = CurAtkMonster.Pos
                                        CurMonsterName = MonsterList(X).Name
                                        Check_Equip CurMonsterName
                                        Check_Accessory CurMonsterName
                                        upd_curMonster
                                    End If
                                    frmMain.tmrAggro.Enabled = False
                                    IsAggro = True
                                End If
                            End If
                        Next
                    End If
                    If P1 < Players(number).HP Then Players(number).HP = Players(number).HP - P1 Else Players(number).HP = 0
                    If (Players(number).MaxHP > 0) Then
                        If (Players(number).HP >= 0) Then frmPlayer.tabHP.width = (Players(number).HP / Players(number).MaxHP) * (frmPlayer.tabHPBg.width - 20)
                        If (Players(number).HP / Players(number).MaxHP > 0.25) Then
                           frmPlayer.tabHP.BackColor = &HC000&
                        Else
                            frmPlayer.tabHP.BackColor = &HC0&
                        End If
                    End If
                    frmPlayer.LabHP.Caption = CStr(Players(number).HP) + "  /  " + CStr(Players(number).MaxHP)
                    frmMain.SendAttack
                    InFight = True
                    IsDamage = True
                    AttackCounter = AttackCounter + 1
                    '------------------------------ Check Aggrolist ------------------------
                    For X = 0 To UBound(Aggro) - 1
                        If Aggro(X).ID = Mid(inData, 3, 4) Then GoTo aggro_end: Exit For
                    Next
                    Aggro(UBound(Aggro)).ID = Mid(inData, 3, 4)
                    ReDim Preserve Aggro(UBound(Aggro) + 1)
                    GoTo aggro_end
aggro_end:
                End If
                Check_Tele
        End Select
    End If
end8A:
Decode_008A = ""
Exit Function
errie:
Decode_008A = "ERROR!!! [Decode_008A] " & Err.Description
Err.Clear
End Function

Function Decode_008D(inData As String) As String
On Error GoTo errie
    'R 008d <len>.w <ID>.l <str>.?B
    Dim ChrID As String, i&, txtChats$, dist&, isTmp As Boolean
    ChrID = Mid(inData, 5, 4)
    txtChats = Mid(inData, 9, MakePort(Mid(inData, 3, 2)) - 8)
    dist = 0
    isTmp = False
    
    If killsteal Then
        If InStr(LCase(inData), "dc") > 0 Then
            Chat "Detected disconnect command and killsteal is enabling, terminate bot", vbRed
            ForceExit
        End If
    End If
    For i = LBound(People) To UBound(People)
        If ChrID = People(i).ID Then
            dist = EvalNorm(curPos, People(i).Pos)
            Chat "[aid:" & MakePort(ChrID) & "] [" & dist & " blks] " & txtChats, MColor.playerchat
            CheckEvent "OnPublicMessage", "posX=" & People(i).Pos.X & Chr(0) & "posY=" & People(i).Pos.Y & Chr(0) & "distance=" & dist & Chr(0) & "AccountID=" & MakePort(ChrID) & Chr(0) & "name=" & Left(txtChats, InStr(txtChats, " : ") - 1) & Chr(0) & "message=" & Right(txtChats, Len(txtChats) - InStr(1, txtChats, " : ") - 2) & Chr(0) & "isavoidid=" & CBool(IsAvoidID(ChrID))
            Exit Function
        End If
    Next
    Chat "[aid:" & MakePort(ChrID) & "] " & txtChats, MColor.playerchat
    CheckEvent "OnPublicMessage", "posX=0" & Chr(0) & "posY=0" & Chr(0) & "distance=-1" & Chr(0) & "AccountID=" & MakePort(ChrID) & Chr(0) & "name=" & Left(txtChats, InStr(txtChats, " : ") - 1) & Chr(0) & "message=" & Right(txtChats, Len(txtChats) - InStr(1, txtChats, " : ") - 2) & Chr(0) & "isavoidid=" & CBool(IsAvoidID(ChrID))
    Decode_008D = ""
Exit Function
errie:
Decode_008D = "ERROR!!! [Decode_008D] " & Err.Description
Err.Clear
End Function

Function Decode_008E(inData As String) As String
On Error GoTo errie
    Chat Mid(inData, 5, MakePort(Mid(inData, 3, 2)) - 4), MColor.mychat
    Decode_008E = ""
Exit Function
errie:
Decode_008E = "ERROR!!! [Decode_008E] " & Err.Description
Err.Clear
End Function

Function Decode_0091(inData As String) As String
On Error GoTo errie
        Dim mapgot As String, eCase As String
         PlayerMoveTime = 0
        frmMain.TmrMove.Enabled = False
        BlockMove = False
        ClearCounter = 0
        Clear_Dot curPos
        eCase = "oldMap=" & MapName & Chr(0) & "oldposX=" & curPos.Y & Chr(0) & "oldposY=" & curPos.X & Chr(0)
        curPos.Y = MakePort(Mid(inData, 19, 2))
        curPos.X = MakePort(Mid(inData, 21, 2))
        frmMain.Label14.Caption = CStr(curPos.X)
        frmMain.Label12.Caption = CStr(curPos.Y)
        ReSetCounter = 0
        mapgot = MakeString(Mid(inData, 3, 16))
        OldMap = CurrentMap
        ResetMod
        CurrentMap = mapgot
        If Not frmMain.tmrPortal.Enabled Then
            frmMain.tmrPortal.Enabled = False
            frmMain.tmrNomons.Enabled = True
            StopAction = False
        End If
        mapgot = Left(mapgot, Len(mapgot) - 4)
        If OldMap <> mapgot Then UpdateMIP mapgot, CurMIP, CurMPort
        CheckEvent "OnMapChange", eCase & "newMap=" & mapgot & Chr(0) & "newposX=" & curPos.Y & Chr(0) & "newposY=" & curPos.X
        If MapName <> mapgot Then
            Stat "Map changed to [" & mapgot & "]... " & vbCrLf
            Reset_Time
            CanusePath = True
            Load_Field mapgot
        End If
        MapName = mapgot
        Load_WayPoint mapgot
        ActionDelay = 3
        If Not IsInLock Then
            MoveOnly = True
        Else
            MoveOnly = False
        End If
        FrmField.PicMap.Refresh
        OldDot.X = 0
        OldDot.Y = 0
        Stat "Teleport to " & MapName & " (" + CStr(curPos.Y) + ":" + CStr(curPos.X) + ")" + vbCrLf
        frmMain.ClearAll
        If Not isUseHaunted Then
            Winsock_SendPacket IntToChr(&H7D), True
            Winsock_SendPacket IntToChr(&H7E) & MakeTickString, True
        End If
        frmMain.Label1.Caption = "Main Status - " & GetMapname(MapName)
        If FrmField.Visible Then Plot_Dot curPos, vbBlue
        If Make_Start_Point(curPos) Then
            If EvalNorm(curPos, WayPoint(StartPoint)) = 0 Then
                Stat "You're on waypoint at (" & CStr(WayPoint(StartPoint).Y) & ":" & CStr(WayPoint(StartPoint).X) & ")" & vbCrLf
            Else
                Stat CStr(EvalNorm(curPos, WayPoint(StartPoint))) & " Block(s) far from closest waypoint at (" & CStr(WayPoint(StartPoint).Y) & ":" & CStr(WayPoint(StartPoint).X) & ")" & vbCrLf
            End If
        End If
        Decode_0091 = ""
        Exit Function
errie:
    Decode_0091 = "ERROR!!! [Decode_0091] " & Err.Description
    Err.Clear
End Function

Function Decode_00A3(inData As String) As String
On Error GoTo errie
    Dim Args As px00A3ex, i&
    Dim bArr() As Byte
    For i = 5 To Len(inData) Step 10
        bArr = Conv2Arr(Mid(inData, i, 10))
        CopyMemory Args, bArr(0), 10
        If Args.Index > UBound(AllInv) Then ReDim Preserve AllInv(Args.Index)
        With AllInv(Args.Index)
            .NameID = MakePort(Mid(inData, i + 2, 2))
            .Name = Return_ItemName(MakeHexName(Mid(inData, i + 2, 2)))
            .Type = Args.ItemType
            .Category = Args.ItemType
            .Amount = Args.Amount
        End With
        CheckCartAI CLng(Args.Index)
    Next
    UpdateInventory
Decode_00A3 = ""
Exit Function
errie:
Decode_00A3 = "ERROR!!! [Decode_00A3] " & Err.Description
Err.Clear
End Function

Function Decode_00A4(inData As String) As String
On Error GoTo errie
        Dim Itemname As String
        Dim ItemData As String
        Dim Index As Long
        'print_packet Left(InData, MakePort(Mid(InData, 3, 2))), "==00A4=="
        ItemData = Mid(inData, 5, MakePort(Mid(inData, 3, 2)) - 4)
        Do While Len(ItemData) > 0
            Index = MakePort(Mid(ItemData, 1, 2))
            If Index > UBound(AllInv) Then ReDim Preserve AllInv(Index)
            Itemname = MakeItemName(Mid(ItemData, 3, 2), Mid(ItemData, 13, 8), Mid(ItemData, 12, 1))
            AllInv(Index).NameID = Trim(STR(MakePort(Mid(ItemData, 3, 2))))
            AllInv(Index).Name = Itemname
            AllInv(Index).Amount = 1
            AllInv(Index).Index = Mid(ItemData, 1, 2)
            AllInv(Index).Category = Asc(Mid(ItemData, 5, 1))
            AllInv(Index).Type = Mid(ItemData, 7, 2)
            AllInv(Index).Pos = MakePort(Mid(ItemData, 9, 2))
            AllInv(Index).Identified = CBool(Asc(Mid(ItemData, 6, 1)))
            If Itemname = tmpEQTeleName And WaitEquipBack And AllInv(Index).Pos > 0 Then tmpEQTelePos = Index
            If Itemname = tmpEQOldName And WaitEquipBack And AllInv(Index).Pos < 1 Then tmpEQOldPos = Index
            CheckCartAI Index
            If MakePort(Mid(ItemData, 9, 2)) > 0 Then frmMain.Update_frmArmor Itemname, MakePort(Mid(ItemData, 9, 2))
            ItemData = Right(ItemData, Len(ItemData) - 20)
        Loop
        If WaitEquipBack And tmpEQOldPos > 0 Then
            frmMain.Send_unEquip tmpEQTelePos
            frmMain.Send_Equip tmpEQOldPos
            WaitEquipBack = False
        End If
        UpdateInventory
        If Not StartBot Then
            If AlwaySit Then
                frmMain.create_chatroom "test", ChatRoomName
                frmMain.Send_Sit
            End If
            CalcModAI "00A4"
            StartBot = True
        End If
        If AutoShare Then frmMain.Set_Share
        If ExAll And (Not BlockMsg) Then frmMain.send_exall
        'ViewState = 0
        Update_FrmItem
    Decode_00A4 = ""
    Exit Function
errie:
Decode_00A4 = "ERROR!!! [Decode_00A4] " & Err.Description
Err.Clear
End Function

Function Decode_010F(inData As String) As String
On Error GoTo errie
        ReDim SkillChar(0)
        Dim Args As px010Fex
        Dim bArr() As Byte
        Dim i&
        frmSkill.lstSkill.Clear
        For i = 5 To Len(inData) Step 37
            bArr = Conv2Arr(Mid(inData, i, 37))
            CopyMemory Args, bArr(0), 37
            '{<skill ID>.w <target type>.w ?.w <lv>.w <sp>.w <range>.w <skill name>.24B <up>.B}.37B*
            '0                          2                          4    6           8               10              12                              36
            With SkillChar(UBound(SkillChar))
                .ID = Args.SkillID
                If Args.Level > 0 Then .MaxLV = Args.Level Else .MaxLV = 1
                .Target = Args.Target
                If Args.SP > 0 Then .SP = Args.SP
                .Name = MakeString(Args.SkillName)
                If LCase(MobSkill.rawname) = LCase(.Name) Then
                    MobSkill.Packet = Chr(&H13) + Chr(1) + IntToChr(CLng(MobSkill.Lv)) & IntToChr(CLng(.ID))
                    MobSkill.SP = .SP
                End If
                If MobSkill2.rawname = .Name Then
                    MobSkill2.Packet = Chr(&H13) + Chr(1) + IntToChr(CLng(MobSkill2.Lv)) & IntToChr(CLng(.ID))
                    MobSkill2.SP = .SP
                End If
            End With
            ReDim Preserve SkillChar(UBound(SkillChar) + 1)
        Next
        Update_AtkSkill
        If find_skill("AL_HEAL") Then
            UseHeal = True
            With FrmHPSPOption
                .imgHeal.Visible = True
                .LabHeal.ForeColor = 0
            End With
        End If
        frmMain.UpdateSkills
Decode_010F = ""
Exit Function
errie:
Decode_010F = "ERROR!!! [Decode_010F] " & Err.Description
Err.Clear
End Function

Function Decode_0114(inData As String) As String
On Error GoTo errie
'R 0114 <skill ID>.w <src ID>.l <dst ID>.l <server tick>.l <src speed>.l <dst speed>.l <param1>.w
'               3                       5                   9               13                      17                          21                      25
'<param2>.w <param3>.w <type>.B
' 27                        29                  30
    Dim SkID As Integer, SkName$, Src$, Des$, EvStat$
    Dim P1&, P2&, P3&, X&
    Dim aType As Integer
    
    SkID = MakePort(Mid(inData, 3, 2))
    If UBound(SkillIDName) > (SkID + 2) Then SkName = SkillIDName(SkID - 1).Name: EvStat = "SkillName=" & SkillIDName(SkID - 1).raw & Chr(0) Else EvStat = "SkillName=Unknown" & Chr(0)
    EvStat = EvStat & "SkillID=" & CStr(SkID) & Chr(0)
    
    P1 = MakePort(Mid(inData, 25, 2)) 'dmg
    P2 = MakePort(Mid(inData, 29, 2)) 'hit
    P3 = MakePort(Mid(inData, 27, 2)) 'skilllv
    aType = Asc(Mid(inData, 30, 1))

    If CurStatus(34).Active And Mid(inData, 9, 4) = AccountID Then P1 = P1 / 7
                
    If Mid(inData, 5, 4) = AccountID Then
        If SkID = 272 Then 'chain combo detected
            UseChain = False
            If Is_UseFCSkill(Get_MonsName(Mid(inData, 9, 4))) And FCSkill.Use And Players(number).SP >= (Players(number).maxsp * FCSkill.SP) And CurSpirit > 0 Then
                UseFinish = True
                X = Find_SkillId("MO_COMBOFINISH")
                If FCSkill.Lv > SkillChar(X).MaxLV Then FCSkill.Lv = SkillChar(X).MaxLV
                If X > 0 Then Send_Use_Skill SkillChar(X).ID, FCSkill.Lv, AccountID
            Else
                UseFinish = False
            End If
        End If
        If SkID = 263 Then 'triple attack detected
            If Is_UseCCSkill(Get_MonsName(Mid(inData, 9, 4))) And CCSkill.Use And Players(number).SP >= (Players(number).maxsp * CCSkill.SP) Then
                UseChain = True
                X = Find_SkillId("MO_CHAINCOMBO")
                If CCSkill.Lv > SkillChar(X).MaxLV Then CCSkill.Lv = SkillChar(X).MaxLV
                If X > 0 Then Send_Use_Skill SkillChar(X).ID, CCSkill.Lv, AccountID
            Else
                UseChain = False
            End If
        End If
        If SkID = 273 Then UseFinish = False
    End If
    
    On Error Resume Next
    If (Mid(inData, 5, 4) = AccountID) Then
        Src = "You"
    ElseIf Get_MonsName(Mid(inData, 5, 4)) <> "Unknow" Then
        Src = "[" & Get_MonsName(Mid(inData, 5, 4)) & "]"
    ElseIf Get_PeopleName(Mid(inData, 5, 4)) <> "Unknow" Then
        Src = "[" & Get_PeopleName(Mid(inData, 5, 4)) & "]"
    Else
        Src = "[Unknown/" & MakePort(Mid(inData, 5, 4)) & "]"
    End If
    EvStat = EvStat & "isSourceAvoidID=" & IsAvoidID(Mid(inData, 5, 4)) & Chr(0)
    EvStat = EvStat & "SourceID=" & MakePort(Mid(inData, 5, 4)) & Chr(0)
    EvStat = EvStat & "Source=" & Src & Chr(0)
    'evstat=evstat &  & chr(0)
    
    If (Mid(inData, 9, 4) = AccountID) Then
        Des = "You"
    ElseIf Get_MonsName(Mid(inData, 9, 4)) <> "Unknow" Then
        Des = "[" & Get_MonsName(Mid(inData, 9, 4)) & "]"
    ElseIf Get_PeopleName(Mid(inData, 9, 4)) <> "Unknow" Then
        Des = "[" & Get_PeopleName(Mid(inData, 9, 4)) & "]"
    Else
        Des = "[Unknown/" & MakePort(Mid(inData, 9, 4)) & "]"
    End If
    EvStat = EvStat & "isDestinationAvoidID=" & IsAvoidID(Mid(inData, 9, 4)) & Chr(0)
    EvStat = EvStat & "DestinationID=" & MakePort(Mid(inData, 9, 4)) & Chr(0)
    EvStat = EvStat & "Destination=" & Des & Chr(0)
    On Error GoTo errie
    If CurAtkMonster.ID = Mid(inData, 9, 4) And Mid(inData, 5, 4) <> AccountID And _
        MakePort(Mid(inData, 3, 2)) = 28 Then
        Chat "[" & Src & "] using Heal your monster!"
        response_mode = 2
        frmMain.tmrChatResponse.Interval = RandomNumber(1000, 1800)
        frmMain.tmrChatResponse.Enabled = True
    End If
    If CurAtkMonster.ID = Mid(inData, 9, 4) And Mid(inData, 5, 4) <> AccountID And _
        MakePort(Mid(inData, 3, 2)) = 29 Then
        Chat "[" & Src & "] using Increase Agi your monster!"
        response_mode = 3
        frmMain.tmrChatResponse.Interval = RandomNumber(1000, 1800)
        frmMain.tmrChatResponse.Enabled = True
    End If
    If CurAtkMonster.ID = Mid(inData, 9, 4) And Mid(inData, 5, 4) <> AccountID And _
        Asc(Mid(inData, 3, 1)) = 34 Then
        Chat "[" & Src & "] using Blessing your monster!"
        response_mode = 4
        frmMain.tmrChatResponse.Interval = RandomNumber(1000, 1800)
        frmMain.tmrChatResponse.Enabled = True
    End If

    If (CurMonsterName <> "" And CurMonsterName <> "None" And CurAtkMonster.NameID > 0) And Mid(inData, 5, 4) = AccountID Then
        If (Not MakeDamage) Then
            Stat "You locked, [" & Des & "] as a Target..." + vbCrLf
            IsLock = True
        End If
        If MakePort(Mid(inData, 21, 2)) > 0 Then
            MakeDamage = True
            SkillWait = False
            frmMain.tmrSkillDelay.Enabled = False
            SkillCounter = SkillCounter + 1
            frmMain.Check_ResetCounter
        End If
        Casting = False
        DamageCounter = 0
        AttackCounter = 0
        AttCounter = AttCounter + 1
    End If


'    type=04 observed when firewall was used. is that the almost same as type=06?
'    type=06 skill for just one hit? param1 is total damage, param2 is level, param3 will always stay 1.
'    type=08 skill for multiple hits? param1 is total damage, param2 is level, param3 will be a number of hit.
    
    EvStat = EvStat & "Damage=" & P1 & Chr(0)
    EvStat = EvStat & "SkillLevel=" & P2 & Chr(0)
    EvStat = EvStat & "Hit=" & P3

    Stat "[0114] " & Src & " used skill [" & SkName & "/Lv:" & CStr(P2) & "]" & IIf(Mid(inData, 5, 4) = AccountID, "/" & SkillCounter, "") & " on " & Des, IIf(Mid(inData, 9, 4) = AccountID, vbRed, 0)
    If P1 = 0 Then
        Stat ", Miss!", vbBlue
    Else
        Stat ", " & CStr(P1) & " Damage", IIf(Mid(inData, 9, 4) = AccountID, vbRed, 0)
    End If
    If P3 > 1 Then Stat " [" & CStr(P3) & " hits]"
    Stat vbCrLf
    checkKS Mid(inData, 5, 4), Mid(inData, 9, 4)
    CheckEvent "OnSkillUse", EvStat
    Exit Function
errie:
Decode_0114 = "ERROR!!! [Decode_0114] " & Err.Description
Err.Clear
End Function

Function Decode_014C(inData As String) As String
On Error GoTo errie
'R 014C <len>.w
'                 3
'(<type>.l <guildID>.l <guild name>.24B).*
'0                4                     8
    Dim pLen&, i&, j&
    pLen = MakePort(Mid(inData, 3, 2))
    For i = 5 To pLen Step 32
        For j = 1 To UBound(GuildAlliance)
            If GuildAlliance(j).ID = Mid(inData, i + 4, 4) Then
                GuildAlliance(j).isAlliance = Not (CBool(Val(MakePort(Mid(inData, i, 4)))))
                GuildAlliance(j).Name = MakeString(Mid(inData, i + 8, 24))
                GoTo stepnext
            End If
        Next
        ReDim Preserve GuildAlliance(UBound(GuildAlliance) + 1)
        GuildAlliance(UBound(GuildAlliance)).ID = Mid(inData, i + 4, 4)
        GuildAlliance(UBound(GuildAlliance)).Name = MakeString(Mid(inData, i + 8, 24))
        GuildAlliance(UBound(GuildAlliance)).isAlliance = Not (CBool(Val(MakePort(Mid(inData, i, 4)))))
stepnext:
    Next
    Exit Function
errie:
    Decode_014C = "ERROR!!! [Decode_014C] " & Err.Description
    Err.Clear
End Function

Function Decode_0150(inData As String) As String
On Error GoTo errie
'R 0150 <guildID>.l <guildLv>.l <connum>.l <Max PPL?>.l <Avl.lvl>.l ?.l <next_exp>.l ?.16B
'               3                   7                   11                      15                      19              23   27                     31
'<guild name>.24B <guild master>.24B ?.16B
'47                                 71                                 95
    With GuildInfo
        .CurOnline = MakePort(Mid(inData, 11, 4))
        .GuID = Mid(inData, 3, 4)
        .GuLV = MakePort(Mid(inData, 7, 4))
        .MaxPPL = MakePort(Mid(inData, 15, 4))
        .GuAvLV = MakePort(Mid(inData, 19, 4))
        .GuEXP = MakePort(Mid(inData, 23, 4))
        .GuNextEXP = MakePort(Mid(inData, 27, 4))
        .Name = MakeString(Mid(inData, 47, 24))
        .GuMaster = MakeString(Mid(inData, 71, 24))
    End With
    UpdateGuild
    Exit Function
errie:
    Decode_0150 = "ERROR!!! [Decode_0150] " & Err.Description
    Err.Clear
End Function

Function Decode_0154(inData As String) As String
On Error GoTo errie
    Dim i&, X As Integer
'R 0154 <len>.w
'{<accID>.l <charactorID>.l <hair type>.w <hair color>.w <sex>.w <job>.w <lvl?>.w
' 0                     4                           8                       10                          12          14          16
'<guild exp>.l <online>.l <Position>.l ?.50B <nick>.24B}*
' 18                    22              26                      30        80
    For i = 5 To MakePort(Mid(inData, 3, 2)) Step 104
        With Guild(UBound(Guild))
            .AccID = Mid(inData, i, 4)
            .CharID = Mid(inData, i + 4, 4)
            If MakePort(Mid(inData, i + 12, 2)) = 0 Then .Sex = " <F>" Else .Sex = " <M>"
            .Class = Return_Class(MakePort(Mid(inData, i + 14, 2)))
            .Lv = MakePort(Mid(inData, i + 16, 2))
            .EXP = MakePort(Mid(inData, i + 18, 4))
            .isOnline = CBool(MakePort(Mid(inData, i + 12, 2)))
            .Name = MakeString(Mid(inData, i + 80, 24))
            .Position = MakePort(Mid(inData, 26, 4))
            .PosName = GetGuildPos(.Position)
        End With
'        For X = 0 To UBound(Guild) - 1
'            If Guild(X).AccID = Mid(InData, i + 26, 4) Then
'                Guild(X).ID = Mid(InData, i, 4)
'                If MakePort(Mid(InData, i + 12, 2)) = 0 Then
'                    Guild(X).Sex = " <F>"
'                Else
'                    Guild(X).Sex = " <M>"
'                End If
'                Guild(X).Class = Return_Class(MakePort(Mid(InData, i + 14, 2)))
'                Guild(X).Lv = MakePort(Mid(InData, i + 16, 2))
'                Guild(X).EXP = GetLong(Mid(InData, i + 18, 4))
'                If MakePort(Mid(InData, i + 22, 2)) = 1 Then
'                    Guild(X).isOnline = True
'                Else
'                    Guild(X).isOnline = False
'                End If
'                Guild(X).Name = MakeString(Mid(InData, i + 80, 24))
'            End If
'        Next
    Next
    UpdateGuild
    Decode_0154 = ""
Exit Function
errie:
Decode_0154 = "ERROR!!! [Decode_0154] " & Err.Description
Err.Clear
End Function

Function Decode_0166(inData As String) As String
On Error GoTo errie
    ReDim GuildPos(0)
    Dim i As Integer
    For i = 5 To Len(inData) Step 28
        GuildPos(UBound(GuildPos)).Position = MakePort(Mid(inData, i, 4))
        GuildPos(UBound(GuildPos)).PosName = MakeString(Mid(inData, i + 4, 24))
        ReDim Preserve GuildPos(UBound(GuildPos) + 1)
    Next
    UpdateGuild
Exit Function
errie:
Decode_0166 = "ERROR!!! [Decode_0166] " & Err.Description
Err.Clear
End Function

Function Decode_01DE(inData As String) As String
On Error GoTo errie
    Dim SkID As Integer, SkName$, Src$, Des$, EvStat$
    Dim P1&, P2&, P3&, X&
    Dim aType As Integer
'DE01      0B00             07D30200  01CB0000   6C267800         51020000         01000000
'R 01DE <skill ID>.w <src ID>.l     <dst ID>.l    <server tick>.l  <src speed>.l <dst speed>.l
'                   3                       5                   9               13                          17                          21
'00000000 0400 0100 05
'25               29      31    33
'<param1>.w <param2>.w <param3>.w <type>.B
'25                     27                      29                  30

    SkID = MakePort(Mid(inData, 3, 2))
    If UBound(SkillIDName) > (SkID + 2) Then SkName = SkillIDName(SkID - 1).Name: EvStat = "SkillName=" & SkillIDName(SkID - 1).raw & Chr(0) Else EvStat = "SkillName=Unknown" & Chr(0)
    EvStat = EvStat & "SkillID=" & CStr(SkID) & Chr(0)
    
    P1 = MakePort(Mid(inData, 25, 4)) 'dmg
    P2 = MakePort(Mid(inData, 29, 2)) 'hit
    P3 = MakePort(Mid(inData, 31, 2)) 'skilllv
    aType = Asc(Mid(inData, 33, 1))
    
    If CurStatus(34).Active And Mid(inData, 9, 4) = AccountID Then P1 = P1 / 7

    If Mid(inData, 5, 4) = AccountID Then
        If SkID = 272 Then 'chain combo detected
            UseChain = False
            If Is_UseFCSkill(Get_MonsName(Mid(inData, 9, 4))) And FCSkill.Use And Players(number).SP >= (Players(number).maxsp * FCSkill.SP) And CurSpirit > 0 Then
                UseFinish = True
                X = Find_SkillId("MO_COMBOFINISH")
                If X > -1 Then
                    If FCSkill.Lv > SkillChar(X).MaxLV Then FCSkill.Lv = SkillChar(X).MaxLV
                End If
                If X > 0 Then Send_Use_Skill SkillChar(X).ID, FCSkill.Lv, AccountID
            Else
                UseFinish = False
            End If
        End If
        If SkID = 263 Then 'triple attack detected
            If Is_UseCCSkill(Get_MonsName(Mid(inData, 9, 4))) And CCSkill.Use And Players(number).SP >= (Players(number).maxsp * CCSkill.SP) Then
                UseChain = True
                X = Find_SkillId("MO_CHAINCOMBO")
                If X > -1 Then
                    If CCSkill.Lv > SkillChar(X).MaxLV Then CCSkill.Lv = SkillChar(X).MaxLV
                End If
                If X > 0 Then Send_Use_Skill SkillChar(X).ID, CCSkill.Lv, AccountID
            Else
                UseChain = False
            End If
        End If
        If SkID = 273 Then UseFinish = False
    End If
    
    On Error Resume Next
    If (Mid(inData, 5, 4) = AccountID) Then
        Src = "You"
    ElseIf Get_MonsName(Mid(inData, 5, 4)) <> "Unknow" Then
        Src = "[" & Get_MonsName(Mid(inData, 5, 4)) & "]"
    ElseIf Left(Get_PeopleName(Mid(inData, 5, 4)), 2) <> "U:" Then
        Src = "[" & Get_PeopleName(Mid(inData, 5, 4)) & "]"
    Else
        Src = "[Unknown/" & MakePort(Mid(inData, 5, 4)) & "]"
    End If
    EvStat = EvStat & "isSourceAvoidID=" & IsAvoidID(Mid(inData, 5, 4)) & Chr(0)
    EvStat = EvStat & "SourceID=" & MakePort(Mid(inData, 5, 4)) & Chr(0)
    EvStat = EvStat & "Source=" & Src & Chr(0)
    'evstat=evstat &  & chr(0)
    
    If (Mid(inData, 9, 4) = AccountID) Then
        Des = "You"
    ElseIf Get_MonsName(Mid(inData, 9, 4)) <> "Unknow" Then
        Des = "[" & Get_MonsName(Mid(inData, 9, 4)) & "]"
    ElseIf Left(Get_PeopleName(Mid(inData, 9, 4)), 2) <> "U:" Then
        Des = "[" & Get_PeopleName(Mid(inData, 9, 4)) & "]"
    Else
        Des = "[Unknown/" & MakePort(Mid(inData, 9, 4)) & "]"
    End If
    EvStat = EvStat & "isDestinationAvoidID=" & IsAvoidID(Mid(inData, 9, 4)) & Chr(0)
    EvStat = EvStat & "DestinationID=" & MakePort(Mid(inData, 9, 4)) & Chr(0)
    EvStat = EvStat & "Destination=" & Des & Chr(0)
    On Error GoTo errie
    
    If CurAtkMonster.ID = Mid(inData, 9, 4) And Mid(inData, 5, 4) <> AccountID And _
        MakePort(Mid(inData, 3, 2)) = 28 Then
        Chat "[" & Src & "] using Heal your monster!"
        response_mode = 2
        frmMain.tmrChatResponse.Interval = RandomNumber(1000, 1800)
        frmMain.tmrChatResponse.Enabled = True
    End If
    If CurAtkMonster.ID = Mid(inData, 9, 4) And Mid(inData, 5, 4) <> AccountID And _
        MakePort(Mid(inData, 3, 2)) = 29 Then
        Chat "[" & Src & "] using Increase Agi your monster!"
        response_mode = 3
        frmMain.tmrChatResponse.Interval = RandomNumber(1000, 1800)
        frmMain.tmrChatResponse.Enabled = True
    End If
    If CurAtkMonster.ID = Mid(inData, 9, 4) And Mid(inData, 5, 4) <> AccountID And _
        Asc(Mid(inData, 3, 1)) = 34 Then
        Chat "[" & Src & "] using Blessing your monster!"
        response_mode = 4
        frmMain.tmrChatResponse.Interval = RandomNumber(1000, 1800)
        frmMain.tmrChatResponse.Enabled = True
    End If

    If (CurMonsterName <> "" And CurMonsterName <> "None" And CurAtkMonster.NameID > 0) And Mid(inData, 5, 4) = AccountID Then
        If (Not MakeDamage) Then
            Stat "You locked, [" & Des & "] as a Target..." + vbCrLf
            IsLock = True
        End If
        If CStr(MakePort(Mid(inData, 21, 2))) > 0 Then
            'Stat "[" + frmSkill.Return_SkillName(Asc(Mid(InData, 3, 1))) + "] Skill to [" + Return_MonsterName(CurAtkMonster.Nameid) + "], " + CStr(MakePort(Mid(InData, 25, 2))) + " Damage" + vbCrLf
            MakeDamage = True
            SkillWait = False
            frmMain.tmrSkillDelay.Enabled = False
            SkillCounter = SkillCounter + 1
            frmMain.Check_ResetCounter
        Else
            'Stat "[" + frmSkill.Return_SkillName(Hex(Asc(Mid(InData, 3, 1)))) + "] Skill to " + Return_MonsterName(CurAtkMonster.Nameid) + ", " + "Miss!" + vbCrLf
        End If
        Casting = False
        DamageCounter = 0
        AttackCounter = 0
        AttCounter = AttCounter + 1
    End If


'    type=04 observed when firewall was used. is that the almost same as type=06?
'    type=06 skill for just one hit? param1 is total damage, param2 is level, param3 will always stay 1.
'    type=08 skill for multiple hits? param1 is total damage, param2 is level, param3 will be a number of hit.
    
    EvStat = EvStat & "Damage=" & P1 & Chr(0)
    EvStat = EvStat & "SkillLevel=" & P2 & Chr(0)
    EvStat = EvStat & "Hit=" & P3

    Stat "[01DE] " & Src & " used skill [" & SkName & "/Lv:" & CStr(P2) & "]" & IIf(Mid(inData, 5, 4) = AccountID, "/" & SkillCounter, "") & " on " & Des, IIf(Mid(inData, 9, 4) = AccountID, vbRed, 0)
    If P1 = 0 Then
        Stat ", Miss!", vbBlue
    Else
        Stat ", " & CStr(P1) & " Damage", IIf(Mid(inData, 9, 4) = AccountID, vbRed, 0)
    End If
    If P3 > 1 Then Stat " [" & CStr(P3) & " hits]"
    Stat vbCrLf
    If aType <> 4 Then checkKS Mid(inData, 5, 4), Mid(inData, 9, 4)
    CheckEvent "OnSkillUse", EvStat
    Exit Function
errie:
Decode_01DE = "ERROR!!! [Decode_01DE] " & Err.Description
Err.Clear
End Function

Function Decode_01EE(inData As String) As String
On Error GoTo errie
    Dim Args As px00A3ex, i&
    Dim bArr() As Byte
    For i = 5 To Len(inData) Step 18
        bArr = Conv2Arr(Mid(inData, i, 10))
        CopyMemory Args, bArr(0), 10
        If Args.Index > UBound(AllInv) Then ReDim Preserve AllInv(Args.Index)
        With AllInv(Args.Index)
            .NameID = MakePort(Mid(inData, i + 2, 2))
            .Name = Return_ItemName(MakeHexName(Mid(inData, i + 2, 2)))
            .Type = Args.ItemType
            .Category = Args.ItemType
            .Amount = Args.Amount
        End With
        CheckCartAI CLng(Args.Index)
    Next
    UpdateInventory
Decode_01EE = ""
Exit Function
errie:
Decode_01EE = "ERROR!!! [Decode_01EE] " & Err.Description
Err.Clear
End Function

