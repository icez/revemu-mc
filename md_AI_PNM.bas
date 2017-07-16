Attribute VB_Name = "md_AI_PNM"
Option Explicit

'// Player, NPC, Monster, People, Portal
Public ExitPortal() As typePNM_Portal
Type typePNM_Portal
    ID As String * 4
    Pos As Coord
End Type

Public NPCList() As typePNM_People
Public People() As typePNM_People
Type typePNM_People
    ID As String * 4
    Pos As Coord
    NextPos As Coord
    Time As Long
    NameID As Long
    Speed As Integer
    Sex As String
    Name As String
    Class As String
    Healed As Integer
End Type

Public Monsters() As typePNM_MData 'monster name data
Type typePNM_MData
    ID As Integer
    Name As String
End Type

Public Attack() As typePNM_MAttack 'old attacking method
Type typePNM_MAttack
    ID As Integer
    Name As String
    lv1 As Byte
    lv2 As Byte
    sp1 As Integer
    sp2 As Integer
    Skill1 As String
    Skill2 As String
    Spell1 As String
    Spell2 As String
    Status As Long
    UTime1 As Integer
    UTime2 As Integer
End Type

'Public SKAttack() As typePNM_MSkillAttack 'new skill attacking method
'Type typePNM_MSkillAttack
'    SkID As Integer
'    SkName As String
'    Sk
'End Type



Function Decode_0078(inData As String) As String
On Error GoTo errie
    If Asc(Mid(inData, 16, 1)) < &H389 Then
        Dim NameID As Integer
        Dim X As Integer
        islist = False
        Dim IsPet As Boolean
        NameID = MakePort(Mid(inData, 15, 2))
        If NameID > 1000 Then GoTo mons
'------------------------------ NPC List ------------------
        If NameID = &H2D Then GoTo CheckPortal
        If NameID < 40 Then GoTo CheckPeople
        'Dim NPCid As Long
        If Mid(inData, 3, 4) = CulVertNPC Then
            Stat "Found [Culvert NPC] at " & CStr(MakeCoords(Mid(inData, 47, 3)).Y) & ":" & CStr(MakeCoords(Mid(inData, 47, 3)).X) & vbCrLf
        End If
'---------- New Add
        Dim foundMynpc As Boolean
        foundMynpc = False
        For X = 0 To UBound(myNPC)
            If Mid(inData, 3, 4) = myNPC(X).ID Then
                myNPC(X).Pos = MakeCoords(Mid(inData, 47, 3))
                myNPC(X).location = MapName
                foundMynpc = True
                Exit For
            End If
        Next
        If Not foundMynpc Then
            If myNPC(0).location <> "" Then ReDim Preserve myNPC(UBound(myNPC) + 1)
            myNPC(UBound(myNPC)).ID = Mid(inData, 3, 4)
            myNPC(UBound(myNPC)).Pos = MakeCoords(Mid(inData, 47, 3))
            myNPC(UBound(myNPC)).location = MapName
            Winsock_SendPacket IntToChr(&H94) & Mid(inData, 3, 4), True
        End If
        
        If UBound(NPCList) > 0 Then
        For X = 0 To UBound(NPCList) - 1
            If Mid(inData, 3, 4) = NPCList(X).ID Then
                Clear_Dot NPCList(X).Pos
                NPCList(X).Pos = MakeCoords(Mid(inData, 47, 3))
                Plot_Dot NPCList(X).Pos, vbYellow
                UpdateNPC
                GoTo end78
            End If
        Next
        End If
        NPCList(UBound(NPCList)).NameID = NameID
        NPCList(UBound(NPCList)).ID = Mid(inData, 3, 4)
        NPCList(UBound(NPCList)).Pos = MakeCoords(Mid(inData, 47, 3))
        Plot_Dot NPCList(UBound(NPCList)).Pos, vbYellow
        NPCList(UBound(NPCList)).NameID = MakePort(Mid(inData, 15, 2))
        ReDim Preserve NPCList(UBound(NPCList) + 1)
        If foundMynpc Then Winsock_SendPacket IntToChr(&H94) & Mid(inData, 3, 4), True
        UpdateNPC
        GoTo end78
'--------- End Add
'------------------------------- Check Exit Portal -----------------------------------
CheckPortal:
        If MakePort(Mid(inData, 15, 2)) = &H2D And Not PartyMode Then
            Stat "Found Exit Portal at " & CStr(MakeCoords(Mid(inData, 47, 3)).Y) & ":" & CStr(MakeCoords(Mid(inData, 47, 3)).X) & " to " & GetWarpInfo(MapName, MakeCoords(Mid(inData, 47, 3)).Y, MakeCoords(Mid(inData, 47, 3)).X) & vbCrLf
            ExitPortal(UBound(ExitPortal)).ID = Mid(inData, 3, 4)
            'ExitPortal(UBound(ExitPortal)).NameID = &H2D
            ExitPortal(UBound(ExitPortal)).Pos = MakeCoords(Mid(inData, 47, 3))
            ReDim Preserve ExitPortal(UBound(ExitPortal) + 1)
            GoTo end78
            If Not CanGO(curPos, ExitPortal(UBound(ExitPortal) - 1).Pos) Then GoTo end78
            If Not MoveOnly And Not SellMode And AutoAI And Not IsOnWayPoint(curPos) Then
                Stat "The Bot's waiting to teleport..." & vbCrLf
                PortalTime = 0
                frmMain.tmrPortal.Enabled = True
            ElseIf Not Dead And Not SellMode And MoveOnly And (Not CanUseWP) And AutoAI Then
                Stat "Map change go to start map..." + vbCrLf
                Winsock_SendPacket IntToChr(&H64 + &H21) + _
                MakeCoordString(MakeCoords(Mid(inData, 47, 3))), True
                
            End If
            DetectPortal = True
            If CurAtkMonster.NameID > 0 And Not MakeDamage And Not IsOnWayPoint(curPos) Then
                SkillCounter = 0
                DamageCounter = 0
                InFight = False
                CurAtkMonster.NameID = 0
                Clear_Dot CurAtkMonster.Pos
                CurAtkMonster.ID = String(4, Chr(0))
                CurAtkMonster.Pos.X = 0
                CurAtkMonster.Pos.Y = 0
                AttCounter = 0
                IsAggro = False
                IsLock = False
                IsDamage = False
                CurMonsterName = "None"
                frmMain.labCurMons.Caption = "[None]"
                Stat "Stop Attacking..." + vbCrLf
                TraceMons = False
                If Make_Start_Point(curPos) Then
                    Stat "Back to waypoint..." & vbCrLf
                    move_to WayPoint(StartPoint)
                    BackWP = True
                End If
            End If
endexit:
            GoTo end78
        End If
 '------------------------------- Update People -------------------------------------
CheckPeople:
        AI_AvoidID (Mid(inData, 3, 4))
        If UBound(People) > 0 Then
        For X = 0 To UBound(People) - 1
            If Mid(inData, 3, 4) = People(X).ID Then
                If Not Disable_frmPeople Then Clear_Dot People(X).Pos
                People(X).Pos = MakeCoords(Mid(inData, 47, 3))
                If Not Disable_frmPeople Then Plot_Dot People(X).Pos, PColor
                'If frmPeople.Visible = True Then
                    UpdatePeople
                'End If
                GoTo end78
            End If
        Next
        End If
        People(UBound(People)).ID = Mid(inData, 3, 4)
        People(UBound(People)).Pos = MakeCoords(Mid(inData, 47, 3))
        If Not Disable_frmPeople Then Plot_Dot People(UBound(People)).Pos, PColor
        People(UBound(People)).NameID = MakePort(Mid(inData, 15, 2))
'R 0078 <ID>.l <speed>.w ?.w ?.w <option>.w <class>.w <hair>.w <weapon>.w <head option bottom>.w <sheild>.w <head option top>.w <head option mid>.w <hair color>.w ?.w <head dir>.w <guild>.w ?.w ?.w <manner>.w <karma>.w ?.B <sex>.B <X_Y_dir>.3B ?.B ?.B <sit>.B
'               3           7                   9       11  13                  15              17              19                  21                                              23                  25                                      27                                  29
        'People(UBound(People)).Shield = MakePort(Mid(inData, 21, 2))
        'People(UBound(People)).HeadB = MakePort(Mid(inData, 23, 2))
        'People(UBound(People)).HeadT = MakePort(Mid(inData, 25, 2))
        'People(UBound(People)).HeadM = MakePort(Mid(inData, 27, 2))
        'People(UBound(People)).HairC = MakePort(Mid(inData, 29, 2))
        'People(UBound(People)).Weapon = MakePort(Mid(inData, 19, 2))
        'People(UBound(People)).Hair = MakePort(Mid(inData, 17, 2))
        People(UBound(People)).Class = Return_Class(People(UBound(People)).NameID)
        If Asc(Mid(inData, 46, 1)) = 0 Then
            People(UBound(People)).Sex = " <F>"
        Else
            People(UBound(People)).Sex = " <M>"
        End If
        If IsAvoidID(Mid(inData, 3, 4)) Then
            CheckEvent "OnGMAppear", "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0) & "job=" & People(UBound(People)).Class & Chr(0) & "posX=" & People(UBound(People)).Pos.Y & Chr(0) & "posY=" & People(UBound(People)).Pos.X & Chr(0) & "distance=" & EvalNorm(People(UBound(People)).Pos, curPos)
        ElseIf IsAvoid(People(UBound(People)).Name) And Len(People(UBound(People)).Name) > 0 Then
            CheckEvent "OnAvoidListAppear", "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0) & "job=" & People(UBound(People)).Class & Chr(0) & "posX=" & People(UBound(People)).Pos.Y & Chr(0) & "posY=" & People(UBound(People)).Pos.X & Chr(0) & "distance=" & EvalNorm(People(UBound(People)).Pos, curPos)
        ElseIf isWarpList(People(UBound(People)).Name) And Len(People(UBound(People)).Name) > 0 Then
            CheckEvent "OnWarpListAppear", "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0) & "job=" & People(UBound(People)).Class & Chr(0) & "posX=" & People(UBound(People)).Pos.Y & Chr(0) & "posY=" & People(UBound(People)).Pos.X & Chr(0) & "distance=" & EvalNorm(People(UBound(People)).Pos, curPos)
        Else
            CheckEvent "OnPlayerAppear", "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0) & "job=" & People(UBound(People)).Class & Chr(0) & "posX=" & People(UBound(People)).Pos.Y & Chr(0) & "posY=" & People(UBound(People)).Pos.X & Chr(0) & "distance=" & EvalNorm(People(UBound(People)).Pos, curPos)
        End If
        Winsock_SendPacket IntToChr(&H94) & Mid(inData, 3, 4), True
        ReDim Preserve People(UBound(People) + 1)
        'If frmPeople.Visible = True Then
            UpdatePeople
        'End If
        GoTo end78
    End If
        Dim found As Boolean
        Dim found2 As Boolean
mons:
        found = False
        islist = False
        IsPet = False
        If Asc(Mid(inData, 17, 1)) > 0 Then IsPet = True
        For X = 0 To UBound(Attack)
            If MakePort(Mid(inData, 15, 2)) = Attack(X).ID Then
                found = True
                Exit For
            End If
        Next
        Dim tmpname As String
        For X = 0 To UBound(Monsters)
            If MakePort(Mid(inData, 15, 2)) = Monsters(X).ID Then
                islist = True
                tmpname = Monsters(X).Name
                Exit For
            End If
        Next
        If Mid(inData, 3, 4) = MyPet.ID And MyPet.Type = "" Then MyPet.Type = tmpname
        'If islist Then
            found2 = False
            If UBound(MonsterList) > 0 Then
            For X = 0 To UBound(MonsterList) - 1
                If Mid(inData, 3, 4) = MonsterList(X).ID Then
                    found2 = True
                    If Not found Then MonsterList(X).NoAttack = True
                    If IsPet Then MonsterList(X).IsPet = True
                    Clear_Dot MonsterList(X).Pos
                    MonsterList(X).Pos = MakeCoords(Mid(inData, 47, 3))
                    MonsterList(X).StatusA = MakePort(Mid(inData, 9, 2))
                    MonsterList(X).StatusB = MakePort(Mid(inData, 11, 2))
                    If (MonsterList(X).StatusA > 0 And MonsterList(X).StatusA < 5) Or MonsterList(X).StatusB > 0 Then MonsterList(X).IsTrap = True
                    If CanGO(curPos, MonsterList(X).Pos) Then MonsterList(X).CantGo = True
                    upd_frmMonster
                    If MyPet.ID = MonsterList(X).ID Then
                        Plot_Dot MonsterList(X).Pos, 16711935
                    ElseIf CurAtkMonster.ID <> MonsterList(X).ID Then
                        Plot_Dot MonsterList(X).Pos, vbRed
                    Else
                        Plot_Dot MonsterList(X).Pos, CurAtkColor
                    End If
                    CheckEvent "OnMonsterAppear", "name=" & MonsterList(X).Name & Chr(0) & "posX=" & MonsterList(X).Pos.Y & Chr(0) & "posY=" & MonsterList(X).Pos.X & Chr(0) & "distance=" & EvalNorm(curPos, MonsterList(X).Pos)
                    Exit For
                End If
            Next
            End If
            If (Not Sitting) And (Not Pickup) And (Not InFight) And (CurAtkMonster.NameID > 0) Then SendAction = True
            If Mid(inData, 3, 4) = CurAtkMonster.ID Then
                    If IsPet Then
                        Stat "This's a pet!, abort target..." & vbCrLf
                        Clear_This_Mons 0
                        GoTo endifcur
                    End If
                    CurAtkMonster.Pos = MakeCoords(Mid(inData, 47, 3))
endifcur:
                     upd_curMonster
            End If
            If Not found2 Then
                If Not found Then MonsterList(UBound(MonsterList)).NoAttack = True
                If IsPet Then MonsterList(UBound(MonsterList)).IsPet = True
                MonsterList(UBound(MonsterList)).ID = Mid(inData, 3, 4)
                MonsterList(UBound(MonsterList)).Pos = MakeCoords(Mid(inData, 47, 3))
                If MyPet.ID = MonsterList(UBound(MonsterList)).ID Then
                    Plot_Dot MonsterList(UBound(MonsterList)).Pos, 16711935
                Else
                    Plot_Dot MonsterList(UBound(MonsterList)).Pos, vbRed
                End If
                MonsterList(UBound(MonsterList)).StatusA = MakePort(Mid(inData, 9, 2))
                MonsterList(UBound(MonsterList)).StatusB = MakePort(Mid(inData, 11, 2))
                If (MakePort(Mid(inData, 9, 2)) > 0 And MakePort(Mid(inData, 9, 2)) < 5) Or MakePort(Mid(inData, 11, 2)) > 0 Then MonsterList(UBound(MonsterList)).IsTrap = True
                MonsterList(UBound(MonsterList)).NameID = MakePort(Mid(inData, 15, 2))
                MonsterList(UBound(MonsterList)).IsAttack = False
                If islist Then MonsterList(UBound(MonsterList)).Name = tmpname
                CheckEvent "OnMonsterAppear", "name=" & MonsterList(UBound(MonsterList)).Name & Chr(0) & "posX=" & MonsterList(UBound(MonsterList)).Pos.Y & Chr(0) & "posY=" & MonsterList(UBound(MonsterList)).Pos.X & Chr(0) & "distance=" & EvalNorm(MonsterList(UBound(MonsterList)).Pos, curPos)
                If IsPet Then Winsock_SendPacket IntToChr(&H94) & Mid(inData, 3, 4), True
                ReDim Preserve MonsterList(UBound(MonsterList) + 1)
                upd_frmMonster
            End If
            If CurAtkMonster.NameID > 0 Then
                If EvalNorm(CurAtkMonster.Pos, curPos) > 2 Then SendAction = True
            End If
        'End If
        If Not islist And NameID > 20 And Not IsPet Then
            Stat "Unknow Monster " & MakeHexName(Mid(inData, 15, 2)) & vbCrLf
            Winsock_SendPacket IntToChr(&H94) & Mid(inData, 3, 4), True
        End If
end78:
Decode_0078 = ""
Exit Function
errie:
Decode_0078 = "ERROR!!! [Decode_0078] " & Err.Description
Err.Clear
End Function

Function Decode_0079(inData As String) As String
On Error GoTo errie
    Dim X As Integer
    If MakePort(Mid(inData, 15, 2)) < &H389 Then
            AI_AvoidID (Mid(inData, 3, 4))
            If UBound(People) > 0 Then
            For X = 0 To UBound(People) - 1
                If Mid(inData, 3, 4) = People(X).ID Then
                    If Not Disable_frmPeople Then Clear_Dot People(X).Pos
                    People(X).Pos = MakeCoords(Mid(inData, 47, 3))
                    If Not Disable_frmPeople Then Plot_Dot People(X).Pos, PColor
                    'If frmPeople.Visible = True Then
                        UpdatePeople
                    'End If
                    GoTo skip2
                End If
            Next
            End If
                People(UBound(People)).ID = Mid(inData, 3, 4)
                People(UBound(People)).Pos = MakeCoords(Mid(inData, 47, 3))
                If Mid(inData, 3, 4) = FollowMode.AID Then move_to People(UBound(People)).Pos, 2
                If Not Disable_frmPeople Then Plot_Dot People(X).Pos, PColor
                People(UBound(People)).NameID = MakePort(Mid(inData, 15, 2))
                'If Mods.STDebug Then Chat "79: " & MakePort(Mid(InData, 21, 2)) & "/" & MakePort(Mid(InData, 23, 2)), &H555555
                'People(UBound(People)).Shield = MakePort(Mid(inData, 21, 2))
                'People(UBound(People)).HeadB = MakePort(Mid(inData, 23, 2))
                'People(UBound(People)).HeadT = MakePort(Mid(inData, 25, 2))
                'People(UBound(People)).HeadM = MakePort(Mid(inData, 27, 2))
                'People(UBound(People)).HairC = MakePort(Mid(inData, 29, 2))
                'People(UBound(People)).Weapon = MakePort(Mid(inData, 19, 2))
                'People(UBound(People)).Hair = MakePort(Mid(inData, 17, 2))
                'print_packet Left(InData, 51), People(UBound(People)).nameid
                If Asc(Mid(inData, 46, 1)) = 0 Then
                    People(UBound(People)).Sex = " <F>"
                Else
                    People(UBound(People)).Sex = " <M>"
                End If
                People(UBound(People)).Class = Return_Class(People(UBound(People)).NameID)
                If IsAvoidID(Mid(inData, 3, 4)) Then
                    CheckEvent "OnGMAppear", "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0) & "job=" & People(UBound(People)).Class & Chr(0) & "posX=" & People(UBound(People)).Pos.Y & Chr(0) & "posY=" & People(UBound(People)).Pos.X & Chr(0) & "distance=" & EvalNorm(People(UBound(People)).Pos, curPos)
                ElseIf IsAvoid(People(UBound(People)).Name) And Len(People(UBound(People)).Name) > 0 Then
                    CheckEvent "OnAvoidListAppear", "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0) & "job=" & People(UBound(People)).Class & Chr(0) & "posX=" & People(UBound(People)).Pos.Y & Chr(0) & "posY=" & People(UBound(People)).Pos.X & Chr(0) & "distance=" & EvalNorm(People(UBound(People)).Pos, curPos)
                ElseIf isWarpList(People(UBound(People)).Name) And Len(People(UBound(People)).Name) > 0 Then
                    CheckEvent "OnWarpListAppear", "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0) & "job=" & People(UBound(People)).Class & Chr(0) & "posX=" & People(UBound(People)).Pos.Y & Chr(0) & "posY=" & People(UBound(People)).Pos.X & Chr(0) & "distance=" & EvalNorm(People(UBound(People)).Pos, curPos)
                Else
                    CheckEvent "OnPlayerAppear", "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0) & "job=" & People(UBound(People)).Class & Chr(0) & "posX=" & People(UBound(People)).Pos.Y & Chr(0) & "posY=" & People(UBound(People)).Pos.X & Chr(0) & "distance=" & EvalNorm(People(UBound(People)).Pos, curPos)
                End If
                'CharHair
                Winsock_SendPacket IntToChr(&H94) + Mid(inData, 3, 4), True
                ReDim Preserve People(UBound(People) + 1)
                'If frmPeople.Visible = True Then
                    UpdatePeople
                'End If
        End If
skip2:
Decode_0079 = ""
Exit Function
errie:
Decode_0079 = "ERROR!!! [Decode_0079] " & Err.Description
Err.Clear
End Function

Function Decode_007B(inData As String) As String
On Error GoTo errie
        'print_packet Left(InData, 58)
        Dim NameID As Integer
        Dim X As Integer
        NameID = MakePort(Mid(inData, 15, 2))
        Dim tmpcoord As Coord
        Monstmppos = MakeCoordsSec(Mid(inData, 53, 3))
        'TmpCoord = MakeCoordsSec(Mid(InData, 53, 3))
        tmpcoord = MakeCoords(Mid(inData, 51, 3))
        Dim IsPet As Boolean
        
        If NameID > 1000 Then GoTo skip3
        If MakePort(Mid(inData, 15, 2)) < &H389 Then
            AI_AvoidID Mid(inData, 3, 4)
            If UBound(People) = 0 Then GoTo ppl
            For X = 0 To UBound(People) - 1
                If Mid(inData, 3, 4) = People(X).ID Then
                    Clear_Dot People(X).Pos
                    People(X).Pos = tmpcoord
                    People(X).NextPos = Monstmppos

                    'People(X).Shield = MakePort(Mid(inData, 21, 2))
                    'People(X).HeadB = MakePort(Mid(inData, 33, 2))
                    'People(X).HeadT = MakePort(Mid(inData, 29, 2))
                    'People(X).HeadM = MakePort(Mid(inData, 31, 2))
                    'People(X).HairC = MakePort(Mid(inData, 33, 2))
                    'People(X).Weapon = MakePort(Mid(inData, 19, 2))
                    'People(X).Hair = MakePort(Mid(inData, 17, 2))
                    
                    People(X).Time = GetTickCount
                    People(X).Speed = MakePort(Mid(inData, 7, 2))
                    If Not Disable_frmPeople Then Plot_Dot People(X).Pos, PColor
                    'If frmPeople.Visible = True Then
                        UpdatePeople
                    'End If
                    If IsAvoidID(Mid(inData, 3, 4)) Then
                        CheckEvent "OnGMMove", "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0) & "job=" & People(X).Class & Chr(0) & "startX=" & People(X).Pos.Y & Chr(0) & "startY=" & People(X).Pos.X & Chr(0) & "endX=" & People(X).NextPos.Y & Chr(0) & "endY=" & People(X).NextPos.X
                    ElseIf IsAvoid(People(X).Name) And Len(People(X).Name) > 0 Then
                        CheckEvent "OnAvoidListMove", "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0) & "job=" & People(X).Class & Chr(0) & "startX=" & People(X).Pos.Y & Chr(0) & "startY=" & People(X).Pos.X & Chr(0) & "endX=" & People(X).NextPos.Y & Chr(0) & "endY=" & People(X).NextPos.X
                    ElseIf isWarpList(People(X).Name) And Len(People(X).Name) > 0 Then
                        CheckEvent "OnWarpListMove", "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0) & "job=" & People(X).Class & Chr(0) & "startX=" & People(X).Pos.Y & Chr(0) & "startY=" & People(X).Pos.X & Chr(0) & "endX=" & People(X).NextPos.Y & Chr(0) & "endY=" & People(X).NextPos.X
                    Else
                        CheckEvent "OnPlayerMove", "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0) & "job=" & People(X).Class & Chr(0) & "startX=" & People(X).Pos.Y & Chr(0) & "startY=" & People(X).Pos.X & Chr(0) & "endX=" & People(X).NextPos.Y & Chr(0) & "endY=" & People(X).NextPos.X
                    End If
                    GoTo endcase
                End If
            Next
ppl:
                People(UBound(People)).ID = Mid(inData, 3, 4)
                People(UBound(People)).Pos = tmpcoord
                People(UBound(People)).NextPos = Monstmppos
                If Mid(inData, 3, 4) = FollowMode.AID Then move_to People(UBound(People)).NextPos, 2
                People(UBound(People)).Speed = MakePort(Mid(inData, 7, 2))
                People(UBound(People)).Time = GetTickCount
                'People(UBound(People)).Shield = MakePort(Mid(inData, 27, 2))
                'People(UBound(People)).HeadB = MakePort(Mid(inData, 21, 2))
                'People(UBound(People)).HeadT = MakePort(Mid(inData, 29, 2))
                'People(UBound(People)).HeadM = MakePort(Mid(inData, 31, 2))
                'People(UBound(People)).HairC = MakePort(Mid(inData, 33, 2))
                'People(UBound(People)).Weapon = MakePort(Mid(inData, 19, 2))
                'People(UBound(People)).Hair = MakePort(Mid(inData, 17, 2))
                If Not Disable_frmPeople Then Plot_Dot People(UBound(People)).Pos, PColor
                People(UBound(People)).NameID = MakePort(Mid(inData, 15, 2))
                If IsPet And NameID > 1000 Then
                    For X = 0 To UBound(Monsters)
                    If MakePort(Mid(inData, 15, 2)) = Monsters(X).ID Then
                        People(UBound(People)).Class = Monsters(X).Name
                        Exit For
                    End If
                    Next
                    People(UBound(People)).Sex = " <Pet>"
                Else
                    People(UBound(People)).Class = Return_Class(People(UBound(People)).NameID)
                    If Asc(Mid(inData, 50, 1)) = 0 Then
                        People(UBound(People)).Sex = " <F>"
                    Else
                        People(UBound(People)).Sex = " <M>"
                    End If
                End If
                'print_packet Left(InData, 58), People(UBound(People)).nameid
                If IsAvoidID(Mid(inData, 3, 4)) Then
                    CheckEvent "OnGMMove", "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0) & "job=" & People(UBound(People)).Class & Chr(0) & "startX=" & People(UBound(People)).Pos.Y & Chr(0) & "startY=" & People(UBound(People)).Pos.X & Chr(0) & "endX=" & People(UBound(People)).NextPos.Y & Chr(0) & "endY=" & People(UBound(People)).NextPos.X
                ElseIf IsAvoid(People(UBound(People)).Name) And Len(People(UBound(People)).Name) > 0 Then
                    CheckEvent "OnAvoidListMove", "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0) & "job=" & People(UBound(People)).Class & Chr(0) & "startX=" & People(UBound(People)).Pos.Y & Chr(0) & "startY=" & People(UBound(People)).Pos.X & Chr(0) & "endX=" & People(UBound(People)).NextPos.Y & Chr(0) & "endY=" & People(UBound(People)).NextPos.X
                ElseIf isWarpList(People(UBound(People)).Name) And Len(People(UBound(People)).Name) > 0 Then
                    CheckEvent "OnWarpListMove", "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0) & "job=" & People(UBound(People)).Class & Chr(0) & "startX=" & People(UBound(People)).Pos.Y & Chr(0) & "startY=" & People(UBound(People)).Pos.X & Chr(0) & "endX=" & People(UBound(People)).NextPos.Y & Chr(0) & "endY=" & People(UBound(People)).NextPos.X
                Else
                    CheckEvent "OnPlayerMove", "name=" & Get_PeopleName(Mid(inData, 3, 4)) & Chr(0) & "job=" & People(UBound(People)).Class & Chr(0) & "startX=" & People(UBound(People)).Pos.Y & Chr(0) & "startY=" & People(UBound(People)).Pos.X & Chr(0) & "endX=" & People(UBound(People)).NextPos.Y & Chr(0) & "endY=" & People(UBound(People)).NextPos.X
                End If
                Winsock_SendPacket IntToChr(&H94) & Mid(inData, 3, 4), True
                ReDim Preserve People(UBound(People) + 1)
                'If frmPeople.Visible = True Then
                UpdatePeople
                'End If
                GoTo endcase
        End If
        Dim found As Boolean
skip3:
        found = False
        islist = False
        IsPet = False
        If Asc(Mid(inData, 17, 1)) > 0 Then IsPet = True
        For X = 0 To UBound(Attack)
            If MakePort(Mid(inData, 15, 2)) = Attack(X).ID Then
                found = True
                Exit For
            End If
        Next
        Dim tmpname As String
        For X = 0 To UBound(Monsters)
            If MakePort(Mid(inData, 15, 2)) = Monsters(X).ID Then
                islist = True
                tmpname = Monsters(X).Name
                Exit For
            End If
        Next
        If Mid(inData, 3, 4) = MyPet.ID And MyPet.Type = "" Then MyPet.Type = tmpname
        Dim found2 As Boolean
        'If islist Then
            found2 = False
            If UBound(MonsterList) > 0 Then
                For X = 0 To UBound(MonsterList) - 1
                    If Mid(inData, 3, 4) = MonsterList(X).ID Then
                        found2 = True
                        If Not found Then MonsterList(X).NoAttack = True
                        If IsPet Then MonsterList(X).IsPet = True
                        Clear_Dot MonsterList(X).Pos
                        MonsterList(X).Pos = MakeCoords(Mid(inData, 51, 3))
                        If MyPet.ID = MonsterList(X).ID Then
                            Plot_Dot MonsterList(X).Pos, 16711935
                        ElseIf CurAtkMonster.ID <> MonsterList(X).ID Then
                            Plot_Dot MonsterList(X).Pos, vbRed
                        Else
                            Plot_Dot MonsterList(X).Pos, CurAtkColor
                        End If
                        MonsterList(X).NextPos = Monstmppos
                        MonsterList(X).Time = GetTickCount()
                        MonsterList(X).Endtime = MonsterList(UBound(MonsterList)).Time + (EvalNorm(tmpcoord, Monstmppos) * MakePort(Mid(inData, 7, 2)))
                        MonsterList(X).StatusA = MakePort(Mid(inData, 9, 2))
                        MonsterList(X).StatusB = MakePort(Mid(inData, 11, 2))
                        MonsterList(X).IsAttack = False
                        MonsterList(X).Speed = MakePort(Mid(inData, 7, 2))
                        If (MakePort(Mid(inData, 9, 2)) > 0 And MakePort(Mid(inData, 9, 2)) < 5) Or MakePort(Mid(inData, 11, 2)) > 0 Then MonsterList(X).IsTrap = True
                        If CanGO(curPos, MonsterList(X).Pos) Then MonsterList(X).CantGo = True
                        upd_frmMonster
                        Exit For
                    End If
                Next
            End If
            If (Not Sitting) And (Not Pickup) And (Not InFight) And (CurAtkMonster.NameID > 0) Then SendAction = True
            'MonsterTime = 0
            'tmrMonsterUpdate.Enabled = True
            

                If Mid(inData, 3, 4) = CurAtkMonster.ID Then
                    'CurAtkMonster.pos = MonsTmpPos
                    CurAtkMonster.Pos = MakeCoords(Mid(inData, 51, 3))
                    CurAtkMonster.NextPos = Monstmppos
                    If IsPet Then
                        Stat "This's a pet!, abort target..." & vbCrLf
                        Clear_This_Mons 0
                        GoTo endifcur
                    End If
                        CurAtkMonster.Time = GetTickCount()
                        CurAtkMonster.Endtime = CurAtkMonster.Time + (EvalNorm(tmpcoord, Monstmppos) * CurAtkMonster.Speed)
                        'TmrMonsMove.Interval = EvalNorm(tmpcoord, Monstmppos) * MakePort(Mid(InData, 7, 2))
                        'TmrMonsMove.Enabled = True

                    CurAtkMonster.Speed = MakePort(Mid(inData, 7, 2))
endifcur:
                upd_curMonster
                End If
         
            
            If Not found2 Then
                If Not found Then MonsterList(UBound(MonsterList)).NoAttack = True
                If IsPet Then MonsterList(UBound(MonsterList)).IsPet = True
                MonsterList(UBound(MonsterList)).ID = Mid(inData, 3, 4)
                MonsterList(UBound(MonsterList)).Pos = MakeCoords(Mid(inData, 51, 3))
                If MyPet.ID = MonsterList(UBound(MonsterList)).ID And MyPet.Name <> "" Then
                    Plot_Dot MonsterList(UBound(MonsterList)).Pos, 16711935
                Else
                    Plot_Dot MonsterList(UBound(MonsterList)).Pos, vbRed
                End If
                MonsterList(UBound(MonsterList)).NextPos = Monstmppos
                MonsterList(UBound(MonsterList)).IsAttack = False
                MonsterList(UBound(MonsterList)).Time = GetTickCount()
                MonsterList(UBound(MonsterList)).Endtime = MonsterList(UBound(MonsterList)).Time + (EvalNorm(tmpcoord, Monstmppos) * MakePort(Mid(inData, 7, 2)))
                MonsterList(UBound(MonsterList)).NameID = MakePort(Mid(inData, 15, 2))
                MonsterList(UBound(MonsterList)).StatusA = MakePort(Mid(inData, 9, 2))
                MonsterList(UBound(MonsterList)).StatusB = MakePort(Mid(inData, 11, 2))
                If (MakePort(Mid(inData, 9, 2)) > 0 And MakePort(Mid(inData, 9, 2)) < 5) Or MakePort(Mid(inData, 11, 2)) > 0 Then MonsterList(UBound(MonsterList)).IsTrap = True
                MonsterList(UBound(MonsterList)).Speed = MakePort(Mid(inData, 7, 2))
                MonsterList(UBound(MonsterList)).IsAttack = False
                If islist Then MonsterList(UBound(MonsterList)).Name = tmpname
                If IsPet Then Winsock_SendPacket IntToChr(&H94) & Mid(inData, 3, 4), True
                ReDim Preserve MonsterList(UBound(MonsterList) + 1)
                upd_frmMonster
            End If
            'If (CurAtkMonster(NumberMons).NameID = 0) And (Not Sitting) And (Not Pickup) And (Not IsDamage) Then
            '    Pickup = True
            '    tmrPickup.Enabled = True
            'End If
        
               
        'End If
        If Not islist And NameID > 20 And Not IsPet Then
            Stat "Unknow Monster " & MakeHexName(Mid(inData, 15, 2)) & vbCrLf
            Winsock_SendPacket IntToChr(&H94) & Mid(inData, 3, 4), True
        End If
endcase:
Decode_007B = ""
Exit Function
errie:
Decode_007B = "ERROR!!! [Decode_007B] " & Err.Description
Err.Clear
End Function


Function Decode_0095(inData As String) As String
On Error GoTo errie
    Dim X As Integer, Y As Integer
    Dim found As Boolean
    
    Select Case MakePort(Mid(inData, 3, 4))
        Case 0 To 100000 'monster/npc
            For X = 0 To UBound(myNPC)
                If Mid(inData, 3, 4) = myNPC(X).ID And myNPC(X).Name = "" Then
                    myNPC(X).Name = Trim(MakeString(Mid(inData, 7, 24)))
                    myNPC(X).Name = Replace(myNPC(X).Name, " ", "_")
                    Save_NPC
                End If
            Next
            If UBound(NPCList) > 0 Then
                For X = 0 To UBound(NPCList) - 1
                    If Mid(inData, 3, 4) = NPCList(X).ID Then
                        NPCList(X).Name = MakeString(Mid(inData, 7, 24))
                        If frmNPC.Visible Then UpdateNPC
                        GoTo end95
                    End If
                Next
            End If
            If UBound(MonsterList) > 0 Then
            For X = 0 To UBound(MonsterList) - 1
                If Mid(inData, 3, 4) = MonsterList(X).ID Then
                    MonsterList(X).Name = MakeString(Mid(inData, 7, 24))
                    If MonsterList(X).IsPet Then
                        GoTo end95
                    End If
                    For Y = 0 To UBound(Monsters)
                        If MonsterList(X).NameID = Monsters(Y).ID Then
                            found = True
                            Exit For
                        End If
                    Next
                    If Not found Then
                        ReDim Preserve Monsters(UBound(Monsters) + 1)
                        Dim tsname As String
                        Monsters(UBound(Monsters)).Name = MakeString(Mid(inData, 7, 24))
                        Monsters(UBound(Monsters)).ID = MonsterList(X).NameID
                        Open App.Path & "\table\data.txt" For Append As #9
                        Print #9, "0" & Hex(MonsterList(X).NameID) & " " & MakeString(Mid(inData, 7, 24))
                        Chat "Updated [0" & Hex(MonsterList(X).NameID) & " " & MakeString(Mid(inData, 7, 24)) & "] to 'table\data.txt'", vbRed
                        Close 9
                        Load_Monster
                    End If
                    found = False
                    For Y = 0 To UBound(Attack)
                        If MonsterList(X).Name = Attack(Y).Name Or (Attack(Y).Name = "/" & MonsterList(X).Name) Then
                            found = True
                            Exit For
                        End If
                    Next
                    If Not found Then
                        ReDim Preserve Attack(UBound(Attack) + 1)
                        Attack(UBound(Attack)).Name = MakeString(Mid(inData, 7, 24))
                        Attack(UBound(Attack)).ID = MonsterList(X).NameID
                        Open App.Path & "\control\attack.txt" For Append As #9
                        Print #9, MakeString(Mid(inData, 7, 24))
                        Chat "Updated [" & MakeString(Mid(inData, 7, 24)) & "] to 'Attack.txt'"
                        Close 9
                        Load_Attack
                        found = False
                    End If
                End If
            Next
            End If

        Case Else 'people
            If UBound(People) > 0 Then
                AI_AvoidID Mid(inData, 3, 4)
                For X = 0 To UBound(People) - 1
                    If Mid(inData, 3, 4) = People(X).ID Then
                        People(X).Name = MakeString(Mid(inData, 7, 24))
                        upd_frmPeople
                        AI_Avoid People(X).Name
                    End If
                Next
            End If
    End Select

end95:
Decode_0095 = ""
Exit Function
errie:
Decode_0095 = "ERROR!!! [Decode_0095] " & Err.Description
Err.Clear
End Function

Function Decode_0195(inData As String) As String
On Error GoTo errie
    Dim X As Long
    If UBound(NPCList) > 0 Then
        For X = 0 To UBound(NPCList) - 1
            If Mid(inData, 3, 4) = NPCList(X).ID Then
                NPCList(X).Name = MakeString(Mid(inData, 7, 24))
                If frmNPC.Visible Then UpdateNPC
                Exit Function
            End If
        Next
    End If
    If UBound(People) > 0 Then
        For X = 0 To UBound(People) - 1
            If Mid(inData, 3, 4) = People(X).ID Then
                People(X).Name = MakeString(Mid(inData, 7, 24))
                AI_Avoid People(X).Name
                upd_frmPeople
                Exit Function
            End If
        Next
    End If
Decode_0195 = ""
Exit Function
errie:
Decode_0195 = "ERROR!!! [Decode_0195] " & Err.Description
Err.Clear
End Function

