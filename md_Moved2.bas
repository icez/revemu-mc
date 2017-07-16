Attribute VB_Name = "md_Moved2"
Option Explicit

Public IsLogin As Boolean
Public TeleportDelay As Integer
Public RestDelay As Integer
Public StandDelay As Integer
Public DelayCheckRest As Integer
Public ActionDelay As Integer
Public DelaymoveCounter As Long
Public UseWingYet As Boolean
Public WarpSaveCount As Byte
Public LastID As String
Public SendSP As Boolean
Public oldSelectPos As Coord
Public noMoveCounter As Integer
Public Casting As Boolean
Public LastMonsID As String
Public CurNPC As String
Public OldNextBEXP As Long
Public OldNextJXP As Long
Public CurEXPMons As Long
Public CurJXPMons As Long
Public DeadMonsName As String
Public ProcessID As Long
Public RandomPoint As Long
Public MonsterMoveTime As Long
Public PlayerEndMoveTime As Long
Public DisPos As Coord
Public tmpdata As String
Public mSkillDelay As Long
Public DelayuseCC As Byte
Public DelayuseFC As Byte
Public ShopStep As Byte
Public MaxShopAmount As Long
Public TmrBackTown_Enabled As Boolean
Public PRange As Integer

Public isulandskill As Boolean
Public IsStanding As Boolean
Public IsSitting As Boolean
Public CryptOn As Boolean
Public RouteCounter As Integer
Public KSID As String * 4
Public DetectPortal As Boolean

Public BackWpCounter As Integer
Public BackWP As Boolean
Public BlockMove As Boolean
Public Monstmppos As Coord
Public tmpPos As Coord
Public OnRoute As Boolean
Public TextDebug As String
Public Sending As Boolean
Public SendSkillMob As Boolean
Public SellNPC As MonsterPos
Public Aggro() As MonsList
Public Isweight1 As Boolean
Public Isweight2 As Boolean
Public Tracing As Boolean
Public Pickuptime As Integer
Public TryPicktime As Integer
Public TraceMons As Boolean
Public GotItem As Boolean
Public countMovewait As Byte
Public MoveWait As Boolean
Public IsRandommove As Boolean
Public SpellCounter As Integer
Public CurItem As Itemname
Public ChopErrorCounter As Integer
Public SkillCounter As Integer
Public SPBound As Integer
Public StartBot As Boolean
Public StopAction As Boolean
Public StartMap As String
Public CurrentMap As String
Public OldMap As String
Public SHour As Integer
Public SMin As Integer
Public SSec As Integer
Public SessionTime As String
Public DamageCounter As Integer
Public Wait As Boolean
Public oldBaseEXP As Long
Public oldJobEXP As Long
Public AttackCounter As Integer
Public ConnState As Integer
Public Oldstate As Integer
Public RecvData As String
Public NumberMons As Integer
Public UsePotCounter As Integer
Public UseHeal As Boolean
Public SessionEXP As Long
Public SessionJEXP As Long

Public ArcherMode As Boolean

Public MyCheck As String

Public StartPos As Coord
Public DeadMonPos As Coord
Public OldPos As Coord
Public WarpPos As Coord
Public ReSetCounter As Integer

Public Sex As Byte
Public islist As Boolean
Public MonsterList() As MonsterPos
Public ClearCounter As Integer
Public CurAtkMonster As MonsterPos
Public NextAtkMonster As MonsterPos
Public Pickup As Boolean
Public InFight As Boolean
Public Items() As ItemPos
Public Connected As Boolean
Public CharNameStart As String
Public CurItemID As String * 4
Public ItemData As String
Public DError As Boolean
Public MakeDamage As Boolean
Public DealtDamage As Boolean
Public CurMonsterName As String
Public IsDamage As Boolean
Public DelayATT As Boolean
Public IsLock As Boolean
Public IsAggro As Boolean
Public Weapon As String
Public NoWalk As Boolean
Public ResponseOK As Boolean
Public CurrentItem As Itemname
Public GotCurItem As Boolean
Public UseArrow As Boolean
Public UseBow As Boolean
Public WeaponName As String
Public SendSell As Boolean
Public SendHeal As Boolean
Public IsSell As Boolean
Public SendUsePot As Boolean
Public SkillWait As Boolean
Public SkillDelay As Integer
Public isWarp As Boolean
Public ResponseCounter As Long
Public LoopWait As Boolean
Public LoopTime As Integer
Public MonsterTime As Integer
Public WarpNumber As Integer
Public tmpSpeed As Integer

Public NomonsTimeCount As Integer

Public checkadmin As String
Public checkvalid As String
Public lastPacket As String

'mc 0.2 variable
Public MODTradeDelay As Long
Public MODTradeStep As Byte

Sub checkKS(Src As String, Target As String)
On Error GoTo errie
    'New Add to Prevent Steal Kill
    Dim X As Integer
    If UBound(MonsterList) > 0 Then
        For X = 0 To UBound(MonsterList) - 1
            If MonsterList(X).ID = Src Then
                    If Target <> Src And MakePort(Target) > 65535 Then MonsterList(X).TargetID = Target
                    If Target = AccountID Then MonsterList(X).IsAttack = False
            End If
            If MonsterList(X).ID = Target Then
                If AccountID <> Src And MakePort(Src) > 65535 And Not InParty(Src) Then
                    If (CurAtkMonster.ID = MonsterList(X).ID And Not MakeDamage And (MonsterList(X).TargetID <> AccountID)) And Mid(Get_PeopleName(Src), 1, 2) <> "U:" Then
                        If (Not killsteal) Then
                            Clear_This_Mons 0
                            Stat "[" & Get_PeopleName(Src) & "] already attack this monster!" + vbCrLf
                        Else
                            Stat "[" & Get_PeopleName(Src) & "] already attack this monster but who cares!" + vbCrLf
                        End If
                        MonsterList(X).IsAttack = True
                    ElseIf KSID <> Src And (MonsterList(X).TargetID = AccountID Or (CurAtkMonster.ID = MonsterList(X).ID And MakeDamage)) And Not InParty(Src) Then
                        Chat "[" & Get_PeopleName(Src) & "] attack your monster!", MColor.Fail
                        KSID = Src
                        response_mode = 1
                        frmMain.tmrChatResponse.Interval = RandomNumber(1500, 3000)
                        frmMain.tmrChatResponse.Enabled = True
                    End If
                    If CurAtkMonster.ID <> MonsterList(X).ID And Not InParty(Src) Then
                        MonsterList(X).IsAttack = True
                    End If
                End If
            ElseIf MonsterList(X).TargetID <> AccountID And MonsterList(X).TargetID <> MonsterList(X).ID And Mid(Get_PeopleName(Src), 1, 2) <> "U:" And Not InParty(Src) Then
                If (CurAtkMonster.ID = MonsterList(X).ID) Then
                    If (Not killsteal) Then
                        Clear_This_Mons 0
                        Stat "Current monster already attack to [" & Get_PeopleName(Src) & "]!" + vbCrLf
                    Else
                        Stat "Current monster already attack to [" & Get_PeopleName(Src) & "]! but who cares!" + vbCrLf
                    End If
                End If
                MonsterList(X).IsAttack = True
            End If
            upd_curMonster
            If CurAtkMonster.NameID = 0 Then
                KSID = Chr(&HFF) & Chr(&HFF) & Chr(&HFF) & Chr(&HFF)
            End If
        Next
    End If
    'End of prevent steal kill
Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "checkKS", Err.number, Err.Description
    Err.Clear
End Sub

Sub Clear_This_Mons(X As Integer)
On Error GoTo errie
    If X = 1 Then
        'Stat "Your target, [" + CurMonsterName + "]" + " Dead..." + vbCrLf
        LastMonsID = CurAtkMonster.ID
    End If
    If IsLock Or InFight Then ModIncMonLog CurAtkMonster.Name
    DeadMonPos = CurAtkMonster.Pos
    TraceMons = False
    PlayerMoveTime = 0
    frmMain.TmrMove.Enabled = False
    BlockMove = False
    BackWP = False
    SkillCounter = 0
    DamageCounter = 0
    InFight = False
    Sitting = False
    ReDim Route(0)
    Current = 0
    CurAtkMonster.NameID = 0
    CurAtkMonster.ID = String(4, Chr(0))
    CurAtkMonster.Pos.X = 0
    CurAtkMonster.Pos.Y = 0
    CurAtkMonster.NextPos.X = 0
    CurAtkMonster.NextPos.Y = 0
    CurAtkMonster.Speed = 0
    AttCounter = 0
    frmMain.tmrAggro.Enabled = True
    IsLock = False
    IsDamage = False
    CurMonsterName = "None"
    frmMain.labCurMons.Caption = "[None]"
    
    If (Not Pickup) And MakeDamage And (Check_Pickup) And X = 1 Then
        Pickup = True
        frmMain.tmrPickDelay.Enabled = True
        frmMain.tmrPickup.Enabled = False
        frmMain.tmrPickup.Interval = TimePickup
        frmMain.tmrPickup.Enabled = True
    End If
    
    MakeDamage = False
    NomonsTimeCount = 0
Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "Clear_This_Mons", Err.number, Err.Description
    Err.Clear
End Sub


Function Get_PeopleName(ByVal ID As String) As String
On Error GoTo errie
    Dim i As Integer
    If UBound(People) = 0 Then GoTo EndFunc
    For i = 0 To UBound(People) - 1
        If ID = People(i).ID And People(i).Name <> "" Then
            Get_PeopleName = People(i).Name
            Exit Function
        End If
    Next
EndFunc:
    Get_PeopleName = "U:" & ChrtoHex(ID)
Exit Function
errie:
    If Err.number > 0 Then print_funcerr "Get_PeopleName", Err.number, Err.Description
    Err.Clear
End Function

Function IsItem(ID As String) As String
On Error GoTo errie
    If UBound(Items) = 0 Then GoTo EndFunc
    Dim i As Integer
    For i = 0 To UBound(Items) - 1
        If Items(i).ID = ID Then
            IsItem = Return_ItemName(Items(i).Name)
            Exit Function
        End If
    Next
EndFunc:
    IsItem = ""
Exit Function
errie:
    If Err.number > 0 Then print_funcerr "IsItem", Err.number, Err.Description
    Err.Clear
End Function


Function Get_MonsName(ByVal ID As String) As String
On Error GoTo errie
    Dim i As Integer
    If UBound(MonsterList) = 0 Then GoTo EndFunc
    For i = 0 To UBound(MonsterList) - 1
        If ID = MonsterList(i).ID And MonsterList(i).Name <> "" Then
            Get_MonsName = MonsterList(i).Name
            Exit Function
        End If
    Next
EndFunc:
    Get_MonsName = "Unknow"
    Exit Function
errie:
    If Err.number > 0 Then print_funcerr "Get_MonsName", Err.number, Err.Description
    Err.Clear
End Function

Function isPlayer(ID As String) As Boolean
On Error GoTo errie
    Dim i As Integer
    If UBound(People) = 0 Then GoTo EndFunc
    For i = 0 To UBound(People)
        If ID = People(i).ID Then
            isPlayer = True
            Exit Function
        End If
    Next
EndFunc:
    isPlayer = False
Exit Function
errie:
    If Err.number > 0 Then print_funcerr "isPlayer", Err.number, Err.Description
    Err.Clear
End Function

Function HaveRare() As Boolean
On Error GoTo errie
    If UBound(Items) = 0 Then GoTo endloop:
    Dim i As Integer
    For i = 0 To UBound(Items) - 1
        If isRare(Items(i).Name) Then
            HaveRare = True
            Exit Function
        End If
    Next
endloop:
    HaveRare = False
Exit Function
errie:
    If Err.number > 0 Then print_funcerr "HaveRare", Err.number, Err.Description
    Err.Clear
End Function

Sub Check_Tele()
On Error GoTo errie
    If (DamageCounter > DamageSet) And (IsDamageDC) Then
        If (AutoDCCase = 0) Then
            Stat "Over " + CStr(DamageSet) + "  Damage and the bot do nothing, Auto-Teleport..." + vbCrLf
            TraceMons = False
            Teleport
        Else
            Stat "Over " + CStr(DamageSet) + "  Damage and the bot do nothing, Auto-Disconnect..." + vbCrLf
            frmMain.ResettoReCon
        End If
        Exit Sub
    End If
    
    If (IsAutoDC) And (Players(number).HP < (Players(number).MaxHP * HPDC)) And (InFight) Then
        Stat "You're nearly to dead, waiting..." + vbCrLf
        DamageCounter = 0
        If (AutoDCCase = 0) Then
            TraceMons = False
            Teleport
        Else
            frmMain.ResettoReCon
        End If
        Exit Sub
    End If
Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "Check_Tele", Err.number, Err.Description
    Err.Clear
End Sub


Sub Teleport()
On Error Resume Next
    TraceMons = False
    BlockMove = False
    Dim tstr As String
    If Sitting Then frmMain.Send_Stand
    Dim i&
    If WaitEquipTele Then Exit Sub
    If Find_Item("Fly_Wing") < 1 And Not WaitEquipBack Then
        tmpEQTelePos = 0
        tmpEQOldPos = 0
        For i = 0 To UBound(AllInv)
            If InStr(LCase(AllInv(i).Name), "teleport ") > 0 And AllInv(i).Pos < 1 Then
                tmpEQTelePos = i
                tmpEQTeleName = AllInv(i).Name
            End If
            If AllInv(i).Pos = 8 Then
                tmpEQOldPos = i
                tmpEQOldName = AllInv(i).Name
            End If
        Next
        If tmpEQTelePos > 0 Then
            Stat "Found [" & AllInv(tmpEQTelePos).Name & "] for teleportation. Equipping." & vbCrLf, 0, False, True
            WaitEquipTele = True
            If tmpEQOldPos > 0 Then frmMain.Send_unEquip tmpEQOldPos
            frmMain.Send_Equip tmpEQTelePos
            Exit Sub
        End If
    End If

    'If AutoWing And (Find_Item("Fly_Wing") > 0) And Not UseWingYet Then
    If find_skill("AL_TELEPORT") Then
        If ForceTeleport Then
            Stat "Found teleport skill..." + vbCrLf
            tstr = Chr(&H1B) & Chr(1) & IntToChr(&H1A) & "Random" & String(10, Chr(0))
            Winsock_SendPacket tstr, True
            TeleportDelay = 3
        Else
            Stat "Found teleport skill, using ..." + vbCrLf
            Winsock_SendPacket Chr(&H13) & Chr(1) & IntToChr(&H1) & IntToChr(&H1A) & AccountID, True
            TeleportDelay = 3
        End If
    ElseIf Find_Item("Fly_Wing") > 0 And TeleportDelay = 0 Then
        Stat "Auto Use Fly_Wing 1 EA..." + vbCrLf
        Winsock_SendPacket IntToChr(&HA7) & IntToChr(Find_Item("Fly_Wing")) & AccountID, True
        UseWingYet = True
        Winsock_SendPacket Chr(&H13) & Chr(1) & IntToChr(&H1) & IntToChr(&H1A) & AccountID, True
        TeleportDelay = 3
    End If
End Sub


Public Function NewCanGO(pt1 As Coord, Pt As Coord) As Boolean
    Dim Des As Coord
    Dim Src As Coord
    Dim count As Integer
    Src = pt1
    Des = Pt
    count = 0
    Do
        If aiMap(Src.X + Sgn(Des.X - Src.X), Src.Y) = 0 Then Src.X = Src.X + Sgn(Des.X - Src.X)
        If aiMap(Src.X, Src.Y + Sgn(Des.Y - Src.Y)) = 0 Then Src.Y = Src.Y + Sgn(Des.Y - Src.Y)
        count = count + 1
        If count > 500 Then
            NewCanGO = False
            Exit Function
        End If
        If Src.X = Des.X And Src.Y = Des.Y Then GoTo endloop
        If Src.X = Des.X Then
            If aiMap(Src.X, Src.Y + Sgn(Des.Y - Src.Y)) <> 0 Then
                NewCanGO = False
                Exit Function
            End If
        End If
        If Src.Y = Des.Y Then
           If aiMap(Src.X + Sgn(Des.X - Src.X), Src.Y) <> 0 Then
                NewCanGO = False
                Exit Function
            End If
        End If
endloop:
    Loop While (Src.X <> Des.X Or Src.Y <> Des.Y)
    NewCanGO = True
End Function

