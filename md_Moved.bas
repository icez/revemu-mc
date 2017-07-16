Attribute VB_Name = "md_Moved"
Option Explicit

'ft devil copy
Public BlockMsg As Boolean
Public UseChain As Boolean
Public UseFinish As Boolean
Public LoginParty As Boolean
Public Sitting As Boolean

Private Inc_Pic As Integer
Public tmpPicStatus() As lPicType
Public NoStoreItem As Boolean
Type lPicType
    isLoaded As Boolean
    Picture As IPictureDisp
End Type

Public Sub AddPicStatus(pic As Long)
On Error GoTo errie
    If FileExists(App.Path & "\effect\" & pic & ".gif") = False Then Exit Sub
    If UBound(tmpPicStatus) < pic Then
        ReDim Preserve tmpPicStatus(pic)
    End If
    If tmpPicStatus(pic).isLoaded = False Then
        Set tmpPicStatus(pic).Picture = LoadPicture(App.Path & "\effect\" & pic & ".gif")
        tmpPicStatus(pic).isLoaded = True
    End If
    For Inc_Pic = 0 To MDIfrmMain.imgStatus.UBound
        If MDIfrmMain.imgStatus(Inc_Pic).Picture = tmpPicStatus(pic).Picture Then Exit Sub
        If (MDIfrmMain.imgStatus(Inc_Pic).Picture = Empty) Or (MDIfrmMain.imgStatus(Inc_Pic).Picture = Null) Then
            MDIfrmMain.imgStatus(Inc_Pic).Picture = tmpPicStatus(pic).Picture
            Exit Sub
        End If
    Next
'
'    Load MDIfrmMain.imgStatus(MDIfrmMain.imgStatus.UBound + 1)
'    MDIfrmMain.imgStatus(MDIfrmMain.imgStatus.UBound).Left = MDIfrmMain.imgStatus(MDIfrmMain.imgStatus.UBound - 1).Left + 540
'    MDIfrmMain.imgStatus(MDIfrmMain.imgStatus.UBound).Picture = tmpPicStatus(pic).Picture
'    MDIfrmMain.imgStatus(MDIfrmMain.imgStatus.UBound).Refresh
    Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "AddPicStatus", Err.number, Err.Description
    Err.Clear
End Sub
Public Sub DelPicStatus(pic As Long)
On Error GoTo errie
    If UBound(tmpPicStatus) < pic Then Exit Sub
    If tmpPicStatus(pic).isLoaded = False Then Exit Sub
    Dim a As Integer
    a = -1
    For Inc_Pic = 0 To MDIfrmMain.imgStatus.UBound
        If (MDIfrmMain.imgStatus(Inc_Pic).Picture = tmpPicStatus(pic).Picture) Then
            a = Inc_Pic
            Exit For
        End If
    Next
    If a > -1 Then
        For Inc_Pic = a To MDIfrmMain.imgStatus.UBound - 1
            MDIfrmMain.imgStatus(Inc_Pic).Picture = MDIfrmMain.imgStatus(Inc_Pic + 1).Picture
        Next
        'Unload MDIfrmMain.imgStatus(MDIfrmMain.imgStatus.UBound)
    End If
Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "DelPicStatus", Err.number, Err.Description
    Err.Clear
End Sub

Sub Update_AtkSkill()
On Error GoTo errie
    Dim i&, Y&
    For Y = 0 To UBound(Attack)
        i = Find_SkillId(Attack(Y).Spell1)
        If i > 0 Then
            If Mods.STSystem Then Stat "Updated skill attack : " & Attack(Y).Name & " [" & Attack(Y).Spell1 & " (" & Attack(Y).lv1 & ")/" & Attack(Y).UTime1 & "] ." & vbCrLf
            Attack(Y).Skill1 = Chr(&H13) + Chr(1) + Chr(Attack(Y).lv1) + Chr(0) + IntToChr(CLng(SkillChar(i).ID))
            Attack(Y).sp1 = Get_UseSPbyName(Attack(Y).Spell1, Attack(Y).lv1)
        End If
        i = Find_SkillId(Attack(Y).Spell2)
        If i > 0 Then
            If Mods.STSystem Then Stat "Updated skill attack : " & Attack(Y).Name & " [" & Attack(Y).Spell2 & " (" & Attack(Y).lv2 & ")/" & Attack(Y).UTime2 & "] .." & vbCrLf
            Attack(Y).Skill2 = Chr(&H13) + Chr(1) + Chr(Attack(Y).lv2) + Chr(0) + IntToChr(CLng(SkillChar(i).ID))
            Attack(Y).sp2 = Get_UseSPbyName(Attack(Y).Spell2, Attack(Y).lv2)
        End If
    Next
    Exit Sub
errie:
    print_funcerr "Update_AtkSkill", Err.number, Err.Description
    Err.Clear
End Sub

Sub Check_ActionTime()
    On Error GoTo errie
    If ModAI Then Exit Sub
    Dim cur_time As Date
    Dim nowsps As Double
    Dim i As Long
    nowsps = (CLng(Players(number).SP) * 100) / Players(number).maxsp
    cur_time = Int(GetTickCount() / 1000)
    If Sitting Then GoTo next_check:
    If IsSelfSkill And AutoSkill(0).Name <> "" And AutoSkill(0).Time <> 0 And DelaySelfSkill = 0 Then
        For i = 0 To UBound(AutoSkill)
            If Not find_skill(AutoSkill(i).Name) Then GoTo end_loop
            
            If (cur_time - AutoSkill(i).TimeCount >= AutoSkill(i).Time) Or _
                (AutoSkill(i).TimeCount = 0) Then
                If CanUseSkill(AutoSkill(i).Name, AutoSkill(i).Level, Players(number).SP) Then
                    If (AutoSkill(i).SPmin < nowsps And (AutoSkill(i).SPmax > nowsps Or AutoSkill(i).SPmax = 0)) Then
                        If (Not AutoSkill(i).Auto_reuse Or AutoSkill(i).StatusNum < 1) And ((IsInLock And (InStr(AutoSkill(i).Mode, "L") Or Len(AutoSkill(i).Mode) = 0)) Or (Not (IsInLock) And InStr(AutoSkill(i).Mode, "M"))) Then
                            Stat CStr(cur_time - AutoSkill(i).TimeCount) & " Time to use skill [" & AutoSkill(i).Name & "] lv." & AutoSkill(i).Level & vbCrLf
                            Send_Use_Skill AutoSkill(i).ID, AutoSkill(i).Level, AccountID
                            DelaySelfSkill = 5
                            GoTo next_check
                        End If
                    End If
                End If
            End If
end_loop:
        Next
    End If
next_check:
    If AutoItem.Auto And UBound(AllInv) > 0 And IsInLock Then
        If (cur_time - AutoItem.TimeCount >= AutoItem.Time) Or AutoItem.TimeCount = 0 Then
            Dim Index As Long
            Index = Find_HealItem(AutoItem.Name)
            If Index > 0 Then
                Stat "Time to use " + AllInv(Index).Name & "..." + vbCrLf
                Winsock_SendPacket IntToChr(&HA7) & IntToChr(Index) & AccountID, True
                AutoItem.TimeCount = cur_time
            End If
        End If
    End If
    
    Exit Sub
errie:
    Open App.Path & "\log\errorlog.txt" For Append As #1
    Print #1, " == Check_ActionTime " & Err.number & "(" & Date & ")@" & Time & " == "
    Print #1, Err.Description
    Close #1
    Err.Clear
    Exit Sub
End Sub
Sub Check_ActionSkill()
On Error GoTo errie
    Dim i As Long
    If Sitting Or ModAI Then Exit Sub
    'If Not IsInLock Then Exit Sub
    If UseAutoSpell Then
        If Find_SkillId("SA_AUTOSPELL") > 0 Then
            If UBound(CurStatus) > 64 Then
                If Not CurStatus(65).Active Then
                    If Find_SkillId(AutoSpell_Name) > 0 Then Winsock_SendPacket Chr(206) & Chr(1) & LngToChr(Find_SkillId(AutoSpell_Name)), True
                    'CE 01 13 00 00 00
                End If
            End If
        End If
    End If
    If (Not IsSelfSkill Or AutoSkill(0).Name = "" Or AutoSkill(0).Time = 0) Then Exit Sub
    Dim nowsps As Double
    nowsps = (CLng(Players(number).SP) * 100) / Players(number).maxsp
    If DelaySelfSkill <> 0 Then Exit Sub
    For i = 0 To UBound(AutoSkill)
        If find_skill(AutoSkill(i).Name) Then
            If AutoSkill(i).Auto_reuse And AutoSkill(i).StatusNum > 0 And _
            ((AutoSkill(i).lusetime = 0) Or (Timer - AutoSkill(i).lusetime > 3)) And _
            CurStatus(AutoSkill(i).StatusNum).Active = False And _
            ((IsInLock And (InStr(AutoSkill(i).Mode, "L") Or Len(AutoSkill(i).Mode) = 0)) Or (Not (IsInLock) And InStr(AutoSkill(i).Mode, "M"))) _
            Then
                If Players(number).SP >= AutoSkill(i).SPNeed Then
                    If (AutoSkill(i).SPmin <= nowsps And (AutoSkill(i).SPmax >= nowsps Or AutoSkill(i).SPmax = 0)) Then
                        'AutoSkill(i).lusetime = Timer
                        Stat "Time to use skill [" & AutoSkill(i).Name & "] lv." & AutoSkill(i).Level & vbCrLf
                        Send_Use_Skill AutoSkill(i).ID, AutoSkill(i).Level, AccountID
                        DelaySelfSkill = 5
                        Exit For
                    Else
                        Stat "Time to use skill [" & AutoSkill(i).Name & "] lv." & AutoSkill(i).Level & " but SP is not in range" & vbCrLf
                        DelaySelfSkill = 5
                    End If
                Else
                    Stat "Time to use skill [" & AutoSkill(i).Name & "] lv." & AutoSkill(i).Level & " but require more SP 3" & vbCrLf
                    'AutoSkill(i).lusetime = Timer
                    DelaySelfSkill = 5
                End If
            End If
        End If
    Next
Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "Check_ActionSkill", Err.number, Err.Description
    Err.Clear
End Sub

Public Function find_skill(ByVal tstr As String) As Boolean
    Dim X As Integer
    For X = 0 To UBound(SkillChar)
        If SkillChar(X).Name = tstr And SkillChar(X).MaxLV > 0 Then
                find_skill = True
                Exit Function
        End If
    Next
    find_skill = False
End Function

Public Sub Send_Use_Skill(ByVal SkillID As Integer, Lv As Byte, TargetID As String)
On Error GoTo errie
    Dim tstr As String
    tstr = Chr(&H13) & Chr(&H1) & IntToChr(CLng(Lv)) & IntToChr(CLng(SkillID)) & TargetID
    Winsock_SendPacket tstr, True
    Exit Sub
errie:
Err.Clear
End Sub

Function ConvPacketData(Packet As String) As String
    Dim tstr$, tsb$
    Dim X As Integer, Y As Integer
    tstr = ""
    For X = 1 To Len(Packet)
        tstr = tstr & ChrtoHex(Mid(Packet, X, 1)) & " "
        'If Asc(Mid(packet, X, 1)) < 16 Then tstr = tstr + "0"
        'tstr = tstr + Hex(Asc(Mid(packet, X, 1))) + " "
        If Asc(Mid(Packet, X, 1)) > 32 And Asc(Mid(Packet, X, 1)) < 127 Then tsb = tsb & Mid(Packet, X, 1) Else tsb = tsb & "."
        If X Mod 16 = 0 Then
            tstr = tstr & "    " & tsb & vbCrLf
            tsb = ""
        End If
    Next
    If (Len(Packet) Mod 16) > 0 Then
        For X = 1 To (16 - (Len(Packet) Mod 16))
            tstr = tstr & "   "
        Next
    End If
    tstr = tstr & "    " & tsb
    ConvPacketData = tstr
    '33-126
End Function

Public Function CheckNPC() As Boolean
On Error GoTo errie
    If frmMain.tmrDealNPC.Enabled Or UBound(NPCList) = 0 Or Not AutoAI Or GetStore Or SendSell Or SendStore Or SendBuy Then
        If UBound(NPCList) = 0 Then
            SendSell = False
            SendBuy = False
            SendStore = False
            GetStore = False
        End If
        CheckNPC = False
        Exit Function
    End If
    Dim i, j As Integer
    Dim dis As Integer
    Dim npccode As String
    
    For i = 0 To UBound(NPCList) - 1
        dis = EvalNorm(NPCList(i).Pos, curPos)
        Dim WeGo As Boolean
        WeGo = CanGO(curPos, NPCList(i).Pos)
        For j = 0 To UBound(ai_npc)
            'Check BUY STORE SELL action to NPC
            If (ai_npc(j).Pos.Y = NPCList(i).Pos.X And ai_npc(j).Pos.X = NPCList(i).Pos.Y) Then
                With ai_npc(j)
                    Select Case .Cause
                        Case "SELL"
                            If HaveSellItem And dis < 10 Then
                                CheckNPC = True
                                ReDim Route(0)
                                Stat "Found [" & NPCList(i).Name & "] in case of <SELL>..." & vbCrLf
                                frmMain.Send_Talk NPCList(i).ID
                                frmMain.tmrDealNPC.Enabled = False
                                frmMain.tmrDealNPC.Enabled = True
                                SendSell = False
                                SendStore = False
                                SendBuy = False
                                Exit Function
                            End If
                        Case "BUY"
                            If HaveSellItem And dis < 10 Then
                                CheckNPC = True
                                ReDim Route(0)
                                Stat "Found [" & NPCList(i).Name & "] in case of <SELL>..." & vbCrLf
                                frmMain.Send_Talk NPCList(i).ID
                                frmMain.tmrDealNPC.Enabled = False
                                frmMain.tmrDealNPC.Enabled = True
                                SendSell = False
                                SendStore = False
                                SendBuy = False
                                Exit Function

                            ElseIf (HaveBuyItem Or (CartBuy And mIsGoBuy)) And dis < 10 And CheckNPCBuy(CLng(j)) Then
                                CheckNPC = True
                                ReDim Route(0)
                                Stat "Found [" & NPCList(i).Name & "] in case of <BUY>..." & vbCrLf
                                frmMain.Send_Talk NPCList(i).ID
                                frmMain.tmrDealNPC.Enabled = False
                                frmMain.tmrDealNPC.Enabled = True
                                SendSell = False
                                SendStore = False
                                SendBuy = False
                                Exit Function

                            End If

                        Case "STORE"
                            If (HaveStoreItem Or HaveGetStorageItem) And dis < 10 And (GetTickCount - LastGetStorage) > 60000 Then
                                CheckNPC = True
                                ReDim Route(0)
                                npc_step = .Script
                                Stat "Found [" & NPCList(i).Name & "] in case of <STORE>..." & vbCrLf
                                frmMain.Send_Talk NPCList(i).ID
                                frmMain.tmrDealNPC.Enabled = False
                                frmMain.tmrDealNPC.Enabled = True
                                SendSell = False
                                SendStore = False
                                SendBuy = False
                                Exit Function
                            End If
                    End Select
                End With
            End If
step_go:
        Next
        
        'Check Warp with NPC
        For j = 0 To UBound(npcwarp)
            With npcwarp(j)
                If (.Pos.Y = NPCList(i).Pos.X And .Pos.X = NPCList(i).Pos.Y) And dis < 10 Then
                    Select Case .Cause
                        Case "MAPROUTE", "WARP"
                            If get_warpsolutions(MapName, NPCList(i).Pos) Then
                                Dim Ksl&, isChecks As Boolean
                                isChecks = False
                                For Ksl = 0 To UBound(MapRoute)
                                    If InStr(.Target, MapRoute(Ksl).Des.Name) And _
                                        MapRoute(Ksl).Src.Name = MapName Then
                                        isChecks = True
                                        Exit For
                                    End If
                                Next
                                If Not isChecks Then GoTo step_go2
                                ReDim Route(0)
                                CheckNPC = True
                                Stat "Found [" & NPCList(i).Name & "] in case of <MAPROUTE>..." & vbCrLf
                                npc_step = Replace(.Script, "SR", Get_NPCWarp_Choice(MapRoute(Ksl).Des.Name, j))
                                frmMain.Send_Talk NPCList(i).ID
                                frmMain.tmrDealNPC.Enabled = False
                                frmMain.tmrDealNPC.Enabled = True
                                Exit Function
                            End If
                    End Select
                End If
            End With
step_go2:
        Next
endloop:
    Next
    
    CheckNPC = False
Exit Function
errie:
    If Err.number > 0 Then print_funcerr "CheckNPC", Err.number, Err.Description
    Err.Clear
End Function


Public Function HaveBuyItem() As Boolean
On Error GoTo EndFunc
    Dim X As Integer
    Dim Index As Integer
    If BuyItem(0).Name = "" Then GoTo EndFunc
    For X = 0 To UBound(BuyItem)
            Index = Find_Item(BuyItem(X).Name)
            If Index = 0 Then
                HaveBuyItem = True
                Exit Function
            ElseIf AllInv(Index).Amount < BuyItem(X).Amount Then
                HaveBuyItem = True
                Exit Function
            End If
            If IsCartWant(BuyItem(X).Name) Then
                HaveBuyItem = True
                Exit Function
            End If
    Next
EndFunc:
    HaveBuyItem = False
    Err.Clear
End Function
    
Function CheckNPCBuy(Index As Long) As Boolean
On Error GoTo errie
    If UBound(ai_npc) < Index Then Exit Function
    If Index < 0 Then Exit Function
    Dim i&, j&, X&, ws() As String
    ws = Split(LCase(ai_npc(Index).Target), "&")
    For i = 0 To UBound(BuyItem)
        For j = 0 To UBound(ws)
            If LCase(ws(j)) = LCase(BuyItem(i).Name) Then
                X = Find_Item(BuyItem(i).Name)
                If X < 1 Then
                    CheckNPCBuy = True
                    Exit Function
                ElseIf AllInv(Index).Amount < BuyItem(i).Amount Then
                    CheckNPCBuy = True
                    Exit Function
                End If
                If IsCartWant(BuyItem(i).Name) Then
                    CheckNPCBuy = True
                    Exit Function
                End If
            End If
        Next
    Next
    CheckNPCBuy = False
Exit Function
errie:
    If Err.number > 0 Then print_funcerr "CheckNPCBuy", Err.number, Err.Description
    Err.Clear
End Function

Public Function HaveGetStorageItem() As Boolean
On Error GoTo EndFunc
    Dim X, i As Integer
    Dim Index As Integer
    If NoStoreItem Then GoTo EndFunc
    If GetStorageItem(0).Name = "" Then GoTo EndFunc
    For X = 0 To UBound(GetStorageItem)
        If Not GetStorageItem(X).NoStore Then
            Index = Find_Item(GetStorageItem(X).Name)
            If Index = 0 Then
                HaveGetStorageItem = True
                Exit Function
            ElseIf AllInv(Index).Amount < GetStorageItem(X).Amount Then
                HaveGetStorageItem = True
                Exit Function
            End If
        End If
    Next
EndFunc:
    If Err.number > 0 Then print_funcerr "HaveGetStorageItem", Err.number, Err.Description
    Err.Clear
    HaveGetStorageItem = False
End Function

Public Function NeedGetStorage() As Boolean
On Error GoTo EndFunc
    Dim i As Integer
    Dim Index As Integer
    If NoStoreItem Then GoTo EndFunc
    If GetStorageItem(0).Name = "" Then GoTo EndFunc
    For i = 0 To UBound(GetStorageItem)
        If Not GetStorageItem(i).NoStore Then
            Index = Find_Item(GetStorageItem(i).Name)
            If Index > 0 Then
                If AllInv(Index).Amount <= GetStorageItem(i).BackNumber Then
                    NeedGetStorage = True
                    Exit Function
                End If
            Else
                NeedGetStorage = True
                Exit Function
            End If
        End If
    Next
EndFunc:
    If Err.number > 0 Then print_funcerr "NeedGetStorage", Err.number, Err.Description
    Err.Clear
    NeedGetStorage = False
End Function

Public Function HaveSellItem() As Boolean
On Error GoTo errie
    Dim X As Integer
    For X = 0 To UBound(SelItem)
        If Check_Sell_Item(SelItem(X).Name) > 0 Then
            HaveSellItem = True
            Exit Function
        'ElseIf Find_Equip(SelItem(X).Name) > 0 Then
        '    HaveSellItem = True
        '    Exit Function
        End If
    Next
    HaveSellItem = False
Exit Function
errie:
    If Err.number > 0 Then print_funcerr "HaveSellItem", Err.number, Err.Description
    Err.Clear
End Function


Public Function HaveStoreItem() As Boolean
On Error GoTo EndFunc
    Dim X As Long, invID As Long
    If Kafra(0).Name = "" Then GoTo EndFunc
    For X = 0 To UBound(Kafra)
        If Not Kafra(X).CantKeep Then
            invID = Find_Item(Kafra(X).Name)
            If invID > 0 Then
                If AllInv(invID).Amount > Kafra(X).Amount Then
                    HaveStoreItem = True
                    Exit Function
                End If
            End If
        End If
    Next
    For X = 0 To UBound(CartCtrl)
        If CartCtrl(X).autostore And Find_CartID(CartCtrl(X).Name) >= 0 And IsCartWant(CartCtrl(X).Name) Then
            HaveStoreItem = True
            Exit Function
        End If
    Next
EndFunc:
    If Err.number > 0 Then print_funcerr "HaveStoreItem", Err.number, Err.Description
    Err.Clear
    HaveStoreItem = False
End Function


Public Function CartBuy() As Boolean
On Error GoTo EndFunc
    Dim i As Integer
    Dim Index As Integer
    For i = 0 To UBound(CartCtrl2)
        If IsCartWant(CartCtrl2(i).Name) Then
            CartBuy = True
            Exit Function
        End If
    Next
    CartBuy = False
    Exit Function
EndFunc:
    If Err.number > 0 Then print_funcerr "CartBuy", Err.number, Err.Description
    Err.Clear
    CartBuy = False
End Function

Public Function NeedBuy() As Boolean
    Dim i As Integer
    Dim Index As Integer
    If BuyItem(0).Name = "" Or Not isBackBuy Then GoTo EndFunc
    For i = 0 To UBound(BuyItem)
        Index = Find_Item(BuyItem(i).Name)
        If Index > 0 Then
            If AllInv(Index).Amount <= BuyItem(i).BackNumber Then
                NeedBuy = True
                Exit Function
            ElseIf AllInv(Index).Amount < BuyItem(i).Amount And ForceBuy And MapName = SaveMapName Then
                NeedBuy = True
                Exit Function
            End If
        Else
            NeedBuy = True
            Exit Function
        End If
    Next
EndFunc:
    NeedBuy = False
End Function

Function Find_Item(Name As String) As Integer
On Error GoTo errie
Dim X As Integer
For X = 0 To UBound(AllInv)
    If (AllInv(X).Amount > 0) And (LCase(AllInv(X).Name) = LCase(Name)) Then
        Find_Item = X
        Exit Function
    End If
Next
errie:
Find_Item = 0
Err.Clear
End Function

Function Check_Sell_Item(Name As String) As Integer
On Error GoTo errie
Dim X As Integer
For X = 0 To UBound(AllInv)
    If (AllInv(X).Amount > 0) And (AllInv(X).Name = Name) And (AllInv(X).Pos = 0) Then
        Check_Sell_Item = X
        Exit Function
    End If
Next
errie:
Check_Sell_Item = 0
End Function


'Function Find_Equip(Name As String) As Integer
'On Error GoTo errie
'Dim X As Integer
'For X = 0 To UBound(AllInv)
'    If (AllInv(X).Amount > 0) And (LCase(AllInv(X).Name) = LCase(Name)) Then
'        Find_Equip = X
'        Exit Function
'    End If
'Next
'errie:
'Find_Equip = 0
'Err.Clear
''ClearAll
'End Function

Sub Check_Route()
On Error GoTo errie
    If UBound(Route) > 0 And ((Not CanFindMonster And CurAtkMonster.NameID = 0) Or MoveOnly Or Not AutoAI) Then
        If EvalNorm(Route(Current), curPos) < 5 Then 'And (GetTickCount >= DelaymoveCounter)
            If Current < UBound(Route) Then
                Current = Current + 1
                move_to Route(Current)
                If Current < UBound(Route) Then move_to Route(Current + 1)
                GoOnRoute = True
                DelaymoveCounter = GetTickCount + (MovementSpeed * (EvalNorm(curPos, Route(Current)) - 3))
            Else
                ReDim Route(0)
                DelaymoveCounter = 0
                Stat "You're reaching destination..." & vbCrLf
                ptEnd.X = 0
                ptEnd.Y = 0
            End If
            CalcModAI "Check_Route"
            If ModAI Then ReDim Route(0)
            If Current < UBound(Route) Then Current = Current + 1 Else Current = UBound(Route)
            RouteCounter = 0
            Exit Sub
        ElseIf RouteCounter > 200 Then
            ReDim Route(0)
            RouteCounter = 0
        ElseIf Not BlockMove Then
            If RouteCounter Mod 50 = 0 Then move_to Route(Current)
            RouteCounter = RouteCounter + 1
            GoOnRoute = True
        End If
    End If
Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "Check_Route", Err.number, Err.Description
    Err.Clear
End Sub

'Procedure to check Destination and using Map ROuting AI to reach
Sub Check_Destination_Route()
On Error GoTo errie
    Dim Index, i As Integer
    Dim indextest As Integer
    'calculate modificate ai
    CalcModAI "check_destination_route"
    If ModAI Then Exit Sub
    
    'Go to Storage NPC if we're in the same map
    Index = Check_sameNPCMap(MapName, "STORE")
    If (HaveStoreItem Or HaveGetStorageItem Or MIsGoStore) And Index > -1 And (GetTickCount - LastGetStorage) > 60000 Then
        ReDim CurRoute(0)
        If Index <= UBound(ai_npc) Then
            'ReDim MapRoute(0)
            Stat "Go to your [kafra] to store..." & vbCrLf
            CurRoute(0).Name = MapName
            CurRoute(0).Pos.Y = ai_npc(Index).Pos.Y
            CurRoute(0).Pos.X = ai_npc(Index).Pos.X
        End If
        GoTo end_check
    End If
    
    'If kafra NPC is not in the same map so we need to Routing to it
    Index = Check_NPCMap(MapName, "STORE")
    If isBackStore And NeedGetStorage And Index > -1 And (GetTickCount - LastGetStorage) > 60000 Then
        If (Find_Item("Butterfly_Wing") > 0 Or SkillExists("AL_TELEPORT")) And ai_npc(Index).location = SaveMapName Then
            frmMain.Warp_Save "Auto-warp to save for store."
            Exit Sub
        End If
        ReDim CurRoute(0)
        'Do we find the solutions yet ?
        indextest = get_buy_solutions(ai_npc(Index).location)
        If indextest < 0 Then
            Stat "Search route from [" & MapName & "] to [" & ai_npc(Index).location & "] for storage, waiting..." & vbCrLf
            'Do Search Map Routing
            Do_Search_Map_Routing MapName, ai_npc(Index).location

            'If we 're on discontinuous MAP then use extend map routing
            'Mode 1 Means Check Portals that can go to NPC room
            Replace_Map_Routing 1, ai_npc(Index).location, ai_npc(Index).Pos, , , True
        End If
        GoTo end_check
    End If

    'Go to Tool Dealer NPC to sell if we're in the same map
    Index = Check_sameNPCMap(MapName, "SELL")
    If Index < 0 Then Index = Check_sameNPCMap(MapName, "BUY")
    If HaveSellItem And Index > -1 Then
        ReDim CurRoute(0)
        If Index <= UBound(ai_npc) Then
            Stat "Go to your [Tool Dealer] to Buy/Sell..." & vbCrLf
            CurRoute(0).Name = MapName
            CurRoute(0).Pos.Y = ai_npc(Index).Pos.Y
            CurRoute(0).Pos.X = ai_npc(Index).Pos.X
        End If
        GoTo end_check
    End If
    
    Index = Check_sameNPCMap(MapName, "BUY")
    Do Until CheckNPCBuy(CLng(Index)) = True Or Index < 0
        Index = Check_sameNPCMap(MapName, "BUY", Index + 1)
    Loop
    If (WantBuy Or (CartBuy And mIsGoBuy)) And Index > -1 Then
        ReDim CurRoute(0)
        If Index <= UBound(ai_npc) Then
            Stat "Go to your [Tool Dealer] to Buy/Sell..." & vbCrLf
            CurRoute(0).Name = MapName
            CurRoute(0).Pos.Y = ai_npc(Index).Pos.Y
            CurRoute(0).Pos.X = ai_npc(Index).Pos.X
        End If
        GoTo end_check
    End If
    
    Index = Check_NPCMap(MapName, "SELL")
    If Index < 0 Then Index = Check_NPCMap(MapName, "BUY")
    If GetWeight >= WeightBackTown And IsBackTown And (HaveSellItem Or HaveStoreItem) And StartBot And Index >= 0 Then
        If (Find_Item("Butterfly_Wing") > 0 Or SkillExists("AL_TELEPORT")) And ai_npc(Index).location = SaveMapName Then
            frmMain.Warp_Save "Auto-warp to save for buy/sell."
            Exit Sub
        End If
        ReDim CurRoute(0)
        'Do we find the solutions yet ?
        indextest = get_buy_solutions(ai_npc(Index).location)
        If indextest < 0 Then
            Stat "Search route from [" & MapName & "] to [" & ai_npc(Index).location & "] for sell, waiting..." & vbCrLf
            'Do Search Map Routing
            Do_Search_Map_Routing MapName, ai_npc(Index).location

            'If we 're on discontinuous MAP then use extend map routing
            'Mode 1 Means Check Portals that can go to NPC room
            Replace_Map_Routing 1, ai_npc(Index).location, ai_npc(Index).Pos, , , True

        End If
        GoTo end_check
    End If
    
    ReDim CurRoute(0)
    'If Tool Dealer NPC is not in the same map so we need to Routing to it
    Dim tmpPortal As MyPortal
    Dim Pt As Coord
    Index = Check_NPCMap(MapName, "BUY")
    Do Until CheckNPCBuy(CLng(Index)) = True Or Index < 0
        Index = Check_sameNPCMap(MapName, "BUY", Index + 1)
    Loop
    If (IsForceBuy) And StartBot And Index > -1 Then
        ReDim CurRoute(0)
        If (Find_Item("Butterfly_Wing") > 0 Or SkillExists("AL_TELEPORT")) And ai_npc(Index).location = SaveMapName Then
            frmMain.Warp_Save "Auto-warp to save for buy/sell."
            Exit Sub
        End If
        'Do we find the solutions yet ?
        indextest = get_buy_solutions(ai_npc(Index).location)
        If indextest < 0 Then
            Stat "Search route from [" & MapName & "] to [" & ai_npc(Index).location & "] for buy, waiting..." & vbCrLf
            'Do Search Map Routing
            Do_Search_Map_Routing MapName, ai_npc(Index).location

            'If we 're on discontinuous MAP then use extend map routing
            'Mode 1 Means Check Portals that can go to NPC room
            Replace_Map_Routing 1, ai_npc(Index).location, ai_npc(Index).Pos, , , True

        End If
        GoTo end_check
    End If
    
    'if you aren't in lock XY position
    If LockMapName = MapName And Not IsInLock And LockXY.X > 0 And LockXY.Y > 0 Then
        ReDim CurRoute(0)
        Dim lockRand As Coord
        lockRand.Y = (LockXY.X - LockXYRand.X) + RandomNumber((LockXYRand.X * 2) + 1, 0)
        lockRand.X = (LockXY.Y - LockXYRand.Y) + RandomNumber((LockXYRand.Y * 2) + 1, 0)
        ReDim CurRoute(0)
        CurRoute(0).Name = MapName
        CurRoute(0).Pos.Y = lockRand.X
        CurRoute(0).Pos.X = lockRand.Y
        Stat "Going to lockmap position at " & CStr(lockRand.Y) & ":" & CStr(lockRand.X) & " [" & EvalNorm(curPos, lockRand) & " blks far]" + vbCrLf, &HFF00FF
        If EvalNorm(curPos, lockRand) < 10 Then
            move_to lockRand
            ReDim CurRoute(0)
        End If
        GoTo end_check
    End If
    
    'If we need to go to the lockmap
    'Do we find the solutions ?
    If get_solutions(MapName) < 0 And LockMapName <> MapName Then 'no ?
        ReDim CurRoute(0)
        Stat "Search route from [" & MapName & "] to [" & LockMapName & "] for lockmap, waiting..." & vbCrLf
        'Do Search Map Routing
        Do_Search_Map_Routing MapName, LockMapName
        'Mode 2 Means Check Portals that can go out from NPC room
        'Replace_Map_Routing 2, MapName, CurPos
        If UBound(MapRoute) > 0 Then
            For i = 1 To UBound(MapRoute)
                If Check_IsManyPortal(MapRoute(i).Src.Name, MapRoute(i - 1).Src.Name) Then
                    If i > 1 Then
                        Replace_Map_Routing 1, MapRoute(i).Src.Name, MapRoute(i).Src.Pos, i - 1, MapRoute(i - 1).Src.Name, True
                    End If
                End If
            Next
        End If
        If Check_IsManyPortal(MapRoute(0).Des.Name, MapRoute(0).Src.Name) Then
            Replace_Map_Routing 2, MapName, curPos, , MapRoute(0).Des.Name, False
        End If
        GoTo end_check
    End If
end_check:
    Check_RouteMap
    'Are we near the Portal Route ?
    Index = Check_Portal_on_Route(MapName)
    If Index > -1 Then
        'move_to ExitPortal(index).pos 'yes, go to it
        Go_NearerPoint curPos, ExitPortal(Index).Pos
        'Stat "Go to portals" & vbCrLf
    End If
    MoveOnly = True
Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "Check_Destination_Route", Err.number, Err.Description
    Err.Clear
End Sub



Public Function CanFindMonster() As Boolean
On Error GoTo errie
    Dim i As Long
    If UBound(Aggro) > 0 Then
        CanFindMonster = True
        Exit Function
    End If
    If UBound(MonsterList) = 0 Then GoTo errie
        For i = 0 To UBound(MonsterList) - 1
            If MonsterList(i).NoAttack Or (IsSMAgg(MonsterList(i).Name) = False And IsSMR(MonsterList(i).Name) = False) Then
                GoTo end_loop
            ElseIf (MonsterList(i).IsAttack Or MonsterList(i).NoAttack) And Not killsteal Then
                GoTo end_loop
            ElseIf MonsterList(i).CantGo Then
                GoTo end_loop
            ElseIf (MonsterList(i).IsFollow Or MonsterList(i).IsTrap) And Not isKillmob Then
                GoTo end_loop
            Else
                CanFindMonster = True
                Exit Function
            End If
end_loop:
        Next
errie:
    If Err.number > 0 Then print_funcerr "CanFindMonster", Err.number, Err.Description
    Err.Clear
    CanFindMonster = False
End Function

Public Function WantBuy() As Boolean
    Dim i As Integer
    Dim Index As Integer
    If BuyItem(0).Name = "" Then GoTo EndFunc
    For i = 0 To UBound(BuyItem)
        Index = Find_Item(BuyItem(i).Name)
        If Index > 0 Then
            If AllInv(Index).Amount < BuyItem(i).Amount Then
                WantBuy = True
                Exit Function
            End If
        Else
            WantBuy = True
            Exit Function
        End If
    Next
EndFunc:
    WantBuy = False
End Function

Private Function Check_Portal_on_Route(MapName As String) As Integer
On Error GoTo errie
    Dim i, j As Integer
    Dim tmpPos As Coord
    If UBound(ExitPortal) > 0 Then
        For i = 0 To UBound(ExitPortal) - 1
            For j = 0 To UBound(MapRoute)
                If MapRoute(j).Src.Name = MapName Then
                    tmpPos.X = ExitPortal(i).Pos.Y
                    tmpPos.Y = ExitPortal(i).Pos.X
                    If EvalNorm(MapRoute(j).Src.Pos, tmpPos) < 5 Then
                        Check_Portal_on_Route = i
                        Exit Function
                    End If
                End If
            Next
        Next
    End If
    Check_Portal_on_Route = -1
Exit Function
errie:
    If Err.number > 0 Then print_funcerr "Check_Portal_on_Route", Err.number, Err.Description
    Err.Clear
End Function

Public Function ETA(STime As Double) As String
    Dim Hrs&, Mns&, Sec&, Dys&, SecTime&
    SecTime = Abs(STime)
    Sec = (SecTime Mod 60)
    Mns = Int((SecTime - Sec) / 60)
    Hrs = Int((Mns - (Mns Mod 60)) / 60)
    Dys = Int(Hrs / 24)
    Hrs = Hrs Mod 24

    ETA = IIf(Dys > 0, Dys & " day" & IIf(Dys > 1, "s ", " "), "") & IIf(Hrs < 10, "0" & CStr(Hrs), CStr(Hrs)) & ":" & IIf(Mns Mod 60 < 10, "0" & CStr(Mns Mod 60), CStr(Mns Mod 60)) & ":" & IIf(Sec < 10, "0" & CStr(Sec), CStr(Sec))
End Function


Public Function MakeTime() As String
    Dim SDay As Long
    SDay = SHour \ 24
    SHour = SHour Mod 24
    MakeTime = CStr(SDay) & ":"
    If (SHour < 10) Then
        MakeTime = MakeTime + "0" + CStr(SHour) + ":"
    Else
        MakeTime = MakeTime + CStr(SHour) + ":"
    End If
    If (SMin < 10) Then
        MakeTime = MakeTime + "0" + CStr(SMin) + ":"
    Else
        MakeTime = MakeTime + CStr(SMin) + ":"
    End If
    If (SSec < 10) Then
        MakeTime = MakeTime + "0" + CStr(SSec)
    Else
        MakeTime = MakeTime + CStr(SSec)
    End If
    
End Function

Function GetWeight() As Single
    If Players(number).MaxWeight > 0 Then GetWeight = CSng(CDbl(Players(number).Weight) / CDbl(Players(number).MaxWeight)) Else GetWeight = 0
End Function
Function GetSP() As Single
    If Players(number).maxsp > 0 Then GetSP = CSng(CDbl(Players(number).SP) / CDbl(Players(number).maxsp)) Else GetSP = 0
End Function
Function GetHP() As Single
    If Players(number).MaxHP > 0 Then GetHP = CSng(CDbl(Players(number).HP) / CDbl(Players(number).MaxHP)) Else GetHP = 0
End Function

Public Function IsForceBuy() As Boolean
On Error GoTo errie
    Dim i As Integer
    Dim Index As Integer
    If BuyItem(0).Name = "" Or Not isBackBuy Then GoTo EndFunc
    For i = 0 To UBound(BuyItem)
        Index = Find_Item(BuyItem(i).Name)
        If Index > 0 Then
            If AllInv(Index).Amount < BuyItem(i).Amount And ForceBuy And MapName = SaveMapName Then
                IsForceBuy = True
                Exit Function
            End If
        Else
            IsForceBuy = True
            Exit Function
        End If
    Next
EndFunc:
    IsForceBuy = False
    Exit Function
errie:
    If Err.number > 0 Then print_funcerr "ForceBuy", Err.number, Err.Description
    Err.Clear
End Function

Sub EstimateClosestItem()
    On Error GoTo errie
    If GetWeight > Weight2 Then Exit Sub
    Dim i As Integer
    Dim BestItem As Integer
    Dim Sel_distance As Integer
    Dim Cur_Distance As Integer
    Dim found As Boolean
    If StopAction Then Exit Sub
    BestItem = 500
    Sel_distance = 30000
    Cur_Distance = 0
    found = False

    If Items(0).Pos.Y = 0 And Items(0).Pos.X = 0 Then
        Pickup = False
        frmMain.tmrPickup.Enabled = False
        Tracing = False
        Exit Sub
    End If
    For i = 0 To UBound(Items) - 1
        Cur_Distance = EvalNorm(Items(i).Pos, curPos)

        If Cur_Distance < Sel_distance Then
            Sel_distance = Cur_Distance
            BestItem = i
            found = True
        End If
    Next
    If CurrentItem.ID = Items(BestItem).ID Then Exit Sub
    CurrentItem.ID = Items(BestItem).ID
    CurrentItem.Name = Items(BestItem).Name
    CurrentItem.Pos = Items(BestItem).Pos
    Stat "Select [" & Return_ItemName(CurrentItem.Name) & "], as a target" & vbCrLf
    GotItem = False
    Tracing = True
    Pickuptime = 0
    TryPicktime = 0
    SendPickup
    Exit Sub
errie:
    Err.Clear
End Sub

Sub SendPickup()
On Error GoTo errie
    Dim found As Boolean
    Dim X As Integer
    Dim Y As Integer
    Dim delete As Boolean
    delete = False
    If Not Pickup Then Exit Sub
    If Pickuptime > 0 Then
        delete = True
    End If
    If TryPicktime > 15 Then
        delete = True
        TryPicktime = 0
    End If
    frmMain.Caption = "pickuptime =" & CStr(Pickuptime)
    found = False
    For X = 0 To UBound(Items) - 1
        If CurrentItem.ID = Items(X).ID Then
            If delete Then
                For Y = X To UBound(Items) - 1
                    Items(Y) = Items(Y + 1)
                Next
                ReDim Preserve Items(UBound(Items) - 1)
                CurrentItem.ID = ""
                CurrentItem.Name = ""
                found = False
                Exit For
            End If
            found = True
            Pickuptime = 0
        End If
    Next
    
    If Not found Then
        If UBound(Items) > 0 Then
            EstimateClosestItem
        Else
            Pickup = False
            Exit Sub
        End If
    End If

    frmMain.labCurMons.Caption = "[" + Return_ItemName(CurrentItem.Name) + "], " _
            & CStr(EvalNorm(CurrentItem.Pos, curPos)) + " Blocks (Picking)"
    
    If EvalNorm(CurrentItem.Pos, curPos) < 2 And IsAutoPick Then
        Winsock_SendPacket Chr(&H64 + &H3B) + Chr(&H0) + CurrentItem.ID, True
        Pickuptime = Pickuptime + 1
        ActionDelay = RandomNumber(3, 2)
    Else
        If TryPicktime > 9 Then
            Winsock_SendPacket Chr(&H64 + &H3B) + Chr(&H0) + CurrentItem.ID, True
        Else
            Winsock_SendPacket IntToChr(&H85) + MakeItemPos(CurrentItem.Pos), True
        End If
        
        TryPicktime = TryPicktime + 1
        'Stat "Go to [" & Return_ItemName(CurrentItem.Name) & "] position..." & vbCrLf
    End If
Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "SendPickup", Err.number, Err.Description
    Err.Clear
End Sub



Function Is_Pickup(itemName As String, Limit As Integer) As Byte
On Error GoTo errie
Dim X As Integer
Dim Ans As Byte
Ans = 0
For X = 0 To UBound(Itempick)
    If itemName = Itempick(X).Name Then
        If Itempick(X).Amount = 0 Then
            Ans = 1
        ElseIf Limit >= Itempick(X).Amount Then
            Ans = 2
        End If
        Is_Pickup = Ans
        Exit For
    End If
Next
    Is_Pickup = Ans
Exit Function
errie:
    If Err.number > 0 Then print_funcerr "Is_Pickup", Err.number, Err.Description
    Err.Clear
End Function

Function isRare(itemName As String) As Boolean
On Error GoTo errie
Dim X As Integer
Dim Ans As Boolean
Ans = False
For X = 0 To UBound(RareItem)
    If itemName = RareItem(X).Name Then
        Ans = True
        Exit For
    End If
Next
    isRare = Ans
Exit Function
errie:
    If Err.number > 0 Then print_funcerr "isRare", Err.number, Err.Description
    Err.Clear
End Function

Function Check_Pickup() As Boolean
On Error GoTo errie
    If Players(number).Weight = 0 Or Players(number).MaxWeight = 0 Then
        Check_Pickup = True
        Exit Function
    End If
    If Players(number).Weight >= (Weight2 * Players(number).MaxWeight) And SWeight2 Then
        Check_Pickup = False
        If Not Isweight2 Then Chat "System : [Overweight],You set to stop pick up...", MColor.Fail
        Isweight2 = True
    Else
        Check_Pickup = True
        Isweight2 = False
    End If
Exit Function
errie:
    If Err.number > 0 Then print_funcerr "Check_Pickup", Err.number, Err.Description
    Err.Clear
End Function

Function Check_Attack() As Boolean
On Error GoTo errie
    If Players(number).Weight = 0 Or Players(number).MaxWeight = 0 Then
        Check_Attack = True
        Exit Function
    End If
    If GetWeight >= (Weight1) And SWeight1 Then
        Check_Attack = False
        If Not Isweight1 Then Chat "System : [Overweight], You set to stop attack...", MColor.Fail
        Isweight1 = True
        'FightMode = False
        'SellMode = True
    Else
        Check_Attack = True
        Isweight1 = False
        'FightMode = True
        'SellMode = False
    End If
Exit Function
errie:
    If Err.number > 0 Then print_funcerr "Check_Attack", Err.number, Err.Description
    Err.Clear
End Function

Function Get_AmountItem(Name As String) As Integer
On Error GoTo errie
Dim X As Integer
For X = 0 To UBound(AllInv)
    If (AllInv(X).Amount > 0) And (LCase(AllInv(X).Name) = LCase(Name)) Then
        Get_AmountItem = AllInv(X).Amount
        Exit Function
    End If
Next
errie:
Get_AmountItem = 0
Err.Clear
End Function

Public Sub UpdateGuild()
'print_errror "sub UpdateGuild"
On Error GoTo errie
    frmGuild.lstGuild.Clear
    If Len(Players(number).Guild) > 0 Then frmGuild.LabGuild.Caption = Players(number).Guild & " - Guild"
    Dim X As Integer
    For X = 0 To UBound(Guild)
        With Guild(X)
            If .Name <> "" Then
                Dim tstr As String
                tstr = .Name & " [" & .PosName & "]," & .Sex & ", " & .Class & " ,Lv:" & .Lv
                If .isOnline Then tstr = tstr & ", Online"
                frmGuild.lstGuild.AddItem tstr
            End If
        End With
    Next
Exit Sub
errie:
'Chat "Error in UpdateGuild : " & Err.Description
Err.Clear
End Sub

