Attribute VB_Name = "md_AI"
Option Explicit
Public StrEncPassword As String
Public LockXY As Coord
Public LockXYRand As Coord
Public MIsGoStore As Boolean
Public mIsGoBuy As Boolean

Public Function EnUser(STR As String, iType As Integer) As String
Dim txtout$, ap$

If iType = 1 Then ap = LngToChr(471242415) & LngToChr(8501) & "•R[e]V=E[m]U•" & Chr(11) & vbCrLf & LngToChr(vbBlue) _
Else: ap = "TDes" & "•R[e]V=E[m]U•" & LngToChr(8501) & vbCrLf & Chr(1) & LngToChr(vbGreen)

txtout = TEncode(STR, ap)
txtout = TEncode(txtout, ap)
EnUser = "Revemu[" & txtout & "]"
End Function

Public Function DeUser(STR As String, iType As Integer) As String
Dim Out$, ap$
Out = Replace(STR, "Revemu[", "")
Out = Left(Out, Len(Out) - 1)

If iType = 1 Then ap = LngToChr(471242415) & LngToChr(8501) & "•R[e]V=E[m]U•" & Chr(11) & vbCrLf & LngToChr(vbBlue) _
Else: ap = "TDes" & "•R[e]V=E[m]U•" & LngToChr(8501) & vbCrLf & Chr(1) & LngToChr(vbGreen)

Out = TDecode(Out, ap)
Out = TDecode(Out, ap)

DeUser = Out
End Function

Function TEncode(inData As String, StrPass As String) As String
Dim strResult As String
strResult = String$(1024, 0)

DESEncrypt inData, StrPass, strResult
If InStr(strResult, Chr$(0)) > 0 Then strResult = Left$(strResult, InStr(strResult, Chr$(0)) - 1)
TEncode = strResult
End Function
Function TDecode(inData As String, StrPass As String) As String
Dim strResult As String
strResult = String$(1024, 0)

DESDecrypt inData, StrPass, strResult
'MsgBox ChrtoHex(strResult)
If InStr(strResult, Chr$(0)) > 0 Then strResult = Left$(strResult, InStr(strResult, Chr$(0)) - 1)
TDecode = strResult
End Function

Sub CheckCartAI(ItemID As Long)
On Error GoTo errie
'inventory > cart
    If Not IsCartOn Then Exit Sub
    Dim i As Long, Amount As Long, tmpAmount&
    For i = 1 To UBound(CartCtrl)
        If LCase(CartCtrl(i).Name) = LCase(AllInv(ItemID).Name) Then
            tmpAmount = AllInv(ItemID).Amount
            If IsCartNeed(AllInv(ItemID).Name) And AllInv(ItemID).Amount > 0 Then
                ' minimum item in cart
                If AllInv(ItemID).Amount > CartNeedAmount(AllInv(ItemID).Name) Then
                    Amount = CartNeedAmount(AllInv(ItemID).Name)
                Else
                    Amount = AllInv(ItemID).Amount
                End If
                ' maximum item in inventory
                If Amount = 0 Then GoTo donext
                pkt_CartGet ItemID, Amount
                tmpAmount = tmpAmount - Amount
                Stat "Cart n auto-add : " & AllInv(ItemID).Name & " " & Amount & "EA" + vbCrLf, vbBlue
            End If
            If CartCtrl(i).autoadd Then
                If tmpAmount > CartCtrl(i).Min Then
                    pkt_CartGet ItemID, tmpAmount - CartCtrl(i).Min
                    Stat "Cart auto-add : " & AllInv(ItemID).Name & "[" & tmpAmount - CartCtrl(i).Min & "EA]" & vbCrLf, vbBlue
                End If
            End If
            Exit Sub
        End If
donext:
    Next
    Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "CheckCartAI", Err.number, Err.Description
    Err.Clear
End Sub
Sub CheckCartInv(cartid As Long)
'cart > inventory
    If Not IsCartOn Then Exit Sub
    Dim j As Long
    For j = 1 To UBound(CartCtrl)
        If LCase(Cart(cartid).Name) = LCase(CartCtrl(j).Name) Then
            If CartCtrl(j).autoget = True And Not IsCartNeed(CartCtrl(j).Name) Then
                Dim autoID As Long, Amount As Long
                autoID = Find_InvID(CartCtrl(j).Name)
                If autoID = -1 Then
                    Amount = IIf(CartCtrl(j).Min > Cart(cartid).Amount, Cart(cartid).Amount, CartCtrl(j).Min)
                    If Amount < CartNeedAmount(CartCtrl(j).Name) Then Exit Sub
                    pkt_CartTake cartid, Amount
                    Stat "Cart auto-get : " & Cart(cartid).Name + vbCrLf, vbBlue
                Else
                    Amount = IIf((CartCtrl(j).Min - AllInv(autoID).Amount > Cart(cartid).Amount), Cart(cartid).Amount, CartCtrl(j).Min - AllInv(autoID).Amount)
                    If Amount > 0 Then
                        If Amount < CartNeedAmount(CartCtrl(j).Name) Then Exit Sub
                        pkt_CartTake cartid, Amount
                        Stat "Cart auto-get : " & Cart(cartid).Name + vbCrLf, vbBlue
                    End If
                End If
            End If
            Exit For
        End If
    Next
End Sub
Sub CheckCartStore(cartid As Long)
    If Not IsCartOn Then Exit Sub
    Dim j As Long
    For j = 1 To UBound(CartCtrl)
        If LCase(Cart(cartid).Name) = LCase(CartCtrl(j).Name) Then
            If CartCtrl(j).autostore = True And Not IsCartWant(Cart(cartid).Name) And Not IsCartNeed(Cart(cartid).Name) Then
                If Cart(cartid).Amount > 0 Then
                    Dim Amount&
                    If CartNeedAmount(Cart(cartid).Name) > Cart(cartid).Amount Then Exit Sub
                    pkt_CartToKafra cartid, Cart(cartid).Amount
                    Stat "Cart auto-storage : " & Cart(cartid).Name + vbCrLf, vbBlue
                End If
            End If
            Exit For
        End If
    Next
End Sub
Sub CheckCartStorage(StoreID As Long)
    If Not IsCartOn Then Exit Sub
Dim takea&
    If IsCartNeed(Storage(StoreID).Name) Then
        If IsCartWant(Storage(StoreID).Name) Then
            If CartWantAmount(Storage(StoreID).Name) > Storage(StoreID).Amount Then takea = Storage(StoreID).Amount Else takea = CartWantAmount(Storage(StoreID).Name)
            pkt_CartFromKafra StoreID, takea
            Exit Sub
        End If
    End If
End Sub
Function IsCartNeed(CName As String) As Boolean
    Dim CtrlS As CartItemCtrl2
    CtrlS = Find_CartAI(CName)
    If CtrlS.Name = "" Then
        IsCartNeed = False
        Exit Function
    End If
    Dim cartid As Long
    cartid = Find_CartID(CName)
    If cartid < 0 And CtrlS.cartmin > 0 Then
        IsCartNeed = True
        Exit Function
    End If
    If Cart(cartid).Amount <= CtrlS.cartmin Then
        IsCartNeed = True
        Exit Function
    End If
    IsCartNeed = False
End Function
Function IsCartWant(CName As String) As Boolean
    Dim CtrlS As CartItemCtrl2
    CtrlS = Find_CartAI(CName)
    If CtrlS.Name = "" Then
        IsCartWant = False
        Exit Function
    End If
    Dim cartid As Long
    cartid = Find_CartID(CName)
    If cartid < 0 And CtrlS.cartmax > 0 Then
        IsCartWant = True
        Exit Function
    End If
    If Cart(cartid).Amount <= CtrlS.cartmax Then
        IsCartWant = True
        Exit Function
    End If
    IsCartWant = False
End Function
Function CartWantAmount(CName As String)
    If Not IsCartWant(CName) Then
        CartWantAmount = 0
        Exit Function
    End If
    Dim CtrlS As CartItemCtrl2
    CtrlS = Find_CartAI(CName)
    Dim cartid As Long
    cartid = Find_CartID(CName)
    If cartid < 0 Then
        CartWantAmount = CtrlS.cartmax
        Exit Function
    End If
    If CtrlS.cartmax > Cart(cartid).Amount Then
        CartWantAmount = CtrlS.cartmax - Cart(cartid).Amount
        Exit Function
    End If
    CartWantAmount = 0
End Function
Function CartNeedAmount(CName As String)
    If Not IsCartNeed(CName) Then
        CartNeedAmount = 0
        Exit Function
    End If
    Dim CtrlS As CartItemCtrl2
    CtrlS = Find_CartAI(CName)
    Dim cartid As Long
    cartid = Find_CartID(CName)
    If cartid < 0 Then
        CartNeedAmount = CtrlS.cartmin
        Exit Function
    End If
    If CtrlS.cartmin > Cart(cartid).Amount Then
        CartNeedAmount = CtrlS.cartmin - Cart(cartid).Amount
        Exit Function
    End If
    CartNeedAmount = 0
End Function
Function Find_CartAI(CName As String) As CartItemCtrl2
    Dim i&
    For i = 1 To UBound(CartCtrl2)
        If LCase(CName) = LCase(CartCtrl2(i).Name) Then
            Find_CartAI = CartCtrl2(i)
            Exit Function
        End If
    Next
    Find_CartAI.Name = ""
End Function

Sub ParseCommand(Nick As String, PrivMsg As String)
    If Left(PrivMsg, Len(MRemote.CommandBegin)) = MRemote.CommandBegin And MRemote.Enabled Then
        Dim mCmd As String
        mCmd = Mid(PrivMsg, Len(MRemote.CommandBegin) + 1, Len(PrivMsg) - Len(MRemote.CommandBegin) - 1)
        Dim mSpl() As String
        If InStr(mCmd, " ") Then
            mSpl = Split(mCmd, " ")
        Else
            ReDim mSpl(0)
            mSpl(0) = mCmd
        End If
        If LCase(mSpl(0)) = "identify" Then
            If UBound(mSpl) = 0 Then Exit Sub
            If mSpl(1) = MRemote.password Then
                MRemote.Owner = Nick
                MRemote.Identified = True
                SendPRIChat Nick, "Password accepted."
            End If
        End If
        If Not MRemote.Identified Or MRemote.Owner <> Nick Then Exit Sub
        Select Case LCase(mSpl(0))
            Case "statinfo"
                With Players(number)
                    SendPRIChat Nick, "Character stat for " & .Name
                    SendPRIChat Nick, "Str: " & .STR & "+" & .Strp & " Atk: " & .ATK & "+" & .ATKp & " Def: " & .Def & "+" & .Defp
                    SendPRIChat Nick, "Agi: " & .AGI & "+" & .Agip & " Matk: " & .MinMatk & "~" & .MaxMatk & " Mdef:" & .mDef & "+" & .mDefp
                    SendPRIChat Nick, "Vit: " & .VIT & "+" & .Vitp & " Hit: " & .Hit & " Flee: " & .Flee & "+" & .Fleep
                    SendPRIChat Nick, "Int: " & .Intl & "+" & .Intp & " Critical: " & .Crit & " Aspd: " & .Aspd
                    SendPRIChat Nick, "Dex: " & .DEX & "+" & .Dexp & " Status point: " & .StatPoint
                    SendPRIChat Nick, "Luk: " & .LUK & "+" & .Lukp
                End With
            Case "charinfo"
                With Players(number)
                    SendPRIChat Nick, "Character info for " & .Name & " [" & .Class & "]"
                    SendPRIChat Nick, "HP: " & .HP & "/" & .MaxHP & " SP: " & .SP & "/" & .maxsp
                    SendPRIChat Nick, "BaseLV: " & .BaseLV & " [" & .BaseExp & "/" & .NextBaseEXP & " - " & (FormatNumber((.BaseExp * 100) / .NextBaseEXP, 2, vbTrue)) & "%]"
                    SendPRIChat Nick, "JobLV: " & .JobLV & " [" & .JobExp & "/" & .MaxJobEXP & " - " & (FormatNumber((.JobExp * 100) / .MaxJobEXP, 2, vbTrue)) & "%]"
                    SendPRIChat Nick, "Weight: " & .Weight & "/" & .MaxWeight & " Zeny: " & .Zeny
                End With
            Case "setpasswd"
                If UBound(mSpl) < 2 Then Exit Sub
                If mSpl(1) = MRemote.password Then
                    MRemote.password = mSpl(2)
                    SendPRIChat Nick, "Password changed to : " & mSpl(2)
                End If
            Case "where"
                SendPRIChat Nick, "Current map : " & MapName & " at (" & curPos.Y & "," & curPos.X & ")"
        End Select
    End If
End Sub

Sub SendPRIChat(ChatName As String, ChatText As String)
    Dim outPacket As String
    Chat "to [" & ChatName & "] : " & ChatText, MColor.mychat
    outPacket = IntToChr(&H96) & IntToChr(Len(ChatText) + 29) & ChatName & String$(24 - Len(ChatName), 0) & ChatText & Chr(0)
    Winsock_SendPacket outPacket, True
End Sub
Sub SendPUBChat(ChatText As String)
    Dim outPacket As String
    outPacket = Players(number).Name & " : " & ChatText & Chr(0)
    outPacket = Chr(140) & Chr(0) & Mid(LngToChr(Len(outPacket) + 4), 1, 2) & outPacket
    Winsock_SendPacket outPacket, True
End Sub
Sub SendPARChat(ChatText As String)
    Dim outPacket As String
    outPacket = Players(number).Name & " : " & ChatText & Chr(0)
    outPacket = Chr(8) & Chr(1) & Mid(LngToChr(Len(outPacket) + 4), 1, 2) & outPacket
    Winsock_SendPacket outPacket, True
End Sub
Sub SendGUIChat(ChatText As String)
    Dim outPacket As String
    outPacket = Players(number).Name & " : " & ChatText & Chr(0)
    outPacket = Chr(126) & Chr(1) & Mid(LngToChr(Len(outPacket) + 4), 1, 2) & outPacket
    Winsock_SendPacket outPacket, True
End Sub

Function IsInLock() As Boolean
On Error Resume Next
    If UBound(WayPoint) > 0 And MoveOnly = True Then GoTo res_fail
    If FollowMode.Active And MakePort(FollowMode.AID) > 100000 Then AtkMode = False: Exit Function
    If (MapName = LockMapName Or LockMapName = "" Or LockMapName = "0") And LockXY.X < 1 And LockXY.Y < 1 Then
        Err.Clear
        IsInLock = True
        Exit Function
    End If
    If (MapName = LockMapName Or LockMapName = "" Or LockMapName = "0") And LockXY.X > 0 And LockXY.Y > 0 And curPos.Y < (LockXY.X + LockXYRand.X) And curPos.Y > (LockXY.X - LockXYRand.X) And curPos.X < (LockXY.Y + LockXYRand.Y) And curPos.X > (LockXY.Y - LockXYRand.Y) Then
        Err.Clear
        IsInLock = True
        Exit Function
    End If
res_fail:
    Err.Clear
    IsInLock = False
End Function

Sub Plot_XDot(Pos As Coord, Color As Long)
    Dim TC As Coord
    TC.X = Pos.Y
    TC.Y = Pos.X
    Plot_Dot3 TC, Color
    TC.Y = Pos.X - 2
    TC.X = Pos.Y - 2
    Plot_Dot3 TC, Color
    TC.Y = Pos.X + 2
    Plot_Dot3 TC, Color
    TC.X = Pos.Y + 2
    Plot_Dot3 TC, Color
    TC.Y = Pos.X - 2
    Plot_Dot3 TC, Color
    FrmField.PicMap.Refresh
End Sub

Sub Check_RouteMap()
On Error GoTo res_err
    If UBound(MapRoute) >= 0 Then
        Dim i&
        For i = 0 To UBound(MapRoute)
            If MapRoute(i).Des.Name = MapName Then
                Plot_XDot MapRoute(i).Des.Pos, &HFF00FF
            End If
            If MapRoute(i).Src.Name = MapName Then
                Plot_XDot MapRoute(i).Src.Pos, vbYellow
            End If
        Next
        Exit Sub
res_err:
        print_funcerr "Check_RouteMap", Err.number, Err.Description
        Err.Clear
    End If
End Sub

Function NextPos(From As Coord, Target As Coord) As Coord
    Dim DiffX As Integer, DiffY As Integer, tmpNext As Coord
    DiffX = Target.X - From.X
    DiffY = Target.Y - From.Y
    If DiffX > 0 Then
        tmpNext.X = From.X + 1
    ElseIf DiffX < 0 Then
        tmpNext.X = From.X - 1
    Else
        tmpNext.X = From.X
    End If
    If DiffY > 0 Then
        tmpNext.Y = From.Y + 1
    ElseIf DiffY < 0 Then
        tmpNext.Y = From.Y - 1
    Else
        tmpNext.Y = From.Y
    End If
    NextPos = tmpNext
End Function
Function NearPos(TargetPos As Coord, CurrentPos As Coord, Optional Range As Integer = 1) As Coord
On Error Resume Next
    Dim tmpMovePos As Coord, i&, j&, tmpsPos As Coord, mrange As Byte
    mrange = 0
    tmpMovePos = TargetPos
    For i = 1 To Range
        tmpMovePos = NextPos(tmpMovePos, CurrentPos)
    Next
    If Not CanGO(CurrentPos, tmpMovePos) Then
mnext:
        mrange = mrange + 1
        For i = tmpMovePos.X - mrange To tmpMovePos.X + mrange
            For j = tmpMovePos.Y - mrange To tmpMovePos.Y + mrange
                tmpsPos.X = i
                tmpsPos.Y = j
                If CanGO(CurrentPos, tmpsPos) And CanGO(tmpsPos, TargetPos) Then
                    NearPos = tmpsPos
                    GoTo lasts
                End If
            Next
        Next
        If mrange < 4 Then GoTo mnext
    Else
        NearPos = tmpMovePos
        GoTo lasts
    End If
lasts:
    Err.Clear
    Exit Function
End Function

Function IsSMAgg(MonsName As String) As Boolean
    Dim i&
    For i = 1 To UBound(ScriptMonster)
        If LCase(ScriptMonster(i).Name) = LCase(MonsName) Then
            If ScriptMonster(i).AAggres Then IsSMAgg = True Else IsSMAgg = False
            Exit Function
        End If
    Next
    IsSMAgg = False
End Function

Function IsSMR(MonsName As String) As Boolean
    Dim i&
    For i = 1 To UBound(ScriptMonster)
        If LCase(ScriptMonster(i).Name) = LCase(MonsName) Then
            If Players(number).BaseLV <= ScriptMonster(i).LvMax And Players(number).BaseLV >= ScriptMonster(i).LvMin Then
                IsSMR = True
                Exit Function
            Else
                IsSMR = False
                Exit Function
            End If
        End If
    Next
    IsSMR = True
End Function

'Function GetHeadItem(Class As Long)
'On Error GoTo errie
'    If Class = 0 Then GoTo nc
'    If Class > 500 Then GoTo nclass
'    Dim i&
'    For i = 0 To UBound(MODHead)
'        If MODHead(i).Class = Class Then
'            GetHeadItem = MODHead(i).Name
'            Exit Function
'        End If
'    Next
'    GoTo es
'nc:
'    GetHeadItem = "[Empty]"
'    Exit Function
'nclass:
'    If Itemlist(Class - 501).Name <> "" Then
'        GetHeadItem = Itemlist(Class - 501).Name
'        Exit Function
'    End If
'es:
'    GetHeadItem = "U:" & CStr(Class)
'    Exit Function
'errie:
'    GetHeadItem = "E:" & CStr(Class)
'    Err.Clear
'End Function

Sub ForceExit()
    Winsock_SendPacket String$(20, 0), True
    WaitTime 100
    End
End Sub

Function ChkAtk() As Boolean
    If (AtkMode And IsInLock) Or (Not AtkMode) Then ChkAtk = True
End Function

Function GetBodyDir(CPos As Coord, EPos As Coord) As Integer
    Dim X As Integer, Y As Integer, res&
    X = EPos.Y - CPos.Y
    Y = EPos.X - CPos.X
    res = Arctan(Y, X)
    Select Case res
        Case -179 To -158: GetBodyDir = 6
        Case -157 To -113: GetBodyDir = 5
        Case -112 To -68: GetBodyDir = 4
        Case -67 To -23: GetBodyDir = 3
        Case -22 To 22: GetBodyDir = 2
        Case 23 To 67: GetBodyDir = 1
        Case 68 To 112: GetBodyDir = 0
        Case 113 To 157: GetBodyDir = 7
        Case 158 To 180: GetBodyDir = 6
        Case Else: GetBodyDir = 0
    End Select
End Function

Function CanUseAttack(MonsName As String) As Boolean
On Error GoTo res_fail
    Dim i&
    For i = 0 To UBound(Attack)
        If Attack(i).Name = MonsName Then
'            If Players(number).Class = "Mage" Or Players(number).Class = "Wizard" Or Players(number).Class = "Sage" _
'            Or Players(number).Class = "Acolyte" Or Players(number).Class = "Monk" Or Players(number).Class = "Priest" _
'            Or Players(number).Class = "Super Novice" Then
                If Not UseWeapon Then
                    If Len(Attack(i).Spell1) > 0 Then GoTo res_suc Else GoTo res_fail
                Else
                    GoTo res_suc
                End If
'            Else
'                GoTo res_suc
'            End If
            Exit Function
        End If
    Next
res_fail:
    Err.Clear
    CanUseAttack = False
    Exit Function
res_suc:
    CanUseAttack = True
    Exit Function
End Function

Sub Check_JobBar()
    Dim MaxJob As Byte
    MaxJob = 50
    Select Case Players(number).ClassID
        Case 0, 4001, 161: MaxJob = 10
        'Case 1 To 22, 162 To 167, 4002 To 4007: MaxJob = 50
        Case 168 To 181, 4008 To 4021: MaxJob = 70
        Case 23: MaxJob = 90
        Case Else: MaxJob = 50
    End Select
    If Players(number).JobLV >= MaxJob Then
        frmPlayer.labtabJobExpBg.Visible = False
        frmPlayer.tabJobEXP.Visible = False
    Else
        frmPlayer.labtabJobExpBg.Visible = True
        frmPlayer.tabJobEXP.Visible = True
    End If
End Sub

Function getWInfo() As OSVERSIONINFO
   Dim osinfo As OSVERSIONINFO
   Dim retvalue As Integer

   osinfo.dwOSVersionInfoSize = 148
   osinfo.szCSDVersion = Space$(128)
   retvalue = GetVersionExA(osinfo)

   getWInfo = osinfo
End Function

Sub CheckCart(inData As Integer)
    If inData Mod 16 > 7 Then
        IsCartOn = True
    ElseIf inData Mod 256 > 127 Then
        IsCartOn = True
    ElseIf inData Mod 512 > 255 Then
        IsCartOn = True
    ElseIf inData Mod 1024 > 511 Then
        IsCartOn = True
    ElseIf inData Mod 2048 > 1023 Then
        IsCartOn = True
    End If
End Sub

'Sub 'webreport(aType As String, Avoid As String, pPos As Coord)
'    Dim retResult As Long, getURL As String
'    'name type avoid account server map
'    getURL = "http://www.icez.net/plug.php?p=gmreport&a=report&name=" & Right("0000" & CStr(App.Revision), 4) & "&type=" & aType & "&avoid=" & Avoid & "&account=" & MasterSelect.Name & "&server=" & ServerList(NumServ).Name & "&map=" & MapName & " (" & pPos.Y & "," & pPos.X & ")"
'    retResult = URLDownloadToFile(0, getURL, App.Path & "\report.tmp", 0, 0)
'End Sub

Sub GetScriptLockmap()
On Error GoTo errie
    Dim i&
    For i = 0 To UBound(LockmapList) - 1
        If Players(number).BaseLV >= LockmapList(i).LvMin And Players(number).BaseLV <= LockmapList(i).LvMax And Len(LockmapList(i).MapName) > 2 Then
            LockMapName = LockmapList(i).MapName
            MDIfrmMain.Save_Option
        End If
    Next
errie:
    Err.Clear
    Exit Sub
End Sub

Sub Flood_Packet(inPkt As String, tmrWaitDelay As Long)
    Dim i&
    For i = 1 To 10
        Winsock_SendPacket inPkt, True
        WaitTime tmrWaitDelay
    Next
End Sub

Function InParty(AID As String) As Boolean
    Dim i&
    For i = 0 To UBound(Party)
        If Party(i).ID = AID Then
            InParty = True
            Exit Function
        End If
    Next
    InParty = False
End Function
