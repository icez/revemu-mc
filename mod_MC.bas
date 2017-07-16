Attribute VB_Name = "md_MC"
Option Explicit

Public SessionID As String * 4
Public AccountID As String * 4
Public CharID As String * 4

Public Const noAtkRange As Long = 3

Public Vending(30) As VendingInfo
Public Mods As ModOption
Public CartNum As Long
Public CartNumM As Long
Public CartWeight As Long
Public CartWeightM As Long
Public Shop(13) As ShopList
Public IsCartOn As Boolean
Public IsShopCreated As Boolean
Public IsWaitShop As Boolean
Public IsWaitChat As Boolean
Public IsVending As Boolean
Public IsChatOC As Boolean
Public IsVendingWait As Boolean
Public SendCBuy As Boolean
Public MColor As MODColor
Public CartCtrl() As CartItemCtrl
Public CartCtrl2() As CartItemCtrl2
'Public IsCartRecv As Boolean
Public ShopRND() As String
Public CSName As Long
Public StartZeny As Long
Public MRemote As MODRemote
Public MyShopID As String * 4
Public ScriptMonster() As MODScriptMonster
Public PacketLenTable As String

Public MChat() As MODChat
Public MShop() As MODShop

Public CurSpirit As Byte

Public MCStartDelay As Long
Public MCDoType As Byte
Public MCShopPacket As String

Type MODScriptMonster
    Name As String
    LvMin As Byte
    LvMax As Byte
    AAggres As Boolean
End Type
Type MODRemote
    CommandBegin As String
    Enabled As Boolean
    Owner As String
    Identified As Boolean
    password As String
End Type
Type MODChat
    Owner As String * 4
    ID As String * 4
    CLimit As Integer
    CUsers As Integer
    IsPub As Byte
    Title As String
    Visible As Boolean
End Type
Type MODShop
    ID As String * 4
    Name As String
    Visible As Boolean
End Type
Type MODColor
    Emotion As Double
    Shop As Double
    Fail As Double
    shopsellitem As Double
    playerchat As Double
    mychat As Double
    whisper As Double
    Party As Double
    trade As Double
    guildannounce As Double
    guildchat As Double
    gmannounce As Double
End Type
Type CartItemCtrl
    Name As String
    Min As Long
    autoadd As Boolean
    autoget As Boolean
    autostore As Boolean
End Type
Type CartItemCtrl2
    Name As String
    cartmin As Long
    cartmax As Long
End Type
Type ShopList
    Index As Long
    ID As Long
    Name As String
    Amount As Long
    Price As Long
End Type
Type VendingInfo
    Name As String
    Price As Long
    Amount As Long
    NPC As String
End Type
Type ModOption
    Vending As Boolean
    Vendingdelay As Integer
    Chatroom As String
    shopname As String
    dcshop As Boolean
    MapCompr As Boolean
    Enabled As Boolean
    AutoSit As Boolean
    minzeny As Long
    isUsePass As Boolean
    
    ReconWhenDead As Boolean
    BodyDir As Integer
    HeadDir As Integer
    
    OC As Boolean
    OCdelay As Integer
    OCcreateshop As Boolean
    OCdisconnect As Boolean
    OCnocalcmoney As Boolean

    GuildText As Boolean
    EmotionText As Boolean
    STChat As Boolean
    STStatus As Boolean
    STWalk As Boolean
    STSystem As Boolean
    STDebug As Boolean
    STSKFail As Boolean
    STParty As Boolean
End Type

Sub ResetMod()
    WaitEquipTele = False
    IsVendingWait = False
    IsVending = False
    IsShopCreated = False
    IsWaitShop = False
    IsWaitChat = False
    IsCartOn = False
    IsChatOC = False
    ActionList = ""
    'ReDim Party(0)
    ReDim Cart(0)
    ReDim MChat(0)
    ReDim MShop(0)
    ReDim GuildAlliance(0)
    ReDim Guild(0)
    'ReadModOption
    MRemote.Identified = False
    UpdateCart
    frmMain.UpdateChatShop
    MRemote.Identified = False
    ShopStep = 0
    MaxShopAmount = 0
    CurSpirit = 0
    
End Sub
Function ModAI() As Boolean
    If Not StartBot Then ModAI = False: Exit Function
    If (Mods.Vending = True Or Mods.OC = True) And Mods.Enabled And _
    Not MIsGoStore And Not mIsGoBuy And IsInLock Then
        ModAI = True
    Else
        ModAI = False
    End If
End Function

Sub CalcModAI(Un As String)
    If Mods.STDebug Then Stat "Debug : CalcModAI [" & Un & "]" & vbCrLf, &HFF00
    If Not StartBot Then Exit Sub
    If Not ModAI Then
        If IsVending Or IsShopCreated Then frmMain.Send_ShopClose
        If IsChatOC Then frmMain.destroy_chatroom
        Exit Sub
    End If
    If CreateShop = False Then
        If CreateOC = False Then
            Mods.Vending = False
            Mods.OC = False
            Chat "Mods can't do anything. Automatic disable mod ai.", MColor.Shop
        End If
    End If
End Sub
Sub CalcShopAI()
    'Dim i&, j&
    If CartBuy Then
        mIsGoBuy = True
        Exit Sub
    End If
    If Mods.Vending Then
        If CreateShop Then Exit Sub
    End If
    If Mods.OC Then CreateOC
    If HaveBuyItem Then
        mIsGoBuy = True
        Exit Sub
    End If
    If HaveStoreItem Then
        MIsGoStore = True
        Exit Sub
    End If
End Sub
Function CreateOC() As Boolean
    If Not ModAI Then
        CreateOC = True
        Exit Function
    End If
    If Mods.OC = False Then
        If IsChatOC Then frmMain.destroy_chatroom
        CreateOC = False
        Exit Function
    End If
    If IsVending Or IsShopCreated Then
        Chat "Create chatroom failed. (you're vending)", MColor.Fail
        CreateOC = True
        Exit Function
    End If
    If Len(Mods.Chatroom) < 1 Then
        Chat "Create chatroom failed. (no chatroom title)", MColor.Fail
        CreateOC = False
        Mods.OC = False
        Exit Function
    End If
    If Mods.minzeny > Players(number).Zeny Then
        If Mods.OCcreateshop Then
            Mods.Vending = True
            Chat "Create chatroom failed. (minimum zeny exceeded) - auto create shop", MColor.Fail
            CreateOC = True
            CreateShop
            Exit Function
        End If
        If Mods.OCdisconnect Then
            Mods.Vending = True
            Chat "Create chatroom failed. (minimum zeny exceeded) - auto disconnect", MColor.Fail
            CreateOC = True
            End
            Exit Function
        End If
    End If
    If IsWaitChat Or IsWaitShop Then
        CreateOC = True
        Exit Function
    End If
    Dim TPass As String
    If IsChatOC Then
        Randomize
        TPass = CStr(Int(Rnd() * 10000000))
        frmMain.edit_chatroom TPass, Mods.Chatroom
        If Mods.AutoSit Then
            frmMain.Send_Sit
            IsSitting = True
            IsStanding = False
        End If
        CreateOC = True
        Exit Function
    End If
    Winsock_SendPacket IntToChr(&H9B) & IntToChr(CLng(Mods.HeadDir)) & Chr(Mods.BodyDir), True
    Chat "Delaying for create chatroom for : " & Mods.OCdelay & " seconds.", MColor.Shop
    IsWaitChat = True
    MCStartDelay = Mods.OCdelay * 100
    MCDoType = 1
    CreateOC = True
End Function
Function CreateShop() As Boolean
    If Not ModAI Then
        CreateShop = True
        Exit Function
    End If
    If Mods.Vending = False Then
        If IsShopCreated = True Then frmMain.Send_ShopClose
        CreateShop = False
        Exit Function
    End If
    If IsShopCreated Then
        CreateShop = True
        Exit Function
    End If
    If IsWaitChat Or IsWaitShop Then
        CreateShop = True
        Exit Function
    End If
    Dim i As Long, j As Long
    i = Find_SkillId("MC_VENDING")
    Dim Length As Integer
    Randomize
    CSName = Int(Rnd() * (UBound(ShopRND) + 1))
    If Mods.STDebug Then Chat "Shop random name : " & CSName, &HAAAAAA
    If CSName > UBound(ShopRND) Then CSName = UBound(ShopRND)
    If CSName < LBound(ShopRND) Then CSName = LBound(ShopRND)
    Mods.shopname = ShopRND(CSName)
    If Len(Mods.shopname) > 80 Then Mods.shopname = Mid(Mods.shopname, 1, 80)
    frmShop.Label1.Caption = "Shop : " & Mods.shopname
    
    If i < 0 Or Len(Mods.shopname) = 0 Then
        Chat "Create shop failed. (Vending skill not found/no shopname)", MColor.Fail
        CreateShop = False
        Mods.Vending = False
        Exit Function
    End If

    Winsock_SendPacket IntToChr(&H9B) & IntToChr(CLng(Mods.HeadDir)) & Chr(Mods.BodyDir), True

    Chat "Delaying for create shop for : " & Mods.Vendingdelay & " seconds.", MColor.Shop
    IsWaitShop = True
    
    MCStartDelay = Mods.Vendingdelay * 100
    MCDoType = 0
    CreateShop = True
End Function

Function Find_SkillId(Name As String) As Long
    On Error GoTo errie
    Dim X As Integer
    For X = 0 To UBound(SkillChar)
        If (SkillChar(X).MaxLV > 0) And (LCase(SkillChar(X).Name) = LCase(Name)) Then
            Find_SkillId = X
            Exit Function
        End If
    Next
errie:
Find_SkillId = -1
End Function
Function SkillExists(Name As String) As Boolean
    On Error GoTo errie
    Dim X As Long
    For X = 0 To UBound(SkillChar)
        If (SkillChar(X).MaxLV > 0) And SkillChar(X).Name = Name Then
            SkillExists = True
            Exit Function
        End If
    Next
errie:
    SkillExists = False
    Err.Clear
End Function
Function Find_CartID(Name As String, Optional begins As Long = 0) As Long
On Error GoTo errie
    If Len(Name) < 1 Then
        Find_CartID = -1
        Exit Function
    End If
    Dim i As Long
    For i = begins To UBound(Cart)
        If LCase(Cart(i).Name) = LCase(Name) And Cart(i).Amount > 0 Then
            Find_CartID = i
            Exit Function
        End If
    Next
errie:
    Err.Clear
    Find_CartID = -1
End Function
Function Find_InvID(Name As String) As Long
    If Len(Name) < 1 Then
        Find_InvID = -1
        Exit Function
    End If
    Dim i As Long
    For i = 0 To UBound(AllInv)
        If LCase(AllInv(i).Name) = LCase(Name) Then
            Find_InvID = i
            Exit Function
        End If
    Next
    Find_InvID = -1
End Function

Function MakeItemName(ITName As String, CardPkt As String, Refine As String) As String
On Error GoTo errie
    Dim Card() As Card_Profile, Itemname As String, i As Integer, Name As String, TmpCard As String, cnumber As Long
    Itemname = Return_ItemName(MakeHexName(ITName))
    If Itemname = "" Then Itemname = "{U:" & CStr(MakeHexName(Mid(CardPkt, 9, 2))) & "}"
    ReDim Card(0)
    Card(0).Name = ""
    For i = 0 To 3
        cnumber = MakePort(Mid(CardPkt, (i * 2) + 1, 2))
        Name = Return_CardNameTable(Trim(STR(cnumber)))
        If cnumber > 3999 And cnumber < 5000 Then
            If Card(0).Name = "" Then
                Card(0).Name = Name
                Card(0).number = 1
            ElseIf Name <> Card(UBound(Card)).Name Then
                ReDim Preserve Card(UBound(Card) + 1)
                Card(UBound(Card)).Name = Name
                Card(UBound(Card)).number = 1
            Else
                Card(UBound(Card)).number = Card(UBound(Card)).number + 1
            End If
        Else
            Exit For
        End If
    Next
    TmpCard = ""
    If Card(0).Name <> "" Then
        For i = UBound(Card) To 0 Step -1
            TmpCard = Card(i).Name & TmpCard
            If Card(i).number = 2 Then
                TmpCard = "Double " & TmpCard
            ElseIf Card(i).number = 3 Then
                TmpCard = "Triple " & TmpCard
            ElseIf Card(i).number = 4 Then
                TmpCard = "Quadruple " & TmpCard
            End If
        Next
    End If
    If TmpCard <> "" Then
        If LCase(Left(TmpCard, 2)) = "of" Then
            Itemname = Itemname & " " & TmpCard
        Else
            Itemname = TmpCard & " " & Itemname
        End If
    End If
    If Mid(CardPkt, 1, 1) = Chr(255) Then Itemname = Get_Element(Asc(Mid(CardPkt, 3, 1))) & Itemname
    If (Asc(Mid(CardPkt, 4, 1)) Mod 5 = 0) And Asc(Mid(CardPkt, 4, 1)) > 0 And Asc(Mid(CardPkt, 4, 1)) < 20 And Card(0).Name = "" Then Itemname = Get_Very(Asc(Mid(CardPkt, 4, 1))) & Itemname
    If Asc(Refine) > 0 Then Itemname = "+" & Asc(Refine) & " " & Itemname
    MakeItemName = Itemname
    Exit Function
errie:
    If Err.number > 0 Then print_funcerr "MakeItemName", Err.number, Err.Description
    Err.Clear
End Function
Function Get_Very(inByte As Byte) As String
On Error Resume Next
    If inByte = 0 Then
        Get_Very = ""
        Exit Function
    End If
    Dim res$
    res = "Strong "
    Dim i As Integer
    For i = 1 To (inByte \ 5)
        res = "Very " & res
    Next
    Get_Very = res
    Err.Clear
End Function
Sub UpdateShop()
    Dim i As Long, isVis As Boolean
    isVis = False
    With frmShop
        .lstShop.Clear
        For i = 1 To 13
            If Shop(i).Amount > 0 Then
                isVis = True
                .lstShop.AddItem i & " - [" & Shop(i).Name & "] " & Shop(i).Amount & "EA " & FormatNumber(Shop(i).Price, 0, vbTrue, vbTrue, vbTrue) & "z"
            End If
        Next
        .Visible = isVis
    End With
End Sub
Sub ModIncMonLog(MonName As String)
On Error GoTo errie
    Dim i&, ri&
    ri = 0
    For i = 1 To UBound(MODMLogN)
        If Players(number).Name = MODMLogN(i) Then
            ri = i
            GoTo finext
        End If
    Next
    ReDim Preserve MODMLogN(UBound(MODMLogN) + 1)
    MODMLogN(UBound(MODMLogN)) = Players(number).Name
    ReDim Preserve MODMLogM(UBound(MODMLogN))
    ReDim Preserve MODMLogM(UBound(MODMLogM)).Names(0)
    ReDim Preserve MODMLogM(UBound(MODMLogM)).Amount(0)
    ri = UBound(MODMLogN)
finext:
    For i = 1 To UBound(MODMLogM(ri).Names)
        If MODMLogM(ri).Names(i) = MonName Then
            MODMLogM(ri).Amount(i) = MODMLogM(ri).Amount(i) + 1
            SaveMonsLog
            Exit Sub
        End If
    Next
    ReDim Preserve MODMLogM(ri).Names(UBound(MODMLogM(ri).Names) + 1)
    ReDim Preserve MODMLogM(ri).Amount(UBound(MODMLogM(ri).Names))
    i = UBound(MODMLogM(ri).Names)
    MODMLogM(ri).Names(i) = MonName
    MODMLogM(ri).Amount(i) = MODMLogM(ri).Amount(i) + 1
    SaveMonsLog
    Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "ModIncMonLog", Err.number, Err.Description
    Err.Clear
End Sub

Sub RefreshMC()
    Dim i As Long
    With frmModConfig
        .txtBelow = Mods.minzeny
        .txtCDelay = Mods.OCdelay
        For i = 0 To 29
            .txtITAmount(i) = Vending(i + 1).Amount
            .txtITName(i) = Vending(i + 1).Name
            .txtITNPC(i) = Vending(i + 1).NPC
            .txtITPrice(i) = Vending(i + 1).Price
        Next
        .txtVDelay = Mods.Vendingdelay
        .chkAutoSit = Abs(CLng(Mods.AutoSit))
        .chkChatroom = Abs(CLng(Mods.OC))
        .chkCreateShop = Abs(CLng(Mods.OCcreateshop))
        .chkDCChat = Abs(CLng(Mods.OCdisconnect))
        .chkNoCalc = Abs(CLng(Mods.OCnocalcmoney))
        .chkDCShop = Abs(CLng(Mods.dcshop))
        .chkEmotion = Abs(CLng(Mods.EmotionText))
        .chkEnabled = Abs(CLng(Mods.Enabled))
        .chkGuildAnn = Abs(CLng(Mods.GuildText))
        .chkNewMapType = btol(Mods.MapCompr)
        .chkSTChat = btol(Mods.STChat)
        .chkSTStatus = btol(Mods.STStatus)
        .chkSTWalk = btol(Mods.STWalk)
        .chkSTSys = btol(Mods.STSystem)
        .chkVending = btol(Mods.Vending)
        .txtCTitle = Mods.Chatroom
    End With
End Sub
Function FileExists(FilePath As String) As Boolean
    On Error GoTo errie
    Dim i&
    i = FileLen(FilePath)
    FileExists = True
    Exit Function
errie:
    FileExists = False
    Err.Clear
End Function
Function GetPacketLen(Header As String) As Integer
    Dim HeaderPos As Integer
    HeaderPos = InStr(1, PacketLenTable, Right$(Header, 2) + Left$(Header, 2))
    If HeaderPos Then
        HeaderPos = HeaderPos + 5
        GetPacketLen = CInt(Mid$(PacketLenTable, HeaderPos, InStr(HeaderPos, PacketLenTable, vbNewLine) - HeaderPos))
    Else
        GetPacketLen = 0
    End If
End Function

Function ChatlogPath() As String
    ChatlogPath = App.Path & "\log\" & Year(Date) & IIf(Val(Month(Date)) < 10, "0" & Month(Date), Month(Date)) & IIf(Val(Day(Date)) < 10, "0" & Day(Date), Day(Date)) & "chatlog.txt"
End Function

Function SplStr(inData As String, Delimiter As String, Index As Long, Optional Limit As Long = -1, Optional Compare As VbCompareMethod = vbBinaryCompare) As String
    Dim spl() As String
    spl = Split(inData, Delimiter, Limit, Compare)
    If Index > UBound(spl) Then SplStr = "" Else SplStr = spl(Index)
End Function

Sub DoConnect(sIP As String, Port As Long)
On Error GoTo errie
    With frmMain.Winsock1
        If .State <> 0 Then .Close
        If IsUseProxy Then
            CurConnIP = sIP
            CurConnPort = Port
            .Connect ProxyIP, ProxyPort
        Else
            .Connect sIP, Port
        End If
    End With
    Exit Sub
errie:
    MsgBox "Error!!! on connecting to " & sIP & ":" & Port & vbCrLf & "Desc: " & Err.Description
    End
End Sub

Function GetWarpInfo(WMap As String, X As Long, Y As Long) As String
    Dim i&
    For i = 0 To UBound(PortalsInfo)
        If PortalsInfo(i).Src.Name = WMap And PortalsInfo(i).Src.Pos.X = X And PortalsInfo(i).Src.Pos.Y = Y Then
            GetWarpInfo = PortalsInfo(i).Des.Name & " (" & PortalsInfo(i).Des.Pos.X & "," & PortalsInfo(i).Des.Pos.Y & ")"
            Exit Function
        End If
    Next
End Function

Public Sub Load_PortalsInfo()
    On Error GoTo errie
    Dim tstr As String
    Dim Index As Integer
    Dim i As Integer, LCount As Long
    ReDim PortalsInfo(0)
    Open App.Path & "\maproute\portals.txt" For Input As #10
    Do While Not EOF(10)
        Line Input #10, tstr
        tstr = Trim(tstr)
        LCount = LCount + 1
        If PortalsInfo(0).Des.Name <> "" Then ReDim Preserve PortalsInfo(UBound(PortalsInfo) + 1)
        For i = 1 To 6
            Index = InStr(tstr, " ")
            Select Case i
                Case 1
                    PortalsInfo(UBound(PortalsInfo)).Src.Name = Left(tstr, Index - 1)
                Case 2
                    PortalsInfo(UBound(PortalsInfo)).Src.Pos.X = Val(Left(tstr, Index - 1))
                Case 3
                    PortalsInfo(UBound(PortalsInfo)).Src.Pos.Y = Val(Left(tstr, Index - 1))
                Case 4
                    PortalsInfo(UBound(PortalsInfo)).Des.Name = Left(tstr, Index - 1)
                Case 5
                    PortalsInfo(UBound(PortalsInfo)).Des.Pos.X = Val(Left(tstr, Index - 1))
                Case 6
                    PortalsInfo(UBound(PortalsInfo)).Des.Pos.Y = Val(tstr)
            End Select
            If i < 6 Then tstr = Trim(Right(tstr, Len(tstr) - Index))
        Next
    Loop
    Close 10
    Exit Sub
errie:
    Close 10
    MsgBox "Error!!! on loading 'maproute\portals.txt' (Load_PortalsInfo) Line:" & LCount & vbCrLf & vbCrLf & Err.Description, vbCritical
End Sub

Function btol(booin As Boolean) As String
    If (booin) Then
        btol = "1"
    Else
        btol = "0"
    End If
End Function

