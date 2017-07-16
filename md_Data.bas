Attribute VB_Name = "md_Data"
Option Explicit

Sub Main()
    Load frmSplash
End Sub

Sub ReadModOption()
    Load_GetStorage
    ReadModOpt
    ReadColorInfo
    ReadCartControl
    ReadCartItemControl
    ReadScriptMonster
    ReadItemCtrl
    ReadHeadItem
    ReadMonsLog
    ReadEventList
    Load_accessory_Profile
    Load_PortalsInfo
    If ReadLenTable(App.Path & "\table\recvpacket.txt") = False Then
        MsgBox "Error on loading 'table\recvpacket.txt'", vbCritical, "Error"
        End
    End If
    Load_LockmapList
End Sub
Sub ReadMonsLog()
On Error GoTo errie
    Dim tmpstr$, Index As Long
    ReDim MODMLogM(0)
    ReDim MODMLogN(0)
    Open App.Path & "\log\monsterlog.txt" For Input As #3
        Do Until EOF(3)
            Line Input #3, tmpstr
            If Left(tmpstr, 1) = "[" And Right(tmpstr, 1) = "]" Then
                ReDim Preserve MODMLogN(UBound(MODMLogN) + 1)
                ReDim Preserve MODMLogM(UBound(MODMLogN))
                ReDim Preserve MODMLogM(UBound(MODMLogN)).Names(0)
                ReDim Preserve MODMLogM(UBound(MODMLogN)).Amount(0)
                MODMLogN(UBound(MODMLogN)) = Mid(tmpstr, 2, Len(tmpstr) - 2)
            Else
                Index = InStr(1, tmpstr, "=")
                If Index > 0 Then
                    ReDim Preserve MODMLogM(UBound(MODMLogN)).Amount(UBound(MODMLogM(UBound(MODMLogN)).Amount) + 1)
                    ReDim Preserve MODMLogM(UBound(MODMLogN)).Names(UBound(MODMLogM(UBound(MODMLogN)).Amount))
                    MODMLogM(UBound(MODMLogN)).Amount(UBound(MODMLogM(UBound(MODMLogN)).Amount)) = Val(Mid(tmpstr, Index + 1, Len(tmpstr) - Index))
                    MODMLogM(UBound(MODMLogN)).Names(UBound(MODMLogM(UBound(MODMLogN)).Amount)) = Left(tmpstr, Index - 1)
                End If
            End If
        Loop
    Close #3
    Exit Sub
errie:
    Err.Clear
End Sub
Sub SaveMonsLog()
On Error GoTo errie
    Dim i&, j&
    Open App.Path & "\log\monsterlog.txt" For Output As #14
    For i = 1 To UBound(MODMLogN)
        Print #14, "[" & MODMLogN(i) & "]"
        For j = 1 To UBound(MODMLogM(i).Names)
            Print #14, MODMLogM(i).Names(j) & "=" & MODMLogM(i).Amount(j)
        Next
        Print #14, ""
    Next
    Close #14
    Exit Sub
errie:
    Err.Clear
End Sub
Sub ReadHeadItem()
On Error GoTo errie
    ReDim MODHead(0)
    MODHead(0).Name = ""
    Dim ass() As String, i&, tmpstr$
    Open App.Path & "\table\charhead.txt" For Input As #3
        Do Until EOF(3)
            Line Input #3, tmpstr
            ass = Split(tmpstr, " ", 3)
            If MODHead(0).Name <> "" Then ReDim Preserve MODHead(UBound(MODHead) + 1)
            MODHead(UBound(MODHead)).Class = CInt(ass(0))
            MODHead(UBound(MODHead)).ItemID = CInt(ass(1))
            MODHead(UBound(MODHead)).Name = ass(2)
        Loop
    Close #3
    Exit Sub
errie:
    Err.Clear
    MsgBox "Error on loading 'table\charhead.txt'"
End Sub
Sub ReadModOpt()
    Mods.Chatroom = ReadINI("Options", "Chatroom")
    Mods.dcshop = CBool(ReadINI("Options", "dcshop", "0"))
    Mods.Enabled = Not CBool(ReadINI("Options", "normal", "0"))
    Mods.OC = CBool(ReadINI("Options", "overcharge", "0"))
    Mods.OCdelay = CInt(ReadINI("Options", "ocdelay", "5"))
    If Mods.OCdelay < 5 Then Mods.OCdelay = 5
    Mods.AutoSit = CBool(ReadINI("Options", "autosit", "0"))
    Mods.ReconWhenDead = CBool(ReadINI("Options", "reconwhendead", "0"))
    'IsCartRecv = CBool(ReadINI("Options", "iscarton", "0"))
    
    Mods.BodyDir = Val(ReadINI("Options", "bodydir", "0"))
    If Mods.BodyDir > 7 Then Mods.BodyDir = 7
    If Mods.BodyDir < 0 Then Mods.BodyDir = 0
    Mods.HeadDir = Val(ReadINI("Options", "headdir", "0"))
    If Mods.HeadDir > 2 Then Mods.HeadDir = 2
    If Mods.HeadDir < 0 Then Mods.HeadDir = 0
    
    Mods.minzeny = Val(ReadINI("buying", "minzeny", "0"))
    Mods.OCcreateshop = CBool(ReadINI("buying", "createshop", "0"))
    Mods.OCdisconnect = CBool(ReadINI("buying", "disconnect", "0"))
    Mods.OCnocalcmoney = CBool(ReadINI("buying", "nocalcmoney", "0"))
    Mods.isUsePass = CBool(ReadINI("buying", "passwordlock", "1"))
    
    MODDC.TAccept = Val(ReadINI("buying", "delay_accept", "1500"))
    MODDC.TCalc = Val(ReadINI("buying", "delay_send", "3000"))
    MODDC.TItem = Val(ReadINI("buying", "no_senditem_timeout", "5000"))
    MODDC.TNItem = Val(ReadINI("buying", "nextitem_timeout", "5000"))
    
    
    MRemote.Enabled = CBool(ReadINI("remote", "enabled", "0"))
    MRemote.CommandBegin = ReadINI("remote", "commandbegin", "!")
    MRemote.password = ReadINI("remote", "password", "iCeZzZz")
    
    frmMain.TmrDeal.Interval = MODDC.TAccept
    frmMain.TmrIT.Interval = MODDC.TItem
    
    ReDim ShopRND(0)
    Dim shopmax As Long
    shopmax = 1
    Do While True
        ShopRND(UBound(ShopRND)) = ReadINI("shopname", CStr(shopmax), "/")
        If ShopRND(UBound(ShopRND)) = "/" Then
            If UBound(ShopRND) = 0 Then ShopRND(0) = InputBox("Please enter shopname", "Shop name didn't completely configured.")
            Exit Do
        End If
        shopmax = shopmax + 1
        ReDim Preserve ShopRND(UBound(ShopRND) + 1)
    Loop
    If UBound(ShopRND) > 0 Then ReDim Preserve ShopRND(UBound(ShopRND) - 1)
    
    Mods.Vending = CBool(ReadINI("Options", "vending", "0"))
    Mods.Vendingdelay = CInt(ReadINI("Options", "vendingdelay", "5"))
    If Mods.Vendingdelay < 5 Then Mods.Vendingdelay = 5
    Mods.MapCompr = CBool(ReadINI("Options", "newmaptype", "0"))
    Mods.GuildText = CBool(ReadINI("message", "guildannounce", "0"))
    Mods.EmotionText = CBool(ReadINI("message", "emotion", "0"))
    Mods.STChat = CBool(ReadINI("message", "chatroom", "0"))
    Mods.STStatus = CBool(ReadINI("message", "status", "0"))
    Mods.STWalk = CBool(ReadINI("message", "walk", "1"))
    Mods.STSystem = CBool(ReadINI("message", "system", "1"))
    Mods.STDebug = CBool(ReadINI("message", "debug", "0"))
    Mods.STSKFail = CBool(ReadINI("message", "skillfail", "0"))
    Mods.STParty = CBool(ReadINI("message", "sysparty", "0"))
    If Not Mods.STParty Then WriteINI "message", "sysparty", CStr(Abs(CLng(Mods.STParty)))
    
    MODDC.DualLogin = CBool(ReadINI("disconnect", "duallogin", "0"))
    MODDC.DualLoginTime = Val(ReadINI("disconnect", "dual_wait", "0"))
    MODDC.AvoidTime = Val(ReadINI("disconnect", "avoid_wait", "0"))
    
    Dim i As Integer
    For i = 1 To 30
        Vending(i).Amount = CInt(ReadINI("Item" & CStr(i), "amount", "0"))
        Vending(i).Price = CLng(ReadINI("Item" & CStr(i), "price", "0"))
        Vending(i).Name = ReadINI("Item" & CStr(i), "name", "")
        Vending(i).NPC = ReadINI("Item" & CStr(i), "npc", "")
    Next
    RefreshMC
End Sub
Sub SaveModConfig()
    WriteINI "Options", "Chatroom", Mods.Chatroom
    WriteINI "Options", "dcshop", CStr(Abs(CLng(Mods.dcshop)))
    WriteINI "Options", "normal", CStr(Abs(CLng(Not Mods.Enabled)))
    WriteINI "Options", "overcharge", CStr(Abs(CLng(Mods.OC)))
    WriteINI "Options", "autosit", CStr(Abs(CLng(Mods.AutoSit)))
    WriteINI "Options", "vending", CStr(Abs(CLng(Mods.Vending)))
    WriteINI "Options", "newmaptype", ""
    WriteINI "Options", "ocdelay", CStr(Mods.OCdelay)
    WriteINI "Options", "vendingdelay", CStr(Mods.Vendingdelay)
    
    WriteINI "buying", "minzeny", CStr(Mods.minzeny)
    WriteINI "buying", "createshop", CStr(Abs(CLng(Mods.OCcreateshop)))
    WriteINI "buying", "disconnect", CStr(Abs(CLng(Mods.OCdisconnect)))
    WriteINI "buying", "nocalcmoney", CStr(Abs(CLng(Mods.OCnocalcmoney)))
    WriteINI "buying", "no_senditem_timeout", CStr(MODDC.TItem)
    WriteINI "buying", "nextitem_timeout", CStr(MODDC.TNItem)
    WriteINI "buying", "delay_accept", CStr(MODDC.TAccept)
    WriteINI "buying", "delay_send", CStr(MODDC.TCalc)
    WriteINI "buying", "passwordlock", btol(Mods.isUsePass)
    
    WriteINI "remote", "commandbegin", MRemote.CommandBegin
    WriteINI "remote", "enabled", CStr(Abs(CLng(MRemote.Enabled)))
    WriteINI "remote", "password", MRemote.password

    Dim i As Integer
    
    WriteINI "disconnect", "duallogin", CStr(Abs(CLng(MODDC.DualLogin)))
    WriteINI "disconnect", "dual_wait", CStr(MODDC.DualLoginTime)
    WriteINI "disconnect", "avoid_wait", CStr(MODDC.AvoidTime)

    WriteINI "message", "guildannounce", CStr(Abs(CLng(Mods.GuildText)))
    WriteINI "message", "emotion", CStr(Abs(CLng(Mods.EmotionText)))
    WriteINI "message", "chatroom", CStr(Abs(CLng(Mods.STChat)))
    WriteINI "message", "status", CStr(Abs(CLng(Mods.STStatus)))
    WriteINI "message", "walk", CStr(Abs(CLng(Mods.STWalk)))
    WriteINI "message", "system", CStr(Abs(CLng(Mods.STSystem)))
    WriteINI "message", "debug", CStr(Abs(CLng(Mods.STDebug)))
    WriteINI "message", "skillfail", CStr(Abs(CLng(Mods.STSKFail)))
    WriteINI "message", "sysparty", CStr(Abs(CLng(Mods.STParty)))

    For i = 1 To 30
        WriteINI "Item" & CStr(i), "name", Vending(i).Name
        WriteINI "Item" & CStr(i), "npc", Vending(i).NPC
        WriteINI "Item" & CStr(i), "price", CStr(Vending(i).Price)
        WriteINI "Item" & CStr(i), "amount", CStr(Vending(i).Amount)
    Next
End Sub
Sub ReadColorInfo()
    MColor.Emotion = Format("&H" & ReverseHex(ReadINI("Color", "emotion", "FF00FF", "table\color.ini")))
    MColor.Fail = Format("&H" & ReverseHex(ReadINI("Color", "fail", "FF0000", "table\color.ini")))
    MColor.Shop = Format("&H" & ReverseHex(ReadINI("Color", "shop", "0000FF", "table\color.ini")))
    MColor.shopsellitem = Format("&H" & ReverseHex(ReadINI("Color", "shop", "00FF00", "table\color.ini")))
    MColor.playerchat = Format("&H" & ReverseHex(ReadINI("Color", "playerchat", "000000", "table\color.ini")))
    MColor.mychat = Format("&H" & ReverseHex(ReadINI("Color", "mychat", "00ADE2", "table\color.ini")))
    MColor.whisper = Format("&H" & ReverseHex(ReadINI("Color", "whisper", "FFFF00", "table\color.ini")))
    MColor.Party = Format("&H" & ReverseHex(ReadINI("Color", "party", "00FFCC", "table\color.ini")))
    MColor.trade = Format("&H" & ReverseHex(ReadINI("Color", "trade", "0000FF", "table\color.ini")))
    MColor.guildannounce = Format("&H" & ReverseHex(ReadINI("Color", "guildannounce", "FFAA00", "table\color.ini")))
    MColor.guildchat = Format("&H" & ReverseHex(ReadINI("Color", "guildchat", "CCCC00", "table\color.ini")))
    MColor.gmannounce = Format("&H" & ReverseHex(ReadINI("Color", "gmannounce", "00FFFF", "table\color.ini")))
End Sub
Sub ReadCartControl()
    Dim lfile As Long, tstr$, strSPL() As String, str2() As String
    lfile = FreeFile
    ReDim CartCtrl(0)
    Open App.Path & "\control\cart_control.txt" For Input As lfile
    Do Until EOF(lfile)
        Input #lfile, tstr
        If InStr(1, tstr, Chr(9)) > 0 And Mid(tstr, 1, 1) <> "'" Then
            strSPL = Split(tstr, Chr(9))
            If InStr(strSPL(1), " ") > 0 Then
                str2 = Split(strSPL(1), " ")
                ReDim Preserve CartCtrl(UBound(CartCtrl) + 1)
                CartCtrl(UBound(CartCtrl)).Name = LCase(strSPL(0))
                CartCtrl(UBound(CartCtrl)).Min = Val(str2(0))
                CartCtrl(UBound(CartCtrl)).autoadd = CBool(Val(str2(1)))
                CartCtrl(UBound(CartCtrl)).autoget = CBool(Val(str2(2)))
                CartCtrl(UBound(CartCtrl)).autostore = CBool(Val(str2(3)))
            End If
        End If
    Loop
    Close #1
End Sub
Sub ReadCartItemControl()
    Dim lfile As Long, tstr$, strSPL() As String, str2() As String
    lfile = FreeFile
    ReDim CartCtrl2(0)
    Open App.Path & "\control\cartitem_control.txt" For Input As lfile
    Do Until EOF(lfile)
        Input #lfile, tstr
        If InStr(1, tstr, Chr(9)) > 0 And Mid(tstr, 1, 1) <> "'" Then
            strSPL = Split(tstr, Chr(9))
            If InStr(strSPL(1), " ") > 0 Then
                str2 = Split(strSPL(1), " ")
                ReDim Preserve CartCtrl2(UBound(CartCtrl2) + 1)
                CartCtrl2(UBound(CartCtrl2)).Name = LCase(strSPL(0))
                CartCtrl2(UBound(CartCtrl2)).cartmin = Val(str2(0))
                CartCtrl2(UBound(CartCtrl2)).cartmax = Val(Val(str2(1)))
            End If
        End If
    Loop
    Close #1
End Sub
Sub ReadScriptMonster()
    Dim lfile As Long, tstr$, strSPL() As String, str2() As String
    lfile = FreeFile
    ReDim ScriptMonster(0)
    Open App.Path & "\profile\script_monster.txt" For Input As lfile
    Do Until EOF(lfile)
        Input #lfile, tstr
        If InStr(1, tstr, Chr(9)) > 0 And Mid(tstr, 1, 1) <> "'" Then
            strSPL = Split(tstr, Chr(9))
            If InStr(strSPL(1), " ") > 0 Then
                str2 = Split(strSPL(1), " ")
                ReDim Preserve ScriptMonster(UBound(ScriptMonster) + 1)
                ScriptMonster(UBound(ScriptMonster)).Name = LCase(strSPL(0))
                ScriptMonster(UBound(ScriptMonster)).LvMin = Val(str2(0))
                ScriptMonster(UBound(ScriptMonster)).LvMax = Val(str2(1))
                ScriptMonster(UBound(ScriptMonster)).AAggres = CBool(Val(str2(2)))
            End If
        End If
    Loop
    Close #1
End Sub
Sub ReadItemCtrl()
On Error GoTo errie:
Open App.Path & "\control\items.txt" For Input As #10
Dim tstr As String
Dim Index As Integer
ReDim ItemCtrl(0)
Do While Not EOF(10)
    Line Input #10, tstr
    Index = InStr(tstr, "#")
    If Index > 0 Then
        If ItemCtrl(0).Name <> "" Then ReDim Preserve ItemCtrl(UBound(ItemCtrl) + 1)
        ItemCtrl(UBound(ItemCtrl)).Lock = True
        ItemCtrl(UBound(ItemCtrl)).Reject = True
    Else
        Index = InStr(tstr, "=")
        If Index > 0 Then
            Select Case LCase(Trim(Left(tstr, Index - 1)))
                Case "name"
                    ItemCtrl(UBound(ItemCtrl)).Name = LCase(Trim(Right(tstr, Len(tstr) - Index)))
                Case "price"
                    ItemCtrl(UBound(ItemCtrl)).Price = Val(Trim(Right(tstr, Len(tstr) - Index)))
                Case "lock"
                    ItemCtrl(UBound(ItemCtrl)).Lock = CBool(Val(Trim(Right(tstr, Len(tstr) - Index))))
                Case "reject"
                    ItemCtrl(UBound(ItemCtrl)).Reject = CBool(Val(Trim(Right(tstr, Len(tstr) - Index))))
            End Select
        End If
    End If
Loop
Close 10
Exit Sub
errie:
Close 10
'ReDim ItemCtrl(0)
MsgBox "Error!!! on loading 'control\items.txt'", vbCritical
End Sub
Sub ReadEventList()
On Error GoTo errie
    Dim tstr$, Index&, tmpstr$, spl() As String
    ReDim Events(0)
    ReDim UserVar(0)
    Events(0).Enabled = True
    isUseEvents = True
    Open App.Path & "\profile\events.txt" For Input As #9
        Do Until EOF(9)
            Line Input #9, tstr
            tstr = Trim(tstr)
            If Mid(tstr, 1, 1) = "'" Then GoTo nextcheck
            Index = InStr(tstr, "#")
            If Index = 1 Then
                If Events(UBound(Events)).Name <> "" Then ReDim Preserve Events(UBound(Events) + 1)
                Events(UBound(Events)).Chance = 100
                Events(UBound(Events)).Enabled = True
            Else
                Index = InStr(tstr, "=")
                If Index > 0 Then
                    Select Case LCase(Trim(Left(tstr, Index - 1)))
                        Case "event"
                            Events(UBound(Events)).Name = Trim(Right(tstr, Len(tstr) - Index))
                        Case "pre-process_var"
                            If Len(Events(UBound(Events)).PreVar) > 0 Then Events(UBound(Events)).PreVar = Events(UBound(Events)).PreVar & Chr(0) & Trim(Right(tstr, Len(tstr) - Index)) Else Events(UBound(Events)).PreVar = Trim(Right(tstr, Len(tstr) - Index))
                        Case "check"
                            If Len(Events(UBound(Events)).Check) > 0 Then Events(UBound(Events)).Check = Events(UBound(Events)).Check & Chr(0) & Trim(Right(tstr, Len(tstr) - Index)) Else Events(UBound(Events)).Check = Trim(Right(tstr, Len(tstr) - Index))
                        Case "action"
                            If Len(Events(UBound(Events)).Action) > 0 Then Events(UBound(Events)).Action = Events(UBound(Events)).Action & Chr(0) & Trim(Right(tstr, Len(tstr) - Index)) Else Events(UBound(Events)).Action = Trim(Right(tstr, Len(tstr) - Index))
                        Case "post-process_var"
                            If Len(Events(UBound(Events)).PostVar) > 0 Then Events(UBound(Events)).PostVar = Events(UBound(Events)).PostVar & Chr(0) & Trim(Right(tstr, Len(tstr) - Index)) Else Events(UBound(Events)).PostVar = Trim(Right(tstr, Len(tstr) - Index))
                        Case "chance"
                            tmpstr = LCase(Trim(Right(tstr, Len(tstr) - Index)))
                            If Right(tmpstr, 1) = "%" Then tmpstr = Left(tmpstr, Len(tmpstr) - 1)
                            If Val(tmpstr) < 0 Or Val(tmpstr) > 100 Then tmpstr = "100"
                            Events(UBound(Events)).Chance = Val(tmpstr)
                        'Case "use_event_profile"
                            'isUseEvents = CBool(Val(Trim(Right(tstr, Len(tstr) - Index))))
                        Case "user-define_variable"
                            spl = Split(Trim(Right(tstr, Len(tstr) - Index)), "=", 2)
                            UserVar(UBound(UserVar)).Variable = spl(0)
                            If UBound(spl) > 0 Then UserVar(UBound(UserVar)).value = spl(1)
                            ReDim Preserve UserVar(UBound(UserVar) + 1)
                    End Select
                End If
            End If
nextcheck:
        Loop
    Close #9
    'ReDim Preserve UserVar(UBound(UserVar) - 1)
    SaveEvents
    Exit Sub
errie:
    MsgBox "Error on loading 'profile\events.txt'"
    End
End Sub
Sub SaveEvents()
On Error GoTo errie
    Dim i&, j&
    Dim spl() As String
    Open App.Path & "\profile\events.txt" For Output As #42
    Print #42, "' This files was auto-regenerated by " & Version & " to arrange some syntax."
    Print #42, "' Note that you can't disable event anymore!!"
    Print #42, "use_event_profile = 1"
    For i = 0 To UBound(UserVar) - 1
        Print #42, "user-define_variable = " & UserVar(i).Variable & "=" & UserVar(i).value
    Next
    For i = 0 To UBound(Events)
        Print #42, "# event id:" & CStr(i)
        If Len(Events(i).Name) > 0 Then
            Print #42, "event = " & Events(i).Name
            spl = Split(Events(i).PreVar, Chr(0))
            For j = 0 To UBound(spl)
                Print #42, "pre-process_var = " & spl(j)
            Next
            spl = Split(Events(i).Check, Chr(0))
            For j = 0 To UBound(spl)
                Print #42, "check = " & spl(j)
            Next
            spl = Split(Events(i).Action, Chr(0))
            For j = 0 To UBound(spl)
                Print #42, "action = " & spl(j)
            Next
            spl = Split(Events(i).PostVar, Chr(0))
            For j = 0 To UBound(spl)
                Print #42, "post-process_var = " & spl(j)
            Next
            Print #42, "chance = " & Events(i).Chance & "%"
        End If
    Next
    Close #42
    Exit Sub
errie:
    MsgBox "Error on saving 'profile\events.txt' : " & Err.Description
End Sub
Function ReadLenTable(FileName As String) As Boolean
    If FileExists(FileName) Then
        Open FileName For Binary Access Read As #1
            PacketLenTable = String$(LOF(1), Chr$(0))
            Get #1, , PacketLenTable
        Close #1
        ReadLenTable = True
    End If
End Function

Sub print_funcerr(Func As String, number As Long, Description As String, Optional Additional As String)
    Open App.Path & "\log\functionlog.txt" For Append As #33
        Print #33, "== Error in : " & Func & " =="
        Print #33, "Time : " & Now
        Print #33, "Code : " & number
        Print #33, "Description : " & Description
        If Len(Additional) > 0 Then Print #33, "Debug Information : " & Additional
        Print #33, ""
    Close #33
End Sub

Sub Load_Sell()
On Error GoTo errie
Dim tstr As String
ReDim SelItem(0)
Open App.Path & "\control\Sell.txt" For Input As #1
Do While Not EOF(1)
    Input #1, tstr
    If Len(tstr) > 1 Then
         If Left(tstr, 1) <> "/" Then
            SelItem(UBound(SelItem)).Name = Right(tstr, Len(tstr))
            ReDim Preserve SelItem(UBound(SelItem) + 1)
        End If
    End If
Loop
'ReDim Preserve SkillList(UBound(SkillList) - 1)
Close 1
Exit Sub
errie:
    MsgBox "Error!!! on loading 'control\sell.txt'", vbCritical
Err.Clear
End Sub

Sub Load_Buy()
On Error GoTo errie
Dim tstr As String
Dim Index As Integer
Dim index2 As Integer
ReDim BuyItem(0)
Open App.Path & "\control\buy.txt" For Input As #1
Do While Not EOF(1)
    Line Input #1, tstr
    Index = InStr(tstr, "=")
    index2 = InStr(tstr, "/")
    If Index > 0 Then
        BuyItem(UBound(BuyItem)).Name = LCase(Trim(Left(tstr, Index - 1)))
        BuyItem(UBound(BuyItem)).Amount = Trim(Mid(tstr, Index + 1, index2 - Index - 1))
        BuyItem(UBound(BuyItem)).BackNumber = Trim(Right(tstr, Len(tstr) - index2))
        ReDim Preserve BuyItem(UBound(BuyItem) + 1)
    End If
Loop
If UBound(BuyItem) > 0 Then ReDim Preserve BuyItem(UBound(BuyItem) - 1)
Close 1
Exit Sub
errie:
    Close 1
    MsgBox "Error on loading 'control\buy.txt'", vbCritical
End Sub

Sub Load_Kafra()
On Error GoTo errie
Dim tstr As String
Dim Index As Integer
ReDim Kafra(0)
Open App.Path & "\control\keep.txt" For Input As #1
Do While Not EOF(1)
    Line Input #1, tstr
    If Len(Trim(tstr)) > 1 Then
        tstr = Trim(tstr)
         If Left(tstr, 1) <> "/" Then
            Index = InStr(tstr, "=")
            If Index > 0 Then
                Kafra(UBound(Kafra)).Amount = Val(Trim(Right(tstr, Len(tstr) - Index)))
                Kafra(UBound(Kafra)).Name = Trim(Left(tstr, Index - 1))
            Else
                Kafra(UBound(Kafra)).Name = tstr
            End If
            ReDim Preserve Kafra(UBound(Kafra) + 1)
        End If
    End If
Loop
If UBound(Kafra) > 0 Then ReDim Preserve Kafra(UBound(Kafra) - 1)
Close 1
Exit Sub
errie:
    Close 1
    MsgBox "Error on loading 'control\keep.txt'", vbCritical
End Sub


Sub Load_GetStorage()
On Error GoTo errie
Dim tstr As String
Dim Index As Integer
Dim index2 As Integer
ReDim GetStorageItem(0)
Open App.Path & "\control\storage.txt" For Input As #1
Do While Not EOF(1)
    Input #1, tstr
    Index = InStr(tstr, "=")
    index2 = InStr(tstr, "/")
    If Index > 0 Then
        GetStorageItem(UBound(GetStorageItem)).Name = Trim(Left(tstr, Index - 1))
        GetStorageItem(UBound(GetStorageItem)).Amount = Trim(Mid(tstr, Index + 1, index2 - Index - 1))
        GetStorageItem(UBound(GetStorageItem)).BackNumber = Trim(Right(tstr, Len(tstr) - index2))
        ReDim Preserve GetStorageItem(UBound(GetStorageItem) + 1)
    End If
Loop
If UBound(GetStorageItem) > 0 Then ReDim Preserve GetStorageItem(UBound(GetStorageItem) - 1)
Close 1
Exit Sub
errie:
    Close 1
    MsgBox "Error on loading 'control\storage.txt'", vbCritical
End Sub

Sub Load_LockmapList()
On Error GoTo errie
    Dim tstr As String
    ReDim LockmapList(0)
    Open App.Path & "\profile\script_lockmap.txt" For Input As #1
        Do Until EOF(1)
            Input #1, tstr
            If Left(tstr, 1) = "#" Then
                If Len(LockmapList(0).MapName) > 0 Then ReDim Preserve LockmapList(UBound(LockmapList) + 1)
            End If
            If InStr(tstr, "=") > 0 Then
                Select Case LCase(Trim(Left(tstr, InStr(tstr, "="))))
                    Case "level_min"
                        If Val(Trim(Right(tstr, Len(tstr) - InStr(tstr, "=")))) < 255 Then LockmapList(UBound(LockmapList)).LvMin = Val(Trim(Right(tstr, Len(tstr) - InStr(tstr, "="))))
                    Case "level_max"
                        If Val(Trim(Right(tstr, Len(tstr) - InStr(tstr, "=")))) < 255 Then LockmapList(UBound(LockmapList)).LvMax = Val(Trim(Right(tstr, Len(tstr) - InStr(tstr, "="))))
                    Case "lockmapname"
                        LockmapList(UBound(LockmapList)).MapName = LCase(Trim(Right(tstr, Len(tstr) - InStr(tstr, "="))))
                End Select
            End If
        Loop
    Close #1
    Exit Sub
errie:
    Close 1
    MsgBox "Error on loading 'profile\script_lockmap.txt' : " & Err.Description, vbCritical
End Sub

