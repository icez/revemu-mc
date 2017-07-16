Attribute VB_Name = "md_Events"
Option Explicit

Public isUseEvents As Boolean
Public ActionList As String
Public Events() As EventList
Public UserVar() As EventProcessing
Type EventList
    Name As String
    Check As String
    Action As String
    Chance As Byte
    Enabled As Boolean
    PreVar As String
    PostVar As String
End Type
Type EventProcessing
    Variable As String
    value As String
End Type

Sub CheckEvent(Name As String, value As String)
On Error GoTo errie
    If Not isUseEvents Then Exit Sub
    Dim i&, j&, K&, str1() As String, str2() As String
    'prepare local variable
    Dim PE() As EventProcessing
    str1 = Split(WithPubDeclare(value), Chr(0))
    ReDim PE(UBound(str1))
    On Error Resume Next
    For i = 0 To UBound(str1)
        str2 = Split(str1(i), "=")
        PE(i).Variable = str2(0)
        If InStr(str1(i), "=") Then PE(i).value = str2(1) Else PE(i).value = "True"
    Next
    Err.Clear
    On Error GoTo errie
    '--'
        Dim tstr$
    For i = 0 To UBound(Events)
        If Not Events(i).Enabled Then GoTo nextcheck
        If Events(i).Name = "" Then GoTo nextcheck
        If Events(i).Name <> Name Then GoTo nextcheck
        If Not RandomUse(Events(i).Chance) Then GoTo nextcheck
        
        EProcess Events(i).PreVar, PE
        str2 = Split(Events(i).Check, Chr(0))
        For j = 0 To UBound(str2)
            If Mods.STDebug Then Chat "Events : [Debug] - Checking variable [" & str2(j) & "] " & CStr(CheckEventCase(PE, str2(j))) & CStr(CheckEventCase(UserVar, str2(j)))
            If (Not CheckEventCase(PE, str2(j))) And (Not CheckEventCase(UserVar, str2(j))) Then GoTo nextcheck
        Next
        EProcess Events(i).PostVar, PE

        Dim tmpAction$
        tmpAction = ReplaceVar(PE, Events(i).Action)
        tmpAction = ReplaceVar(UserVar, tmpAction)
        tmpAction = Replace(tmpAction, "disablethisevent", "disableevent:" & i)
        tmpAction = Replace(tmpAction, "enablethisevent", "disableevent:" & i)
        If Len(ActionList) = 0 Then
            ActionList = tmpAction
        Else
            Do Until Len(ActionList) = 0
                Sleep 1
                DoEvents
            Loop
            ActionList = tmpAction
        End If
nextcheck:
    Next
    '--'
    Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "CheckEvent", Err.number, Err.Description
    Err.Clear
End Sub
Function CanFindVar(VarName As String) As Long
On Error GoTo errie
    Dim i&
    For i = 0 To UBound(UserVar)
        If UserVar(i).Variable = VarName Then
            CanFindVar = i
            Exit Function
        End If
    Next
errie:
    If Err.number > 0 Then print_funcerr "CanFindVar", Err.number, Err.Description
    Err.Clear
    CanFindVar = -1
End Function
Sub EProcess(EList As String, ProcE() As EventProcessing)
    On Error GoTo errie
    Dim spl() As String, spm() As String, spn() As String, spo() As String
    Dim i&, vpos&, tmps$, tmpl&
    spl = Split(EList, Chr(0))
    For i = 0 To UBound(spl)
        spm = Split(spl(i), "=", 2)
        vpos = CanFindVar(spm(0))
        If vpos < 0 Then
            Chat "Events : [Error] - Variable not defined : " & spm(0), vbRed
            GoTo donext
        End If
        spn = Split(spm(1), ":", 2)
        If UBound(spn) < 1 Then Chat "Events : [Error] - incorrect syntax of function : " & spn(0), vbRed: GoTo donext
        Select Case spn(0)
            Case "set" '+
                spn(1) = ReplaceVar(ProcE, spn(1))
                spn(1) = ReplaceVar(UserVar, spn(1))
                UserVar(vpos).value = spn(1)

            Case "plus" '+
                spn(1) = ReplaceVar(ProcE, spn(1))
                spn(1) = ReplaceVar(UserVar, spn(1))
                spo = Split(spn(1), ",", 2)
                If UBound(spo) < 1 Then Chat "Events : [Error] - incorrect syntax of function 'plus' near : " & spn(1), vbRed: GoTo donext
                UserVar(vpos).value = CStr(Val(spo(0)) + Val(spo(1)))
                
            Case "minus" '-
                spn(1) = ReplaceVar(ProcE, spn(1))
                spn(1) = ReplaceVar(UserVar, spn(1))
                spo = Split(spn(1), ",", 2)
                If UBound(spo) < 1 Then Chat "Events : [Error] - incorrect syntax of function 'minus' near : " & spn(1), vbRed: GoTo donext
                UserVar(vpos).value = CStr(Val(spo(0)) - Val(spo(1)))
                
            Case "mod" '-
                spn(1) = ReplaceVar(ProcE, spn(1))
                spn(1) = ReplaceVar(UserVar, spn(1))
                spo = Split(spn(1), ",", 2)
                If UBound(spo) < 1 Then Chat "Events : [Error] - incorrect syntax of function 'mod' near : " & spn(1), vbRed: GoTo donext
                UserVar(vpos).value = CStr(Val(spo(0)) Mod Val(spo(1)))
                
            Case "multiply" '*
                spn(1) = ReplaceVar(ProcE, spn(1))
                spn(1) = ReplaceVar(UserVar, spn(1))
                spo = Split(spn(1), ",", 2)
                If UBound(spo) < 1 Then Chat "Events : [Error] - incorrect syntax of function 'multiply' near : " & spn(1), vbRed: GoTo donext
                UserVar(vpos).value = CStr(Val(spo(0)) * Val(spo(1)))
                
            Case "divide" '\
                spn(1) = ReplaceVar(ProcE, spn(1))
                spn(1) = ReplaceVar(UserVar, spn(1))
                spo = Split(spn(1), ",", 2)
                If UBound(spo) < 1 Then Chat "Events : [Error] - incorrect syntax of function 'divide' near : " & spn(1), vbRed: GoTo donext
                UserVar(vpos).value = CStr(Val(spo(0)) \ Val(spo(1)))
                
            Case "itemcount" 'get currrent amount of item in inventory
                spn(1) = ReplaceVar(ProcE, spn(1))
                spn(1) = ReplaceVar(UserVar, spn(1))
                tmpl = Find_Item(spn(1))
                If tmpl < 1 Then
                    UserVar(vpos).value = "0"
                Else
                    UserVar(vpos).value = AllInv(tmpl).Amount
                End If

            Case "cartitemcount" 'get current amount of item in cart
                spn(1) = ReplaceVar(ProcE, spn(1))
                spn(1) = ReplaceVar(UserVar, spn(1))
                tmpl = Find_CartID(spn(1))
                If tmpl < 0 Then
                    UserVar(vpos).value = "0"
                Else
                    UserVar(vpos).value = Cart(tmpl).Amount
                End If
                
            Case "statusactive" 'is skill active
                spn(1) = ReplaceVar(ProcE, spn(1))
                spn(1) = ReplaceVar(UserVar, spn(1))
                If Val(spn(1)) > UBound(CurStatus) Then
                    UserVar(vpos).value = "False"
                Else
                    UserVar(vpos).value = CStr(CurStatus(Val(spn(1))).Active)
                End If

            Case Else
                 Chat "Events : [Error] - undefined function : " & spn(0), vbRed
                 GoTo donext
        End Select
donext:
    Next
errie:
    If Err.number > 0 Then print_funcerr "EProcess", Err.number, Err.Description
    Err.Clear
End Sub

Function Find_EventID(eCase As String) As Long
On Error GoTo errie
    If Not isUseEvents Then GoTo res_fail
    Dim i&
    For i = 0 To UBound(Events)
        If Events(i).Name = eCase Then
            Find_EventID = i
            Exit Function
        End If
    Next
res_fail:
    Find_EventID = -1
    Exit Function
errie:
    Err.Clear
    Find_EventID = -1
End Function

Function ReplaceVar(ProcE() As EventProcessing, LAction As String) As String
    Dim i&, tmpRes$
    tmpRes = LAction
    For i = 0 To UBound(ProcE)
        tmpRes = Replace(tmpRes, "$" & ProcE(i).Variable, ProcE(i).value)
    Next
    ReplaceVar = tmpRes
End Function

Function WithPubDeclare(value As String) As String
On Error GoTo errie
    Dim tstr$, testbound&
    On Error Resume Next
        testbound = UBound(Players)
        If Err.number > 0 Then testbound = 1 Else testbound = 0
    On Error GoTo errie
    tstr = value & Chr(0) & "isinlock=" & CStr(IsInLock)
    tstr = tstr & Chr(0) & "ischatopen=" & CStr(IsChatOC)
    tstr = tstr & Chr(0) & "connstate=" & CStr(ConnState)
    If testbound = 1 Then
        tstr = tstr & Chr(0) & "curWeight=0"
        tstr = tstr & Chr(0) & "maxWeight=0"
        tstr = tstr & Chr(0) & "percentWeight=0"
        tstr = tstr & Chr(0) & "curPlayerHP=0"
        tstr = tstr & Chr(0) & "maxPlayerHP=0"
        tstr = tstr & Chr(0) & "percentPlayerHP=0"
        tstr = tstr & Chr(0) & "curPlayerSP=0"
        tstr = tstr & Chr(0) & "maxPlayerSP=0"
        tstr = tstr & Chr(0) & "percentPlayerSP=0"
    Else
        tstr = tstr & Chr(0) & "curWeight=" & CInt(Players(number).Weight)
        tstr = tstr & Chr(0) & "maxWeight=" & CInt(Players(number).MaxWeight)
        tstr = tstr & Chr(0) & "percentWeight=" & Int(GetWeight * 100)
        tstr = tstr & Chr(0) & "curPlayerHP=" & Players(number).HP
        tstr = tstr & Chr(0) & "maxPlayerHP=" & Players(number).MaxHP
        tstr = tstr & Chr(0) & "percentPlayerHP=" & Int(GetHP * 100)
        tstr = tstr & Chr(0) & "curPlayerSP=" & Players(number).SP
        tstr = tstr & Chr(0) & "maxPlayerSP=" & Players(number).maxsp
        tstr = tstr & Chr(0) & "percentPlayerSP=" & Int(GetSP * 100)
    End If
    tstr = tstr & Chr(0) & "isvending=" & CStr(IsVending)
    tstr = tstr & Chr(0) & "isinfight=" & CStr(MakeDamage)
    tstr = tstr & Chr(0) & "playerX=" & CStr(curPos.Y)
    tstr = tstr & Chr(0) & "playerY=" & CStr(curPos.X)
    tstr = tstr & Chr(0) & "playerAID=" & CStr(MakePort(AccountID))
    tstr = tstr & Chr(0) & "mapname=" & CStr(MapName)

    WithPubDeclare = tstr
    Exit Function
errie:
    If Err.number > 0 Then print_funcerr "WithPubDeclare", Err.number, Err.Description
    WithPubDeclare = value
    Err.Clear
End Function

Function CheckEventCase(ProcE() As EventProcessing, strChk As String) As Boolean
On Error GoTo errie
    Dim Index&, i&, chkMode As Byte, cCase$, cVal$
    Index = InStr(strChk, "=")
    chkMode = 1 'equal
    If Index > 0 Then GoTo loopchk
    
    Index = InStr(strChk, "<>")
    chkMode = 2 'not equal
    If Index > 0 Then GoTo loopchk
    
    Index = InStr(strChk, ">")
    chkMode = 3 'more than
    If Index > 0 Then GoTo loopchk
    
    Index = InStr(strChk, "<")
    chkMode = 4 'less than
    If Index > 0 Then GoTo loopchk
    
    Index = InStr(strChk, "@")
    chkMode = 5 'like
    If Index > 0 Then GoTo loopchk
    
    Index = InStr(strChk, "\")
    chkMode = 6 'not like
    If Index > 0 Then GoTo loopchk
    
    Index = InStr(strChk, "!")
    chkMode = 7 'boolean false
    If Index > 0 Then GoTo loopchk
    
    Index = 0
    chkMode = 8 'boolean true
    
loopchk:
    'MsgBox chkMode
    If Index > 0 Then cCase = Mid(strChk, 1, Index - 1) Else cCase = ""
    cVal = Right(strChk, Len(strChk) - Index - IIf(chkMode = 2, 1, 0))
    
    If chkMode > 6 Then
        If Index = 0 Then cCase = strChk Else cCase = Mid(strChk, 2, Len(strChk) - 1)
        If chkMode = 7 Then cVal = "False" Else cVal = "True"
    End If
    
    'MsgBox cCase & " ::: " & cVal
    For i = 0 To UBound(ProcE)
        If ProcE(i).Variable <> cCase And cCase <> "" Then GoTo nextchk
        If cCase = "" And chkMode < 7 Then GoTo nextchk
        Select Case chkMode
            Case 1 'equal
                Dim chkspl() As String, j&
                chkspl = Split(cVal, "/")
                For j = 0 To UBound(chkspl)
                    If ProcE(i).value = chkspl(j) Then
                        CheckEventCase = True
                        Exit Function
                    End If
                Next
            Case 2 'not equal
                If ProcE(i).value <> cVal Then
                    CheckEventCase = True
                    Exit Function
                End If
            Case 3 'more than
                If IsNumeric(ProcE(i).value) And IsNumeric(cVal) Then
                    If Val(ProcE(i).value) > Val(cVal) Then
                        CheckEventCase = True
                        Exit Function
                    End If
                End If
            Case 4 'less then
                If IsNumeric(ProcE(i).value) And IsNumeric(cVal) Then
                    If Val(ProcE(i).value) < Val(cVal) Then
                        CheckEventCase = True
                        Exit Function
                    End If
                End If
            Case 5 'like
                If MODCheckStr(ProcE(i).value, cVal) Then
                    CheckEventCase = True
                    Exit Function
                End If
            Case 6 'not like
                If Not MODCheckStr(ProcE(i).value, cVal) Then
                    CheckEventCase = True
                    Exit Function
                End If
            Case 7, 8
                If ProcE(i).value = cVal Then
                    CheckEventCase = True
                    Exit Function
                End If
            Case Else
                Err.Raise 0, "", "Invalid case number " & chkMode
        End Select
nextchk:
    Next
    CheckEventCase = False
    Exit Function
errie:
    print_funcerr "CheckEventCase", Err.number, Err.Description
    Err.Clear
End Function

Function RandomUse(Chance As Byte) As Boolean
    Dim rndval As Integer
    Randomize
    rndval = Int(Rnd() * 100)
    If rndval < Chance Then RandomUse = True Else RandomUse = False
End Function

Sub ProcessAction()
On Error GoTo errie
    If Len(ActionList) > 0 Then
        Dim tmpAct$, mDist&, rndWaitTime&
        tmpAct = ActionList
        ActionList = ""
        Dim spl() As String, i&, spl2() As String, spl3() As String, selmsg As Long
        Dim X As Long
        Dim mCoord As Coord
        spl = Split(tmpAct, Chr(0))
        For i = 0 To UBound(spl)
            spl2 = Split(spl(i), ":", 2)
            If Mods.STDebug Then
                If UBound(spl2) > 0 Then Chat "Debug : [Events] - action list processing : " & spl2(0) & " - " & spl2(1), &HAAAAAA Else Chat "Event action list processing : " & spl2(0), &HAAAAAA
            End If
            'spl2(0) is an action name
            'spl2(1) is an action value
            Select Case spl2(0)
                Case "delay"
                    If InStr(spl2(1), ",") > 0 Then
                        spl3 = Split(spl2(1), ",", 3)
                        If Val(spl3(0)) > 0 And Val(spl3(1)) > Val(spl3(0)) Then
                            rndWaitTime = RandomVal(Val(spl3(0)), Val(spl3(1)))
                            If Mods.STDebug Then Chat "Debug : [Events] - action 'delay' waiting for " & rndWaitTime & " milliseconds"
                            WaitTime rndWaitTime
                        Else
                            If Val(spl3(0)) > 0 Then
                                If Mods.STDebug Then Chat "Debug : [Events] - action 'delay' waiting for " & Val(spl3(0)) & " milliseconds"
                                WaitTime Val(spl3(0))
                            Else
                                If Mods.STDebug Then Chat "Debug : [Events] - action 'delay' variable passed into this function error!"
                            End If
                        End If
                    Else
                        If Val(spl2(1)) > 0 Then
                            If Mods.STDebug Then Chat "Debug : [Events] - action 'delay' waiting for " & Val(spl2(1)) & " milliseconds"
                            WaitTime Val(spl2(1))
                        Else
                            If Mods.STDebug Then Chat "Debug : [Events] - action 'delay' variable passed into this function error!!"
                        End If
                    End If

                Case "say"
                    spl3 = Split(spl2(1), ",")
                    selmsg = RandomVal(0, UBound(spl3))
                    SendPUBChat spl3(selmsg)

                Case "useskill"
                    spl3 = Split(spl2(1), ",", 4)
                    X = Find_SkillId(UCase(spl3(1)))
                    If X > 0 Then Send_Use_Skill SkillChar(X).ID, CByte(Val(spl3(2))), LngToChr(Val(spl3(0)))

                Case "savelog"
                    Dim FFile As Long
                    FFile = FreeFile
                    spl3 = Split(spl2(1), ",", 2)
                    Open App.Path & "\log\" & spl3(0) For Append As FFile
                        Print #FFile, spl3(1)
                    Close FFile

                Case "chatmsg"
                    Chat spl2(1)

                Case "msgbox"
                    MsgBox spl2(1)

                Case "sit"
                    frmMain.Send_Sit
                    IsSitting = True
                    IsStanding = False
                
                Case "stand"
                    frmMain.Send_Stand
                    IsSitting = False
                    IsStanding = True
                
                Case "equipitem"
                    X = Find_Item(Trim(spl2(1)))
                    If X > -1 Then
                        With AllInv(X)
                            If (((.Category > 3 And .Category < 6) Or .Category > 7) And .Category <> 10 And .Pos = 0 And .Name <> "") Then
                                Winsock_SendPacket IntToChr(&HA9) & IntToChr(X) & .Type, True
                                Stat "Events : Equipping [" & .Name & "]", vbBlue
                            End If
                        End With
                    End If
                
                Case "unequipitem"
                    X = Find_Item(Trim(spl2(1)))
                    If X > -1 Then
                        If AllInv(X).Pos > 0 Then
                            Winsock_SendPacket IntToChr(&HAB) & IntToChr(X), True
                            Chat "Events : Un-Equipping [" & AllInv(X).Name & "]", vbBlue
                        End If
                    End If
                
                Case "pm"
                    spl3 = Split(spl2(1), "|", 2)
                    Dim tNick$, tMsg$
                    tNick = spl3(0)
                    tMsg = spl3(1)
                    spl3 = Split(tMsg, ",")
                    selmsg = RandomVal(0, UBound(spl3))
                    SendPRIChat tNick, spl3(selmsg)

                Case "guildmsg"
                    spl3 = Split(spl2(1), ",")
                    selmsg = RandomVal(0, UBound(spl3))
                    SendGUIChat spl3(selmsg)

                Case "partymsg"
                    spl3 = Split(spl2(1), ",")
                    selmsg = RandomVal(0, UBound(spl3))
                    SendPARChat spl3(selmsg)

                Case "disableevent"
                    Events(Val(spl2(1))).Enabled = False

                Case "enableevent"
                    Events(Val(spl2(1))).Enabled = False
                
                Case "warpsave"
                    frmMain.Warp_Save "Event script : teleport back to town."

                Case "createarrow"
                    If UBound(spl2) > 0 Then pkt_CreateArrow spl2(1)

                Case "emotion"
                    spl3 = Split(spl2(1), ",")
                    selmsg = RandomVal(0, UBound(spl3))
                    frmMain.Send_Emoticon Get_Emotion_Code(spl3(selmsg))

                Case "teleport"
                    Teleport

                Case "reconnect"
                    frmMain.ResettoReCon

                Case "disableai"
                    AutoAI = False
                    FrmField.update_ImgAI

                Case "enableai"
                    AutoAI = True
                    FrmField.update_ImgAI

                Case "disconnect"
                    If UBound(spl2) > 0 Then
                        frmMain.Winsock1.Close
                        MODDelay.DualLogin = Val(spl2(1)) * 6000
                        ConnState = 1
                    Else
                        ForceExit
                    End If
                
                Case "terminate"
                    ForceExit

                Case "moveto"
                    spl3 = Split(spl2(1), ",", 2)
                    mCoord.Y = Val(spl3(0))
                    mCoord.X = Val(spl3(1))
                    move_to mCoord

                Case "nearmoveto"
                    spl3 = Split(spl2(1), "|", 2)
                    mDist = Val(spl3(1))
                    spl3 = Split(spl3(0), ",", 2)
                    mCoord.Y = Val(spl3(0))
                    mCoord.X = Val(spl3(1))
                    For X = 1 To mDist
                        mCoord = NextPos(mCoord, curPos)
                    Next
                    move_to mCoord

                Case "route_moveto"
                    AutoAI = True
                    spl3 = Split(spl2(1), "|", 2)
                    mDist = Val(spl3(1))
                    spl3 = Split(spl3(0), ",", 2)
                    mCoord.Y = Val(spl3(0))
                    mCoord.X = Val(spl3(1))
'                    For X = 1 To mDist
'                        mCoord = NextPos(mCoord, CurPos)
'                    Next
                    ptEnd.X = mCoord.X
                    ptEnd.Y = mCoord.Y
                    FrmField.Run_Search

                Case "useitem"
                    X = Find_HealItem(spl2(1))
                    If X > 0 Then Winsock_SendPacket IntToChr(&HA7) & IntToChr(X) & AccountID, True

                Case "dropitem"
                    spl3 = Split(spl2(1), ",")
                    X = Find_Item(spl3(0))
                    'S 00a2 <index>.w <amount>.w
                    If UBound(spl3) = 0 And X > 0 Then
                        Winsock_SendPacket IntToChr(&HA2) & IntToChr(CLng(AllInv(X).Index)) & IntToChr(AllInv(X).Amount), True
                    ElseIf UBound(spl3) > 0 And X > 0 Then
                        If AllInv(X).Amount > Val(spl3(1)) Then
                            Winsock_SendPacket IntToChr(&HA2) & IntToChr(CLng(AllInv(X).Index)) & IntToChr(Val(spl3(1))), True
                        Else
                            Winsock_SendPacket IntToChr(&HA2) & IntToChr(CLng(AllInv(X).Index)) & IntToChr(AllInv(X).Amount), True
                        End If
                    End If

                Case Else
                    Chat "Events : [Error] - Undefined action : " & spl2(0)
            End Select
        Next
    End If
    Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "ProcessAction", Err.number, Err.Description
    Err.Clear
End Sub
