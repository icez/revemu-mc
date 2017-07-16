Attribute VB_Name = "md_ChatResp"
Option Explicit

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal Milliseconds As Long)

Public ChatResp() As MODChatResp
Public UseResp As Boolean

Public MChatDelay As Long
Public MChatRestype As String
Public MChatRest As Byte
Public MChatResNick As String

Type MODChatResp
    rangeMin As Long
    rangeMax As Long
    Delay As Long
    delayrand As Long
    inText As String
    OutText As String
    doTele As Boolean
    onPriv As Boolean
    onPub As Boolean
    onGuild As Boolean
    onParty As Boolean
    percent As Long
End Type

Sub ReadChatResponse()
On Error GoTo errie
    UseResp = False
    Dim lfile As Long, tstr$, index&
    lfile = FreeFile
    ReDim ChatResp(0)
    Open App.Path & "\profile\chat_response.txt" For Input As lfile
        Do Until EOF(lfile)
            Line Input #lfile, tstr
            index = InStr(tstr, "=")
            If Mid(Trim(tstr), 1, 1) = "'" Then index = 0
            If tstr = "#" Then
                ReDim Preserve ChatResp(UBound(ChatResp) + 1)
            Else
                If index > 0 Then
                    Select Case LCase(Trim(Left(tstr, index - 1)))
                        Case "mindist"
                            ChatResp(UBound(ChatResp)).rangeMin = Val(Trim(Right(tstr, Len(tstr) - index)))
                        Case "maxdist"
                            ChatResp(UBound(ChatResp)).rangeMax = Val(Trim(Right(tstr, Len(tstr) - index)))
                        Case "onprivate"
                            ChatResp(UBound(ChatResp)).onPriv = CBool(Val(Trim(Right(tstr, Len(tstr) - index))))
                        Case "onpublic"
                            ChatResp(UBound(ChatResp)).onPub = CBool(Val(Trim(Right(tstr, Len(tstr) - index))))
                        Case "onparty"
                            ChatResp(UBound(ChatResp)).onParty = CBool(Val(Trim(Right(tstr, Len(tstr) - index))))
                        Case "onguild"
                            ChatResp(UBound(ChatResp)).onGuild = CBool(Val(Trim(Right(tstr, Len(tstr) - index))))
                        Case "teleport"
                            ChatResp(UBound(ChatResp)).doTele = CBool(Val(Trim(Right(tstr, Len(tstr) - index))))
                        Case "intext"
                            ChatResp(UBound(ChatResp)).inText = LCase(Trim(Right(tstr, Len(tstr) - index - 1)))
                        Case "outtext"
                            ChatResp(UBound(ChatResp)).OutText = ChatResp(UBound(ChatResp)).OutText & "|" & Trim(Right(tstr, Len(tstr) - index))
                        Case "delay"
                            ChatResp(UBound(ChatResp)).Delay = Val(Trim(Right(tstr, Len(tstr) - index)))
                        Case "delayrand"
                            ChatResp(UBound(ChatResp)).delayrand = Val(Trim(Right(tstr, Len(tstr) - index)))
                        Case "percent"
                            ChatResp(UBound(ChatResp)).percent = Val(Trim(Right(tstr, Len(tstr) - index)))
                        Case "use_chat_response"
                            UseResp = CBool(Val(Trim(Right(tstr, Len(tstr) - index))))
                    End Select
                End If
            End If
        Loop
    Close lfile
    Exit Sub
errie:
    Close lfile
    MsgBox "Error on loading 'profile\chat_response.txt' : " & Err.Description
    ReDim ChatResp(0)
    UseResp = False
End Sub

Sub CheckChatResponse(ChatText As String, InType As Byte, InDist As Long)
    If Not UseResp Then Exit Sub
    Dim i&
    If MChatDelay = 0 Then MChatRestype = ""
    MChatRest = InType
    Dim CNick$, CText$, Choosen$, delays&, randPC&
    CNick = Left(ChatText, InStr(1, ChatText, " : "))
    If CNick = Players(number).Name Then Exit Sub
    MChatResNick = CNick
    CText = Right(ChatText, Len(ChatText) - InStr(1, ChatText, " : ") - 2)
    'Stat CText, 0, True, False, True
    For i = 0 To UBound(ChatResp)
        If MODCheckStr(CText, ChatResp(i).inText) Then
            Randomize
            MChatDelay = ((ChatResp(i).Delay * 100) + Int((Rnd() * (ChatResp(i).delayrand + 1)) * 100))
            If Not RandomUse(CByte(ChatResp(i).percent)) Then GoTo casenext
            Choosen = MODChooseText(ChatResp(i).OutText)
            If Len(Choosen) < 1 Then GoTo casenext
            Select Case InType
                Case 0 'pub
                    If Not ChatResp(i).onPub Then GoTo casenext
                    If InDist < ChatResp(i).rangeMin Or InDist > ChatResp(i).rangeMax Then GoTo casenext
                    MChatRestype = Choosen
                    If ChatResp(i).doTele Then
                       MChatRestype = "dotele"
                    End If
                Case 1 'priv
                    If Not ChatResp(i).onPriv Then GoTo casenext
                    MChatRestype = Choosen
                    If ChatResp(i).doTele Then
                       MChatRestype = "dotele"
                    End If
                Case 2 'party
                    If Not ChatResp(i).onParty Then GoTo casenext
                    MChatRestype = Choosen
                    If ChatResp(i).doTele Then
                       MChatRestype = "dotele"
                    End If
                Case 3 'guild
                    If Not ChatResp(i).onGuild Then GoTo casenext
                    MChatRestype = Choosen
                    If ChatResp(i).doTele Then
                       MChatRestype = "dotele"
                    End If
                Case Else: Err.Raise 1, "CheckChatResponse", "Unknown InType"
            End Select
            Exit For
        End If
casenext:
    Next
    Exit Sub
errie:
    MsgBox "Error in CheckChatResponse : " & Err.Description
    Exit Sub
End Sub

Function MODChooseText(inText As String)
    Dim tstr() As String
    tstr = Split(Mid(inText, 2, Len(inText) - 1), "|")
    Randomize
    MODChooseText = tstr(Int(Rnd() * (UBound(tstr) + 1)))
End Function

Function MODCheckStr(inString As String, CheckCase As String) As Boolean
On Error GoTo errie
    Dim inStr1$, inStr2$
    inStr1 = LCase(inString)
    inStr2 = LCase(CheckCase)
    If Len(inStr2) = 0 Or Len(inStr1) = 0 Then Exit Function
    Dim Srch1() As String, Srch2() As String
    Srch1 = Split(inStr2, ",")
    Dim i&, j&, iCheck&, jCheck&, resJ As Boolean
    iCheck = 0
    For i = 0 To UBound(Srch1)
        If Len(Srch1(i)) > 0 Then
            If InStr(1, Srch1(i), "/") Then
                Srch2 = Split(Srch1(i), "/")
                jCheck = 0
                resJ = False
                For j = 0 To UBound(Srch2)
                    If Len(Srch2(j)) > 0 Then
                        jCheck = InStr(iCheck + 1, inStr1, Srch2(j))
                        If jCheck > iCheck And iCheck > 0 Then
                            iCheck = jCheck
                            resJ = True
                            Exit For
                        End If
                    End If
                Next
                If Not resJ Then GoTo falsecheck
            Else
                If InStr(iCheck + 1, inStr1, Srch1(i)) = 0 Then GoTo falsecheck
                If InStr(iCheck + 1, inStr1, Srch1(i)) > iCheck Then
                    iCheck = InStr(iCheck + 1, inStr1, Srch1(i))
                End If
            End If
        End If
    Next
truecheck:
    MODCheckStr = True
    Exit Function
falsecheck:
    MODCheckStr = False
    Exit Function
errie:
    MODCheckStr = False
    MsgBox "Error in MODCheckStr : " & Err.Description
End Function
Sub WaitTime(timeMS As Long)
    Dim etime As Double
    etime = Timer + (timeMS / 1000)
    Do Until Timer >= etime
        'Sleep 1
        DoEvents
    Loop
End Sub
