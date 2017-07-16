Attribute VB_Name = "Mod_NPC_Profile"
Type NPC_Profile
    Cause As String
    Script As String
    location As String
    target As String
    pos As Coord
End Type

Public use_npc_profile As Boolean
Public ai_npc() As NPC_Profile
Public npc_step As String

Public npcwarp() As NPC_Profile
Public UseNPCWarp As Boolean

Public Function Check_sameNPCMap(Name As String, Cause As String, Optional chkStart As Long = 0) As Integer
    Dim i As Integer
    If chkStart > UBound(ai_npc) Then
        Check_sameNPCMap = -1
        Exit Function
    End If
    For i = chkStart To UBound(ai_npc)
        If ai_npc(i).location = MapName Then
            If InStr(Cause, ai_npc(i).Cause) > 0 Then
                'Dim dis As Integer
                'Dim tmpPos As Coord
                'tmpPos.X = ai_npc(i).pos.Y
                'tmpPos.Y = ai_npc(i).pos.X
                'dis = EvalNorm(CurPos, tmpPos)
                'If dis >= 10 Then
                    Check_sameNPCMap = i
                    Exit Function
                'Else
                '    Check_sameNPCMap = UBound(ai_npc) + 1
                '    Exit Function
                'End If
            End If
        End If
    Next
    Check_sameNPCMap = -1
End Function

Public Function Check_NPCMap(Name As String, Cause As String) As Integer
    Dim i As Integer
    For i = 0 To UBound(ai_npc)
       If InStr(Cause, ai_npc(i).Cause) > 0 And ai_npc(i).location <> "" Then
            Check_NPCMap = i
            Exit Function
        End If
    Next
    Check_NPCMap = -1
End Function

Public Function Get_NPCWarp_Choice(MapName As String, Index As Integer) As String
    Dim tstr As String, tstr2 As String
    Dim i As Integer
    tstr = npcwarp(Index).target
    i = 0
    Do While tstr <> ""
        Index = InStr(tstr, " ")
        i = i + 1
        If Index > 0 Then
            tstr2 = Trim(Left(tstr, Index))
            tstr = Trim(Right(tstr, Len(tstr) - Index))
        Else
            tstr2 = tstr
            tstr = ""
        End If
        If MapName = tstr2 Then
            Get_NPCWarp_Choice = "S" & CStr(i)
            Exit Function
        End If
    Loop
    Get_NPCWarp_Choice = ""
End Function

Public Sub Load_NPCWARP()
On Error GoTo errie:
Open App.Path & "\maproute\npcwarp.txt" For Input As #10
Dim tstr, tstr2 As String
Dim Index, i As Integer
Dim tmp As Integer
UseNPCWarp = False
ReDim npcwarp(0)
Do While Not EOF(10)
    Line Input #10, tstr
    tstr = Trim(tstr)
    Index = InStr(tstr, "#")
    If Index > 0 Then
        If npcwarp(0).Cause <> "" Then ReDim Preserve npcwarp(UBound(npcwarp) + 1)
    Else
        Index = InStr(tstr, "=")
        If Index > 0 Then
            Select Case LCase(Trim(Left(tstr, Index - 1)))
                Case "location"
                    tstr = Trim(Right(tstr, Len(tstr) - Index))
                    For i = 1 To 3
                        Index = InStr(tstr, " ")
                        Select Case i
                            Case 1
                                npcwarp(UBound(npcwarp)).location = Trim(Left(tstr, Index - 1))
                            Case 2
                                npcwarp(UBound(npcwarp)).pos.X = Val(Trim(Left(tstr, Index - 1)))
                            Case 3
                                npcwarp(UBound(npcwarp)).pos.Y = Val(Trim(tstr))
                        End Select
                        If i < 3 Then tstr = Trim(Right(tstr, Len(tstr) - Index))
                    Next
                Case "action_script"
                    npcwarp(UBound(npcwarp)).Script = Trim(Right(tstr, Len(tstr) - Index))
                Case "cause"
                    npcwarp(UBound(npcwarp)).Cause = Trim(Right(tstr, Len(tstr) - Index))
                Case "target"
                    npcwarp(UBound(npcwarp)).target = Trim(Right(tstr, Len(tstr) - Index))
                    tmp = UBound(npcwarp)
            End Select
        End If
    End If
Loop
Close 10
Exit Sub
errie:
Close 10
MsgBox "Error!!! on loading 'maproute\npcwarp.txt' " & Err.Description, vbCritical
'Unload MDIfrmMain
End Sub

Public Sub Load_NPC_Profile()
On Error GoTo errie:
Open App.Path & "\profile\npc.txt" For Input As #10
Dim tstr, tstr2 As String
Dim Index, i As Integer
Dim tmp As Integer
use_npc_profile = False
ReDim ai_npc(0)
Do While Not EOF(10)
    Line Input #10, tstr
    tstr = Trim(tstr)
    Index = InStr(tstr, "#")
    If Index > 0 Then
        If ai_npc(0).Cause <> "" Then ReDim Preserve ai_npc(UBound(ai_npc) + 1)
    Else
        Index = InStr(tstr, "=")
        If Index > 0 Then
            Select Case LCase(Trim(Left(tstr, Index - 1)))
                Case "location"
                    tstr = Trim(Right(tstr, Len(tstr) - Index))
                    For i = 1 To 3
                        Index = InStr(tstr, " ")
                        Select Case i
                            Case 1
                                ai_npc(UBound(ai_npc)).location = Trim(Left(tstr, Index - 1))
                            Case 2
                                ai_npc(UBound(ai_npc)).pos.X = Val(Trim(Left(tstr, Index - 1)))
                            Case 3
                                ai_npc(UBound(ai_npc)).pos.Y = Val(Trim(tstr))
                        End Select
                        If i < 3 Then tstr = Trim(Right(tstr, Len(tstr) - Index))
                    Next
                Case "action_script"
                    ai_npc(UBound(ai_npc)).Script = Trim(Right(tstr, Len(tstr) - Index))
                Case "target"
                    ai_npc(UBound(ai_npc)).target = Trim(Right(tstr, Len(tstr) - Index))
                Case "cause"
                    ai_npc(UBound(ai_npc)).Cause = Trim(Right(tstr, Len(tstr) - Index))
                Case "use_npc_profile"
                    If Trim(Right(tstr, Len(tstr) - Index)) = "1" Then use_npc_profile = True
            End Select
        End If
    End If
Loop
Close 10
Exit Sub
errie:
Close 10
MsgBox "Error!!! on loading 'profile\npc.txt' " & Err.Description, vbCritical
'Unload MDIfrmMain
End Sub

