Attribute VB_Name = "Mod_Profile"
Type Equip_Profile
    unequip As Boolean
    Equipment As String
    Monster As String
    Arrow As String
End Type

Type Recovery_Profile
    Name As String
    RecovItem As String
    RecovSkill As String
End Type

Type Accessoty_Profile
    Equip_L As String
    Equip_R As String
    Monster As String
End Type

Type Pet
    ID As String
    Name As String
    Type As String
    Level As Byte
    Status As Byte
    pos As Coord
    Relation As Integer
    Equipment As String
    AutoFeed As Boolean
    DelayFeed As Byte
    Delay As Byte
    FeedLimit As Byte
End Type

Type Card_Profile
    Name As String
    number As Byte
End Type

Type Emoticon
    Key As String
    detail As String
End Type

Public Emotions() As Emoticon
Public PetWinClose As Boolean
Public MyPet As Pet
Public use_recovery_profile As Boolean
Public delay_recovery As Integer
Public delay_count As Integer
Public ai_recovery() As Recovery_Profile
Public ai_Accessory() As Accessoty_Profile
Public use_accessory_profile As Boolean
Public use_eq_profile As Boolean
Public Auto_Chatlog As Boolean
Public ai_equip() As Equip_Profile



Public Sub change_equip(tstr As String, isBoth As Boolean)
    Dim i As Long
    If isBoth Then
        For i = 0 To UBound(AllInv)
            With AllInv(i)
            If (.pos = 32 Or .pos = 34 Or .pos = 2) And (InStr(tstr, .Name) = 0) Then
                Winsock_SendPacket IntToChr(&HAB) & IntToChr(i), True
            End If
            End With
        Next
    End If
    Dim tspl() As String, j&
    tspl = Split(tstr, ",")
    For j = 0 To UBound(tspl)
        For i = 0 To UBound(AllInv)
            With AllInv(i)
                If (((.Category > 3 And .Category < 6) Or .Category > 7) And .Category <> 10 And .pos = 0) Then
                    If (InStr(1, tspl(j), .Name, vbTextCompare) > 0) Then
                       Winsock_SendPacket IntToChr(&HA9) & IntToChr(i) & AllInv(i).Type, True
                        Stat "Try to equip [" & .Name & "]..." & vbCrLf
                    End If
                End If
                If (.Category = 10 And .pos = 0) Then
                    If (InStr(1, tspl(j), .Name, vbTextCompare) > 0) Then 'And AllInv(ArrowNumber).Name <> .Name Then
                        Winsock_SendPacket IntToChr(&HA9) & IntToChr(i) & Chr(0) & Chr(0), True
                        Stat "Try to arrow [" & .Name & "]..." & vbCrLf
                        ArrowChangeNumber = Val(i)
                    End If
                End If
            End With
        Next
    Next
End Sub

Public Sub Check_Equip(Name As String)
    If UBound(ai_equip) = 0 Or Not use_eq_profile Then Exit Sub
    Dim i As Integer
    For i = 1 To UBound(ai_equip)
        If InStr(ai_equip(i).Monster, Name) > 0 Then
            change_equip ai_equip(i).Equipment, ai_equip(i).unequip
            Exit Sub
        End If
    Next
    change_equip (ai_equip(0).Equipment), ai_equip(0).unequip
End Sub

Public Sub Load_Equip_Profile()
On Error GoTo errie:
Open App.Path & "\profile\equip_monster.txt" For Input As #10
Dim tstr As String
Dim Index As Integer
use_eq_profile = False
ReDim ai_equip(0)
ai_equip(0).unequip = False
Do While Not EOF(10)
    Line Input #10, tstr
    Index = InStr(tstr, "#")
    If Index > 0 Then
        If ai_equip(0).Monster <> "" Then
            ReDim Preserve ai_equip(UBound(ai_equip) + 1)
            ai_equip(UBound(ai_equip)).unequip = False
        End If
    Else
        Index = InStr(tstr, "=")
        If Index > 0 Then
            Select Case LCase(Trim(Left(tstr, Index - 1)))
                Case "equipment"
                    ai_equip(UBound(ai_equip)).Equipment = Trim(Right(tstr, Len(tstr) - Index))
                Case "monster"
                    ai_equip(UBound(ai_equip)).Monster = Trim(Right(tstr, Len(tstr) - Index))
                Case "arrow"
                    ai_equip(UBound(ai_equip)).Arrow = Trim(Right(tstr, Len(tstr) - Index))
                Case "unequip_both_hand"
                    'If Trim(Right(tstr, Len(tstr) - index)) = "1" Then ai_equip(UBound(ai_equip)).unequip = True
                Case "use_equipment_profile"
                    If Trim(Right(tstr, Len(tstr) - Index)) = "1" Then use_eq_profile = True
            End Select
        End If
    End If
Loop
Close 10
Exit Sub
errie:
Close 10
MsgBox "Error!!! on loading 'profile\equip_monster.txt'", vbCritical
'Unload MDIfrmMain
End Sub

Public Sub Load_SelfSkill_Profile()
On Error GoTo errie:
Open App.Path & "\profile\selfskill.txt" For Input As #10
Dim tstr As String
Dim Index As Integer
IsSelfSkill = False
ReDim AutoSkill(0)
Do While Not EOF(10)
    Line Input #10, tstr
    Index = InStr(tstr, "#")
    If Index > 0 Then
        If AutoSkill(0).Name <> "" Then
            ReDim Preserve AutoSkill(UBound(AutoSkill) + 1)
        End If
        AutoSkill(UBound(AutoSkill)).lusetime = 0
    Else
        Index = InStr(tstr, "=")
        If Index > 0 Then
            Select Case LCase(Trim(Left(tstr, Index - 1)))
                Case "skill_name"
                    AutoSkill(UBound(AutoSkill)).Name = Trim(Right(tstr, Len(tstr) - Index))
                    AutoSkill(UBound(AutoSkill)).ID = GetSkillIDbyName(AutoSkill(UBound(AutoSkill)).Name)
                Case "level"
                    AutoSkill(UBound(AutoSkill)).Level = Val(Trim(Right(tstr, Len(tstr) - Index)))
                    AutoSkill(UBound(AutoSkill)).SPNeed = Get_UseSPbyID(AutoSkill(UBound(AutoSkill)).ID, AutoSkill(UBound(AutoSkill)).Level)
                Case "loop_time"
                    AutoSkill(UBound(AutoSkill)).Time = Val(Trim(Right(tstr, Len(tstr) - Index)))
                    AutoSkill(UBound(AutoSkill)).TimeCount = 0
                Case "spmin"
                    AutoSkill(UBound(AutoSkill)).SPmin = Val(Trim(Right(tstr, Len(tstr) - Index)))
                Case "spmax"
                    AutoSkill(UBound(AutoSkill)).SPmax = Val(Trim(Right(tstr, Len(tstr) - Index)))
                Case "statusnum"
                    AutoSkill(UBound(AutoSkill)).StatusNum = Val(Trim(Right(tstr, Len(tstr) - Index)))
                Case "mode"
                    AutoSkill(UBound(AutoSkill)).Mode = Trim(Right(tstr, Len(tstr) - Index))
                Case "auto_reuse"
                    AutoSkill(UBound(AutoSkill)).Auto_reuse = CBool(Val(Trim(Right(tstr, Len(tstr) - Index))))
                Case "use_selfskill_profile"
                    If Trim(Right(tstr, Len(tstr) - Index)) = "1" Then IsSelfSkill = True
            End Select
        End If
    End If
Loop
Close 10
Exit Sub
errie:
Close 10
MsgBox "Error!!! on loading 'profile\selfskill.txt'", vbCritical
'Unload MDIfrmMain
End Sub

Public Sub Load_Recovery_Profile()
On Error GoTo errie:
Open App.Path & "\profile\statusrecovery.txt" For Input As #10
Dim tstr As String
Dim Index As Integer
Dim ID As Integer
use_recovery_profile = False
ReDim ai_recovery(0)
Do While Not EOF(10)
    Line Input #10, tstr
    Index = InStr(tstr, "#")
    If Index > 0 Then
        If ai_recovery(0).Name <> "" Then
            ReDim Preserve ai_recovery(UBound(ai_recovery) + 1)
        End If
    Else
        Index = InStr(tstr, "=")
        If Index > 0 Then
            Select Case LCase(Trim(Left(tstr, Index - 1)))
                Case "status_name"
                    ai_recovery(UBound(ai_recovery)).Name = Trim(Right(tstr, Len(tstr) - Index))
                Case "delay"
                    delay_recovery = Trim(Right(tstr, Len(tstr) - Index))
                Case "recovery_item"
                    ai_recovery(UBound(ai_recovery)).RecovItem = Trim(Right(tstr, Len(tstr) - Index))
                Case "recovery_skill"
                    ai_recovery(UBound(ai_recovery)).RecovSkill = Trim(Right(tstr, Len(tstr) - Index))
                Case "use_recovery_profile"
                    If Trim(Right(tstr, Len(tstr) - Index)) = "1" Then use_recovery_profile = True
            End Select
        End If
    End If
Loop
Close 10
Exit Sub
errie:
Close 10
MsgBox "Error!!! on loading 'profile\statusrecovery.txt'", vbCritical
'Unload MDIfrmMain
End Sub

Public Sub Change_Accessory(acc_L As String, acc_R As String)
    Dim i, j As Long
    For i = 0 To UBound(AllInv)
        With AllInv(i)
        If (.Category = 4 And .pos = 0) Then
            If (InStr(acc_L, .Name) > 0) Then
                For j = 0 To UBound(AllInv)
                    With AllInv(j)
                        If (.pos = 8) And (InStr(acc_L, .Name) = 0) Then
                            Winsock_SendPacket IntToChr(&HAB) & IntToChr(j), True
                        End If
                    End With
                    DoEvents
                Next
               Winsock_SendPacket IntToChr(&HA9) & IntToChr(CLng(i)) & AllInv(i).Type, True
                Stat "Try to equip [" & .Name & "][Left Accessory]..." & vbCrLf
            End If
        End If
      End With
    Next
    
     For i = 0 To UBound(AllInv)
        With AllInv(i)
        If (.Category = 4 And .pos = 0) Then
             If (InStr(acc_R, .Name) > 0) Then
                 For j = 0 To UBound(AllInv)
                    With AllInv(j)
                        If (.pos = 128) And (InStr(acc_R, .Name) = 0) Then
                            Winsock_SendPacket IntToChr(&HAB) & IntToChr(j), True
                         End If
                    End With
                 Next
               Winsock_SendPacket IntToChr(&HA9) & IntToChr(CLng(i)) & AllInv(i).Type, True
                Stat "Try to equip [" & .Name & "][Right Accessory]..." & vbCrLf
            End If
       End If
       End With
    Next
End Sub
Public Sub Check_Accessory(Name As String)
    If UBound(ai_Accessory) = 0 Or Not use_accessory_profile Then Exit Sub
    Dim i As Integer
    For i = 1 To UBound(ai_Accessory)
        If InStr(ai_Accessory(i).Monster, Name) > 0 Then
            Change_Accessory ai_Accessory(i).Equip_L, ai_Accessory(i).Equip_R
            Exit Sub
        End If
        DoEvents
    Next
    Change_Accessory ai_Accessory(0).Equip_L, ai_Accessory(0).Equip_R
End Sub
Public Sub Load_accessory_Profile()
On Error GoTo errie:
Open App.Path & "\profile\equip_accessory.txt" For Input As #10
Dim tstr As String
Dim Index As Integer
Dim ID As Integer
use_accessory_profile = False
ReDim ai_Accessory(0)
Do While Not EOF(10)
    Line Input #10, tstr
    Index = InStr(tstr, "#")
    If Index > 0 Then
        If ai_Accessory(0).Equip_L <> "" Then
            ReDim Preserve ai_Accessory(UBound(ai_Accessory) + 1)
        End If
    Else
        Index = InStr(tstr, "=")
        If Index > 0 Then
            Select Case LCase(Trim(Left(tstr, Index - 1)))
                Case "equip_left"
                    ai_Accessory(UBound(ai_Accessory)).Equip_L = Trim(Right(tstr, Len(tstr) - Index))
                Case "equip_right"
                    ai_Accessory(UBound(ai_Accessory)).Equip_R = Trim(Right(tstr, Len(tstr) - Index))
                Case "monster"
                    ai_Accessory(UBound(ai_Accessory)).Monster = Trim(Right(tstr, Len(tstr) - Index))
                Case "use_accessory_profile"
                    If Trim(Right(tstr, Len(tstr) - Index)) = "1" Then use_accessory_profile = True
            End Select
        End If
    End If
Loop
Close 10
Exit Sub
errie:
Close 10
Chat "Error!!! on loading 'profile\equip_accessory.txt", 6
'Unload MDIfrmMain
End Sub

