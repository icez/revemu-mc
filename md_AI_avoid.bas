Attribute VB_Name = "md_AI_avoid"
Option Explicit
Public AvoidID() As Long

Sub AI_Avoid(PPLName As String)
On Error GoTo errie
    If IsAvoid(PPLName) And Not MoveOnly Then
        If Find_EventID("OnAvoidListAppear") < 0 Then
            Open App.Path & "\log\warning.txt" For Append As #8
            Chat Date & "@" & Time & ":" & "Detect avoid name [" & PPLName & "], Closed program...", MColor.Fail
            Print #8, Date & "@" & Time & ":" & "Detect avoid name [" & PPLName & "], Closed program..."
            Close 8
            Winsock_SendPacket String(15, 0), True
            ForceExit
        Else
            Open App.Path & "\log\warning.txt" For Append As #8
            Chat Date & "@" & Time & ":" & "Detect avoid name [" & PPLName & "], Do events...", MColor.Fail
            Print #8, Date & "@" & Time & ":" & "Detect avoid name [" & PPLName & "], Do events..."
            Close 8
            CheckEvent "OnAvoidListAppear", "name=" & PPLName
        End If
    End If
    Exit Sub
errie:
    print_funcerr "AI_Avoid", Err.number, Err.Description
    Err.Clear
End Sub

Sub AI_AvoidID(ByVal AID As String, Optional PacketID As String)
On Error GoTo errie
    If IsAvoidID(AID) Then
        Open App.Path & "\log\warning.txt" For Append As #8
        Chat Date & "@" & Time & ": Detect avoid ID [" & MakePort(AID) & "], " & IIf(MODDC.AvoidTime > 0, "Delaying for login", "Closed program") & "..."
        Print #8, Date & "@" & Time & ": Detect avoid ID [" & CStr(MakePort(AID)) & "] (" & MakeHex(People(UBound(People)).ID) & ") at [" & MapName & "] (" & People(UBound(People)).Pos.Y & ":" & People(UBound(People)).Pos.X & "), You're at (" & curPos.Y & ":" & curPos.X & ") ..."
        If Find_EventID("OnGMAppear") < 0 Or Len(PacketID) > 0 Then
            If MODDC.AvoidTime > 0 Then
                Stat "Delaying to login for " & MODDC.AvoidTime & " minutes (" & PacketID & ")", vbRed
                Print #8, "Delaying to login for " & MODDC.AvoidTime & " minutes (" & PacketID & ")"
                Print #8, ""
                Close 8
                MODDelay.DualLogin = MODDC.AvoidTime * 6000
                frmMain.Winsock1.Close
                Exit Sub
            End If
            Print #8, "Close program"
            Print #8, ""
            Winsock_SendPacket String(15, 0), True
            ForceExit
        End If
        Close 8
    End If
Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "AI_AvoidID", Err.number, Err.Description
    Err.Clear
End Sub

Sub AI_AVCheck()
    Exit Sub
    Dim i&
    For i = 0 To UBound(AvoidID) - 1
        Winsock_SendPacket IntToChr(&H94) & LngToChr(AvoidID(i)), True
    Next
End Sub
