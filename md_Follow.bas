Attribute VB_Name = "md_Follow"
Option Explicit

Sub FollowCheck()
On Error GoTo errie
    'If Not FollowMode.Active Then Exit Sub
    Dim i&, testMap$, testCur$
    For i = 0 To UBound(Party)
        If LCase(Party(i).Name) = LCase(FollowMode.Name) Then Exit For
    Next
    'If Not Party(i).Online Then
    '    Chat "System : [Party] - Partner not online. PM Checking..."
    '    SendPARChat "MC:CHECK:" & FollowMode.Name
    '    FollowCancel
    '    Exit Sub
    'End If
    If InStr(Party(i).Map, ".gat") > 0 Then testMap = Replace(Party(i).Map, ".gat", "") Else testMap = Party(i).Map
    If InStr(MapName, ".gat") > 0 Then testCur = Replace(MapName, ".gat", "") Else testCur = MapName
    If testMap <> testCur Then
        Chat "System : [Party] - Follow target is online but isn't in same map. Paused follow mode."
    '    FollowCancel
        Exit Sub
    End If
    If Party(i).ID <> FollowMode.AID Then
        Chat "System : [Party] - Found follow target [" & FollowMode.Name & "]. Locking . . ."
    End If
    FollowMode.AID = Party(i).ID
    FollowMode.curPos = Party(i).pos
    Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "FollowCheck", Err.number, Err.Description
    Err.Clear
End Sub
