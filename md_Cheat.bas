Attribute VB_Name = "md_Cheat"
Option Explicit

Public CurCIP As String
Public CurMIP As String
Public CurMPort As Integer
Public WaitCheat As Byte
Public WaitCMap As String
Type mcIPSList
    MName As String
    MIP As String
    MPort As Integer
End Type
Public IPList() As mcIPList
Type mcIPList
    CName As String
    CIP As String
    MList() As mcIPSList
End Type

Sub Load_IPList()
On Error GoTo errie
    Dim tstr As String, tspl() As String
    ReDim IPList(0)
    ReDim IPList(0).MList(0)
    Open App.Path & "\profile\iplist.txt" For Input As #24
        Do Until EOF(24)
            Line Input #24, tstr
            If Left(tstr, 1) = "#" Then
                If Len(IPList(UBound(IPList)).CName) > 0 Then
                    ReDim Preserve IPList(UBound(IPList) + 1)
                    ReDim IPList(UBound(IPList)).MList(0)
                End If
            End If
            If InStr(tstr, "=") > 0 Then
                    Select Case LCase(Trim(Left(tstr, InStr(tstr, "=") - 1)))
                        Case "cserver"
                            IPList(UBound(IPList)).CName = Trim(Right(tstr, Len(tstr) - InStr(tstr, "=")))
                        Case "cserverip"
                            IPList(UBound(IPList)).CIP = Trim(Right(tstr, Len(tstr) - InStr(tstr, "=")))
                        Case "map"
                            tspl = Split(Trim(Right(tstr, Len(tstr) - InStr(tstr, "="))), "#", 2)
                            If Val(tspl(0)) > UBound(IPList(UBound(IPList)).MList) Then ReDim Preserve IPList(UBound(IPList)).MList(Val(tspl(0)))
                            IPList(UBound(IPList)).MList(Val(tspl(0))).MName = tspl(1)
                        Case "mapip"
                            tspl = Split(Trim(Right(tstr, Len(tstr) - InStr(tstr, "="))), "#", 2)
                            If Val(tspl(0)) > UBound(IPList(UBound(IPList)).MList) Then ReDim Preserve IPList(UBound(IPList)).MList(Val(tspl(0)))
                            IPList(UBound(IPList)).MList(Val(tspl(0))).MIP = tspl(1)
                        Case "mapport"
                            tspl = Split(Trim(Right(tstr, Len(tstr) - InStr(tstr, "="))), "#", 2)
                            If Val(tspl(0)) > UBound(IPList(UBound(IPList)).MList) Then ReDim Preserve IPList(UBound(IPList)).MList(Val(tspl(0)))
                            IPList(UBound(IPList)).MList(Val(tspl(0))).MPort = Val(tspl(1))
                    End Select
            End If
        Loop
    Close #24
    Save_IPList
    Exit Sub
errie:
    Close #24
    MsgBox "Error on loading 'profile\iplist.txt' : " & Err.Description, vbCritical
    Err.Clear
End Sub
Sub Save_IPList()
On Error GoTo errie
    Dim i&, j&
    Open App.Path & "\profile\iplist.txt" For Output As #24
        For i = 0 To UBound(IPList)
            Print #24, "#"
            Print #24, "cserver = " & IPList(i).CName
            Print #24, "cserverip = " & IPList(i).CIP
            For j = 0 To UBound(IPList(i).MList)
                Print #24, "map = " & j & "#" & IPList(i).MList(j).MName
                Print #24, "mapip = " & j & "#" & IPList(i).MList(j).MIP
                Print #24, "mapport = " & j & "#" & IPList(i).MList(j).MPort
            Next
        Next
    Close #24
    Exit Sub
errie:
    Close #24
    print_funcerr "Save_IPList", Err.number, Err.Description
    Err.Clear
End Sub
Sub UpdateMIP(MNames As String, MIPs As String, MPorts As Integer)
On Error GoTo errie
    Dim i&, j&
    For i = 0 To UBound(IPList)
        If IPList(i).CIP = CurCIP Then
            For j = 0 To UBound(IPList(i).MList)
                If MNames = IPList(i).MList(j).MName Then
                    IPList(i).MList(j).MIP = MIPs
                    IPList(i).MList(j).MPort = MPorts
                    Save_IPList
                    Exit Sub
                End If
            Next
            If Len(IPList(i).MList(UBound(IPList(i).MList)).MName) > 0 Then ReDim Preserve IPList(i).MList(UBound(IPList(i).MList) + 1)
            IPList(i).MList(UBound(IPList(i).MList)).MName = MNames
            IPList(i).MList(UBound(IPList(i).MList)).MIP = MIPs
            IPList(i).MList(UBound(IPList(i).MList)).MPort = MPorts
            Save_IPList
            Exit Sub
        End If
    Next
    If Len(IPList(UBound(IPList)).CName) > 0 Then ReDim Preserve IPList(UBound(IPList) + 1)
    IPList(UBound(IPList)).CIP = CurCIP
    IPList(UBound(IPList)).CName = ServerList(NumServ).Name
    ReDim IPList(UBound(IPList)).MList(0)
    IPList(UBound(IPList)).MList(0).MName = MNames
    IPList(UBound(IPList)).MList(0).MIP = MIPs
    IPList(UBound(IPList)).MList(0).MPort = MPorts
    Save_IPList
    Exit Sub
errie:
    print_funcerr "UpdateMIP", Err.number, Err.Description
    Err.Clear
End Sub
Sub ChangeMap(MNames As String)
On Error GoTo errie
    WaitCMap = MNames
    Dim i&, j&
    For i = 0 To UBound(IPList)
        If IPList(i).CIP = CurCIP Then
            For j = 0 To UBound(IPList(i).MList)
                If MNames = IPList(i).MList(j).MName Then
                    Chat "System : [cheat] - Changing map to [" & MNames & "]", vbRed
                    MapName = MNames
                    CurrentMap = MNames
                    Load_WayPoint MNames
                    Load_Field MNames
                    LockMapName = MNames
                    DoConnect IPList(i).MList(j).MIP, CLng(IPList(i).MList(j).MPort)
                    WaitCheat = 1
                    Exit Sub
                End If
            Next
        End If
    Next
    Chat "System : [cheat] - Map IP of [" & MNames & "] not found in database."
    Exit Sub
errie:
    print_funcerr "ChangeMap", Err.number, Err.Description
    Err.Clear
End Sub
