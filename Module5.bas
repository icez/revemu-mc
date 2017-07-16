Attribute VB_Name = "Mod_File"
Type ItemAuto
    Auto As Boolean
    Name As String
    Time As Integer
    TimeCount As Date
End Type

Type CharStatus
    Name As String
End Type

Type SpecialStatus
    Name As String
    Active As Boolean
End Type

Type SkillSP
    Name As String
    SP As String
End Type

Type SelfSkill
    ID As Integer
    Name As String
    Time As Integer
    TimeCount As Date
    Level As Byte
    SPNeed As Integer
    SPmin As Byte
    SPmax As Byte
    Auto_reuse As Boolean
    StatusNum As Long
    lusetime As Double
    Mode As String
End Type

Type NPC
    ID As String
    pos As Coord
    NameID As Long
    Name As String
    location As String
End Type

'Character Status
Public CurCharStatus() As CharStatus
Public CurStatus() As SpecialStatus

'Sefskill Profile
Public IsSelfSkill As Boolean
Public tmpSelfskill As String
Public AutoSkill() As SelfSkill
Public DelaySelfSkill As Integer
Public UsingSelfSkill As Boolean
Public SelfSkillDelay As Integer

'Auto use Item
Public AutoItem As ItemAuto
Public SPUse() As SkillSP

'NPC.txt Related
Public myNPC() As NPC

'Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Public TheWindowsDirectory As String

Public Function IsAvoidID(AID As String)
    If AvoidID(0) > 0 Then
        For i = 0 To UBound(AvoidID)
            If MakePort(AID) = AvoidID(i) Then
                IsAvoidID = True
                Exit Function
            End If
        Next
    End If
    IsAvoidID = False
End Function

Public Sub Load_AvoidID()
On Error GoTo Out
Dim tstr As String
Dim myVal As Long, tmp As Long
ReDim AvoidID(0)
'Close 1
Open App.Path & "\avoid\avoidid.txt" For Input As #1
Do While Not EOF(1)
    Line Input #1, tstr
    myVal = Val(Left(tstr, Len(tstr)))
    AvoidID(UBound(AvoidID)) = myVal
    'tmp = GetLong(AvoidID(UBound(AvoidID)))
    ReDim Preserve AvoidID(UBound(AvoidID) + 1)
Loop
ReDim Preserve AvoidID(UBound(AvoidID) - 1)
Close 1
Exit Sub
Out:
Close 1
If Err.number <> 53 Then MsgBox "Error! on loading 'avoid\avoidid.txt'", vbCritical, Err.Description & "!"
'Unload MDIfrmMain
End Sub

Public Sub Save_NPC()
    Open App.Path & "\maproute\npcinfo.txt" For Output As #1
    Dim i As Integer
    For i = 0 To UBound(myNPC)
        If Len(myNPC(i).Name) > 0 Then
            With myNPC(i)
                If Len(Trim(.Name)) > 0 Then Print #1, MakeHex(.ID) & " " & Replace(Trim(.Name), " ", "_") & " " & .location & " " & .pos.Y & " " & .pos.X
            End With
        End If
    Next
    Close 1
End Sub

Public Sub Load_NPC()
On Error GoTo errie
    Dim tstr As String
    Dim Index, i, j As Integer
    Dim ID As Integer
    ReDim myNPC(0)
    Open App.Path & "\maproute\npcinfo.txt" For Input As #10
    Do While Not EOF(10)
        Line Input #10, tstr
        If tstr = "" Then GoTo end_loop
        If myNPC(0).location <> "" Then ReDim Preserve myNPC(UBound(myNPC) + 1)
        For i = 1 To 5
            Index = InStr(tstr, " ")
            Dim tstr2 As String
            Select Case i
                Case 1
                    tstr2 = Trim(Left(tstr, Index - 1))
                    myNPC(UBound(myNPC)).ID = ""
                    For j = 1 To 7 Step 2
                        myNPC(UBound(myNPC)).ID = myNPC(UBound(myNPC)).ID & Chr(Val("&H" & Mid(tstr2, j, 2)))
                    Next
                Case 2
                    myNPC(UBound(myNPC)).Name = Trim(Left(tstr, Index - 1))
                Case 3
                    myNPC(UBound(myNPC)).location = Trim(Left(tstr, Index - 1))
                Case 4
                    myNPC(UBound(myNPC)).pos.Y = Val(Trim(Left(tstr, Index - 1)))
                Case 5
                    myNPC(UBound(myNPC)).pos.X = Val(Trim(tstr))
            End Select
            If i < 5 Then tstr = Trim(Right(tstr, Len(tstr) - Index))
        Next
end_loop:
    Loop
    Close 10
    Exit Sub
errie:
    Close 10
    MsgBox "Error!!! on loading 'maproute\npcinfo.txt'", vbCritical
    Unload MDIfrmMain
End Sub

Public Sub Reset_Time()
    If AutoSkill(0).Name <> "" Then
        Dim i As Integer
        For i = 0 To UBound(AutoSkill)
            AutoSkill(i).TimeCount = 0
        Next
    End If
    AutoItem.TimeCount = 0
    reset_status
End Sub

Public Sub reset_status()
    Dim i As Integer
    For i = 0 To UBound(CurStatus)
        CurStatus(i).Active = False
    Next
    frmStatus.lstStatus.Clear
End Sub

Public Sub Load_Char_Status()
On Error GoTo errie
    Dim tstr As String
    Dim Index As Integer
    Dim ID As Integer
    ReDim CurCharStatus(0)
    Open App.Path & "\table\status.txt" For Input As #10
    Do While Not EOF(10)
        Line Input #10, tstr
        Index = InStr(tstr, " ")
        ID = Val(Left(tstr, Index))
        If ID > UBound(CurCharStatus) Then ReDim Preserve CurCharStatus(ID)
        CurCharStatus(ID).Name = Right(tstr, Len(tstr) - Index)
    Loop
    Close 10
    Exit Sub
errie:
    Close 10
    MsgBox "Error!!! on loading 'table\status.txt'", vbCritical
    Unload MDIfrmMain
End Sub

Public Sub Load_Special_Status()
On Error GoTo errie
    Dim tstr As String
    Dim Index As Integer
    Dim ID As Integer
    ReDim CurStatus(0)
    Open App.Path & "\table\specialstatus.txt" For Input As #10
    Do While Not EOF(10)
        Line Input #10, tstr
        Index = InStr(tstr, " ")
        ID = Val(Left(tstr, Index))
        If ID > UBound(CurStatus) Then ReDim Preserve CurStatus(ID)
        CurStatus(ID).Name = Trim(Right(tstr, Len(tstr) - Index))
    Loop
    Close 10
    Exit Sub
errie:
    Close 10
    MsgBox "Error!!! on loading 'table\specialstatus.txt'", vbCritical
    Unload MDIfrmMain
End Sub

Public Function Get_SkillUseSp(Name As String) As String
On Error GoTo errie
    Dim tstr As String
    Dim Index As Integer
    Dim i As Integer
    Dim number As Integer
    Dim text As String
    Dim tmp As String
    text = ""
    Open App.Path & "\table\leveluseskillspamount.txt" For Input As #10
    i = 0
    Do While Not EOF(10)
        Line Input #10, tstr
        Index = InStr(tstr, "#")
        If Index > 0 Then
            If Name = Left(tstr, Len(tstr) - 1) Then
                text = ""
                i = 1
            ElseIf i > 0 Then
                Dim SP As Integer
                SP = Val(Left(tstr, Len(tstr) - 1))
                If SP < 10 Then
                    tmp = "00"
                ElseIf SP < 100 Then
                    tmp = "0"
                End If
                tmp = tmp & CStr(SP)
                text = text & tmp
                i = i + 1
            End If
        Else
            Index = InStr(tstr, "@")
            If Index > 0 And text <> "" Then
                i = 0
                Get_SkillUseSp = text
                Close 10
                Exit Function
            End If
        End If
    Loop
    Get_SkillUseSp = ""
    Close 10
    Exit Function
errie:
    Close 10
    MsgBox "Error!!! on loading 'table\leveluseskillspamount.txt'", vbCritical
    Unload MDIfrmMain
End Function

