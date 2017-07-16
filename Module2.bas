Attribute VB_Name = "Module2"
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
'dss
Public Const HTLEFT = 10
Public Const HTRight = 11
Public Const HTUP = 12
Public Const HTDown = 13
Type SkillName
    Name As String
    raw As String
    Sp As String
    maxsp As Integer
End Type

Public SkillIDName() As SkillName

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Function GetSkillIDbyName(tstr As String) As Integer
    Dim i As Integer
    If UBound(SkillIDName) > 0 Then
        For i = 0 To UBound(SkillIDName)
            If SkillIDName(i).raw = tstr Then
                GetSkillIDbyName = i + 1
                Exit Function
            End If
        Next
    End If
    GetSkillIDbyName = 0
End Function

Public Function CanUseSkill(ByVal Name As String, level As Byte, ByVal Sp As Integer) As Boolean
    Dim requiresp As Integer
    requiresp = Get_UseSPbyName(Name, level)
    If Sp >= requiresp And requiresp <> 0 Then
        'Stat "Can use " & name & vbCrLf
        CanUseSkill = True
    Else
        'Stat sp & " Can't use " & CStr(level) & " " & name & " " & CStr(requiresp) & vbCrLf
        CanUseSkill = False
    End If
End Function

Public Function Get_UseSPbyName(ByVal Name As String, ByVal level As Byte) As Integer
On Error GoTo errie
    Dim i As Integer
    For i = 0 To UBound(SkillIDName)
        If SkillIDName(i).raw = Name Then
            If SkillIDName(i).Sp <> "" Then
                Index = 1
                If level > 1 Then Index = (level * 3) - 2
                Get_UseSPbyName = Val(Mid(SkillIDName(i).Sp, Index, 3))
            Else
                Get_UseSPbyName = SkillIDName(i).maxsp
            End If
            Exit Function
        End If
    Next
errie:
    Get_UseSPbyName = 0
End Function

Public Function Get_UseSPbyID(ByVal ID As Integer, ByVal level As Byte) As Integer
On Error GoTo errie
    Dim Index As Integer
    If UBound(SkillIDName) >= ID - 1 Then
        If SkillIDName(ID - 1).Sp <> "" Then
            If SkillIDName(i).Sp <> "" Then
                Index = 1
                If level > 1 Then Index = (level * 3) - 2
                Get_UseSPbyID = Val(Mid(SkillIDName(i).Sp, Index, 3))
            Else
                Get_UseSPbyID = SkillIDName(i).maxsp
            End If
            Exit Function
        End If
    End If
errie:
    Get_UseSPbyID = 0
End Function

Public Sub Load_SkillName()
On Error GoTo Out
Dim tstr As String
ReDim SkillIDName(0)
'Close 1
Open App.Path & "\table\Skillname.txt" For Input As #1
Do While Not EOF(1)
    Line Input #1, tstr
    Index = InStr(tstr, "#")
    If Index > 0 Then
        SkillIDName(UBound(SkillIDName)).raw = Left(tstr, Index - 1)
        tstr = Trim(Right(tstr, Len(tstr) - Index))
        SkillIDName(UBound(SkillIDName)).Name = Left(tstr, Len(tstr))
        SkillIDName(UBound(SkillIDName)).Sp = Get_SkillUseSp(SkillIDName(UBound(SkillIDName)).raw)
        ReDim Preserve SkillIDName(UBound(SkillIDName) + 1)
    End If
Loop
If (UBound(SkillIDName) > 0) Then ReDim Preserve SkillIDName(UBound(SkillIDName) - 1)
Close 1
Exit Sub
Out:
Close 1
MsgBox "Error!!! on loading 'table\skillname.txt' : " & Err.Description, vbCritical
End Sub

Public Function Get_SkillName(rawname As String) As String
Open App.Path & "\table\Skillname.txt" For Input As #1
Dim tstr As String
Dim text As String
Dim Index As Integer
text = "#"
Do While Not EOF(1)
    Input #1, tstr
    Index = InStr(1, tstr, text, vbTextCompare)
    If Index > 0 Then
        If Left(tstr, Index - 1) = rawname Then
            Get_SkillName = Mid(tstr, Index + 1, Len(tstr) - Index)
            Close 1
            Exit Function
        End If
    End If
Loop
Get_SkillName = rawname
Close 1
End Function

Public Function Get_SkillD(rawname As String) As String
Dim tstr As String
Dim text As String
Dim Index As Integer
text = "#"
Open App.Path & "\table\SkillD.txt" For Input As #1
Do While Not EOF(1)
    Input #1, tstr
    Index = InStr(1, tstr, text, vbTextCompare)
    If Index > 0 Then
        If Left(tstr, Index - 1) = rawname Then
            Input #1, tstr
            Input #1, tstr
            Do
            Get_SkillD = Get_SkillD & tstr
            Input #1, tstr
            If (tstr <> "#") Then Get_SkillD = Get_SkillD & vbCrLf
            Loop While (tstr <> "#")
            Close 1
            Exit Function
        End If
    End If
Loop
Get_SkillD = "Nothing"
Close 1
End Function

