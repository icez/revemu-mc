Attribute VB_Name = "Warp"
Type warp
    name As String
End Type

Type RandomMessage
    message As String
End Type

Public warpuser() As warp

Public register() As RandomMessage
Public change_map() As RandomMessage
Public invalid_command() As RandomMessage
Public invalid_map() As RandomMessage
Public finish() As RandomMessage

Public warpportal  As Coord
Public warppass As String
Public warpmap As String

Public Function random_message(incoming() As RandomMessage) As String
    Dim x As Integer
    x = UBound(incoming)
    random_message = incoming(RandomNumber(x, 0)).message
End Function

Public Sub Load_Warpconfig()
    ReDim register(0)
    ReDim change_map(0)
    ReDim invalid_command(0)
    ReDim invalid_map(0)
    ReDim finish(0)
    Open "table\warpconfig.txt" For Input As #1
    Dim text As String
    Dim test As String
    Dim tstr   As String
    text = "="
    Do While Not EOF(1)
        Input #1, tstr
        If tstr = "[end]" Then
            ReDim Preserve register(UBound(register) - 1)
            ReDim Preserve change_map(UBound(change_map) - 1)
            ReDim Preserve invalid_command(UBound(invalid_command) - 1)
            ReDim Preserve invalid_map(UBound(invalid_map) - 1)
            ReDim Preserve finish(UBound(finish) - 1)
            Close 1
            Exit Sub
        End If
        text = "="
        index = InStr(1, tstr, text, vbTextCompare) - 1
        If index > 0 Then
            If LCase(Trim(Left(tstr, index))) = "register_password" Then
                warpass = Trim(Right(tstr, Len(tstr) - index - 1))
'---------------------------------------------------------------------------------------------------------------------
            ElseIf LCase(Trim(Left(tstr, index))) = "portal_location_x" Then
                warpportal.x = Val(Trim(Right(tstr, Len(tstr) - index - 1)))
'---------------------------------------------------------------------------------------------------------------------
            ElseIf LCase(Trim(Left(tstr, index))) = "portal_location_y" Then
                warpportal.y = Val(Trim(Right(tstr, Len(tstr) - index - 1)))
'---------------------------------------------------------------------------------------------------------------------
            ElseIf LCase(Trim(Left(tstr, index))) = "register_message" Then
                register(UBound(register)).message = Trim(Right(tstr, Len(tstr) - index - 1))
                ReDim Preserve register(UBound(register) + 1)
'---------------------------------------------------------------------------------------------------------------------
            ElseIf LCase(Trim(Left(tstr, index))) = "changemap_message" Then
                change_map(UBound(change_map)).message = Trim(Right(tstr, Len(tstr) - index - 1))
                ReDim Preserve change_map(UBound(change_map) + 1)
'---------------------------------------------------------------------------------------------------------------------
            ElseIf LCase(Trim(Left(tstr, index))) = "invalidcommand_message" Then
                invalid_command(UBound(invalid_command)).message = Trim(Right(tstr, Len(tstr) - index - 1))
                ReDim Preserve invalid_command(UBound(invalid_command) + 1)
'---------------------------------------------------------------------------------------------------------------------
            ElseIf LCase(Trim(Left(tstr, index))) = "invalidmap_message" Then
                invalid_map(UBound(invalid_map)).message = Trim(Right(tstr, Len(tstr) - index - 1))
                ReDim Preserve invalid_map(UBound(invalid_map) + 1)
'---------------------------------------------------------------------------------------------------------------------
            ElseIf LCase(Trim(Left(tstr, index))) = "finish_message" Then
                finish(UBound(finish)).message = Trim(Right(tstr, Len(tstr) - index - 1))
                ReDim Preserve finish(UBound(finish) + 1)
            End If
        End If
    Loop
    Close 1
End Sub

Public Sub whisper(message As String, username As String)
    frmMain.Chat "to " & username & " : " & message
    frmMain.Winsock_SendPacket Chr(&H64 + &H32) + Chr(0) + Make2Byte(Len(message) + 29) + username + String(24 - Len(username), Chr(0)) + message + Chr(0)
End Sub

Public Sub Load_WarpUser()
    ReDim warpuser(0)
    Open "table\warpuser.txt" For Input As #1
    Dim text As String
    Dim test As String
    Dim tstr   As String
    text = "="
    Do While Not EOF(1)
        Input #1, tstr
            warpuser(UBound(warpuser)).name = Trim(tstr)
            ReDim Preserve warpuser(UBound(warpuser) + 1)
    Loop
    Close 1
    ReDim Preserve warpuser(UBound(warpuser) - 1)
End Sub

Public Function check_user(name As String) As Boolean
    Dim i As Integer
    For i = 0 To UBound(warpuser)
        If warpuser(i).name = name Then
            check_user = True
            Exit Function
        End If
    Next
    check_user = False
End Function

Public Function Decode_Whisper(message As String, username As String) As String
    Dim test As String
    Decode_Whisper = "none"
    If InStr(message, "register") > 0 Then
        message = Trim(Right(message, Len(message) - Len("register")))
        If message = warppass Then
            If Not check_user(username) Then
                message = random_message(register) & ", " & username
                'whisper message, username
                Decode_Whisper = message
            End If
        End If
        Exit Function
    End If
    
    If Not check_user(username) Then Exit Function
    Dim text As String
    If InStr(message, "change map") Then
        message = Trim(Right(message, Len(message) - Len("change map")))
        warpmap = Return_MapWarp(message)
        text = username & " " & random_message(change_map) & " [" & warpmap & ".rsw]"
        'whisper message, username
        Decode_Whisper = text
        Exit Function
    End If
    
End Function

Public Function Return_MapWarp(name As String) As String
    
    Open "table\warpmap.txt" For Input As #1
    Dim text, test, tstr As String
    Do While Not EOF(1)
        Input #1, tstr
        text = "="
        index = InStr(1, tstr, text, vbTextCompare) - 1
        If index > 0 Then
            test = LCase(Trim(Left(tstr, index)))
            If test = name Then
                Return_MapWarp = Trim(Right(tstr, Len(tstr) - index - 1))
                Close #1
                Exit Function
            End If
        End If
    Loop
    Close #1
    Return_MapWarp = "[none]"
End Function
