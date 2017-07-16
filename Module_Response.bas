Attribute VB_Name = "Mod_Response"
Public mons_heal() As String
Public mons_jam() As String
Public mons_agi() As String
Public mons_bless() As String
Public mons_heal_response As Boolean
Public mons_jam_response As Boolean
Public mons_agi_response As Boolean
Public mons_bless_response As Boolean
Public response_mode As Byte
Public delay_response As Byte

Public Sub Load_Response()
    Dim tstr As String
    ReDim mons_heal(0)
    ReDim mons_jam(0)
    ReDim mons_agi(0)
    ReDim mons_bless(0)
    Open App.Path & "\control\response.txt" For Input As #1
    Do While Not EOF(1)
        Input #1, tstr
        index = InStr(tstr, "#")
        If index > 0 Then
            Select Case LCase(Trim(Left(tstr, index - 1)))
                Case "mons_heal_response"
                    mons_heal_response = CBool(Val(Trim(Right(tstr, Len(tstr) - index))))

                Case "mons_jam_response"
                    mons_jam_response = CBool(Val(Trim(Right(tstr, Len(tstr) - index))))
                    
                Case "mons_agi_response"
                    mons_agi_response = CBool(Val(Trim(Right(tstr, Len(tstr) - index))))
                    
                Case "mons_bless_response"
                    mons_bless_response = CBool(Val(Trim(Right(tstr, Len(tstr) - index))))
                    
                Case "mons_heal"
                    tstr = Trim(Right(tstr, Len(tstr) - index))
                    mons_heal(UBound(mons_heal)) = tstr
                    ReDim Preserve mons_heal(UBound(mons_heal) + 1)
                    
                Case "mons_agi"
                    tstr = Trim(Right(tstr, Len(tstr) - index))
                    mons_agi(UBound(mons_agi)) = tstr
                    ReDim Preserve mons_agi(UBound(mons_agi) + 1)
                    
                Case "mons_bless"
                    tstr = Trim(Right(tstr, Len(tstr) - index))
                    mons_bless(UBound(mons_bless)) = tstr
                    ReDim Preserve mons_bless(UBound(mons_bless) + 1)
                    
                Case "mons_jam"
                    tstr = Trim(Right(tstr, Len(tstr) - index))
                    mons_jam(UBound(mons_jam)) = tstr
                    ReDim Preserve mons_jam(UBound(mons_jam) + 1)
                    
            End Select
        End If
    Loop
    Close 1
    If UBound(mons_heal) > 0 Then ReDim Preserve mons_heal(UBound(mons_heal) - 1)
    If UBound(mons_jam) > 0 Then ReDim Preserve mons_jam(UBound(mons_jam) - 1)
    If UBound(mons_agi) > 0 Then ReDim Preserve mons_agi(UBound(mons_agi) - 1)
    If UBound(mons_bless) > 0 Then ReDim Preserve mons_bless(UBound(mons_bless) - 1)
End Sub

Public Function Random_Mons_Heal_Message() As String
    If mons_heal(0) = "" Or Not mons_heal_response Then Exit Function
    Random_Mons_Heal_Message = mons_heal(RandomNumber(0, UBound(mons_heal)))
End Function

Public Function Random_Mons_Jam_Message() As String
    If mons_jam(0) = "" Or Not mons_jam_response Then Exit Function
    Random_Mons_Jam_Message = mons_jam(RandomNumber(0, UBound(mons_jam)))
End Function

Public Function Random_Mons_Agi_Message() As String
    If mons_agi(0) = "" Or Not mons_agi_response Then Exit Function
    Random_Mons_Agi_Message = mons_agi(RandomNumber(0, UBound(mons_agi)))
End Function

Public Function Random_Mons_Bless_Message() As String
    If mons_bless(0) = "" Or Not mons_bless_response Then Exit Function
    Random_Mons_Bless_Message = mons_bless(RandomNumber(0, UBound(mons_bless)))
End Function

