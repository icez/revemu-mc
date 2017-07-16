Attribute VB_Name = "Module4"
Type ItemCode
    Name As String
End Type

Type BackCode
    Name As String
    Amount As String
    BackNumber As Integer
End Type

Type ServerCode
    Name As String
    Port As Integer
    IP As String
    code As Byte
    Version As Integer
    start As Integer
    IsLoginCrypt As Boolean
    Encrypt As Integer
    enctype As Integer
    encRequest As String
    pkserver As Byte
End Type

Public ViewState As Byte
Public isWarpAll As Boolean
Public isBackBuy As Boolean
Public MasterSelect As ServerCode
'Public BackItem() As BackCode
Public ROServer() As ServerCode
Public isKillmob As Boolean
Public SendBuy As Boolean
Public SendStore As Boolean

Public Itemlist() As ItemCode
Public PrivateKey As String
Public AvoidList() As ItemCode
Public MonsWarplist As String
Public WarpList() As ItemCode
'Public MobName() As String
'Public MobName2() As String

Public Function Get_Status(ID As Integer) As String
On Error GoTo errie
    Select Case ID
    Case 0
        Get_Status = ""
    Case 1
        Get_Status = "Poisoned"
    Case 4
        Get_Status = "Sleep"
    Case 16
        Get_Status = "Blind"
    Case Else
        Get_Status = CStr(ID)
    End Select
Exit Function
errie:
    Get_Status = ""
End Function

Public Function Get_StatusA(Key As Integer) As String
On Error GoTo errie
    Select Case Key
    Case 1
        Get_StatusA = "Stone Cursed"
    Case 2
        Get_StatusA = "Frozen"
    Case 3
        Get_StatusA = "Dizzy"
    Case 4
        Get_StatusA = "Sleep"
    Case 5
        Get_StatusA = "Immortal"
    Case Else
        Get_StatusA = "StatusA " & CStr(Key)
    End Select
Exit Function
errie:
    Get_StatusA = ""
End Function

Public Function Get_StatusB(Key As Integer) As String
On Error GoTo errie
    Select Case Key
    Case 1
        Get_StatusB = "Envenomed"
    Case 2
        Get_StatusB = "Trapped"
    Case 5
        Get_StatusB = "Silenced"
    Case Else
        Get_StatusB = "StatusB " & CStr(Key)
    End Select
    Exit Function
errie:
    Get_StatusB = ""
End Function

Public Function Get_Emotion_Code(Key As String) As Byte
On Error GoTo errie
If UBound(Emotions) = 0 Then GoTo errie
    Dim i As Integer
    For i = 0 To UBound(Emotions)
        If Emotions(i).Key = Key Then
            Get_Emotion_Code = i
            Exit Function
        End If
    Next
errie:
    MsgBox "Error!!! on loading 'table\emotions.txt'", vbCritical
    Get_Emotion_Code = UBound(Emotions) + 5
End Function

Public Sub Load_Emotion()
On Error GoTo errie:
    Dim tstr As String
    Dim Index As Integer
    Dim number As Integer
    Open App.Path & "\table\emotions.txt" For Input As #10
    Do While Not EOF(10)
        Line Input #10, tstr
        Index = InStr(tstr, "/")
        If Index > 0 Then
            number = Val("&H" & Trim(Left(tstr, Index - 1)))
            ReDim Preserve Emotions(number)
            tstr = Right(tstr, Len(tstr) - Index + 1)
            Index = InStr(tstr, "#")
            Emotions(number).Key = Trim(Left(tstr, Index - 1))
            Emotions(number).detail = Trim(Right(tstr, Len(tstr) - Index))
        End If
    Loop
errie:
    Close 10
End Sub

Public Sub Update_FrmItem()
    Select Case ViewState
        Case 0
            frmItem.Picture1.Picture = LoadPicture(App.Path & "\interface\item_bar.gif")
        Case 1
            frmItem.Picture1.Picture = LoadPicture(App.Path & "\interface\equip_bar.gif")
        Case 2
            frmItem.Picture1.Picture = LoadPicture(App.Path & "\interface\etc_bar.gif")
    End Select
End Sub

Public Function GetSkillName(code As Integer) As String
On Error GoTo errie:
    Dim tstr As String
    Open App.Path & "\table\skillname.txt" For Input As #10
    For i = 1 To code
        Line Input #10, tstr
        If i = code Then
            GetSkillName = Trim(Mid(tstr, InStr(tstr, "#") + 1, Len(tstr) - InStr(tstr, "#") - 1))
            Close 10
            Exit Function
        End If
    Next
errie:
    Close 10
    GetSkillName = CStr(code)
End Function

Public Sub Load_Server()
On Error GoTo errie:
Open App.Path & "\table\server.txt" For Input As #10
Dim tstr As String
Dim Index As Integer
ReDim ROServer(0)
ROServer(0).Port = 5000
ROServer(0).code = &HFF
ROServer(0).Version = &HFF
Do While Not EOF(10)
    Line Input #10, tstr
    Index = InStr(tstr, "#")
    If Index > 0 Then
        If ROServer(0).Name <> "" Then
            With ROServer(UBound(ROServer))
                If .IP = "" Or .Port = 5000 Or .code = &HFF Or .Version = &HFF Then MsgBox "Need to properly config server in 'table/server.txt'", vbCritical
            End With
            ReDim Preserve ROServer(UBound(ROServer) + 1)
            With ROServer(UBound(ROServer))
                .code = &HFF
                .Port = 5000
                .Version = &HFF
            End With
        End If
    Else
        Index = InStr(tstr, "=")
        If Index > 0 Then
            Select Case LCase(Trim(Left(tstr, Index - 1)))
            Case "name"
                ROServer(UBound(ROServer)).Name = Trim(Right(tstr, Len(tstr) - Index))
            Case "ip"
                ROServer(UBound(ROServer)).IP = Trim(Right(tstr, Len(tstr) - Index))
            Case "port"
                ROServer(UBound(ROServer)).Port = Val(Trim(Right(tstr, Len(tstr) - Index)))
            Case "pkserver"
                ROServer(UBound(ROServer)).pkserver = Val(Trim(Right(tstr, Len(tstr) - Index)))
            Case "code"
                ROServer(UBound(ROServer)).code = Val("&H" & Trim(Right(tstr, Len(tstr) - Index)))
            Case "version"
                ROServer(UBound(ROServer)).Version = Val("&H" & Trim(Right(tstr, Len(tstr) - Index)))
            Case "encrypt_request"
                ROServer(UBound(ROServer)).encRequest = Trim(Right(tstr, Len(tstr) - Index))
            Case "char_data_start"
                ROServer(UBound(ROServer)).start = Val(Trim(Right(tstr, Len(tstr) - Index)))
            Case "encrypt_login"
                If Val(Trim(Right(tstr, Len(tstr) - Index))) > 0 Then
                    ROServer(UBound(ROServer)).IsLoginCrypt = True
                    ROServer(UBound(ROServer)).Encrypt = Val(Trim(Right(tstr, Len(tstr) - Index)))
                Else
                    ROServer(UBound(ROServer)).IsLoginCrypt = False
                End If
            Case "encrypt_login_code"
                ROServer(UBound(ROServer)).enctype = Val(Trim(Right(tstr, Len(tstr) - Index)))
            End Select
        End If
    End If
    With ROServer(UBound(ROServer))
        If .IP <> "" And .Port <> 5000 And .code <> &HFF And .Version <> &HFF Then
            If .Name = MasterSelect.Name Then MasterSelect = ROServer(UBound(ROServer))
        End If
        
    End With
Loop
Close 10
Exit Sub
errie:
Close 10
MsgBox "Error!!! on loading 'table\server.txt'", vbCritical
Unload MDIfrmMain
End Sub


Public Function Get_Element(code As Byte) As String
    Select Case code
        Case 0
            Get_Element = ""
        Case 1
            Get_Element = "Ice "
        Case 2
            Get_Element = "Earth "
        Case 3
            Get_Element = "Fire "
        Case 4
            Get_Element = "Wind "
        Case Else
            Get_Element = "Unknow [" & CStr(code) & "] "
    End Select
End Function

Public Function CloseAnyPlayer(pos1 As Coord, pos2 As Coord) As Boolean
Dim i As Integer
If People(0).Pos.X = 0 Then GoTo EndFunc
    For i = 0 To UBound(People) - 1
        If EvalNorm(pos2, People(i).Pos) < noAtkRange Then 'Or EvalNorm(pos2, People(i).pos) < 2 Then
            CloseAnyPlayer = True
            Exit Function
        End If
    Next
EndFunc:
CloseAnyPlayer = False
End Function

Public Function Is_Buy(Name As String) As Integer
On Error GoTo errie
    'If UBound(BuyItem) = 0 Then GoTo errie
    Dim i As Integer
    For i = 0 To UBound(BuyItem)
        If Name = BuyItem(i).Name Then
            Is_Buy = BuyItem(i).Amount
            Exit Function
        End If
    Next
errie:
    Is_Buy = 0
End Function

Public Function Is_Sell(Name As String) As Boolean
On Error GoTo errie
    'If UBound(SelItem) = 0 Then GoTo errie
    Dim i As Integer
    For i = 0 To UBound(SelItem)
        If Name = SelItem(i).Name Then
            Is_Sell = True
            Exit Function
        End If
    Next
errie:
    Is_Sell = False
End Function

Public Function GetEncodePos(Coords As Coord) As String
    Dim tstr As String
    Dim test As String
    Dim X As Integer
    tstr = tstr + Chr(Int((Coords.Y) / 4))
    tstr = tstr + Chr(((Coords.Y) Mod 4) * 64 + Int((Coords.X) / 16))
    tstr = tstr + Chr(((Coords.X) Mod 16) * 16)
    For X = 1 To Len(tstr)
       If Asc(Mid(tstr, X, 1)) < 16 Then test = test + "0"
       test = test + Hex(Asc(Mid(tstr, X, 1))) + " "
       'If x Mod 16 = 0 Then test = test & vbCrLf
    Next
    GetEncodePos = test
End Function

Public Function MakeCoords(rawCoords As String) As Coord
''print_errror "sub MakeCoords"
On Error GoTo Out
Dim xint As Long
Dim yint As Long
yint = Asc(Mid(rawCoords, 1, 1)) * 4
yint = yint + (Asc(Mid(rawCoords, 2, 1)) And &HC0) / 64
xint = (Asc(Mid(rawCoords, 2, 1)) And &H3F) * 16
xint = xint + (Asc(Mid(rawCoords, 3, 1)) And &HF0) / 16
MakeCoords.Y = yint
MakeCoords.X = xint
Exit Function
Out:
MakeCoords.Y = 0
MakeCoords.X = 0
End Function

Public Function Return_CardNameTable(CardName As String) As String
On Error GoTo errie:
Open App.Path & "\table\cardprefixnametable.txt" For Input As #10
Dim tstr As String
Dim Index As Integer

Do While Not EOF(10)
    Input #10, tstr
    Index = InStr(tstr, "#")
    If Index > 0 Then
        If CardName = Left(tstr, Index - 1) Then
            Close 10
            Return_CardNameTable = Mid(tstr, Index + 1, Len(tstr) - Index - 1)
            Exit Function
        End If
    End If
Loop
errie:
Return_CardNameTable = ""
Close 10
End Function

Public Function IsAvoid(Name As String) As Boolean
Dim found As Boolean
Dim i As Long
found = False
If UBound(AvoidList) = 0 Then GoTo EndFunc
    For i = 0 To UBound(AvoidList)
        If InStr(LCase(Name), LCase(AvoidList(i).Name)) > 0 Then
            found = True
            Exit For
        End If
    Next
EndFunc:
    IsAvoid = found
End Function

Public Function isWarpList(Name As String) As Boolean
Dim found As Boolean
Dim i As Integer
If WarpList(0).Name = "" Then GoTo EndFunc
    For i = 0 To UBound(WarpList)
        If InStr(Name, WarpList(i).Name) > 0 Then
            isWarpList = True
            Exit Function
        End If
    Next
EndFunc:
    isWarpList = False
End Function

Public Function isMonsWarp(Name As String) As Boolean
Dim ispl() As String, i&
ispl = Split(MonsWarplist, Chr(0))
    
For i = 0 To UBound(ispl)
    If LCase(Name) = ispl(i) Then
        isMonsWarp = True
        Exit Function
    End If
Next
    isMonsWarp = False
End Function

Public Sub Load_Warplist()
On Error GoTo Out
Dim tstr As String
ReDim WarpList(0)
'Close 1
Open App.Path & "\avoid\warplist.txt" For Input As #1
Do While Not EOF(1)
    Input #1, tstr
    WarpList(UBound(WarpList)).Name = tstr
    ReDim Preserve WarpList(UBound(WarpList) + 1)
Loop
ReDim Preserve WarpList(UBound(WarpList) - 1)
Close 1
Exit Sub
Out:
Close 1
MsgBox "Error!!! on loading 'avoid\warplist.txt'", vbCritical
'Unload MDIfrmMain
End Sub

Public Sub Load_Monswarplist()
On Error GoTo Out
Dim tstr As String
MonsWarplist = ""
'Close 1
Open App.Path & "\avoid\monswarplist.txt" For Input As #1
Do While Not EOF(1)
    Line Input #1, tstr
    If MonsWarplist = "" Then MonsWarplist = tstr Else MonsWarplist = MonsWarplist & Chr(0) & tstr
Loop
Close 1
MonsWarplist = LCase(MonsWarplist)
Exit Sub
Out:
Close 1
MsgBox "Error!!! on loading 'avoid\monswarplist.txt'", vbCritical
End Sub

Public Sub Load_Avoidlist()
On Error GoTo Out
Dim tstr As String
ReDim AvoidList(0)
'Close 1
Open App.Path & "\avoid\avoidlist.txt" For Input As #1
Do While Not EOF(1)
    Line Input #1, tstr
    AvoidList(UBound(AvoidList)).Name = tstr
    ReDim Preserve AvoidList(UBound(AvoidList) + 1)
Loop
ReDim Preserve AvoidList(UBound(AvoidList) - 1)
Close 1
Exit Sub
Out:
Close 1
MsgBox "Error!!! on loading 'avoid\avoidlist.txt'", vbCritical
'Unload MDIfrmMain
End Sub

Public Sub Load_Rarelist()
On Error GoTo Out
Dim tstr As String
ReDim RareItem(0)
'Close 1
Open App.Path & "\control\rarelist.txt" For Input As #1
Do While Not EOF(1)
    Line Input #1, tstr
    RareItem(UBound(RareItem)).Name = tstr
    ReDim Preserve RareItem(UBound(RareItem) + 1)
    
Loop
ReDim Preserve RareItem(UBound(RareItem) - 1)
Close 1
Exit Sub
Out:
Close 1
MsgBox "Error!!! on loading 'control\rarelist.txt'", vbCritical
End Sub

Public Sub Load_Droplist()
On Error GoTo Out
Dim tstr As String
ReDim Itempick(0)
'Close 1
Open App.Path & "\control\droplist.txt" For Input As #1
Do While Not EOF(1)
    Line Input #1, tstr
    Dim Index As Integer
    Index = InStr(tstr, "=")
    If Index > 0 Then
        Itempick(UBound(Itempick)).Amount = Val(Trim(Right(tstr, Len(tstr) - Index)))
        Itempick(UBound(Itempick)).Name = Trim(Left(tstr, Index - 1))
    Else
        Itempick(UBound(Itempick)).Name = tstr
    End If
    ReDim Preserve Itempick(UBound(Itempick) + 1)
    
Loop
ReDim Preserve Itempick(UBound(Itempick) - 1)
Close 1
Exit Sub
Out:
Close 1
MsgBox "Error!!! on loading 'control\droplist.txt'", vbCritical
End Sub

Public Function Return_Raw_Skill(ID As Byte) As String
    Open App.Path & "\table\Skillname.txt" For Input As #1
    Dim i As Integer
    Dim tstr As String
    Dim text As String
    Dim Index As Integer
    text = "#"
    For i = 1 To ID
    Input #1, tstr
    Next
    Index = InStr(1, tstr, text, vbTextCompare)
    If Index > 0 Then
        Return_Raw_Skill = Left(tstr, Index - 1)
    End If
    Close 1
    Exit Function
End Function

Function isClassMage() As Boolean
On Error GoTo errie
    Select Case Players(number).ClassID
        Case 2, 9, 16, 4003, 163, 4010, 170, 4017, 177
            isClassMage = True
        Case Else
            isClassMage = False
    End Select
errie:
    Err.Clear
    isClassMage = False
End Function
Function isClassAco() As Boolean
On Error GoTo errie
    Select Case Players(number).ClassID
        Case 4, 8, 15, 4005, 165, 4009, 169, 4016, 176
            isClassAco = True
        Case Else
            isClassAco = False
    End Select
errie:
    Err.Clear
    isClassAco = False
End Function
Function isClassArcher() As Boolean
On Error GoTo errie
    Select Case Players(number).ClassID
        Case 3, 11, 19, 20, 4004, 164, 4013, 173, 4020, 4021, 180, 181
            isClassArcher = True
        Case Else
            isClassArcher = False
    End Select
errie:
    Err.Clear
    isClassArcher = False
End Function

Public Function Return_Class(ID As Long) As String
    Select Case ID
    Case 0
        Return_Class = "Novice"
    Case 1
        Return_Class = "Swordsman"
    Case 2
        Return_Class = "Mage"
    Case 3
        Return_Class = "Archer"
    Case 4
        Return_Class = "Acolyte"
    Case 5
        Return_Class = "Merchant"
    Case 6
        Return_Class = "Thief"
    Case 7
        Return_Class = "Knight"
    Case 8
        Return_Class = "Priest"
    Case 9
        Return_Class = "Wizard"
    Case 10
        Return_Class = "Blacksmith"
    Case 11
        Return_Class = "Hunter"
    Case 12
        Return_Class = "Assassin"
    Case 13
        Return_Class = "KnightP"
    Case 14
        Return_Class = "Crusader"
    Case 15
        Return_Class = "Monk"
    Case 16
        Return_Class = "Sage"
    Case 17
        Return_Class = "Rogue"
    Case 18
        Return_Class = "Alchemist"
    Case 19
        Return_Class = "Bard"
    Case 20
        Return_Class = "Dancer"
    Case 21
        Return_Class = "CrusaderP"
    Case 22
        Return_Class = "Wedding"
    Case 23
        Return_Class = "Super Novice"
    Case 4001, 161
        Return_Class = "Novice High"
    Case 4002, 162
        Return_Class = "Swordsman High"
    Case 4003, 163
        Return_Class = "Mage High"
    Case 4004, 164
        Return_Class = "Archer High"
    Case 4005, 165
        Return_Class = "Acolyte High"
    Case 4006, 166
        Return_Class = "Merchant High"
    Case 4007, 167
        Return_Class = "Thief High"
    Case 4008, 168
        Return_Class = "Lord Knight"
    Case 4009, 169
        Return_Class = "High Priest"
    Case 4010, 170
        Return_Class = "High Wizard"
    Case 4011, 171
        Return_Class = "Whitesmith"
    Case 4012, 172
        Return_Class = "Sniper"
    Case 4013, 173
        Return_Class = "Assassin Cross"
    Case 4014, 174
        Return_Class = "Lord KnightP"
    Case 4015, 175
        Return_Class = "Paladin"
    Case 4016, 176
        Return_Class = "Champion"
    Case 4017, 177
        Return_Class = "Professor"
    Case 4018, 178
        Return_Class = "Stalker"
    Case 4019, 179
        Return_Class = "Creator"
    Case 4020, 180
        Return_Class = "Clown"
    Case 4021, 181
        Return_Class = "Gypsy"
    Case Else
        Return_Class = "Unknown_" & STR(ID)
    End Select

End Function

Public Function RandomNumber(Upper As Long, _
    Lower As Long) As Long
  On Error GoTo LocalError
  'Generates a Random Number BETWEEN then LOWER
  'and UPPER values
  Randomize
  RandomNumber = ((Upper - Lower + 1) * Rnd + Lower)
  Exit Function
LocalError:
  RandomNumber = 1
End Function

Public Function RandomVal(Min As Long, Max As Long) As Long
    Randomize
    RandomVal = Int(Rnd() * (Max - Min + 1)) + Min
End Function

Public Sub Load_Item()
On Error GoTo Out:
'DESDecryptFile "table\data.grf", "124589", "table\data.txt"
Open App.Path & "\table\data.txt" For Input As #1
Dim tstr As String
Dim text As String
Dim Index As Integer
Dim index2 As Integer
Dim num As Long
ReDim Itemlist(0)
text = "#"
Do While Not EOF(1) And tstr <> "[ItemData]"
    Line Input #1, tstr
Loop
Do While Not EOF(1) And (tstr <> "[MonsterData]")
    Line Input #1, tstr
    If tstr = "[MonsterData]" Then Exit Do
    Index = InStr(1, tstr, text, vbTextCompare)
    index2 = InStr(Index + 1, tstr, text, vbTextCompare)
    If Index > 0 Then
            num = Val(Left(tstr, Index - 1)) - 501
            If num > UBound(Itemlist) Then ReDim Preserve Itemlist(num)
            Itemlist(num).Name = Mid(tstr, Index + 1, index2 - Index - 1)
    End If
Loop
Close 1
Exit Sub
Out:
Close 1
MsgBox "Error!!! on loading 'table\data.txt' (Item Data)", vbCritical
End Sub

Public Sub Load_Monster()
On Error GoTo errie
Dim tstr As String
Dim X As Integer
ReDim Monsters(0)
'Close 1
Open App.Path & "\table\data.txt" For Input As #1
Input #1, tstr
Do While tstr <> "[MonsterData]" And Not (EOF(1))
    Line Input #1, tstr
Loop
Do While Not EOF(1)
    Line Input #1, tstr
    If Len(tstr) > 3 Then
        If Val("&H" + Left(tstr, 4)) > 0 Then
            Monsters(UBound(Monsters)).ID = Val("&H" + Left(tstr, 4))
            Monsters(UBound(Monsters)).Name = Right(tstr, Len(tstr) - 5)
            ReDim Preserve Monsters(UBound(Monsters) + 1)
        End If
    End If
Loop
ReDim Preserve Monsters(UBound(Monsters) - 1)
Close 1
Exit Sub
errie:
Close 1
MsgBox "Error!!! on loading 'table\data.txt (Monster Data)' : " & Err.Description, vbCritical
End Sub

Public Function return_monsid(MonsName As String) As Integer
Dim X As Integer
    For X = 0 To UBound(Monsters)
        If MonsName = Monsters(X).Name Then
            return_monsid = Monsters(X).ID
            Exit Function
        End If
    Next
    return_monsid = 0
End Function

Public Function Return_ItemName(itemName As String) As String
On Error GoTo endsub
    Dim Index As Integer
    Index = Val("&H" & itemName)
    If Itemlist(Index - 501).Name <> "" Then
        Return_ItemName = Itemlist(Index - 501).Name
    Else
        Return_ItemName = CStr(Index)
    End If
    Exit Function
endsub:
Return_ItemName = itemName
End Function

'Public Function Return_ItemD(Itemname As String) As String
'On Error GoTo Out:
'Dim tstr As String
'Dim text As String
'Dim Index As Integer
'text = "#"
'Open App.Path & "\table\item_detail.txt" For Input As #1
'Do While Not EOF(1)
'    Input #1, tstr
'    Index = InStr(1, tstr, text, vbTextCompare)
'    If Index > 0 Then
'        If Left(tstr, Index - 1) = Itemname Then
'            Input #1, tstr
'            Do
'            Return_ItemD = Return_ItemD & tstr
'            Input #1, tstr
'            If (tstr <> "#") Then Return_ItemD = Return_ItemD & vbCrLf
'            Loop While (tstr <> "#")
'            Close 1
'            Exit Function
'        End If
'    End If
'Loop
'endsub:
'Return_ItemD = "Nothing"
'Close 1
'Exit Function
'Out:
'Close 1
'MsgBox "Error!!! on loading 'table\item_detail.txt'", vbCritical
'End Function

Public Sub Load_Attack()
On Error GoTo Out:
Dim text As String
Dim Index As Integer
Dim index2 As Integer
Dim tstr As String
Dim count As Integer
Dim blens As Long
Dim i, Y As Integer
ReDim Attack(0)
Open App.Path & "\control\Attack.txt" For Input As #1
Do While Not EOF(1)
    Line Input #1, tstr
    If Left(tstr, 1) <> "/" Then
        Index = InStr(tstr, " - ")
        index2 = 0
        If Index = 0 Then
            Attack(UBound(Attack)).Name = Trim(tstr)
            Attack(UBound(Attack)).ID = return_monsid(Trim(tstr))
        End If
        Do While Index > 0
            text = Trim(Left(tstr, Index))
            tstr = Trim(Right(tstr, Len(tstr) - Index))
            If text = "-" Then GoTo endloop
            Select Case index2
            Case 0
                Attack(UBound(Attack)).Name = text
                Attack(UBound(Attack)).ID = return_monsid(text)
            Case 1
                blens = Len(text)
                text = Replace(text, "/", "")
                Attack(UBound(Attack)).Spell1 = text
                Attack(UBound(Attack)).UTime1 = blens - Len(text)
                If InStr(tstr, " ") = 0 Then Attack(UBound(Attack)).lv1 = Val(tstr)
            Case 2
                Attack(UBound(Attack)).lv1 = Val(text)
            Case 3
                blens = Len(text)
                text = Replace(text, "/", "")
                Attack(UBound(Attack)).Spell2 = text
                Attack(UBound(Attack)).UTime2 = blens - Len(text)
                Attack(UBound(Attack)).lv2 = Val(tstr)
            End Select
            index2 = index2 + 1
endloop:
            Index = InStr(tstr, " ")
        Loop
            ReDim Preserve Attack(UBound(Attack) + 1)
    End If
Loop
ReDim Preserve Attack(UBound(Attack) - 1)
Close 1
If UBound(SkillChar) > 0 Then Update_AtkSkill
Exit Sub
Out:
Close 1
MsgBox "Error!!! on loading 'control\attack.txt' : " & Err.Description, vbCritical
End Sub
