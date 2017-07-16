Attribute VB_Name = "md_Forms"
Option Explicit
Public Disable_frmPeople As Boolean

Sub UpdateInventory()
On Error GoTo errie
frmItem.lstInvent.Clear
If ConnState < 4 Then Exit Sub
Dim tstr As String
Dim X As Integer
For X = 0 To UBound(AllInv)
    If AllInv(X).Amount > 0 Or Mods.STDebug Then
        With AllInv(X)
            Select Case ViewState
                Case 0
                    'cate 0-2
                    If (.Category < 3) Then frmItem.lstInvent.AddItem CStr(X) + " : [" + AllInv(X).Name + "] " + CStr(AllInv(X).Amount) + " EA  [" & AllInv(X).Category & "]"
                Case 1
                    'cate 4-5,8-9,else
                    If ((.Category > 3 And .Category < 6) Or .Category > 7) And .Category <> 10 Then
                        tstr = CStr(X) + " : [" + AllInv(X).Name + "] " + CStr(AllInv(X).Amount) + " EA  [" & AllInv(X).Category & "]"
                        If .Pos > 0 Then
                            tstr = tstr & " (Equipped)"
                        ElseIf Not .Identified Then
                            tstr = tstr & " (Not Identified)"
                        End If
                        frmItem.lstInvent.AddItem tstr
                    End If
                Case Else
                    '3,6,10
                    If (.Category = 3 Or .Category = 6 Or .Category = 10) Then frmItem.lstInvent.AddItem CStr(X) + " : [" + AllInv(X).Name + "] " + CStr(AllInv(X).Amount) + " EA  [" & AllInv(X).Category & "]"
            End Select
        End With
    End If
Next
Exit Sub
errie:
Err.Clear
'ClearAll
End Sub

Public Sub UpdateCart()
'print_errror "sub UpdateCart"
On Error GoTo errie
    frmCart.lstCart.Clear
    Dim X As Integer
    For X = 0 To UBound(Cart)
        With Cart(X)
            If .Amount > 0 Then
                Dim tstr As String
                tstr = CStr(X) & " : [" & .Name & "] " & CStr(.Amount) & " EA"
                If (.Category > 3 And .Category < 6) Or (.Category = 9 Or .Category = 8) Then
                   If Not (.Identified) Then tstr = tstr & " (Not Identified)"
                End If
                frmCart.lstCart.AddItem tstr
            End If
        End With
    Next
Exit Sub
errie:
Chat "Error in UpdateCart : " & Err.Description + vbCrLf
Err.Clear
End Sub

Sub upd_frmPeople()
On Error GoTo errie
    If Not frmPeople.Visible Or Disable_frmPeople Then Exit Sub
    Dim X&
    
    frmPeople.lstPeople.Clear
    For X = 0 To UBound(People)
        If (Val(People(X).NameID) >= 0) And Len(People(X).Name) > 0 Then
            frmPeople.lstPeople.AddItem "[" & EvalNorm(curPos, People(X).Pos) & " blks] " & People(X).Name & " (" & CStr(People(X).Pos.Y) _
            & ":" & CStr(People(X).Pos.X) & ")" & " - " & People(X).Class & "," _
            & People(X).Sex & ", [" & MakeHex(People(X).ID) & "]"
            '[" & GetHeadItem(People(X).HeadT) & "/" & GetHeadItem(People(X).HeadM) & "/" & GetHeadItem(People(X).HeadB) & "], [E:" & GetHeadItem(People(X).Weapon) & "/" & GetHeadItem(People(X).Shield) & "], [H:" & People(X).Hair & "/" & People(X).HairC & "],
        End If
    Next
Exit Sub
errie:
print_funcerr "upd_frmPeople", Err.number, Err.Description
Err.Clear
End Sub

Sub upd_frmMonster()
On Error GoTo errie
    Dim X As Integer
    Dim tstr As String
    frmMonster.lstMonster.Clear
    If UBound(MonsterList) = 0 Then Exit Sub
    For X = 0 To UBound(MonsterList) - 1
        tstr = ""
        If MyPet.ID = MonsterList(X).ID And MyPet.Name <> "" Then
            tstr = tstr & ", <Your PET>"
        ElseIf MonsterList(X).IsPet Then
            tstr = tstr & ", <PET>"
        Else
            MonsterList(X).NoAttack = Not CanUseAttack(MonsterList(X).Name)
            If MonsterList(X).NoAttack Or (IsSMAgg(MonsterList(X).Name) = False And IsSMR(MonsterList(X).Name) = False) Then tstr = ", [Not Target]"
            If MonsterList(X).StatusA > 0 Then tstr = tstr & ", [" & Get_StatusA(MonsterList(X).StatusA) & "]"
            If MonsterList(X).StatusB > 0 Then tstr = tstr & ", [" & Get_StatusB(MonsterList(X).StatusB) & "]"
            If MonsterList(X).ID = CurAtkMonster.ID Then tstr = tstr & ", [ATK]"
            If isMonsWarp(MonsterList(X).Name) And TeleportDelay = 0 And MonsterList(X).Name <> "" Then
                If CanGO(MonsterList(X).Pos, curPos) Then
                    Stat "Found [" & MonsterList(X).Name & "], Teleport..." & vbCrLf
                    Teleport
                    Exit Sub
                End If
            End If
        End If
        If frmMonster.Visible Then
            If MyPet.ID <> MonsterList(X).ID Then
                frmMonster.lstMonster.AddItem "[" & EvalNorm(curPos, MonsterList(X).Pos) & " blks] " & MonsterList(X).Name & " - (" & CStr(MonsterList(X).Pos.Y) & _
                ":" + CStr(MonsterList(X).Pos.X) & ")" & tstr
            Else
                frmMonster.lstMonster.AddItem "[" & EvalNorm(curPos, MonsterList(X).Pos) & " blks] " & MyPet.Name & " - (" & CStr(MonsterList(X).Pos.Y) & _
                ":" + CStr(MonsterList(X).Pos.X) & ")" & tstr
            End If
        End If
endloop:
    Next
Exit Sub
errie:
    If Err.number > 0 Then print_funcerr "upd_frmMonster", Err.number, Err.Description
    Err.Clear
End Sub
Sub upd_curMonster()
On Error GoTo errie
    If (CurAtkMonster.NameID > 0) Then
        frmMain.labCurMons.Caption = "[" + Return_MonsterName(CurAtkMonster.NameID) + "], " & CStr(EvalNorm(CurAtkMonster.Pos, curPos)) + " Blocks"
    ElseIf (Not IsAggro) And Not Pickup Then
        frmMain.labCurMons.Caption = "[None]"
    End If
Exit Sub
errie:
Err.Clear
'ClearAll
End Sub
Sub upd_frmStorage()
On Error GoTo errie
frmStorage.lstStorage.Clear
Dim X As Integer
For X = 0 To UBound(Storage)
    If Storage(X).Amount > 0 Then
        frmStorage.lstStorage.AddItem CStr(X) & " : " & Storage(X).Name & " " & CStr(Storage(X).Amount) & " EA" & IIf(Not Storage(X).Identified, " [not identified]", "")
    End If
Next
Exit Sub
errie:
Chat "Error in upd_frmStorage" + vbCrLf
Err.Clear
End Sub
Sub UpdateNPC()
On Error GoTo errie
frmNPC.lstNPC.Clear
Dim X As Integer
Dim npccode As String
For X = 0 To UBound(NPCList) - 1
    If (Val(NPCList(X).NameID) >= 0) And NPCList(X).Name <> "" Then
            frmNPC.lstNPC.AddItem "[" & EvalNorm(curPos, NPCList(X).Pos) & " blks] " & NPCList(X).Name & " (" & CStr(NPCList(X).Pos.Y) _
            & ":" & CStr(NPCList(X).Pos.X) & "), [" & MakeHex(Mid(NPCList(X).ID, 1, 4)) & "]"
    End If
Next
Exit Sub
errie:
Chat "Error in UpdateNPC" + vbCrLf
Err.Clear
End Sub

Public Sub UpdatePeople()
On Error GoTo errie
If frmPeople.Visible Then frmPeople.lstPeople.Clear
Dim X As Integer

If UBound(People) = 0 Then Exit Sub

For X = 0 To UBound(People) - 1
    If Asc(Mid(People(X).ID, 3, 1)) > 0 And isWarpAll And Not MoveOnly Then
        Stat "Oh People!!! Leave me alone!, Teleport away..." & vbCrLf
        Teleport
        Exit Sub
    End If
    If (InStr(LCase(JobTele), LCase(People(X).Class)) > 0 Or LCase(JobTele) = "all") And JTele And (UBound(WayPoint) = 0 Or Not MoveOnly) Then
        Chat "Found [" & People(X).Name & "] [" & People(X).Class & "] at (" & People(X).Pos.Y & ":" & People(X).Pos.X & "), Teleport away...", vbRed
        Teleport
        Exit Sub
    End If
'    If (Val(People(X).NameID) >= 0) And Len(People(X).Name) > 0 Then
'        If frmPeople.Visible Then frmPeople.lstPeople.AddItem "[" & EvalNorm(curPos, People(X).pos) & " blks] " & People(X).Name & " (" & CStr(People(X).pos.Y) _
        & ":" & CStr(People(X).pos.X) & ")" & " - " & People(X).Class & "," _
'        & People(X).Sex & ", [" & GetHeadItem(People(X).HeadT) & "/" & GetHeadItem(People(X).HeadM) & "/" & GetHeadItem(People(X).HeadB) & "], [E:" & GetHeadItem(People(X).Weapon) & "/" & GetHeadItem(People(X).Shield) & "], [H:" & People(X).Hair & "/" & People(X).HairC & "] , [" & MakeHex(People(X).ID) & "]"
'    End If
Next
Exit Sub
errie:
Chat "Error in UpdatePeople : " & Err.Description
Err.Clear
End Sub

