Attribute VB_Name = "md_GetData"
Option Explicit

Function Return_MonsterName(MonsterID As Integer) As String
    On Error GoTo errie
    Dim i As Integer
    For i = 0 To UBound(Monsters)
        If MonsterID = Monsters(i).ID Then
            Return_MonsterName = Monsters(i).Name
            Exit Function
        End If
    Next
    Return_MonsterName = "U:" & CStr(MonsterID)
    Exit Function
errie:
    If Err.number > 0 Then print_funcerr "Return_MonsterName", Err.number, Err.Description
    Err.Clear
End Function

