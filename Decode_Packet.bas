Attribute VB_Name = "Decode_Packet"
Public Const MageRange As Integer = 8
Public Const ArcherRange As Integer = 6
Public Const SMRange As Integer = 3
Public Const ThiefRange As Integer = 2
Public Const TimePickup As Integer = 800
Public Const TimeTick As Integer = 1000
Public Const TimeAnswer As Long = 25000
Public Const PColor As Long = &HEFAE00
Public Const CurAtkColor As Long = &H1D94F7

Sub SendClient(inData As String)
    If Winsock1.State <> 7 Then Exit Sub
    frmMain.Winsock1.SendData Chr(&H52) & IntToChr(Len(inData)) & inData
    If MDIfrmMain.mnuPKTLOG.CheckED Then
        Open App.Path & "\packet.txt" For Append As #76
            Print #76, "Recv : "
            Print #76, ConvPacketData(inData)
            Print #76, ""
        Close #76
    End If
End Sub

Public Sub Winsock_SendPacket(pkt As String, pass As Boolean)
On Error GoTo Out:
    Dim Packet As String
    Packet = pkt

    'If UsingSelfSkill Then Exit Sub
    If Mid(Packet, 1, 2) = IntToChr(&H89) And Not AutoAI Then Packet = Right(Packet, Len(Packet) - 7)
    If Len(Packet) < 1 Then Exit Sub
    If (frmMain.Winsock1.State = 7) And (Not TraceMons And Not Sending Or pass) Then
        If isUseHaunted Then
            frmMain.Winsock1.SendData Chr(&H53) & IntToChr(Len(Packet)) & Packet
        Else
            frmMain.Winsock1.SendData Packet
        End If
        If MDIfrmMain.mnuPKTLOG.CheckED Then
            Open App.Path & "\packet.txt" For Append As #76
                Print #76, "Send : "
                Print #76, ConvPacketData(Packet)
                Print #76, ""
            Close #76
        End If
    End If
Out:
Err.Clear
End Sub

