Attribute VB_Name = "md_Packet"
Option Explicit

Sub pkt_ReqName(AccID As String)
    Winsock_SendPacket Chr(148) & Chr(0) & AccID, True
End Sub
Sub pkt_CreateArrow(eqName As String)
    Dim X As Integer
    X = Find_Item(eqName)
    If X > 0 Then Winsock_SendPacket IntToChr(430) & IntToChr(CInt(Val(AllInv(X).NameID))), True
End Sub
Public Sub pkt_StorageAdd(ByVal Index As Long, Amount As Long)
    Winsock_SendPacket IntToChr(&HF3) & IntToChr(Index) & LngToChr(Amount), True
End Sub
Public Sub pkt_StorageGet(ByVal Index As Long, Amount As Long)
    Winsock_SendPacket IntToChr(&HF5) & IntToChr(Index) & LngToChr(Amount), True
End Sub
Public Sub pkt_StorageClose()
    Winsock_SendPacket IntToChr(&HF7), True
End Sub
Public Sub pkt_CartGet(ByVal Index As Long, Amount As Long)
'S 0126 <index>.w <amount>.l
'inventory > cart
    If Not IsCartOn Then Exit Sub
    Winsock_SendPacket Chr(&H26) & Chr(1) & IntToChr(Index) & IntToChr(Amount) & String(2, Chr(0)), True
End Sub
Public Sub pkt_CartTake(ByVal Index As Long, Amount As Long)
'S 0127 <index>.w <amount>.l
'cart > inventory
    If Not IsCartOn Then Exit Sub
    Winsock_SendPacket Chr(&H27) & Chr(1) & IntToChr(Index) & IntToChr(Amount) & Chr(0) & Chr(0), True
End Sub
Public Sub pkt_CartFromKafra(ByVal Index As Long, Amount As Long)
'S 0128 <index>.w <amount>.l
'kafra > cart
    If Not IsCartOn Then Exit Sub
    Winsock_SendPacket Chr(&H28) & Chr(1) & IntToChr(Index) & IntToChr(Amount) & String(2, Chr(0)), True
End Sub
Public Sub pkt_CartToKafra(ByVal Index As Long, Amount As Long)
'S 0129 <index>.w <amount>.l
'cart > kafra
    If Not IsCartOn Then Exit Sub
    Winsock_SendPacket Chr(&H29) & Chr(1) & IntToChr(Index) & IntToChr(Amount) & String(2, Chr(0)), True
End Sub

