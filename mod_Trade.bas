Attribute VB_Name = "md_Trade"
Option Explicit

Public MTrade() As TYTrade
Public MTradeStep As Byte
Public MTPartner As String
Public MTStatus As TYPartner
Public MZenyPacket As String
Public TOTPrice As Long

Type TYTrade
    ItemID As Long
    Itemname As String
    Amount As Long
    Identified As Boolean
End Type
Type TYPartner
    You As Boolean
    Partner As Boolean
End Type
'R 00e9 <amount>.l <type ID>.w <identify flag>.B <attribute?>.B <refine>.B <card>.4w

'Sub Send_TradeAddItem(id As Integer, amount As Long)
'    Winsock_SendPacket Chr(&HE8) & Chr(0), True
'End Sub

Sub NResetTrade()
    MODTradeStep = 0
    MODTradeDelay = 0
    ReDim MTrade(0)
    MTStatus.Partner = False
    MTStatus.You = False
    MTradeStep = 0
    MTPartner = ""
End Sub

Function IsTradeAccept(ITName As String) As Boolean
    Dim i&
    For i = 0 To UBound(ItemCtrl)
        If LCase(ITName) = ItemCtrl(i).Name And Not ItemCtrl(i).Reject Then
            IsTradeAccept = True
            Exit Function
        End If
    Next
    IsTradeAccept = False
End Function
Function TradeID(ITName As String) As Long
    Dim i&
    For i = 0 To UBound(ItemCtrl)
        If LCase(ITName) = ItemCtrl(i).Name Then
            TradeID = i
            Exit Function
        End If
    Next
    TradeID = -1
End Function

Sub CalcTrade()
    Dim i&
    TOTPrice = 0
    If Mods.OCnocalcmoney Then
        Chat "System : [Trade] No calculate pricing to : " & MTPartner, MColor.trade
        GoTo sendmoney
    End If
    Dim tID&
    For i = 0 To (UBound(MTrade) - 1)
        tID = TradeID(MTrade(i).Itemname)
        If tID >= 0 Then TOTPrice = TOTPrice + ((ItemCtrl(tID).Price) * MTrade(i).Amount)
    Next
    If TOTPrice = 0 Then
        Chat "System : [Trade] no item added. auto-cancel", MColor.trade
        MODTradeStep = 3
        MODTradeDelay = GetTickCount + RandomNumber(3500, 1500)
        Exit Sub
    End If
    If TOTPrice > Players(number).Zeny Then
        Chat "System : [Trade] Not enough zeny (" & TOTPrice & "z). auto-cancel", MColor.trade
        MODTradeStep = 3
        MODTradeDelay = GetTickCount + MODDC.TCalc + RandomNumber(1000, 0)
        Exit Sub
    End If
    Players(number).Zeny = Players(number).Zeny - TOTPrice
    frmMain.UpdatePlayer
    Chat "System : [Trade] Sending money (" & TOTPrice & "z) to " & MTPartner, MColor.trade
sendmoney:
    MZenyPacket = IntToChr(&HE8) & Chr(0) & Chr(0) & LngToChr(TOTPrice)
    MODTradeStep = 6
    MODTradeDelay = GetTickCount + MODDC.TCalc
End Sub

Sub Send_Response(restype As String)
    'tmp
End Sub
