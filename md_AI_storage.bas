Attribute VB_Name = "md_AI_storage"
Option Explicit
Public Storage() As ItemInv
Sub GSNoItemReset()
    Dim i&
    For i = 0 To UBound(GetStorageItem)
        GetStorageItem(i).NoStore = True
    Next
End Sub
Sub GSNoItemTrue(Name As String)
    Dim i&
    For i = 0 To UBound(GetStorageItem)
        If LCase(Name) = LCase(GetStorageItem(i).Name) Then GetStorageItem(i).NoStore = True
    Next
End Sub
Sub GSNoItemFalse(Name As String)
    Dim i&
    For i = 0 To UBound(GetStorageItem)
        If LCase(Name) = LCase(GetStorageItem(i).Name) Then GetStorageItem(i).NoStore = False
    Next
End Sub

Public Function Is_Keep(Name As String) As Boolean
On Error GoTo errie
    If Kafra(0).Name = "" Then GoTo errie
    Dim i As Integer
    For i = 0 To UBound(Kafra)
        If LCase(Name) = LCase(Kafra(i).Name) Then
            Is_Keep = True
            Exit Function
        End If
    Next
errie:
    Is_Keep = False
End Function
Public Function ai_storage_Keepamount(Name As String) As Long
On Error GoTo errie
    If Kafra(0).Name = "" Then GoTo errie
    Dim i As Integer
    For i = 0 To UBound(Kafra)
        If LCase(Name) = LCase(Kafra(i).Name) Then
            ai_storage_Keepamount = Kafra(i).Amount
            Exit Function
        End If
    Next
errie:
End Function
Function Find_StorageID(Name As String) As Integer
On Error GoTo errie
Dim X As Integer
For X = 0 To UBound(Storage)
    If (Storage(X).Amount > 0) And (LCase(Storage(X).Name) = LCase(Name)) Then
        Find_StorageID = X
        Exit Function
    End If
Next
errie:
Find_StorageID = 0
Err.Clear
End Function

'/////////////////////////////////////////////////////////////////
'////////////////// PACKET PARSER ENGINE //////////////////
'/////////////////////////////////////////////////////////////////

Function Decode_00A5(inData As String) As String
On Error GoTo errie
    'R 00a5 <len>.w
    '               3
    '{<index>.w <item ID>.w <type>.B <identify flag>.B <amount>.w ?.2B}.10B*
    '  0                    2                   4               5                               6                       8
    ReDim Storage(0)
    GetStore = False
    Dim i As Integer
    Dim ChopNumber As Long
    Dim STindex&, Itemname$, NameID$, Amount&
    ChopNumber = MakePort(Mid(inData, 3, 2))
    For i = 5 To ChopNumber Step 10
        STindex = MakePort(Mid(inData, i, 2))
        NameID = MakeHexName(Mid(inData, i + 2, 2))
        Itemname = Return_ItemName(NameID)
        Amount = MakePort(Mid(inData, i + 6, 4))
        If STindex > UBound(Storage) Then ReDim Preserve Storage(STindex)
        With Storage(STindex)
            .Amount = Amount
            .Identified = CBool(Asc(Mid(inData, i + 5, 1)))
            .Index = STindex
            .NameID = NameID
            .Name = Itemname
        End With
    Next
    GSNoItemReset
    If AutoAI Then
            Dim GetAmount As Long, CGetAmount As Long
            NoStoreItem = True
            LastGetStorage = GetTickCount()
            For i = 0 To UBound(GetStorageItem)
                Dim X As Integer, Y As Integer
                X = Find_StorageID(GetStorageItem(i).Name)
                Y = Find_Item(GetStorageItem(i).Name)
                If X > 0 Then
                    If Y > 0 Then
                        GetAmount = GetStorageItem(i).Amount - AllInv(Y).Amount
                    Else
                        GetAmount = GetStorageItem(i).Amount
                    End If
                    If Storage(X).Amount > GetAmount Then GSNoItemFalse Storage(X).Name
                    If GetAmount > Storage(X).Amount Then GetAmount = Storage(X).Amount
                    If GetAmount > 0 Then
                        pkt_StorageGet X, GetAmount
                        NoStoreItem = False
                    End If
                    If IsCartWant(Storage(X).Name) Then
                        CGetAmount = CartWantAmount(Storage(X).Name)
                        If (Storage(X).Amount - GetAmount) > CGetAmount Then GSNoItemFalse Storage(X).Name
                        If (Storage(X).Amount - GetAmount) < CGetAmount Then CGetAmount = Storage(X).Amount - GetAmount
                        If CGetAmount > 0 Then
                            pkt_CartFromKafra X, CGetAmount
                            NoStoreItem = False
                        End If
                    End If
                    GetStore = True
                End If
            Next
            Dim KeepAmount As Long
            For i = 0 To UBound(AllInv)
               If AllInv(i).Amount > 0 Then
                    If Is_Keep(AllInv(i).Name) Then
                        KeepAmount = ai_storage_Keepamount(AllInv(i).Name)
                        If AllInv(i).Amount > KeepAmount Then pkt_StorageAdd i, AllInv(i).Amount - KeepAmount
                        SendStore = True
                    End If
               End If
            Next
            SendSell = False
            frmMain.tmrDealNPC.Enabled = False
            frmMain.tmrDealNPC.Enabled = True
    Else
            If UBound(Storage) > 0 Then ReDim Preserve Storage(UBound(Storage) - 1)
            frmStorage.Visible = True
            upd_frmStorage
            GetStore = False
    End If
    Decode_00A5 = ""
    Exit Function
errie:
    Decode_00A5 = "ERROR!!! [Decode_00A5] " & Err.Description
    Err.Clear
End Function

Function Decode_00A6(inData As String) As String
On Error GoTo errie
    'ReDim Storage(0)
    Dim StorageIndex As Long
    Dim i As Integer
    Dim NameID As String
    Dim Itemname As String
    Dim ChopNumber As Long
'R 00a6 <len>.w
'{<index>.w <item ID>.w <type>.B <identify flag>.B <equip type>.w <equip point>.w <attribute?>.B <refine>.B <card>.4w}.20B*
' 0                     2                   4                   5                           6                               8                           10                      11                  12
    ChopNumber = MakePort(Mid(inData, 3, 2))
    For i = 5 To ChopNumber Step 20
        StorageIndex = MakePort(Mid(inData, i, 2))
        NameID = MakeHexName(Mid(inData, i + 2, 2))
        If StorageIndex > UBound(Storage) Then ReDim Preserve Storage(StorageIndex)
        If MakePort(Mid(inData, i + 2, 2)) > 0 Then
            'If Storage(0).Name <> "" Then ReDim Preserve Storage(UBound(Storage) + 1)
            Storage(StorageIndex).Index = StorageIndex
            Storage(StorageIndex).NameID = NameID
            Storage(StorageIndex).Name = MakeItemName(Mid(inData, i + 2, 2), Mid(inData, i + 12, 8), Mid(inData, i + 11, 1))
            Storage(StorageIndex).Identified = CBool(Asc(Mid(inData, 5, 1)))
            Storage(StorageIndex).Type = Asc(Mid(inData, 4, 1))
            Storage(StorageIndex).Amount = 1
        End If
    Next
    If Not AutoAI Then
        frmStorage.Visible = True
        upd_frmStorage
    End If
    Decode_00A6 = ""
    Exit Function
errie:
    Decode_00A6 = "ERROR!!! [Decode_00A6] " & Err.Description
    Err.Clear
End Function
Function Decode_00F4(inData As String) As String
On Error GoTo errie
'R 00f4 <index>.w <amount>.l <type ID>.w <identify flag>.B <attribute?>.B <refine>.B <card>.4w
'               3                   5                   9                       11                          12                      13                  14
    Dim StorageIndex As Integer
    Dim NameID As String
    Dim Itemname As String
    StorageIndex = MakePort(Mid(inData, 3, 2))
    NameID = MakeHexName(Mid(inData, 9, 2))
    Itemname = MakeItemName(Mid(inData, 9, 2), Mid(inData, 14, 8), Mid(inData, 13, 1))
    Dim i As Integer
    Dim Index As Integer
    If StorageIndex > UBound(Storage) Then
        ReDim Preserve Storage(StorageIndex)
        Storage(StorageIndex).Index = StorageIndex
        Storage(StorageIndex).NameID = NameID
        Storage(StorageIndex).Name = Itemname
        Storage(StorageIndex).Amount = MakePort(Mid(inData, 5, 4))
        Storage(StorageIndex).Identified = CBool(Asc(Mid(inData, 11, 1)))
        CheckCartStorage CLng(StorageIndex)
    Else
        'Storage(index).index = StorageIndex
        'Storage(index).Nameid = Nameid
        'Storage(index).name = Itemname
        Storage(StorageIndex).Amount = Storage(StorageIndex).Amount + MakePort(Mid(inData, 5, 4))
    End If
    upd_frmStorage
    Stat "Add [" & Itemname & "] " & CStr(MakePort(Mid(inData, 5, 4))) & " EA to Storage..." & vbCrLf
Decode_00F4 = ""
Exit Function
errie:
Decode_00F4 = "ERROR!!! [Decode_00F4] " & Err.Description
Err.Clear
End Function
Function Decode_00F6(inData As String) As String
On Error GoTo errie
    Dim StorageIndex As Integer
    StorageIndex = MakePort(Mid(inData, 3, 2))
    If UBound(Storage) >= StorageIndex Then
        Storage(StorageIndex).Amount = Storage(StorageIndex).Amount - MakePort(Mid(inData, 5, 4))
    End If
    upd_frmStorage
Decode_00F6 = ""
Exit Function
errie:
Decode_00F6 = "ERROR!!! [Decode_00F6] " & Err.Description
Err.Clear
End Function
Function Decode_01F0(inData As String) As String
On Error GoTo errie
    ReDim Storage(0)
    GetStore = False
    Dim i As Integer
    Dim ChopNumber As Long
    Dim STindex&, Itemname$, NameID$, Amount&
    ChopNumber = MakePort(Mid(inData, 3, 2))
    For i = 5 To ChopNumber Step 18
        STindex = MakePort(Mid(inData, i, 2))
        NameID = MakeHexName(Mid(inData, i + 2, 2))
        Itemname = Return_ItemName(NameID)
        Amount = MakePort(Mid(inData, i + 6, 4))
        If STindex > UBound(Storage) Then ReDim Preserve Storage(STindex)
        With Storage(STindex)
            .Amount = Amount
            .Identified = True
            .Index = STindex
            .NameID = NameID
            .Name = Itemname
        End With
    Next
    GSNoItemReset
    If AutoAI Then
            Dim GetAmount As Long, CGetAmount As Long
            NoStoreItem = True
            LastGetStorage = GetTickCount()
            For i = 0 To UBound(GetStorageItem)
                Dim X As Integer, Y As Integer
                X = Find_StorageID(GetStorageItem(i).Name)
                Y = Find_Item(GetStorageItem(i).Name)
                If X > 0 Then
                    If Y > 0 Then
                        GetAmount = GetStorageItem(i).Amount - AllInv(Y).Amount
                    Else
                        GetAmount = GetStorageItem(i).Amount
                    End If
                    If Storage(X).Amount > GetAmount Then GSNoItemFalse Storage(X).Name
                    If GetAmount > Storage(X).Amount Then GetAmount = Storage(X).Amount
                    If GetAmount > 0 Then
                        pkt_StorageGet X, GetAmount
                        NoStoreItem = False
                    End If
                    If IsCartWant(Storage(X).Name) Then
                        CGetAmount = CartWantAmount(Storage(X).Name)
                        If (Storage(X).Amount - GetAmount) > CGetAmount Then GSNoItemFalse Storage(X).Name
                        If (Storage(X).Amount - GetAmount) < CGetAmount Then CGetAmount = Storage(X).Amount - GetAmount
                        If CGetAmount > 0 Then
                            pkt_CartFromKafra X, CGetAmount
                            NoStoreItem = False
                        End If
                    End If
                    GetStore = True
                End If
            Next
            Dim KeepAmount As Long
            For i = 0 To UBound(AllInv)
               If AllInv(i).Amount > 0 Then
                    If Is_Keep(AllInv(i).Name) Then
                        KeepAmount = ai_storage_Keepamount(AllInv(i).Name)
                        If AllInv(i).Amount > KeepAmount Then pkt_StorageAdd i, AllInv(i).Amount - KeepAmount
                        SendStore = True
                    End If
               End If
            Next
            SendSell = False
            frmMain.tmrDealNPC.Enabled = False
            frmMain.tmrDealNPC.Enabled = True
    Else
            If UBound(Storage) > 0 Then ReDim Preserve Storage(UBound(Storage) - 1)
            frmStorage.Visible = True
            upd_frmStorage
            GetStore = False
    End If
    Decode_01F0 = ""
    Exit Function
errie:
    Decode_01F0 = "ERROR!!! [Decode_01F0] " & Err.Description
    Err.Clear
End Function

