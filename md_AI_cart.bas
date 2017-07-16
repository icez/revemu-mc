Attribute VB_Name = "md_AI_Cart"
Option Explicit

Public Cart() As typeC_CartList
Type typeC_CartList
'    Index As Integer
    ID As Long
    Name As String
    Amount As Long
    Category As Integer
    Identified As Boolean
    Type As String
    Pos As Long
    CheckED As Boolean
End Type

Function Decode_0122(inData As String)
'R 0122 <len>.w
'{<index>.w <item ID>.w <type>.B <identify flag>.B <equip type>.w <equip point>.w <attribute?>.B <refine>.B <card>.4w}.20B*
' 1                     3                    5                  6                           7                               9                              11                       12              13,15,17,19
On Error GoTo errie
    Dim Args As px0122ex
    Dim bArr() As Byte ', Itemname As String
    Dim i&
    For i = 5 To Len(inData) Step 20
        bArr = Conv2Arr(Mid(inData, i, 20))
        CopyMemory Args, bArr(0), 20
        If Args.Index > UBound(Cart) Then ReDim Preserve Cart(Args.Index)
        With Cart(Args.Index)
            .ID = Args.ItemID
            .Amount = 1
            .Category = Args.Attribute
            .Type = Args.EqType
            .Pos = Args.EqPlace
            .Identified = CBool(Args.Identified)
            .Name = MakeItemName(Mid(inData, i + 2, 2), Args.Note, Mid(inData, i + 11, 1))
        End With
        CheckCartInv CLng(Args.Index)
    Next
    IsCartOn = True
    UpdateCart
    CalcModAI "0122"
Exit Function
errie:
Decode_0122 = "ERROR!!! [Decode_0122] " & Err.Description
frmMain.print_packet inData, "0122"
Err.Clear
End Function


