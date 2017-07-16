VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMapRoute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Point
    X As Long
    Y As Long
End Type

Private Type Map
    Name As String
    Pos As Point
End Type

Private Type Portal
    Src As Map
    Des As Map
End Type

Private Type Link
    NodeIndex As Portal
    f As Long
    g As Long
    h As Long
End Type

Private Type Node
    Current As Link
    Parent As Link
End Type

Private Portals() As Portal
Private ExcludeMap() As String
Private ExcludeLoaded As Boolean
Private WarpPoint() As Portal
Private Sucessor() As Portal
Private OpenList() As Node
Private CloseList() As Node
Private NodeCurrent As Node
Private MySearch As Portal
Private Finish As Boolean
Private AllSearch() As Node
Private Solutions() As Portal
Private UseNPCWarp As Boolean

Public Function MapStepNext(Index As Integer, ByRef SName As String, ByRef Sx As Long, ByRef Sy As Long, ByRef DName As String, ByRef Dx As Long, ByRef Dy As Long) As Boolean
    Dim continue As Boolean
    continue = False
    If Index > UBound(Solutions) Then
        GoTo end_func
    Else
        If Index < UBound(Solutions) Then continue = True
        SName = Solutions(UBound(Solutions) - Index).Src.Name
        Sx = Solutions(UBound(Solutions) - Index).Src.Pos.X
        Sy = Solutions(UBound(Solutions) - Index).Src.Pos.Y
        DName = Solutions(UBound(Solutions) - Index).Des.Name
        Dx = Solutions(UBound(Solutions) - Index).Des.Pos.X
        Dy = Solutions(UBound(Solutions) - Index).Des.Pos.Y
    End If
end_func:
MapStepNext = continue
End Function

Private Sub Remove_List(ByVal Index As Long, ByRef List() As Node)
    Dim X As Long
    For X = Index To UBound(List) - 1
        List(X) = List(X + 1)
    Next
    If UBound(List) > 0 Then ReDim Preserve List(UBound(List) - 1)
End Sub

Private Sub Get_Successor()
    Dim i As Long
    If Not ExcludeLoaded Then Load_ExcludeMap
    ReDim Sucessor(0)
    For i = 0 To UBound(Portals)
        
        If (Portals(i).Src.Name = NodeCurrent.Current.NodeIndex.Des.Name And _
        Not (NodeCurrent.Current.NodeIndex.Src.Name = Portals(i).Des.Name And _
             NodeCurrent.Current.NodeIndex.Src.Pos.X = Portals(i).Des.Pos.X And _
             NodeCurrent.Current.NodeIndex.Src.Pos.Y = Portals(i).Des.Pos.Y) And _
        Not (NodeCurrent.Parent.NodeIndex.Src.Name <> "" And _
             NodeCurrent.Parent.NodeIndex.Src.Name = Portals(i).Des.Name And _
             NodeCurrent.Parent.NodeIndex.Src.Pos.X = Portals(i).Des.Pos.X And _
             NodeCurrent.Parent.NodeIndex.Src.Pos.Y = Portals(i).Des.Pos.Y)) And _
             Not IsExclude(Portals(i).Des.Name) Then
                'If (Portals(i).Des.Name <> Portals(i).Src.Name) Then
                    If Sucessor(0).Src.Name = "" Then
                        Sucessor(0) = Portals(i)
                    Else
                        ReDim Preserve Sucessor(UBound(Sucessor) + 1)
                        Sucessor(UBound(Sucessor)) = Portals(i)
                    End If
                'End If
        End If
    Next
    
    'To use teleport from kafra
    If UseNPCWarp Then
    
    For i = 0 To UBound(WarpPoint)
        If (WarpPoint(i).Src.Name = NodeCurrent.Current.NodeIndex.Des.Name And _
        Not (NodeCurrent.Current.NodeIndex.Src.Name = WarpPoint(i).Des.Name And _
             NodeCurrent.Current.NodeIndex.Src.Pos.X = WarpPoint(i).Des.Pos.X And _
             NodeCurrent.Current.NodeIndex.Src.Pos.Y = WarpPoint(i).Des.Pos.Y) And _
        Not (NodeCurrent.Parent.NodeIndex.Src.Name <> "" And _
             NodeCurrent.Parent.NodeIndex.Src.Name = WarpPoint(i).Des.Name And _
             NodeCurrent.Parent.NodeIndex.Src.Pos.X = WarpPoint(i).Des.Pos.X And _
             NodeCurrent.Parent.NodeIndex.Src.Pos.Y = WarpPoint(i).Des.Pos.Y)) Then
                    If Sucessor(0).Src.Name = "" Then
                        Sucessor(0) = WarpPoint(i)
                    Else
                        ReDim Preserve Sucessor(UBound(Sucessor) + 1)
                        Sucessor(UBound(Sucessor)) = WarpPoint(i)
                    End If
        End If
    Next
    
    End If
End Sub

Private Sub Search_Step()
    Dim i, j As Integer
    Dim myindex As Integer
    Dim newg As Integer
    myindex = 0
    For i = 0 To UBound(OpenList)
        If OpenList(i).Current.NodeIndex.Des.Name = MySearch.Des.Name Then
            myindex = i
            Exit For
        End If
    Next
    NodeCurrent = OpenList(myindex)
    If AllSearch(0).Current.NodeIndex.Des.Name <> "" Then ReDim Preserve AllSearch(UBound(AllSearch) + 1)
    AllSearch(UBound(AllSearch)) = NodeCurrent
    
    If UBound(OpenList) > 0 Then
        Remove_List myindex, OpenList
    Else
        ReDim OpenList(0)
        If UBound(AllSearch) > 0 Then
            Finish = True
            'errorcode = 2
            Exit Sub
        End If
    End If
    
    If NodeCurrent.Current.NodeIndex.Des.Name = MySearch.Des.Name Then
        If MySearch.Des.Pos.X <> 0 Then
            If NodeCurrent.Current.NodeIndex.Des.Pos.X = MySearch.Des.Pos.X And _
                NodeCurrent.Current.NodeIndex.Des.Pos.Y = MySearch.Des.Pos.Y Then
                Finish = True
                Exit Sub
            End If
        Else
            Finish = True
            Exit Sub
        End If
    End If
    
    
    Get_Successor
    newg = NodeCurrent.Current.g + 1
    Dim Openi As Long
    Dim Closei As Long
    Dim found As Boolean
    Openi = -1
    Closei = -1
    For i = 0 To UBound(Sucessor)
        found = False
        For j = 0 To UBound(OpenList)
            If Sucessor(i).Des.Name = OpenList(j).Current.NodeIndex.Des.Name And _
               Sucessor(i).Des.Pos.X = OpenList(j).Current.NodeIndex.Des.Pos.X And _
               Sucessor(i).Des.Pos.Y = OpenList(j).Current.NodeIndex.Des.Pos.Y Then
                If newg >= OpenList(j).Current.g Then found = True
                Openi = j
                Exit For
            End If
            'DoEvents
        Next
        If found Then GoTo end_outer
        
        found = False
        For j = 0 To UBound(CloseList)
            If Sucessor(i).Des.Name = CloseList(j).Current.NodeIndex.Des.Name And _
               Sucessor(i).Des.Pos.X = CloseList(j).Current.NodeIndex.Des.Pos.X And _
               Sucessor(i).Des.Pos.Y = CloseList(j).Current.NodeIndex.Des.Pos.Y Then
                If newg >= CloseList(j).Current.g Then found = True
                Closei = j
                Exit For
            End If
            'DoEvents
        Next
        If found Then GoTo end_outer
        
        If Openi >= 0 Then Remove_List Openi, OpenList()
        If Closei >= 0 Then Remove_List Closei, CloseList()
        
        If OpenList(0).Current.NodeIndex.Des.Name <> "" Then
            ReDim Preserve OpenList(UBound(OpenList) + 1)
        End If
        
        OpenList(UBound(OpenList)).Current.NodeIndex = Sucessor(i)
        OpenList(UBound(OpenList)).Current.g = newg
        OpenList(UBound(OpenList)).Current.f = newg + 0
        OpenList(UBound(OpenList)).Parent = NodeCurrent.Current
end_outer:
    Next
    
    If CloseList(0).Current.NodeIndex.Des.Name <> "" Then
        ReDim Preserve CloseList(UBound(CloseList) + 1)
    End If
        
    CloseList(UBound(CloseList)) = NodeCurrent

end_sub:
End Sub

Private Sub UpdateCurrentNode()
    Dim i As Integer
    For i = 0 To UBound(AllSearch)
        If NodeCurrent.Parent.NodeIndex.Des.Name = AllSearch(i).Current.NodeIndex.Des.Name And _
           NodeCurrent.Parent.NodeIndex.Des.Pos.X = AllSearch(i).Current.NodeIndex.Des.Pos.X And _
           NodeCurrent.Parent.NodeIndex.Des.Pos.Y = AllSearch(i).Current.NodeIndex.Des.Pos.Y And _
           NodeCurrent.Parent.NodeIndex.Src.Name = AllSearch(i).Current.NodeIndex.Src.Name And _
           NodeCurrent.Parent.NodeIndex.Src.Pos.X = AllSearch(i).Current.NodeIndex.Src.Pos.X And _
           NodeCurrent.Parent.NodeIndex.Src.Pos.Y = AllSearch(i).Current.NodeIndex.Src.Pos.Y Then
            
                NodeCurrent = AllSearch(i)
                Exit Sub
        End If
    Next
End Sub

Public Function Map_Search(ByVal start As String, Destination As String, TimeLimit As Integer, Optional Sx As Long = 0, Optional Sy As Long = 0, Optional Dx As Long = 0, Optional Dy As Long = 0, Optional UseWarp As Boolean = False) As Byte
    Dim errcode As Byte
    Load_Portal
    Load_WarpPoint
    MySearch.Src.Name = start
    MySearch.Des.Name = Destination
    MySearch.Des.Pos.X = Sx
    MySearch.Des.Pos.Y = Sy
    MySearch.Des.Pos.X = Dx
    MySearch.Des.Pos.Y = Dy
    ReDim OpenList(0)
    ReDim CloseList(0)
    ReDim AllSearch(0)
    ReDim Solutions(0)
    errcode = 0
    UseNPCWarp = UseWarp
    OpenList(0).Current.NodeIndex.Des = MySearch.Src
    Finish = False
    Dim reftime As Date
    reftime = Time
    Do While Not Finish And DateDiff("s", reftime, Time) < TimeLimit
        Search_Step
        DoEvents
    Loop
    
    If DateDiff("s", reftime, Time) >= TimeLimit Then
        errcode = 1
        GoTo end_func
    End If
    
    If AllSearch(UBound(AllSearch)).Current.NodeIndex.Des.Name = MySearch.Des.Name Then
    NodeCurrent = AllSearch(UBound(AllSearch))
    Solutions(0) = AllSearch(UBound(AllSearch)).Current.NodeIndex
    Do While Solutions(UBound(Solutions)).Des.Name <> MySearch.Src.Name
        UpdateCurrentNode
        ReDim Preserve Solutions(UBound(Solutions) + 1)
        Solutions(UBound(Solutions)) = NodeCurrent.Current.NodeIndex
        DoEvents
    Loop
    End If
    
    If UBound(Solutions) > 0 Then
        ReDim Preserve Solutions(UBound(Solutions) - 1)
    Else
        errcode = 2
    End If
end_func:

    ReDim OpenList(0)
    ReDim CloseList(0)
    ReDim AllSearch(0)
    Map_Search = errcode
End Function

Private Sub Load_WarpPoint()
    On Error GoTo errie
    Dim tstr As String
    Dim Index As Integer
    Dim i As Integer
    ReDim WarpPoint(0)
    Open App.Path & "\maproute\warppoint.txt" For Input As #10
    Do While Not EOF(10)
        Line Input #10, tstr
        If WarpPoint(0).Des.Name <> "" Then ReDim Preserve WarpPoint(UBound(WarpPoint) + 1)
        For i = 1 To 6
            Index = InStr(tstr, " ")
            Select Case i
                Case 1
                    WarpPoint(UBound(WarpPoint)).Src.Name = Left(tstr, Index - 1)
                Case 2
                    WarpPoint(UBound(WarpPoint)).Src.Pos.X = Val(Left(tstr, Index - 1))
                Case 3
                    WarpPoint(UBound(WarpPoint)).Src.Pos.Y = Val(Left(tstr, Index - 1))
                Case 4
                    WarpPoint(UBound(WarpPoint)).Des.Name = Left(tstr, Index - 1)
                Case 5
                    WarpPoint(UBound(WarpPoint)).Des.Pos.X = Val(Left(tstr, Index - 1))
                Case 6
                    WarpPoint(UBound(WarpPoint)).Des.Pos.Y = Val(tstr)
            End Select
            If i < 6 Then tstr = Trim(Right(tstr, Len(tstr) - Index))
        Next
    Loop
    Close 10
    Exit Sub
errie:
    Close 10
    MsgBox "Error!!! on loading 'maproute\warppoint.txt'", vbCritical
End Sub

Private Sub Load_Portal()
    On Error GoTo errie
    Dim tstr As String
    Dim Index As Integer
    Dim i As Integer, LCount As Long
    ReDim Portals(0)
    Open App.Path & "\maproute\portals.txt" For Input As #10
    Do While Not EOF(10)
        Line Input #10, tstr
        tstr = Trim(tstr)
        LCount = LCount + 1
        If Portals(0).Des.Name <> "" Then ReDim Preserve Portals(UBound(Portals) + 1)
        For i = 1 To 6
            Index = InStr(tstr, " ")
            Select Case i
                Case 1
                    Portals(UBound(Portals)).Src.Name = Left(tstr, Index - 1)
                Case 2
                    Portals(UBound(Portals)).Src.Pos.X = Val(Left(tstr, Index - 1))
                Case 3
                    Portals(UBound(Portals)).Src.Pos.Y = Val(Left(tstr, Index - 1))
                Case 4
                    Portals(UBound(Portals)).Des.Name = Left(tstr, Index - 1)
                Case 5
                    Portals(UBound(Portals)).Des.Pos.X = Val(Left(tstr, Index - 1))
                Case 6
                    Portals(UBound(Portals)).Des.Pos.Y = Val(tstr)
            End Select
            If i < 6 Then tstr = Trim(Right(tstr, Len(tstr) - Index))
        Next
    Loop
    Close 10
    Exit Sub
errie:
    Close 10
    MsgBox "Error!!! on loading 'maproute\portals.txt' (Load_Portal) Line:" & LCount & vbCrLf & vbCrLf & Err.Description, vbCritical
End Sub

Sub Load_ExcludeMap()
    ReDim ExcludeMap(0)
    Dim tstr As String
    Open App.Path & "\maproute\mapexclude.txt" For Input As #10
    Do Until EOF(10)
        ReDim Preserve ExcludeMap(UBound(ExcludeMap) + 1)
        Input #10, tstr
        ExcludeMap(UBound(ExcludeMap)) = LCase(Trim(tstr))
    Loop
    Close 10
    ExcludeLoaded = True
    Exit Sub
errie:
    Close 10
    MsgBox "Error!!! on loading 'maproute\mapexclude.txt'", vbCritical
End Sub

Function IsExclude(MapName As String) As Boolean
    If UBound(ExcludeMap) < 1 Then
        IsExclude = False
        Exit Function
    End If
    Dim i&
    For i = 1 To UBound(ExcludeMap)
        If InStr(1, MapName, ExcludeMap(i)) Then
            IsExclude = True
            Exit Function
        End If
    Next
    IsExclude = False
End Function
