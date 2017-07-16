VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Type tNode
  X          As Long
  Y          As Long
  lValue     As Long        'the Path value (sum) so far
  lHeuristic As Long        'a weighted value for choices
  lLink      As Long        'a link between subsequent nodes in the chain
End Type

Private Type inPoint
    X       As Long
    Y       As Long
End Type

Private Type inMap
    Name As String
    Pos As inPoint
End Type

Private Type inPortal
    Src As inMap
    Des As inMap
End Type

Private Const LOW_BOUND As Long = 1

Public HIGH_BOUND      As Long
Public bAStar          As Boolean '// Running A* and not Dijkstra

Private Portals()      As inPortal
Private tNodes()       As tNode
Private lNodes         As Long    'The number of nodes in operation
Private lNodeCur       As Long    'The index of the one we are working on
Private lFirstEmpty    As Long    'Where to start searching for an empty node

Private lBeenHere()    As Long    '0 if not, if so index into tNodes(), -2 if impassable
Private lAIMap()       As Byte    'The connection set aiMAP

Private lEndX          As Long
Private lEndY          As Long    'To store the end points

Private xDim           As Long
Private yDim           As Long    'RVT The X and Y dimensions of the tableau (it doesnt need to be square anymore)

Private Sub Class_Terminate()
    Erase tNodes
    Erase lBeenHere
    Erase lAIMap
End Sub

Private Sub FindNearerPoint(ByRef X, Y As Long)
    Dim Des As inPoint
    Dim i, j As Integer
    Des.X = X
    Des.Y = Y
    If lAIMap(Des.X, Des.Y) = 0 Then Exit Sub
    For i = Des.X - 1 To Des.X + 1
        For j = Des.Y - 1 To Des.Y + 1
            If i >= 0 And j >= 0 And i < xDim And j < yDim And i <> Des.X And j <> Des.Y Then
                If lAIMap(i, j) = 0 Then
                    X = i
                    Y = j
                    Exit Sub
                End If
            End If
        Next
    Next
    
    For i = Des.X - 2 To Des.X + 2 Step 2
        For j = Des.Y - 2 To Des.Y + 2 Step 2
            If i >= 0 And j >= 0 And i < xDim And j < yDim Then
                If lAIMap(i, j) = 0 Then
                    X = i
                    Y = j
                    Exit Sub
                End If
            End If
        Next
    Next
End Sub

Public Sub Load_Portal(MapName As String)
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
        Index = InStr(tstr, MapName & " ")
        If Index = 0 Then GoTo end_loop
        Index = InStr(Index + 1, tstr, MapName & " ")
        If Index = 0 Then GoTo end_loop
        If Portals(0).Des.Name <> "" Then ReDim Preserve Portals(UBound(Portals) + 1)
        For i = 1 To 6
            Index = InStr(tstr, " ")
            Select Case i
                Case 1
                    Portals(UBound(Portals)).Src.Name = Left(tstr, Index - 1)
                Case 2
                    Portals(UBound(Portals)).Src.Pos.X = Val(Left(tstr, Index - 1))
                Case 3
                    Portals(UBound(Portals)).Src.Pos.Y = yDim - Val(Left(tstr, Index - 1))
                Case 4
                    Portals(UBound(Portals)).Des.Name = Left(tstr, Index - 1)
                Case 5
                    Portals(UBound(Portals)).Des.Pos.X = Val(Left(tstr, Index - 1))
                Case 6
                    Portals(UBound(Portals)).Des.Pos.Y = yDim - Val(tstr)
            End Select
            If i < 6 Then
                tstr = Trim(Right(tstr, Len(tstr) - Index))
            Else
                With Portals(UBound(Portals)).Src.Pos
                If lAIMap(.X, .Y) > 0 Then FindNearerPoint .X, .Y
                End With
                With Portals(UBound(Portals)).Des.Pos
                If lAIMap(.X, .Y) > 0 Then FindNearerPoint .X, .Y
                End With
            End If
        Next
end_loop:
    Loop
    Close 10
    Exit Sub
errie:
    Close 10
    MsgBox "Error!!! on loading 'maproute\portals.txt' (Load_Portal2) Line:" & LCount & vbCrLf & vbCrLf & Err.Description, vbCritical
End Sub

Public Function RunSearch(lXStart As Long, lYStart As Long, lXEnd As Long, lYEnd As Long, aiMap() As Byte, MapName As String) As Boolean
    '// Returns TRUE if path found
  
    StartSearch lXStart, lYStart, lXEnd, lYEnd, aiMap()
    Load_Portal MapName
    
    If bAStar Then                                            'AStar using the NextNode routine
        Do Until NextNode Or RunSearch
            RunSearch = UpdateCurrentNode
            DoEvents
        Loop
    Else                                                      'DIJKSTRA (with a little optimisation)
        Do Until lNodeCur > lNodes Or RunSearch
            If tNodes(lNodeCur).lValue = PATH_EMPTY Then
                RunSearch = UpdateCurrentNode
                DoEvents
            End If
            lNodeCur = lNodeCur + 1
        Loop
    End If
    BackTracePath                     'Go back from goal to start and record
End Function

'================================================================================================================

Public Sub StartSearch(ByRef lXStart As Long, ByRef lYStart As Long, ByRef lXEnd As Long, ByRef lYEnd As Long, aiMap() As Byte)
  
    Dim xLD As Long, xUD As Long, yLD As Long, yUD As Long
      
    xLD = LBound(aiMap, 1)
    xUD = UBound(aiMap, 1)
    xDim = xUD - xLD + 1
    
    yLD = LBound(aiMap, 2)
    yUD = UBound(aiMap, 2)
    yDim = yUD - yLD + 1
    
    
    
    'take a copy of the connection map - the costs of crossing a tile
    ReDim lAIMap(xLD To xUD, yLD To yUD) As Byte
    lAIMap() = aiMap()
    
    If lAIMap(lXStart, lYStart) <> 0 Then
        FindNearerPoint lXStart, lYStart
    End If
    If lAIMap(lXEnd, lYEnd) <> 0 Then
        FindNearerPoint lXEnd, lYEnd
    End If
    'A record of which Nodes have been visited ie. are on the queue
    ReDim lBeenHere(xLD To xUD, yLD To yUD) As Long
    
    HIGH_BOUND = xDim * yDim
    ReDim tNodes(LOW_BOUND To HIGH_BOUND) As tNode
  
    '// Store the end point
    lEndX = lXEnd
    lEndY = lYEnd
    
    '// Store the start point
    
    Add lXStart, lYStart 'This one now Current
    lNodeCur = LOW_BOUND 'Make it the current one
    With tNodes(lNodeCur)
        .lValue = 0          'StartofPath then Fist 4 nodes of near neigbours
        If .X > 0 Then Add .X - 1, .Y '// IMP
        If .Y > 0 Then Add .X, .Y - 1 '// IMP
        If (.X + 1) < xDim Then Add .X + 1, .Y
        If (.Y + 1) < yDim Then Add .X, .Y + 1
    End With
    
    lNodeCur = lNodeCur + 1 'the first node to be examined is this one
    lFirstEmpty = lNodeCur
End Sub

Private Sub Add(lx As Long, ly As Long)
    Dim lDistX As Long, lDistY As Long
    
    If lBeenHere(lx, ly) = 0 Then                'Is it already there ie. seen?
        If lAIMap(lx, ly) = 0 Then   'and its a useful point to move to ie dont add duds
            'havent been here and it isn't impassable so add it
            lNodes = lNodes + 1
            With tNodes(lNodes)
            .X = lx
            .Y = ly
            .lValue = PATH_EMPTY
            .lLink = 0                             'to keep track of where we came from
    
            '// When adding a point to the list, figure out a heuristic if in A* mode
            If bAStar Then
                lDistX = lAIMap(lx, ly) * (lx - lEndX) 'increase the disadvantage of expensive paths
                lDistY = lAIMap(lx, ly) * (ly - lEndY)
                .lHeuristic = (lDistX * lDistX + lDistY * lDistY)
            End If
            End With
            lBeenHere(lx, ly) = lNodes       'mark as seen and in queue at this place
        Else
            lBeenHere(lx, ly) = PATH_IMPOSSIBLE      'seen but not in queue (impassable)
        End If
    End If
End Sub

Public Function NextNode() As Boolean
    Dim i As Long
    
    If bAStar Then
        'RVT The A* algorithm looks for the EMPTY_PATH node with the smallest heuristic value.
        '    So if we keep track of the first empty one and then look for the
        '    smallest from there we save ourselves a sort and one and a bit linear time searches
        
        For i = lFirstEmpty To lNodes    'look for the first empty node from where we last looked
            If tNodes(i).lValue = PATH_EMPTY Then Exit For
        Next

        If i <= lNodes Then              'we found one
            lFirstEmpty = i               'This says where now to look from next time
            lNodeCur = SmallestEmptyPath 'This one is therefore the SMALLEST HEURISTIC EMPTY_PATH NODE
            Exit Function
        End If
    Else
    
    '// IMP
'        Dim lBnd&
'        lBnd = lNodes - NUMBER_OF_TILES
'        If lBnd <= LOW_BOUND Then lBnd = LOW_BOUND
        For i = lFirstEmpty To lNodes 'lFirstEmpty increases monotonically (usually check only 1 or 2)
            If tNodes(i).lValue = PATH_EMPTY Then
'                If I <= lBnd Then Stop
                lNodeCur = i
                lFirstEmpty = i
                Exit Function
            End If
        Next
    End If
    
    NextNode = True 'At this point the search is at an end - It may not have been successful (deadended)...
                    'BTW searches dont always end here - if the end point is encountered they stop anyway
End Function

Private Function SmallestEmptyPath() As Long 'Find the smallest Heuristic Path Value from lFirstEmpty on
    Dim i As Long, minHeuristic As Long, minNode As Long
    
    minNode = lFirstEmpty
    minHeuristic = tNodes(lFirstEmpty).lHeuristic       'remember where and what the Heuristic value is
    
    For i = lFirstEmpty + 1 To lNodes                   'now check the rest keeping track of the smallest lHeuristic
        If tNodes(i).lValue = PATH_EMPTY Then           'must be empty
            If tNodes(i).lHeuristic < minHeuristic Then 'and its smaller
                minNode = i
                minHeuristic = tNodes(i).lHeuristic
            End If
        End If
    Next
  
    If minNode <> lFirstEmpty Then
        SwapNodes minNode, lFirstEmpty
    End If                                            'we swap them around so that the pending list gradually shrinks
  
    SmallestEmptyPath = lFirstEmpty
End Function

Private Sub SwapNodes(ByVal lNodeIdx1 As Long, ByVal lNodeIdx2 As Long)
    Dim z As tNode
    
    'do the swap
    z = tNodes(lNodeIdx1)
    tNodes(lNodeIdx1) = tNodes(lNodeIdx2)
    tNodes(lNodeIdx2) = z
    
    'fix the BeenHere pointers
    lBeenHere(tNodes(lNodeIdx1).X, tNodes(lNodeIdx1).Y) = lNodeIdx1
    lBeenHere(tNodes(lNodeIdx2).X, tNodes(lNodeIdx2).Y) = lNodeIdx2
End Sub

Public Function BackTracePath()
    Dim NodeFrom As Long, NodeTo As Long, z As Long
    
    z = lBeenHere(lEndX, lEndY)
    If z Then lNodeCur = z
    NodeTo = lNodeCur
    NodeFrom = tNodes(lNodeCur).lLink
    tNodes(NodeTo).lLink = 0                 'mark as end of chain
    If NodeFrom Then
        Do
            z = tNodes(NodeFrom).lLink             'need this for a while
            tNodes(NodeFrom).lLink = NodeTo        'mark the new link
            NodeTo = NodeFrom
            NodeFrom = z
        Loop Until NodeFrom = 0                     'start has lLink=0
    End If
End Function

Public Function PathStepNext(ByRef lx As Long, ByRef ly As Long, ByRef Result As Integer) As Long
    lNodeCur = tNodes(lBeenHere(lx, ly)).lLink
    If lNodeCur Then GetNode lx, ly
    Result = isportal(lx, ly)
    PathStepNext = lNodeCur
End Function

Public Function GetNode(ByRef lx As Long, ByRef ly As Long)
    lx = tNodes(lNodeCur).X
    ly = tNodes(lNodeCur).Y
End Function

Public Function GetNodeValue(ByVal lx As Long, ByVal ly As Long) As Long
    '// Cannot use Current Node
    Dim i As Long
                                            'lBeenHere() records the index of Node(x,y) in the nodelist
    i = lBeenHere(lx, ly)                   'a constant time lookup
    If i > 0 Then                           'its in the queue
        GetNodeValue = tNodes(i).lValue
    Else
        GetNodeValue = PATH_EMPTY
    End If
End Function

Private Function isportal(lx As Long, ly As Long) As Integer
    Dim X As Integer
    For X = 0 To UBound(Portals)
        If Abs(lx - Portals(X).Src.Pos.X) = 0 And Abs(ly - Portals(X).Src.Pos.Y) = 0 Then
            isportal = X
            Exit Function
        End If
    Next
    isportal = -1
End Function


Public Function UpdateCurrentNode() As Boolean      'TRUE if this node was the END Node
    
    Dim lVal As Long, lPathMinVal As Long
    Dim i As Long, X As Long, Y As Long
    Dim number As Integer
    lPathMinVal = PATH_HUGE                  'assume a really large value
    number = RandomNumber(1, 500) Mod 30
    'number = 15
    With tNodes(lNodeCur)
        
        X = isportal(.X, .Y)
        If (X > -1) Then
            i = lBeenHere(Portals(X).Des.Pos.X, Portals(X).Des.Pos.Y)
            If i = 0 Then
                Add Portals(X).Des.Pos.X, Portals(X).Des.Pos.Y                      'if not there Then Expand to it for next time
            ElseIf i > 0 Then
                lVal = tNodes(i).lValue
                If lVal >= 0 Then
                    If lVal < lPathMinVal Then
                        lPathMinVal = lVal
                        .lLink = i                'new link from here
                    End If
                End If
            End If
        End If
        
        If number < 16 Then
            For X = .X - 1 To .X + 1 Step 2        'Check LEFT and RIGHT
                If X >= 0 Then
                    If X < xDim Then
                        i = lBeenHere(X, .Y)             'CHECK point
                        If i = 0 Then
                            Add X, .Y                      'if not there Then Expand to it for next time
                        ElseIf i > 0 Then
                            lVal = tNodes(i).lValue
                            If lVal >= 0 Then
                                If lVal < lPathMinVal Then
                                    lPathMinVal = lVal
                                    .lLink = i                'new link from here
                                End If
                            End If
                        End If
                    End If
                End If
            Next
    
            For Y = .Y - 1 To .Y + 1 Step 2       'CHECK UP and DOWN
                If Y >= 0 Then
                    If Y < yDim Then
                        i = lBeenHere(.X, Y)             'CHECK point
                        If i = 0 Then
                            Add .X, Y                      'if not there Then Expand to it for next time
                        ElseIf i > 0 Then
                            lVal = tNodes(i).lValue
                            If lVal >= 0 Then
                                If lVal < lPathMinVal Then
                                    lPathMinVal = lVal
                                    .lLink = i                'new link from here
                                End If
                            End If
                        End If
                    End If
                End If
            Next
    ElseIf number < 20 Then
        For X = .X - 1 To .X + 1 Step 1       'Check LEFT and RIGHT
            For Y = .Y - 1 To .Y + 1 Step 1
                If X >= 0 And Y >= 0 And (X <> .X Or Y <> .Y) Then
                    If X < xDim And Y < yDim Then
                        i = lBeenHere(X, Y)             'CHECK point
                        If i = 0 Then
                            Add X, Y                      'if not there Then Expand to it for next time
                        ElseIf i > 0 Then
                            lVal = tNodes(i).lValue
                            If lVal >= 0 Then
                                If lVal < lPathMinVal Then
                                    lPathMinVal = lVal
                                    .lLink = i                'new link from here
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        Next
    Else
        For Y = .Y - 1 To .Y + 1 Step 1
            For X = .X + 1 To .X - 1 Step -1       'Check LEFT and RIGHT
                If X >= 0 And Y >= 0 And (X <> .X Or Y <> .Y) Then
                    If X < xDim And Y < yDim Then
                        i = lBeenHere(X, Y)             'CHECK point
                        If i = 0 Then
                            Add X, Y                      'if not there Then Expand to it for next time
                        ElseIf i > 0 Then
                            lVal = tNodes(i).lValue
                            If lVal >= 0 Then
                                If lVal < lPathMinVal Then
                                    lPathMinVal = lVal
                                    .lLink = i                'new link from here
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        Next
    End If
    .lValue = lPathMinVal + lAIMap(.X, .Y)                'currently min Path Cost
    UpdateCurrentNode = ((.X = lEndX) And (.Y = lEndY))   'TRUE if we found the END
    End With
End Function

