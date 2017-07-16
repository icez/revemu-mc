VERSION 5.00
Begin VB.Form frmConfAttack 
   Caption         =   "Monster Attack Option"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   4815
   Begin VB.Frame Frame1 
      Caption         =   "Skill Slot"
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   4575
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update or Add to Attack List"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   4335
      End
      Begin VB.TextBox cmbCount2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   15
         Text            =   "0"
         Top             =   975
         Width           =   375
      End
      Begin VB.ComboBox cmbLv2 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   960
         Width           =   615
      End
      Begin VB.ComboBox cmbSlot2 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtMons 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Top             =   210
         Width           =   2775
      End
      Begin VB.TextBox cmbCount1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   7
         Text            =   "0"
         Top             =   660
         Width           =   375
      End
      Begin VB.ComboBox cmbLv1 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   645
         Width           =   615
      End
      Begin VB.ComboBox cmbSlot1 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   645
         Width           =   1935
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Count: "
         Height          =   255
         Left            =   3480
         TabIndex        =   14
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Lv: "
         Height          =   255
         Left            =   2520
         TabIndex        =   12
         Top             =   1005
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Slot 2 : "
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Skill slot for monster: "
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Count: "
         Height          =   255
         Left            =   3480
         TabIndex        =   6
         Top             =   690
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Lv: "
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   690
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Slot 1 : "
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   690
         Width           =   615
      End
   End
   Begin VB.ListBox lstAtk 
      Appearance      =   0  'Flat
      Height          =   2730
      ItemData        =   "frmConfAttack.frx":0000
      Left            =   120
      List            =   "frmConfAttack.frx":0002
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmConfAttack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public frmLoaded As Boolean

Private Sub cmbCount1_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub cmbCount2_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    'tmp
    frmLoaded = False
    LoadFormPos Me
    Update_AtkList
    Update_ListSkill
    Frame1.Enabled = False
    frmLoaded = True
End Sub

Private Sub Form_Resize()
    If Me.width < 4800 Then
        Me.width = 4800
        Exit Sub
    End If
    If Me.height < 5100 Then
        Me.height = 5220
        Exit Sub
    End If
    lstAtk.width = Me.width - 360
    lstAtk.height = Me.height - 2490
    Me.height = lstAtk.height + 2490
    Frame1.width = lstAtk.width
    Frame1.Top = lstAtk.height + 270
    cmbSlot1.width = Frame1.width - 2640
    Label2.Left = cmbSlot1.width + 585
    cmbLv1.Left = cmbSlot1.width + 945
    Label3.Left = cmbSlot1.width + 1545
    cmbCount1.Left = cmbSlot1.width + 2145
    cmbSlot2.width = Frame1.width - 2640
    Label6.Left = cmbSlot2.width + 585
    cmbLv2.Left = cmbSlot2.width + 945
    Label7.Left = cmbSlot2.width + 1545
    cmbCount2.Left = cmbSlot2.width + 2145
    cmdUpdate.width = Frame1.width - 240
    txtMons.width = Frame1.width - 1800
    SaveFormPos Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub

Public Sub Update_AtkList()
    Dim i&, j&
    lstAtk.Clear
    For i = 0 To UBound(Monsters)
        lstAtk.AddItem Monsters(i).Name
    Next
    For i = 0 To UBound(Attack)
        For j = 0 To lstAtk.ListCount - 1
            If LCase(Attack(i).Name) = LCase(lstAtk.List(j)) Then
                lstAtk.Selected(j) = True
                lstAtk.List(j) = lstAtk.List(j) & IIf(Len(Attack(i).Spell1) > 0, ":" & Attack(i).Spell1 & "/Lv-" & Attack(i).lv1 & "/Cnt-" & Attack(i).UTime1, "") & IIf(Len(Attack(i).Spell2) > 0, ":" & Attack(i).Spell2 & "/Lv-" & Attack(i).lv2 & "/Cnt-" & Attack(i).UTime2, "")
            End If
            If LCase(Attack(i).Name) = "/" & LCase(lstAtk.List(j)) Then
                lstAtk.List(j) = lstAtk.List(j) & IIf(Len(Attack(i).Spell1) > 0, ":" & Attack(i).Spell1 & "/Lv-" & Attack(i).lv1 & "/Cnt-" & Attack(i).UTime1, "") & IIf(Len(Attack(i).Spell2) > 0, ":" & Attack(i).Spell2 & "/Lv-" & Attack(i).lv2 & "/Cnt-" & Attack(i).UTime2, "")
            End If
        Next
    Next
    lstAtk.ListIndex = -1
End Sub

Public Sub Refresh_AtkList(Optional MonsName As String = "")
    Dim i&, olist&, j&
    olist = lstAtk.ListIndex
    lstAtk.Visible = False
    For i = 0 To UBound(Attack)
        If InStr(Attack(i).Name, MonsName) > 0 Then
            For j = 0 To lstAtk.ListCount - 1
                If LCase(Attack(i).Name) = LCase(SplStr(lstAtk.List(j), ":", 0)) Then
                    lstAtk.Selected(j) = True
                    lstAtk.List(j) = SplStr(lstAtk.List(j), ":", 0) & IIf(Len(Attack(i).Spell1) > 0, ":" & Attack(i).Spell1 & "/Lv-" & Attack(i).lv1 & "/Cnt-" & Attack(i).UTime1, "") & IIf(Len(Attack(i).Spell2) > 0, ":" & Attack(i).Spell2 & "/Lv-" & Attack(i).lv2 & "/Cnt-" & Attack(i).UTime2, "")
                End If
            Next
        End If
    Next
    lstAtk.Visible = True
    lstAtk.ListIndex = olist
End Sub

Public Sub Update_ListSkill()
    Dim i&
    cmbSlot1.Clear
    cmbSlot2.Clear
    cmbSlot1.AddItem ""
    cmbSlot2.AddItem ""
    For i = 0 To UBound(SkillIDName)
        cmbSlot1.AddItem SkillIDName(i).Name & " ::: " & SkillIDName(i).raw
        cmbSlot2.AddItem SkillIDName(i).Name & " ::: " & SkillIDName(i).raw
    Next
    cmbLv1.Clear
    cmbLv2.Clear
    For i = 1 To 10
        cmbLv1.AddItem CStr(i)
        cmbLv2.AddItem CStr(i)
    Next
End Sub

Public Sub Clear_Frame()
    If Not frmLoaded Then Exit Sub
    txtMons.text = ""
    cmbSlot1.ListIndex = -1
    cmbSlot2.ListIndex = -1
    cmbLv1.ListIndex = -1
    cmbLv2.ListIndex = -1
    cmbCount1.text = "0"
    cmbCount2.text = "0"
End Sub
Public Sub Refresh_Frame()
    'tmp
    If Not frmLoaded Then Exit Sub
    Dim i&, j&
    Clear_Frame
    If lstAtk.Selected(lstAtk.ListIndex) = False Then Frame1.Enabled = False:  Exit Sub
    txtMons.text = SplStr(lstAtk.List(lstAtk.ListIndex), ":", 0)
    For i = 0 To UBound(Attack)
        If LCase(Attack(i).Name) = LCase(SplStr(lstAtk.List(lstAtk.ListIndex), ":", 0)) Then
            For j = 0 To UBound(SkillIDName)
                If SkillIDName(j).raw = Attack(i).Spell1 Then cmbSlot1.ListIndex = j + 1
                If SkillIDName(j).raw = Attack(i).Spell2 Then cmbSlot2.ListIndex = j + 1
            Next
            cmbLv1.ListIndex = Attack(i).lv1 - 1
            cmbLv2.ListIndex = Attack(i).lv2 - 1
            cmbCount1.text = Attack(i).UTime1
            cmbCount2.text = Attack(i).UTime2
            Exit Sub
        End If
    Next
End Sub

Public Sub Save_Frame()
    'tmp
    If Not frmLoaded Then Exit Sub
    Dim i&, FF&
    FF = -1
    For i = 0 To UBound(Attack)
        If LCase(Attack(i).Name) = LCase(txtMons.text) Then FF = i
    Next
    If FF < 0 Then ReDim Preserve Attack(UBound(Attack) + 1): FF = UBound(Attack)
    With Attack(FF)
        .Name = txtMons.text
        .lv1 = cmbLv1.ListIndex + 1
        .lv2 = cmbLv2.ListIndex + 1
        .Spell1 = SplStr(cmbSlot1.List(cmbSlot1.ListIndex), " ::: ", 1)
        .Spell2 = SplStr(cmbSlot2.List(cmbSlot2.ListIndex), " ::: ", 1)
        .UTime1 = Val(cmbCount1.text)
        .UTime2 = Val(cmbCount2.text)
    End With
    Refresh_AtkList txtMons.text
    Save_ConfAttack
End Sub

Private Sub cmdUpdate_Click()
    Save_Frame
End Sub

Private Sub lstAtk_Click()
    If Not frmLoaded Then Exit Sub
    If lstAtk.ListIndex > -1 Then
        Frame1.Enabled = True
        Refresh_Frame
    Else
        Frame1.Enabled = False
        Clear_Frame
    End If
End Sub

Private Sub lstAtk_ItemCheck(Item As Integer)
    If Not frmLoaded Then Exit Sub
    Dim i&, tmp1$, tmp2$, tmp11$, ffound&
    If lstAtk.Selected(Item) = True Then
        ffound = -1
        For i = 0 To UBound(Attack)
            If Attack(i).Name = SplStr(lstAtk.List(Item), ":", 0) Or Attack(i).Name = "/" & SplStr(lstAtk.List(Item), ":", 0) Then
                ffound = i
                Exit For
            End If
        Next
        If ffound < 0 Then ReDim Preserve Attack(UBound(Attack) + 1): ffound = UBound(Attack)
        Attack(ffound).Name = SplStr(lstAtk.List(Item), ":", 0)
        tmp1 = SplStr(lstAtk.List(Item), ":", 1)
        tmp2 = SplStr(lstAtk.List(Item), ":", 2)
        Attack(ffound).Spell1 = SplStr(tmp1, "/Lv-", 0)
        Attack(ffound).Spell2 = SplStr(tmp2, "/Lv-", 0)
        tmp11 = SplStr(tmp1, "/Lv-", 1)
        Attack(ffound).lv1 = Val(SplStr(tmp11, "/Cnt-", 0))
        Attack(ffound).UTime1 = Val(SplStr(tmp11, "/Cnt-", 1))
        tmp11 = SplStr(tmp2, "/Lv-", 1)
        Attack(ffound).lv2 = Val(SplStr(tmp11, "/Cnt-", 0))
        Attack(ffound).UTime2 = Val(SplStr(tmp11, "/Cnt-", 1))
        'attack(ubound(attack))
    Else
        For i = 0 To UBound(Attack)
            If Attack(i).Name = SplStr(lstAtk.List(Item), ":", 0) Then
                Attack(i).Name = "/" & Attack(i).Name
                Exit For
            End If
        Next
    End If
    Save_ConfAttack
End Sub

Sub Save_ConfAttack()
    Dim i&, tmp1$, tmp2$, tmp11$, tmp12$, res$
    ReDim Attack(lstAtk.ListCount - 1)
    Open App.Path & "\control\attack.txt" For Output As #27
    For i = 0 To lstAtk.ListCount - 1
        'If lstAtk.Selected(i) = True Then
            res = ""
            res = SplStr(lstAtk.List(i), ":", 0)
            Attack(i).Name = IIf(lstAtk.Selected(i) = False, "/", "") & SplStr(lstAtk.List(i), ":", 0)
            Attack(i).ID = return_monsid(Attack(i).Name)
            
            tmp1 = SplStr(lstAtk.List(i), ":", 1)
            tmp2 = SplStr(lstAtk.List(i), ":", 2)
            
            tmp11 = SplStr(tmp1, "/Lv-", 0)
            tmp12 = SplStr(tmp1, "/Lv-", 1)
            Attack(i).Spell1 = tmp11
            Attack(i).UTime1 = Val(SplStr(tmp12, "/Cnt-", 1))
            Attack(i).lv1 = Val(SplStr(tmp12, "/Cnt-", 0))
            If Len(tmp11) > 0 Then res = res & " - " & tmp11 & String$(Val(SplStr(tmp12, "/Cnt-", 1)), "/") & " " & Val(SplStr(tmp12, "/Cnt-", 0))
            
            tmp11 = SplStr(tmp2, "/Lv-", 0)
            tmp12 = SplStr(tmp2, "/Lv-", 1)
            
            Attack(i).Spell2 = tmp11
            Attack(i).UTime2 = Val(SplStr(tmp12, "/Cnt-", 1))
            Attack(i).lv2 = Val(SplStr(tmp12, "/Cnt-", 0))
            If Len(tmp11) > 0 Then res = res & " - " & tmp11 & String$(Val(SplStr(tmp12, "/Cnt-", 1)), "/") & " " & Val(SplStr(tmp12, "/Cnt-", 0))
            Print #27, IIf(lstAtk.Selected(i) = False, "/", "") & res
        'Else
        '    Print #27, "/" & SplStr(lstAtk.List(i), ":", 0)
        '    Attack(i).Name = "/" & SplStr(lstAtk.List(i), ":", 0)
        'End If
    Next
    Close #27
    Update_AtkSkill
End Sub
