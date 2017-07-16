VERSION 5.00
Begin VB.Form frmSkillLV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Skill LV?"
   ClientHeight    =   690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1935
   ControlBox      =   0   'False
   Icon            =   "frmSkillLV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   690
   ScaleWidth      =   1935
   Begin VB.TextBox txtLV 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAF9FF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1054
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   250
      Picture         =   "frmSkillLV.frx":0E42
      Top             =   360
      Width           =   630
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   1050
      Picture         =   "frmSkillLV.frx":1346
      Top             =   360
      Width           =   630
   End
   Begin VB.Image Image3 
      Height          =   1545
      Left            =   0
      Picture         =   "frmSkillLV.frx":1884
      Top             =   -840
      Width           =   4200
   End
End
Attribute VB_Name = "frmSkillLV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
If (Val(txtLV.Text) < SkillChar(SkillNumber).MaxLV) And (Val(txtLV.Text) > 0) Then
SkillPacket = Chr(&H13) + Chr(1) + Chr(Val(txtLV.Text)) + Chr(0) + Chr(SkillChar(SkillNumber).ID) + Chr(0)
frmMain.Stat "Use Skill " + SkillChar(SkillNumber).Name + " LV." + txtLV.Text + vbCrLf
ElseIf (Val(txtLV.Text) = SkillChar(SkillNumber).MaxLV) Then
SkillPacket = Chr(&H13) + Chr(1) + Chr(SkillChar(SkillNumber).MaxLV) + Chr(1) + Chr(SkillChar(SkillNumber).ID) + Chr(0)
frmMain.Stat "Use Skill " + SkillChar(SkillNumber).Name + " LV." + CStr(SkillChar(SkillNumber).MaxLV) + vbCrLf
Else
SkillPacket = Chr(&H13) + Chr(1) + Chr(SkillChar(SkillNumber).MaxLV) + Chr(1) + Chr(SkillChar(SkillNumber).ID) + Chr(0)
frmMain.Stat "ERROR: Defualt Use Skill " + SkillChar(SkillNumber).Name + " LV." + CStr(SkillChar(SkillNumber).MaxLV) + vbCrLf
End If
IsSelectSkill = True
frmSkillLV.Visible = False
frmSkill.Enabled = True
frmSkill.Visible = False
End Sub

Private Sub Image2_Click()
frmSkillLV.Visible = False
frmSkill.Enabled = True
End Sub
