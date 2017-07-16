Attribute VB_Name = "md_VarFT"
Option Explicit

Public IsWantAgi As Boolean
Public WantAgiTime As Integer
Public IsWantBles As Boolean
Public WantBlesTime As Integer
Public IsAutoChainCombo As Boolean
Public ChainComboLv As Integer
Public SpChain As Double
Public IsAutoFinishCombo As Boolean
Public FinishComboLv As Integer
Public SpFinish As Double
Public IsMapAvoid As Boolean
Public IsAutoSpirits As Boolean
Public BallSpirits As Integer
Public SpSpirits As Double
Public AutoTele As Boolean
Public DeadRecon As Boolean
Public ExAll As Boolean
Public uTime As Long
Public WarpAll As Boolean
Public JTele As Boolean
Public JobTele As String
Public PAvoid As Boolean
Public GSonyou As Boolean
Public GSnearyou As Boolean
Public MGSonyou As Boolean
Public MGSnearyou As Boolean
Public IsSpeedPot As Boolean
Public SpeedPotName As String
Public SpeedPotTime As Integer
Public IsDcArrow As Boolean
Public ArrowNumber As Integer
Public ArrowChangeNumber As Integer
Public ViewGuild As Byte

Public GetStore As Boolean
Public GetStorageItem() As GetStorageCode

Public CCSkill As SkillMonk
Public FCSkill As SkillMonk

Type GetStorageCode
    Name As String
    Amount As Long
    BackNumber As Integer
    NoStore As Boolean
End Type

Type SkillMonk
    Use As Boolean
    Monster As String
    Lv As Byte
    SP As Single
End Type

Public Function Is_UseCCSkill(Name As String) As Boolean
Dim found As Boolean
Dim X As Integer
found = False
If LCase(CCSkill.Monster) = "all" Then
    found = True
    GoTo EndFunc
End If
    If InStr(Name, CCSkill.Monster) Then found = True
EndFunc:
    Is_UseCCSkill = found
End Function

Public Function Is_UseFCSkill(Name As String) As Boolean
Dim found As Boolean
Dim X As Integer
found = False
If LCase(FCSkill.Monster) = "all" Then
    found = True
    GoTo EndFunc
End If
    If InStr(Name, FCSkill.Monster) Then found = True
EndFunc:
    Is_UseFCSkill = found
End Function

Public Sub Update_frmGuild()
    Select Case ViewGuild
        Case 0
            frmGuild.Picture1.Picture = LoadPicture(App.Path & "\interface\info_guild_bar.gif")
            If frmGuild.Frame1.Visible Then
                frmGuild.lstGuild.Visible = True
                frmGuild.Frame1.Visible = False
                If StartBot Then frmMain.Send_Guildinfo 0
            End If
        Case 1
            frmGuild.Picture1.Picture = LoadPicture(App.Path & "\interface\mem_guild_bar.gif")
            If frmGuild.lstGuild.Visible Then
                frmGuild.Frame1.Visible = True
                frmGuild.lstGuild.Visible = False
                If StartBot Then frmMain.Send_Guildinfo 1
            End If
    End Select
End Sub

