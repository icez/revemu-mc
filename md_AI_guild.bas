Attribute VB_Name = "md_AI_guild"
Option Explicit

Public GuildAlliance() As GuildAllianceInfo
Public Guild() As GuildMember
Public GuildPos() As GuildPosition
Public GuildInfo As GuildInfoz
Type GuildPosition
    Position As Long
    PosName As String
End Type
Type GuildAllianceInfo
    isAlliance As Boolean
    ID As String
    Name As String
End Type
'R 0154 <len>.w {<accID>.l <charactorID>.l <hair type>.w <hair color>.w <sex>.w <job>.w <lvl?>.w
'<guild exp>.l <online>.l <Position>.l ?.50B <nick>.24B}*
Type GuildMember
    AccID As String
    CharID As String
    Sex As String
    Class As String
    Lv As Integer
    EXP As Long
    isOnline As Long
    Position As Long
    PosName As String
    Name As String
End Type
'R 0150 <guildID>.l <guildLv>.l <connum>.l <Max PPL?>.l <Avl.lvl>.l ?.l <next_exp>.l ?.16B
'<guild name>.24B <guild master>.24B ?.16B
Type GuildInfoz
    GuID As String * 4
    GuLV As Long
    CurOnline As Long
    MaxPPL As Long
    GuAvLV As Long
    GuEXP As Long
    GuNextEXP As Long
    Name As String
    GuMaster As String
End Type
Function GetGuildPos(Position As Long) As String
    Dim i&
    For i = 0 To UBound(GuildPos) - 1
        If GuildPos(i).Position = Position Then
            GetGuildPos = GuildPos(i).PosName
            Exit Function
        End If
    Next
End Function

