VERSION 5.00
Begin VB.Form frmRegistration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registration"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
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
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1080
      Width           =   4455
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   3495
   End
   Begin VB.TextBox txtKey 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000003&
      Height          =   255
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "You can request a key for free from MSN: 'icez@icez.cc'"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   4455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "*Don't close this windows until you got the registration code*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   375
      Width           =   4455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "= Enter your registration code below ="
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Your name : "
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Unique key : "
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PKey As String
Private PCode As String
Private PName As String
Private isLoadKey As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRegister_Click()
    Save_Key "PKey", txtKey.text
    PKey = LCase(txtKey.text)
    Save_Key "PName", txtName.text
    PName = txtName.text
    Save_Key "PCode", txtCode.text
    PCode = LCase(txtCode.text)
    If Check_Key = True Then MsgBox "Registration success.": Unload Me Else MsgBox "Registration failed"
End Sub

Private Sub Form_Load()
    gVersion
    If Check_Key Then Unload Me: Exit Sub
    Generate_Key
    Me.Show
End Sub

Sub Load_Key()
    If isLoadKey Then Exit Sub
    PKey = LCase(GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Revemu\", "PKey", ""))
    PName = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Revemu\", "PName", "")
    PCode = LCase(GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Revemu\", "PCode", ""))
    isLoadKey = True
End Sub

Function Check_Key() As Boolean
On Error GoTo res_fail
    If Not isLoadKey Then Load_Key
    Dim cmd5 As MD5
    Set cmd5 = New MD5

    If Len(PKey) = 0 Or Len(PName) = 0 Or Len(PCode) = 0 Then GoTo res_fail

    Dim TEnc$, tRes$, Conv1$, Conv2$, Conv3$, ConvT$, KeyTick$, KeyOS$
    ConvT = HexDecrypt(PKey)
    ConvT = HextoChr(ConvT)
    KeyTick = Left(ConvT, 4)
    ConvT = Right(ConvT, Len(ConvT) - 4)
    Conv1 = Left(ConvT, 16)
    ConvT = Right(ConvT, Len(ConvT) - 16)
    Conv2 = Left(ConvT, InStr(ConvT, Chr(0)) - 1)
    ConvT = Right(ConvT, Len(ConvT) - InStr(ConvT, Chr(0)))
    KeyOS = Conv2
    If Conv2 <> CStr(getWInfo.dwMajorVersion & "." & getWInfo.dwMinorVersion & "." & getWInfo.dwBuildNumber) Then GoTo res_fail
    If Conv1 <> cmd5.DigestStrToChar(Conv2) Then GoTo res_fail
    Conv1 = Left(ConvT, 16)
    ConvT = Right(ConvT, Len(ConvT) - 16)
    Conv2 = Left(ConvT, InStr(ConvT, Chr(0)) - 1)
    KeyOS = KeyOS & "/" & Conv2
    ConvT = Right(ConvT, Len(ConvT) - InStr(ConvT, Chr(0)))
    Conv3 = getWInfo.szCSDVersion
    If InStr(Conv3, Chr(0)) > 0 Then Conv3 = Left$(Conv3, InStr(Conv3, Chr(0)) - 1)
    If Conv2 <> Conv3 Then GoTo res_fail
    If Conv1 <> cmd5.DigestStrToChar(Conv2) Then GoTo res_fail
    Conv1 = Left(ConvT, 16)
    ConvT = Right(ConvT, Len(ConvT) - 16)
    Conv2 = ConvT
    If Conv2 <> Version Then GoTo res_fail
    If Conv1 <> cmd5.DigestStrToChar(Conv2) Then GoTo res_fail

    Conv1 = cmd5.DigestStrToHexStr(KeyOS & "+" & Version & KeyTick)
    
    Conv2 = TEncode(PName, "nameencryptpassword")
    Conv2 = cmd5.DigestStrToHexStr(Conv2 & KeyTick) & Conv2
    
    Conv3 = TEncode(cmd5.DigestStrToHexStr(KeyOS) & cmd5.DigestStrToHexStr(Version) & _
        Conv1 & Conv2, "resultkey")
    
    If LCase(PCode) = LCase(HexEncrypt(Conv3)) Then
        MDIfrmMain.mnuRegis.CheckED = nKillSteal
        MDIfrmMain.mnuRegis.Caption = "Enable KillSteal"
        killsteal = nKillSteal
        curEnableKey = True
        Check_Key = True
        Exit Function
    End If
res_fail:
    If Err.number > 0 Then MsgBox Err.number & " - " & Err.Description
    Check_Key = False
    MDIfrmMain.mnuRegis.Caption = "&Register"
    Exit Function
End Function

Sub Save_Key(KeyName As String, KeyVal As String)
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Revemu\", KeyName, KeyVal
End Sub

Sub Generate_Key()
    'tmp
    Dim md5Test As MD5
    Set md5Test = New MD5
    Dim tstr$
    tstr = getWInfo.szCSDVersion
    If InStr(tstr, Chr(0)) > 0 Then tstr = Left$(tstr, InStr(tstr, Chr(0)) - 1)
    'Dim PKey As String
    PKey = LngToChr(GetTickCount) & _
        md5Test.DigestStrToChar(getWInfo.dwMajorVersion & "." & getWInfo.dwMinorVersion & "." & getWInfo.dwBuildNumber) & _
        getWInfo.dwMajorVersion & "." & getWInfo.dwMinorVersion & "." & getWInfo.dwBuildNumber & _
        Chr(0) & md5Test.DigestStrToChar(tstr) & tstr & _
        Chr(0) & md5Test.DigestStrToChar(Version) & Version
    PKey = ChrtoHex(PKey)
    PKey = HexEncrypt(PKey)
    txtKey.text = PKey
    Save_Key "PKey", PKey
End Sub

Function HexEncrypt(inHex As String) As String
'sxg/zwa/qre/dct/fhv/b
    Dim i&, res$
    res = ""
    For i = 1 To Len(inHex)
        Select Case LCase(Mid(inHex, i, 1))
            Case "0": res = res & "s"
            Case "1": res = res & "x"
            Case "2": res = res & "g"
            Case "3": res = res & "k"
            Case "4": res = res & "w"
            Case "5": res = res & "y"
            Case "6": res = res & "q"
            Case "7": res = res & "r"
            Case "8": res = res & "u"
            Case "9": res = res & "p"
            Case "a": res = res & "m"
            Case "b": res = res & "t"
            Case "c": res = res & "l"
            Case "d": res = res & "h"
            Case "e": res = res & "v"
            Case "f": res = res & "j"
        End Select
    Next
    HexEncrypt = res
End Function

Function HexDecrypt(inEnc As String) As String
    Dim i&, res$
    res = ""
    For i = 1 To Len(inEnc)
        Select Case LCase(Mid(inEnc, i, 1))
            Case "s": res = res & "0"
            Case "x": res = res & "1"
            Case "g": res = res & "2"
            Case "k": res = res & "3"
            Case "w": res = res & "4"
            Case "y": res = res & "5"
            Case "q": res = res & "6"
            Case "r": res = res & "7"
            Case "u": res = res & "8"
            Case "p": res = res & "9"
            Case "m": res = res & "a"
            Case "t": res = res & "b"
            Case "l": res = res & "c"
            Case "h": res = res & "d"
            Case "v": res = res & "e"
            Case "j": res = res & "f"
        End Select
    Next
    HexDecrypt = res
End Function
