VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   2250
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   5595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   2250
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TLoad 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4680
      Top             =   1560
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   1320
      Picture         =   "frmSplash.frx":45E0
      ScaleHeight     =   120
      ScaleWidth      =   2910
      TabIndex        =   0
      Top             =   1800
      Width           =   2910
      Begin VB.PictureBox PLoad 
         BorderStyle     =   0  'None
         Height          =   60
         Left            =   40
         Picture         =   "frmSplash.frx":5862
         ScaleHeight     =   60
         ScaleWidth      =   15
         TabIndex        =   1
         Top             =   30
         Width           =   15
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Free Download at : http://www.revemu.org << click"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   5595
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MeMode As Boolean

Private Sub Form_Click()
    Shell "explorer http://www.revemu.org", vbMaximizedFocus
End Sub

Public Function getWVersion() As String
   Dim osinfo As OSVERSIONINFO
   Dim retvalue As Integer

   osinfo.dwOSVersionInfoSize = 148
   osinfo.szCSDVersion = Space$(128)
   retvalue = GetVersionExA(osinfo)

   getWVersion = osinfo.dwPlatformId
End Function

Private Sub Form_Load()
    On Error GoTo errie:
    Dim SStep As Byte
    gVersion
    
    If getWVersion = 1 Then MeMode = True

    SStep = 1
    isUseHaunted = CBool(ReadINI("haunted", "enabled", "0"))
    ProcessName = ReadINI("haunted", "exename", "ragexe.exe")
    If isUseHaunted Then SetPriorityClass GetCurrentProcess, HIGH_PRIORITY_CLASS
    SStep = 2
'begin date check
    Dim strRegPath As String
    Dim dtMax As String, dtMax2 As String
    strRegPath = "Software\" & App.CompanyName & "\" & App.Title
    dtMax = GetSettingString(HKEY_CURRENT_USER, strRegPath, "LastRun", "")
    SStep = 3
    If Len(dtMax) <> 0 Then
        Dim dtChk As Date
        On Error Resume Next
        SStep = 4
        dtChk = CDate(dtMax)
        On Error GoTo errie
        SStep = 5
        If DateDiff("s", dtChk, Now) < 0 Then
            SStep = 6
            MsgBox Version & " was expired or you're rolling back your computer time." & vbCrLf & "Please visit http://www.revemu.org for release detail.", vbCritical
            Shell "explorer http://www.icez.net", vbMaximizedFocus
            End
        End If
        SStep = 7
    End If
    
    SStep = 8
    If DateDiff("s", CDate(DateSerial(2005, 9, 30) & " 00:00:00"), Now) >= 0 Then
        MsgBox Version & " has expired." & vbCrLf & "You can get a new version of Revemu-MC at http://www.revemu.org for free.", vbCritical
        Shell "explorer http://www.icez.net", vbMaximizedFocus
        End
    End If
    SStep = 9
    If DateDiff("s", CDate(DateSerial(2005, 6, 4) & " 00:00:00"), Now) < 0 Then
        MsgBox Version & " wasn't activated." & vbCrLf & "Please visit http://www.revemu.org for release detail.", vbCritical
        Shell "explorer http://www.icez.net", vbMaximizedFocus
        End
    End If
    SStep = 10
    SaveSettingString HKEY_CURRENT_USER, strRegPath, "LastRun", CStr(Now)

    SStep = 11
    TLoad.Enabled = True
    'If Not MeMode Then MeMode = CBool(ReadINI("compability", "win9x/me", "0"))
    SStep = 12
    MakeTopMost Me.hWnd
    ReDim WayPoint(0)
    SStep = 13
    If Not MeMode Then
        SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
        'SetLayeredWindowAttributes Me.hWnd, 0, 0, LWA_ALPHA   'make it not transparent
        SStep = 14
        SetLayeredWindowAttributes Me.hWnd, 0, 255, LWA_ALPHA
    End If
    Me.Show
    Exit Sub
errie:
    MsgBox "Error!!! on initializing splash screen" & vbCrLf & vbCrLf & "Step: " & CStr(SStep) & vbCrLf & "Code: " & Err.number & vbCrLf & "Desc: " & Err.Description, vbOKOnly, "Error !!!"
    End
End Sub

Private Sub Label1_Click()
    Shell "explorer http://www.revemu.org", vbMaximizedFocus
End Sub

Private Sub TLoad_Timer()
On Error GoTo errie
    Dim lStep As Byte
    PLoad.width = PLoad.width + 30
    lStep = 1
    If PLoad.width > 2910 Then
        TLoad.Enabled = False
        lStep = 2
        MakeNormal frmSplash.hWnd
        lStep = 3
        
        'GoTo dcheck
        'If getWVersion = 1 Then MeMode = True: GoTo dcheck
        'begin duplication check
        If (LCase(App.EXEName) <> "revemu-mc") And (LCase(App.EXEName) <> "project1") Then
            lStep = 4
            MsgBox "Don't rename 'revemu-mc.exe'."
            End
        End If
        lStep = 5
        Load MDIfrmMain
        lStep = 6
        Unload frmSplash
        Exit Sub
    End If
    If Not MeMode Then
        'SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
        lStep = 7
        SetLayeredWindowAttributes Me.hWnd, 0, (PLoad.width * 255) / 2910, LWA_ALPHA
    End If
    Exit Sub
errie:
    MsgBox "Error!!! on loading splash screen" & vbCrLf & vbCrLf & "Step: " & CStr(lStep) & vbCrLf & "Code: " & Err.number & vbCrLf & "Desc: " & Err.Description, vbOKOnly, "Error !!!"
    End
End Sub
