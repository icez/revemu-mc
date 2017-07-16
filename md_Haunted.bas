Attribute VB_Name = "md_Haunted"
Public Const PROCESS_ALL_ACCESS = &H1F0FFF

Declare Function InjectLibrary Lib "madCodeHookLib.dll" (ByVal ID As Long, ByVal tstr As String) As Long

Public ProcessName As String
Public isUseHaunted As Boolean
Public Declare Function OpenProcess Lib "kernel32" _
(ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
ByVal dwProcessId As Long) As Long

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" _
  (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long

Private Declare Function Process32First Lib "kernel32.dll" _
  (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

Private Declare Function Process32Next Lib "kernel32.dll" _
  (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

Private Declare Function CloseHandle Lib "kernel32.dll" _
  (ByVal hObject As Long) As Long

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" _
  (ByVal lpString As String) As Long

Private Type PROCESSENTRY32
  dwSize As Long
  cntUsage As Long
  th32ProcessID As Long
  th32DefaultHeapID As Long
  th32ModuleID As Long
  cntThreads As Long
  th32ParentProcessID As Long
  pcPriClassBase As Long
  dwFlags As Long
  szExeFile As String * 260
End Type

Private Const TH32CS_INHERIT = &H80000000
Private Const TH32CS_SNAPALL = &HF
Private Const TH32CS_SNAPHEAPLIST = &H1
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPTHREAD = &H4
Private Const TH32CS_SNAPMODULE = &H8

Public Function GetProcessByName(Name As String) As Long
 Dim test As String
  Dim RetVal As Long
  Dim hSnap As Long
  Dim PInfo As PROCESSENTRY32

  ' Snapshot vom gesamten System erstellen
  hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
  If hSnap = -1 Then
    MsgBox "Der System-Snapshot konnte nicht erstellt werden.", _
    vbInformation, "Fehler"
    Exit Function
  End If

  PInfo.dwSize = Len(PInfo)
  RetVal = Process32First(hSnap, PInfo) ' ersten Prozess ermitteln

  Do Until RetVal = 0
    With PInfo
      .szExeFile = Trim$(Left$(.szExeFile, lstrlen(.szExeFile))) _
      ' VBNullChar abtrennen
      test = LCase(Trim(Mid$(.szExeFile, InStrRev(.szExeFile, "\") + 1)))
      If test = Name Then
        GetProcessByName = .th32ProcessID
        Exit Function
      End If
    End With
    
    RetVal = Process32Next(hSnap, PInfo) ' nächsten Prozess ermitteln
    DoEvents
  Loop
  
  CloseHandle hSnap ' Snapshot zerstören
  GetProcessByName = 0
End Function

Public Function GetProcessName(Name As String) As String
 Dim test As String
  Dim RetVal As Long
  Dim hSnap As Long
  Dim PInfo As PROCESSENTRY32

  ' Snapshot vom gesamten System erstellen
  hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
  If hSnap = -1 Then
    MsgBox "Der System-Snapshot konnte nicht erstellt werden.", _
    vbInformation, "Fehler"
    Exit Function
  End If

  PInfo.dwSize = Len(PInfo)
  RetVal = Process32First(hSnap, PInfo) ' ersten Prozess ermitteln

  Do Until RetVal = 0
    With PInfo
      .szExeFile = Trim$(Left$(.szExeFile, lstrlen(.szExeFile))) _
      ' VBNullChar abtrennen
      test = LCase(Trim(Mid$(.szExeFile, InStrRev(.szExeFile, "\") + 1)))
      If test = Name Then
        GetProcessName = Trim(Mid$(.szExeFile, InStrRev(.szExeFile, "\") + 1))
        Exit Function
      End If
    End With
    
    RetVal = Process32Next(hSnap, PInfo) ' nächsten Prozess ermitteln
    DoEvents
  Loop
  
  CloseHandle hSnap ' Snapshot zerstören
  GetProcessName = ""
End Function

Function CheckProc() As Integer
    Dim res As Long, objProcess, objWMIService, colProcesses
    res = 0
    Set objWMIService = GetObject("winmgmts:")
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process")
    For Each objProcess In colProcesses
        If LCase(objProcess.Caption) = "revemu-mc.exe" Then res = res + 1
        If LCase(objProcess.Caption) = "pub-revemu.exe" Then res = res + 1
        If LCase(objProcess.Caption) = "devil-revemu.exe" Then res = res + 1
        If LCase(objProcess.Caption) = "revemu-plus.exe" Then res = res + 1
        If LCase(objProcess.Caption) = "project1.exe" Then res = res + 1
    Next
    CheckProc = res
End Function
'FindProcess("revemu-mc.exe") + FindProcess("pub-revemu.exe") + FindProcess("devil-revemu.exe") + FindProcess("revemu-plus.exe") + FindProcess("project1.exe")
