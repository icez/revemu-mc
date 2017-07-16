Attribute VB_Name = "md_IO"
Option Explicit

#If Win16 Then
        Private Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal FileName As String) As Integer
        Private Declare Function GetPrivateProfileString Lib "Kernel" Alias "GetPrivateProfilestring" (ByVal AppName As String, ByVal KeyName As Any, ByVal Default As String, ByVal ReturnedString As String, ByVal MAXSIZE As Integer, ByVal FileName As String) As Integer
#Else
        Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
        Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If
' INI

Public Function ReadINI(Section, KeyName, Optional Default As String, Optional FileName As String = "control\mods.ini") As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, Default, sRet, 255, _
                                           IIf(InStr(1, FileName, ":") = 0, App.Path + "\", vbNullString) + FileName))
    If LenB(ReadINI) = 0 Or Left$(ReadINI, 6) = "Error " Then ReadINI = Default
End Function

Public Sub WriteINI(sSection As String, sKeyName As String, sNewString, Optional FileName As String = "control\mods.ini")
    Dim r As Long
    r = WritePrivateProfileString(sSection, _
                                  sKeyName, _
                                  CStr(sNewString), _
                                  IIf(InStr(1, FileName, ":") = 0, App.Path + "\", vbNullString) + FileName)
End Sub

