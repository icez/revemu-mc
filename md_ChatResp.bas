Attribute VB_Name = "md_ChatResp"
Option Explicit

Public Declare Sub Sleep Lib "kernel32.dll" (ByVal Milliseconds As Long)

Function gVersion() As String
    Version = "Revemu-MC (" & App.Major & "." & App.Minor & "#" & Right("0000" & CStr(App.Revision), 4) & ")"
End Function

Function MODCheckStr(inString As String, CheckCase As String) As Boolean
On Error GoTo errie
    Dim inStr1$, inStr2$
    inStr1 = LCase(inString)
    inStr2 = LCase(CheckCase)
    If Len(inStr2) = 0 Or Len(inStr1) = 0 Then Exit Function
    Dim Srch1() As String, Srch2() As String
    Srch1 = Split(inStr2, ",")
    Dim i&, j&, iCheck&, jCheck&, resJ As Boolean
    iCheck = 0
    For i = 0 To UBound(Srch1)
        If Len(Srch1(i)) > 0 Then
            If InStr(1, Srch1(i), "/") Then
                Srch2 = Split(Srch1(i), "/")
                jCheck = 0
                resJ = False
                For j = 0 To UBound(Srch2)
                    If Len(Srch2(j)) > 0 Then
                        jCheck = InStr(iCheck + 1, inStr1, Srch2(j))
                        If jCheck > iCheck And iCheck >= 0 Then
                            iCheck = jCheck
                            resJ = True
                            Exit For
                        End If
                    End If
                Next
                If Not resJ Then GoTo falsecheck
            Else
                If InStr(iCheck + 1, inStr1, Srch1(i)) = 0 Then GoTo falsecheck
                If InStr(iCheck + 1, inStr1, Srch1(i)) > iCheck Then
                    iCheck = InStr(iCheck + 1, inStr1, Srch1(i))
                End If
            End If
        End If
    Next
truecheck:
    MODCheckStr = True
    Exit Function
falsecheck:
    MODCheckStr = False
    Exit Function
errie:
    MODCheckStr = False
    MsgBox "Error in MODCheckStr : " & Err.Description
End Function
Sub WaitTime(timeMS As Long)
On Error Resume Next
    Dim etime As Double
    etime = Timer + (timeMS / 1000)
    Do Until Timer >= etime
        Sleep 10
        DoEvents
    Loop
    Err.Clear
End Sub

Function IPToStr(inIP As String) As String
    Dim sspl() As String
    sspl = Split(inIP, ".")
    IPToStr = Chr(Val(sspl(0))) & Chr(Val(sspl(1))) & Chr(Val(sspl(2))) & Chr(Val(sspl(3)))
End Function

Function CheckIP(inIP As String) As Boolean
    Dim sspl() As String
    sspl = Split(inIP, ".")
    If UBound(sspl) <> 3 Then CheckIP = False: Exit Function
    If (Not IsNumeric(sspl(0))) Or (Not IsNumeric(sspl(1))) Or (Not IsNumeric(sspl(2))) Or (Not IsNumeric(sspl(3))) Then CheckIP = False: Exit Function
    CheckIP = True
End Function

Function HextoChr(inHex As String) As String ' 4142 > "AB"
    If (Len(inHex) Mod 2) <> 0 Then
        'SaveError "HextoChr", 0, "Hex input error : " & inHex
        Exit Function
    End If
    Dim i As Long, tChr As String, tRes As String
    For i = 1 To Len(inHex) Step 2
        tChr = Mid(inHex, i, 2)
        tRes = tRes & Chr(Format("&H" & tChr))
    Next
    HextoChr = tRes
End Function
