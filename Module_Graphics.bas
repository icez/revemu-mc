Attribute VB_Name = "Mod_Graphics"
Option Explicit

Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal Color As Long) As Long

'Prevent a form being resized smaller than a certain width and height
'-- Start Code --'
Public Declare Function SetWindowLong Lib "user32" Alias _
       "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
       ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias _
       "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
       ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, _
       ByVal lParam As Long) As Long

Public Const GWL_WNDPROC = (-4)
Public Const GWL_USERDATA = (-21)

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
       (Destination As Any, Source As Any, ByVal Length As Long)

Public Type MINMAXINFO
        ptReserved As Coord
        ptMaxSize As Coord
        ptMaxPosition As Coord
        ptMinTrackSize As Coord
        ptMaxTrackSize As Coord
End Type

' Windows Messages
' Look them up in the API Text Viewer. All Windows Messages
' are Constants and starts with WM_
Public Const WM_SIZE = &H5
Public Const WM_GETMINMAXINFO = &H24
Public Inithwnd As Long
Public Function SubWndProc(ByVal hwnd As Long, ByVal uMsg As Long, _
       ByVal wParam As Long, ByVal lParam As Long) As Long
   If hwnd <> Inithwnd Then Exit Function
   SubWndProc = frmItem.WindowProc(hwnd, uMsg, wParam, lParam)
        
End Function
'-- End --'




