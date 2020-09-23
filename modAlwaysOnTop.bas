Attribute VB_Name = "modAlwaysOnTop"
' Copyright©2002-2004 CP_You Software
'

' All variables MUST be declared
Option Explicit

' Used to set the windows pos
Public Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' Used to set the window as always on top
Public Const HWND_TOPMOST = -1

' Used to show the window
Public Const SWP_SHOWWINDOW = &H40

' I am not sure, but you need it
Public Const SWP_NOACTIVATE = &H10

' Make sure the window cannot be sized
Public Const SWP_NOSIZE = &H1

Public Function AlwaysOnTop(hWnd As Long)

' Set the window as topmost
SetWindowPos hWnd, HWND_TOPMOST, 50, 50, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOSIZE


End Function

