Attribute VB_Name = "modMoveWindow"
' Copyright©2002-2004 CP_You Software
'

' All variables MUST be declared
Option Explicit

' Used to send a message to a window
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' Let go of the mouse capture
Public Declare Sub ReleaseCapture Lib "User32" ()

' Set the button down
Public Const WM_NCLBUTTONDOWN = &HA1

' Sets the button down on a caption
Public Const HTCAPTION = 2

' Used for the return value of the sendmessage
Public RetVal As Long
