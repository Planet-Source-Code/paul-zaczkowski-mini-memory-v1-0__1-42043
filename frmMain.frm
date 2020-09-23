VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   465
   ClientLeft      =   15
   ClientTop       =   2565
   ClientWidth     =   855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   31
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   57
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdQuit 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   135
   End
   Begin MSComctlLib.ProgressBar pbPercent 
      Height          =   135
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblUsage 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   -120
      TabIndex        =   1
      Top             =   0
      Width           =   840
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright©2002-2004 CP_You Software
'

' All variables MUST be declared
Option Explicit

' Used to hold the percent of memory usage
Dim MemUsagePercent As Integer

' Used for msgbox answers
Dim UserAnswer As Long

Private Sub cmdQuit_Click()

' Ask the user if he really wants to quit
UserAnswer = MsgBox("Are you sure you want to quit?", vbYesNo + vbQuestion, "Are you sure ...")

' If the user answers yes, then quit
If (UserAnswer = vbYes) Then

    ' Unload the form
    Unload Me
    
    ' End the program
    End
End If

End Sub

Private Sub Form_Activate()

' If an instance of memStats is already open, then close this one
If (App.PrevInstance) Then

    ' Close this one
    Unload Me
    
    ' End the program
    End

End If

' Make the form always on top
AlwaysOnTop Me.hWnd

' Start reading the memory status
GetMemoryUsage



End Sub

Private Sub GetMemoryUsage()
' Begin finding the used memory status

' Loop forever and let the computer think if it has to
Do While DoEvents
    
    ' Find the memory usage
    MemUsagePercent = GetMemUsage()
    
    ' Display the results
    lblUsage = MemUsagePercent & " %"
    
    pbPercent.Value = MemUsagePercent

Loop
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

' If the button is the left then move the form
If (Button = 1) Then

    ' Release the button
    ReleaseCapture
    
    ' Move the form
    RetVal = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)

End If


End Sub

Private Sub lblUsage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

' If the button is the left then move the form
If (Button = 1) Then

    ' Release the button
    ReleaseCapture
    
    ' Move the form
    RetVal = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)

End If

End Sub

Private Sub pbPercent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

' If the button is the left then move the form
If (Button = 1) Then

    ' Release the button
    ReleaseCapture
    
    ' Move the form
    RetVal = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)

End If


End Sub
