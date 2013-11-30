VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Desktop Enhancer"
   ClientHeight    =   660
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   2730
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   660
   ScaleWidth      =   2730
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrUpdateCaptionDelay 
      Interval        =   500
      Left            =   120
      Top             =   0
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "Hidden"
      Begin VB.Menu mnuDesktops 
         Caption         =   "1"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuDesktops 
         Caption         =   "2"
         Index           =   2
      End
      Begin VB.Menu mnuDesktops 
         Caption         =   "3"
         Index           =   3
      End
      Begin VB.Menu mnuDesktops 
         Caption         =   "4"
         Index           =   4
      End
      Begin VB.Menu mnuDesktops 
         Caption         =   "5"
         Index           =   5
      End
      Begin VB.Menu mnuDesktops 
         Caption         =   "6"
         Index           =   6
      End
      Begin VB.Menu mnuDesktops 
         Caption         =   "7"
         Index           =   7
      End
      Begin VB.Menu mnuDesktops 
         Caption         =   "8"
         Index           =   8
      End
      Begin VB.Menu mnuDesktops 
         Caption         =   "9"
         Index           =   9
      End
      Begin VB.Menu mnuDesktops 
         Caption         =   "10"
         Index           =   10
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVisualTaskManager 
         Caption         =   "Visual Task Manager"
      End
      Begin VB.Menu mnuChangeResolution 
         Caption         =   "Resolution"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lDesktopEnhancer As clsDesktopEnhancer
Private lSystray As clsSystray

Private Sub Form_Load()
On Local Error GoTo ErrHandler
Set lSystray = New clsSystray
Set lDesktopEnhancer = New clsDesktopEnhancer
lSystray.LoadTrayIcon
Me.Hide
Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Local Error Resume Next
lSystray.MoveMoveEvent Button, Shift, x, Y, Me
End Sub

Private Sub Form_Resize()
On Local Error GoTo ErrHandler
If frmMain.WindowState = vbMinimized Then frmMain.Hide
Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error Resume Next
lSystray.RemoveTrayIcon
Set lSystray = Nothing
Set lDesktopEnhancer = Nothing
End Sub

Private Sub mnuAbout_Click()
On Local Error Resume Next
frmAbout.Show 1
End Sub

Private Sub mnuChangeResolution_Click()
On Local Error Resume Next
frmChangeResolution.Show
End Sub

Private Sub mnuDesktops_Click(Index As Integer)
On Local Error GoTo ErrHandler
Dim b As Boolean, i As Integer
For i = 1 To 10
    mnuDesktops(i).Checked = False
Next i
mnuDesktops(Index).Checked = True
b = lDesktopEnhancer.SwitchDesktop(lDesktopEnhancer.ReturnCurrentDesktop, Index)
tmrUpdateCaptionDelay.Enabled = True
Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub mnuExit_Click()
On Local Error Resume Next
Unload Me
End Sub

Private Sub mnuVisualTaskManager_Click()
On Local Error Resume Next
Shell App.Path & "\Visual Task Manager.exe", vbNormalFocus
End Sub

Private Sub tmrUpdateCaptionDelay_Timer()
On Local Error Resume Next
SetDesktopMenuCaption
End Sub
