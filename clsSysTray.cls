VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSystray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private NotifyIcon As NOTIFYICONDATA
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const NIF_TIP = &H4
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const WM_MOUSEMOVE = &H200
Private Const WM_RBUTTONUP = &H205
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Sub LoadTrayIcon()
On Local Error GoTo ErrHandler
With NotifyIcon
    .cbSize = Len(NotifyIcon)
    .hWnd = frmMain.hWnd
    .uId = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uCallBackMessage = WM_MOUSEMOVE
    .hIcon = frmMain.Icon
    .szTip = "Desktop Enhancer" & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, NotifyIcon
Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Public Sub RemoveTrayIcon()
On Local Error GoTo ErrHandler
Shell_NotifyIcon NIM_DELETE, NotifyIcon
Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Public Sub MoveMoveEvent(Button As Integer, Shift As Integer, x As Single, Y As Single, lForm As Form)
On Local Error GoTo ErrHandler
Dim l As Long, i As Long
If lForm.ScaleMode = vbPixels Then
    i = x
Else
    i = x / Screen.TwipsPerPixelX
End If
If i = WM_RBUTTONUP Then
    l = SetForegroundWindow(lForm.hWnd)
    lForm.PopupMenu lForm.mnuHidden
End If
Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub
