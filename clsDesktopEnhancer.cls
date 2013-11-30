VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDesktopEnhancer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal wIndx As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal wIndx As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpSting As String, ByVal nMaxCount As Long) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDNEXT = 2
Private Const GWL_STYLE = (-16)
Private Const SW_HIDE = 0
Private Const SW_SHOW = 5
Private Const WM_CLOSE = &H10
Private Const WS_VISIBLE = &H10000000
Private Const WS_BORDER = &H800000
Private lOpenWindows(0 To 10, 0 To 1023) As Long
Private lOpenWindowsCount(0 To 10) As Long
Private lCurrentDesktop As Integer
Private lPastDesktop As Integer
Private IsTask As Long

Public Function SwitchDesktop(lFromDesktop As Integer, lGotoDesktop As Integer) As Boolean
On Local Error GoTo ErrHandler
Dim lHwnd As Long, i As Long, msg As String, j As Integer, l As Long
SetDesktopMenuCaption
IsTask = WS_VISIBLE Or WS_BORDER
j = 0
lHwnd = GetWindow(frmMain.hWnd, GW_HWNDFIRST)
Do While lHwnd
    If lHwnd <> frmMain.hWnd And TaskWindow(lHwnd) Then
        i = GetWindowTextLength(lHwnd) + 1
        msg = Space$(i)
        i = GetWindowText(lHwnd, msg, i)
        If i > 0 Then
            If lHwnd <> frmMain.hWnd Then
                l = ShowWindow(lHwnd, SW_HIDE)
                lOpenWindows(lFromDesktop, j) = lHwnd
                j = j + 1
            End If
        End If
    End If
    lHwnd = GetWindow(lHwnd, GW_HWNDNEXT)
Loop
lOpenWindowsCount(lFromDesktop) = j
j = 0
While j < lOpenWindowsCount(lGotoDesktop)
    l = ShowWindow(lOpenWindows(lGotoDesktop, j), SW_SHOW)
    j = j + 1
Wend
lPastDesktop = lFromDesktop
lCurrentDesktop = lGotoDesktop
SwitchDesktop = True
For i = 0 To 10
    
Next i
Exit Function
ErrHandler:
    MsgBox Err.Description
End Function

Public Function TaskWindow(lHwnd As Long) As Long
On Local Error GoTo ErrHandler
Dim l As Long
l = GetWindowLong(lHwnd, GWL_STYLE)
If (l And IsTask) = IsTask Then TaskWindow = True
Exit Function
ErrHandler:
    MsgBox Err.Description
End Function

Public Function ExitDesktopEnhancer(lMoveToMainDesktop As Boolean) As Boolean
On Local Error GoTo ErrHandler
Dim l As Integer, i As Integer, m As Long
If lMoveToMainDesktop = True Then
    l = 1
    While l < 10
        i = 0
        While i < lOpenWindowsCount(l)
            m = ShowWindow(lOpenWindows(l, i), SW_SHOW)
            i = i + 1
        Wend
        l = l + 1
    Wend
    End
Else
    l = 2
    While l < 10
        i = 0
        While i < lOpenWindowsCount(l)
            m = SendMessage(lOpenWindows(l, i), WM_CLOSE, 0, 0)
            i = i + 1
        Wend
        l = l + 1
    Wend
    End
End If
ExitDesktopEnhancer = True
Exit Function
ErrHandler:
    MsgBox Err.Description
End Function

Public Function ReturnCurrentDesktop() As Integer
On Local Error GoTo ErrHandler
ReturnCurrentDesktop = lCurrentDesktop
Exit Function
ErrHandler:
    MsgBox Err.Description
End Function

Private Sub Class_Initialize()
On Local Error Resume Next
lCurrentDesktop = 1
lPastDesktop = 1
End Sub

Private Sub Class_Terminate()
On Local Error Resume Next
ExitDesktopEnhancer True
End Sub
