Attribute VB_Name = "mdlWindowFunctions"
Option Explicit
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal wIndx As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal wIndx As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpSting As String, ByVal nMaxCount As Long) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GWL_STYLE = (-16)
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const WM_CLOSE = &H10
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WS_VISIBLE = &H10000000
Public Const WS_BORDER = &H800000
Public openWindows(0 To 10, 0 To 1023) As Long
Public openWindowsCount(0 To 10) As Long
Public currentDesktop As Integer
Public pastDesktop As Integer
Public NotifyIcon As NOTIFYICONDATA
Public IsTask As Long

Public Function SwitchDesktop(lFromDesktop As Integer, lGotoDesktop As Integer)
On Local Error GoTo ErrHandler
Dim lhWnd As Long, i As Long, msg As String, j As Integer
IsTask = WS_VISIBLE Or WS_BORDER
j = 0
lhWnd = GetWindow(frmMain.hWnd, GW_HWNDFIRST)
Do While lhWnd
    If lhWnd <> frmMain.hWnd And TaskWindow(lhWnd) Then
        i = GetWindowTextLength(lhWnd) + 1
        msg = Space$(i)
        i = GetWindowText(lhWnd, msg, i)
        If i > 0 Then
            If lhWnd <> frmMain.hWnd Then
                RetVal = ShowWindow(lhWnd, SW_HIDE)
                openWindows(lFromDesktop, j) = lhWnd
                j = j + 1
            End If
        End If
    End If
    lhWnd = GetWindow(lhWnd, GW_HWNDNEXT)
Loop
openWindowsCount(lFromDesktop) = j
j = 0
While j < openWindowsCount(lGotoDesktop)
    RetVal = ShowWindow(openWindows(lGotoDesktop, j), SW_SHOW)
    j = j + 1
Wend
pastDesktop = lFromDesktop
currentDesktop = lGotoDesktop
Exit Function
ErrHandler:
    MsgBox Err.Description
End Function

Public Function TaskWindow(lhWnd As Long) As Long
On Local Error GoTo ErrHandler
Dim lngStyle As Long
lngStyle = GetWindowLong(lhWnd, GWL_STYLE)
If (lngStyle And IsTask) = IsTask Then TaskWindow = True
Exit Function
ErrHandler:
    MsgBox Err.Description
End Function

Public Function ExitDesktopEnhancer(lMoveToMainDesktop As Boolean) As Boolean
On Local Error GoTo ErrHandler
Dim l As Integer, i As Integer
If lMoveToMainDesktop = True Then
    l = 1
    While l < 10
        i = 0
        While i < openWindowsCount(l)
            RetVal = ShowWindow(openWindows(l, i), SW_SHOW)
            i = i + 1
        Wend
        l = l + 1
    Wend
    Shell_NotifyIcon NIM_DELETE, NotifyIcon
    End
Else
    l = 2
    While l < 10
        i = 0
        While i < openWindowsCount(l)
            RetVal = SendMessage(openWindows(l, i), WM_CLOSE, 0, 0)
            i = i + 1
        Wend
        l = l + 1
    Wend
    Shell_NotifyIcon NIM_DELETE, NotifyIcon
    End
End If
Exit Function
ErrHandler:
    MsgBox Err.Description
End Function
