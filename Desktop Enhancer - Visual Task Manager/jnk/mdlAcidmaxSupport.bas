Attribute VB_Name = "mdlAcidmaxSupport"
Option Explicit
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Any) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const conSwNormal = 1
Global lAcidmaxHwnd As Long

Public Sub FormDrag(lFormname As Form)
On Local Error Resume Next
ReleaseCapture
Call SendMessage(lFormname.hWnd, &HA1, 2, 0&)
End Sub

Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)
On Local Error Resume Next
Dim lFlag As Integer
If SetOnTop Then
    lFlag = HWND_TOPMOST
Else
    lFlag = HWND_NOTOPMOST
End If
SetWindowPos myfrm.hWnd, lFlag, myfrm.Left / Screen.TwipsPerPixelX, myfrm.Top / Screen.TwipsPerPixelY, myfrm.Width / Screen.TwipsPerPixelX, myfrm.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Public Function WndEnumProc(ByVal hWnd As Long, ByVal lParam As Form) As Long
On Local Error Resume Next
Dim WText As String * 512, bRet As Long, WLen As Long, WClass As String * 50
WLen = GetWindowTextLength(hWnd)
bRet = GetWindowText(hWnd, WText, WLen + 1)
GetClassName hWnd, WClass, 50
If WLen <> 0 Then
    WText = Trim(WText)
    If InStr(1, LCase(WText), "Nexgen Acidmax", vbTextCompare) Then
        lAcidmaxHwnd = hWnd
    End If
End If
WndEnumProc = 1
End Function

Public Sub SetAcidmaxHwnd(lForm As Form)
On Local Error Resume Next
Dim l As Long
l = EnumWindows(AddressOf WndEnumProc, lForm)
End Sub

Public Sub SetParentAcidmax(lForm As Form)
On Local Error Resume Next
SetAcidmaxHwnd lForm
If lAcidmaxHwnd <> 0 Then
    SetParent lForm.hWnd, lAcidmaxHwnd
Else
    MsgBox "Nexgen Acidmax is not loaded. Click 'OK' to terminate program", vbExclamation
    End
End If
End Sub

