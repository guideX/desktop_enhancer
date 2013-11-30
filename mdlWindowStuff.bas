Attribute VB_Name = "mdlWindowStuff"
Option Explicit
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Const WM_SETTEXT = &HC
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_GETTEXT = &HD
Private i As Integer, msg As String, l As Long

Public Sub SetDesktopMenuCaption()
On Local Error Resume Next
msg = ""
For i = 1 To 10
    If frmMain.mnuDesktops(i).Checked = True Then
        msg = GetText(GetForegroundWindow())
        If LCase(msg) <> "desktop enhancer" And Len(msg) <> 0 Then frmMain.mnuDesktops(i).Caption = Trim(Str(i)) & " (" & msg & ")"
    End If
Next i
End Sub

Public Function GetText(lHwnd As Long) As String
On Local Error Resume Next
l = 0
l = SendMessage(lHwnd, WM_GETTEXTLENGTH, 0, 0)
If l = 0 Then
    GetText = ""
    Exit Function
End If
l = l + 1
msg = Space(l)
l = SendMessage(lHwnd, WM_GETTEXT, l, ByVal msg)
GetText = Left(msg, l)
If Err.Number <> 0 Then
    Err.Clear
End If
End Function
