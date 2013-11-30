Attribute VB_Name = "mdlReadINI"
Option Explicit
#If Win32 Then
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#Else
    Private Declare Function ShellExecute Lib "shell.dll" (ByVal hWnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
#End If

Public Function ReadINI(ByVal lFile As String, ByVal Section As String, ByVal Key As String, Optional lDefault As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
Dim msg As String, RetVal As String, Worked As Integer
RetVal = String$(255, 0)
Worked = GetPrivateProfileString(Section, Key, "", RetVal, Len(RetVal), lFile)
If Worked = 0 Then
    ReadINI = lDefault
Else
    ReadINI = Left(RetVal, InStr(RetVal, Chr(0)) - 1)
End If
End Function

Public Sub WriteINI(ByVal lFile As String, ByVal Section As String, ByVal Key As String, ByVal Value As String)
If lSettings.sHandleErrors = True Then On Local Error Resume Next
WritePrivateProfileString Section, Key, Value, lFile
End Sub
