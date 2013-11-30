Attribute VB_Name = "mdlUpdateAdvancedAddons"
Option Explicit

Sub Main()
On Local Error Resume Next
Pause 3
Kill App.Path & "\AdvancedAddons.exe"
DoEvents
Pause 0.2
If DoesFileExist(App.Path & "\AdvancedAddons.exe") = True Then
    If Err.Number <> 0 Then
        MsgBox "Advanced Addons is still running. This program will now terminate.", vbCritical
        End
    End If
End If
DownloadFile "http://www.team-nexgen.com/downloads/addons/AdvancedAddons.exe", App.Path & "\AdvancedAddons.exe": DoEvents
Pause 0.2
If DoesFileExist(App.Path & "\AdvancedAddons.exe") = True Then
    Shell App.Path & "\AdvancedAddons.exe", vbNormalFocus
End If
End Sub

Public Sub DownloadFile(lUrl As String, lFile As String)
On Local Error Resume Next
Dim f As New frmDownloadFile
Set f = New frmDownloadFile
f.Show
f.ClickDownloadButton lUrl, lFile
End Sub

Public Function DoesFileExist(lFilename As String) As Boolean
On Local Error Resume Next
Dim msg As String
msg = Dir(lFilename)
If msg <> "" Then
    DoesFileExist = True
Else
    DoesFileExist = False
End If
End Function

Public Sub Pause(interval)
On Local Error Resume Next
Dim Current
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

