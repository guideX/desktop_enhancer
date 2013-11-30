Attribute VB_Name = "mdlTaskManager"
Option Explicit
Private Type gTaskManagerImage
    tImage As String
    tDescription As String
End Type
Private Type gTaskManagerImages
    tImages(100) As gTaskManagerImage
    tCount As Integer
End Type
Private lTaskManagerImages As gTaskManagerImages

Public Sub Pause(interval)
'On Local Error Resume Next
Dim Current
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

Public Sub DownloadFile(lUrl As String, lFile As String)
'On Local Error Resume Next
Dim f As New frmDownloadFile
Set f = New frmDownloadFile
f.Show
f.ClickDownloadButton lUrl, lFile
End Sub

Public Function Parse(lWhole As String, lStart As String, lEnd As String)
'On Local Error GoTo ErrHandler
Dim len1 As Integer, len2 As Integer, Str1 As String, Str2 As String
len1 = InStr(lWhole, lStart)
len2 = InStr(lWhole, lEnd)
Str1 = Right(lWhole, Len(lWhole) - len1)
Str2 = Right(lWhole, Len(lWhole) - len2)
Parse = Left(Str1, Len(Str1) - Len(Str2) - 1)
ErrHandler:
End Function

Public Sub Surf(lUrl As String, lHwnd As Long)
'On Local Error Resume Next
Dim msg As Long, mbox As VbMsgBoxResult
'On Local Error Resume Next
If Left(LCase(lUrl), 7) <> "http://" Then lUrl = "http://" & lUrl
msg = ShellExecute(lHwnd, vbNullString, lUrl, vbNullString, "c:\", SW_SHOWNORMAL)
End Sub

Public Sub EndProgram()
'On Local Error Resume Next

Unload frmDock
End
End Sub

Public Function DoesFileExist(lFilename As String) As Boolean
'On Local Error Resume Next
Dim msg As String
msg = Dir(lFilename)
If msg <> "" Then
    DoesFileExist = True
Else
    DoesFileExist = False
End If
End Function

Sub Main()
'On Local Error Resume Next
LoadTaskManagerImages: DoEvents
frmDock.Show
End Sub

Public Sub LoadTaskManagerImages()
'On Local Error Resume Next
Dim i As Integer, msg() As String
lIniFile = App.Path & "\tasks.ini"
lTaskManagerImages.tCount = ReadINI(lIniFile, "Settings", "Count", 0)
If lTaskManagerImages.tCount <> 0 Then
    For i = 1 To lTaskManagerImages.tCount
        With lTaskManagerImages
            .tImages(i).tImage = ReadINI(lIniFile, Trim(str(i)), "Image", "")
            .tImages(i).tDescription = ReadINI(lIniFile, Trim(str(i)), "Description", "")
        End With
    Next i
End If
End Sub
