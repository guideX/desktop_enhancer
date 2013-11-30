Attribute VB_Name = "mdlAddons"
Option Explicit
Private Type gAddon
    aName As String
    aVersion As String
    aExeUrl As String
    aPngUrl As String
    aFile As String
End Type
Private Type gAddons
    aAddons(150) As gAddon
    aCount As Integer
End Type
Private Type gSettings
    sDockVisible As Boolean
End Type
Private Const SW_SHOWNORMAL = 1
Global lSettings As gSettings
Global lAddons As gAddons
Global lIniFile As String

'Sub Main()
'On Local Error Resume Next
'LoadAddons
'frmDock.Show
'End Sub

Public Sub DownloadFile(lUrl As String, lFile As String)
On Local Error Resume Next
Dim f As New frmDownloadFile
Set f = New frmDownloadFile
f.Show
f.ClickDownloadButton lUrl, lFile
End Sub

Public Sub SaveAddons()
On Local Error Resume Next
Dim i As Integer
If DoesFileExist(lIniFile) = True Then Kill lIniFile
WriteINI lIniFile, "Settings", "Count", Trim(str(lAddons.aCount))
For i = 1 To lAddons.aCount
    WriteINI lIniFile, Trim(str(i)), "ExeUrl", lAddons.aAddons(i).aExeUrl
    WriteINI lIniFile, Trim(str(i)), "File", lAddons.aAddons(i).aFile
    WriteINI lIniFile, Trim(str(i)), "Name", lAddons.aAddons(i).aName
    WriteINI lIniFile, Trim(str(i)), "PngUrl", lAddons.aAddons(i).aPngUrl
    WriteINI lIniFile, Trim(str(i)), "Version", lAddons.aAddons(i).aVersion
Next i
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

Public Function IsAddonInstalled(lName As String) As Boolean
On Local Error Resume Next
Dim i As Integer
For i = 1 To 150
    If LCase(lAddons.aAddons(i).aName) = LCase(lName) Then
        IsAddonInstalled = True
        Exit For
    End If
Next i
End Function

Public Function FindAddonIndex(lName As String) As Integer
On Local Error Resume Next
Dim i As Integer
For i = 1 To 150
    If LCase(lAddons.aAddons(i).aName) = LCase(lName) Then
        FindAddonIndex = i
        Exit For
    End If
Next i
End Function

Public Sub LoadAddons()
On Local Error Resume Next
Dim i As Integer
lIniFile = App.Path & "\addons.ini"
lAddons.aCount = ReadINI(lIniFile, "Settings", "Count", 0)
If lAddons.aCount <> 0 Then
    For i = 1 To lAddons.aCount
        lAddons.aAddons(i).aName = ReadINI(lIniFile, Trim(str(i)), "Name", "")
        lAddons.aAddons(i).aFile = App.Path & "\icons\" & lAddons.aAddons(i).aName & ".exe"
        lAddons.aAddons(i).aVersion = ReadINI(lIniFile, Trim(str(i)), "Version", "")
        If DoesFileExist(lAddons.aAddons(i).aFile) = True And Len(lAddons.aAddons(i).aName) <> 0 Then
            lAddons.aAddons(i).aExeUrl = ReadINI(lIniFile, Trim(str(i)), "ExeUrl", "")
            lAddons.aAddons(i).aPngUrl = ReadINI(lIniFile, Trim(str(i)), "PngUrl", "")
        Else
            lAddons.aAddons(i).aFile = ""
            lAddons.aAddons(i).aName = ""
        End If
    Next i
End If
End Sub

Public Sub CheckForNewAddons()
On Local Error Resume Next
Dim c As Integer, f As Integer, i As Integer, msg As String, msg2 As String, msg3 As String, msg4 As String, mbox As VbMsgBoxResult
If DoesFileExist(App.Path & "\addons.ini") = True Then Kill App.Path & "\addons.ini"
DownloadFile "http://www.team-nexgen.com/downloads/addons/addons.ini", App.Path & "\addons.ini"
DoEvents
If DoesFileExist(App.Path & "\addons.ini") = True Then
    c = Int(ReadINI(App.Path & "\addons.ini", "Settings", "Count", 0))
    If c <> 0 Then
        For i = 1 To c
            msg = ReadINI(lIniFile, Trim(str(i)), "Name", "")
            If IsAddonInstalled(msg) = False Then
                msg2 = ReadINI(lIniFile, Trim(str(i)), "Description", "")
                msg3 = ReadINI(lIniFile, Trim(str(i)), "ExeUrl", "")
                msg4 = ReadINI(lIniFile, Trim(str(i)), "PngUrl", "")
                mbox = MsgBox("A new Addon is available" & vbCrLf & "Name: " & msg & vbCrLf & "Description: " & msg2 & vbCrLf & vbCrLf & "Would you like to download it?", vbYesNo + vbQuestion)
                If mbox = vbYes And Len(msg3) <> 0 And Len(msg4) <> 0 Then
                    DownloadFile msg4, App.Path & "\icons\" & UCase(msg) & ".PNG": DoEvents
                    DownloadFile msg3, App.Path & "\icons\" & msg & ".exe": DoEvents
                    lAddons.aCount = lAddons.aCount + 1
                    lAddons.aAddons(lAddons.aCount).aName = msg
                    lAddons.aAddons(lAddons.aCount).aFile = App.Path & "\icons\" & msg & ".exe"
                    f = f + 1
                End If
            End If
        Next i
    Else
        SaveAddons
    End If
Else
    SaveAddons
End If

If f = 0 Then
    If c = lAddons.aCount Then
        MsgBox "No new Addons are currently available", vbInformation
    Else
        MsgBox "No new Addons were installed", vbInformation
    End If
Else
    Shell App.Path & "\AdvancedAddons.exe", vbNormalFocus
    EndProgram
End If
End Sub

Public Sub Pause(interval)
On Local Error Resume Next
Dim Current
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

Public Sub UpdateExistingAddons()
On Local Error Resume Next
Dim c As Integer, f As Integer, b As Integer, i As Integer, msg As String, msg2 As String, msg3 As String, msg4 As String, mbox As VbMsgBoxResult
If DoesFileExist(App.Path & "\addons.ini") = True Then Kill App.Path & "\addons.ini"
DownloadFile "http://www.team-nexgen.com/downloads/addons/addons.ini", App.Path & "\addons.ini"
DoEvents
If DoesFileExist(App.Path & "\addons.ini") = True Then
    If Len(LCase(ReadINI(App.Path & "\addons.ini", "Settings", "Version", ""))) <> 0 And LCase(ReadINI(App.Path & "\addons.ini", "Settings", "Version", "")) <> App.Major & "." & App.Minor & App.Revision Then
        mbox = MsgBox("A newer version of the Advanced Addons program dock is now available for Download. " & vbCrLf & "New Version: " & LCase(ReadINI(App.Path & "\addons.ini", "Settings", "Version", "")) & vbCrLf & "Your Version: " & App.Major & "." & App.Minor & App.Revision & vbCrLf & "Would you like to download it?", vbYesNo + vbQuestion)
        If mbox = vbYes Then
            Shell App.Path & "\UpdateAdvancedAddons.exe", vbNormalFocus
            End
        End If
    End If
    c = Int(ReadINI(App.Path & "\addons.ini", "Settings", "Count", 0))
    If c <> 0 Then
        For i = 1 To c
            msg = ReadINI(lIniFile, Trim(str(i)), "Name", "")
            If IsAddonInstalled(msg) = True Then
                f = FindAddonIndex(msg)
                msg2 = ReadINI(App.Path & "\addons.ini", Trim(str(i)), "Version", "")
                msg3 = ReadINI(App.Path & "\addons.ini", Trim(str(i)), "ExeUrl", "")
                If Len(msg2) <> 0 Then
                    If LCase(msg2) <> LCase(lAddons.aAddons(f).aVersion) Then
                        mbox = MsgBox("An Update to '" & lAddons.aAddons(f).aName & "' is currently available." & vbCrLf & "New Version: " & msg2 & vbCrLf & "Your Version: " & lAddons.aAddons(f).aVersion & vbCrLf & " would you like to update this Addon?", vbYesNo + vbQuestion)
                        If mbox = vbYes Then
                            Kill lAddons.aAddons(f).aFile: DoEvents
                            Pause 0.2
                            If DoesFileExist(lAddons.aAddons(f).aFile) = False Then
                                DownloadFile msg3, lAddons.aAddons(f).aFile: DoEvents
                                lAddons.aAddons(f).aVersion = msg2
                                b = b + 1
                            Else
                                MsgBox "Unable to delete the file '" & lAddons.aAddons(f).aFile & "'. Please make sure you are not running this Addon and try again!", vbCritical
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        Next i
    Else
        SaveAddons
    End If
Else
    SaveAddons
End If
    If b = 0 Then
        MsgBox "Nothing was updated/no updates were available", vbInformation
    Else
        MsgBox str(b) & " Addon(s) were updated", vbInformation
    End If
End Sub

Public Sub Surf(lUrl As String, lHwnd As Long)
On Local Error Resume Next
Dim msg As Long, mbox As VbMsgBoxResult
On Local Error Resume Next
If Left(LCase(lUrl), 7) <> "http://" Then lUrl = "http://" & lUrl
msg = ShellExecute(lHwnd, vbNullString, lUrl, vbNullString, "c:\", SW_SHOWNORMAL)
End Sub

Public Sub EndProgram()
On Local Error Resume Next
If lSettings.sDockVisible = True Then Unload frmDock
End
End Sub
